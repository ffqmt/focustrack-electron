require('dotenv').config();

const express = require('express');
const cors = require('cors');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const {
  getSite,
  getListById,
  getListItems,
  createListItem,
  updateListItem,
  getListItemById,
  getListItemsByQuery,
  getTokenDebugInfo,
  testConnection
} = require('./services/sharepoint');

const app = express();
const finishLocks = new Set();

app.use(cors());
app.use(express.json());

/**
 * ============================================================
 * CONFIGURAÇÃO SHAREPOINT
 * ============================================================
 */

function getSharePointConfig() {
  return {
    siteUrl: process.env.SHAREPOINT_SITE_URL,
    ticketsListId: process.env.SHAREPOINT_TICKETS_LIST_ID,
    timeEntriesListId: process.env.SHAREPOINT_TIME_ENTRIES_LIST_ID,

    /**
     * Campos Pessoa/Grupo
     */
    ticketRequesterLookupField:
      process.env.SP_FIELD_TICKET_REQUESTER_LOOKUP || 'SolicitanteLookupId',

    timeResponsibleLookupField:
      process.env.SP_FIELD_TIME_RESPONSIBLE_LOOKUP || 'Respons_x00e1_velLookupId',

    /**
     * Campo Data Criação customizado da lista Demandas/CentralDemandas
     *
     * Campo visual: Data Criação
     * Nome interno: DataCria_x00e7__x00e3_o
     */
    ticketCreatedAtField:
      process.env.SP_FIELD_TICKET_CREATED_AT || 'DataCria_x00e7__x00e3_o',

    /**
     * Campos de acompanhamento da demanda/chamado.
     *
     * Campos confirmados na lista Demandas/CentralDemandas:
     * Status => Status
     * Iniciado em => InicioPlanejado
     * Data Conclusão => DataConclus_x00e3_o
     * Tempo Gasto => TempoGasto
     */
    ticketStatusField:
      process.env.SP_FIELD_TICKET_STATUS || 'Status',

    ticketStartedAtField:
      process.env.SP_FIELD_TICKET_STARTED_AT || 'InicioPlanejado',

    ticketFinishedAtField:
      process.env.SP_FIELD_TICKET_FINISHED_AT || 'DataConclus_x00e3_o',

    ticketTimeSpentField:
      process.env.SP_FIELD_TICKET_TIME_SPENT || 'TempoGasto',

    ticketStatusInProgressValue:
      process.env.SP_VALUE_TICKET_STATUS_IN_PROGRESS || 'Em andamento',

    ticketStatusFinishedValue:
      process.env.SP_VALUE_TICKET_STATUS_FINISHED || 'Concluído',

    /**
     * Campos técnicos para integração com Power Automate.
     * Usados na lista LancamentoTempo.
     */
    timeCentralCommentField:
      process.env.SP_FIELD_TIME_CENTRAL_COMMENT || 'ComentarioCentralDemandas',

    timeSendCentralCommentField:
      process.env.SP_FIELD_TIME_SEND_CENTRAL_COMMENT || 'EnviarComentarioCentral',

    timeCentralCommentSentField:
      process.env.SP_FIELD_TIME_CENTRAL_COMMENT_SENT || 'ComentarioCentralEnviado',

    timeCentralCommentErrorField:
      process.env.SP_FIELD_TIME_CENTRAL_COMMENT_ERROR || 'ErroComentarioCentral',

    /**
     * Comentários nativos do item no SharePoint/Microsoft Lists.
     */
    nativeCommentsEnabled:
      String(
        process.env.SP_NATIVE_COMMENTS_ENABLED ||
        process.env.SHAREPOINT_NATIVE_COMMENTS_ENABLED ||
        'true'
      ).toLowerCase() !== 'false'
  };
}

function validateRequiredConfig() {
  const { siteUrl, ticketsListId, timeEntriesListId } = getSharePointConfig();

  const missing = [];

  if (!siteUrl) missing.push('SHAREPOINT_SITE_URL');
  if (!ticketsListId) missing.push('SHAREPOINT_TICKETS_LIST_ID');
  if (!timeEntriesListId) missing.push('SHAREPOINT_TIME_ENTRIES_LIST_ID');

  return {
    ok: missing.length === 0,
    missing
  };
}

async function resolveSiteId() {
  const { siteUrl } = getSharePointConfig();
  const site = await getSite(siteUrl);
  return site.id;
}

/**
 * ============================================================
 * HELPERS GERAIS
 * ============================================================
 */

function removeNullishFields(obj) {
  return Object.fromEntries(
    Object.entries(obj).filter(([, value]) => {
      return value !== null && value !== undefined && value !== '';
    })
  );
}

function calculateMinutes(startAt, endAt) {
  if (!startAt || !endAt) return 0;

  const diffMs = new Date(endAt).getTime() - new Date(startAt).getTime();

  if (!Number.isFinite(diffMs) || diffMs <= 0) {
    return 0;
  }

  return Math.max(1, Math.ceil(diffMs / 60000));
}

function normalizeEmail(email) {
  return String(email || '').trim().toLowerCase();
}

function toNumberOrNull(value) {
  if (value === null || value === undefined || value === '') return null;

  const n = Number(value);

  if (!Number.isFinite(n) || n <= 0) {
    return null;
  }

  return n;
}

function toSafeNumber(value) {
  if (value === null || value === undefined || value === '') return 0;

  const n = Number(value);

  if (!Number.isFinite(n)) {
    return 0;
  }

  return n;
}

function isFilledValue(value) {
  return value !== null && value !== undefined && value !== '';
}

function getTimeEntriesSumTop() {
  const value = Number(process.env.SP_TIME_ENTRIES_SUM_TOP || 999);

  if (!Number.isFinite(value) || value <= 0) {
    return 999;
  }

  return Math.min(value, 999);
}

function normalizeDateTimeOrNow(value) {
  if (!value) {
    return new Date().toISOString();
  }

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return new Date().toISOString();
  }

  return date.toISOString();
}

function formatDateTimeBR(value) {
  if (!value) return '';

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return String(value);
  }

  return new Intl.DateTimeFormat('pt-BR', {
    timeZone: process.env.TZ || 'America/Cuiaba',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

function formatTimeBR(value) {
  if (!value) return '';

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return String(value);
  }

  return new Intl.DateTimeFormat('pt-BR', {
    timeZone: process.env.TZ || 'America/Cuiaba',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

function sanitizeText(value) {
  return String(value || '').trim();
}

function ptbrDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function getCentralCommentMaxChars() {
  return Number(process.env.SP_CENTRAL_COMMENT_MAX_CHARS || 2000);
}

function getUserLookupMap() {
  try {
    return JSON.parse(process.env.SP_USER_LOOKUP_MAP || '{}');
  } catch {
    return {};
  }
}

function getDefaultUserEmail() {
  return process.env.SP_DEFAULT_USER_EMAIL || '';
}

function getDefaultUserLookupId() {
  return Number(process.env.SP_DEFAULT_USER_LOOKUP_ID || 0);
}

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text || '').replace(/[&<>"']/g, m => map[m]);
}

function renderMetricCell(label, value, color) {
  return `
<td width="25%" style="padding:0 6px 0 0;">
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-top:4px solid ${color};border:1px solid #d9e2ec;border-top:4px solid ${color};border-collapse:collapse;">
    <tr><td style="padding:18px 16px 4px 16px;font-size:32px;font-weight:bold;color:#0d2b4c;font-family:Arial,sans-serif;line-height:1; text-align: center;">${value}</td></tr>
    <tr><td style="padding:0 16px 16px 16px;font-size:10px;color:#6b7a90;font-family:Arial,sans-serif;letter-spacing:1px;text-transform:uppercase; text-align: center;">${label}</td></tr>
  </table>
</td>`;
}

function setDynamicField(fields, fieldName, value) {
  if (!fieldName) return;
  if (value === null || value === undefined || value === '') return;

  fields[fieldName] = value;
}

function formatMinutesToHours(minutes) {
  const m = toSafeNumber(minutes);
  const hours = Math.floor(m / 60);
  const remainingMinutes = m % 60;
  return `${hours}h${String(remainingMinutes).padStart(2, '0')}`;
}

function hasWinTag(text) {
  return /(^|\s)#win\b/i.test(String(text || ''));
}

function cleanWinTag(text) {
  return String(text || '')
    .replace(/(^|\s)#win\b/ig, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// ── Backlog por status real ─────────────────────────────────
function normalizeStatusLabel(status) {
  return String(status || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function isBacklogStatus(status) {
  const s = normalizeStatusLabel(status);
  return [
    'novo',
    'triagem',
    'em triagem',
    'backlog',
    'pendente',
    'aguardando priorizacao',
    'aguardando analise'
  ].includes(s);
}

function isInProgressStatus(status) {
  const s = normalizeStatusLabel(status);
  return [
    'em andamento',
    'andamento',
    'em execucao',
    'executando',
    'doing',
    'in progress'
  ].includes(s);
}

function isFinishedStatus(status) {
  const s = normalizeStatusLabel(status);
  return ['concluido', 'finalizado', 'done', 'finished'].includes(s);
}

// ── Tags manuais ────────────────────────────────────────────
const REPORT_TAGS = {
  DECISAO: ['#decisao', '#decisao_gestao', '#gestao'],
  RISCO: ['#risco', '#atencao', '#alerta'],
  DEPENDENCIA: ['#dependencia', '#bloqueio', '#aguardando'],
  AUTOMACAO: ['#automacao', '#economia', '#bot', '#robo'],
  URGENTE: ['#urgente', '#prioridade', '#critico']
};

function extractTagsFromText(text) {
  const raw = String(text || '').toLowerCase();
  const found = {};
  let economia = 0;

  for (const [key, aliases] of Object.entries(REPORT_TAGS)) {
    if (aliases.some(alias => raw.includes(alias))) {
      found[key] = true;
    }
  }

  // Parse #economia:4.5 or #economia 4.5h
  const econMatch = raw.match(/#(?:economia|automacao)[:\s](\d+(?:\.\d+)?)(?:h\b)?/);
  if (econMatch) {
    economia = parseFloat(econMatch[1]);
  }

  return { 
    tags: Object.keys(found),
    economia
  };
}

function cleanManualSignals(text) {
  let out = String(text || '');
  for (const tag of MANUAL_SIGNAL_TAGS) {
    out = out.replace(new RegExp(`#${tag}\\b`, 'gi'), '');
  }
  return out.replace(/\s+/g, ' ').trim();
}

// ── Frente de gestão ────────────────────────────────────────
function inferManagementFront({ demanda, comentarios }) {
  if (isBacklogStatus(demanda?.status)) return 'Backlog e triagem';
  const text = [
    demanda?.titulo,
    demanda?.descricao,
    demanda?.origem,
    ...(comentarios || []).map(c => typeof c === 'string' ? c : (c.texto || c.comentario || c.body || ''))
  ].join(' ').toLowerCase();

  if (/ia\b|pdf|ocr|imagem|extrato|planilha|dados|excel|monica|agente|automa|robo|robô/.test(text)) return 'Automação, IA e dados';
  if (/dctfweb|sped|onix|integra|fiscal|api|contador|competência|sefaz|nfse|nfe/.test(text)) return 'Fiscal e integrações';
  if (/onedrive|teams|azure|sharepoint|outlook|usuário|grupo|e-mail|email/.test(text)) return 'Microsoft 365, acessos e rede';
  if (/rede|máquina|\bpc\b|servidor|office|windows|domínio|impressora|acesso/.test(text)) return 'Infraestrutura e suporte';
  if (/contai|focustrack|manager|sistema|plataforma|electron/.test(text)) return 'Sistemas internos e plataformas';
  if (/documentação|manual|procedimento|passo a passo|padronização/.test(text)) return 'Documentação e padronização';
  return 'Suporte operacional';
}

// ── Classificação executiva de comentário ───────────────────
const SIGNAL_META = {
  WIN:     { tipo: 'vitoria',       peso: 5, leitura: 'Vitória sinalizada manualmente, indicando avanço percebido como relevante na semana.' },
  RISCO:   { tipo: 'risco',         peso: 5, leitura: 'Risco, bloqueio ou dependência apontado no comentário, exigindo acompanhamento.' },
  DECISAO: { tipo: 'decisao',       peso: 5, leitura: 'Comentário indica necessidade de decisão ou alinhamento da gestão.' },
  IMPACTO: { tipo: 'impacto',       peso: 5, leitura: 'Comentário destaca impacto operacional, ganho ou benefício para a organização.' },
  PROXIMO: { tipo: 'proximo_passo', peso: 4, leitura: 'Comentário registra próximo passo claro para continuidade da demanda.' },
  DOC:     { tipo: 'documentacao',  peso: 4, leitura: 'Comentário relacionado à documentação, padronização ou transferência de conhecimento.' },
  IA:      { tipo: 'ia',            peso: 4, leitura: 'Comentário associado ao uso de IA, automação inteligente ou estruturação de dados.' }
};
const SIGNAL_PRIORITY = ['WIN', 'RISCO', 'DECISAO', 'IMPACTO', 'PROXIMO', 'DOC', 'IA'];

function classifyCommentForExecutiveReport({ comentario, demanda }) {
  const texto_original = typeof comentario === 'string' ? comentario : (comentario?.texto || '');
  const sinais = extractManualSignals(texto_original);
  const texto_limpo = cleanManualSignals(texto_original);
  const frente = inferManagementFront({ demanda, comentarios: [texto_original] });

  const tagPrioritaria = SIGNAL_PRIORITY.find(t => sinais.includes(t));
  const meta = tagPrioritaria ? SIGNAL_META[tagPrioritaria] : null;

  let tipo, peso, leitura;
  if (meta) {
    tipo = meta.tipo;
    peso = meta.peso;
    leitura = meta.leitura;
  } else if (texto_limpo.length > 60) {
    tipo = 'evidencia';
    peso = 3;
    leitura = 'Comentário com detalhamento operacional relevante para entender o andamento da demanda.';
  } else {
    tipo = 'comentario';
    peso = texto_limpo.length > 20 ? 2 : 1;
    leitura = '';
  }

  return {
    texto_original,
    texto_limpo,
    demanda_id: demanda?.id || '',
    demanda_titulo: demanda?.titulo || '',
    status: demanda?.status || '',
    origem: demanda?.origem || '',
    sinais,
    frente,
    tipo,
    peso_executivo: peso,
    leitura_executiva: leitura
  };
}

// ── Carteira executiva ──────────────────────────────────────
function buildExecutivePortfolio({ demandas, comentariosExecutivos, backlogExecutivo }) {
  const winIds = new Set(
    (comentariosExecutivos || []).filter(c => c.sinais.includes('WIN')).map(c => String(c.demanda_id))
  );

  const porFrenteMap = {};
  (demandas || []).forEach(d => {
    const comentsDemanda = (comentariosExecutivos || [])
      .filter(c => String(c.demanda_id) === String(d.id))
      .map(c => c.texto_original);
    const frente = inferManagementFront({ demanda: d, comentarios: comentsDemanda });
    if (!porFrenteMap[frente]) porFrenteMap[frente] = [];
    porFrenteMap[frente].push({ demanda_id: d.id, titulo: d.titulo, status: d.status, origem: d.origem });
  });

  const FRENTE_LEITURA = {
    'Automação, IA e dados': 'Frente com concentração de demandas ligadas a automação, IA, dados ou estruturação de informações.',
    'Fiscal e integrações': 'Frente com foco em compliance fiscal, integrações de sistemas e obrigações acessórias.',
    'Microsoft 365, acessos e rede': 'Frente de gestão de acessos, identidade e plataformas Microsoft 365.',
    'Infraestrutura e suporte': 'Frente de sustentação de infraestrutura, equipamentos e suporte operacional.',
    'Sistemas internos e plataformas': 'Frente de evolução e manutenção de sistemas internos da organização.',
    'Documentação e padronização': 'Frente de documentação, padronização de processos e transferência de conhecimento.',
    'Backlog e triagem': 'Demandas aguardando triagem ou definição de escopo e prioridade.',
    'Suporte operacional': 'Frente de suporte direto a usuários e operações internas.'
  };

  const porFrente = Object.entries(porFrenteMap)
    .filter(([frente]) => frente !== 'Backlog e triagem')
    .map(([frente, arr]) => ({
      frente,
      quantidade: arr.length,
      demandas: arr,
      leitura: FRENTE_LEITURA[frente] || 'Frente com demandas de natureza mista registradas no período.'
    }));

  const emEvolucao = (demandas || []).filter(d => {
    const peso = Math.max(0, ...(comentariosExecutivos || [])
      .filter(c => String(c.demanda_id) === String(d.id))
      .map(c => c.peso_executivo));
    const temTag = (comentariosExecutivos || [])
      .filter(c => String(c.demanda_id) === String(d.id))
      .some(c => ['PROXIMO','DOC','IA'].some(t => c.sinais.includes(t)));
    return d.status === 'Em andamento' || peso >= 3 || temTag;
  }).map(d => ({ demanda_id: d.id, titulo: d.titulo, status: d.status }));

  const vitoriasParciais = (demandas || [])
    .filter(d => d.status !== 'Concluído' && winIds.has(String(d.id)))
    .map(d => ({ demanda_id: d.id, titulo: d.titulo, status: d.status }));

  const semAvanco = (demandas || []).filter(d => {
    if (d.status === 'Concluído') return false;
    if (isBacklogStatus(d.status)) return false;
    const temComentario = (comentariosExecutivos || []).some(c => String(c.demanda_id) === String(d.id) && c.peso_executivo >= 2);
    return !temComentario && !winIds.has(String(d.id));
  }).map(d => ({ demanda_id: d.id, titulo: d.titulo, status: d.status }));

  return {
    por_frente: porFrente,
    backlog_executivo: backlogExecutivo || [],
    em_evolucao: emEvolucao,
    vitorias_parciais: vitoriasParciais,
    sem_avanco_claro: semAvanco
  };
}

// ── Temas e insights ────────────────────────────────────────
function buildWeeklyThemes({ resumoParaIa, comentariosExecutivos, carteiraExecutiva }) {
  const temas = [];
  const seen = new Set();

  const add = (tema, evidencia, impacto_gestao) => {
    if (temas.length >= 5 || seen.has(tema)) return;
    seen.add(tema);
    temas.push({ tema, evidencia, impacto_gestao });
  };

  const frentes = (carteiraExecutiva?.por_frente || []);
  frentes.slice(0, 3).forEach(f => {
    add(f.frente, f.leitura, `Frente com ${f.quantidade} demanda${f.quantidade > 1 ? 's' : ''} — monitorar evolução e resultados.`);
  });

  const haveIa = (comentariosExecutivos || []).some(c => c.sinais.includes('IA'));
  if (haveIa) add('Automação e IA aplicada a dados', 'Comentários indicam uso de agente/IA para extração ou estruturação de dados.', 'Pode reduzir retrabalho manual e abrir caminho para padronização operacional.');

  const haveDoc = (comentariosExecutivos || []).some(c => c.sinais.includes('DOC'));
  if (haveDoc) add('Documentação e padronização', 'Comentários registram produção de documentação ou padronização de processo.', 'Facilita transferência de conhecimento e continuidade operacional.');

  const haveRisco = (resumoParaIa?.sinais_manuais?.riscos || []).length > 0;
  if (haveRisco) add('Riscos e dependências identificados', 'Comentários registram pontos de atenção que requerem acompanhamento.', 'Demandam decisão ou acompanhamento para não impactar a continuidade.');

  return temas;
}

function buildOperationalInsights({ resumoParaIa, carteiraExecutiva }) {
  const insights = [];

  const wins = (resumoParaIa?.sinais_manuais?.wins || []);
  if (wins.length > 0) insights.push({ titulo: 'Avanços da semana', descricao: `${wins.length} vitória${wins.length > 1 ? 's' : ''} sinalizada${wins.length > 1 ? 's' : ''} pela equipe — resultado percebido com impacto direto.`, sinal: 'avanço' });

  const riscos = (resumoParaIa?.sinais_manuais?.riscos || []);
  if (riscos.length > 0) insights.push({ titulo: 'Riscos e bloqueios', descricao: `${riscos.length} ponto${riscos.length > 1 ? 's' : ''} de risco ou bloqueio registrado${riscos.length > 1 ? 's' : ''} em comentários — exigem acompanhamento.`, sinal: 'atenção' });

  const decisoes = (resumoParaIa?.sinais_manuais?.decisoes || []);
  if (decisoes.length > 0) insights.push({ titulo: 'Decisões pendentes', descricao: `${decisoes.length} comentário${decisoes.length > 1 ? 's' : ''} sinaliza${decisoes.length > 1 ? 'm' : ''} necessidade de decisão ou alinhamento.`, sinal: 'decisão' });

  const backlog = (resumoParaIa?.backlog_executivo || []);
  if (backlog.length > 0) insights.push({ titulo: 'Triagem pendente', descricao: `${backlog.length} demanda${backlog.length > 1 ? 's' : ''} em status Novo ou Em triagem aguardam definição de escopo e prioridade.`, sinal: 'continuidade' });

  const emEvolucao = (carteiraExecutiva?.em_evolucao || []);
  if (emEvolucao.length > 0) insights.push({ titulo: 'Frentes em evolução', descricao: `${emEvolucao.length} demanda${emEvolucao.length > 1 ? 's' : ''} com movimentação relevante no período.`, sinal: 'eficiência' });

  return insights;
}

function parseComments(rawText) {
  if (!rawText) return [];
  const lines = rawText.split('\n');
  const focusIndex = lines.findIndex(l => l.includes('Comentários do foco:'));
  if (focusIndex === -1) return [];

  const comments = [];
  for (let i = focusIndex + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    // Match pattern "- HH:MM — texto" or "- HH:MM - texto"
    const match = line.match(/^-\s*(\d{2}:\d{2})\s*[-—]\s*(.*)$/);
    if (match) {
      comments.push({
        horario: match[1],
        texto: match[2].trim()
      });
    }
  }
  return comments;
}

function calculateMinutesInRange(start, end, rangeStart, rangeEnd) {
  if (!start) return 0;
  
  const s = new Date(start).getTime();
  const e = end ? new Date(end).getTime() : Date.now();
  const rs = new Date(rangeStart).getTime();
  const re = new Date(rangeEnd).getTime();

  const effectiveStart = Math.max(s, rs);
  const effectiveEnd = Math.min(e, re);

  if (effectiveEnd <= effectiveStart) return 0;

  const diffMs = effectiveEnd - effectiveStart;
  return Math.ceil(diffMs / 60000);
}

/**
 * ============================================================
 * SHAREPOINT REST - COMENTÁRIOS NATIVOS DO ITEM
 * ============================================================
 */

function getSharePointSiteUrl() {
  const { siteUrl } = getSharePointConfig();
  return String(siteUrl || '').trim().replace(/\/+$/, '');
}

function getSharePointOrigin() {
  const siteUrl = getSharePointSiteUrl();

  if (!siteUrl) {
    return '';
  }

  try {
    return new URL(siteUrl).origin;
  } catch (_error) {
    return '';
  }
}

function getMsalClient() {
  const msalConfig = {
    auth: {
      clientId: process.env.CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
      clientSecret: process.env.CLIENT_SECRET
    }
  };

  return new ConfidentialClientApplication(msalConfig);
}

async function getSharePointRestAccessToken() {
  const origin = getSharePointOrigin();

  if (!origin) {
    throw new Error(
      'SHAREPOINT_SITE_URL inválido. Não foi possível descobrir o origin do SharePoint.'
    );
  }

  const cca = getMsalClient();

  const tokenResult = await cca.acquireTokenByClientCredential({
    scopes: [`${origin}/.default`]
  });

  const token = tokenResult?.accessToken;

  if (!token) {
    throw new Error('Não foi possível obter token REST para SharePoint.');
  }

  return token;
}

async function readResponseSafely(response) {
  const rawText = await response.text();

  if (!rawText) {
    return null;
  }

  try {
    return JSON.parse(rawText);
  } catch (_error) {
    return rawText;
  }
}

function normalizeGuidForRest(value) {
  return String(value || '')
    .trim()
    .replace(/^\{/, '')
    .replace(/\}$/, '');
}

async function getSharePointRequestDigest(accessToken) {
  const siteUrl = getSharePointSiteUrl();

  const response = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata'
    }
  });

  const data = await readResponseSafely(response);

  if (!response.ok) {
    throw new Error(
      `Erro ao obter RequestDigest do SharePoint (${response.status}): ${JSON.stringify(data)}`
    );
  }

  return (
    data?.FormDigestValue ||
    data?.d?.GetContextWebInformation?.FormDigestValue ||
    data?.GetContextWebInformation?.FormDigestValue ||
    null
  );
}

async function addSharePointListItemComment({ listId, itemId, text }) {
  const { nativeCommentsEnabled } = getSharePointConfig();

  if (!nativeCommentsEnabled) {
    return {
      ok: false,
      skipped: true,
      message: 'Comentários nativos desabilitados por configuração.'
    };
  }

  const cleanText = sanitizeText(text);

  if (!cleanText) {
    return {
      ok: false,
      skipped: true,
      message: 'Comentário vazio. Nada foi enviado ao SharePoint.'
    };
  }

  const safeListId = normalizeGuidForRest(listId);
  const numericItemId = toNumberOrNull(itemId);

  if (!safeListId) {
    throw new Error('listId é obrigatório para comentar no item.');
  }

  if (!numericItemId) {
    throw new Error('itemId inválido para comentar no item.');
  }

  const siteUrl = getSharePointSiteUrl();
  const accessToken = await getSharePointRestAccessToken();

  let requestDigest = null;

  try {
    requestDigest = await getSharePointRequestDigest(accessToken);
  } catch (digestError) {
    console.warn(
      'Aviso: não foi possível obter RequestDigest. Tentando comentar apenas com Bearer token.',
      digestError?.message || digestError
    );
  }

  const url =
    `${siteUrl}/_api/web/lists(guid'${safeListId}')` +
    `/items(${numericItemId})/Comments()`;

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata'
  };

  if (requestDigest) {
    headers['X-RequestDigest'] = requestDigest;
  }

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      text: cleanText
    })
  });

  const data = await readResponseSafely(response);

  if (!response.ok) {
    throw new Error(
      `Erro ao criar comentário nativo no SharePoint (${response.status}): ${JSON.stringify(data)}`
    );
  }

  return {
    ok: true,
    itemId: numericItemId,
    listId: safeListId,
    comment: data
  };
}

async function getSharePointListItemComments({ listId, itemId }) {
  const safeListId = normalizeGuidForRest(listId);
  const numericItemId = toNumberOrNull(itemId);

  if (!safeListId) {
    throw new Error('listId é obrigatório para listar comentários.');
  }

  if (!numericItemId) {
    throw new Error('itemId inválido para listar comentários.');
  }

  const siteUrl = getSharePointSiteUrl();
  const accessToken = await getSharePointRestAccessToken();

  const url =
    `${siteUrl}/_api/web/lists(guid'${safeListId}')` +
    `/items(${numericItemId})/Comments()`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=nometadata'
    }
  });

  const data = await readResponseSafely(response);

  if (!response.ok) {
    throw new Error(
      `Erro ao listar comentários nativos no SharePoint (${response.status}): ${JSON.stringify(data)}`
    );
  }

  return data?.value || data?.d?.results || data || [];
}

/**
 * ============================================================
 * MAPA E-MAIL -> LOOKUPID DO SHAREPOINT
 * ============================================================
 */

function getUserLookupMap() {
  const raw = String(process.env.SHAREPOINT_USER_LOOKUP_MAP || '').trim();

  if (!raw) {
    return {};
  }

  try {
    const parsed = JSON.parse(raw);

    return Object.fromEntries(
      Object.entries(parsed).map(([email, lookupId]) => [
        normalizeEmail(email),
        toNumberOrNull(lookupId)
      ])
    );
  } catch (_error) {
    console.warn(
      'Aviso: SHAREPOINT_USER_LOOKUP_MAP inválido no .env. Use JSON válido. Exemplo: {"usuario@contaudi.com.br":10}'
    );

    return {};
  }
}

function getUserLookupIdByEmail(email) {
  const normalizedEmail = normalizeEmail(email);

  if (!normalizedEmail) {
    return null;
  }

  const map = getUserLookupMap();
  return map[normalizedEmail] || null;
}

function resolvePersonLookupId({ explicitLookupId, email }) {
  const direct = toNumberOrNull(explicitLookupId);

  if (direct) {
    return direct;
  }

  return getUserLookupIdByEmail(email);
}

function getDefaultUserEmail() {
  return normalizeEmail(
    process.env.FOCUSTRACK_DEFAULT_USER_EMAIL ||
    process.env.DEFAULT_USER_EMAIL ||
    Object.keys(getUserLookupMap())[0] ||
    ''
  );
}

function getDefaultUserLookupId() {
  const explicitDefault = toNumberOrNull(
    process.env.FOCUSTRACK_DEFAULT_USER_LOOKUP_ID ||
    process.env.DEFAULT_USER_LOOKUP_ID
  );

  if (explicitDefault) {
    return explicitDefault;
  }

  const defaultEmail = getDefaultUserEmail();

  if (defaultEmail) {
    return getUserLookupIdByEmail(defaultEmail);
  }

  return null;
}

/**
 * ============================================================
 * HELPERS DE FOCO / COMENTÁRIO
 * ============================================================
 */

function normalizeFocusComments(input) {
  if (!input) {
    return [];
  }

  if (Array.isArray(input)) {
    return input
      .map(comment => {
        if (typeof comment === 'string') {
          return {
            at: null,
            text: sanitizeText(comment)
          };
        }

        return {
          at: comment?.at || comment?.createdAt || comment?.time || null,
          text: sanitizeText(comment?.text || comment?.comment || comment?.message)
        };
      })
      .filter(comment => comment.text);
  }

  const text = sanitizeText(input);

  if (!text) {
    return [];
  }

  return text
    .split('\n')
    .map(line => sanitizeText(line))
    .filter(Boolean)
    .map(line => ({
      at: null,
      text: line
    }));
}

function getResponsibleDisplayNameFromRequest(req) {
  return (
    sanitizeText(req.body?.responsibleName) ||
    sanitizeText(req.body?.currentUser?.name) ||
    sanitizeText(req.body?.user?.name) ||
    sanitizeText(req.body?.responsibleEmail) ||
    sanitizeText(req.body?.currentUser?.email) ||
    sanitizeText(req.body?.user?.email) ||
    sanitizeText(getDefaultUserEmail()) ||
    'Usuário FocusTrack'
  );
}

function getTicketIdFromTimeEntryFields(fields) {
  return (
    toNumberOrNull(fields?.DemandaLookupId) ||
    toNumberOrNull(fields?.DemandaId) ||
    toNumberOrNull(fields?.Demanda) ||
    null
  );
}

function buildFocusNativeComment({
  statusLabel,
  startAt,
  endAt,
  minutes,
  responsibleName,
  focusObjective,
  focusComments,
  rawCommentText
}) {
  const comments = normalizeFocusComments(
    focusComments ||
    rawCommentText ||
    ''
  );

  const lines = [];

  lines.push(`[FocusTrack] Apontamento ${statusLabel || 'registrado'}`);
  lines.push('');
  lines.push(`Entrada: ${formatDateTimeBR(startAt) || '-'}`);
  lines.push(`Saída: ${formatDateTimeBR(endAt) || '-'}`);
  lines.push(`Duração: ${minutes || 0} min`);
  lines.push(`Responsável: ${responsibleName || 'Usuário FocusTrack'}`);



  if (comments.length > 0) {
    lines.push('');
    lines.push('Comentários do foco:');

    comments.forEach(comment => {
      const timePrefix = comment.at ? `${formatTimeBR(comment.at)} — ` : '';
      lines.push(`- ${timePrefix}${comment.text}`);
    });
  }

  return lines.join('\n');
}

/**
 * ============================================================
 * HELPERS DE ENVIO DE COMENTÁRIOS DO FOCUSTRACK
 * ============================================================
 *
 * Estratégia:
 * - Comentários são enviados diretamente como comentários nativos na demanda.
 * - Não usamos ComentarioCentralDemandas como canal principal no pause/finish,
 *   para evitar sobrescrita por chamadas posteriores.
 * - Comentários grandes são divididos em partes.
 * - Se houver comentários no payload e o envio falhar, a rota retorna erro,
 *   para o frontend não limpar os comentários locais.
 */

function getFocusCommentMaxChars() {
  const value = Number(
    process.env.FOCUSTRACK_COMMENT_MAX_CHARS ||
    process.env.CENTRAL_COMMENT_MAX_CHARS ||
    2000
  );

  if (!Number.isFinite(value) || value <= 500) {
    return 2000;
  }

  return Math.min(value, 10000);
}

function shouldBlockOnFocusCommentFailure() {
  return String(
    process.env.FOCUSTRACK_BLOCK_ON_COMMENT_FAILURE || 'true'
  ).toLowerCase() !== 'false';
}

function countChars(text) {
  return Array.from(String(text || '')).length;
}

function sliceByCharCount(text, maxChars) {
  const chars = Array.from(String(text || ''));
  const parts = [];

  for (let i = 0; i < chars.length; i += maxChars) {
    parts.push(chars.slice(i, i + maxChars).join(''));
  }

  return parts.filter(part => sanitizeText(part));
}

function splitTextByCharLimit(text, limit) {
  const cleanText = String(text || '').trim();

  if (!cleanText) {
    return [];
  }

  if (countChars(cleanText) <= limit) {
    return [cleanText];
  }

  const parts = [];
  const lines = cleanText.split('\n');

  let current = '';

  for (const line of lines) {
    const candidate = current ? `${current}\n${line}` : line;

    if (countChars(candidate) <= limit) {
      current = candidate;
      continue;
    }

    if (current) {
      parts.push(current);
      current = '';
    }

    if (countChars(line) <= limit) {
      current = line;
    } else {
      const sliced = sliceByCharCount(line, limit);

      for (const slice of sliced) {
        parts.push(slice);
      }
    }
  }

  if (current) {
    parts.push(current);
  }

  return parts.filter(part => sanitizeText(part));
}

function getRawFocusCommentTextFromRequest(req) {
  return sanitizeText(
    req.body?.focusCommentText ||
    req.body?.commentText ||
    req.body?.ticketCommentText ||
    ''
  );
}

function getFocusCommentsFromRequest(req) {
  const comments = normalizeFocusComments(req.body?.focusComments);
  const rawText = getRawFocusCommentTextFromRequest(req);

  if (comments.length > 0) {
    return comments;
  }

  if (rawText) {
    return normalizeFocusComments(rawText);
  }

  return [];
}

function hasFocusCommentsPayload(req) {
  return getFocusCommentsFromRequest(req).length > 0;
}

function hasFocusObjectivePayload(req) {
  return Boolean(
    sanitizeText(req.body?.focusObjective || req.body?.objective)
  );
}

function shouldPostFocusCommentFromRequest(req) {
  if (req.body?.postCommentToTicket === false) return false;
  if (req.body?.skipTicketComment === true) return false;

  /**
   * Regra importante:
   * Por padrão, só cria comentário nativo se houver comentário real.
   *
   * Isso evita o erro atual:
   * - pause envia comentários
   * - finish vem sem comentários, mas com objetivo
   * - finish cria comentário só com objetivo
   */
  if (hasFocusCommentsPayload(req)) return true;

  /**
   * Se algum fluxo futuro quiser forçar comentário só com objetivo,
   * pode enviar forceTicketComment: true.
   */


  return false;
}

async function addFocusCommentToTicketInChunks({
  ticketId,
  statusLabel,
  startAt,
  endAt,
  minutes,
  req
}) {
  const { ticketsListId } = getSharePointConfig();

  if (!ticketId) {
    return {
      ok: false,
      skipped: true,
      message: 'Sem ticketId. Comentário nativo não criado.'
    };
  }

  if (!shouldPostFocusCommentFromRequest(req)) {
    return {
      ok: false,
      skipped: true,
      message: 'Payload sem comentários pendentes. Comentário nativo não criado.'
    };
  }

  const responsibleName = getResponsibleDisplayNameFromRequest(req);

  const fullText = buildFocusNativeComment({
    statusLabel,
    startAt,
    endAt,
    minutes,
    responsibleName,
    focusObjective: req.body?.focusObjective || req.body?.objective,
    focusComments: req.body?.focusComments,
    rawCommentText: getRawFocusCommentTextFromRequest(req)
  });

  const maxChars = getFocusCommentMaxChars();

  /**
   * Reserva espaço para cabeçalho de parte.
   */
  const chunkLimit = Math.max(500, maxChars - 180);
  const parts = splitTextByCharLimit(fullText, chunkLimit);

  if (!parts.length) {
    return {
      ok: false,
      skipped: true,
      message: 'Comentário final ficou vazio após normalização.'
    };
  }

  const results = [];

  for (let index = 0; index < parts.length; index += 1) {
    const part = parts[index];

    const partText =
      parts.length > 1
        ? `[FocusTrack] ${statusLabel || 'registro'} — Parte ${index + 1}/${parts.length}\n\n${part}`
        : part;

    const result = await addSharePointListItemComment({
      listId: ticketsListId,
      itemId: ticketId,
      text: partText
    });

    results.push(result);
  }

  return {
    ok: true,
    ticketId,
    chunks: parts.length,
    maxChars,
    results
  };
}



async function tryAddFocusCommentToTicket({
  ticketId,
  statusLabel,
  startAt,
  endAt,
  minutes,
  req
}) {
  const { ticketsListId } = getSharePointConfig();

  const shouldPostComment =
    req.body?.postCommentToTicket !== false &&
    req.body?.skipTicketComment !== true;

  if (!shouldPostComment) {
    return {
      ok: false,
      skipped: true,
      message: 'Comentário nativo ignorado por opção do payload.'
    };
  }

  if (!ticketId) {
    return {
      ok: false,
      skipped: true,
      message: 'Sem DemandaLookupId no apontamento. Comentário nativo não enviado.'
    };
  }

  const hasFocusPayload =
    sanitizeText(req.body?.focusCommentText) ||
    sanitizeText(req.body?.commentText) ||
    sanitizeText(req.body?.ticketCommentText) ||
    Array.isArray(req.body?.focusComments);

  if (!hasFocusPayload && req.body?.forceTicketComment !== true) {
    return {
      ok: false,
      skipped: true,
      message:
        'Nenhum focusComments/focusCommentText enviado. Comentário nativo não criado.'

    };
  }

  const responsibleName = getResponsibleDisplayNameFromRequest(req);

  const text = buildFocusNativeComment({
    statusLabel,
    startAt,
    endAt,
    minutes,
    responsibleName,
    focusObjective: req.body?.focusObjective || req.body?.objective,
    focusComments: req.body?.focusComments,
    rawCommentText:
      req.body?.focusCommentText ||
      req.body?.commentText ||
      req.body?.ticketCommentText ||
      ''
  });

  return addSharePointListItemComment({
    listId: ticketsListId,
    itemId: ticketId,
    text
  });
}

/**
 * ============================================================
 * MAPEAMENTO DE TICKETS/DEMANDAS
 * ============================================================
 */

function mapTicketItem(item) {
  const f = item.fields || {};

  return {
    id: String(item.id),
    number: String(item.id),
    title: f.Title || '',
    description: f.Descri_x00e7__x00e3_o || '',
    status: f.Status || '',
    type: f.TipodeChamado || '',
    department: f.Departamento || '',
    origin: f.Origem || '',

    createdAt:
      f.DataCria_x00e7__x00e3_o ||
      f.Created ||
      item.createdDateTime ||
      null,

    startPlannedAt: f.InicioPlanejado || null,
    startedAt: f.InicioPlanejado || null,

    endPlannedAt: f.FimPlanejado || null,

    finishedAt: f.DataConclus_x00e3_o || null,
    concludedAt: f.DataConclus_x00e3_o || null,

    timeSpentMinutes: toSafeNumber(f.TempoGasto),

    requesterId:
      f.SolicitanteLookupId ||
      f.SolicitanteId ||
      f.Solicitante ||
      null,

    webUrl: item.webUrl || ''
  };
}

/**
 * ============================================================
 * ROTAS BÁSICAS
 * ============================================================
 */

app.get('/health', (_req, res) => {
  res.json({
    ok: true,
    service: 'FocusTrack SharePoint Backend'
  });
});

app.get('/api/sharepoint/token-info', async (_req, res) => {
  try {
    const result = await getTokenDebugInfo();
    res.json(result);
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error?.message || 'Erro ao obter informações do token.'
    });
  }
});

app.post('/api/sharepoint/test-connection', async (req, res) => {
  try {
    const result = await testConnection(req.body || {});
    res.status(result.ok ? 200 : 400).json(result);
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error?.message || 'Erro interno no backend.'
    });
  }
});

app.post('/api/sharepoint/site', async (req, res) => {
  try {
    const siteUrl = String(req.body?.siteUrl || '').trim();
    const result = await getSite(siteUrl);

    res.json({
      ok: true,
      site: result
    });
  } catch (error) {
    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao obter site.'
    });
  }
});

app.post('/api/sharepoint/list', async (req, res) => {
  try {
    const siteId = String(req.body?.siteId || '').trim();
    const listId = String(req.body?.listId || '').trim();

    const result = await getListById(siteId, listId);

    res.json({
      ok: true,
      list: result
    });
  } catch (error) {
    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao obter lista.'
    });
  }
});

app.post('/api/sharepoint/list-items', async (req, res) => {
  try {
    const siteId = String(req.body?.siteId || '').trim();
    const listId = String(req.body?.listId || '').trim();
    const top = Number(req.body?.top || 10);

    const result = await getListItems(siteId, listId, top);

    res.json({
      ok: true,
      items: result
    });
  } catch (error) {
    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao obter itens da lista.'
    });
  }
});

app.post('/api/sharepoint/list-item', async (req, res) => {
  try {
    const siteId = String(req.body?.siteId || '').trim();
    const listId = String(req.body?.listId || '').trim();
    const fields = req.body?.fields || {};

    const result = await createListItem(siteId, listId, fields);

    res.status(201).json({
      ok: true,
      item: result
    });
  } catch (error) {
    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao criar item.'
    });
  }
});

app.patch('/api/sharepoint/list-item/:itemId', async (req, res) => {
  try {
    const siteId = String(req.body?.siteId || '').trim();
    const listId = String(req.body?.listId || '').trim();
    const itemId = String(req.params?.itemId || '').trim();
    const fields = req.body?.fields || {};

    const result = await updateListItem(siteId, listId, itemId, fields);

    res.json(result);
  } catch (error) {
    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao atualizar item.'
    });
  }
});

/**
 * Lista colunas da lista.
 */
app.post('/api/sharepoint/list-columns', async (req, res) => {
  try {
    const siteId = String(req.body?.siteId || '').trim();
    const listId = String(req.body?.listId || '').trim();

    if (!siteId || !listId) {
      return res.status(400).json({
        ok: false,
        message: 'siteId e listId são obrigatórios.'
      });
    }

    const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
    const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

    const cca = getMsalClient();

    const tokenResult = await cca.acquireTokenByClientCredential({
      scopes: [GRAPH_SCOPE]
    });

    const token = tokenResult?.accessToken;

    if (!token) {
      return res.status(500).json({
        ok: false,
        message: 'Não foi possível obter token de acesso.'
      });
    }

    const response = await fetch(
      `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json'
        }
      }
    );

    const data = await response.json();

    if (!response.ok) {
      return res.status(400).json({
        ok: false,
        message: 'Erro ao obter colunas da lista.',
        details: data
      });
    }

    res.json({
      ok: true,
      columns: (data.value || []).map(col => ({
        id: col.id,
        displayName: col.displayName,
        name: col.name,
        lookupListId: col.lookup?.listId || null,
        lookupColumnName: col.lookup?.columnName || null,
        type: Object.keys(col).filter(k =>
          [
            'text',
            'number',
            'choice',
            'lookup',
            'dateTime',
            'personOrGroup',
            'boolean',
            'currency'
          ].includes(k)
        )
      }))
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error?.message || 'Erro ao listar colunas.'
    });
  }
});

app.get('/api/sharepoint/user-lookup-map', (_req, res) => {
  const map = getUserLookupMap();

  res.json({
    ok: true,
    count: Object.keys(map).length,
    defaultUser: {
      email: getDefaultUserEmail() || null,
      lookupId: getDefaultUserLookupId() || null
    },
    map
  });
});

app.post('/api/sharepoint/list-item/get/:itemId', async (req, res) => {
  try {
    const siteId = String(req.body?.siteId || '').trim();
    const listId = String(req.body?.listId || '').trim();
    const itemId = String(req.params?.itemId || '').trim();

    if (!siteId || !listId || !itemId) {
      return res.status(400).json({
        ok: false,
        message: 'siteId, listId e itemId são obrigatórios.'
      });
    }

    const item = await getListItemById(siteId, listId, itemId);

    res.json({
      ok: true,
      item,
      fields: item?.fields || {}
    });
  } catch (error) {
    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao obter item.'
    });
  }
});

/**
 * ============================================================
 * COMENTÁRIOS NATIVOS SHAREPOINT
 * ============================================================
 */

app.post('/api/tickets/:ticketId/comments', async (req, res) => {
  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const { ticketsListId } = getSharePointConfig();
    const ticketId = String(req.params.ticketId || '').trim();

    const text =
      sanitizeText(req.body?.text) ||
      sanitizeText(req.body?.comment) ||
      sanitizeText(req.body?.message);

    if (!ticketId) {
      return res.status(400).json({
        ok: false,
        message: 'ticketId é obrigatório.'
      });
    }

    if (!text) {
      return res.status(400).json({
        ok: false,
        message: 'Texto do comentário é obrigatório. Envie text, comment ou message.'
      });
    }

    const result = await addSharePointListItemComment({
      listId: ticketsListId,
      itemId: ticketId,
      text
    });

    res.status(201).json({
      ok: true,
      message: 'Comentário nativo criado no item da demanda.',
      result
    });
  } catch (error) {
    console.error('ERRO /api/tickets/:ticketId/comments =>', error);

    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao criar comentário nativo no item.'
    });
  }
});

app.get('/api/tickets/:ticketId/comments', async (req, res) => {
  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const { ticketsListId } = getSharePointConfig();
    const ticketId = String(req.params.ticketId || '').trim();

    if (!ticketId) {
      return res.status(400).json({
        ok: false,
        message: 'ticketId é obrigatório.'
      });
    }

    const comments = await getSharePointListItemComments({
      listId: ticketsListId,
      itemId: ticketId
    });

    res.json({
      ok: true,
      ticketId,
      count: Array.isArray(comments) ? comments.length : null,
      comments
    });
  } catch (error) {
    console.error('ERRO GET /api/tickets/:ticketId/comments =>', error);

    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao listar comentários nativos do item.'
    });
  }
});

/**
 * ============================================================
 * DEMANDAS / CHAMADOS
 * ============================================================
 */

app.get('/api/tickets', async (req, res) => {
  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const siteId = await resolveSiteId();
    const { ticketsListId } = getSharePointConfig();

    const result = await getListItems(siteId, ticketsListId, 100);
    const items = Array.isArray(result?.value) ? result.value : [];

    const search = String(req.query.search || '').trim().toLowerCase();
    const status = String(req.query.status || '').trim().toLowerCase();

    let tickets = items.map(mapTicketItem);

    if (search) {
      tickets = tickets.filter(ticket =>
        [
          ticket.number,
          ticket.title,
          ticket.description,
          ticket.status,
          ticket.type,
          ticket.department,
          ticket.origin
        ]
          .filter(Boolean)
          .some(value => String(value).toLowerCase().includes(search))
      );
    }

    if (status) {
      tickets = tickets.filter(
        ticket => String(ticket.status || '').toLowerCase() === status
      );
    }

    res.json({
      ok: true,
      count: tickets.length,
      tickets
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error?.message || 'Erro ao listar chamados.'
    });
  }
});

app.post('/api/tickets', async (req, res) => {
  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const siteId = await resolveSiteId();

    const {
      ticketsListId,
      ticketRequesterLookupField,
      ticketCreatedAtField
    } = getSharePointConfig();

    const requesterEmail = normalizeEmail(
      req.body?.requesterEmail ||
      req.body?.solicitanteEmail ||
      req.body?.requester?.email ||
      req.body?.solicitante?.email ||
      req.body?.currentUser?.email ||
      req.body?.user?.email ||
      req.body?.responsibleEmail ||
      getDefaultUserEmail()
    );

    const requesterLookupId =
      resolvePersonLookupId({
        explicitLookupId:
          req.body?.requesterLookupId ||
          req.body?.solicitanteLookupId ||
          req.body?.requester?.lookupId ||
          req.body?.solicitante?.lookupId ||
          req.body?.currentUser?.lookupId ||
          req.body?.user?.lookupId ||
          req.body?.responsibleLookupId,
        email: requesterEmail
      }) ||
      getDefaultUserLookupId();

    const createdAt = normalizeDateTimeOrNow(
      req.body?.createdAt ||
      req.body?.dataCriacao ||
      req.body?.dataCriacaoEm
    );

    const fields = {
      Title: req.body?.title || 'Novo chamado',
      Status: req.body?.status || 'Novo',
      TipodeChamado: req.body?.type || 'Chamado',
      Departamento: req.body?.department || null,
      Origem: req.body?.origin || 'FocusTrack',
      InicioPlanejado: req.body?.startPlannedAt || null,
      FimPlanejado: req.body?.endPlannedAt || null,
      Descri_x00e7__x00e3_o:
        req.body?.description ||
        req.body?.descricao ||
        req.body?.notes ||
        null
    };

    setDynamicField(fields, ticketCreatedAtField, createdAt);
    setDynamicField(fields, ticketRequesterLookupField, requesterLookupId);

    const cleanFields = removeNullishFields(fields);

    console.log('BODY /api/tickets =>', req.body);
    console.log('Solicitante/Data resolvidos =>', {
      requesterEmail,
      requesterLookupId,
      ticketRequesterLookupField,
      createdAt,
      ticketCreatedAtField,
      defaultUserEmail: getDefaultUserEmail(),
      defaultUserLookupId: getDefaultUserLookupId()
    });
    console.log('FIELDS enviados ao SharePoint /api/tickets =>', cleanFields);

    const created = await createListItem(siteId, ticketsListId, cleanFields);

    const warnings = [];

    if (!requesterLookupId) {
      warnings.push(
        'Solicitante não foi preenchido porque não foi possível resolver LookupId. Verifique SHAREPOINT_USER_LOOKUP_MAP ou FOCUSTRACK_DEFAULT_USER_LOOKUP_ID.'
      );
    }

    if (!createdAt) {
      warnings.push(
        'Data Criação não foi preenchida porque não foi possível gerar uma data válida.'
      );
    }

    let nativeCommentResult = null;

    const initialNativeComment =
      sanitizeText(req.body?.initialComment) ||
      sanitizeText(req.body?.nativeComment) ||
      sanitizeText(req.body?.ticketComment);

    if (initialNativeComment && created?.id) {
      try {
        nativeCommentResult = await addSharePointListItemComment({
          listId: ticketsListId,
          itemId: created.id,
          text: initialNativeComment
        });
      } catch (commentError) {
        warnings.push(
          `Chamado criado, mas não foi possível criar comentário nativo: ${commentError?.message || commentError}`
        );
      }
    }

    res.status(201).json({
      ok: true,
      message: 'Chamado criado com sucesso.',
      warnings,
      requester: {
        email: requesterEmail || null,
        lookupId: requesterLookupId || null,
        field: ticketRequesterLookupField
      },
      createdAt: {
        value: createdAt || null,
        field: ticketCreatedAtField
      },
      nativeComment: nativeCommentResult,
      ticket: created
    });
  } catch (error) {
    console.error('ERRO /api/tickets =>', error);

    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao criar chamado.'
    });
  }
});

/**
 * ============================================================
 * COMENTÁRIO CENTRAL DEMANDAS
 * ============================================================
 *
 * Solução sem criar nova lista:
 * - O backend grava o comentário no próprio item de LancamentoTempo.
 * - O Power Automate lê ComentarioCentralDemandas + EnviarComentarioCentral.
 * - Não usamos comentário nativo direto no pause/finish porque o SharePoint
 *   retornou 401 "Unsupported app only token".
 * - Só gera comentário se houver Comentários do foco.
 * - Não gera comentário apenas com objetivo.
 * - Não sobrescreve comentário existente com texto sem "Comentários do foco".
 * - Se passar do limite configurado, retorna erro 413 para o frontend
 *   preservar os comentários locais.
 */

function getCentralCommentMaxChars() {
  const value = Number(
    process.env.FOCUSTRACK_COMMENT_MAX_CHARS ||
    process.env.CENTRAL_COMMENT_MAX_CHARS ||
    2000
  );

  if (!Number.isFinite(value) || value <= 500) {
    return 2000;
  }

  return Math.min(value, 10000);
}

function countCentralCommentChars(text) {
  return Array.from(String(text || '')).length;
}

function getCentralRawFocusCommentTextFromRequest(req) {
  return sanitizeText(
    req.body?.focusCommentText ||
    req.body?.commentText ||
    req.body?.ticketCommentText ||
    ''
  );
}

function getCentralFocusCommentsFromRequest(req) {
  const comments = normalizeFocusComments(req.body?.focusComments);
  const rawText = getCentralRawFocusCommentTextFromRequest(req);

  if (comments.length > 0) {
    return comments;
  }

  if (rawText) {
    return normalizeFocusComments(rawText);
  }

  return [];
}

function hasCentralFocusCommentsPayload(req) {
  return getCentralFocusCommentsFromRequest(req).length > 0;
}

function hasCentralFocusObjectivePayload(req) {
  return Boolean(
    sanitizeText(req.body?.focusObjective || req.body?.objective)
  );
}

function shouldCreateCentralComment(req) {
  if (req.body?.postCommentToTicket === false) return false;
  if (req.body?.skipTicketComment === true) return false;

  /**
   * Regra principal:
   * Só cria comentário para o Power Automate se houver comentário real.
   * Isso impede o bug do finish criar comentário apenas com objetivo.
   */
  if (hasCentralFocusCommentsPayload(req)) return true;

  /**
   * Só permite comentário sem comentários do foco se for explicitamente forçado.
   * Uso normal do FocusTrack NÃO deve depender disso.
   */


  return false;
}

function validateCentralCommentLimit(text) {
  const maxChars = getCentralCommentMaxChars();
  const totalChars = countCentralCommentChars(text);

  if (totalChars > maxChars) {
    return {
      ok: false,
      maxChars,
      totalChars,
      message:
        `Comentário do FocusTrack possui ${totalChars} caracteres, ` +
        `mas o limite configurado é ${maxChars}. ` +
        `Reduza o texto ou pause o foco antes de acumular comentários muito longos.`
    };
  }

  return {
    ok: true,
    maxChars,
    totalChars
  };
}

function hasFocusCommentsInText(text) {
  return String(text || '').includes('Comentários do foco:');
}

function shouldPreserveExistingCentralComment(existingComment, newComment) {
  return (
    hasFocusCommentsInText(existingComment) &&
    !hasFocusCommentsInText(newComment)
  );
}

function buildFocusCentralComment({
  statusLabel,
  startAt,
  endAt,
  minutes,
  responsibleName,
  focusObjective,
  focusComments,
  rawCommentText
}) {
  const comments = normalizeFocusComments(
    focusComments ||
    rawCommentText ||
    ''
  );

  const lines = [];

  lines.push(`[FocusTrack] Apontamento ${statusLabel || 'registrado'}`);
  lines.push('');
  lines.push(`Entrada: ${formatDateTimeBR(startAt) || '-'}`);
  lines.push(`Saída: ${formatDateTimeBR(endAt) || '-'}`);
  lines.push(`Duração: ${minutes || 0} min`);
  lines.push(`Responsável: ${responsibleName || 'Usuário FocusTrack'}`);



  if (comments.length > 0) {
    lines.push('');
    lines.push('Comentários do foco:');

    comments.forEach(comment => {
      const timePrefix = comment.at ? `${formatTimeBR(comment.at)} — ` : '';
      lines.push(`- ${timePrefix}${comment.text}`);
    });
  }

  return lines.join('\n');
}

function applyCentralCommentFields({
  fields,
  req,
  existingFields,
  endAt,
  minutes,
  statusLabel
}) {
  if (!shouldCreateCentralComment(req)) {
    return null;
  }

  const {
    timeCentralCommentField,
    timeSendCentralCommentField,
    timeCentralCommentSentField,
    timeCentralCommentErrorField
  } = getSharePointConfig();

  const responsibleName = getResponsibleDisplayNameFromRequest(req);

  let text = buildFocusCentralComment({
    statusLabel,
    startAt: existingFields?.Inicio,
    endAt,
    minutes,
    responsibleName,
    focusObjective: req.body?.focusObjective || req.body?.objective,
    focusComments: req.body?.focusComments,
    rawCommentText: getCentralRawFocusCommentTextFromRequest(req)
  });

  const existingCentralComment = existingFields?.[timeCentralCommentField] || '';

  /**
   * Proteção contra sobrescrita destrutiva:
   * Se já havia comentário com "Comentários do foco:"
   * e o novo texto não tem esse bloco, preserva o existente.
   */
  if (shouldPreserveExistingCentralComment(existingCentralComment, text)) {
    console.warn(
      '[FocusTrack] Comentário central existente preservado: novo texto não contém Comentários do foco.'
    );

    text = existingCentralComment;
  }

  const validation = validateCentralCommentLimit(text);

  if (!validation.ok) {
    const error = new Error(validation.message);
    error.status = 413;
    error.code = 'FOCUS_COMMENT_TOO_LARGE';
    error.details = validation;
    throw error;
  }

  setDynamicField(fields, timeCentralCommentField, text);
  setDynamicField(fields, timeSendCentralCommentField, true);
  setDynamicField(fields, timeCentralCommentSentField, false);

  if (timeCentralCommentErrorField) {
    fields[timeCentralCommentErrorField] = null;
  }

  return {
    ok: true,
    mode: 'central-field',
    text,
    field: timeCentralCommentField,
    chars: validation.totalChars,
    maxChars: validation.maxChars
  };
}

/**
 * ============================================================
 * ATUALIZAÇÃO DE CAMPOS DA DEMANDA PELO FOCO
 * ============================================================
 */

async function sumTicketTimeSpentMinutes(siteId, ticketId) {
  const numericTicketId = toNumberOrNull(ticketId);

  if (!numericTicketId) {
    return 0;
  }

  const { timeEntriesListId } = getSharePointConfig();

  const result = await getListItems(
    siteId,
    timeEntriesListId,
    getTimeEntriesSumTop()
  );

  const items = Array.isArray(result?.value) ? result.value : [];

  return items.reduce((total, item) => {
    const fields = item?.fields || {};
    const itemTicketId = getTicketIdFromTimeEntryFields(fields);

    if (Number(itemTicketId) !== Number(numericTicketId)) {
      return total;
    }

    return total + toSafeNumber(fields.Minutos);
  }, 0);
}

async function updateTicketTrackingOnStart({ siteId, ticketId, startAt }) {
  const numericTicketId = toNumberOrNull(ticketId);

  if (!numericTicketId) {
    return {
      ok: false,
      skipped: true,
      message: 'Sem ticketId válido para atualizar demanda no início.'
    };
  }

  const {
    ticketsListId,
    ticketStatusField,
    ticketStartedAtField,
    ticketStatusInProgressValue
  } = getSharePointConfig();

  const ticket = await getListItemById(
    siteId,
    ticketsListId,
    String(numericTicketId)
  );

  const fields = ticket?.fields || {};
  const updateFields = {};

  setDynamicField(
    updateFields,
    ticketStatusField,
    ticketStatusInProgressValue
  );

  if (!isFilledValue(fields?.[ticketStartedAtField])) {
    setDynamicField(updateFields, ticketStartedAtField, startAt);
  }

  const cleanFields = removeNullishFields(updateFields);

  if (Object.keys(cleanFields).length === 0) {
    return {
      ok: true,
      skipped: true,
      message: 'Nenhum campo de início precisou ser atualizado.'
    };
  }

  await updateListItem(
    siteId,
    ticketsListId,
    String(numericTicketId),
    cleanFields
  );

  return {
    ok: true,
    ticketId: numericTicketId,
    fields: cleanFields
  };
}

async function syncTicketTrackingAfterTimeEntryClosed({
  siteId,
  ticketId,
  endAt,
  markFinished
}) {
  const numericTicketId = toNumberOrNull(ticketId);

  if (!numericTicketId) {
    return {
      ok: false,
      skipped: true,
      message: 'Sem ticketId válido para sincronizar tempo gasto da demanda.'
    };
  }

  const {
    ticketsListId,
    ticketStatusField,
    ticketFinishedAtField,
    ticketTimeSpentField,
    ticketStatusInProgressValue,
    ticketStatusFinishedValue
  } = getSharePointConfig();

  const totalMinutes = await sumTicketTimeSpentMinutes(siteId, numericTicketId);

  const updateFields = {};

  setDynamicField(updateFields, ticketTimeSpentField, totalMinutes);

  if (markFinished) {
    setDynamicField(updateFields, ticketStatusField, ticketStatusFinishedValue);
    setDynamicField(updateFields, ticketFinishedAtField, endAt);
  } else {
    setDynamicField(updateFields, ticketStatusField, ticketStatusInProgressValue);
  }

  const cleanFields = removeNullishFields(updateFields);

  await updateListItem(
    siteId,
    ticketsListId,
    String(numericTicketId),
    cleanFields
  );

  return {
    ok: true,
    ticketId: numericTicketId,
    totalMinutes,
    fields: cleanFields
  };
}

/**
 * ============================================================
 * LANÇAMENTOS DE TEMPO
 * ============================================================
 */

app.post('/api/time-entries/start', async (req, res) => {
  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const siteId = await resolveSiteId();

    const {
      timeEntriesListId,
      timeResponsibleLookupField
    } = getSharePointConfig();

    const ticketId = Number(req.body?.ticketId || 0);
    const startAt = normalizeDateTimeOrNow(
      req.body?.startAt || new Date().toISOString()
    );

    const responsibleEmail = normalizeEmail(
      req.body?.responsibleEmail ||
      req.body?.currentUser?.email ||
      req.body?.user?.email ||
      getDefaultUserEmail()
    );

    const responsibleLookupId =
      resolvePersonLookupId({
        explicitLookupId:
          req.body?.responsibleLookupId ||
          req.body?.currentUser?.lookupId ||
          req.body?.user?.lookupId,
        email: responsibleEmail
      }) ||
      getDefaultUserLookupId();

    const fields = {
      Title: req.body?.title || `Apontamento #${ticketId || 'sem-demanda'}`,
      DemandaLookupId: ticketId || null,
      Inicio: startAt,
      Minutos: 0,
      Observacao: req.body?.notes || ''
    };

    setDynamicField(fields, timeResponsibleLookupField, responsibleLookupId);

    const cleanFields = removeNullishFields(fields);

    console.log('BODY /api/time-entries/start =>', req.body);
    console.log('Responsável resolvido =>', {
      responsibleEmail,
      responsibleLookupId,
      timeResponsibleLookupField,
      defaultUserEmail: getDefaultUserEmail(),
      defaultUserLookupId: getDefaultUserLookupId()
    });
    console.log('FIELDS enviados ao SharePoint /api/time-entries/start =>', cleanFields);

    const created = await createListItem(siteId, timeEntriesListId, cleanFields);

    const warnings = [];
    let ticketTrackingStartResult = null;

    if (!responsibleLookupId) {
      warnings.push(
        'Responsável não foi preenchido porque não foi possível resolver LookupId. Verifique SHAREPOINT_USER_LOOKUP_MAP ou FOCUSTRACK_DEFAULT_USER_LOOKUP_ID.'
      );
    }

    if (ticketId) {
      try {
        ticketTrackingStartResult = await updateTicketTrackingOnStart({
          siteId,
          ticketId,
          startAt
        });
      } catch (trackingError) {
        warnings.push(
          `Apontamento iniciado, mas não foi possível atualizar Status/Iniciado em da demanda: ${trackingError?.message || trackingError}`
        );
      }
    }

    res.status(201).json({
      ok: true,
      message: 'Apontamento iniciado com sucesso.',
      warnings,
      responsible: {
        email: responsibleEmail || null,
        lookupId: responsibleLookupId || null,
        field: timeResponsibleLookupField
      },
      ticketTracking: ticketTrackingStartResult,
      timeEntry: created
    });
  } catch (error) {
    console.error('ERRO /api/time-entries/start =>', error);

    res.status(400).json({
      ok: false,
      message: error?.message || 'Erro ao iniciar apontamento.'
    });
  }
});

app.post('/api/time-entries/:id/pause', async (req, res) => {
  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const siteId = await resolveSiteId();
    const { timeEntriesListId } = getSharePointConfig();
    const itemId = String(req.params.id || '').trim();

    const existing = await getListItemById(siteId, timeEntriesListId, itemId);
    const existingFields = existing?.fields || {};

    const endAt = normalizeDateTimeOrNow(
      req.body?.endAt || new Date().toISOString()
    );

    let minutos = req.body?.minutes;

    if (minutos === undefined || minutos === null) {
      minutos = calculateMinutes(existingFields.Inicio, endAt);
    }

    minutos = toSafeNumber(minutos);

    if (minutos <= 0) {
      minutos = calculateMinutes(existingFields.Inicio, endAt);
    }

    const updateFields = {
      Fim: endAt,
      Minutos: minutos,
      Observacao: req.body?.notes || existingFields.Observacao || ''
    };

    /**
     * Sem comentário nativo direto.
     * Aqui gravamos ComentarioCentralDemandas no LancamentoTempo
     * para o Power Automate processar.
     */
    const centralCommentPrepared = applyCentralCommentFields({
      fields: updateFields,
      req,
      existingFields,
      endAt,
      minutes: minutos,
      statusLabel: 'pausado'
    });

    console.log('BODY /api/time-entries/:id/pause =>', {
      itemId,
      hasFocusComments: hasCentralFocusCommentsPayload(req),
      skipTicketComment: req.body?.skipTicketComment === true,
      centralCommentPrepared: Boolean(centralCommentPrepared)
    });

    console.log(
      'FIELDS enviados ao SharePoint /api/time-entries/:id/pause =>',
      updateFields
    );

    await updateListItem(
      siteId,
      timeEntriesListId,
      itemId,
      removeNullishFields(updateFields)
    );

    const updated = await getListItemById(siteId, timeEntriesListId, itemId);

    const warnings = [];
    let ticketTrackingResult = null;

    if (!centralCommentPrepared && hasCentralFocusCommentsPayload(req)) {
      warnings.push(
        'Havia comentários no payload, mas o comentário central não foi preparado. Verifique skipTicketComment/postCommentToTicket.'
      );
    }

    const ticketId = getTicketIdFromTimeEntryFields(
      updated?.fields || existingFields
    );

    if (ticketId) {
      try {
        ticketTrackingResult = await syncTicketTrackingAfterTimeEntryClosed({
          siteId,
          ticketId,
          endAt,
          markFinished: false
        });
      } catch (trackingError) {
        warnings.push(
          `Apontamento pausado, mas não foi possível atualizar Tempo Gasto/Status da demanda: ${trackingError?.message || trackingError}`
        );
      }
    }

    res.json({
      ok: true,
      message: 'Apontamento pausado com sucesso.',
      centralCommentPrepared,
      nativeComment: null,
      ticketTracking: ticketTrackingResult,
      warnings,
      timeEntry: updated
    });
  } catch (error) {
    console.error('ERRO /api/time-entries/:id/pause =>', error);

    res.status(error?.status || 400).json({
      ok: false,
      code: error?.code || null,
      message: error?.message || 'Erro ao pausar apontamento.',
      details: error?.details || null
    });
  }
});

app.post('/api/time-entries/:id/finish', async (req, res) => {
  const itemId = String(req.params.id || '').trim();

  if (!itemId) {
    return res.status(400).json({
      ok: false,
      message: 'ID do apontamento não informado.'
    });
  }

  if (finishLocks.has(itemId)) {
    return res.status(409).json({
      ok: false,
      code: 'FINISH_ALREADY_RUNNING',
      message: 'Este apontamento já está sendo concluído. Aguarde alguns segundos.'
    });
  }

  finishLocks.add(itemId);

  try {
    const configCheck = validateRequiredConfig();

    if (!configCheck.ok) {
      return res.status(500).json({
        ok: false,
        message: `Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`
      });
    }

    const siteId = await resolveSiteId();
    const { timeEntriesListId } = getSharePointConfig();

    const existing = await getListItemById(siteId, timeEntriesListId, itemId);
    const existingFields = existing?.fields || {};

    const endAt = normalizeDateTimeOrNow(
      req.body?.endAt || new Date().toISOString()
    );

    let minutos = req.body?.minutes;

    if (minutos === undefined || minutos === null) {
      minutos = calculateMinutes(existingFields.Inicio, endAt);
    }

    minutos = toSafeNumber(minutos);

    if (minutos <= 0) {
      minutos = calculateMinutes(existingFields.Inicio, endAt);
    }

    const updateFields = {
      Fim: endAt,
      Minutos: minutos,
      Observacao: req.body?.notes || existingFields.Observacao || ''
    };

    /**
     * Sem comentário nativo direto.
     *
     * Regra:
     * - Se houver comentários pendentes, grava ComentarioCentralDemandas.
     * - Se não houver comentários, não cria comentário apenas com objetivo.
     */
    const centralCommentPrepared = applyCentralCommentFields({
      fields: updateFields,
      req,
      existingFields,
      endAt,
      minutes: minutos,
      statusLabel: 'concluído'
    });

    console.log('BODY /api/time-entries/:id/finish =>', {
      itemId,
      hasFocusComments: hasCentralFocusCommentsPayload(req),
      skipTicketComment: req.body?.skipTicketComment === true,
      centralCommentPrepared: Boolean(centralCommentPrepared)
    });

    console.log(
      'FIELDS enviados ao SharePoint /api/time-entries/:id/finish =>',
      updateFields
    );

    await updateListItem(
      siteId,
      timeEntriesListId,
      itemId,
      removeNullishFields(updateFields)
    );

    const updated = await getListItemById(siteId, timeEntriesListId, itemId);

    const warnings = [];
    let ticketTrackingResult = null;

    if (!centralCommentPrepared && hasCentralFocusCommentsPayload(req)) {
      warnings.push(
        'Havia comentários no payload, mas o comentário central não foi preparado. Verifique skipTicketComment/postCommentToTicket.'
      );
    }

    const ticketId = getTicketIdFromTimeEntryFields(
      updated?.fields || existingFields
    );

    if (ticketId) {
      try {
        ticketTrackingResult = await syncTicketTrackingAfterTimeEntryClosed({
          siteId,
          ticketId,
          endAt,
          markFinished: true
        });
      } catch (trackingError) {
        warnings.push(
          `Apontamento concluído, mas não foi possível atualizar Tempo Gasto/Data Conclusão/Status da demanda: ${trackingError?.message || trackingError}`
        );
      }
    }

    res.json({
      ok: true,
      message: 'Apontamento concluído com sucesso.',
      centralCommentPrepared,
      nativeComment: null,
      ticketTracking: ticketTrackingResult,
      warnings,
      timeEntry: updated
    });
  } catch (error) {
    console.error('ERRO /api/time-entries/:id/finish =>', error);

    res.status(error?.status || 400).json({
      ok: false,
      code: error?.code || null,
      message: error?.message || 'Erro ao concluir apontamento.',
      details: error?.details || null
    });
  } finally {
    finishLocks.delete(itemId);
  }
});

/**
 * ============================================================
 * LÓGICA DO RELATÓRIO SEMANAL (REUTILIZÁVEL)
 * ============================================================
 */

async function getWeeklyReportData(startStr) {
  const configCheck = validateRequiredConfig();
  if (!configCheck.ok) {
    const err = new Error(`Variáveis ausentes no .env: ${configCheck.missing.join(', ')}`);
    err.status = 500;
    throw err;
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(startStr)) {
    const err = new Error('Formato inválido para "start". Use YYYY-MM-DD.');
    err.status = 400;
    throw err;
  }

  const [year, month, day] = startStr.split('-').map(Number);
  const weekStartUtc = new Date(Date.UTC(year, month - 1, day, 4, 0, 0));
  const weekEndUtc = new Date(weekStartUtc.getTime() + (7 * 24 * 60 * 60 * 1000) - 1000);
  const now = new Date();

  const siteId = await resolveSiteId();
  const config = getSharePointConfig();

  const timeEntries = await getListItemsByQuery(siteId, config.timeEntriesListId, {
    filter: `fields/Inicio ge '${weekStartUtc.toISOString()}' and fields/Inicio le '${weekEndUtc.toISOString()}'`,
    expand: 'fields',
    top: 999
  });

  const sortedEntries = timeEntries.sort((a, b) => 
    new Date(a.fields?.Inicio).getTime() - new Date(b.fields?.Inicio).getTime()
  );

  const ticketIds = new Set();
  const apontamentosSemDemanda = [];
  const apontamentosPorDemanda = {};
  const alertas = [];

  sortedEntries.forEach(item => {
    const f = item.fields || {};
    const ticketId = getTicketIdFromTimeEntryFields(f);
    const minutosOriginal = toSafeNumber(f.Minutos);
    const minutosNaSemana = calculateMinutesInRange(f.Inicio, f.Fim, weekStartUtc, weekEndUtc);
    const atravessaLimite = (new Date(f.Inicio).getTime() < weekStartUtc.getTime()) || 
                            (f.Fim && new Date(f.Fim).getTime() > weekEndUtc.getTime());

    const entryData = {
      id: String(item.id),
      inicio: f.Inicio || null,
      fim: f.Fim || null,
      minutos: minutosOriginal,
      minutos_original: minutosOriginal,
      minutos_na_semana: minutosNaSemana,
      atravessa_limite_semana: atravessaLimite,
      responsavel: f[config.timeResponsibleLookupField] || '',
      observacao: f.Observacao || '',
      comentario_enviado: f[config.timeCentralCommentSentField] === true || f[config.timeCentralCommentSentField] === 'true',
      comentario_raw: f[config.timeCentralCommentField] || null,
      comentarios: parseComments(f[config.timeCentralCommentField])
    };

    if (ticketId) {
      ticketIds.add(String(ticketId));
      if (!apontamentosPorDemanda[ticketId]) apontamentosPorDemanda[ticketId] = [];
      apontamentosPorDemanda[ticketId].push(entryData);
    } else {
      apontamentosSemDemanda.push(entryData);
    }
  });

  const demandasMap = {};
  for (const tid of ticketIds) {
    try {
      const ticketItem = await getListItemById(siteId, config.ticketsListId, tid);
      const tf = ticketItem?.fields || {};
      demandasMap[tid] = {
        id: String(tid),
        titulo: tf.Title || '',
        status: tf.Status || '',
        tipo: tf.TipodeChamado || '',
        departamento: tf.Departamento || 'Sem departamento informado',
        origem: tf.Origem || '',
        descricao: tf.Descri_x00e7__x00e3_o || '',
        data_criacao: tf.DataCria_x00e7__x00e3_o || tf.Created || null,
        data_conclusao: tf.DataConclus_x00e7__x00e3_o || null,
        inicio_planejado: tf.InicioPlanejado || null,
        fim_planejado: tf.FimPlanejado || null,
        tempo_total_historico_min: toSafeNumber(tf.TempoGasto),
        tempo_total_historico_horas: formatMinutesToHours(tf.TempoGasto)
      };
    } catch (err) {
      console.warn(`Erro ao buscar demanda ${tid}:`, err.message);
      demandasMap[tid] = { id: tid, error: 'Não encontrada' };
    }
  }

  const finalDemandas = Object.entries(apontamentosPorDemanda).map(([tid, entries]) => {
    const d = demandasMap[tid] || { id: tid };
    const totalMinSemana = entries.reduce((sum, e) => sum + e.minutos_na_semana, 0);
    const comentariosConsolidados = entries.flatMap(e => e.comentarios.map(c => c.texto));
    
    // Status normalization
    const normalizedStatus = normalizeStatusLabel(d.status);
    const inProgress = isInProgressStatus(d.status);
    const backlog = isBacklogStatus(d.status);
    const finished = isFinishedStatus(d.status);

    // Date analysis
    const concluidaNestaSemana = d.data_conclusao && 
      (new Date(d.data_conclusao) >= weekStartUtc && new Date(d.data_conclusao) <= weekEndUtc);
    const criadaNestaSemana = d.data_criacao && 
      (new Date(d.data_criacao) >= weekStartUtc && new Date(d.data_criacao) <= weekEndUtc);

    // Deadline analysis
    const fimPlanejado = d.fim_planejado ? new Date(d.fim_planejado) : null;
    const dataConclusao = d.data_conclusao ? new Date(d.data_conclusao) : null;
    
    let deadlineStatus = 'Sem prazo';
    if (fimPlanejado) {
      if (dataConclusao) {
        deadlineStatus = dataConclusao <= fimPlanejado ? 'No prazo' : 'Com atraso';
      } else {
        deadlineStatus = fimPlanejado < now ? 'Atrasado' : 'No prazo (aberto)';
      }
    }

    // Tag analysis
    const allText = `${d.titulo} ${d.descricao} ${comentariosConsolidados.join(' ')}`;
    const { tags, economia } = extractTagsFromText(allText);

    return {
      ...d,
      tags,
      economia_estimada: economia,
      normalized_status: { in_progress: inProgress, backlog, finished },
      deadline_status: deadlineStatus,
      semana: {
        tempo_minutos: totalMinSemana,
        tempo_horas: formatMinutesToHours(totalMinSemana),
        concluida_nesta_semana: concluidaNestaSemana,
        criada_nesta_semana: criadaNestaSemana,
        quantidade_apontamentos: entries.length
      },
      comentarios_semana: comentariosConsolidados
    };
  });

  // Global metrics
  const totalMinutosSemana = finalDemandas.reduce((sum, d) => sum + d.semana.tempo_minutos, 0) + 
                             apontamentosSemDemanda.reduce((sum, e) => sum + e.minutos_na_semana, 0);

  const meta = {
    semana: {
      inicio: startStr,
      fim: new Date(weekEndUtc.getTime() - (4 * 60 * 60 * 1000)).toISOString().split('T')[0],
      label: `Semana de ${ptbrDate(startStr)} a ${ptbrDate(new Date(weekEndUtc.getTime() - (4 * 60 * 60 * 1000)).toISOString().split('T')[0])}`
    },
    gerado_em_br: new Intl.DateTimeFormat('pt-BR', { timeZone: 'America/Cuiaba', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).format(new Date()),
    total_demandas_tocadas: finalDemandas.length,
    total_apontamentos: sortedEntries.length,
    total_horas_semana: formatMinutesToHours(totalMinutosSemana),
    total_minutos_semana: totalMinutosSemana,
    total_economia_horas: finalDemandas.reduce((sum, d) => sum + (d.economia_estimada || 0), 0)
  };

  // Group by department
  const deptoMap = {};
  finalDemandas.forEach(d => {
    const dep = d.departamento || 'Sem departamento informado';
    if (!deptoMap[dep]) {
      deptoMap[dep] = {
        nome: dep,
        total: 0,
        em_andamento: 0,
        backlog: 0,
        concluidos_semana: 0,
        atrasados: 0,
        decisoes: 0,
        riscos: 0,
        demandas: []
      };
    }
    const m = deptoMap[dep];
    m.total++;
    if (d.normalized_status.in_progress) m.em_andamento++;
    if (d.normalized_status.backlog) m.backlog++;
    if (d.semana.concluida_nesta_semana) m.concluidos_semana++;
    if (d.deadline_status === 'Atrasado') m.atrasados++;
    if (d.tags.includes('DECISAO')) m.decisoes++;
    if (d.tags.includes('RISCO')) m.riscos++;
    m.demandas.push(d.titulo);
  });

  const departamentos = Object.values(deptoMap);

  // Resume for IA
  const resumoParaIa = {
    periodo: meta.semana.label,
    total_horas: meta.total_horas_semana,
    total_economia_horas: meta.total_economia_horas,
    total_demandas_movimentadas: meta.total_demandas_tocadas,
    metricas_globais: {
      em_andamento: finalDemandas.filter(d => d.normalized_status.in_progress).length,
      backlog: finalDemandas.filter(d => d.normalized_status.backlog).length,
      concluidos_na_semana: finalDemandas.filter(d => d.semana.concluida_nesta_semana).length,
      criados_na_semana: finalDemandas.filter(d => d.semana.criada_nesta_semana).length,
      atrasados: finalDemandas.filter(d => d.deadline_status === 'Atrasado').length,
      concluidos_no_prazo: finalDemandas.filter(d => d.deadline_status === 'No prazo' && d.semana.concluida_nesta_semana).length,
      concluidos_com_atraso: finalDemandas.filter(d => d.deadline_status === 'Com atraso' && d.semana.concluida_nesta_semana).length
    },
    departamentos,
    blocos_tematicos: {
      decisoes: finalDemandas.filter(d => d.tags.includes('DECISAO')).map(d => ({ titulo: d.titulo, descricao: d.comentarios_semana.find(c => c.toLowerCase().includes('#decisao')) || d.descricao, depto: d.departamento })),
      riscos: finalDemandas.filter(d => d.tags.includes('RISCO') || d.deadline_status === 'Atrasado').map(d => ({ titulo: d.titulo, motivo: d.deadline_status === 'Atrasado' ? 'Prazo vencido' : 'Sinalizado como risco' })),
      dependencias: finalDemandas.filter(d => d.tags.includes('DEPENDENCIA')).map(d => ({ titulo: d.titulo, descricao: d.comentarios_semana.find(c => c.toLowerCase().includes('#dependencia')) || 'Aguardando ação externa' })),
      automacoes: finalDemandas.filter(d => d.tags.includes('AUTOMACAO')).map(d => ({ titulo: d.titulo, status: d.status })),
    },
    prazos: {
      concluidos_no_prazo: finalDemandas.filter(d => d.deadline_status === 'No prazo' && d.semana.concluida_nesta_semana).length,
      concluidos_com_atraso: finalDemandas.filter(d => d.deadline_status === 'Com atraso' && d.semana.concluida_nesta_semana).length,
      em_aberto_atrasados: finalDemandas.filter(d => d.deadline_status === 'Atrasado').length,
      em_aberto_no_prazo: finalDemandas.filter(d => d.deadline_status === 'No prazo (aberto)').length
    },
    demandas_detalhe: finalDemandas.map(d => ({
      id: d.id,
      titulo: d.titulo,
      status: d.status,
      depto: d.departamento,
      tempo: d.semana.tempo_horas,
      tags: d.tags,
      prazo: d.deadline_status
    }))
  };

  return {
    meta,
    resumo_para_ia: resumoParaIa,
    demandas: finalDemandas
  };
}

/**
 * ============================================================
 * RELATÓRIO SEMANAL - ENDPOINTS
 * ============================================================
 */

app.get('/api/report/week', async (req, res) => {
  try {
    const report = await getWeeklyReportData(req.query.start);
    res.json({ ok: true, ...report });
  } catch (error) {
    console.error('ERRO /api/report/week =>', error);
    res.status(error.status || 500).json({
      ok: false,
      message: error.message
    });
  }
});

/**
 * ============================================================
 * IA NARRATIVA SEMANAL
 * ============================================================
 */

function buildWeeklyNarrativePrompt(resumoParaIa) {
  const r = resumoParaIa;

  return `Você é um assistente executivo sênior especializado em relatórios de gestão operacional para a Franco Sistemas / FocusTrack.
Sua tarefa é gerar um relatório executivo semanal completo baseado EXCLUSIVAMENTE nos dados fornecidos abaixo.

DIRETRIZES DE ESTILO:
- Tom executivo, direto, profissional e focado em resultados/impacto.
- Não use palavras vazias ou clichês corporativos (ex: "empenho total", "busca constante").
- Transforme métricas em conclusões úteis (ex: em vez de "10 frentes", diga "A operação mantém 10 frentes simultâneas...").
- Escreva em português brasileiro (PT-BR).
- Nunca mencione "JSON", "IA", "tags" ou o processo de geração.

REGRAS DE DADOS (CRÍTICO):
- NÃO INVENTE NÚMEROS. Use exatamente o que está no JSON.
- Se uma métrica for 0, mencione como um ponto de estabilidade ou ausência (ex: "Sem registros de atrasos na semana").
- Use os departamentos para dar contexto geográfico/setorial ao trabalho.
- DISTINÇÃO IMPORTANTE: Não misture Riscos com Dependências. 
  - RISCOS: Ameaças ao prazo ou qualidade originadas no desenvolvimento ou ambiente.
  - DEPENDÊNCIAS: Bloqueios externos, aguardando aprovação ou ação de terceiros/clientes.

ESTRUTURA DO RELATÓRIO (Retorne em JSON):
{
  "chamada_capa": "Frase curta e impactante (título do e-mail)",
  "resumo_executivo": "Visão geral da semana em 2 parágrafos",
  "leitura_30s": {
    "entrega_principal": "Destaque de conclusão",
    "foco_operacional": "Onde o tempo foi mais investido",
    "ponto_gestao": "Decisão ou risco que exige atenção"
  },
  "analise_por_departamento": "Narrativa sobre como as áreas se comportaram",
  "vitorias_da_semana": [{ "titulo": "", "descricao": "" }],
  "principais_entregas": [{ "titulo": "", "descricao": "", "tempo": "" }],
  "demandas_em_andamento": [{ "titulo": "", "descricao": "", "tempo": "" }],
  "decisoes_e_prazos": "Narrativa integrando as decisões tomadas e o status de cumprimento de prazos",
  "pontos_de_atencao": [{ "titulo": "", "descricao": "" }],
  "proximos_passos": [{ "titulo": "", "descricao": "" }],
  "texto_email_curto": "Corpo de e-mail para encaminhamento do relatório"
}

DADOS DA SEMANA:
${JSON.stringify(r, null, 2)}`;
}

function getAiConfig() {
  let baseUrl = process.env.AI_API_BASE_URL || 'https://api.openai.com/v1';
  if (baseUrl.endsWith('/')) baseUrl = baseUrl.slice(0, -1);

  return {
    baseUrl,
    apiKey: process.env.AI_API_KEY || process.env.OPENAI_API_KEY,
    model: process.env.AI_MODEL || process.env.OPENAI_MODEL || 'gpt-4o-mini',
    responseFormat: process.env.AI_API_RESPONSE_FORMAT || 'json_object'
  };
}

async function generateWeeklyNarrative(resumoParaIa) {
  const config = getAiConfig();
  if (!config.baseUrl || !config.apiKey) {
    throw new Error('Configuração de IA incompleta (AI_API_BASE_URL/AI_API_KEY).');
  }

  const prompt = buildWeeklyNarrativePrompt(resumoParaIa);
  const url = `${config.baseUrl}/chat/completions`;

  const payload = {
    model: config.model,
    messages: [
      { role: 'system', content: 'Responda exclusivamente com JSON válido conforme a estrutura solicitada.' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.2
  };

  if (config.responseFormat === 'json_object') {
    payload.response_format = { type: 'json_object' };
  }

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.apiKey}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(`Erro na API de IA (${response.status}): ${JSON.stringify(errorData)}`);
  }

  const data = await response.json();
  let content = data.choices?.[0]?.message?.content?.trim();

  if (!content) throw new Error('A API de IA retornou uma resposta sem conteúdo.');

  if (content.startsWith('```json')) content = content.replace(/^```json/, '').replace(/```$/, '').trim();
  else if (content.startsWith('```')) content = content.replace(/^```/, '').replace(/```$/, '').trim();

  try {
    return JSON.parse(content);
  } catch (err) {
    console.error('Erro ao fazer parse da resposta da IA:', content);
    throw new Error('Resposta da IA não é um JSON válido.');
  }
}


app.get('/api/report/week/narrative', async (req, res) => {
  try {
    const report = await getWeeklyReportData(req.query.start);
    
    const narrativa = await generateWeeklyNarrative(report.resumo_para_ia);

    res.json({
      ok: true,
      meta: report.meta,
      narrativa,
      ...(req.query.debug === '1'
        ? { dados_base: { resumo_para_ia: report.resumo_para_ia } }
        : {})
    });
  } catch (error) {
    console.error('ERRO /api/report/week/narrative =>', error);
    res.status(error.status || 500).json({
      ok: false,
      message: error.message || "Erro ao gerar narrativa semanal com IA.",
      details: error.details || null
    });
  }
});


/**
 * ============================================================
 * EMAIL HTML - HELPERS
 * ============================================================
 */

function isWeakText(value, minLength) {
  const min = minLength == null ? 80 : minLength;
  const text = String(value || '').trim();
  if (!text || text.length < min) return true;
  return [
    /^resumo da produtividade/i,
    /^sem informa/i,
    /^não definido/i,
    /^não há/i
  ].some(re => re.test(text));
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatPeriodForEmail(meta) {
  try {
    const inicio = meta.semana.inicio;
    const fim = meta.semana.fim;
    const [, sm, sd] = inicio.split('-');
    const [ey, em, ed] = fim.split('-');
    return `${String(sd).padStart(2,'0')}/${String(sm).padStart(2,'0')} a<br>${String(ed).padStart(2,'0')}/${String(em).padStart(2,'0')}/${ey}`;
  } catch (_) {
    return escapeHtml(meta?.semana?.label || '');
  }
}

function renderMetricCell(value, label, borderColor) {
  const safeVal = escapeHtml(value);
  const safeLbl = escapeHtml(label);
  return `<td width="25%" style="padding:0 6px 0 0;">
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-top:4px solid ${borderColor};border:1px solid #d9e2ec;border-top:4px solid ${borderColor};border-collapse:collapse;">
    <tr><td style="padding:18px 16px 4px 16px;font-size:32px;font-weight:bold;color:#0d2b4c;font-family:Arial,sans-serif;line-height:1;">${safeVal}</td></tr>
    <tr><td style="padding:0 16px 16px 16px;font-size:10px;color:#6b7a90;font-family:Arial,sans-serif;letter-spacing:1px;text-transform:uppercase;">${safeLbl}</td></tr>
  </table>
</td>`;
}

function renderEmailItems(items, options) {
  const limit = options?.limit || 8;
  const fallback = options?.fallback || '';
  const borderColor = options?.borderColor || '#1a73be';
  const titleColor = options?.titleColor || '#0d2b4c';

  const arr = Array.isArray(items) ? items.slice(0, limit) : [];
  if (arr.length === 0) {
    return fallback
      ? `<tr><td style="padding:10px 0;font-size:13px;color:#777;font-family:Arial,sans-serif;">${escapeHtml(fallback)}</td></tr>`
      : '';
  }

  return arr.map(item => {
    const titulo = typeof item === 'string' ? item : (item.titulo || item.frente || '');
    const descricao = typeof item === 'string' ? '' : (item.descricao || item.texto || '');
    const tempo = typeof item === 'object' ? (item.tempo || '') : '';
    const tempoHtml = tempo ? ` <span style="color:#888;font-size:11px;">(${escapeHtml(tempo)})</span>` : '';
    return `<tr>
  <td style="padding:8px 0 8px 12px;border-left:3px solid ${borderColor};margin-bottom:6px;font-family:Arial,sans-serif;">
    <div style="font-size:13px;font-weight:700;color:${titleColor};">${escapeHtml(titulo)}${tempoHtml}</div>
    ${descricao ? `<div style="font-size:12px;color:#444;margin-top:3px;line-height:1.5;">${escapeHtml(descricao)}</div>` : ''}
  </td>
</tr>
<tr><td style="height:6px;"></td></tr>`;
  }).join('');
}

function renderVictoryItems(items) {
  const arr = Array.isArray(items) ? items.slice(0, 6) : [];
  if (arr.length === 0) return '';
  return arr.map(item => {
    const titulo = typeof item === 'string' ? item : (item.titulo || '');
    const descricao = typeof item === 'string' ? '' : (item.descricao || '');
    const origem = typeof item === 'object' ? (item.origem || '') : '';
    return `<tr>
  <td style="padding:0 0 8px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #ccebdd;border-left:4px solid #2fbf88;background-color:#f0faf6;">
      <tr><td style="padding:14px 16px;">
        <div style="font-size:13px;font-weight:bold;color:#0d4a2e;font-family:Arial,sans-serif;">${escapeHtml(titulo)}</div>
        ${descricao ? `<div style="font-size:12px;color:#2d5a3d;margin-top:5px;line-height:1.6;font-family:Arial,sans-serif;">${escapeHtml(descricao)}</div>` : ''}
        ${origem ? `<div style="font-size:11px;color:#5a8a6a;margin-top:6px;font-family:Arial,sans-serif;">${escapeHtml(origem)}</div>` : ''}
      </td></tr>
    </table>
  </td>
</tr>`;
  }).join('');
}

function renderAttentionItems(items) {
  const arr = Array.isArray(items) ? items.slice(0, 6) : [];
  if (arr.length === 0) {
    return `<tr><td style="padding:10px 12px;background:#fffbea;border-left:4px solid #f2b84b;border-radius:3px;font-size:13px;color:#555;font-family:Arial,sans-serif;">Nenhum ponto crítico identificado para o período.</td></tr>`;
  }
  return arr.map(item => {
    const titulo = typeof item === 'string' ? item : (item.titulo || item.descricao || String(item));
    const descricao = typeof item === 'string' ? '' : (item.descricao && item.titulo ? item.descricao : '');
    return `<tr>
  <td style="padding:10px 12px;background:#fffbea;border-left:4px solid #f2b84b;border-radius:3px;font-family:Arial,sans-serif;">
    <div style="font-size:13px;font-weight:700;color:#7a5800;">${escapeHtml(titulo)}</div>
    ${descricao ? `<div style="font-size:12px;color:#444;margin-top:3px;line-height:1.5;">${escapeHtml(descricao)}</div>` : ''}
  </td>
</tr>
<tr><td style="height:6px;"></td></tr>`;
  }).join('');
}

function renderProjectsTable(items) {
  const arr = Array.isArray(items) ? items.slice(0, 6) : [];
  if (arr.length === 0) return '';
  const thStyle = 'padding:8px 10px;background:#0d2b4c;color:#fff;font-size:11px;font-family:Arial,sans-serif;text-align:left;font-weight:700;';
  const tdStyle = 'padding:8px 10px;font-size:12px;color:#333;font-family:Arial,sans-serif;border-bottom:1px solid #e5e5e5;vertical-align:top;';
  const rows = arr.map((item, i) => {
    const bg = i % 2 === 0 ? '#fff' : '#f7f9fc';
    return `<tr style="background:${bg};">
  <td style="${tdStyle}">${escapeHtml(item.frente || '')}</td>
  <td style="${tdStyle}">${escapeHtml(item.status || '')}</td>
  <td style="${tdStyle}">${escapeHtml(item.impacto || '')}</td>
  <td style="${tdStyle}">${escapeHtml(item.proximo_passo || '')}</td>
</tr>`;
  }).join('');
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:1px solid #ddd;border-radius:4px;overflow:hidden;">
  <tr>
    <th style="${thStyle}">Frente</th>
    <th style="${thStyle}">Status</th>
    <th style="${thStyle}">Impacto</th>
    <th style="${thStyle}">Próximo passo</th>
  </tr>
  ${rows}
</table>`;
}

function escapeHtml(text) {
  if (typeof text !== 'string') return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function renderMetricCell(value, label, color) {
  return `
    <td width="24%" style="vertical-align:top;padding:15px;background:#f7f9fc;border-bottom:3px solid ${color};">
      <div style="font-size:22px;font-weight:bold;color:${color};font-family:Arial,sans-serif;">${value}</div>
      <div style="font-size:10px;color:#6b7a90;text-transform:uppercase;letter-spacing:1px;margin-top:4px;font-family:Arial,sans-serif;">${label}</div>
    </td>
    <td width="1%"></td>
  `;
}

function formatPeriodForEmail(meta) {
  if (!meta.inicio || !meta.fim) return 'Período não definido';
  const i = new Date(meta.inicio);
  const f = new Date(meta.fim);
  const d = (date) => date.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' });
  return `${d(i)} a ${d(f)}`;
}

function buildWeeklyReportEmailHtml({ meta, narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};
  const g = r.metricas_globais || {};

  const TEAMS_LINK = process.env.FOCUSTRACK_TEAMS_LINK || '';
  const LOGO_FRANCO = 'https://raw.githubusercontent.com/ffqmt/Images/15116cdbaa87af68eb9eaf9a1bea9ee7502bb9f7/FRANCO%20LOGO.png';
  const LOGO_CONTAUDI = 'https://raw.githubusercontent.com/ffqmt/Images/15116cdbaa87af68eb9eaf9a1bea9ee7502bb9f7/LOGO%20IMAGEM.png';

  const SEC = 'padding:30px 34px;border-bottom:1px solid #d9e2ec;';
  const SEC_LABEL = 'margin:0 0 6px 0;font-size:10px;color:#1a73be;letter-spacing:3px;text-transform:uppercase;font-weight:bold;font-family:Arial,sans-serif;';
  const SEC_TITLE = 'margin:0 0 20px 0;font-size:20px;line-height:26px;color:#0d2b4c;font-weight:bold;font-family:Arial,sans-serif;';
  
  const sectionHeader = (label, title) => `<p style="${SEC_LABEL}">${label}</p><h2 style="${SEC_TITLE}">${title}</h2>`;

  const renderSimpleTable = (items, color = '#1a73be') => {
    if (!items || items.length === 0) return '<tr><td style="font-size:12px;color:#777;padding:10px 0;">Nenhum item registrado.</td></tr>';
    return items.map(item => `
      <tr>
        <td style="padding:10px 12px;border-left:4px solid ${color};background:#f8fafd;margin-bottom:8px;font-family:Arial,sans-serif;">
          <div style="font-size:13px;font-weight:bold;color:#0d2b4c;">${escapeHtml(item.titulo)}</div>
          ${item.descricao ? `<div style="font-size:12px;color:#444;margin-top:4px;">${escapeHtml(item.descricao)}</div>` : ''}
          ${item.tempo ? `<div style="font-size:11px;color:#888;margin-top:4px;">Tempo: ${item.tempo}</div>` : ''}
        </td>
      </tr>
      <tr><td style="height:6px;"></td></tr>
    `).join('');
  };

  const renderThematicBlock = (items, color, label) => {
    if (!items || items.length === 0) return '';
    return `
      <tr>
        <td style="${SEC}">
          ${sectionHeader('Destaque', label)}
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            ${items.map(item => `
              <tr>
                <td style="padding:12px;background:#f0f5fa;border-left:4px solid ${color};">
                  <div style="font-size:13px;font-weight:bold;color:#0d2b4c;">${escapeHtml(item.titulo)}</div>
                  <div style="font-size:12px;color:#444;margin-top:4px;">${escapeHtml(item.descricao || item.motivo || '')}</div>
                </td>
              </tr>
              <tr><td style="height:8px;"></td></tr>
            `).join('')}
          </table>
        </td>
      </tr>
    `;
  };

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>Relatório Weekly FocusTrack</title></head>
<body style="margin:0;padding:0;background-color:#edf1f6;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#edf1f6">
<tr><td align="center" style="padding:20px;">
  <table width="800" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid #d9e2ec;">
    
    <!-- HEADER -->
    <tr>
      <td style="padding:20px 34px;border-bottom:1px solid #d9e2ec;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="${LOGO_FRANCO}" width="120"></td>
            <td align="right"><img src="${LOGO_CONTAUDI}" width="120"></td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- TITLE & COVER -->
    <tr>
      <td style="padding:40px 34px;background:#f8fafd;border-bottom:1px solid #d9e2ec;">
        <p style="${SEC_LABEL}">${meta.semana.label}</p>
        <h1 style="font-size:28px;color:#0d2b4c;margin:0 0 15px 0;">${escapeHtml(n.chamada_capa || 'Relatório de Gestão Operacional')}</h1>
        <p style="font-size:14px;color:#3d5370;line-height:1.6;margin:0;">${escapeHtml(n.resumo_executivo || '')}</p>
      </td>
    </tr>

    <!-- METRICS -->
    <tr>
      <td style="${SEC}">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            ${renderMetricCell(meta.total_horas_semana || '0h', 'Total Horas', '#1a73be')}
            ${renderMetricCell(g.concluidos_na_semana || 0, 'Concluídas', '#2fbf88')}
            ${renderMetricCell(g.em_andamento || 0, 'Em Andamento', '#0d2b4c')}
            ${renderMetricCell(g.backlog || 0, 'Backlog / Triagem', '#f2b84b')}
          </tr>
        </table>
      </td>
    </tr>

    <!-- LEITURA 30s -->
    <tr>
      <td style="${SEC}">
        ${sectionHeader('Rápido', 'Leitura em 30 segundos')}
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="32%" style="padding:15px;background:#f4fbf7;border-top:3px solid #2fbf88;vertical-align:top;">
              <div style="font-size:11px;font-weight:bold;color:#2fbf88;margin-bottom:5px;">ENTREGA PRINCIPAL</div>
              <div style="font-size:12px;color:#333;">${escapeHtml(n.leitura_30s?.entrega_principal || 'N/A')}</div>
            </td>
            <td width="2%"></td>
            <td width="32%" style="padding:15px;background:#f7fbff;border-top:3px solid #1a73be;vertical-align:top;">
              <div style="font-size:11px;font-weight:bold;color:#1a73be;margin-bottom:5px;">FOCO OPERACIONAL</div>
              <div style="font-size:12px;color:#333;">${escapeHtml(n.leitura_30s?.foco_operacional || 'N/A')}</div>
            </td>
            <td width="2%"></td>
            <td width="32%" style="padding:15px;background:#fffcf0;border-top:3px solid #f2b84b;vertical-align:top;">
              <div style="font-size:11px;font-weight:bold;color:#c48b00;margin-bottom:5px;">PONTO DE GESTÃO</div>
              <div style="font-size:12px;color:#333;">${escapeHtml(n.leitura_30s?.ponto_gestao || 'N/A')}</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- DEPARTAMENTAL -->
    <tr>
      <td style="${SEC}">
        ${sectionHeader('Departamentos', 'Análise por Frente de Trabalho')}
        <p style="font-size:13px;color:#333;line-height:1.6;margin-bottom:15px;">${escapeHtml(n.analise_por_departamento || '')}</p>
        <table width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;border:1px solid #d9e2ec;font-size:11px;">
          <tr style="background:#0d2b4c;color:#fff;">
            <th style="padding:8px;text-align:left;">Departamento</th>
            <th style="padding:8px;">Total</th>
            <th style="padding:8px;">Em Andamento</th>
            <th style="padding:8px;">Backlog</th>
            <th style="padding:8px;">Atrasados</th>
          </tr>
          ${(r.departamentos || []).map(d => `
            <tr>
              <td style="padding:8px;font-weight:bold;">${escapeHtml(d.nome)}</td>
              <td style="padding:8px;text-align:center;">${d.total}</td>
              <td style="padding:8px;text-align:center;">${d.em_andamento}</td>
              <td style="padding:8px;text-align:center;">${d.backlog}</td>
              <td style="padding:8px;text-align:center;color:${d.atrasados > 0 ? '#d93025' : '#777'};">${d.atrasados}</td>
            </tr>
          `).join('')}
        </table>
      </td>
    </tr>

    <!-- PRAZOS E DECISÕES -->
    <tr>
      <td style="${SEC}">
        ${sectionHeader('Operacional', 'Prazos e Decisões de Gestão')}
        <div style="padding:15px;background:#f7f9fc;border-left:4px solid #0d2b4c;font-size:13px;color:#333;line-height:1.7;">
          ${escapeHtml(n.decisoes_e_prazos || 'Nenhuma decisão ou observação de prazo relevante registrada.')}
        </div>
      </td>
    </tr>

    <!-- THEMATIC BLOCKS -->
    ${renderThematicBlock(r.blocos_tematicos?.decisoes, '#0d2b4c', 'Decisões da Semana')}
    ${renderThematicBlock(r.blocos_tematicos?.riscos, '#d93025', 'Riscos e Atenção')}
    ${renderThematicBlock(r.blocos_tematicos?.dependencias, '#f2b84b', 'Dependências e Bloqueios')}

    <!-- AUTOMATION SECTION -->
    <tr>
      <td style="${SEC}background:#f0f7ff;">
        ${sectionHeader('Eficiência', 'Automações e Economia de Tempo')}
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="60%" style="vertical-align:top;">
              <p style="font-size:13px;color:#0d2b4c;margin:0 0 10px 0;">Frentes de automação identificadas ou em desenvolvimento:</p>
              ${renderSimpleTable(r.blocos_tematicos?.automacoes, '#1a73be')}
            </td>
            <td width="40%" style="padding-left:20px;vertical-align:top;">
              <div style="padding:15px;background:#fff;border:1px dashed #1a73be;text-align:center;">
                <div style="font-size:10px;color:#1a73be;margin-bottom:5px;">ECONOMIA ESTIMADA</div>
                <div style="font-size:24px;font-weight:bold;color:#0d2b4c;">${meta.total_economia_horas || 0}</div>
                <div style="font-size:10px;color:#888;margin-top:5px;">Horas/Mês</div>
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- LISTS -->
    <tr>
      <td style="${SEC}">
        ${sectionHeader('Entregas', 'Principais Conclusões')}
        <table width="100%" cellpadding="0" cellspacing="0">${renderSimpleTable(n.principais_entregas, '#2fbf88')}</table>
      </td>
    </tr>

    <!-- PRÓXIMOS PASSOS -->
    <tr>
      <td style="${SEC}">
        ${sectionHeader('Próximos', 'Foco Planejado')}
        <table width="100%" cellpadding="0" cellspacing="0">${renderSimpleTable(n.proximos_passos, '#0d2b4c')}</table>
      </td>
    </tr>

    <!-- FOOTER -->
    <tr>
      <td style="padding:30px;background:#0d2b4c;color:#fff;text-align:center;">
        <div style="font-size:12px;font-weight:bold;">Franco Sistemas & Contaudi</div>
        <div style="font-size:10px;color:#7fb3d3;margin-top:5px;">Relatório Weekly Automatizado &middot; FocusTrack v2</div>
      </td>
    </tr>

  </table>
</td></tr>
</table>
</body>
</html>`;
}

/**
 * ============================================================
 * EMAIL HTML - ENDPOINT
 * ============================================================
 */

app.get('/api/report/week/email-html', async (req, res) => {
  try {
    const report = await getWeeklyReportData(req.query.start);
    const narrativa = await generateWeeklyNarrative(report.resumo_para_ia);
    const html = buildWeeklyReportEmailHtml({
      meta: report.meta,
      narrativa,
      resumoParaIa: report.resumo_para_ia
    });

    res.json({
      ok: true,
      meta: report.meta,
      html,
      ...(req.query.debug === '1'
        ? { dados_base: { resumo_para_ia: report.resumo_para_ia, narrativa } }
        : {})
    });
  } catch (error) {
    console.error('ERRO /api/report/week/email-html =>', error);
    res.status(error.status || 500).json({
      ok: false,
      message: error.message || 'Erro ao gerar HTML do relatório semanal.',
      details: error.details || null
    });
  }
});

/**
 * ============================================================
 * START SERVER
 * ============================================================
 */

const port = Number(process.env.PORT || 3001);

app.listen(port, () => {
  console.log(`Backend rodando em http://localhost:${port}`);

  const config = getSharePointConfig();

  console.log('Config SharePoint carregada:', {
    siteUrl: config.siteUrl,
    ticketsListId: config.ticketsListId,
    timeEntriesListId: config.timeEntriesListId,

    ticketRequesterLookupField: config.ticketRequesterLookupField,
    ticketCreatedAtField: config.ticketCreatedAtField,

    ticketStatusField: config.ticketStatusField,
    ticketStartedAtField: config.ticketStartedAtField,
    ticketFinishedAtField: config.ticketFinishedAtField,
    ticketTimeSpentField: config.ticketTimeSpentField,
    ticketStatusInProgressValue: config.ticketStatusInProgressValue,
    ticketStatusFinishedValue: config.ticketStatusFinishedValue,

    timeResponsibleLookupField: config.timeResponsibleLookupField,
    focusCommentMaxChars: getCentralCommentMaxChars(),
    commentDeliveryMode: 'central-field',

    nativeCommentsEnabled: config.nativeCommentsEnabled,
    sharePointRestOrigin: getSharePointOrigin() || null,

    userLookupMapCount: Object.keys(getUserLookupMap()).length,
    defaultUserEmail: getDefaultUserEmail() || null,
    defaultUserLookupId: getDefaultUserLookupId() || null,

    timeCentralCommentField: config.timeCentralCommentField,
    timeSendCentralCommentField: config.timeSendCentralCommentField,
    timeCentralCommentSentField: config.timeCentralCommentSentField,
    timeCentralCommentErrorField: config.timeCentralCommentErrorField
  });
});
