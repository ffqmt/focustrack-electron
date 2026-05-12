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
    .replace(/[̀-ͯ]/g, '')
    .trim()
    .toLowerCase();
}

function isBacklogStatus(status) {
  const s = normalizeStatusLabel(status);
  return s === 'novo' || s === 'em triagem' || s === 'triagem';
}

// ── Tags manuais ────────────────────────────────────────────
const MANUAL_SIGNAL_TAGS = ['WIN', 'RISCO', 'DECISAO', 'PROXIMO', 'IMPACTO', 'DOC', 'IA'];

function extractManualSignals(text) {
  const raw = String(text || '');
  return MANUAL_SIGNAL_TAGS.filter(tag => new RegExp(`#${tag}\\b`, 'i').test(raw));
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
  // Segunda-feira 00:00 Cuiabá (UTC-4) = Segunda-feira 04:00 UTC
  const weekStartUtc = new Date(Date.UTC(year, month - 1, day, 4, 0, 0));
  // Domingo 23:59:59 Cuiabá (UTC-4) = Segunda-feira 03:59:59 UTC
  const weekEndUtc = new Date(weekStartUtc.getTime() + (7 * 24 * 60 * 60 * 1000) - 1000);

  const siteId = await resolveSiteId();
  const config = getSharePointConfig();

  const timeEntries = await getListItemsByQuery(siteId, config.timeEntriesListId, {
    filter: `fields/Inicio ge '${weekStartUtc.toISOString()}' and fields/Inicio le '${weekEndUtc.toISOString()}'`,
    expand: 'fields',
    top: 999
  });

  // Ordenar por Inicio crescente
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

    if (atravessaLimite) {
      alertas.push({
        tipo: 'apontamento_atravessa_semana',
        demanda_id: ticketId ? String(ticketId) : null,
        apontamento_id: entryData.id,
        mensagem: `Apontamento #${entryData.id} atravessa o limite da semana (Original: ${minutosOriginal}min, Na Semana: ${minutosNaSemana}min).`
      });
    }

    if (ticketId) {
      ticketIds.add(String(ticketId));
      if (!apontamentosPorDemanda[ticketId]) {
        apontamentosPorDemanda[ticketId] = [];
      }
      apontamentosPorDemanda[ticketId].push(entryData);
    } else {
      apontamentosSemDemanda.push(entryData);
      alertas.push({
        tipo: 'apontamento_sem_demanda',
        demanda_id: null,
        apontamento_id: entryData.id,
        mensagem: `Apontamento #${entryData.id} não possui demanda vinculada.`
      });
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
        departamento: tf.Departamento || '',
        origem: tf.Origem || '',
        descricao: tf.Descri_x00e7__x00e3_o || '',
        data_criacao: tf.DataCria_x00e7__x00e3_o || tf.Created || null,
        data_conclusao: tf.DataConclus_x00e3_o || null,
        tempo_total_historico_min: toSafeNumber(tf.TempoGasto),
        tempo_total_historico_horas: formatMinutesToHours(tf.TempoGasto)
      };
    } catch (err) {
      console.warn(`Erro ao buscar demanda ${tid}:`, err.message);
      demandasMap[tid] = { id: tid, error: 'Não encontrada' };
    }
  }

  let countComentarioPendente = 0;
  let countDemandaSemComentario = 0;

  const finalDemandas = Object.entries(apontamentosPorDemanda).map(([tid, entries]) => {
    const demandaInfo = demandasMap[tid] || { id: tid };
    const totalMinOriginal = entries.reduce((sum, e) => sum + e.minutos_original, 0);
    const totalMinSemana = entries.reduce((sum, e) => sum + e.minutos_na_semana, 0);
    const temAtravessado = entries.some(e => e.atravessa_limite_semana);
    
    // Flags de semana
    const iniciadaNestaSemana = demandaInfo.data_criacao ? 
      (new Date(demandaInfo.data_criacao) >= weekStartUtc && new Date(demandaInfo.data_criacao) <= weekEndUtc) : false;
    const concluidaNestaSemana = demandaInfo.data_conclusao ? 
      (new Date(demandaInfo.data_conclusao) >= weekStartUtc && new Date(demandaInfo.data_conclusao) <= weekEndUtc) : false;

    const temComentario = entries.some(e => e.comentario_raw);
    if (!temComentario) {
      countDemandaSemComentario++;
      alertas.push({
        tipo: 'demanda_sem_comentario',
        demanda_id: String(tid),
        apontamento_id: null,
        mensagem: `Demanda #${tid} não possui comentários centrais em nenhum apontamento desta semana.`
      });
    }

    entries.forEach(e => {
      if (e.comentario_raw && !e.comentario_enviado) {
        countComentarioPendente++;
        alertas.push({
          tipo: 'comentario_nao_confirmado',
          demanda_id: String(tid),
          apontamento_id: e.id,
          mensagem: `Apontamento #${e.id} possui comentário pendente de envio.`
        });
      }
    });

    const comentariosConsolidados = [];
    entries.forEach(e => {
      e.comentarios.forEach(c => {
        comentariosConsolidados.push(c);
      });
    });

    return {
      ...demandaInfo,
      semana: {
        tempo_minutos_original: totalMinOriginal,
        tempo_minutos: totalMinSemana,
        tempo_horas: formatMinutesToHours(totalMinSemana),
        quantidade_apontamentos: entries.length,
        concluida_nesta_semana: concluidaNestaSemana,
        iniciada_nesta_semana: iniciadaNestaSemana,
        tem_apontamento_atravessado: temAtravessado
      },
      apontamentos: entries,
      comentarios_consolidados: comentariosConsolidados
    };
  });

  const totalMinutosOriginais = sortedEntries.reduce((sum, item) => sum + toSafeNumber(item.fields?.Minutos), 0);
  const totalMinutosSemana = finalDemandas.reduce((sum, d) => sum + d.semana.tempo_minutos, 0) + 
                             apontamentosSemDemanda.reduce((sum, e) => sum + e.minutos_na_semana, 0);

  const [endYear, endMonth, endDate] = new Date(weekEndUtc.getTime() - (4 * 60 * 60 * 1000)).toISOString().split('T')[0].split('-').map(Number);
  const labelFim = `${String(endDate).padStart(2, '0')}/${String(endMonth).padStart(2, '0')}/${endYear}`;
  const [startYear, startMonth, startDay] = startStr.split('-').map(Number);
  const labelInicio = `${String(startDay).padStart(2, '0')}/${String(startMonth).padStart(2, '0')}/${startYear}`;

  const geradoEmBr = new Intl.DateTimeFormat('pt-BR', {
    timeZone: 'America/Cuiaba',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  }).format(new Date());

  const meta = {
    semana: {
      inicio: startStr,
      fim: new Date(weekEndUtc.getTime() - (4 * 60 * 60 * 1000)).toISOString().split('T')[0],
      label: `Semana de ${labelInicio} a ${labelFim}`
    },
    gerado_em: new Date().toISOString(),
    gerado_em_br: geradoEmBr,
    total_demandas_tocadas: finalDemandas.length,
    total_apontamentos: sortedEntries.length,
    total_minutos_originais_lancamentos: totalMinutosOriginais,
    total_minutos_semana: totalMinutosSemana,
    total_horas_semana: formatMinutesToHours(totalMinutosSemana),
    fonte: 'FocusTrack + SharePoint'
  };

  const metricas = {
    demandas_sem_comentario: countDemandaSemComentario,
    apontamentos_sem_demanda: apontamentosSemDemanda.length,
    apontamentos_com_comentario_pendente: countComentarioPendente
  };

  const mapResumoDemanda = (d) => ({
    id: d.id,
    titulo: d.titulo,
    status: d.status,
    tipo: d.tipo,
    departamento: d.departamento,
    origem: d.origem,
    tempo_horas: d.semana.tempo_horas,
    tempo_minutos: d.semana.tempo_minutos,
    quantidade_apontamentos: d.semana.quantidade_apontamentos,
    concluida_nesta_semana: d.semana.concluida_nesta_semana,
    iniciada_nesta_semana: d.semana.iniciada_nesta_semana,
    comentarios: d.comentarios_consolidados.map(c => c.texto)
  });

  const concluidas = finalDemandas.filter(d => d.status === 'Concluído').map(mapResumoDemanda);
  const emAndamento = finalDemandas.filter(d => d.status === 'Em andamento').map(mapResumoDemanda);
  const outras = finalDemandas.filter(d => d.status !== 'Concluído' && d.status !== 'Em andamento').map(mapResumoDemanda);

  const backlogExecutivo = finalDemandas
    .filter(d => isBacklogStatus(d.status))
    .map(d => {
      const leituraBacklog = normalizeStatusLabel(d.status) === 'novo'
        ? 'Demanda registrada, ainda sem triagem concluída. Deve ser avaliada para definição de prioridade, responsável e próximo passo.'
        : 'Demanda em avaliação inicial. Requer definição de escopo, impacto e encaminhamento operacional.';
      const comentsDemanda = d.comentarios_consolidados.map(c => c.texto);
      return {
        demanda_id: d.id,
        titulo: d.titulo,
        status: d.status,
        origem: d.origem,
        tempo_horas: d.semana.tempo_horas,
        tempo_minutos: d.semana.tempo_minutos,
        comentarios: comentsDemanda,
        ultimo_comentario: comentsDemanda[comentsDemanda.length - 1] || '',
        frente: inferManagementFront({ demanda: d, comentarios: comentsDemanda }),
        leitura_executiva: leituraBacklog
      };
    });

  const resumoParaIa = {
    periodo: meta.semana.label,
    total_horas: meta.total_horas_semana,
    total_minutos: meta.total_minutos_semana,
    demandas_movimentadas: meta.total_demandas_tocadas,
    total_apontamentos: meta.total_apontamentos,
    demandas_por_status: {
      concluidas,
      em_andamento: emAndamento,
      outras
    },
    top_demandas_por_tempo: [...finalDemandas]
      .sort((a, b) => b.semana.tempo_minutos - a.semana.tempo_minutos)
      .slice(0, 10)
      .map(mapResumoDemanda),
    entregas_concluidas: concluidas,
    demandas_em_andamento: emAndamento,
    comentarios_relevantes: [],
    comentarios_executivos: [],
    pontos_de_atencao: [],
    vitorias_sinalizadas: [],
    backlog_executivo: backlogExecutivo,
    sinais_manuais: { wins: [], riscos: [], decisoes: [], proximos: [], impactos: [], documentacao: [], ia: [] },
    carteira_executiva: null,
    temas_da_semana: [],
    insights_operacionais: [],
    base_narrativa: `Na ${meta.semana.label}, foram registradas ${meta.total_horas_semana} em ${meta.total_demandas_tocadas} demandas, distribuídas em ${meta.total_apontamentos} apontamentos. A semana registrou ${concluidas.length} demandas concluídas e ${emAndamento.length} em andamento, com destaque para as demandas de maior tempo registrado.`
  };

  const SINAL_CHAVE = { WIN: 'wins', RISCO: 'riscos', DECISAO: 'decisoes', PROXIMO: 'proximos', IMPACTO: 'impactos', DOC: 'documentacao', IA: 'ia' };

  finalDemandas.forEach(d => {
    d.apontamentos.forEach(a => {
      a.comentarios.forEach(c => {
        resumoParaIa.comentarios_relevantes.push({
          demanda_id: d.id,
          titulo: d.titulo,
          horario: c.horario,
          texto: c.texto
        });

        const classificado = classifyCommentForExecutiveReport({ comentario: c.texto, demanda: d });
        resumoParaIa.comentarios_executivos.push({
          ...classificado,
          horario: c.horario,
          tempo_horas: d.semana.tempo_horas,
          tempo_minutos: d.semana.tempo_minutos
        });

        const sinalItem = {
          demanda_id: d.id,
          titulo: d.titulo,
          status: d.status,
          origem: d.origem,
          texto_limpo: classificado.texto_limpo,
          frente: classificado.frente,
          leitura_executiva: classificado.leitura_executiva,
          tempo_horas: d.semana.tempo_horas,
          tempo_minutos: d.semana.tempo_minutos
        };

        classificado.sinais.forEach(tag => {
          const chave = SINAL_CHAVE[tag];
          if (chave) resumoParaIa.sinais_manuais[chave].push(sinalItem);
        });

        if (hasWinTag(c.texto)) {
          resumoParaIa.vitorias_sinalizadas.push({
            demanda_id: d.id,
            titulo: d.titulo,
            status: d.status,
            tipo: d.tipo,
            departamento: d.departamento,
            origem: d.origem,
            horario: c.horario,
            texto: cleanWinTag(c.texto),
            tempo_horas: d.semana.tempo_horas,
            tempo_minutos: d.semana.tempo_minutos
          });
        }
      });
    });
  });

  resumoParaIa.carteira_executiva = buildExecutivePortfolio({
    demandas: finalDemandas,
    comentariosExecutivos: resumoParaIa.comentarios_executivos,
    backlogExecutivo
  });

  resumoParaIa.temas_da_semana = buildWeeklyThemes({
    resumoParaIa,
    comentariosExecutivos: resumoParaIa.comentarios_executivos,
    carteiraExecutiva: resumoParaIa.carteira_executiva
  });

  resumoParaIa.insights_operacionais = buildOperationalInsights({
    resumoParaIa,
    carteiraExecutiva: resumoParaIa.carteira_executiva
  });

  if (metricas.apontamentos_com_comentario_pendente > 0) {
    resumoParaIa.pontos_de_atencao.push(`Existem ${metricas.apontamentos_com_comentario_pendente} apontamentos com comentário pendente de envio pelo Power Automate.`);
  }
  if (metricas.demandas_sem_comentario > 0) {
    resumoParaIa.pontos_de_atencao.push(`Exist${metricas.demandas_sem_comentario > 1 ? 'em' : 'e'} ${metricas.demandas_sem_comentario} demanda${metricas.demandas_sem_comentario > 1 ? 's' : ''} sem comentário central nesta semana.`);
  }
  if (alertas.some(a => a.tipo === 'apontamento_atravessa_semana')) {
    resumoParaIa.pontos_de_atencao.push("Existem apontamentos que atravessam o limite da semana; os minutos foram recortados para o período solicitado.");
  }
  if (metricas.apontamentos_sem_demanda > 0) {
    resumoParaIa.pontos_de_atencao.push(`Existem ${metricas.apontamentos_sem_demanda} apontamentos sem demanda vinculada.`);
  }

  return {
    meta,
    metricas,
    alertas,
    resumo_para_ia: resumoParaIa,
    demandas: finalDemandas,
    apontamentos_sem_demanda: apontamentosSemDemanda
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
  const dadosCompactos = {
    periodo: resumoParaIa.periodo,
    total_horas: resumoParaIa.total_horas,
    total_apontamentos: resumoParaIa.total_apontamentos,
    demandas_movimentadas: resumoParaIa.demandas_movimentadas,
    entregas_concluidas: resumoParaIa.entregas_concluidas,
    demandas_em_andamento: resumoParaIa.demandas_em_andamento,
    top_demandas_por_tempo: resumoParaIa.top_demandas_por_tempo,
    vitorias_sinalizadas: resumoParaIa.vitorias_sinalizadas,
    sinais_manuais: resumoParaIa.sinais_manuais,
    backlog_executivo: resumoParaIa.backlog_executivo,
    carteira_executiva: resumoParaIa.carteira_executiva,
    temas_da_semana: resumoParaIa.temas_da_semana,
    insights_operacionais: resumoParaIa.insights_operacionais,
    comentarios_executivos: (resumoParaIa.comentarios_executivos || []).filter(c => c.peso_executivo >= 3),
    pontos_de_atencao: resumoParaIa.pontos_de_atencao,
    base_narrativa: resumoParaIa.base_narrativa
  };

  return `Você é um assistente executivo sênior especializado em relatórios de gestão operacional.
Sua tarefa é gerar um relatório executivo semanal completo, baseado nos dados fornecidos em JSON.

TOM E ESTILO:
- Escreva para gestão, não para desenvolvedores.
- Tom executivo, claro, confiante e objetivo.
- Não liste tarefas de forma fria — interprete avanço, impacto, risco, decisão e continuidade.
- Transforme comentários em narrativa de impacto.
- Evite frases genéricas como "A equipe continua focada", "Próximos passos não definidos" ou "Resumo da produtividade".
- Evite adjetivos exagerados (incrível, extraordinário, extraordinária).
- Escreva em português brasileiro (PT-BR).
- Não mencione "JSON", "IA", "modelo", "#WIN", "#RISCO" ou qualquer tag técnica no texto final.

REGRAS CRÍTICAS:
- Não invente demandas, horas, números ou nomes.
- Use apenas os dados presentes no JSON fornecido.
- Se uma demanda não tiver comentários, use título, status e tempo.
- Cada seção deve ter função própria — não repita o mesmo tema em seções diferentes.

REGRAS DE VITÓRIAS:
- 'sinais_manuais.wins' contém vitórias sinalizadas manualmente pelo operador — têm prioridade absoluta.
- Se houver itens em 'sinais_manuais.wins', 'vitorias_da_semana' NÃO pode ficar vazio.
- Vitória não é sinônimo de demanda concluída — uma demanda em andamento pode ter vitória parcial.
- Trate vitórias de demandas em andamento como avanços relevantes, marcos parciais ou entregas intermediárias.

REGRAS DE BACKLOG:
- Backlog vem exclusivamente de 'backlog_executivo', que contém demandas com status Novo ou Em triagem.
- Se 'backlog_executivo' estiver vazio, não invente backlog.
- Não use tags para inferir backlog.

REGRAS DE CARTEIRA:
- Use 'carteira_executiva.por_frente' para compor 'carteira_por_frente'.
- Use 'temas_da_semana' para compor 'projetos_interesse_gestao'.

REGRAS DE RISCOS E DECISÕES:
- 'sinais_manuais.riscos' e 'sinais_manuais.decisoes' alimentam 'pontos_de_atencao'.
- 'sinais_manuais.proximos' alimenta 'proximos_passos_sugeridos'.

FUNÇÃO DE CADA SEÇÃO:
- resumo_executivo: visão geral executiva da semana
- chamada_capa: frase de impacto para título do e-mail (até 15 palavras)
- leitura_30s: 3 frases de leitura rápida para gestão
- contexto_operacional: contexto do período — o que moldou a semana
- vitorias_da_semana: avanços percebidos como relevantes (obrigatório se houver wins)
- principais_entregas: demandas concluídas com contexto de valor
- demandas_em_andamento: frentes em evolução, não apenas listagem
- avancos_estruturais: melhorias de processo, sistema ou automação
- carteira_por_frente: organização por natureza do trabalho
- pontos_de_atencao: riscos, bloqueios e dependências
- projetos_interesse_gestao: temas que merecem acompanhamento executivo
- proximos_passos_sugeridos: ações concretas e prioritárias
- fechamento_executivo: encerramento com sentido de avanço
- texto_email: corpo curto de e-mail para encaminhar o relatório

ESTRUTURA DE RETORNO (retorne exatamente estas chaves, sem adicionar nem remover):
{
  "resumo_executivo": "",
  "chamada_capa": "",
  "leitura_30s": {
    "o_que_avancou": "",
    "o_que_foi_entregue": "",
    "o_que_exige_gestao": ""
  },
  "contexto_operacional": { "titulo": "", "texto": "" },
  "vitorias_da_semana": [{ "titulo": "", "descricao": "", "demanda_id": "", "origem": "" }],
  "principais_entregas": [{ "titulo": "", "descricao": "", "demanda_id": "", "tempo": "" }],
  "demandas_em_andamento": [{ "titulo": "", "descricao": "", "demanda_id": "", "tempo": "" }],
  "avancos_estruturais": [{ "titulo": "", "descricao": "" }],
  "carteira_por_frente": [{ "frente": "", "descricao": "" }],
  "pontos_de_atencao": [{ "titulo": "", "descricao": "" }],
  "projetos_interesse_gestao": [{ "frente": "", "status": "", "impacto": "", "proximo_passo": "" }],
  "proximos_passos_sugeridos": [{ "titulo": "", "descricao": "" }],
  "fechamento_executivo": "",
  "texto_email": ""
}

DADOS PARA O RELATÓRIO:
${JSON.stringify(dadosCompactos, null, 2)}`;
}

function getAiConfig() {
  let baseUrl = process.env.AI_API_BASE_URL || 'https://api.openai.com/v1';
  // Normalizar baseUrl para remover barra final
  if (baseUrl.endsWith('/')) {
    baseUrl = baseUrl.slice(0, -1);
  }

  const apiKey = process.env.AI_API_KEY || process.env.OPENAI_API_KEY;
  const model = process.env.AI_MODEL || process.env.OPENAI_MODEL || 'gpt-4o-mini';
  const responseFormat = process.env.AI_API_RESPONSE_FORMAT || 'json_object';

  return {
    baseUrl,
    apiKey,
    model,
    responseFormat
  };
}

async function generateWeeklyNarrative(resumoParaIa) {
  const config = getAiConfig();

  if (!config.baseUrl) {
    throw new Error('AI_API_BASE_URL não configurada para geração de narrativa semanal.');
  }
  if (!config.apiKey) {
    throw new Error('AI_API_KEY não configurada para geração de narrativa semanal.');
  }
  if (!config.model) {
    throw new Error('AI_MODEL não configurado para geração de narrativa semanal.');
  }

  const prompt = buildWeeklyNarrativePrompt(resumoParaIa);
  const url = `${config.baseUrl}/chat/completions`;

  const payload = {
    model: config.model,
    messages: [
      {
        role: 'system',
        content: 'Você é um assistente executivo especializado em transformar dados semanais de produtividade em narrativa clara, objetiva e profissional. Responda exclusivamente com JSON válido, sem markdown, sem bloco de código e sem texto fora do JSON.'
      },
      {
        role: 'user',
        content: prompt
      }
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
    const err = new Error(`Erro na API de IA: ${response.status} - ${JSON.stringify(errorData)}`);
    err.status = response.status;
    err.details = errorData;
    throw err;
  }

  const data = await response.json();
  let content = data.choices?.[0]?.message?.content;

  if (!content) {
    const err = new Error('A API de IA retornou uma resposta sem conteúdo.');
    err.details = data;
    throw err;
  }

  // Limpeza robusta antes do parse
  content = content.trim();
  if (content.startsWith('```json')) {
    content = content.replace(/^```json/, '').replace(/```$/, '').trim();
  } else if (content.startsWith('```')) {
    content = content.replace(/^```/, '').replace(/```$/, '').trim();
  }

  try {
    const parsed = JSON.parse(content);
    return normalizeNarrative(parsed, resumoParaIa);
  } catch (err) {
    console.error('Erro ao fazer parse da resposta da IA:', content);
    const parseErr = new Error('Resposta da IA não é um JSON válido.');
    parseErr.details = content;
    throw parseErr;
  }
}

function normalizeNarrative(raw, resumoParaIa) {
  const toObjArray = (arr, fallbackKey) => {
    if (!Array.isArray(arr)) return [];
    return arr.map(item => {
      if (typeof item === 'string') return { [fallbackKey]: item, descricao: '' };
      return item;
    });
  };

  const r = resumoParaIa || {};
  const sinais = r.sinais_manuais || {};
  const resumoExecutivo = raw.resumo_executivo || r.base_narrativa || '';

  // ── vitórias ────────────────────────────────────────────────
  const vitoriasIa = toObjArray(raw.vitorias_da_semana, 'titulo');
  const idsIa = new Set(vitoriasIa.map(v => String(v.demanda_id || '')).filter(Boolean));
  const vitoriasFallbackSinais = (sinais.wins || [])
    .filter(v => !idsIa.has(String(v.demanda_id || '')))
    .map(v => ({ titulo: v.titulo, descricao: v.texto_limpo || v.leitura_executiva, demanda_id: v.demanda_id, origem: v.origem }));
  const vitoriasFallbackSinalizadas = (r.vitorias_sinalizadas || [])
    .filter(v => !idsIa.has(String(v.demanda_id || '')) && !vitoriasFallbackSinais.some(s => String(s.demanda_id) === String(v.demanda_id)))
    .map(v => ({ titulo: v.titulo, descricao: v.texto, demanda_id: v.demanda_id, origem: v.origem }));
  const vitoriasFinais = [...vitoriasIa, ...vitoriasFallbackSinais, ...vitoriasFallbackSinalizadas];

  // ── pontos de atenção ───────────────────────────────────────
  let pontosIa = toObjArray(raw.pontos_de_atencao, 'titulo');
  if (pontosIa.length === 0) {
    pontosIa = [
      ...(sinais.riscos || []).map(s => ({ titulo: s.titulo || 'Risco identificado', descricao: s.texto_limpo || s.leitura_executiva })),
      ...(sinais.decisoes || []).map(s => ({ titulo: s.titulo || 'Decisão pendente', descricao: s.texto_limpo || s.leitura_executiva }))
    ].slice(0, 6);
  }

  // ── próximos passos ─────────────────────────────────────────
  let proximosIa = toObjArray(raw.proximos_passos_sugeridos, 'titulo');
  if (proximosIa.length === 0) {
    proximosIa = (sinais.proximos || [])
      .map(s => ({ titulo: s.titulo || 'Próximo passo', descricao: s.texto_limpo || s.leitura_executiva }))
      .slice(0, 6);
  }

  // ── carteira por frente ─────────────────────────────────────
  let carteiraIa = toObjArray(raw.carteira_por_frente, 'frente');
  if (carteiraIa.length === 0) {
    carteiraIa = ((r.carteira_executiva || {}).por_frente || [])
      .map(f => ({ frente: f.frente, descricao: f.leitura }))
      .slice(0, 6);
  }

  // ── projetos de gestão ──────────────────────────────────────
  let projetosIa = Array.isArray(raw.projetos_interesse_gestao) ? raw.projetos_interesse_gestao : [];
  if (projetosIa.length === 0) {
    projetosIa = (r.temas_da_semana || [])
      .map(t => ({ frente: t.tema, status: 'Em acompanhamento', impacto: t.impacto_gestao, proximo_passo: t.evidencia }))
      .slice(0, 5);
  }

  // ── leitura 30s ─────────────────────────────────────────────
  const leitura30s = (raw.leitura_30s && typeof raw.leitura_30s === 'object') ? raw.leitura_30s : {};
  const concluidas = r.entregas_concluidas || [];
  const emAndamento = r.demandas_em_andamento || [];
  const primeiroPonto = Array.isArray(r.pontos_de_atencao) && r.pontos_de_atencao.length ? String(r.pontos_de_atencao[0]) : '';

  return {
    resumo_executivo: resumoExecutivo,
    chamada_capa: raw.chamada_capa || resumoExecutivo.split('.')[0].slice(0, 80),
    leitura_30s: {
      o_que_avancou: leitura30s.o_que_avancou || (emAndamento[0] ? `Avanço em ${emAndamento[0].titulo || emAndamento[0]}` : ''),
      o_que_foi_entregue: leitura30s.o_que_foi_entregue || (concluidas[0] ? `Entregue: ${concluidas[0].titulo || concluidas[0]}` : ''),
      o_que_exige_gestao: leitura30s.o_que_exige_gestao || primeiroPonto
    },
    contexto_operacional: (raw.contexto_operacional && typeof raw.contexto_operacional === 'object')
      ? raw.contexto_operacional
      : { titulo: 'Contexto da Semana', texto: resumoExecutivo },
    vitorias_da_semana: vitoriasFinais,
    principais_entregas: toObjArray(raw.principais_entregas, 'titulo'),
    demandas_em_andamento: toObjArray(raw.demandas_em_andamento, 'titulo'),
    avancos_estruturais: toObjArray(raw.avancos_estruturais, 'titulo'),
    carteira_por_frente: carteiraIa,
    pontos_de_atencao: pontosIa,
    projetos_interesse_gestao: projetosIa,
    proximos_passos_sugeridos: proximosIa,
    fechamento_executivo: raw.fechamento_executivo || raw.texto_email || r.base_narrativa || '',
    texto_email: raw.texto_email || raw.fechamento_executivo || ''
  };
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

function renderNextSteps(items) {
  const arr = Array.isArray(items) ? items.slice(0, 6) : [];
  if (arr.length === 0) return '';
  const mid = Math.ceil(arr.length / 2);
  const left = arr.slice(0, mid);
  const right = arr.slice(mid);
  const renderCol = (list, startIndex) => list.map((item, i) => {
    const n = startIndex + i + 1;
    const titulo = typeof item === 'string' ? item : (item.titulo || '');
    const descricao = typeof item === 'string' ? '' : (item.descricao || '');
    return `<tr>
  <td style="padding:6px 0 6px 10px;border-left:3px solid #1a73be;font-family:Arial,sans-serif;vertical-align:top;">
    <div style="font-size:13px;font-weight:700;color:#0d2b4c;">${n}. ${escapeHtml(titulo)}</div>
    ${descricao ? `<div style="font-size:12px;color:#555;margin-top:2px;line-height:1.4;">${escapeHtml(descricao)}</div>` : ''}
  </td>
</tr>
<tr><td style="height:6px;"></td></tr>`;
  }).join('');
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td width="49%" style="vertical-align:top;padding-right:10px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">${renderCol(left, 0)}</table>
    </td>
    <td width="2%"></td>
    <td width="49%" style="vertical-align:top;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">${renderCol(right, mid)}</table>
    </td>
  </tr>
</table>`;
}

function buildExecutiveCoverText({ meta, narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const candidato = n.chamada_capa || n.resumo_executivo || '';
  if (!isWeakText(candidato)) return candidato;

  const total = r.total_horas || meta.total_horas_semana || '';
  const apontamentos = r.total_apontamentos || meta.total_apontamentos || 0;
  const movimentadas = r.demandas_movimentadas || meta.total_demandas_tocadas || 0;
  const concluidas = (r.entregas_concluidas || r.demandas_por_status?.concluidas || []).length;
  const andamento = (r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []).length;
  const wins = (r.vitorias_sinalizadas || []).length;
  const topDemanda = (r.top_demandas_por_tempo || [])[0];

  let partes = [];
  partes.push(
    `Semana com ${apontamentos} registros distribuídos em ${movimentadas} demandas` +
    (total ? `, totalizando ${total} de trabalho registrado` : '') + '.'
  );
  if (concluidas > 0 || andamento > 0) {
    const c = concluidas > 0 ? `${concluidas} entrega${concluidas > 1 ? 's concluídas' : ' concluída'}` : '';
    const a = andamento > 0 ? `${andamento} frente${andamento > 1 ? 's' : ''} em andamento` : '';
    partes.push(`O período combinou ${[c, a].filter(Boolean).join(' e ')}.`);
  }
  if (wins > 0) {
    partes.push(`${wins > 1 ? `${wins} vitórias foram` : 'Uma vitória foi'} sinalizadas pela equipe ao longo da semana.`);
  }
  if (topDemanda) {
    partes.push(`Destaque para a demanda "${topDemanda.titulo}" como principal frente em tempo registrado no período.`);
  }
  return partes.join(' ');
}

function buildManagementReading({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const leitura = n.leitura_30s || {};
  const avancou = leitura.o_que_avancou || '';
  const entregue = leitura.o_que_foi_entregue || '';

  if (!isWeakText(avancou, 40) || !isWeakText(entregue, 40)) {
    return '';
  }

  const andamento = (r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []);
  const comentariosRel = (r.comentarios_relevantes || []);
  const topDemanda = (r.top_demandas_por_tempo || [])[0];

  let texto = 'A semana combinou sustentação operacional e avanço em frentes estratégicas.';
  if (andamento.length > 0) {
    texto += ` A carteira ativa conta com ${andamento.length} frente${andamento.length > 1 ? 's' : ''} em desenvolvimento.`;
  }
  if (comentariosRel.length > 0) {
    texto += ' Os registros mostram atividade detalhada com comentários ao longo dos apontamentos.';
  }
  if (topDemanda) {
    texto += ` A demanda de maior tempo registrado foi "${topDemanda.titulo}".`;
  }
  return texto;
}

function buildStructuralAdvances({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const existentes = Array.isArray(n.avancos_estruturais) ? n.avancos_estruturais : [];
  if (existentes.length >= 2) return existentes.slice(0, 4);

  const result = [...existentes];
  const seen = new Set(existentes.map(e => (e.titulo || '').toLowerCase()));

  const add = (titulo, descricao) => {
    if (result.length >= 4) return;
    if (seen.has(titulo.toLowerCase())) return;
    seen.add(titulo.toLowerCase());
    result.push({ titulo, descricao });
  };

  (r.vitorias_sinalizadas || []).slice(0, 2).forEach(v => {
    add(v.titulo || 'Marco relevante', v.texto || 'Avanço sinalizado manualmente pela equipe operacional.');
  });

  const comentarios = (r.comentarios_relevantes || []);
  const longos = comentarios.filter(c => (c.texto || '').length > 60);
  if (longos.length > 0) {
    add('Detalhamento operacional', 'Registros com comentários extensos indicam envolvimento técnico relevante no período.');
  }

  const top = (r.top_demandas_por_tempo || []).slice(0, 2);
  top.forEach(d => {
    add(d.titulo || 'Frente de destaque', `Demanda com maior tempo registrado na semana (${d.tempo_horas || ''}).`);
  });

  return result;
}

function buildPortfolioByFront({ resumoParaIa }) {
  const r = resumoParaIa || {};
  const n = r.narrativa || {};

  const existentes = Array.isArray(n && n.carteira_por_frente) ? n.carteira_por_frente : [];
  if (existentes.length >= 2) return existentes.slice(0, 6);

  const demandas = [
    ...(r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []),
    ...(r.demandas_por_status?.outras || [])
  ];

  const frentes = {
    'Automação, IA e dados': [],
    'Fiscal e integrações': [],
    'Infraestrutura e suporte': [],
    'Microsoft 365, acessos e rede': [],
    'Backlog / replanejamento': []
  };

  const keywords = {
    'Automação, IA e dados': /automaç|ia\b|ocr|pdf|extrat|dados|robô|bot|python|script/i,
    'Fiscal e integrações': /fiscal|nfse|nfe|sped|irpf|ecd|xml|nota|tribut|api/i,
    'Infraestrutura e suporte': /servidor|acesso|senha|bloqueio|vpn|rede|suporte|infraestrut/i,
    'Microsoft 365, acessos e rede': /teams|sharepoint|outlook|365|office|onedrive|exchange/i
  };

  demandas.forEach(d => {
    const titulo = String(d.titulo || '').toLowerCase();
    let colocado = false;
    for (const [frente, re] of Object.entries(keywords)) {
      if (re.test(titulo)) {
        frentes[frente].push(d.titulo);
        colocado = true;
        break;
      }
    }
    if (!colocado) frentes['Backlog / replanejamento'].push(d.titulo);
  });

  return Object.entries(frentes)
    .filter(([, arr]) => arr.length > 0)
    .slice(0, 6)
    .map(([frente, arr]) => ({
      frente,
      descricao: arr.slice(0, 3).join(', ') + (arr.length > 3 ? ` e mais ${arr.length - 3}` : '') + '.'
    }));
}

function buildManagementProjects({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const existentes = Array.isArray(n.projetos_interesse_gestao) ? n.projetos_interesse_gestao : [];
  if (existentes.length >= 2) return existentes.slice(0, 6);

  const result = [...existentes];
  const seen = new Set(existentes.map(e => (e.frente || '').toLowerCase()));

  const add = (frente, status, impacto, proximo_passo) => {
    if (result.length >= 6) return;
    if (seen.has(frente.toLowerCase())) return;
    seen.add(frente.toLowerCase());
    result.push({ frente, status, impacto, proximo_passo });
  };

  const top = (r.top_demandas_por_tempo || []).slice(0, 3);
  top.forEach(d => {
    const status = d.status || 'Em andamento';
    add(d.titulo || 'Demanda principal', status, `${d.tempo_horas || ''} registradas no período`, 'Acompanhar evolução e validar entregas parciais.');
  });

  (r.vitorias_sinalizadas || []).slice(0, 2).forEach(v => {
    add(v.titulo || 'Frente com vitória', v.status || 'Em andamento', 'Vitória sinalizada pela equipe', 'Documentar resultado e avaliar replicabilidade.');
  });

  const andamento = (r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []).slice(0, 2);
  andamento.forEach(d => {
    add(d.titulo || 'Frente em andamento', 'Em andamento', 'Frente ativa com impacto operacional direto', 'Acompanhar progresso e garantir continuidade.');
  });

  return result;
}

function buildNextStepsFallback({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const existentes = Array.isArray(n.proximos_passos_sugeridos) ? n.proximos_passos_sugeridos : [];
  if (existentes.length >= 2) return existentes.slice(0, 6);

  const result = [...existentes];
  const seen = new Set(existentes.map(e => (e.titulo || '').toLowerCase()));

  const add = (titulo, descricao) => {
    if (result.length >= 6) return;
    if (seen.has(titulo.toLowerCase())) return;
    seen.add(titulo.toLowerCase());
    result.push({ titulo, descricao });
  };

  const andamento = (r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []);
  if (andamento.length > 0) {
    add('Acompanhar demandas em andamento', `Monitorar evolução das ${andamento.length} frente${andamento.length > 1 ? 's' : ''} ativas e garantir continuidade.`);
  }

  const wins = (r.vitorias_sinalizadas || []);
  if (wins.length > 0) {
    add('Consolidar vitórias registradas', 'Documentar os resultados sinalizados com #WIN e avaliar replicabilidade nas demais frentes.');
  }

  const concluidas = (r.entregas_concluidas || r.demandas_por_status?.concluidas || []);
  if (concluidas.length > 0) {
    add('Revisar entregas concluídas', 'Validar qualidade e registrar aprendizados das demandas encerradas no período.');
  }

  const top = (r.top_demandas_por_tempo || [])[0];
  if (top) {
    add(`Monitorar: ${top.titulo}`, 'Demanda com maior tempo registrado — acompanhar status e avaliar necessidade de repriorizacão.');
  }

  const outras = (r.demandas_por_status?.outras || []);
  if (outras.length > 0) {
    add('Organizar backlog', `Revisar as ${outras.length} demanda${outras.length > 1 ? 's' : ''} no backlog e definir prioridade para a próxima semana.`);
  }

  add('Atualizar registro de atividades', 'Garantir que todos os apontamentos da semana estejam documentados e com comentários atualizados.');

  return result;
}

function buildExecutiveClosing({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const candidato = n.fechamento_executivo || n.texto_email || '';
  if (!isWeakText(candidato)) return candidato;

  const concluidas = (r.entregas_concluidas || r.demandas_por_status?.concluidas || []).length;
  const andamento = (r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []).length;
  const total = r.total_horas || '';
  const wins = (r.vitorias_sinalizadas || []).length;

  let partes = [];
  if (concluidas > 0 || andamento > 0) {
    const c = concluidas > 0 ? `${concluidas} entrega${concluidas > 1 ? 's' : ''} concluída${concluidas > 1 ? 's' : ''}` : '';
    const a = andamento > 0 ? `${andamento} frente${andamento > 1 ? 's' : ''} em andamento` : '';
    partes.push(`A semana encerra com ${[c, a].filter(Boolean).join(' e ')}` + (total ? `, totalizando ${total} de esforço registrado` : '') + '.');
  }
  if (wins > 0) {
    partes.push(`${wins > 1 ? `${wins} vitórias foram sinalizadas` : 'Uma vitória foi sinalizada'} pela operação, evidenciando entregas de valor além do registro técnico.`);
  }
  partes.push('O registro contínuo sustenta o acompanhamento gerencial e a rastreabilidade das atividades da equipe.');
  return partes.join(' ');
}

function buildEnrichedDeliveries({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const iaItems = Array.isArray(n.principais_entregas) ? n.principais_entregas : [];
  const seen = new Set();
  const result = [];

  const push = (item) => {
    if (result.length >= 9) return;
    const key = (item.demanda_id ? String(item.demanda_id) : '') || (item.titulo || '').toLowerCase().slice(0, 30);
    if (!key || seen.has(key)) return;
    seen.add(key);
    result.push(item);
  };

  iaItems.forEach(push);

  (r.entregas_concluidas || r.demandas_por_status?.concluidas || []).forEach(d => {
    const descFallback =
      d.departamento === 'Fiscal' ? 'Entrega concluída com impacto na frente Fiscal, removendo pendência operacional do período.' :
      d.origem === 'Chat Teams' ? 'Atendimento originado no Teams, com resolução direta da necessidade do usuário.' :
      d.tipo === 'Chamado' ? 'Chamado concluído no período, contribuindo para a sustentação operacional e continuidade dos serviços.' :
      'Demanda encerrada no período, com registro de conclusão no FocusTrack.';
    push({ titulo: d.titulo, descricao: d.comentarios?.length ? d.comentarios[0] : descFallback, demanda_id: d.id, tempo: d.tempo_horas });
  });

  return result;
}

function buildEnrichedInProgress({ narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  const iaItems = Array.isArray(n.demandas_em_andamento) ? n.demandas_em_andamento : [];
  const winsById = new Set((r.vitorias_sinalizadas || []).map(v => String(v.demanda_id || '')).filter(Boolean));
  const seen = new Set();
  const result = [];

  const push = (item) => {
    if (result.length >= 8) return;
    const key = (item.demanda_id ? String(item.demanda_id) : '') || (item.titulo || '').toLowerCase().slice(0, 30);
    if (!key || seen.has(key)) return;
    seen.add(key);
    result.push(item);
  };

  iaItems.forEach(push);

  (r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []).forEach(d => {
    const temWin = winsById.has(String(d.id || ''));
    const temComentarios = Array.isArray(d.comentarios) && d.comentarios.length > 1;
    const descFallback = temWin
      ? 'Frente em andamento com vitória parcial já sinalizada no período. Evolução relevante registrada.'
      : temComentarios
        ? 'Demanda com detalhamento extenso nos registros — evolução e acompanhamento ativos no período.'
        : 'Demanda em progresso com apontamentos registrados na semana.';
    push({ titulo: d.titulo, descricao: descFallback, demanda_id: d.id, tempo: d.tempo_horas });
  });

  return result;
}

function buildWeeklyReportEmailHtml({ meta, narrativa, resumoParaIa }) {
  const n = narrativa || {};
  const r = resumoParaIa || {};

  // ── constantes de URLs ──────────────────────────────────────
  const TEAMS_LINK = 'https://teams.microsoft.com/l/entity/26bc2873-6023-480c-a11b-76b66605ce8c/_djb2_msteams_prefix_2810599455?context=%7B%22channelId%22%3A%2219%3A8m7ZGscFI8DiTnbcgxhzOc9_1aqEYMtlZtw6odWaShw1%40thread.tacv2%22%7D&tenantId=d2f6807c-d369-4371-93d4-16a3561f25f2';
  const LOGO_FRANCO = 'https://raw.githubusercontent.com/ffqmt/Images/15116cdbaa87af68eb9eaf9a1bea9ee7502bb9f7/FRANCO%20LOGO.png';
  const LOGO_CONTAUDI = 'https://raw.githubusercontent.com/ffqmt/Images/15116cdbaa87af68eb9eaf9a1bea9ee7502bb9f7/LOGO%20IMAGEM.png';

  // ── conteúdo enriquecido ────────────────────────────────────
  const capaTexto = escapeHtml(buildExecutiveCoverText({ meta, narrativa, resumoParaIa }));
  const leituraGestao = escapeHtml(buildManagementReading({ narrativa, resumoParaIa }));
  const fechamentoTexto = escapeHtml(buildExecutiveClosing({ narrativa, resumoParaIa }));

  // ── vitórias: IA > sinais_manuais.wins > vitorias_sinalizadas ──
  const vitoriasArr = (() => {
    const ia = Array.isArray(n.vitorias_da_semana) ? n.vitorias_da_semana : [];
    if (ia.length) return ia;
    const wins = (r.sinais_manuais?.wins || []).map(v => ({ titulo: v.titulo, descricao: v.texto_limpo || v.leitura_executiva, demanda_id: v.demanda_id, origem: v.origem }));
    if (wins.length) return wins;
    return (r.vitorias_sinalizadas || []).map(v => ({ titulo: v.titulo, descricao: v.texto, demanda_id: v.demanda_id, origem: v.origem }));
  })();

  // ── carteira: IA > carteira_executiva.por_frente > buildPortfolioByFront ──
  const carteiraArr = (() => {
    const ia = Array.isArray(n.carteira_por_frente) ? n.carteira_por_frente : [];
    if (ia.length) return ia;
    const exec = ((r.carteira_executiva || {}).por_frente || []).map(f => ({ frente: f.frente, descricao: f.leitura }));
    if (exec.length) return exec;
    return buildPortfolioByFront({ resumoParaIa });
  })();

  // ── pontos de atenção: IA > sinais riscos/decisoes ─────────
  const pontosAtencaoArr = (() => {
    const ia = Array.isArray(n.pontos_de_atencao) ? n.pontos_de_atencao : [];
    if (ia.length) return ia;
    return [
      ...(r.sinais_manuais?.riscos || []).map(s => ({ titulo: s.titulo || 'Risco identificado', descricao: s.texto_limpo || s.leitura_executiva })),
      ...(r.sinais_manuais?.decisoes || []).map(s => ({ titulo: s.titulo || 'Decisão pendente', descricao: s.texto_limpo || s.leitura_executiva }))
    ].slice(0, 6);
  })();

  // ── próximos passos: IA > sinais.proximos > buildNextStepsFallback ──
  const proximosArr = (() => {
    const ia = Array.isArray(n.proximos_passos_sugeridos) ? n.proximos_passos_sugeridos : [];
    if (ia.length) return ia;
    const prox = (r.sinais_manuais?.proximos || []).map(s => ({ titulo: s.titulo || 'Próximo passo', descricao: s.texto_limpo || s.leitura_executiva }));
    if (prox.length) return prox;
    return buildNextStepsFallback({ narrativa, resumoParaIa });
  })();

  // ── projetos de gestão: IA > temas_da_semana > buildManagementProjects ──
  const projetosArr = (() => {
    const ia = Array.isArray(n.projetos_interesse_gestao) ? n.projetos_interesse_gestao : [];
    if (ia.length) return ia;
    const temas = (r.temas_da_semana || []).map(t => ({ frente: t.tema, status: 'Em acompanhamento', impacto: t.impacto_gestao, proximo_passo: t.evidencia }));
    if (temas.length) return temas;
    return buildManagementProjects({ narrativa, resumoParaIa });
  })();

  const avancosArr = buildStructuralAdvances({ narrativa, resumoParaIa });
  const entregasArr = buildEnrichedDeliveries({ narrativa, resumoParaIa });
  const andamentoArr = buildEnrichedInProgress({ narrativa, resumoParaIa });

  // ── renderização ────────────────────────────────────────────
  const vitoriasHtml = renderVictoryItems(vitoriasArr);
  const entregasHtml = renderEmailItems(entregasArr, { limit: 9, borderColor: '#2fbf88', titleColor: '#0d4a2e' });
  const andamentoHtml = renderEmailItems(andamentoArr, { limit: 8, borderColor: '#1a73be', titleColor: '#0d2b4c' });
  const avancosHtml = renderEmailItems(avancosArr, { limit: 4, borderColor: '#0d2b4c', titleColor: '#0d2b4c' });
  const carteiraHtml = renderEmailItems(carteiraArr, { limit: 6, borderColor: '#1a73be', titleColor: '#0d2b4c' });
  const atencaoHtml = renderAttentionItems(pontosAtencaoArr);
  const projetosHtml = renderProjectsTable(projetosArr);
  const proximosHtml = renderNextSteps(proximosArr);

  // ── métricas ────────────────────────────────────────────────
  const totalApontamentos = String(meta.total_apontamentos || r.total_apontamentos || 0);
  const totalConcluidas = String((r.entregas_concluidas || r.demandas_por_status?.concluidas || []).length);
  const totalAndamento = String((r.demandas_em_andamento || r.demandas_por_status?.em_andamento || []).length);
  const totalBacklog = String(
    Array.isArray(r.backlog_executivo) && r.backlog_executivo.length > 0
      ? r.backlog_executivo.length
      : (r.demandas_por_status?.outras || []).length
  );

  // ── leitura 30s ─────────────────────────────────────────────
  const l30 = n.leitura_30s || {};
  const l30avancou = escapeHtml(l30.o_que_avancou || '');
  const l30entregue = escapeHtml(l30.o_que_foi_entregue || '');
  const l30gestao = escapeHtml(l30.o_que_exige_gestao || '');

  // ── textos de cabeçalho ─────────────────────────────────────
  const periodo = formatPeriodForEmail(meta);
  const geradoEm = escapeHtml(meta.gerado_em_br || '');
  const semanaLabel = escapeHtml(meta.semana?.label || r.periodo || '');
  const contextoTitulo = escapeHtml(n.contexto_operacional?.titulo || 'Contexto operacional da semana');
  const contextoTexto = escapeHtml(n.contexto_operacional?.texto || r.base_narrativa || '');

  // ── flags de visibilidade ───────────────────────────────────
  const showVitorias = vitoriasArr.length > 0;
  const showAvancos = avancosArr.length > 0;
  const showCarteira = carteiraArr.length > 0;
  const showProjetos = projetosArr.length > 0;

  // ── estilos reutilizáveis ───────────────────────────────────
  const SEC = 'padding:30px 34px;border-bottom:1px solid #d9e2ec;';
  const SEC_LABEL = 'margin:0 0 6px 0;font-size:10px;color:#1a73be;letter-spacing:3px;text-transform:uppercase;font-weight:bold;font-family:Arial,sans-serif;';
  const SEC_TITLE = 'margin:0 0 20px 0;font-size:20px;line-height:26px;color:#0d2b4c;font-weight:bold;font-family:Arial,sans-serif;';
  const COL3 = 'vertical-align:top;border-top:3px solid ';

  const sectionHeader = (label, title) =>
    `<p style="${SEC_LABEL}">${label}</p><h2 style="${SEC_TITLE}">${title}</h2>`;

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>Relat&oacute;rio Semanal &mdash; FocusTrack</title>
</head>
<body style="margin:0;padding:0;background-color:#edf1f6;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#edf1f6">
<tr><td align="center" style="padding:28px 16px;">

<table width="800" cellpadding="0" cellspacing="0" border="0" style="width:800px;max-width:800px;background-color:#ffffff;border-collapse:collapse;border:1px solid #d9e2ec;">

  <!-- CABEÇALHO INSTITUCIONAL -->
  <tr>
    <td style="padding:22px 34px 18px 34px;background-color:#ffffff;border-bottom:1px solid #d9e2ec;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="vertical-align:middle;">
            <img src="${LOGO_FRANCO}" alt="Franco Sistemas" width="118" style="display:block;width:118px;max-width:118px;border:0;">
          </td>
          <td width="18" style="vertical-align:middle;border-left:1px solid #d9e2ec;padding-left:18px;">
            <img src="${LOGO_CONTAUDI}" alt="Contaudi" width="122" style="display:block;width:122px;max-width:122px;border:0;">
          </td>
          <td align="right" style="vertical-align:middle;">
            <div style="font-size:10px;color:#8a9ab5;font-family:Arial,sans-serif;text-align:right;">Uso interno &mdash; gest&atilde;o</div>
            <div style="font-size:10px;color:#b0bcc8;font-family:Arial,sans-serif;text-align:right;margin-top:4px;">Gerado em ${geradoEm} &middot; Cuiab&aacute;/MT</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- LINHA COLORIDA -->
  <tr>
    <td style="padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="44%" style="height:5px;background-color:#0d2b4c;font-size:0;">&nbsp;</td>
          <td width="22%" style="height:5px;background-color:#1a73be;font-size:0;">&nbsp;</td>
          <td width="18%" style="height:5px;background-color:#2fbf88;font-size:0;">&nbsp;</td>
          <td width="16%" style="height:5px;background-color:#f2b84b;font-size:0;">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- CAPA EXECUTIVA -->
  <tr>
    <td style="padding:34px 42px 30px 42px;background-color:#f8fafd;border-bottom:1px solid #d9e2ec;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="vertical-align:top;" width="62%">
            <p style="margin:0 0 10px 0;font-size:10px;color:#1a73be;letter-spacing:3px;text-transform:uppercase;font-weight:bold;font-family:Arial,sans-serif;">Relat&oacute;rio semanal</p>
            <h1 style="margin:0 0 16px 0;font-size:30px;line-height:38px;color:#0d2b4c;font-weight:bold;letter-spacing:-0.6px;font-family:Arial,sans-serif;">Entregas, Carteira<br>e Automa&ccedil;&atilde;o</h1>
            <p style="margin:0;font-size:13px;color:#3d5370;font-family:Arial,sans-serif;line-height:1.7;">${capaTexto}</p>
          </td>
          <td width="4%"></td>
          <td style="vertical-align:top;" width="34%">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #d9e2ec;background-color:#ffffff;">
              <tr><td style="padding:20px;text-align:center;">
                <div style="font-size:9px;color:#8a9ab5;font-family:Arial,sans-serif;text-transform:uppercase;letter-spacing:2px;margin-bottom:10px;">Per&iacute;odo</div>
                <div style="font-size:18px;font-weight:bold;color:#0d2b4c;font-family:Arial,sans-serif;line-height:1.4;">${periodo}</div>
                <div style="margin-top:10px;padding-top:10px;border-top:1px solid #e5eaf0;font-size:11px;color:#6b7a90;font-family:Arial,sans-serif;">${semanaLabel}</div>
              </td></tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- LEITURA 30s -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Gest&atilde;o', 'Leitura em 30 segundos')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="32%" style="${COL3}#1a73be;padding:14px 14px 14px 14px;background-color:#f7fbff;vertical-align:top;">
            <div style="font-size:10px;font-weight:bold;color:#1a73be;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;font-family:Arial,sans-serif;">O que avan&ccedil;ou</div>
            <div style="font-size:12px;color:#2d3e52;line-height:1.6;font-family:Arial,sans-serif;">${l30avancou}</div>
          </td>
          <td width="2%"></td>
          <td width="32%" style="${COL3}#2fbf88;padding:14px;background-color:#f4fbf7;vertical-align:top;">
            <div style="font-size:10px;font-weight:bold;color:#2fbf88;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;font-family:Arial,sans-serif;">O que foi entregue</div>
            <div style="font-size:12px;color:#2d3e52;line-height:1.6;font-family:Arial,sans-serif;">${l30entregue}</div>
          </td>
          <td width="2%"></td>
          <td width="32%" style="${COL3}#f2b84b;padding:14px;background-color:#fffcf0;vertical-align:top;">
            <div style="font-size:10px;font-weight:bold;color:#c48b00;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;font-family:Arial,sans-serif;">O que exige gest&atilde;o</div>
            <div style="font-size:12px;color:#2d3e52;line-height:1.6;font-family:Arial,sans-serif;">${l30gestao}</div>
          </td>
        </tr>
      </table>
      ${leituraGestao ? `<p style="margin:16px 0 0 0;font-size:12px;color:#4a5a70;font-family:Arial,sans-serif;line-height:1.7;padding:14px;background-color:#f7f9fc;border-left:3px solid #8aa4c0;">${leituraGestao}</p>` : ''}
    </td>
  </tr>

  <!-- MOVIMENTO DA CARTEIRA -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('M&eacute;tricas', 'Movimento da carteira')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          ${renderMetricCell(totalApontamentos, 'Registros da semana', '#1a73be')}
          ${renderMetricCell(totalConcluidas, 'Entregas concluídas', '#2fbf88')}
          ${renderMetricCell(totalAndamento, 'Frentes em andamento', '#0d2b4c')}
          ${renderMetricCell(totalBacklog, 'Backlog / triagem', '#f2b84b')}
        </tr>
      </table>
      <p style="margin:16px 0 0 0;font-size:12px;color:#4a5a70;font-family:Arial,sans-serif;line-height:1.7;padding:14px;background-color:#f7f9fc;border-left:3px solid #8aa4c0;"><strong>Leitura de gest&atilde;o:</strong> A base formal de chamados representa a movimenta&ccedil;&atilde;o registrada no FocusTrack. Tamb&eacute;m devem ser considerados coment&aacute;rios, vit&oacute;rias sinalizadas, automa&ccedil;&otilde;es, suporte e frentes estruturais registradas no per&iacute;odo.</p>
    </td>
  </tr>

  <!-- ACOMPANHAMENTO CONTÍNUO -->
  <tr>
    <td style="padding:0;border-bottom:1px solid #d9e2ec;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f0f5fa;">
        <tr>
          <td style="padding:22px 34px;">
            <p style="margin:0 0 4px 0;font-size:10px;color:#6b7a90;letter-spacing:2px;text-transform:uppercase;font-family:Arial,sans-serif;">Acompanhamento cont&iacute;nuo</p>
            <p style="margin:0 0 10px 0;font-size:15px;font-weight:bold;color:#0d2b4c;font-family:Arial,sans-serif;">Lista oficial de demandas</p>
            <p style="margin:0 0 16px 0;font-size:12px;color:#4a5a70;font-family:Arial,sans-serif;line-height:1.6;">Acesse o painel centralizado de demandas para acompanhar status, prazos e hist&oacute;rico de atividades em tempo real.</p>
            <table cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td style="background-color:#0d2b4c;padding:10px 22px;">
                  <a href="${TEAMS_LINK}" target="_blank" style="font-size:13px;font-weight:bold;color:#ffffff;font-family:Arial,sans-serif;text-decoration:none;">Abrir lista de demandas</a>
                </td>
              </tr>
            </table>
            <p style="margin:10px 0 0 0;font-size:10px;color:#8a9ab5;font-family:Arial,sans-serif;">Acesso via Microsoft Teams</p>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- CONTEXTO OPERACIONAL -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Contexto', contextoTitulo)}
      <p style="margin:0;font-size:13px;color:#3d5370;font-family:Arial,sans-serif;line-height:1.8;">${contextoTexto}</p>
    </td>
  </tr>

  ${showVitorias ? `<!-- VITÓRIAS -->
  <tr>
    <td style="${SEC}background-color:#f4fbf7;">
      ${sectionHeader('Destaque da semana', 'Vit&oacute;rias da semana')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        ${vitoriasHtml}
      </table>
    </td>
  </tr>` : ''}

  <!-- ENTREGAS CONCLUÍDAS -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Resultados', 'Entregas conclu&iacute;das na semana')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        ${entregasHtml || `<tr><td style="font-size:13px;color:#6b7a90;font-family:Arial,sans-serif;padding:10px 0;">Nenhuma entrega conclu&iacute;da registrada neste per&iacute;odo.</td></tr>`}
      </table>
    </td>
  </tr>

  <!-- DEMANDAS EM ANDAMENTO -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Carteira ativa', 'Demandas em andamento e projetos de melhoria')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        ${andamentoHtml || `<tr><td style="font-size:13px;color:#6b7a90;font-family:Arial,sans-serif;padding:10px 0;">Nenhuma demanda em andamento no per&iacute;odo.</td></tr>`}
      </table>
    </td>
  </tr>

  ${showAvancos ? `<!-- AVANÇOS ESTRUTURAIS -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Evolu&ccedil;&atilde;o', 'Avan&ccedil;os estruturais')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        ${avancosHtml}
      </table>
    </td>
  </tr>` : ''}

  ${showCarteira ? `<!-- CARTEIRA POR FRENTE -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Distribui&ccedil;&atilde;o', 'Carteira ativa por frente')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        ${carteiraHtml}
      </table>
    </td>
  </tr>` : ''}

  <!-- PONTOS DE ATENÇÃO -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Riscos e depend&ecirc;ncias', 'Pontos de aten&ccedil;&atilde;o e depend&ecirc;ncias')}
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        ${atencaoHtml}
      </table>
    </td>
  </tr>

  ${showProjetos ? `<!-- PROJETOS DE GESTÃO -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Gest&atilde;o estrat&eacute;gica', 'Projetos de interesse da gest&atilde;o')}
      ${projetosHtml}
    </td>
  </tr>` : ''}

  <!-- PRÓXIMOS PASSOS -->
  <tr>
    <td style="${SEC}">
      ${sectionHeader('Planejamento', 'Foco planejado &mdash; pr&oacute;xima semana')}
      ${proximosHtml}
    </td>
  </tr>

  <!-- FECHAMENTO EXECUTIVO -->
  <tr>
    <td style="padding:0;border-bottom:1px solid #d9e2ec;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#0d2b4c;">
        <tr>
          <td width="6" style="background-color:#1a73be;font-size:0;">&nbsp;</td>
          <td style="padding:26px 30px;">
            <p style="margin:0 0 8px 0;font-size:10px;color:#7fb3d3;letter-spacing:2px;text-transform:uppercase;font-family:Arial,sans-serif;">Fechamento executivo</p>
            <p style="margin:0;font-size:13px;color:#c8ddf0;font-family:Arial,sans-serif;line-height:1.8;">${fechamentoTexto}</p>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- RODAPÉ -->
  <tr>
    <td style="padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="44%" style="height:3px;background-color:#0d2b4c;font-size:0;">&nbsp;</td>
          <td width="22%" style="height:3px;background-color:#1a73be;font-size:0;">&nbsp;</td>
          <td width="18%" style="height:3px;background-color:#2fbf88;font-size:0;">&nbsp;</td>
          <td width="16%" style="height:3px;background-color:#f2b84b;font-size:0;">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f7f9fc">
        <tr>
          <td style="padding:20px 34px;text-align:center;border-top:1px solid #d9e2ec;">
            <div style="font-size:12px;font-weight:bold;color:#0d2b4c;font-family:Arial,sans-serif;">Franco Sistemas &middot; Contaudi Assessoria Cont&aacute;bil</div>
            <div style="font-size:11px;color:#6b7a90;font-family:Arial,sans-serif;margin-top:5px;">Relat&oacute;rio semanal de entregas, carteira, documenta&ccedil;&atilde;o e automa&ccedil;&atilde;o</div>
          </td>
        </tr>
      </table>
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
