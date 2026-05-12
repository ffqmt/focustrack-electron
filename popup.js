let currentTicket = null;
let timerInterval = null;
let searchTicketsCache = [];

function truncate(text, max = 40) {
  if (!text) return '';
  return text.length > max ? `${text.slice(0, max)}...` : text;
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatTime(ms) {
  const totalSeconds = Math.max(0, Math.floor(Number(ms || 0) / 1000));
  const hours = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
  const minutes = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
  const seconds = String(totalSeconds % 60).padStart(2, '0');
  return `${hours}:${minutes}:${seconds}`;
}

function getStatusMeta(status) {
  switch (status) {
    case 'running':
      return {
        label: 'Em andamento',
        dotClass: 'status-dot status-dot--active',
        cardClass: 'current-card--active'
      };

    case 'paused':
      return {
        label: 'Pausado',
        dotClass: 'status-dot status-dot--paused',
        cardClass: 'current-card--paused'
      };

    case 'selected':
      return {
        label: 'Selecionado',
        dotClass: 'status-dot status-dot--waiting',
        cardClass: 'current-card--selected'
      };

    case 'finished':
      return {
        label: 'Concluído',
        dotClass: 'status-dot status-dot--waiting',
        cardClass: 'current-card--finished'
      };

    default:
      return {
        label: 'Selecionado',
        dotClass: 'status-dot status-dot--waiting',
        cardClass: 'current-card--selected'
      };
  }
}

function normalizeTicket(ticket) {
  if (!ticket) return null;

  return {
    id: Number(ticket.id || ticket.ticketId),
    title: ticket.title || ticket.ticketTitle || '',
    team: ticket.team || 'Geral'
  };
}

/* =========================
   Teams
========================= */

function getTeamsDemandsUrl() {
  const config = window.FocusTrackConfigStore?.load?.() || window.FocusTrackConfig || {};

  return (
    config?.links?.teamsDemandsUrl ||
    window.FocusTrackConfig?.links?.teamsDemandsUrl ||
    ''
  );
}

async function openTeamsDemands() {
  const url = getTeamsDemandsUrl();

  console.log('[FocusTrack] Clique no botão Teams detectado.');
  console.log('[FocusTrack] URL Teams:', url);
  console.log('[FocusTrack] focusTrack disponível:', window.focusTrack);

  if (!url || url === 'COLE_AQUI_O_LINK_DO_TEAMS') {
    await window.FocusTrackUI.alert({
      type: 'warning',
      title: 'Link do Teams não configurado',
      message: 'Configure o link da lista de demandas no config.js.',
      confirmText: 'Entendi'
    });

    return;
  }

  try {
    if (window.focusTrack?.openExternal) {
      const result = await window.focusTrack.openExternal(url);

      console.log('[FocusTrack] Resultado openExternal:', result);

      if (!result?.ok) {
        throw new Error(result?.message || 'Não foi possível abrir o link do Teams.');
      }

      return;
    }

    console.warn('[FocusTrack] window.focusTrack.openExternal não encontrado. Tentando window.open.');
    window.open(url, '_blank');
  } catch (error) {
    console.error('[FocusTrack] Erro ao abrir Teams:', error);

    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Erro ao abrir Teams',
      message: error?.message || 'Não foi possível abrir a lista de demandas.',
      confirmText: 'Entendi'
    });
  }
}

/* =========================
   Settings
========================= */

function getSettingsElements() {
  return {
    modal: document.getElementById('settingsModal'),
    backdrop: document.getElementById('settingsBackdrop'),
    settingsBtn: document.getElementById('settingsBtn'),
    closeBtn: document.getElementById('closeSettingsBtn'),
    cancelBtn: document.getElementById('cancelSettingsBtn'),
    testBtn: document.getElementById('testSettingsBtn'),
    saveBtn: document.getElementById('saveSettingsBtn'),
    userEmail: document.getElementById('configUserEmail'),
    status: document.getElementById('settingsInlineStatus')
  };
}

function setSettingsStatus(type, message) {
  const statusEl = document.getElementById('settingsInlineStatus');
  if (!statusEl) return;

  if (!message) {
    statusEl.textContent = '';
    statusEl.className = 'settings-inline-status hidden';
    return;
  }

  statusEl.textContent = message;
  statusEl.className = `settings-inline-status settings-inline-status--${type}`;
}

function formatConnectionResult(result) {
  if (!result) {
    return 'Sem resposta no teste de conexão.';
  }

  const lines = [];

  if (Array.isArray(result.checks)) {
    for (const check of result.checks) {
      lines.push(`${check.ok ? '✅' : '❌'} ${check.message}`);
    }
  }

  if (!lines.length && result.message) {
    lines.push(result.message);
  }

  return lines.join(' | ');
}

function readSettingsForm() {
  const els = getSettingsElements();

  return {
    userEmail: els.userEmail?.value?.trim() || ''
  };
}

function writeSettingsForm(data) {
  const els = getSettingsElements();

  if (els.userEmail) {
    els.userEmail.value = data?.userEmail || '';
  }
}

function validateSettings(config) {
  if (!config.userEmail) {
    return { valid: false, message: 'Informe o e-mail do usuário.' };
  }

  const emailOk = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(config.userEmail);

  if (!emailOk) {
    return { valid: false, message: 'Informe um e-mail válido.' };
  }

  return { valid: true };
}

function loadSettings() {
  const config = window.FocusTrackConfigStore?.load?.() || window.FocusTrackConfig || {};

  writeSettingsForm({
    userEmail: config?.currentUser?.email || ''
  });

  setSettingsStatus('', '');
}

function saveSettings() {
  const config = readSettingsForm();
  const validation = validateSettings(config);

  if (!validation.valid) {
    setSettingsStatus('error', validation.message);
    return false;
  }

  if (!window.FocusTrackConfigStore?.updateCurrentUser) {
    setSettingsStatus('error', 'Não foi possível salvar o e-mail.');
    return false;
  }

  window.FocusTrackConfigStore.updateCurrentUser({
    email: config.userEmail
  });

  setSettingsStatus('success', 'E-mail salvo com sucesso.');
  return true;
}

async function testSettings() {
  const formConfig = readSettingsForm();
  const validation = validateSettings(formConfig);

  if (!validation.valid) {
    setSettingsStatus('error', validation.message);
    return;
  }

  if (!window.focusTrack?.testConnection) {
    setSettingsStatus('error', 'Teste de conexão indisponível.');
    return;
  }

  const localConfig = window.FocusTrackConfigStore?.load?.() || window.FocusTrackConfig || {};

  const payload = {
    userEmail: formConfig.userEmail,
    sharePoint: {
      siteUrl: localConfig?.sharePoint?.siteUrl || '',
      lists: {
        tickets: localConfig?.sharePoint?.lists?.tickets || '',
        timeEntries: localConfig?.sharePoint?.lists?.timeEntries || '',
        ticketsId: localConfig?.sharePoint?.lists?.ticketsId || '',
        timeEntriesId: localConfig?.sharePoint?.lists?.timeEntriesId || ''
      }
    }
  };

  setSettingsStatus('loading', 'Testando conexão com SharePoint...');

  try {
    const result = await window.focusTrack.testConnection(payload);

    if (result?.ok) {
      setSettingsStatus(
        'success',
        formatConnectionResult(result) || 'Conexão validada com sucesso.'
      );
    } else {
      setSettingsStatus(
        'error',
        formatConnectionResult(result) || 'Falha ao testar conexão.'
      );
    }
  } catch (error) {
    setSettingsStatus('error', error?.message || 'Erro ao testar conexão.');
  }
}

function openSettingsModal() {
  loadSettings();
  const els = getSettingsElements();
  els.modal?.classList.remove('hidden');
}

function closeSettingsModal() {
  const els = getSettingsElements();
  els.modal?.classList.add('hidden');
  setSettingsStatus('', '');
}

/* =========================
   Busca de chamados
========================= */

async function renderResults(query) {
  const searchResults = document.getElementById('searchResults');
  const newTicketBtn = document.getElementById('newTicketBtn');
  const trimmed = String(query || '').trim();

  if (!searchResults || !newTicketBtn) return;

  if (!trimmed) {
    searchTicketsCache = [];
    searchResults.innerHTML = '';
    searchResults.classList.add('hidden');
    newTicketBtn.classList.add('new-ticket-btn--inactive');
    newTicketBtn.classList.remove('new-ticket-btn--active');
    return;
  }

  try {
    const filtered = await window.FocusTrackTickets.searchTickets(trimmed);
    searchTicketsCache = Array.isArray(filtered) ? filtered : [];
  } catch (error) {
    console.error('Erro ao buscar tickets:', error);
    searchTicketsCache = [];
  }

  newTicketBtn.classList.add('new-ticket-btn--active');
  newTicketBtn.classList.remove('new-ticket-btn--inactive');

  let html = `
    <button class="result-row result-create" type="button">
      <span class="result-create-left">
        <span class="result-badge">✣</span>
        <span class="result-create-text">
          Criar novo chamado <strong>"${escapeHtml(trimmed)}"</strong>
        </span>
      </span>
      <span class="result-arrow">→</span>
    </button>
  `;

  if (searchTicketsCache.length === 0) {
    html += `
      <div class="result-empty">
        Nenhum chamado encontrado com esse título
      </div>
    `;
  } else {
    html += searchTicketsCache
      .slice(0, 6)
      .map((ticket) => {
        const id = Number(ticket.id);
        const safeTitle = escapeHtml(ticket.title || '');
        const safeTeam = escapeHtml(ticket.team || 'Geral');

        return `
          <button
            class="result-row result-ticket"
            type="button"
            data-ticket-id="${id}"
            data-ticket-title="${safeTitle}"
            data-ticket-team="${safeTeam}"
          >
            <span class="result-hash">#</span>
            <span class="result-ticket-text">
              ${id} ${escapeHtml(truncate(ticket.title))}
            </span>
          </button>
        `;
      })
      .join('');
  }

  searchResults.innerHTML = html;
  searchResults.classList.remove('hidden');
}

/* =========================
   Worklist
========================= */

function getWorkItems() {
  const worklist = window.FocusTrackFlows?.getWorklist?.() || [];
  const sessions = window.FocusTrackFlows?.getAllSessions?.() || {};

  const map = new Map();

  for (const item of worklist) {
    const normalized = normalizeTicket(item);
    if (!normalized?.id) continue;

    map.set(String(normalized.id), {
      ...normalized,
      source: 'worklist'
    });
  }

  for (const session of Object.values(sessions || {})) {
    if (!session?.ticketId) continue;

    const id = Number(session.ticketId);
    const key = String(id);
    const existing = map.get(key) || {};

    map.set(key, {
      id,
      title: session.ticketTitle || existing.title || '',
      team: session.team || existing.team || 'Geral',
      source: existing.source || 'session'
    });
  }

  return Array.from(map.values());
}

function getTicketSession(ticketId) {
  if (!ticketId) return null;
  return window.FocusTrackFlows?.getSession?.(ticketId) || null;
}

function getTicketStatus(ticket) {
  const session = getTicketSession(ticket.id);

  if (session?.status) {
    return session.status;
  }

  return 'selected';
}

function getTicketElapsedMs(ticket) {
  const session = getTicketSession(ticket.id);

  if (session) {
    return Number(window.FocusTrackFlows.getElapsedMs(session) || 0);
  }

  return 0;
}

function setCurrentTicket(ticket) {
  const normalized = normalizeTicket(ticket);

  if (!normalized?.id) {
    currentTicket = null;
    return;
  }

  const session = getTicketSession(normalized.id);

  currentTicket = {
    id: normalized.id,
    title: session?.ticketTitle || normalized.title || '',
    team: session?.team || normalized.team || 'Geral',
    elapsedMs: session ? Number(session.elapsedMs || 0) : 0,
    uiStatus: session?.status || 'selected'
  };
}

function selectFirstAvailableTicket() {
  const items = getWorkItems();

  if (!items.length) {
    currentTicket = null;
    return;
  }

  const sessions = window.FocusTrackFlows?.getAllSessions?.() || {};
  const sessionList = Object.values(sessions || {});

  const running = sessionList.find((session) => session.status === 'running');
  const paused = sessionList.find((session) => session.status === 'paused');

  if (running) {
    setCurrentTicket({
      id: running.ticketId,
      title: running.ticketTitle,
      team: running.team
    });
    return;
  }

  if (paused) {
    setCurrentTicket({
      id: paused.ticketId,
      title: paused.ticketTitle,
      team: paused.team
    });
    return;
  }

  setCurrentTicket(items[0]);
}

function renderSummary(items) {
  const statusDot = document.getElementById('currentStatusDot');
  const statusLabel = document.getElementById('currentStatusLabel');
  const currentTime = document.getElementById('currentTime');

  if (!statusDot || !statusLabel || !currentTime) return;

  const sessions = window.FocusTrackFlows?.getAllSessions?.() || {};
  const sessionList = Object.values(sessions || {});

  const runningCount = sessionList.filter((session) => session.status === 'running').length;
  const pausedCount = sessionList.filter((session) => session.status === 'paused').length;

  const totalElapsedMs = sessionList.reduce((acc, session) => {
    return acc + Number(window.FocusTrackFlows.getElapsedMs(session) || 0);
  }, 0);

  if (!items.length) {
    statusDot.className = 'status-dot status-dot--waiting';
    statusLabel.textContent = 'Nenhum chamado';
    currentTime.textContent = '00:00:00';
    return;
  }

  if (runningCount > 0) {
    statusDot.className = 'status-dot status-dot--active';
  } else if (pausedCount > 0) {
    statusDot.className = 'status-dot status-dot--paused';
  } else {
    statusDot.className = 'status-dot status-dot--waiting';
  }

  const totalLabel = items.length === 1 ? '1 chamado' : `${items.length} chamados`;

  statusLabel.textContent = `${totalLabel} · ${runningCount} rodando · ${pausedCount} pausado(s)`;
  currentTime.textContent = formatTime(totalElapsedMs);
}

function renderWorklist() {
  const container = document.getElementById('focusWorklist');

  if (!container) return;

  const items = getWorkItems();

  if (!currentTicket && items.length) {
    selectFirstAvailableTicket();
  }

  renderSummary(items);

  if (!items.length) {
    container.innerHTML = `
      <div class="focus-worklist-empty">
        <strong>Nenhum chamado na lista.</strong>
        <span>
          Busque um chamado acima para adicioná-lo à lista de foco.
        </span>
      </div>
    `;
    return;
  }

  container.innerHTML = items
    .map((item) => {
      const session = getTicketSession(item.id);
      const status = getTicketStatus(item);
      const meta = getStatusMeta(status);
      const elapsedMs = getTicketElapsedMs(item);
      const isCurrent = currentTicket && Number(currentTicket.id) === Number(item.id);

      let primaryAction = 'start';
      let primaryText = 'Iniciar';

      if (status === 'running') {
        primaryAction = 'pause';
        primaryText = 'Pausar';
      } else if (status === 'paused') {
        primaryAction = 'resume';
        primaryText = 'Retomar';
      }

      const canFinish = status === 'running' || status === 'paused';
      const canRemove = !session;

      return `
        <article
          class="current-card focus-worklist-card ${meta.cardClass || ''} ${isCurrent ? 'focus-worklist-card--selected' : ''}"
          data-ticket-id="${item.id}"
        >
          <div class="focus-ticket-layout">
            <div
              class="focus-worklist-main"
              data-action="select"
              data-ticket-id="${item.id}"
            >
              <div class="focus-ticket-kicker">
                <span class="${meta.dotClass}"></span>
                <span>#${item.id}</span>
                <span>·</span>
                <span>${escapeHtml(item.team || 'Geral')}</span>
              </div>

              <div class="focus-ticket-title">
                ${escapeHtml(item.title || 'Sem título')}
              </div>

              <div class="focus-ticket-status">
                ${escapeHtml(meta.label)} · ${formatTime(elapsedMs)}
              </div>
            </div>

            <div class="focus-ticket-actions">
              <button
                class="pause-chip"
                type="button"
                data-action="${primaryAction}"
                data-ticket-id="${item.id}"
              >
                ${primaryText}
              </button>

              <button
                class="play-chip"
                type="button"
                title="Abrir widget"
                aria-label="Abrir widget"
                data-action="open"
                data-ticket-id="${item.id}"
              >
                ⚡
              </button>

              <button
                class="finish-chip finish-chip--icon"
                type="button"
                title="Concluir apontamento"
                data-action="finish"
                data-ticket-id="${item.id}"
                ${canFinish ? '' : 'disabled'}
              >
                ✓
              </button>

              <button
                class="finish-chip finish-chip--icon"
                type="button"
                title="Remover da lista"
                data-action="remove"
                data-ticket-id="${item.id}"
                ${canRemove ? '' : 'disabled'}
              >
                ×
              </button>
            </div>
          </div>
        </article>
      `;
    })
    .join('');
}

/* =========================
   Timer
========================= */

function startTimer() {
  stopTimer();

  timerInterval = setInterval(() => {
    renderWorklist();
  }, 1000);
}

function stopTimer() {
  if (timerInterval) {
    clearInterval(timerInterval);
    timerInterval = null;
  }
}

/* =========================
   Ações de ticket
========================= */

async function selectCurrentTicket(ticket) {
  const normalized = normalizeTicket(ticket);

  if (!normalized?.id) return;

  const result = window.FocusTrackFlows.addTicketToWorklist(normalized);

  if (!result?.ok) {
    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Não foi possível adicionar',
      message: result?.msg || 'Não foi possível adicionar chamado à lista.',
      confirmText: 'Entendi'
    });

    return;
  }

  setCurrentTicket(normalized);
  renderWorklist();
}

function getTicketById(ticketId) {
  const id = Number(ticketId);

  const item = getWorkItems().find((ticket) => Number(ticket.id) === id);
  const session = getTicketSession(id);

  if (session) {
    return {
      id: Number(session.ticketId),
      title: session.ticketTitle || item?.title || '',
      team: session.team || item?.team || 'Geral'
    };
  }

  if (item) {
    return {
      id: Number(item.id),
      title: item.title || '',
      team: item.team || 'Geral'
    };
  }

  return null;
}

function openWidget(ticket = currentTicket) {
  const normalized = normalizeTicket(ticket);

  if (!normalized?.id) return;

  if (window.focusTrack?.openWidget) {
    window.focusTrack.openWidget({
      id: Number(normalized.id),
      title: normalized.title || '',
      team: normalized.team || 'Geral'
    });
    return;
  }

  console.log('Abrir widget:', normalized);
}

async function createNewTicket(title) {
  if (!title) return;

  try {
    const created = await window.FocusTrackTickets.createTicket(title);

    if (!created) {
      await window.FocusTrackUI.alert({
        type: 'error',
        title: 'Chamado não criado',
        message: 'Não foi possível criar o chamado.',
        confirmText: 'Entendi'
      });

      return;
    }

    await selectCurrentTicket(created);

    const searchResults = document.getElementById('searchResults');
    const searchInput = document.getElementById('searchInput');
    const newTicketBtn = document.getElementById('newTicketBtn');

    if (searchResults) {
      searchResults.innerHTML = '';
      searchResults.classList.add('hidden');
    }

    if (searchInput) {
      searchInput.value = '';
    }

    if (newTicketBtn) {
      newTicketBtn.classList.add('new-ticket-btn--inactive');
      newTicketBtn.classList.remove('new-ticket-btn--active');
    }
  } catch (error) {
    console.error('Erro ao criar chamado:', error);

    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Erro ao criar chamado',
      message: error?.message || 'Erro ao criar chamado.',
      confirmText: 'Entendi'
    });
  }
}

async function startTicketById(ticketId) {
  const ticket = getTicketById(ticketId);
  if (!ticket) return;

  try {
    const result = await window.FocusTrackFlows.startTicket(ticket);

    if (!result?.ok) {
      await window.FocusTrackUI.alert({
        type: 'error',
        title: 'Erro ao iniciar',
        message: result?.msg || 'Erro ao iniciar apontamento.',
        confirmText: 'Entendi'
      });

      return;
    }

    setCurrentTicket(ticket);
    openWidget(ticket);
    renderWorklist();
  } catch (error) {
    console.error('Erro ao iniciar chamado:', error);

    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Erro ao iniciar chamado',
      message: error?.message || 'Erro ao iniciar chamado.',
      confirmText: 'Entendi'
    });
  }
}

async function pauseTicketById(ticketId) {
  const ticket = getTicketById(ticketId);
  if (!ticket) return;

  try {
    const result = await window.FocusTrackFlows.pauseTicket(ticket);

    if (!result?.ok) {
      await window.FocusTrackUI.alert({
        type: 'error',
        title: 'Erro ao pausar',
        message: result?.msg || 'Erro ao pausar apontamento.',
        confirmText: 'Entendi'
      });

      return;
    }

    setCurrentTicket(ticket);
    renderWorklist();
  } catch (error) {
    console.error('Erro ao pausar chamado:', error);

    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Erro ao pausar chamado',
      message: error?.message || 'Erro ao pausar chamado.',
      confirmText: 'Entendi'
    });
  }
}

async function resumeTicketById(ticketId) {
  const ticket = getTicketById(ticketId);
  if (!ticket) return;

  try {
    const result = await window.FocusTrackFlows.resumeTicket(ticket);

    if (!result?.ok) {
      await window.FocusTrackUI.alert({
        type: 'error',
        title: 'Erro ao retomar',
        message: result?.msg || 'Erro ao retomar apontamento.',
        confirmText: 'Entendi'
      });

      return;
    }

    setCurrentTicket(ticket);
    openWidget(ticket);
    renderWorklist();
  } catch (error) {
    console.error('Erro ao retomar chamado:', error);

    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Erro ao retomar chamado',
      message: error?.message || 'Erro ao retomar chamado.',
      confirmText: 'Entendi'
    });
  }
}

async function finishTicketById(ticketId) {
  const ticket = getTicketById(ticketId);
  if (!ticket) return;

  const confirmed = await window.FocusTrackUI.confirm({
    type: 'warning',
    title: 'Concluir apontamento?',
    message: 'Essa ação vai registrar o tempo e remover o chamado da lista de foco.',
    confirmText: 'Concluir',
    cancelText: 'Cancelar'
  });

  if (!confirmed) return;

  try {
    const result = await window.FocusTrackFlows.concludeTicket(ticket);

    if (!result?.ok) {
      await window.FocusTrackUI.alert({
        type: 'error',
        title: 'Erro ao concluir',
        message: result?.msg || 'Erro ao concluir apontamento.',
        confirmText: 'Entendi'
      });

      return;
    }

    if (currentTicket && Number(currentTicket.id) === Number(ticket.id)) {
      currentTicket = null;
      selectFirstAvailableTicket();
    }

    renderWorklist();
  } catch (error) {
    console.error('Erro ao concluir chamado:', error);

    await window.FocusTrackUI.alert({
      type: 'error',
      title: 'Erro ao concluir chamado',
      message: error?.message || 'Erro ao concluir chamado.',
      confirmText: 'Entendi'
    });
  }
}

function removeTicketById(ticketId) {
  const ticket = getTicketById(ticketId);
  if (!ticket) return;

  const session = getTicketSession(ticket.id);

  if (session) {
    window.FocusTrackUI.alert({
      type: 'warning',
      title: 'Chamado em andamento',
      message: 'Este chamado possui apontamento iniciado ou pausado. Conclua o apontamento antes de remover da lista.',
      confirmText: 'Entendi'
    });

    return;
  }

  window.FocusTrackFlows.removeTicketFromWorklist(ticket.id);

  if (currentTicket && Number(currentTicket.id) === Number(ticket.id)) {
    currentTicket = null;
    selectFirstAvailableTicket();
  }

  renderWorklist();
}

/* =========================
   Bootstrap
========================= */

window.addEventListener('DOMContentLoaded', async () => {
  const searchInput = document.getElementById('searchInput');
  const searchResults = document.getElementById('searchResults');
  const newTicketBtn = document.getElementById('newTicketBtn');
  const focusWorklist = document.getElementById('focusWorklist');
  const openTeamsBtn = document.getElementById('openTeamsBtn');

  const settingsEls = getSettingsElements();

  console.log('[FocusTrack] DOMContentLoaded popup.');
  console.log('[FocusTrack] openTeamsBtn:', openTeamsBtn);

  loadSettings();
  selectFirstAvailableTicket();
  renderWorklist();
  startTimer();

  openTeamsBtn?.addEventListener('click', async (event) => {
    event.preventDefault();
    event.stopPropagation();

    console.log('[FocusTrack] Botão T clicado.');

    await openTeamsDemands();
  });

  searchInput?.addEventListener('input', async (event) => {
    await renderResults(event.target.value);
  });

  newTicketBtn?.addEventListener('click', async () => {
    const query = searchInput?.value?.trim() || '';
    if (!query) return;
    await createNewTicket(query);
  });

  searchResults?.addEventListener('click', async (event) => {
    const ticketButton = event.target.closest('.result-ticket');

    if (ticketButton) {
      const ticket = {
        id: Number(ticketButton.dataset.ticketId),
        title: ticketButton.dataset.ticketTitle,
        team: ticketButton.dataset.ticketTeam || 'Geral'
      };

      await selectCurrentTicket(ticket);
      searchResults.classList.add('hidden');

      if (searchInput) {
        searchInput.value = '';
      }

      return;
    }

    const createButton = event.target.closest('.result-create');

    if (createButton) {
      const query = searchInput?.value?.trim() || '';
      if (!query) return;
      await createNewTicket(query);
    }
  });

  focusWorklist?.addEventListener('click', async (event) => {
    const actionEl = event.target.closest('[data-action]');
    if (!actionEl) return;

    const action = actionEl.dataset.action;
    const ticketId = Number(actionEl.dataset.ticketId);

    if (!ticketId) return;

    const ticket = getTicketById(ticketId);

    if (action === 'select') {
      if (ticket) {
        setCurrentTicket(ticket);
        renderWorklist();
      }

      return;
    }

    if (action === 'start') {
      await startTicketById(ticketId);
      return;
    }

    if (action === 'pause') {
      await pauseTicketById(ticketId);
      return;
    }

    if (action === 'resume') {
      await resumeTicketById(ticketId);
      return;
    }

    if (action === 'open') {
      if (ticket) {
        setCurrentTicket(ticket);
        openWidget(ticket);
        renderWorklist();
      }

      return;
    }

    if (action === 'finish') {
      await finishTicketById(ticketId);
      return;
    }

    if (action === 'remove') {
      removeTicketById(ticketId);
    }
  });

  settingsEls.settingsBtn?.addEventListener('click', () => {
    openSettingsModal();
  });

  settingsEls.closeBtn?.addEventListener('click', () => {
    closeSettingsModal();
  });

  settingsEls.cancelBtn?.addEventListener('click', () => {
    closeSettingsModal();
  });

  settingsEls.backdrop?.addEventListener('click', () => {
    closeSettingsModal();
  });

  settingsEls.testBtn?.addEventListener('click', async () => {
    await testSettings();
  });

  settingsEls.saveBtn?.addEventListener('click', () => {
    const ok = saveSettings();

    if (ok) {
      setTimeout(() => {
        closeSettingsModal();
      }, 350);
    }
  });

  window.addEventListener('focustrack:sessions-changed', () => {
    renderWorklist();
  });

  window.addEventListener('focustrack:worklist-changed', () => {
    renderWorklist();
  });

  window.addEventListener('focustrack:state-changed', () => {
    renderWorklist();
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      if (!settingsEls.modal?.classList.contains('hidden')) {
        closeSettingsModal();
        return;
      }

      window.focusTrack?.closeWindow?.();
    }
  });
});

window.addEventListener('beforeunload', () => {
  stopTimer();
});

console.log('POPUP JS MULTI-CRONÔMETROS VISUAL AJUSTADO CARREGADO');
