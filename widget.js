function formatTime(seconds) {
  const hrs = String(Math.floor(seconds / 3600)).padStart(2, '0');
  const mins = String(Math.floor((seconds % 3600) / 60)).padStart(2, '0');
  const secs = String(seconds % 60).padStart(2, '0');
  return `${hrs}:${mins}:${secs}`;
}

function formatMs(ms) {
  return formatTime(Math.floor(Number(ms || 0) / 1000));
}

function formatCommentTime(value) {
  const date = value ? new Date(value) : new Date();

  if (Number.isNaN(date.getTime())) {
    return '--:--';
  }

  return date.toLocaleTimeString('pt-BR', {
    hour: '2-digit',
    minute: '2-digit'
  });
}

function insertCommentTagIntoField(field, tag) {
  if (!field || !tag) return;

  const currentValue = field.value || '';

  const escapedTag = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const existingRe = new RegExp(`(^|\\s)${escapedTag}(?=\\s|$)`, 'i');

  if (existingRe.test(currentValue)) {
    field.focus();
    return;
  }

  const tagText = `${tag} `;

  const start = typeof field.selectionStart === 'number'
    ? field.selectionStart
    : currentValue.length;

  const end = typeof field.selectionEnd === 'number'
    ? field.selectionEnd
    : currentValue.length;

  const before = currentValue.slice(0, start);
  const after = currentValue.slice(end);

  const needsSpaceBefore = before.length > 0 && !/\s$/.test(before);
  const needsSpaceAfter = after.length > 0 && !/^\s/.test(after);

  const insertText = `${needsSpaceBefore ? ' ' : ''}${tagText}${needsSpaceAfter ? ' ' : ''}`;

  const nextValue = before + insertText + after;
  field.value = nextValue;

  const cursorPos = before.length + insertText.length;
  field.focus();

  try {
    field.setSelectionRange(cursorPos, cursorPos);
  } catch (_) {
    // campo não suporta seleção
  }

  field.dispatchEvent(new Event('input', { bubbles: true }));
  field.dispatchEvent(new Event('change', { bubbles: true }));
}

function findCommentFieldForTagButton(button) {
  const container =
    button.closest('.widget-comment-box') ||
    button.closest('.comment-box') ||
    button.closest('form') ||
    document;

  return (
    container.querySelector('#focusCommentInput') ||
    container.querySelector('textarea') ||
    container.querySelector('input[type="text"]') ||
    document.getElementById('focusCommentInput')
  );
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function getStatusUi(status) {
  switch (status) {
    case 'running':
      return {
        label: 'Em andamento',
        dotClass: 'status-dot--active',
        cardClass: 'widget-card--active'
      };
    case 'paused':
      return {
        label: 'Pausado',
        dotClass: 'status-dot--paused',
        cardClass: 'widget-card--paused'
      };
    case 'selected':
      return {
        label: 'Selecionado',
        dotClass: 'status-dot--waiting',
        cardClass: 'widget-card--waiting'
      };
    default:
      return {
        label: 'Selecionado',
        dotClass: 'status-dot--waiting',
        cardClass: 'widget-card--waiting'
      };
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  const card = document.querySelector('.widget-card');
  const timerElement = document.getElementById('widgetTimer');
  const togglePauseBtn = document.getElementById('togglePauseBtn');
  const finishBtn = document.getElementById('finishBtn');
  const statusDot = document.getElementById('widgetStatusDot');
  const statusText = document.getElementById('widgetStatusText');
  const ticketIdEl = document.getElementById('widgetTicketId');
  const titleEl = document.getElementById('widgetTitle');
  const teamEl = document.getElementById('widgetTeam');

  const focusObjectiveInput = document.getElementById('focusObjectiveInput');
  const focusCommentInput = document.getElementById('focusCommentInput');
  const addFocusCommentBtn = document.getElementById('addFocusCommentBtn');
  const focusCommentsList = document.getElementById('focusCommentsList');
  const widgetInlineStatus = document.getElementById('widgetInlineStatus');

  const focusCommentLimitInfo = document.getElementById('focusCommentLimitInfo');
  const focusCommentCounter = document.getElementById('focusCommentCounter');
  const focusCommentLimitHint = document.getElementById('focusCommentLimitHint');

  const toggleNotesPanelBtn = document.getElementById('toggleNotesPanelBtn');
  const notesPanelContent = document.getElementById('notesPanelContent');
  const notesToggleIcon = document.getElementById('notesToggleIcon');
  const notesCountBadge = document.getElementById('notesCountBadge');

  const params = new URLSearchParams(window.location.search);

  const widgetTicketId = Number(params.get('ticketId'));
  const widgetTicketTitle = params.get('title') || '';
  const widgetTicketTeam = params.get('team') || 'Geral';

  let ticket = null;
  let notesExpanded = false;
  let renderTimer = null;

  function getSession() {
    if (!widgetTicketId) return null;
    return window.FocusTrackFlows?.getSession?.(widgetTicketId) || null;
  }

  function getTicketFromUrl() {
    if (!widgetTicketId) return null;

    return {
      id: widgetTicketId,
      title: widgetTicketTitle || '',
      team: widgetTicketTeam || 'Geral',
      uiStatus: 'selected'
    };
  }

  function setNotesExpanded(expanded) {
  notesExpanded = Boolean(expanded);

  const widgetCard = document.querySelector('.widget-card');

  if (widgetCard) {
    widgetCard.classList.toggle('widget-card--notes-expanded', notesExpanded);
    widgetCard.classList.toggle('widget-card--notes-collapsed', !notesExpanded);
  }

  if (notesPanelContent) {
    notesPanelContent.classList.toggle('collapsed', !notesExpanded);
  }

  if (notesToggleIcon) {
    notesToggleIcon.textContent = notesExpanded ? '▴' : '▾';
  }

  if (toggleNotesPanelBtn) {
    toggleNotesPanelBtn.classList.toggle('expanded', notesExpanded);
  }

  /**
   * Redimensiona a janela junto com o estado das anotações.
   */
  window.setTimeout(() => {
    if (!window.focusTrack?.resizeCurrentWindow) return;

    if (notesExpanded) {
      window.focusTrack.resizeCurrentWindow({
        width: 320,
        height: 560
      });
    } else {
      window.focusTrack.resizeCurrentWindow({
        width: 320,
        height: 430
      });
    }
  }, 100);
}


  function setInlineStatus(type, message) {
    if (!widgetInlineStatus) return;

    if (!message) {
      widgetInlineStatus.textContent = '';
      widgetInlineStatus.className = 'widget-inline-status hidden';
      return;
    }

    widgetInlineStatus.textContent = message;
    widgetInlineStatus.className = `widget-inline-status widget-inline-status--${type}`;

    window.clearTimeout(setInlineStatus._timer);

    setInlineStatus._timer = window.setTimeout(() => {
      widgetInlineStatus.textContent = '';
      widgetInlineStatus.className = 'widget-inline-status hidden';
    }, 3500);
  }


  function getConfiguredFocusCommentMaxChars() {
    const config =
      window.FocusTrackConfigStore?.load?.() ||
      window.FocusTrackConfig ||
      {};

    const value = Number(
      config.focusCommentMaxChars ||
      config.centralCommentMaxChars ||
      config.commentMaxChars ||
      3000
    );

    if (!Number.isFinite(value) || value <= 500) {
      return 3000;
    }

    return Math.min(value, 10000);
  }

  function countChars(value) {
    return Array.from(String(value || '')).length;
  }

  function getNormalizedFocusComments(comments) {
    if (window.FocusTrackFlows?.normalizeFocusComments) {
      return window.FocusTrackFlows.normalizeFocusComments(comments || []);
    }

    if (!Array.isArray(comments)) return [];

    return comments
      .map((comment) => {
        if (typeof comment === 'string') {
          return {
            at: new Date().toISOString(),
            text: comment.trim()
          };
        }

        return {
          at: comment?.at || new Date().toISOString(),
          text: String(comment?.text || '').trim()
        };
      })
      .filter((comment) => comment.text);
  }

  function estimateFinalFocusCommentChars(session, comments) {

    const normalizedComments = getNormalizedFocusComments(comments || []);

    const lines = [];

    /**
     * Esta estimativa tenta imitar o texto final montado no server.js.
     * Não precisa ser 100% idêntica, mas precisa ser conservadora.
     */
    lines.push('[FocusTrack] Apontamento registrado');
    lines.push('');
    lines.push('Entrada: 00/00/0000 00:00');
    lines.push('Saída: 00/00/0000 00:00');
    lines.push('Duração: 000 min');
    lines.push('Responsável: Usuário FocusTrack');



    if (normalizedComments.length > 0) {
      lines.push('');
      lines.push('Comentários do foco:');

      normalizedComments.forEach((comment) => {
        lines.push(`- 00:00 — ${comment.text}`);
      });
    }

    return countChars(lines.join('\n'));
  }

  function buildCommentsWithDraft(session) {
    const currentComments = getNormalizedFocusComments(session?.focusComments || []);
    const draftText = String(focusCommentInput?.value || '').trim();

    if (!draftText) {
      return currentComments;
    }

    return [
      ...currentComments,
      {
        at: new Date().toISOString(),
        text: draftText
      }
    ];
  }

  function validateDraftFocusCommentLimit() {
    const session = getSession();
    const maxChars = getConfiguredFocusCommentMaxChars();

    if (!session) {
      return {
        ok: true,
        totalChars: 0,
        maxChars,
        remainingChars: maxChars,
        exceededBy: 0
      };
    }

    const commentsWithDraft = buildCommentsWithDraft(session);

    /**
     * Se o flows.js já tiver a validação centralizada, usa ela.
     * Isso mantém o widget alinhado com o restante do app.
     */
    if (typeof window.FocusTrackFlows?.validateFocusCommentLimit === 'function') {
      const validationSession = {
        ...session,
        focusObjective: String(
          focusObjectiveInput?.value ||
          session.focusObjective ||
          ''
        ).trim()
      };

      return window.FocusTrackFlows.validateFocusCommentLimit(
        validationSession,
        commentsWithDraft
      );
    }

    /**
     * Fallback local caso o flows.js ainda não tenha validateFocusCommentLimit.
     */
    const totalChars = estimateFinalFocusCommentChars(session, commentsWithDraft);
    const remainingChars = maxChars - totalChars;

    if (totalChars > maxChars) {
      return {
        ok: false,
        code: 'FOCUS_COMMENT_TOO_LARGE',
        totalChars,
        maxChars,
        remainingChars,
        exceededBy: totalChars - maxChars,
        msg:
          `O texto final ficaria com ${totalChars} caracteres, ` +
          `mas o limite configurado é ${maxChars}. ` +
          `Reduza pelo menos ${totalChars - maxChars} caracteres.`
      };
    }

    return {
      ok: true,
      totalChars,
      maxChars,
      remainingChars,
      exceededBy: 0
    };
  }

  function updateFocusCommentCounter() {
    if (!focusCommentLimitInfo || !focusCommentCounter || !focusCommentLimitHint) {
      return;
    }

    const session = getSession();
    const draftText = String(focusCommentInput?.value || '').trim();

    if (!session) {
      focusCommentLimitInfo.classList.add('hidden');
      focusCommentCounter.textContent = '';
      focusCommentLimitHint.textContent = '';

      if (addFocusCommentBtn) {
        addFocusCommentBtn.disabled = true;
      }

      return;
    }

    const validation = validateDraftFocusCommentLimit();
    const hasDraftText = Boolean(draftText);

    focusCommentLimitInfo.classList.remove('hidden');

    focusCommentCounter.textContent =
      `${validation.totalChars}/${validation.maxChars} caracteres`;

    focusCommentLimitInfo.classList.toggle('is-danger', !validation.ok);
    focusCommentLimitInfo.classList.toggle(
      'is-warning',
      validation.ok && validation.remainingChars <= 300
    );

    if (!validation.ok) {
      focusCommentLimitHint.textContent =
        `Excedeu ${validation.exceededBy} caracteres.`;

      if (addFocusCommentBtn) {
        addFocusCommentBtn.disabled = true;
        addFocusCommentBtn.title = validation.msg || 'Comentário acima do limite.';
      }

      return;
    }

    if (validation.remainingChars <= 300) {
      focusCommentLimitHint.textContent =
        `Restam ${validation.remainingChars} caracteres.`;
    } else {
      focusCommentLimitHint.textContent = '';
    }

    if (addFocusCommentBtn) {
      addFocusCommentBtn.disabled = !hasDraftText;
      addFocusCommentBtn.title = '';
    }
  }



  function syncFromSession() {
    const session = getSession();

    if (session) {
      ticket = {
        id: Number(session.ticketId),
        title: session.ticketTitle || widgetTicketTitle || '',
        team: session.team || widgetTicketTeam || 'Geral',
        uiStatus: session.status || 'selected'
      };

      return;
    }

    ticket = getTicketFromUrl();
  }

  function updateNotesBadge(session) {
    const commentsCount = Array.isArray(session?.focusComments)
      ? session.focusComments.length
      : 0;

    if (!notesCountBadge) return;

    if (commentsCount > 0) {
      notesCountBadge.textContent = String(commentsCount);
      notesCountBadge.classList.remove('hidden');
    } else {
      notesCountBadge.textContent = '0';
      notesCountBadge.classList.add('hidden');
    }
  }

  function renderComments(session) {
    if (!focusCommentsList) return;

    const comments = Array.isArray(session?.focusComments)
      ? session.focusComments
      : [];

    if (!session) {
      focusCommentsList.innerHTML = `
        <div class="widget-comments-empty">
          Inicie o foco para adicionar comentários.
        </div>
      `;
      return;
    }

    if (!comments.length) {
      focusCommentsList.innerHTML = `
        <div class="widget-comments-empty">
          Nenhum comentário neste intervalo.
        </div>
      `;
      return;
    }

    focusCommentsList.innerHTML = comments
      .map((comment, index) => {
        const time = formatCommentTime(comment?.at);
        const text = escapeHtml(comment?.text || '');

        return `
          <div class="widget-comment-item">
            <div class="widget-comment-content">
              <span class="widget-comment-time">${time}</span>
              <span class="widget-comment-text">${text}</span>
            </div>

            <button
              class="widget-comment-remove"
              type="button"
              data-comment-index="${index}"
              title="Remover comentário"
            >
              ×
            </button>
          </div>
        `;
      })
      .join('');
  }

  function renderEmpty() {
    if (statusText) statusText.textContent = 'Sem foco';
    if (statusDot) statusDot.className = 'status-dot status-dot--waiting';
    if (card) card.className = 'widget-card widget-card--waiting';
    if (ticketIdEl) ticketIdEl.textContent = '--';
    if (titleEl) titleEl.textContent = 'Nenhum chamado vinculado';
    if (teamEl) teamEl.textContent = '';
    if (timerElement) timerElement.textContent = '00:00:00';

    if (togglePauseBtn) {
      togglePauseBtn.textContent = 'Iniciar';
      togglePauseBtn.disabled = true;
    }

    if (finishBtn) {
      finishBtn.disabled = true;
    }

    if (toggleNotesPanelBtn) {
      toggleNotesPanelBtn.disabled = true;
    }

    if (focusObjectiveInput) {
      focusObjectiveInput.value = '';
      focusObjectiveInput.disabled = true;
    }

    if (focusCommentInput) {
      focusCommentInput.value = '';
      focusCommentInput.disabled = true;
    }

    if (addFocusCommentBtn) {
      addFocusCommentBtn.disabled = true;
    }

    renderComments(null);
  }

  function render() {
    syncFromSession();

    const session = getSession();

    updateNotesBadge(session);

    if (!ticket) {
      renderEmpty();
      return;
    }

    const uiStatus = session?.status || ticket.uiStatus || 'selected';
    const statusUi = getStatusUi(uiStatus);

    if (ticketIdEl) ticketIdEl.textContent = ticket.id;
    if (titleEl) titleEl.textContent = ticket.title || '';
    if (teamEl) teamEl.textContent = ticket.team || 'Geral';

    if (timerElement) {
      timerElement.textContent = session
        ? formatMs(window.FocusTrackFlows.getElapsedMs(session))
        : '00:00:00';
    }

    if (card) {
      card.className = `widget-card ${statusUi.cardClass}`;
    }

    if (statusDot) {
      statusDot.className = `status-dot ${statusUi.dotClass}`;
    }

    if (statusText) {
      statusText.textContent = statusUi.label;
    }

    if (togglePauseBtn) {
      if (uiStatus === 'running') {
        togglePauseBtn.textContent = 'Pausar';
      } else if (uiStatus === 'paused') {
        togglePauseBtn.textContent = 'Retomar';
      } else {
        togglePauseBtn.textContent = 'Iniciar';
      }

      togglePauseBtn.disabled = false;
    }

    if (finishBtn) {
      finishBtn.disabled = !(uiStatus === 'running' || uiStatus === 'paused');
    }

    if (toggleNotesPanelBtn) {
      toggleNotesPanelBtn.disabled = false;
    }

    if (focusObjectiveInput) {
      focusObjectiveInput.disabled = false;

      if (document.activeElement !== focusObjectiveInput) {
        focusObjectiveInput.value = session?.focusObjective || focusObjectiveInput.value || '';
      }
    }

    if (focusCommentInput) {
      focusCommentInput.disabled = !session;
    }

    if (addFocusCommentBtn) {
      addFocusCommentBtn.disabled = !session;
    }

    renderComments(session);
    updateFocusCommentCounter();
  }

  function saveObjectiveFromInput() {
    /**
     * Campo objetivo foi descontinuado.
     * Mantemos a função para não quebrar os pontos onde ela é chamada.
     */
    return;
  }


  function addCommentFromInput() {
    const session = getSession();

    if (!session) {
      setInlineStatus('error', 'Inicie o foco antes de adicionar comentários.');
      return;
    }

    const value = String(focusCommentInput?.value || '').trim();

    if (!value) {
      setInlineStatus('error', 'Digite um comentário antes de adicionar.');
      updateFocusCommentCounter();
      return;
    }

    const validation = validateDraftFocusCommentLimit();

    if (!validation.ok) {
      setInlineStatus(
        'error',
        validation.msg ||
          `Comentário acima do limite. Excedeu ${validation.exceededBy || 0} caracteres.`
      );

      updateFocusCommentCounter();
      return;
    }

    const result = window.FocusTrackFlows.addFocusComment(value, widgetTicketId);

    if (!result?.ok) {
      setInlineStatus('error', result?.msg || 'Não foi possível adicionar comentário.');
      updateFocusCommentCounter();
      return;
    }

    focusCommentInput.value = '';
    setInlineStatus('success', 'Comentário adicionado.');
    setNotesExpanded(true);
    render();
    updateFocusCommentCounter();
  }


  async function startCurrentTicket() {
  if (!ticket?.id) return;

  const result = await window.FocusTrackFlows.startTicket(
    {
      id: Number(ticket.id),
      title: ticket.title || '',
      team: ticket.team || 'Geral'
    },
    {
      focusObjective: ''
    }
  );

  return result;
}


  toggleNotesPanelBtn?.addEventListener('click', () => {
    setNotesExpanded(!notesExpanded);
  });

  focusObjectiveInput?.addEventListener('input', () => {
    saveObjectiveFromInput();
    updateFocusCommentCounter();
  });


  focusObjectiveInput?.addEventListener('blur', () => {
    saveObjectiveFromInput();
  });

  addFocusCommentBtn?.addEventListener('click', () => {
    addCommentFromInput();
  });
    focusCommentInput?.addEventListener('input', () => {
    updateFocusCommentCounter();
  });


  focusCommentInput?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();

      const validation = validateDraftFocusCommentLimit();

      if (!validation.ok) {
        setInlineStatus(
          'error',
          validation.msg ||
            `Comentário acima do limite. Excedeu ${validation.exceededBy || 0} caracteres.`
        );

        updateFocusCommentCounter();
        return;
      }

      addCommentFromInput();
    }
  });


  focusCommentsList?.addEventListener('click', (event) => {
    const button = event.target.closest('[data-comment-index]');

    if (!button) return;

    const index = Number(button.dataset.commentIndex);

    const result = window.FocusTrackFlows.removeFocusComment(
      index,
      widgetTicketId
    );

    if (!result?.ok) {
      setInlineStatus('error', result?.msg || 'Não foi possível remover comentário.');
      return;
    }

    setInlineStatus('success', 'Comentário removido.');
    render();
  });

  togglePauseBtn?.addEventListener('click', async () => {
    syncFromSession();

    if (!ticket?.id) return;

    saveObjectiveFromInput();

    const session = getSession();

    const current = {
      id: Number(ticket.id),
      title: ticket.title || '',
      team: ticket.team || 'Geral'
    };

    try {
      let result;

      if (!session) {
        result = await startCurrentTicket();
      } else if (session.status === 'paused') {
        result = await window.FocusTrackFlows.resumeTicket(current);
      } else if (session.status === 'running') {
        result = await window.FocusTrackFlows.pauseTicket(current);
      } else {
        result = await startCurrentTicket();
      }

      if (!result?.ok) {
        await window.FocusTrackUI.alert({
          type: 'error',
          title: 'Erro ao alterar status',
          message: result?.msg || 'Erro ao alterar status.',
          confirmText: 'Entendi'
        });

        render();
        return;
      }

      if (session?.status === 'running') {
        setInlineStatus('success', 'Foco pausado. Comentários enviados.');
      } else if (session?.status === 'paused') {
        setInlineStatus('success', 'Foco retomado.');
      } else {
        setInlineStatus('success', 'Foco iniciado.');
      }

      render();
    } catch (error) {
      console.error('Erro no widget:', error);
      alert(error?.message || 'Erro ao alterar status do apontamento.');
    }
  });

  finishBtn?.addEventListener('click', async () => {
    const session = getSession();

    if (!session) return;

    saveObjectiveFromInput();

    const current = {
      id: Number(session.ticketId),
      title: session.ticketTitle,
      team: session.team || 'Geral'
    };

    try {
      const result = await window.FocusTrackFlows.concludeTicket(current);

      if (!result?.ok) {
        await window.FocusTrackUI.alert({
          type: 'error',
          title: 'Não foi possível concluir',
          message: result?.msg || 'Erro ao concluir.',
          confirmText: 'Entendi'
        });

        render();
        return;
      }

      await window.FocusTrackUI.alert({
        type: 'success',
        title: 'Apontamento concluído',
        message: 'O tempo foi registrado com sucesso.',
        confirmText: 'Fechar'
      });

      render();
      window.focusTrack?.closeWindow?.();

    } catch (error) {
      console.error('Erro no widget ao concluir:', error);
      await window.FocusTrackUI.alert({
        type: 'error',
        title: 'Erro ao concluir',
        message: error?.message || 'Erro ao concluir apontamento.',
        confirmText: 'Entendi'
      });

    }
  });

  window.addEventListener('focustrack:sessions-changed', () => {
    render();
  });

  window.addEventListener('focustrack:state-changed', () => {
    render();
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      window.focusTrack?.closeWindow?.();
    }
  });

  if (!window.__focusTrackCommentTagShortcutsBound) {
    window.__focusTrackCommentTagShortcutsBound = true;

    document.addEventListener('click', (event) => {
      const button = event.target.closest('[data-comment-tag]');
      if (!button) return;

      event.preventDefault();

      const tag = button.getAttribute('data-comment-tag');
      const field = findCommentFieldForTagButton(button);

      insertCommentTagIntoField(field, tag);
    });
  }

  setNotesExpanded(false);

  renderTimer = window.setInterval(() => {
    render();
  }, 1000);

  window.addEventListener('beforeunload', () => {
    if (renderTimer) {
      window.clearInterval(renderTimer);
      renderTimer = null;
    }
  });

  render();
});
