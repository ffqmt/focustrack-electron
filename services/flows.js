(function () {
  const LEGACY_STORAGE_KEY = 'focustrack:active-session';
  const SESSIONS_KEY = 'focustrack:sessions';
  const WORKLIST_KEY = 'focustrack:worklist';

  /**
   * Locks em localStorage para proteger contra:
   * - clique duplo
   * - popup e widget concluindo ao mesmo tempo
   * - eventos duplicados em janelas diferentes do Electron
   */
  const FINISH_LOCK_PREFIX = 'focustrack:finish-lock:';
  const PAUSE_LOCK_PREFIX = 'focustrack:pause-lock:';
  const LOCK_TTL_MS = 30000;

  /**
   * Limite real desejado para o texto final do comentário.
   * Esse total inclui:
   * - cabeçalho FocusTrack
   * - entrada/saída
   * - duração
   * - responsável
   * - título "Comentários do foco:"
   * - comentários digitados
   *
   * O campo "Objetivo" foi descontinuado e não entra mais na conta.
   */
  const DEFAULT_FOCUS_COMMENT_MAX_CHARS = 2000;

  function getConfig() {
    return window.FocusTrackConfigStore?.load?.() || window.FocusTrackConfig || {};
  }

  function getApiBaseUrl() {
    return getConfig().apiBaseUrl || 'http://localhost:3001';
  }

  function getCurrentUser() {
    const config = getConfig();
    const currentUser = config?.currentUser || {};

    return {
      name: String(currentUser.name || '').trim(),
      email: String(currentUser.email || '').trim()
    };
  }

  function emitStateChanged() {
    try {
      window.dispatchEvent(new CustomEvent('focustrack:sessions-changed'));
      window.dispatchEvent(new CustomEvent('focustrack:worklist-changed'));
      window.dispatchEvent(new CustomEvent('focustrack:state-changed'));
    } catch {
      // silencioso
    }
  }

  function safeJsonParse(value, fallback) {
    try {
      if (!value) return fallback;
      return JSON.parse(value);
    } catch {
      return fallback;
    }
  }

  function acquireOperationLock(prefix, id) {
    const numericId = Number(id);

    if (!numericId) return false;

    const key = `${prefix}${numericId}`;
    const now = Date.now();
    const current = Number(localStorage.getItem(key) || 0);

    if (current && now - current < LOCK_TTL_MS) {
      return false;
    }

    localStorage.setItem(key, String(now));
    return true;
  }

  function releaseOperationLock(prefix, id) {
    const numericId = Number(id);

    if (!numericId) return;

    const key = `${prefix}${numericId}`;
    localStorage.removeItem(key);
  }

  function acquireFinishLock(ticketId) {
    return acquireOperationLock(FINISH_LOCK_PREFIX, ticketId);
  }

  function releaseFinishLock(ticketId) {
    releaseOperationLock(FINISH_LOCK_PREFIX, ticketId);
  }

  function acquirePauseLock(ticketId) {
    return acquireOperationLock(PAUSE_LOCK_PREFIX, ticketId);
  }

  function releasePauseLock(ticketId) {
    releaseOperationLock(PAUSE_LOCK_PREFIX, ticketId);
  }

  function loadSessions() {
    migrateLegacySessionIfNeeded();
    return safeJsonParse(localStorage.getItem(SESSIONS_KEY), {});
  }

  function saveSessions(sessions) {
    localStorage.setItem(SESSIONS_KEY, JSON.stringify(sessions || {}));
    emitStateChanged();
    return sessions || {};
  }

  function getSession(ticketId) {
    if (!ticketId) return null;

    const sessions = loadSessions();
    return sessions[String(Number(ticketId))] || null;
  }

  function setSession(ticketId, session) {
    const id = Number(ticketId);

    if (!id) return null;

    const sessions = loadSessions();

    sessions[String(id)] = {
      ...session,
      ticketId: id,

      /**
       * Campo objetivo descontinuado.
       * Mantemos vazio para limpar sessões antigas sem quebrar compatibilidade.
       */
      focusObjective: ''
    };

    saveSessions(sessions);

    return sessions[String(id)];
  }

  function removeSession(ticketId) {
    const id = Number(ticketId);

    if (!id) return;

    const sessions = loadSessions();
    delete sessions[String(id)];
    saveSessions(sessions);
  }

  function getAllSessions() {
    return loadSessions();
  }

  function clearAllSessions() {
    localStorage.removeItem(SESSIONS_KEY);
    localStorage.removeItem(LEGACY_STORAGE_KEY);
    emitStateChanged();
  }

  function migrateLegacySessionIfNeeded() {
    const alreadyHasSessions = localStorage.getItem(SESSIONS_KEY);

    if (alreadyHasSessions) return;

    const legacy = safeJsonParse(localStorage.getItem(LEGACY_STORAGE_KEY), null);

    if (!legacy?.ticketId) return;

    /**
     * Ao migrar sessão antiga, já descartamos focusObjective.
     */
    const normalizedLegacy = {
      ...legacy,
      focusObjective: ''
    };

    const sessions = {
      [String(Number(legacy.ticketId))]: normalizedLegacy
    };

    localStorage.setItem(SESSIONS_KEY, JSON.stringify(sessions));
    localStorage.removeItem(LEGACY_STORAGE_KEY);
  }

  function loadWorklist() {
    return safeJsonParse(localStorage.getItem(WORKLIST_KEY), []);
  }

  function saveWorklist(items) {
    localStorage.setItem(
      WORKLIST_KEY,
      JSON.stringify(Array.isArray(items) ? items : [])
    );

    emitStateChanged();
    return Array.isArray(items) ? items : [];
  }

  function getWorklist() {
    return loadWorklist();
  }

  function addTicketToWorklist(ticket) {
    const normalized = normalizeTicket(ticket);

    if (!normalized?.id) {
      return {
        ok: false,
        msg: 'Chamado inválido.'
      };
    }

    const items = loadWorklist();
    const exists = items.some((item) => Number(item.id) === Number(normalized.id));

    if (exists) {
      return {
        ok: true,
        msg: 'Chamado já está na lista.',
        worklist: items
      };
    }

    const nextItems = [
      ...items,
      {
        id: normalized.id,
        title: normalized.title || '',
        team: normalized.team || 'Geral',
        addedAt: new Date().toISOString()
      }
    ];

    saveWorklist(nextItems);

    return {
      ok: true,
      msg: 'Chamado adicionado à lista.',
      worklist: nextItems
    };
  }

  function removeTicketFromWorklist(ticketId) {
    const id = Number(ticketId);

    if (!id) {
      return {
        ok: false,
        msg: 'Chamado inválido.'
      };
    }

    const nextItems = loadWorklist().filter((item) => Number(item.id) !== id);
    saveWorklist(nextItems);

    return {
      ok: true,
      worklist: nextItems
    };
  }

  function clearWorklist() {
    localStorage.removeItem(WORKLIST_KEY);
    emitStateChanged();

    return {
      ok: true,
      worklist: []
    };
  }

  function getElapsedMs(session) {
    if (!session) return 0;

    const elapsedMs = Number(session.elapsedMs || 0);

    if (session.status === 'running' && session.startedAt) {
      return elapsedMs + Math.max(0, Date.now() - Number(session.startedAt));
    }

    return elapsedMs;
  }

  async function postJson(url, body) {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body || {})
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      const error = new Error(data?.message || `Erro HTTP ${response.status}`);
      error.status = response.status;
      error.data = data;
      error.code = data?.code || null;
      error.details = data?.details || null;
      error.url = url;
      throw error;
    }

    return data;
  }

  function isTimeEntryNotFoundError(error) {
    const message = String(error?.message || '').toLowerCase();
    const dataText = JSON.stringify(error?.data || {}).toLowerCase();

    return (
      error?.status === 404 ||
      message.includes('404') ||
      message.includes('not found') ||
      message.includes('item not found') ||
      message.includes('specified list item') ||
      message.includes('não encontrado') ||
      message.includes('nao encontrado') ||
      dataText.includes('not found') ||
      dataText.includes('item not found') ||
      dataText.includes('specified list item') ||
      dataText.includes('não encontrado') ||
      dataText.includes('nao encontrado')
    );
  }

  function isFocusCommentPostFailedError(error) {
    const message = String(error?.message || '').toLowerCase();
    const dataText = JSON.stringify(error?.data || {}).toLowerCase();

    return (
      error?.code === 'FOCUS_COMMENT_POST_FAILED' ||
      error?.data?.code === 'FOCUS_COMMENT_POST_FAILED' ||
      message.includes('comentário nativo') ||
      message.includes('comentario nativo') ||
      message.includes('focus_comment_post_failed') ||
      dataText.includes('focus_comment_post_failed')
    );
  }

  function isFocusCommentTooLargeError(error) {
    const message = String(error?.message || '').toLowerCase();
    const dataText = JSON.stringify(error?.data || {}).toLowerCase();

    return (
      error?.status === 413 ||
      error?.code === 'FOCUS_COMMENT_TOO_LARGE' ||
      error?.data?.code === 'FOCUS_COMMENT_TOO_LARGE' ||
      message.includes('focus_comment_too_large') ||
      message.includes('comentário do focustrack possui') ||
      message.includes('comentario do focustrack possui') ||
      message.includes('limite configurado') ||
      dataText.includes('focus_comment_too_large') ||
      dataText.includes('limite configurado')
    );
  }

  function isAlreadyFinishingError(error) {
    return (
      error?.status === 409 ||
      error?.code === 'FINISH_ALREADY_RUNNING' ||
      error?.data?.code === 'FINISH_ALREADY_RUNNING'
    );
  }

  function getCommentDeliveryOk(data) {
    return Boolean(
      data?.centralCommentPrepared?.ok ||
      data?.nativeComment?.ok
    );
  }

  function getCommentDeliveryMode(data) {
    if (data?.centralCommentPrepared?.ok) {
      return data?.centralCommentPrepared?.mode || 'central-field';
    }

    if (data?.nativeComment?.ok) {
      return 'native-comment';
    }

    return null;
  }

  function normalizeTicket(ticket) {
    if (!ticket) return null;

    return {
      id: Number(ticket.id || ticket.ticketId),
      title: ticket.title || ticket.ticketTitle || '',
      team: ticket.team || 'Geral'
    };
  }

  function normalizeFocusComments(comments) {
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

  function getFocusCommentMaxChars() {
    const config = getConfig();

    const value = Number(
      config.focusCommentMaxChars ||
      config.centralCommentMaxChars ||
      config.commentMaxChars ||
      DEFAULT_FOCUS_COMMENT_MAX_CHARS
    );

    if (!Number.isFinite(value) || value <= 500) {
      return DEFAULT_FOCUS_COMMENT_MAX_CHARS;
    }

    return Math.min(value, 10000);
  }

  function countChars(text) {
    return Array.from(String(text || '')).length;
  }

  /**
   * Estima o tamanho do comentário final completo.
   *
   * IMPORTANTE:
   * O campo "Objetivo" foi removido do cálculo e do envio.
   */
  function estimateFocusCommentEnvelopeChars(session, comments) {
    const normalizedComments = normalizeFocusComments(comments || []);

    const lines = [];

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

  function validateFocusCommentLimit(session, nextComments) {
    const maxChars = getFocusCommentMaxChars();
    const totalChars = estimateFocusCommentEnvelopeChars(session, nextComments);
    const remainingChars = maxChars - totalChars;

    if (totalChars > maxChars) {
      return {
        ok: false,
        code: 'FOCUS_COMMENT_TOO_LARGE',
        maxChars,
        totalChars,
        exceededBy: totalChars - maxChars,
        remainingChars,
        msg:
          `O texto final ficaria com ${totalChars} caracteres, ` +
          `mas o limite configurado é ${maxChars}. ` +
          `Reduza o comentário em pelo menos ${totalChars - maxChars} caracteres.`
      };
    }

    return {
      ok: true,
      maxChars,
      totalChars,
      remainingChars
    };
  }

  function hasPendingFocusComments(session) {
    return normalizeFocusComments(session?.focusComments || []).length > 0;
  }

  function getPendingFocusComments(session) {
    return normalizeFocusComments(session?.focusComments || []);
  }

  function getFocusCommentBatchesSent(session) {
    return Number(session?.focusCommentBatchesSent || 0);
  }

  function getFocusCommentsSentCount(session) {
    return Number(session?.focusCommentsSentCount || 0);
  }

  /**
   * Monta payload para pause/finish.
   *
   * Campo "Objetivo" foi descontinuado:
   * - não envia focusObjective
   * - não envia objective
   * - não força comentário só por objetivo
   */
  function buildCommentPayload(session, notes, options = {}) {
    const focusComments = normalizeFocusComments(session?.focusComments || []);

    const payload = {
      notes,
      focusComments,

      /**
       * Mantém compatibilidade com backends que ainda leem commentText.
       * O texto é composto apenas pelos comentários reais.
       */
      commentText: focusComments.map((comment) => comment.text).join('\n')
    };

    if (options.skipTicketComment === true) {
      payload.skipTicketComment = true;
    }

    if (options.forceTicketComment === true) {
      payload.forceTicketComment = true;
    }

    if (options.postCommentToTicket === true) {
      payload.postCommentToTicket = true;
    }

    if (options.postCommentToTicket === false) {
      payload.postCommentToTicket = false;
    }

    return payload;
  }

  function buildPausePayload(session) {
    const pendingComments = getPendingFocusComments(session);
    const hasComments = pendingComments.length > 0;

    return {
      endAt: new Date().toISOString(),
      ...buildCommentPayload(session, 'Pausado pela aplicação', {
        skipTicketComment: !hasComments,
        postCommentToTicket: hasComments
      })
    };
  }

  function buildFinishPayload(session) {
    const pendingComments = getPendingFocusComments(session);
    const hasComments = pendingComments.length > 0;

    if (hasComments) {
      return {
        endAt: new Date().toISOString(),
        ...buildCommentPayload(session, 'Concluído pela aplicação', {
          skipTicketComment: false,
          postCommentToTicket: true
        })
      };
    }

    return {
      endAt: new Date().toISOString(),
      notes: 'Concluído pela aplicação',
      skipTicketComment: true,
      postCommentToTicket: false
    };
  }

  function resolveTargetSession(ticketId) {
    if (ticketId) {
      return getSession(ticketId);
    }

    return getActiveSession();
  }

  /**
   * Compatibilidade com versões antigas do widget.
   *
   * O campo objetivo foi descontinuado.
   * Se algum código antigo chamar essa função, ela apenas limpa o objetivo
   * da sessão para evitar que sessões antigas continuem carregando texto.
   */
  function setFocusObjective(_objective, ticketId) {
    const session = resolveTargetSession(ticketId);

    if (!session) {
      return { ok: false, msg: 'Nenhuma sessão encontrada.' };
    }

    const nextSession = {
      ...session,
      focusObjective: ''
    };

    setSession(nextSession.ticketId, nextSession);

    return {
      ok: true,
      session: nextSession,
      deprecated: true,
      msg: 'Campo objetivo foi descontinuado. Use comentários do foco.'
    };
  }

  function addFocusComment(text, ticketId) {
    const session = resolveTargetSession(ticketId);
    const cleanText = String(text || '').trim();

    if (!session) {
      return { ok: false, msg: 'Nenhuma sessão encontrada.' };
    }

    if (!cleanText) {
      return { ok: false, msg: 'Digite um comentário.' };
    }

    const comment = {
      at: new Date().toISOString(),
      text: cleanText
    };

    const nextComments = [
      ...normalizeFocusComments(session.focusComments || []),
      comment
    ];

    const validation = validateFocusCommentLimit(session, nextComments);

    if (!validation.ok) {
      return {
        ok: false,
        code: validation.code,
        msg: validation.msg,
        maxChars: validation.maxChars,
        totalChars: validation.totalChars,
        exceededBy: validation.exceededBy,
        remainingChars: validation.remainingChars
      };
    }

    const nextSession = {
      ...session,
      focusObjective: '',
      focusComments: nextComments
    };

    setSession(nextSession.ticketId, nextSession);

    return {
      ok: true,
      session: nextSession,
      comment,
      limit: validation
    };
  }

  function removeFocusComment(index, ticketId) {
    const session = resolveTargetSession(ticketId);

    if (!session) {
      return { ok: false, msg: 'Nenhuma sessão encontrada.' };
    }

    const comments = normalizeFocusComments(session.focusComments || []);
    const targetIndex = Number(index);

    if (
      Number.isNaN(targetIndex) ||
      targetIndex < 0 ||
      targetIndex >= comments.length
    ) {
      return { ok: false, msg: 'Comentário inválido.' };
    }

    comments.splice(targetIndex, 1);

    const nextSession = {
      ...session,
      focusObjective: '',
      focusComments: comments
    };

    setSession(nextSession.ticketId, nextSession);

    return {
      ok: true,
      session: nextSession
    };
  }

  function clearFocusComments(ticketId) {
    const session = resolveTargetSession(ticketId);

    if (!session) {
      return { ok: false, msg: 'Nenhuma sessão encontrada.' };
    }

    const nextSession = {
      ...session,
      focusObjective: '',
      focusComments: []
    };

    setSession(nextSession.ticketId, nextSession);

    return {
      ok: true,
      session: nextSession
    };
  }

  async function startTicket(ticket, options = {}) {
    const normalized = normalizeTicket(ticket);

    if (!normalized?.id) {
      return { ok: false, msg: 'Selecione um ticket válido.' };
    }

    addTicketToWorklist(normalized);

    const currentSession = getSession(normalized.id);

    if (currentSession?.status === 'running') {
      return {
        ok: true,
        msg: 'Este ticket já está em andamento.',
        session: currentSession
      };
    }

    if (currentSession?.status === 'paused') {
      return resumeTicket(normalized);
    }

    const currentUser = getCurrentUser();

    const payload = {
      ticketId: normalized.id,
      title: `Apontamento demanda ${normalized.id}`,
      startAt: new Date().toISOString(),
      notes: '',
      responsibleName: currentUser.name || currentUser.email || '',
      responsibleEmail: currentUser.email || ''
    };

    console.log('START payload', payload);

    const apiBaseUrl = getApiBaseUrl();
    const data = await postJson(`${apiBaseUrl}/api/time-entries/start`, payload);
    const timeEntry = data?.timeEntry;

    const nextSession = {
      ticketId: normalized.id,
      ticketTitle: normalized.title,
      team: normalized.team,
      status: 'running',
      startedAt: Date.now(),
      elapsedMs: Number(currentSession?.elapsedMs || 0),
      timeEntryId: timeEntry?.id || null,

      /**
       * Objetivo descontinuado.
       */
      focusObjective: '',

      focusComments: normalizeFocusComments(
        options?.focusComments ||
        currentSession?.focusComments ||
        []
      ),

      focusCommentBatchesSent: Number(currentSession?.focusCommentBatchesSent || 0),
      focusCommentsSentCount: Number(currentSession?.focusCommentsSentCount || 0),
      lastFocusCommentSentAt: currentSession?.lastFocusCommentSentAt || null
    };

    setSession(normalized.id, nextSession);

    return {
      ok: true,
      msg: 'Apontamento iniciado com sucesso.',
      session: nextSession,
      timeEntry
    };
  }

  async function pauseTicket(ticket) {
    const normalized = normalizeTicket(ticket);

    if (!normalized?.id) {
      return { ok: false, msg: 'Selecione um ticket válido.' };
    }

    if (!acquirePauseLock(normalized.id)) {
      return {
        ok: false,
        code: 'PAUSE_ALREADY_RUNNING',
        msg: 'Este apontamento já está sendo pausado. Aguarde alguns segundos.'
      };
    }

    try {
      const session = getSession(normalized.id);

      if (!session || !session.timeEntryId) {
        return {
          ok: false,
          msg: 'Nenhum apontamento ativo encontrado para este ticket.'
        };
      }

      if (session.status !== 'running') {
        return { ok: false, msg: 'Este ticket não está em andamento.' };
      }

      const elapsedMs = getElapsedMs(session);
      const pendingComments = getPendingFocusComments(session);
      const pendingCommentsCount = pendingComments.length;

      /**
       * Validação preventiva antes de enviar ao backend.
       */
      const validation = validateFocusCommentLimit(session, pendingComments);

      if (pendingCommentsCount > 0 && !validation.ok) {
        return {
          ok: false,
          code: validation.code,
          preserveComments: true,
          msg: validation.msg,
          details: validation
        };
      }

      const payload = buildPausePayload(session);

      console.log('PAUSE payload', payload, 'timeEntryId', session.timeEntryId);

      const apiBaseUrl = getApiBaseUrl();

      let data;

      try {
        data = await postJson(
          `${apiBaseUrl}/api/time-entries/${session.timeEntryId}/pause`,
          payload
        );
      } catch (error) {
        if (isTimeEntryNotFoundError(error)) {
          removeSession(normalized.id);

          return {
            ok: false,
            code: 'TIME_ENTRY_NOT_FOUND',
            sessionCleared: true,
            msg: 'Este apontamento não existe mais no SharePoint. O foco local foi encerrado.'
          };
        }

        if (isFocusCommentTooLargeError(error)) {
          return {
            ok: false,
            code: 'FOCUS_COMMENT_TOO_LARGE',
            preserveComments: true,
            msg:
              error?.message ||
              'O comentário ficou maior que o limite permitido. Os comentários foram preservados localmente.',
            details: error?.details || error?.data?.details || null,
            error
          };
        }

        if (isFocusCommentPostFailedError(error)) {
          return {
            ok: false,
            code: 'FOCUS_COMMENT_POST_FAILED',
            preserveComments: true,
            msg:
              error?.message ||
              'O tempo foi pausado, mas os comentários não foram enviados. Eles foram preservados localmente.',
            error
          };
        }

        throw error;
      }

      const commentDeliveryOk = getCommentDeliveryOk(data);
      const commentDeliveryMode = getCommentDeliveryMode(data);

      const shouldClearComments =
        pendingCommentsCount === 0 || commentDeliveryOk;

      const sentNowCount = shouldClearComments ? pendingCommentsCount : 0;

      const nextSession = {
        ...session,
        status: 'paused',
        elapsedMs,
        startedAt: null,
        focusObjective: '',

        focusComments: shouldClearComments
          ? []
          : normalizeFocusComments(session.focusComments || []),

        focusCommentBatchesSent:
          getFocusCommentBatchesSent(session) + (sentNowCount > 0 ? 1 : 0),

        focusCommentsSentCount:
          getFocusCommentsSentCount(session) + sentNowCount,

        lastFocusCommentSentAt:
          sentNowCount > 0
            ? new Date().toISOString()
            : session.lastFocusCommentSentAt || null
      };

      setSession(normalized.id, nextSession);

      return {
        ok: true,
        msg: commentDeliveryOk
          ? 'Apontamento pausado e comentários preparados para envio com sucesso.'
          : 'Apontamento pausado com sucesso.',
        session: nextSession,
        timeEntry: data?.timeEntry,
        centralCommentPrepared: data?.centralCommentPrepared || null,
        nativeComment: data?.nativeComment || null,
        commentDeliveryOk,
        commentDeliveryMode,
        warnings: data?.warnings || []
      };
    } finally {
      releasePauseLock(normalized.id);
    }
  }

  async function resumeTicket(ticket) {
    const normalized = normalizeTicket(ticket);

    if (!normalized?.id) {
      return { ok: false, msg: 'Selecione um ticket válido.' };
    }

    addTicketToWorklist(normalized);

    const session = getSession(normalized.id);

    if (!session) {
      return startTicket(normalized);
    }

    if (session.status === 'running') {
      return {
        ok: true,
        msg: 'Este ticket já está em andamento.',
        session
      };
    }

    const currentUser = getCurrentUser();

    const payload = {
      ticketId: normalized.id,
      title: `Apontamento demanda ${normalized.id}`,
      startAt: new Date().toISOString(),
      notes: 'Retomado pela aplicação',
      responsibleName: currentUser.name || currentUser.email || '',
      responsibleEmail: currentUser.email || ''
    };

    console.log('RESUME payload', payload);
    console.log('CURRENT USER', currentUser);

    const apiBaseUrl = getApiBaseUrl();
    const data = await postJson(`${apiBaseUrl}/api/time-entries/start`, payload);
    const timeEntry = data?.timeEntry;

    const nextSession = {
      ...session,
      ticketId: normalized.id,
      ticketTitle: normalized.title || session.ticketTitle,
      team: normalized.team || session.team || 'Geral',
      status: 'running',
      startedAt: Date.now(),
      elapsedMs: Number(session.elapsedMs || 0),
      timeEntryId: timeEntry?.id || null,

      focusObjective: '',
      focusComments: normalizeFocusComments(session.focusComments || []),

      focusCommentBatchesSent: Number(session.focusCommentBatchesSent || 0),
      focusCommentsSentCount: Number(session.focusCommentsSentCount || 0),
      lastFocusCommentSentAt: session.lastFocusCommentSentAt || null
    };

    setSession(normalized.id, nextSession);

    return {
      ok: true,
      msg: 'Apontamento retomado com sucesso.',
      session: nextSession,
      timeEntry
    };
  }

  async function concludeTicket(ticket) {
    const normalized = normalizeTicket(ticket);

    if (!normalized?.id) {
      return { ok: false, msg: 'Selecione um ticket válido.' };
    }

    if (!acquireFinishLock(normalized.id)) {
      return {
        ok: false,
        code: 'FINISH_ALREADY_RUNNING',
        msg: 'Este apontamento já está sendo concluído. Aguarde alguns segundos.'
      };
    }

    try {
      const session = getSession(normalized.id);

      if (!session || !session.timeEntryId) {
        return {
          ok: false,
          msg: 'Nenhum apontamento ativo encontrado para este ticket.'
        };
      }

      if (session.status !== 'running' && session.status !== 'paused') {
        return { ok: false, msg: 'Não há sessão válida para concluir.' };
      }

      const elapsedMs = getElapsedMs(session);
      const pendingComments = getPendingFocusComments(session);
      const pendingCommentsCount = pendingComments.length;
      const hadPendingComments = pendingCommentsCount > 0;

      /**
       * Validação preventiva antes de enviar ao backend.
       */
      const validation = validateFocusCommentLimit(session, pendingComments);

      if (hadPendingComments && !validation.ok) {
        return {
          ok: false,
          code: validation.code,
          preserveComments: true,
          msg: validation.msg,
          details: validation
        };
      }

      const payload = buildFinishPayload(session);

      console.log('FINISH payload', payload, 'timeEntryId', session.timeEntryId);

      const apiBaseUrl = getApiBaseUrl();

      let data;

      try {
        data = await postJson(
          `${apiBaseUrl}/api/time-entries/${session.timeEntryId}/finish`,
          payload
        );
      } catch (error) {
        if (isTimeEntryNotFoundError(error)) {
          removeSession(normalized.id);
          removeTicketFromWorklist(normalized.id);

          return {
            ok: false,
            code: 'TIME_ENTRY_NOT_FOUND',
            sessionCleared: true,
            msg: 'Este apontamento não existe mais no SharePoint. O foco local foi encerrado.'
          };
        }

        if (isAlreadyFinishingError(error)) {
          return {
            ok: false,
            code: 'FINISH_ALREADY_RUNNING',
            msg: 'Este apontamento já está sendo concluído. Aguarde alguns segundos.',
            error
          };
        }

        if (isFocusCommentTooLargeError(error)) {
          return {
            ok: false,
            code: 'FOCUS_COMMENT_TOO_LARGE',
            preserveComments: true,
            msg:
              error?.message ||
              'O comentário ficou maior que o limite permitido. Os comentários foram preservados localmente.',
            details: error?.details || error?.data?.details || null,
            error
          };
        }

        if (isFocusCommentPostFailedError(error)) {
          return {
            ok: false,
            code: 'FOCUS_COMMENT_POST_FAILED',
            preserveComments: true,
            msg:
              error?.message ||
              'O apontamento foi processado, mas os comentários não foram enviados. Eles foram preservados localmente.',
            error
          };
        }

        throw error;
      }

      const commentDeliveryOk = getCommentDeliveryOk(data);
      const commentDeliveryMode = getCommentDeliveryMode(data);

      if (hadPendingComments && !commentDeliveryOk) {
        const preservedSession = {
          ...session,
          status: 'paused',
          elapsedMs,
          startedAt: null,
          focusObjective: '',
          focusComments: normalizeFocusComments(session.focusComments || [])
        };

        setSession(normalized.id, preservedSession);

        return {
          ok: false,
          code: 'FOCUS_COMMENT_NOT_CONFIRMED',
          preserveComments: true,
          msg:
            'Não foi possível confirmar o preparo/envio dos comentários. A sessão foi preservada para evitar perda de dados.',
          session: preservedSession,
          timeEntry: data?.timeEntry,
          centralCommentPrepared: data?.centralCommentPrepared || null,
          nativeComment: data?.nativeComment || null,
          commentDeliveryOk,
          commentDeliveryMode,
          warnings: data?.warnings || []
        };
      }

      const finishedSession = {
        ...session,
        status: 'finished',
        elapsedMs,
        startedAt: null,

        focusObjective: '',
        focusComments: [],

        focusCommentBatchesSent:
          getFocusCommentBatchesSent(session) + (pendingCommentsCount > 0 ? 1 : 0),

        focusCommentsSentCount:
          getFocusCommentsSentCount(session) + pendingCommentsCount,

        lastFocusCommentSentAt:
          pendingCommentsCount > 0
            ? new Date().toISOString()
            : session.lastFocusCommentSentAt || null
      };

      removeSession(normalized.id);
      removeTicketFromWorklist(normalized.id);

      return {
        ok: true,
        msg: commentDeliveryOk
          ? 'Apontamento concluído e comentários preparados para envio com sucesso.'
          : 'Apontamento concluído com sucesso.',
        session: finishedSession,
        timeEntry: data?.timeEntry,
        centralCommentPrepared: data?.centralCommentPrepared || null,
        nativeComment: data?.nativeComment || null,
        commentDeliveryOk,
        commentDeliveryMode,
        warnings: data?.warnings || []
      };
    } finally {
      releaseFinishLock(normalized.id);
    }
  }

  function getActiveSession() {
    const sessions = loadSessions();
    const list = Object.values(sessions || {});

    return (
      list.find((session) => session.status === 'running') ||
      list.find((session) => session.status === 'paused') ||
      null
    );
  }

  function clearSession(ticketId) {
    if (ticketId) {
      removeSession(ticketId);
      return;
    }

    localStorage.removeItem(LEGACY_STORAGE_KEY);
  }

  window.FocusTrackFlows = {
    startTicket,
    pauseTicket,
    resumeTicket,
    concludeTicket,
    getFocusCommentMaxChars,
    countChars,
    estimateFocusCommentEnvelopeChars,
    validateFocusCommentLimit,

    getSession,
    setSession,
    removeSession,
    getAllSessions,
    getActiveSession,
    getElapsedMs,
    clearSession,
    clearAllSessions,

    getWorklist,
    addTicketToWorklist,
    removeTicketFromWorklist,
    clearWorklist,

    /**
     * Mantido só por compatibilidade.
     * O objetivo foi descontinuado.
     */
    setFocusObjective,

    addFocusComment,
    removeFocusComment,
    clearFocusComments,

    /**
     * Expostos para debug/diagnóstico, se precisar testar no console.
     */
    normalizeFocusComments,
    hasPendingFocusComments,

    /**
     * Debug do fluxo de comentário.
     */
    getCommentDeliveryOk,
    getCommentDeliveryMode
  };
})();
