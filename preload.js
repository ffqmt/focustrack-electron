const { contextBridge, ipcRenderer } = require('electron');

const API_BASE = 'http://localhost:3001';

async function request(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, {
    method: options.method || 'GET',
    headers: {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  let data = null;

  try {
    data = await response.json();
  } catch {
    data = {
      ok: false,
      message: `Resposta inválida do backend (${response.status})`
    };
  }

  if (!response.ok && data?.ok === undefined) {
    return {
      ok: false,
      message: data?.message || `Erro HTTP ${response.status}`
    };
  }

  return data;
}

contextBridge.exposeInMainWorld('focusTrack', {
  closeWindow: () => ipcRenderer.send('focustrack:close-window'),
  minimizeWindow: () => ipcRenderer.send('focustrack:minimize-window'),
  toggleMaximizeWindow: () => ipcRenderer.send('focustrack:toggle-maximize-window'),
  isWindowMaximized: () => ipcRenderer.invoke('focustrack:is-window-maximized'),
  openWidget: (ticket) => ipcRenderer.send('focustrack:open-widget', ticket),
  testConnection: (payload) => ipcRenderer.invoke('focustrack:test-connection', payload),
  resizeCurrentWindow: (size) => ipcRenderer.send('focustrack:resize-current-window', size),
  openExternal: (url) => ipcRenderer.invoke('focustrack:open-external', url),

  searchTickets: (query = '') =>
    request(`/api/tickets${query ? `?search=${encodeURIComponent(query)}` : ''}`),

  createTicket: (payload) =>
    request('/api/tickets', {
      method: 'POST',
      body: payload
    }),

  startTimeEntry: (payload) =>
    request('/api/time-entries/start', {
      method: 'POST',
      body: payload
    }),

  pauseTimeEntry: (id, payload = {}) =>
    request(`/api/time-entries/${id}/pause`, {
      method: 'POST',
      body: payload
    }),

  finishTimeEntry: (id, payload = {}) =>
    request(`/api/time-entries/${id}/finish`, {
      method: 'POST',
      body: payload
    })
});
