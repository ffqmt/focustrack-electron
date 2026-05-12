const { app, BrowserWindow, Tray, Menu, screen, ipcMain, shell } = require('electron');

const path = require('path');

app.disableHardwareAcceleration();
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-gpu-compositing');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('in-process-gpu');
app.commandLine.appendSwitch('no-sandbox');
app.commandLine.appendSwitch('use-angle', 'swiftshader');

const APP_BACKEND_CONFIG = {
  appName: 'FocusTrack',
  mode: 'mock',
  sharepoint: {
    siteUrl: 'https://contaudiassessoria.sharepoint.com/sites/TecnologiaContaudi',
    ticketsListId: '21d2acbc-82d5-419c-b4a5-3a1a962a69f0',
    timeEntriesListId: 'e6a41895-5b9f-438b-9873-ddeaa3364fb8',
    ticketsListName: 'Chamados',
    timeEntriesListName: 'Apontamentos'
  },
  flows: {
    createTicket: '',
    playTicket: 'PLAY_LancamentoTempo',
    pauseTicket: '',
    concludeTicket: ''
  }
};

let tray = null;
let popupWindow = null;
const widgetWindows = new Map();

function createPopupWindow() {
  popupWindow = new BrowserWindow({
    width: 440,
    height: 620,
    minWidth: 380,
    minHeight: 500,
    show: false,
    frame: false,
    transparent: true,
    resizable: true,
    movable: true,
    minimizable: true,
    maximizable: true,
    fullscreenable: false,
    skipTaskbar: false,
    alwaysOnTop: false,
    backgroundColor: '#00000000',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  const popupHtmlPath = path.join(__dirname, 'popup.html');
  console.log('Carregando popup de:', popupHtmlPath);

  popupWindow.loadFile(popupHtmlPath);

  popupWindow.on('closed', () => {
    popupWindow = null;
  });
}


function createWidgetWindow(ticket) {
  const safeTicket = {
    id: ticket?.id || 'sem-id',
    title: ticket?.title || 'Sem título',
    team: ticket?.team || 'Suporte Interno'
  };

  const key = String(safeTicket.id);

  const existingWidget = widgetWindows.get(key);

  if (existingWidget && !existingWidget.isDestroyed()) {
    existingWidget.show();
    existingWidget.focus();
    return existingWidget;
  }

  const widget = new BrowserWindow({
    width: 380,
    height: 540,

    minWidth: 320,
    minHeight: 280,

    show: true,
    frame: false,
    transparent: true,

    resizable: true,
    movable: true,
    minimizable: true,
    maximizable: true,
    fullscreenable: true,

    skipTaskbar: false,
    alwaysOnTop: true,
    backgroundColor: '#00000000',

    /**
     * No Windows, thickFrame ajuda janela sem frame a manter
     * comportamento melhor de resize/snap/maximize.
     */
    thickFrame: true,
    hasShadow: true,

    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });


  const display = screen.getPrimaryDisplay();
  const workArea = display.workArea;
  const offset = widgetWindows.size * 28;

  const initialWidth = 380;
  const initialHeight = 360;

  const x = workArea.x + workArea.width - initialWidth - 24 - offset;
  const y = workArea.y + workArea.height - initialHeight - 24 - offset;


  widget.setPosition(x, y, false);

  const fileUrl = `file://${path.join(__dirname, 'widget.html')}?ticketId=${encodeURIComponent(safeTicket.id)}&title=${encodeURIComponent(safeTicket.title)}&team=${encodeURIComponent(safeTicket.team)}`;

  widget.loadURL(fileUrl);

  widget.on('closed', () => {
    widgetWindows.delete(key);
  });

  widgetWindows.set(key, widget);

  return widget;
}

function getTrayBoundsSafe() {
  if (!tray) return null;

  try {
    return tray.getBounds();
  } catch {
    return null;
  }
}

function showPopupNearTray() {
  if (!popupWindow) return;

  const trayBounds = getTrayBoundsSafe();
  const display = screen.getPrimaryDisplay();
  const workArea = display.workArea;
  const windowBounds = popupWindow.getBounds();

  let x = workArea.x + workArea.width - windowBounds.width - 16;
  let y = workArea.y + workArea.height - windowBounds.height - 16;

  if (trayBounds) {
    x = Math.round(trayBounds.x + trayBounds.width / 2 - windowBounds.width / 2);
    y = Math.round(trayBounds.y - windowBounds.height - 10);

    if (y < workArea.y + 8) {
      y = Math.round(trayBounds.y + trayBounds.height + 10);
    }

    if (x + windowBounds.width > workArea.x + workArea.width) {
      x = workArea.x + workArea.width - windowBounds.width - 8;
    }

    if (x < workArea.x + 8) {
      x = workArea.x + 8;
    }
  }

  popupWindow.setPosition(x, y, false);
  popupWindow.show();
  popupWindow.focus();
}

function togglePopup() {
  if (!popupWindow) return;

  if (popupWindow.isVisible()) {
    popupWindow.hide();
  } else {
    showPopupNearTray();
  }
}

function toggleLatestWidget() {
  const widgets = Array.from(widgetWindows.values());
  const latest = widgets[widgets.length - 1];

  if (!latest) return;

  if (latest.isVisible()) {
    latest.hide();
  } else {
    latest.show();
    latest.focus();
  }
}

ipcMain.on('focustrack:close-window', (event) => {
  const senderWindow = BrowserWindow.fromWebContents(event.sender);
  if (senderWindow) {
    senderWindow.hide();
  }
});

ipcMain.on('focustrack:minimize-window', (event) => {
  const senderWindow = BrowserWindow.fromWebContents(event.sender);
  if (senderWindow) {
    senderWindow.minimize();
  }
});

ipcMain.on('focustrack:toggle-maximize-window', (event) => {
  const senderWindow = BrowserWindow.fromWebContents(event.sender);
  if (!senderWindow) return;

  if (senderWindow.isMaximized()) {
    senderWindow.unmaximize();
  } else {
    senderWindow.maximize();
  }
});
ipcMain.on('focustrack:resize-current-window', (event, size = {}) => {
  const senderWindow = BrowserWindow.fromWebContents(event.sender);

  if (!senderWindow || senderWindow.isDestroyed()) return;

  /**
   * Não redimensiona se estiver maximizada ou fullscreen.
   */
  if (senderWindow.isMaximized() || senderWindow.isFullScreen()) return;

  const width = Number(size.width || 380);
  const height = Number(size.height || 360);

  const safeWidth = Math.max(320, Math.min(width, 900));
  const safeHeight = Math.max(280, Math.min(height, 900));

  const bounds = senderWindow.getBounds();
  const display = screen.getDisplayMatching(bounds);
  const workArea = display.workArea;

  let x = bounds.x;
  let y = bounds.y;

  /**
   * Mantém a janela dentro da área útil da tela.
   */
  if (x + safeWidth > workArea.x + workArea.width) {
    x = workArea.x + workArea.width - safeWidth - 8;
  }

  if (y + safeHeight > workArea.y + workArea.height) {
    y = workArea.y + workArea.height - safeHeight - 8;
  }

  if (x < workArea.x) x = workArea.x + 8;
  if (y < workArea.y) y = workArea.y + 8;

  senderWindow.setBounds(
    {
      x,
      y,
      width: safeWidth,
      height: safeHeight
    },
    true
  );
});


ipcMain.on('focustrack:open-widget', (_, ticket) => {
  createWidgetWindow(ticket);
});

ipcMain.handle('focustrack:is-window-maximized', (event) => {
  const senderWindow = BrowserWindow.fromWebContents(event.sender);
  return senderWindow ? senderWindow.isMaximized() : false;
});

ipcMain.handle('focustrack:test-connection', async (_, payload) => {
  try {
    const response = await fetch('http://localhost:3001/api/sharepoint/test-connection', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload || {})
    });

    const data = await response.json();
    return data;
  } catch (error) {
    return {
      ok: false,
      message: 'Não foi possível conectar ao backend local.',
      checks: [
        {
          step: 'backend',
          ok: false,
          message: error?.message || 'Backend indisponível.'
        }
      ]
    };
  }
});

ipcMain.handle('focustrack:open-external', async (_, url) => {
  const rawUrl = String(url || '').trim();

  if (!rawUrl) {
    return {
      ok: false,
      message: 'URL não informada.'
    };
  }

  try {
    const parsed = new URL(rawUrl);

    const allowedProtocols = ['https:', 'http:', 'msteams:'];

    if (!allowedProtocols.includes(parsed.protocol)) {
      return {
        ok: false,
        message: 'Protocolo de URL não permitido.'
      };
    }

    await shell.openExternal(rawUrl);

    return {
      ok: true
    };
  } catch (error) {
    return {
      ok: false,
      message: error?.message || 'Não foi possível abrir o link.'
    };
  }
});

async function testConnection(input = {}) {
  const checks = [];

  function pushCheck(step, ok, message, extra = {}) {
    const item = { step, ok, message, ...extra };
    checks.push(item);
    return item;
  }

  function normalizeSiteUrl(url) {
    return String(url || '').trim().replace(/\/$/, '');
  }

  function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
  }

  function buildListUrlByTitle(siteUrl, listTitle, suffix = '') {
    const escapedTitle = String(listTitle || '').replace(/'/g, "''");
    return `${siteUrl}/_api/web/lists/GetByTitle('${escapedTitle}')${suffix}`;
  }

  function buildListUrlById(siteUrl, listId, suffix = '') {
    return `${siteUrl}/_api/web/lists(guid'${listId}')${suffix}`;
  }

  async function spFetchJson(url, options = {}) {
    const controller = new AbortController();
    const timeoutMs = options.timeoutMs || 12000;

    const timer = setTimeout(() => {
      controller.abort();
    }, timeoutMs);

    try {
      const response = await fetch(url, {
        method: options.method || 'GET',
        headers: {
          Accept: 'application/json;odata=verbose',
          ...(options.headers || {})
        },
        credentials: 'include',
        signal: controller.signal,
        body: options.body
      });

      const raw = await response.text();

      let data = null;
      try {
        data = raw ? JSON.parse(raw) : null;
      } catch {
        data = raw || null;
      }

      return { response, data, raw };
    } finally {
      clearTimeout(timer);
    }
  }

  try {
    const userEmail = String(
      input?.userEmail ||
      input?.currentUser?.email ||
      ''
    ).trim();

    const rendererSiteUrl =
      input?.sharePointSiteUrl ||
      input?.sharepointSiteUrl ||
      input?.sharePoint?.siteUrl ||
      '';

    const siteUrl = normalizeSiteUrl(
      rendererSiteUrl || APP_BACKEND_CONFIG.sharepoint?.siteUrl || ''
    );

    const ticketsListName =
      input?.ticketsListName ||
      input?.sharePoint?.lists?.tickets ||
      input?.lists?.tickets ||
      APP_BACKEND_CONFIG.sharepoint?.ticketsListName ||
      '';

    const timeEntriesListName =
      input?.timeEntriesListName ||
      input?.sharePoint?.lists?.timeEntries ||
      input?.lists?.timeEntries ||
      APP_BACKEND_CONFIG.sharepoint?.timeEntriesListName ||
      '';

    const ticketsListId =
      input?.ticketsListId ||
      input?.sharePoint?.lists?.ticketsId ||
      input?.lists?.ticketsId ||
      APP_BACKEND_CONFIG.sharepoint?.ticketsListId ||
      '';

    const timeEntriesListId =
      input?.timeEntriesListId ||
      input?.sharePoint?.lists?.timeEntriesId ||
      input?.lists?.timeEntriesId ||
      APP_BACKEND_CONFIG.sharepoint?.timeEntriesListId ||
      '';

    if (!userEmail) {
      pushCheck('config-user-email', false, 'E-mail do usuário não informado.');
      return {
        ok: false,
        message: 'Informe o e-mail do usuário.',
        checks
      };
    }

    if (!isValidEmail(userEmail)) {
      pushCheck('config-user-email', false, 'E-mail do usuário inválido.');
      return {
        ok: false,
        message: 'Informe um e-mail válido.',
        checks
      };
    }

    pushCheck('config-user-email', true, `E-mail informado: ${userEmail}`);

    if (!siteUrl) {
      pushCheck('config-site-url', false, 'URL do SharePoint não informada.');
      return {
        ok: false,
        message: 'URL do SharePoint não informada.',
        checks
      };
    }

    pushCheck('config-site-url', true, 'URL do SharePoint configurada.');
    pushCheck('config', true, 'Configuração básica válida.');

    const siteRequest = await spFetchJson(
      `${siteUrl}/_api/web?$select=Title,Url,Id`
    );

    if (!siteRequest.response.ok) {
      pushCheck(
        'site',
        false,
        `Falha ao acessar o site SharePoint (${siteRequest.response.status}).`,
        {
          httpStatus: siteRequest.response.status
        }
      );

      return {
        ok: false,
        message: getHttpMessage(siteRequest.response.status, 'site SharePoint'),
        checks,
        debug: {
          response: siteRequest.data
        }
      };
    }

    const siteInfo = siteRequest.data?.d || siteRequest.data || {};

    pushCheck(
      'site',
      true,
      `Site SharePoint acessível: ${siteInfo.Title || 'Sem título'}.`,
      {
        siteTitle: siteInfo.Title,
        siteUrl: siteInfo.Url,
        siteId: siteInfo.Id
      }
    );

    const ticketsListRequest = await fetchListInfo({
      siteUrl,
      listName: ticketsListName,
      listId: ticketsListId,
      step: 'tickets-list',
      label: 'lista de chamados',
      pushCheck,
      spFetchJson
    });

    if (!ticketsListRequest.ok) {
      return {
        ok: false,
        message: ticketsListRequest.message,
        checks,
        debug: ticketsListRequest.debug
      };
    }

    const ticketsReadRequest = await fetchListItems({
      siteUrl,
      listName: ticketsListName,
      listId: ticketsListId,
      step: 'tickets-read',
      label: ticketsListName || 'lista de chamados',
      pushCheck,
      spFetchJson
    });

    if (!ticketsReadRequest.ok) {
      return {
        ok: false,
        message: ticketsReadRequest.message,
        checks,
        debug: ticketsReadRequest.debug
      };
    }

    if (timeEntriesListName || timeEntriesListId) {
      const timeEntriesListRequest = await fetchListInfo({
        siteUrl,
        listName: timeEntriesListName,
        listId: timeEntriesListId,
        step: 'time-entries-list',
        label: 'lista de apontamentos',
        pushCheck,
        spFetchJson
      });

      if (!timeEntriesListRequest.ok) {
        return {
          ok: false,
          message: timeEntriesListRequest.message,
          checks,
          debug: timeEntriesListRequest.debug
        };
      }

      const timeEntriesReadRequest = await fetchListItems({
        siteUrl,
        listName: timeEntriesListName,
        listId: timeEntriesListId,
        step: 'time-entries-read',
        label: timeEntriesListName || 'lista de apontamentos',
        pushCheck,
        spFetchJson
      });

      if (!timeEntriesReadRequest.ok) {
        return {
          ok: false,
          message: timeEntriesReadRequest.message,
          checks,
          debug: timeEntriesReadRequest.debug
        };
      }
    }

    return {
      ok: true,
      message: 'Conexão validada com sucesso.',
      checks,
      details: {
        userEmail,
        siteTitle: siteInfo.Title || '',
        siteUrl: siteInfo.Url || siteUrl,
        ticketsListName: ticketsListName || null,
        ticketsListId: ticketsListId || null,
        timeEntriesListName: timeEntriesListName || null,
        timeEntriesListId: timeEntriesListId || null
      }
    };
  } catch (error) {
    pushCheck(
      'unexpected-error',
      false,
      error?.name === 'AbortError'
        ? 'Tempo excedido ao testar conexão com SharePoint.'
        : error?.message || 'Erro inesperado ao testar conexão.'
    );

    return {
      ok: false,
      message:
        error?.name === 'AbortError'
          ? 'Tempo excedido ao testar conexão com SharePoint.'
          : error?.message || 'Erro inesperado ao testar conexão.',
      checks
    };
  }

  function getListBaseUrl(siteUrlValue, listName, listId) {
    if (listId) {
      return buildListUrlById(siteUrlValue, listId);
    }

    if (listName) {
      return buildListUrlByTitle(siteUrlValue, listName);
    }

    return '';
  }

  async function fetchListInfo({
    siteUrl,
    listName,
    listId,
    step,
    label,
    pushCheck,
    spFetchJson
  }) {
    const baseUrl = getListBaseUrl(siteUrl, listName, listId);

    if (!baseUrl) {
      pushCheck(step, false, `Configuração da ${label} não informada.`);
      return {
        ok: false,
        message: `Configuração da ${label} não informada.`
      };
    }

    const request = await spFetchJson(
      `${baseUrl}?$select=Title,Id,ItemCount,Hidden`
    );

    if (!request.response.ok) {
      pushCheck(
        step,
        false,
        `Falha ao acessar ${label} (${request.response.status}).`,
        {
          httpStatus: request.response.status,
          listName,
          listId
        }
      );

      return {
        ok: false,
        message: getHttpMessage(request.response.status, label),
        debug: {
          response: request.data
        }
      };
    }

    const listInfo = request.data?.d || request.data || {};

    pushCheck(
      step,
      true,
      `${label.charAt(0).toUpperCase() + label.slice(1)} acessível.`,
      {
        listName: listInfo.Title || listName || null,
        listId: listInfo.Id || listId || null,
        itemCount: listInfo.ItemCount,
        hidden: listInfo.Hidden
      }
    );

    return {
      ok: true,
      data: listInfo
    };
  }

  async function fetchListItems({
    siteUrl,
    listName,
    listId,
    step,
    label,
    pushCheck,
    spFetchJson
  }) {
    const baseUrl = getListBaseUrl(siteUrl, listName, listId);

    if (!baseUrl) {
      pushCheck(step, false, `Configuração da ${label} não informada.`);
      return {
        ok: false,
        message: `Configuração da ${label} não informada.`
      };
    }

    const request = await spFetchJson(
      `${baseUrl}/items?$top=1&$select=Id,Title`
    );

    if (!request.response.ok) {
      pushCheck(
        step,
        false,
        `Não foi possível ler itens de ${label} (${request.response.status}).`,
        {
          httpStatus: request.response.status,
          listName,
          listId
        }
      );

      return {
        ok: false,
        message: getHttpMessage(request.response.status, label),
        debug: {
          response: request.data
        }
      };
    }

    const items =
      request.data?.d?.results ||
      request.data?.value ||
      [];

    pushCheck(
      step,
      true,
      `Leitura de ${label} validada.`,
      {
        sampleCount: Array.isArray(items) ? items.length : 0
      }
    );

    return {
      ok: true,
      data: items
    };
  }

  function getHttpMessage(status, targetLabel) {
    if (status === 401) {
      return `Não autenticado para acessar ${targetLabel}.`;
    }

    if (status === 403) {
      return `Sem permissão para acessar ${targetLabel}.`;
    }

    if (status === 404) {
      return `${targetLabel.charAt(0).toUpperCase() + targetLabel.slice(1)} não encontrado(a).`;
    }

    return `Falha ao acessar ${targetLabel} (${status}).`;
  }
}

app.whenReady().then(() => {
  createPopupWindow();

  tray = new Tray(path.join(__dirname, 'assets', 'tray.png'));
  tray.setToolTip('FocusTrack');

  const contextMenu = Menu.buildFromTemplate([
    { label: 'Abrir demandas', click: togglePopup },
    { label: 'Mostrar último widget', click: toggleLatestWidget },
    { type: 'separator' },
    { label: 'Sair', click: () => app.quit() }
  ]);

  tray.setContextMenu(contextMenu);
  tray.on('click', togglePopup);

  app.on('activate', () => {
    if (!popupWindow) {
      createPopupWindow();
    } else {
      showPopupNearTray();
    }
  });
});

app.on('window-all-closed', (event) => {
  event.preventDefault();
});
