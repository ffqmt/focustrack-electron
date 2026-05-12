(() => {
  const DEFAULT_CONFIG = {
    appName: 'FocusTrack',
    mode: 'real',
    apiBaseUrl: 'http://localhost:3001',

    currentUser: {
      name: '',
      email: ''
    },

    sharePoint: {
      siteUrl: 'https://contaudiassessoria.sharepoint.com/sites/TecnologiaContaudi',
      lists: {
        tickets: 'Chamados',
        timeEntries: 'LancamentosTempo',
        ticketsId: '21d2acbc-82d5-419c-b4a5-3a1a962a69f0',
        timeEntriesId: 'e6a41895-5b9f-438b-9873-ddeaa3364fb8'
      }
    },

    links: {
      teamsDemandsUrl: 'https://teams.microsoft.com/l/entity/26bc2873-6023-480c-a11b-76b66605ce8c/_djb2_msteams_prefix_2810599455?context=%7B%22channelId%22%3A%2219%3A8m7ZGscFI8DiTnbcgxhzOc9_1aqEYMtlZtw6odWaShw1%40thread.tacv2%22%7D&tenantId=d2f6807c-d369-4371-93d4-16a3561f25f2'
    }

  };

  const STORAGE_KEY = 'focustrack:user-config';

  function clone(data) {
    return JSON.parse(JSON.stringify(data));
  }

  function mergeConfig(parsed = {}) {
    const base = clone(DEFAULT_CONFIG);

    return {
      ...base,
      ...parsed,

      currentUser: {
        ...base.currentUser,
        ...(parsed.currentUser || {})
      },

      sharePoint: {
        ...base.sharePoint,
        ...(parsed.sharePoint || {}),
        lists: {
          ...base.sharePoint.lists,
          ...(parsed.sharePoint?.lists || {})
        }
      },

      links: {
        ...base.links,
        ...(parsed.links || {})
      }
    };
  }

  function loadConfig() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);

      if (!raw) {
        return clone(DEFAULT_CONFIG);
      }

      const parsed = JSON.parse(raw);
      return mergeConfig(parsed);
    } catch (error) {
      console.error('Erro ao carregar config:', error);
      return clone(DEFAULT_CONFIG);
    }
  }

  function saveConfig(nextConfig) {
    const merged = mergeConfig(nextConfig);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
    window.FocusTrackConfig = merged;
    return merged;
  }

  function updateCurrentUser(nextUser) {
    const current = loadConfig();

    const nextConfig = {
      ...current,
      currentUser: {
        ...current.currentUser,
        ...(nextUser || {})
      }
    };

    return saveConfig(nextConfig);
  }

  function updateSharePoint(nextSharePoint) {
    const current = loadConfig();

    const nextConfig = {
      ...current,
      sharePoint: {
        ...current.sharePoint,
        ...(nextSharePoint || {}),
        lists: {
          ...current.sharePoint?.lists,
          ...(nextSharePoint?.lists || {})
        }
      }
    };

    return saveConfig(nextConfig);
  }

  function updateApiBaseUrl(apiBaseUrl) {
    const current = loadConfig();

    const nextConfig = {
      ...current,
      apiBaseUrl: String(apiBaseUrl || '').trim() || DEFAULT_CONFIG.apiBaseUrl
    };

    return saveConfig(nextConfig);
  }

  function updateLinks(nextLinks) {
    const current = loadConfig();

    const nextConfig = {
      ...current,
      links: {
        ...current.links,
        ...(nextLinks || {})
      }
    };

    return saveConfig(nextConfig);
  }

  function resetConfig() {
    localStorage.removeItem(STORAGE_KEY);
    window.FocusTrackConfig = clone(DEFAULT_CONFIG);
    return window.FocusTrackConfig;
  }

  window.FocusTrackDefaultConfig = clone(DEFAULT_CONFIG);
  window.FocusTrackConfig = loadConfig();

  window.FocusTrackConfigStore = {
    load: loadConfig,
    save: saveConfig,
    reset: resetConfig,
    updateCurrentUser,
    updateSharePoint,
    updateApiBaseUrl,
    updateLinks
  };
})();
