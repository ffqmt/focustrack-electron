const { ConfidentialClientApplication } = require('@azure/msal-node');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const cca = new ConfidentialClientApplication(msalConfig);

async function acquireGraphToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: [GRAPH_SCOPE]
  });

  if (!result?.accessToken) {
    throw new Error('Não foi possível obter token de acesso do Microsoft Graph.');
  }

  return result.accessToken;
}

function decodeJwt(token) {
  const parts = String(token || '').split('.');
  if (parts.length < 2) {
    return null;
  }

  const payload = parts[1]
    .replace(/-/g, '+')
    .replace(/_/g, '/');

  const padded = payload + '='.repeat((4 - (payload.length % 4)) % 4);
  const json = Buffer.from(padded, 'base64').toString('utf8');
  return JSON.parse(json);
}

async function graphRequest(path, options = {}) {
  const token = await acquireGraphToken();

  const response = await fetch(`${GRAPH_BASE}${path}`, {
    method: options.method || 'GET',
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
      ...(options.headers || {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  const raw = await response.text();

  let data = null;
  try {
    data = raw ? JSON.parse(raw) : null;
  } catch {
    data = raw || null;
  }

  return {
    ok: response.ok,
    status: response.status,
    data,
    headers: Object.fromEntries(response.headers.entries())
  };
}

function normalizeSiteUrl(siteUrl) {
  return String(siteUrl || '').trim().replace(/\/$/, '');
}

function extractHostname(siteUrl) {
  return new URL(normalizeSiteUrl(siteUrl)).hostname;
}

function extractSitePath(siteUrl) {
  return new URL(normalizeSiteUrl(siteUrl)).pathname;
}

async function getSite(siteUrl) {
  const hostname = extractHostname(siteUrl);
  const path = extractSitePath(siteUrl);

  const result = await graphRequest(`/sites/${hostname}:${path}`);

  if (!result.ok) {
    throw new Error(
      `Falha ao localizar site no Graph (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  return result.data;
}

async function getListById(siteId, listId) {
  const result = await graphRequest(`/sites/${siteId}/lists/${listId}`);

  if (!result.ok) {
    throw new Error(
      `Falha ao localizar lista ${listId} (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  return result.data;
}

async function getListItems(siteId, listId, top = 10) {
  const result = await graphRequest(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&top=${Number(top) || 10}`
  );

  if (!result.ok) {
    throw new Error(
      `Falha ao obter itens da lista ${listId} (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  return result.data;
}

async function getListItemsByQuery(siteId, listId, queryParams = {}) {
  const { filter, top, expand, select } = queryParams;
  const params = new URLSearchParams();
  if (filter) params.append('$filter', filter);
  if (top) params.append('$top', top);
  if (expand) params.append('$expand', expand);
  if (select) params.append('$select', select);

  const queryString = params.toString() ? `?${params.toString()}` : '';
  const basePath = `/sites/${siteId}/lists/${listId}/items${queryString}`;

  let allItems = [];
  let nextPath = basePath;

  while (nextPath) {
    const result = await graphRequest(nextPath);

    if (!result.ok) {
      throw new Error(
        `Falha ao obter itens da lista ${listId} (${result.status}): ${JSON.stringify(result.data)}`
      );
    }

    const data = result.data;
    if (data.value) {
      allItems = allItems.concat(data.value);
    }

    if (data['@odata.nextLink']) {
      nextPath = data['@odata.nextLink'].replace(GRAPH_BASE, '');
    } else {
      nextPath = null;
    }
  }

  return allItems;
}

async function createListItem(siteId, listId, fields) {
  const result = await graphRequest(`/sites/${siteId}/lists/${listId}/items`, {
    method: 'POST',
    body: {
      fields
    }
  });

  if (!result.ok) {
    throw new Error(
      `Falha ao criar item na lista ${listId} (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  return result.data;
}

async function updateListItem(siteId, listId, itemId, fields) {
  const result = await graphRequest(
    `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
    {
      method: 'PATCH',
      body: fields
    }
  );

  if (!result.ok) {
    throw new Error(
      `Falha ao atualizar item ${itemId} da lista ${listId} (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  return {
    ok: true,
    status: result.status,
    data: result.data
  };
}

async function validateListsAccess(siteId, ticketsId, timeEntriesId) {
  const output = {
    ok: true,
    checks: []
  };

  async function validateOne(name, listId) {
    if (!listId) {
      output.ok = false;
      output.checks.push({
        step: `${name}-config`,
        ok: false,
        message: `ID da lista ${name} não informado.`
      });
      return;
    }

    try {
      const list = await getListById(siteId, listId);
      output.checks.push({
        step: `${name}-exists`,
        ok: true,
        message: `Lista ${name} encontrada.`,
        list: {
          id: list.id,
          name: list.name,
          displayName: list.displayName,
          webUrl: list.webUrl
        }
      });

      const items = await getListItems(siteId, listId, 3);
      output.checks.push({
        step: `${name}-read`,
        ok: true,
        message: `Leitura da lista ${name} validada.`,
        countReturned: Array.isArray(items?.value) ? items.value.length : 0
      });
    } catch (error) {
      output.ok = false;
      output.checks.push({
        step: `${name}-access`,
        ok: false,
        message: error.message
      });
    }
  }

  await validateOne('tickets', ticketsId);
  await validateOne('timeEntries', timeEntriesId);

  return output;
}

async function getTokenDebugInfo() {
  const token = await acquireGraphToken();
  const claims = decodeJwt(token);

  return {
    ok: true,
    message: 'Token obtido com sucesso.',
    claims: claims
      ? {
          aud: claims.aud || null,
          iss: claims.iss || null,
          tid: claims.tid || null,
          appid: claims.appid || null,
          idtyp: claims.idtyp || null,
          roles: claims.roles || [],
          scp: claims.scp || null
        }
      : null
  };
}

async function testConnection(payload = {}) {
  const checks = [];

  function pushCheck(step, ok, message, extra = {}) {
    checks.push({ step, ok, message, ...extra });
  }

  try {
    const userEmail = String(payload?.userEmail || '').trim();
    const siteUrl = normalizeSiteUrl(
      payload?.sharePoint?.siteUrl || process.env.SHAREPOINT_SITE_URL || ''
    );
    const ticketsId = String(
      payload?.sharePoint?.lists?.ticketsId ||
      process.env.SHAREPOINT_TICKETS_LIST_ID ||
      ''
    ).trim();

    const timeEntriesId = String(
      payload?.sharePoint?.lists?.timeEntriesId ||
      process.env.SHAREPOINT_TIME_ENTRIES_LIST_ID ||
      ''
    ).trim();


    if (!userEmail) {
      pushCheck('config-user-email', false, 'E-mail do usuário não informado.');
      return { ok: false, message: 'Informe o e-mail do usuário.', checks };
    }

    pushCheck('config-user-email', true, `E-mail informado: ${userEmail}`);

    if (!siteUrl) {
      pushCheck('config-site-url', false, 'URL do SharePoint não informada.');
      return { ok: false, message: 'URL do SharePoint não informada.', checks };
    }

    pushCheck('config-site-url', true, 'URL do SharePoint configurada.');

    const token = await acquireGraphToken();
    const claims = decodeJwt(token);

    pushCheck('auth', true, 'Token obtido com sucesso via Microsoft Graph.', {
      aud: claims?.aud || null,
      roles: claims?.roles || [],
      idtyp: claims?.idtyp || null
    });

    const site = await getSite(siteUrl);

    pushCheck('site', true, `Site acessível: ${site.displayName || site.name || 'Sem nome'}.`, {
      siteId: site.id,
      siteName: site.displayName || site.name || '',
      siteWebUrl: site.webUrl || siteUrl
    });

    const listValidation = await validateListsAccess(site.id, ticketsId, timeEntriesId);

    for (const item of listValidation.checks) {
      checks.push(item);
    }

    return {
      ok: listValidation.ok,
      message: listValidation.ok
        ? 'Conexão validada com sucesso via Microsoft Graph.'
        : 'Conexão com o site ok, mas houve falha na validação das listas.',
      checks,
      details: {
        siteId: site.id,
        siteName: site.displayName || site.name || '',
        siteWebUrl: site.webUrl || siteUrl
      }
    };
  } catch (error) {
    pushCheck('unexpected-error', false, error?.message || 'Erro inesperado.');
    return {
      ok: false,
      message: error?.message || 'Erro inesperado ao testar conexão.',
      checks
    };
  }
}
async function getListItemById(siteId, listId, itemId) {
  const result = await graphRequest(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`);

  if (!result.ok) {
    throw new Error(
      `Falha ao obter item ${itemId} da lista ${listId} (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  return result.data;
}

async function acquireSharePointToken(siteUrl) {
  const hostname = extractHostname(siteUrl);
  const scope = `https://${hostname}/.default`;

  const result = await cca.acquireTokenByClientCredential({
    scopes: [scope]
  });

  if (!result?.accessToken) {
    throw new Error('Não foi possível obter token de acesso do SharePoint.');
  }

  return result.accessToken;
}

async function sharePointRequest(siteUrl, path, options = {}) {
  const token = await acquireSharePointToken(siteUrl);
  const normalizedSiteUrl = normalizeSiteUrl(siteUrl);

  const response = await fetch(`${normalizedSiteUrl}${path}`, {
    method: options.method || 'GET',
    headers: {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      Authorization: `Bearer ${token}`,
      ...(options.headers || {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  const raw = await response.text();

  let data = null;
  try {
    data = raw ? JSON.parse(raw) : null;
  } catch {
    data = raw || null;
  }

  return {
    ok: response.ok,
    status: response.status,
    data,
    raw
  };
}

function buildEnsureUserLoginName(email) {
  return `i:0#.f|membership|${String(email || '').trim().toLowerCase()}`;
}

async function ensureSiteUser(siteUrl, email) {
  const normalizedEmail = String(email || '').trim().toLowerCase();

  if (!normalizedEmail) {
    throw new Error('E-mail do usuário não informado para resolução no SharePoint.');
  }

  const result = await sharePointRequest(siteUrl, '/_api/web/ensureuser', {
    method: 'POST',
    body: {
      logonName: buildEnsureUserLoginName(normalizedEmail)
    }
  });

  if (!result.ok) {
    throw new Error(
      `Falha ao resolver usuário no SharePoint (${result.status}): ${JSON.stringify(result.data)}`
    );
  }

  const payload = result.data?.d || result.data || {};
  const userId = Number(payload.Id || payload.id || 0);

  if (!userId) {
    throw new Error(`Não foi possível obter o ID do usuário ${normalizedEmail} no SharePoint.`);
  }

  return {
    id: userId,
    title: payload.Title || '',
    email: payload.Email || normalizedEmail,
    loginName: payload.LoginName || buildEnsureUserLoginName(normalizedEmail),
    raw: payload
  };
}



module.exports = {
  getSite,
  getListById,
  getListItems,
  getListItemsByQuery,
  getListItemById,
  createListItem,
  updateListItem,
  validateListsAccess,
  getTokenDebugInfo,
  testConnection,
  ensureSiteUser
};
