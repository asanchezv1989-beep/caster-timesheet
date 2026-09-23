// Conexión con OneDrive/SharePoint vía Microsoft Graph.
// Inicio de sesión con MSAL (redirect: funciona dentro de la PWA en iPhone).
// El Client ID sale de la app registrada en Microsoft Entra; se puede pegar en la
// pestaña Archivo (se guarda en este dispositivo) o fijar en CLIENT_ID.

const CLIENT_ID = '191f59ed-6671-4ca7-b327-46b1bb57e3f1'; // app "Novelis" registrada en Entra (Southern Welding)
const SCOPES = ['Files.ReadWrite', 'User.Read'];
const GRAPH = 'https://graph.microsoft.com/v1.0';

let pca = null;
let account = null;

export function clientId() {
  try { return localStorage.getItem('cts_client_id') || CLIENT_ID; } catch { return CLIENT_ID; }
}
export function setClientId(id) {
  try { id ? localStorage.setItem('cts_client_id', id.trim()) : localStorage.removeItem('cts_client_id'); } catch {}
  pca = null;
  account = null;
}
export const configured = () => /^[0-9a-f-]{36}$/i.test(clientId());
export const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, '');
export const user = () => account;

function loadMsal() {
  if (globalThis.msal) return Promise.resolve();
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'vendor/msal-browser.min.js';
    s.onload = res;
    s.onerror = () => rej(new Error('No se pudo cargar MSAL'));
    document.head.append(s);
  });
}

// Llamar al arrancar: procesa la vuelta del login y devuelve la acción pendiente (si había)
export async function init() {
  if (!configured()) return null;
  await loadMsal();
  if (!pca) {
    pca = new globalThis.msal.PublicClientApplication({
      auth: { clientId: clientId(), authority: 'https://login.microsoftonline.com/organizations', redirectUri: redirectUri() },
      cache: { cacheLocation: 'localStorage' },
    });
    await pca.initialize();
  }
  const r = await pca.handleRedirectPromise();
  account = r?.account || pca.getAllAccounts()[0] || null;
  let pending = null;
  try {
    pending = JSON.parse(sessionStorage.getItem('cts_pending') || 'null');
    sessionStorage.removeItem('cts_pending');
  } catch {}
  return r ? pending : null;
}

export async function signIn(pendingAction) {
  if (!pca) await init();
  try { sessionStorage.setItem('cts_pending', JSON.stringify(pendingAction || null)); } catch {}
  await pca.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
}
export async function signOut() {
  if (!pca) await init();
  const a = account;
  account = null;
  if (a) await pca.logoutRedirect({ account: a, postLogoutRedirectUri: redirectUri() });
}

class NeedLogin extends Error {}
export const isNeedLogin = (e) => e instanceof NeedLogin;

async function token() {
  if (!pca) await init();
  if (!account) throw new NeedLogin('login');
  try {
    return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
  } catch {
    throw new NeedLogin('login');
  }
}

async function g(path, { method = 'GET', body, headers = {}, raw = false } = {}) {
  const url = path.startsWith('http') ? path : GRAPH + path;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, {
      method,
      headers: { Authorization: 'Bearer ' + (await token()), ...(body ? { 'Content-Type': 'application/json' } : {}), ...headers },
      body: body ? JSON.stringify(body) : undefined,
    });
    if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 5) {
      const wait = +(res.headers.get('Retry-After') || 2 ** attempt);
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (res.status === 401) throw new NeedLogin('login');
    if (!res.ok) {
      let msg = res.status + ' ' + res.statusText;
      try { msg = (await res.json()).error?.message || msg; } catch {}
      throw new Error(msg);
    }
    if (raw) return res;
    return res.status === 204 ? null : res.json();
  }
}

const itemPath = (ref) => `/drives/${ref.driveId}/items/${ref.itemId}`;
const toRef = (it) => ({
  driveId: it.parentReference?.driveId || it.remoteItem?.parentReference?.driveId,
  itemId: it.remoteItem?.id || it.id,
  name: it.name,
  webUrl: it.webUrl,
  modified: it.lastModifiedDateTime,
  folder: it.parentReference?.path?.replace(/^\/drive\/root:?/, '') || '',
});

// Archivos .xlsx: recientes, o búsqueda por texto
export async function findFiles(q) {
  const sel = '$select=id,name,webUrl,lastModifiedDateTime,parentReference,remoteItem,file';
  const r = q
    ? await g(`/me/drive/root/search(q='${encodeURIComponent(q.replace(/'/g, "''"))}')?${sel}&$top=100`)
    : await g(`/me/drive/recent?$top=100`);
  return (r.value || [])
    .filter((it) => /\.xlsx$/i.test(it.name) && (it.file || it.remoteItem?.file))
    .map(toRef)
    .filter((f) => f.driveId && f.itemId)
    .sort((a, b) => (b.modified || '').localeCompare(a.modified || ''));
}

export async function download(ref) {
  const res = await g(itemPath(ref) + '/content', { raw: true });
  return res.arrayBuffer();
}

export async function meta(ref) {
  return toRef(await g(itemPath(ref) + '?$select=id,name,webUrl,lastModifiedDateTime,parentReference'));
}

// Lee la hoja en vivo (valores + fórmulas) para ubicar filas y detectar cambios ajenos
export async function readSheet(ref, sheet, session) {
  const h = session ? { 'workbook-session-id': session } : {};
  const ws = `/workbook/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;
  const r = await g(itemPath(ref) + ws + `/usedRange?$select=address,values,formulas`, { headers: h });
  const m = r.address.match(/!\$?([A-Z]+)\$?(\d+)/);
  return { values: r.values, formulas: r.formulas, startCol: m ? m[1] : 'A', startRow: m ? +m[2] : 1 };
}

export async function openSession(ref) {
  const r = await g(itemPath(ref) + '/workbook/createSession', { method: 'POST', body: { persistChanges: true } });
  return r.id;
}
export async function closeSession(ref, session) {
  try { await g(itemPath(ref) + '/workbook/closeSession', { method: 'POST', headers: { 'workbook-session-id': session } }); } catch {}
}

// Escribe celdas [{addr:'BT15', v}] en lotes de 20 (Graph $batch). Devuelve las que fallaron.
export async function writeCells(ref, sheet, session, cells, onProgress) {
  const ws = `/workbook/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;
  const failed = [];
  for (let i = 0; i < cells.length; i += 20) {
    const chunk = cells.slice(i, i + 20);
    const r = await g('/$batch', {
      method: 'POST',
      body: {
        requests: chunk.map((c, k) => ({
          id: String(k),
          method: 'PATCH',
          url: itemPath(ref) + ws + `/range(address='${c.addr}')`,
          headers: { 'Content-Type': 'application/json', 'workbook-session-id': session },
          body: { values: [[c.v ?? '']] },
        })),
      },
    });
    for (const resp of r.responses || []) {
      if (resp.status >= 400) failed.push({ ...chunk[+resp.id], error: resp.body?.error?.message || String(resp.status) });
    }
    onProgress?.(Math.min(i + 20, cells.length), cells.length);
  }
  return failed;
}
