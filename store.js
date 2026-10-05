/* Capa de datos.
   - Modo Microsoft 365: la base vive como un archivo JSON en una biblioteca de SharePoint
     (o en OneDrive) y se lee/escribe con Microsoft Graph usando el inicio de sesión corporativo.
     SharePoint conserva el historial de versiones del archivo; además se guarda una copia diaria.
   - Modo local: si no hay configuración de Microsoft 365, los datos quedan en este navegador.
   Concurrencia: cada registro lleva updatedAt; al guardar se usa el eTag del archivo y,
   si alguien guardó antes, se fusiona registro por registro (gana el cambio más reciente). */
(function (root) {
  const U = root.U;
  const CFG = root.APP_CONFIG || {};
  const COLLS = ['proyectos', 'actividades', 'pendientes', 'hse', 'contratistas', 'contratos', 'actas', 'bitacora'];
  const CACHE_KEY = 'enelctl.cache.v1';

  const emptyDB = () => ({
    meta: {
      version: 1, updatedAt: 0, corte: U.todayISO(),
      pesos: { informe: 0.05, presupuesto: 0.05, preop: 0.10, obra: 0.70, liquidacion: 0.05, dossier: 0.05 },
      hseRequisitos: null
    },
    proyectos: [], actividades: [], pendientes: [], hse: [], contratistas: [], contratos: [], actas: [], bitacora: []
  });

  const S = {
    db: emptyDB(),
    mode: (CFG.msal && CFG.msal.clientId) ? 'm365' : 'local',
    user: { name: '', email: '' },
    status: 'idle',          // idle | loading | saving | saved | dirty | error | offline
    lastSaved: null, error: '',
    listeners: new Set(),
    _etag: null, _dirty: false, _msal: null, _account: null, _base: null,
    rev: 0                   // aumenta cuando llegan cambios de otra persona (para refrescar la vista)
  };

  const emit = () => S.listeners.forEach(f => { try { f(S); } catch (e) { console.error(e); } });
  S.on = (f) => { S.listeners.add(f); return () => S.listeners.delete(f); };
  const setStatus = (st, err = '') => { S.status = st; S.error = err; emit(); };

  function normalize(db) {
    const base = emptyDB();
    const out = Object.assign(base, db || {});
    out.meta = Object.assign(emptyDB().meta, (db && db.meta) || {});
    COLLS.forEach(c => { if (!Array.isArray(out[c])) out[c] = []; });
    return out;
  }

  /* ── Fusión registro por registro ── */
  function merge(a, b) {
    const out = emptyDB();
    out.meta = (a.meta.updatedAt || 0) >= (b.meta.updatedAt || 0) ? a.meta : b.meta;
    COLLS.forEach(c => {
      const m = new Map();
      [...(a[c] || []), ...(b[c] || [])].forEach(r => {
        const prev = m.get(r.id);
        if (!prev || (r.updatedAt || 0) > (prev.updatedAt || 0)) m.set(r.id, r);
      });
      out[c] = [...m.values()];
    });
    return out;
  }

  /* ── API de registros ── */
  S.stamp = (r) => Object.assign(r, { updatedAt: Date.now(), updatedBy: S.user.name || S.user.email || 'usuario' });
  S.list = (c, f) => S.db[c].filter(r => !r.deleted && (!f || f(r)));
  S.get = (c, id) => S.db[c].find(r => r.id === id && !r.deleted) || null;
  S.upsert = (c, rec) => {
    if (!rec.id) rec.id = U.uid(c.slice(0, 2));
    S.stamp(rec);
    const i = S.db[c].findIndex(r => r.id === rec.id);
    if (i >= 0) S.db[c][i] = Object.assign(S.db[c][i], rec); else S.db[c].push(rec);
    touch();
    return rec;
  };
  S.remove = (c, id) => {
    const r = S.db[c].find(x => x.id === id);
    if (r) { r.deleted = true; S.stamp(r); touch(); }
  };
  S.setMeta = (patch) => { Object.assign(S.db.meta, patch); touch(); };
  S.log = (proyectoId, texto) => S.upsert('bitacora', { proyectoId, fecha: new Date().toISOString(), texto });

  function touch() {
    S.db.meta.updatedAt = Date.now();
    S._dirty = true;
    writeCache();
    setStatus('dirty');
    scheduleSave();
  }
  function writeCache() {
    try { localStorage.setItem(CACHE_KEY, JSON.stringify({ mode: S.mode, user: S.user.email || '', etag: S._etag, dirty: S._dirty, db: S.db })); }
    catch (e) { console.warn('No se pudo escribir la caché local', e); }
  }
  function readCache() {
    try {
      const x = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null');
      if (!x || x.mode !== S.mode) return null;
      // La caché solo sirve al mismo usuario que la creó: otra cuenta en el mismo equipo no la ve
      if (S.mode === 'm365' && x.user !== (S.user.email || '')) { localStorage.removeItem(CACHE_KEY); return null; }
      return x;
    }
    catch (e) { return null; }
  }
  const scheduleSave = U.debounce(() => S.save(), 1800);

  /* ── Microsoft 365 / Graph ── */
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const scopes = () => (CFG.msal.scopes && CFG.msal.scopes.length) ? CFG.msal.scopes
    : (CFG.storage && CFG.storage.type === 'onedrive' ? ['User.Read', 'Files.ReadWrite'] : ['User.Read', 'Files.ReadWrite.All', 'Sites.ReadWrite.All']);

  async function initMsal() {
    if (!root.msal) throw new Error('No cargó la librería de Microsoft (MSAL). Revisa la conexión a internet.');
    S._msal = new root.msal.PublicClientApplication({
      auth: {
        clientId: CFG.msal.clientId,
        authority: 'https://login.microsoftonline.com/' + (CFG.msal.tenantId || 'organizations'),
        redirectUri: CFG.msal.redirectUri || (location.origin + location.pathname),
        navigateToLoginRequestUrl: false
      },
      cache: { cacheLocation: 'localStorage' }
    });
    await S._msal.initialize();
    const res = await S._msal.handleRedirectPromise();
    if (res && res.account) S._msal.setActiveAccount(res.account);
    S._account = S._msal.getActiveAccount() || S._msal.getAllAccounts()[0] || null;
    if (S._account) {
      S._msal.setActiveAccount(S._account);
      S.user = { name: S._account.name || S._account.username, email: S._account.username };
    }
    if (/[#&](code|error)=/.test(location.hash)) history.replaceState(null, '', location.pathname + '#/tablero');
    return !!S._account;
  }
  S.login = () => S._msal.loginRedirect({ scopes: scopes(), prompt: 'select_account' });
  S.logout = async () => {
    if (S._dirty) { try { await S.save(); } catch (e) { /* se intenta guardar antes de salir */ } }
    if (S._dirty && !confirm('Hay cambios que no se han podido guardar en SharePoint. Si cierras sesión se perderán. ¿Salir de todas formas?')) return;
    localStorage.removeItem(CACHE_KEY);          // no dejar datos en equipos compartidos
    Object.keys(localStorage).filter(k => k.startsWith('enelctl.backup.')).forEach(k => localStorage.removeItem(k));
    return S._msal ? S._msal.logoutRedirect({ account: S._account }) : null;
  };

  async function token() {
    try {
      const r = await S._msal.acquireTokenSilent({ scopes: scopes(), account: S._account });
      return r.accessToken;
    } catch (e) {
      if (e instanceof root.msal.InteractionRequiredAuthError) { await S._msal.acquireTokenRedirect({ scopes: scopes() }); }
      throw e;
    }
  }
  async function graph(path, opts = {}) {
    const t = await token();
    const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
      ...opts, headers: Object.assign({ Authorization: 'Bearer ' + t }, opts.headers || {})
    });
    return res;
  }
  const enc = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

  async function base() {
    if (S._base) return S._base;
    const st = CFG.storage || {};
    if (st.type === 'onedrive') { S._base = '/me/drive/root:/' + enc(st.folder || 'Control ENEL'); return S._base; }
    const r = await graph(`/sites/${st.hostname}:${st.sitePath}`);
    if (!r.ok) throw new Error(`No se encontró el sitio de SharePoint ${st.hostname}${st.sitePath} (${r.status}). Revisa config.js y que tengas acceso al sitio.`);
    const site = await r.json();
    S._base = `/sites/${site.id}/drive/root:/` + enc(st.folder || 'Control ENEL');
    if (st.library) {   // biblioteca distinta a "Documentos"
      const dl = await graph(`/sites/${site.id}/drives`);
      const drives = (await dl.json()).value || [];
      const d = drives.find(x => U.norm(x.name) === U.norm(st.library));
      if (!d) throw new Error(`No existe la biblioteca “${st.library}” en el sitio.`);
      S._base = `/drives/${d.id}/root:/` + enc(st.folder || 'Control ENEL');
    }
    return S._base;
  }
  const fileName = () => (CFG.storage && CFG.storage.file) || 'control-enel-db.json';

  async function remoteRead() {
    const b = await base();
    const r = await graph(`${b}/${encodeURIComponent(fileName())}`);
    if (r.status === 404) return { db: null, etag: null };
    if (!r.ok) throw new Error('No se pudo leer la base en Microsoft 365 (' + r.status + ').');
    const item = await r.json();
    const c = await fetch(item['@microsoft.graph.downloadUrl']);
    return { db: normalize(await c.json()), etag: item.eTag, webUrl: item.webUrl };
  }
  async function remoteWrite(db, etag) {
    const b = await base();
    const url = `${b}/${encodeURIComponent(fileName())}:/content` + (etag ? '' : '?@microsoft.graph.conflictBehavior=fail');
    const headers = { 'Content-Type': 'application/json' };
    if (etag) headers['If-Match'] = etag;
    const r = await graph(url, { method: 'PUT', headers, body: JSON.stringify(db) });
    if (r.status === 412 || r.status === 409) return { conflict: true };
    if (!r.ok) throw new Error('No se pudo guardar en Microsoft 365 (' + r.status + ').');
    const item = await r.json();
    return { etag: item.eTag, webUrl: item.webUrl };
  }
  async function dailyBackup(db) {
    const key = 'enelctl.backup.' + U.todayISO();
    if (localStorage.getItem(key)) return;
    try {
      const b = await base();
      const name = 'copias/' + fileName().replace('.json', '') + '-' + U.todayISO() + '.json';
      const r = await graph(`${b}/${enc(name)}:/content`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(db) });
      if (r.ok) localStorage.setItem(key, '1');
    } catch (e) { console.warn('Copia diaria no realizada', e); }
  }

  /* Adjuntos (evidencias HSE, soportes) a la carpeta de la base */
  S.upload = async (file, carpeta) => {
    if (S.mode !== 'm365') throw new Error('Los adjuntos requieren conexión a Microsoft 365.');
    if (file.size > 4 * 1024 * 1024) throw new Error('El archivo supera 4 MB. Súbelo directamente a SharePoint y pega el enlace.');
    const b = await base();
    const safe = file.name.replace(/[#%&{}\\<>*?/$!'":@+`|=]/g, '_');
    const path = 'adjuntos/' + (carpeta || 'general') + '/' + Date.now() + '-' + safe;
    const r = await graph(`${b}/${enc(path)}:/content`, { method: 'PUT', headers: { 'Content-Type': file.type || 'application/octet-stream' }, body: file });
    if (!r.ok) throw new Error('No se pudo subir el adjunto (' + r.status + ').');
    const it = await r.json();
    return { nombre: file.name, url: it.webUrl };
  };

  /* ── Ciclo de vida ── */
  S.init = async () => {
    setStatus('loading');
    if (S.mode === 'local') {
      const c = readCache();
      S.db = normalize(c && c.db);
      S.user = { name: localStorage.getItem('enelctl.user') || '', email: '' };
      setStatus('saved');
      return { needLogin: false };
    }
    const ok = await initMsal();
    if (!ok) { setStatus('idle'); return { needLogin: true }; }
    const c = readCache();
    if (c) { S.db = normalize(c.db); S._etag = c.etag; S._dirty = !!c.dirty; }
    try {
      await S.pull();
      if (S._dirty) S.save();
    } catch (e) {
      console.error(e);
      setStatus(c ? 'offline' : 'error', e.message);
    }
    return { needLogin: false };
  };

  S.pull = async () => {
    if (S.mode !== 'm365') return;
    setStatus('loading');
    const r = await remoteRead();
    if (r.db && r.etag !== S._etag) {
      S.db = S._dirty ? merge(S.db, r.db) : r.db;
      S.rev++;
    }
    if (r.db) { S._etag = r.etag; S.webUrl = r.webUrl; }
    writeCache();
    setStatus(S._dirty ? 'dirty' : 'saved');
  };

  let saving = null;
  S.save = async () => {
    if (S.mode === 'local') { S._dirty = false; writeCache(); S.lastSaved = new Date(); setStatus('saved'); return; }
    if (saving) return saving;
    if (!S._dirty) return;
    saving = (async () => {
      setStatus('saving');
      try {
        for (let intento = 0; intento < 4; intento++) {
          const snapshot = JSON.parse(JSON.stringify(S.db));
          const w = await remoteWrite(snapshot, S._etag);
          if (w.conflict) {                       // alguien guardó antes: fusionar y reintentar
            const r = await remoteRead();
            if (r.db) { S.db = merge(S.db, r.db); S._etag = r.etag; S.rev++; }
            continue;
          }
          S._etag = w.etag; S.webUrl = w.webUrl || S.webUrl;
          if ((S.db.meta.updatedAt || 0) <= (snapshot.meta.updatedAt || 0)) S._dirty = false;
          S.lastSaved = new Date();
          writeCache();
          setStatus(S._dirty ? 'dirty' : 'saved');
          dailyBackup(snapshot);
          return;
        }
        throw new Error('Varias personas están guardando al tiempo. Intenta de nuevo en unos segundos.');
      } catch (e) {
        console.error(e);
        setStatus(navigator.onLine ? 'error' : 'offline', e.message);
      } finally { saving = null; if (S._dirty && S.status !== 'error') scheduleSave(); }
    })();
    return saving;
  };

  S.replaceAll = (db) => { S.db = normalize(db); touch(); };
  S.exportJSON = () => JSON.stringify(S.db, null, 1);
  S.merge = merge;
  S.normalize = normalize;
  S.COLLS = COLLS;

  window.addEventListener('focus', () => { if (S.mode === 'm365' && S._account && !S._dirty && S.status !== 'saving') S.pull().catch(() => {}); });
  window.addEventListener('online', () => { if (S._dirty) S.save(); });
  window.addEventListener('beforeunload', (e) => { if (S.mode === 'm365' && S._dirty) { e.preventDefault(); e.returnValue = ''; } });

  root.Store = S;
})(window);
