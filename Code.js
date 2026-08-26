// TaskTime — server (Google Apps Script WebApp).
//
// Two spreadsheets:
//   - MASTER (one, configured via SPREADSHEET_ID Script Property or
//     configurarSpreadsheetMaestro API): holds the auth tables (Usuarios,
//     Spreadsheets, HojasUsuarios, Tokens, Config).
//   - PER-USER (one per linked user, listed in HojasUsuarios): holds the
//     time-tracking data (Tasks, Labels, TimeEntries). Auto-created by
//     ensureSchema_ when bootstrap runs.
//
// API surface: doGet / doPost dispatch via dispatchApi_(). Frontend calls
// `call('action', ...args)` from a static page hosted elsewhere, with
// `{ token }` pulled from localStorage.

const SHEETS = {
  Tasks:              ['id', 'name', 'labelId', 'completed', 'createdAt'],
  Labels:             ['id', 'name', 'color',   'createdAt'],
  TimeEntries:        ['id', 'taskId', 'startTime', 'endTime', 'durationMinutes', 'source', 'notes', 'createdAt'],
  ActiveTimer:        ['taskId', 'startTime'],
  // ponytail: rutinas viven en tablas separadas de Tasks. Routines = "qué se repite".
  // Completions se reparten en tres hojas (una por periodo) con shape [periodKey, ids]
  // donde ids es la lista delimitada por comas de los routineId marcados en ese periodoKey.
  // periodKey natural: yyyy-mm-dd / yyyy-Www / yyyy-mm.
  Routines:           ['id', 'name', 'period', 'labelId', 'createdAt'],
  DailyCompletions:   ['periodKey', 'ids'],
  WeeklyCompletions:  ['periodKey', 'ids'],
  MonthlyCompletions: ['periodKey', 'ids']
};

const AUTH_SCHEMA = {
  Usuarios:       ['username', 'password_hash', 'salt', 'rol', 'activo', 'fecha_creacion'],
  Spreadsheets:   ['spreadsheet_id', 'nombre', 'descripcion', 'fecha_alta'],
  HojasUsuarios:  ['username', 'spreadsheet_id', 'por_defecto', 'fecha_alta'],
  Tokens:         ['token', 'username', 'fecha_creacion'],
  Config:         ['clave', 'valor']
};

const ROLES = { ADMIN: 'admin', BASICO: 'basico' };
const APP_VERSION = 'ms-http-1';
const MASTER_PROP_KEY = 'SPREADSHEET_ID';
const CONFIG_MASTER_SHEET_ID = 'MASTER_SPREADSHEET_ID';
const DEFAULT_ADMIN_USERNAME = 'admin';
const DEFAULT_ADMIN_PASSWORD = 'admin1234';

const PALETTE = ['#4285F4','#EA4335','#FBBC04','#34A853','#A142F4','#24C1E0','#FF6D00','#7B1FA2'];

// ───────── Script Cache (CacheService) ─────────
// Caché compartida entre ejecuciones del Web App. Reduce lecturas a Sheets en
// el path caliente de cada request autenticada: validar token y resolver hoja
// activa pasan de abrir el master spreadsheet a un CacheService.get.
// Límites GAS: ~100 KB/entrada, TTL máx. 21600 s. Fallos de caché se ignoran.
const CACHE_TTL_TOKEN_SEC = 30 * 60;       // 30 min — validación de sesión
const CACHE_TTL_AUTH_SHEET_SEC = 5 * 60;   // 5 min — hojas del master sheet
const CACHE_MAX_JSON_CHARS = 90000;        // margen bajo el límite ~100 KB

function scriptCache_() { return CacheService.getScriptCache(); }

function cachePutJson_(key, value, ttlSec) {
  try {
    const json = JSON.stringify(value);
    if (!json || json.length > CACHE_MAX_JSON_CHARS) return false;
    scriptCache_().put(key, json, Math.min(ttlSec || CACHE_TTL_AUTH_SHEET_SEC, 21600));
    return true;
  } catch (e) { return false; }
}

function cacheGetJson_(key) {
  try {
    const raw = scriptCache_().get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) { return null; }
}

function cacheRemove_(key) {
  try { scriptCache_().remove(key); } catch (e) { /* ignore */ }
}

// ponytail: payloads auth (Usuarios/Tokens/HojasUsuarios/Spreadsheets/Config) son
// pequeños (<10 KB típicamente), gzip solo mete overhead. Si hojas auth crecen
// hasta chocar con el límite de ~100 KB, añadir prefijo gzip como en
// finanzasFamiliaMs (cacheSerialize_ / cacheDeserialize_ con Utilities.gzip).
function tokenCacheKey_(token) { return 'tok:' + String(token || '').trim(); }
function authSheetCacheKey_(nombre) { return 'auth:' + String(nombre || ''); }

// ───────── Auth primitives ─────────

function bytesHex_(bytes) {
  return bytes.map(function (b) {
    const h = (b < 0 ? b + 256 : b).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}

function sha256Hex_(text) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytesHex_(digest);
}

// 1500 iteraciones: en GAS cada Utilities.computeDigest es costoso.
// 12000 hacía que login/cambio de contraseña tardara varios segundos.
// 1500 sigue siendo resistente a fuerza bruta offline y es ~8× más rápido.
// Usuarios existentes con hash antiguo deberán resetear contraseña (admin)
// o cambiarla una vez (el nuevo hash se escribe al cambiar).
const PASSWORD_ITERATIONS = 1500;

function passwordHash_(password, salt) {
  let h = String(salt || '') + '|' + String(password || '');
  for (let i = 0; i < PASSWORD_ITERATIONS; i++) h = sha256Hex_(h);
  return h;
}

function isoAhora_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Madrid', "yyyy-MM-dd'T'HH:mm:ss");
}

// ───────── Spreadsheet access ─────────

function getMasterSpreadsheetId_() {
  // ping() llama a esto en cada arranque de página; leer Script Property es
  // instantáneo y evita abrir el master spreadsheet (que puede estar inaccesible
  // y bloquear `SpreadsheetApp.openById` durante minutos).
  return PropertiesService.getScriptProperties().getProperty(MASTER_PROP_KEY) || '';
}

function setMasterSpreadsheetId_(id) {
  const limpio = String(id || '').trim();
  if (!limpio) throw new Error('Indica un spreadsheet ID válido');
  PropertiesService.getScriptProperties().setProperty(MASTER_PROP_KEY, limpio);
  return limpio;
}

// Caché de SpreadsheetApp por request (openById es caro).
let _ssCache = {};          // id → Spreadsheet
let _userReadCache = {};    // sheetName → rows[] (datos de usuario)
let _headersCache = {};     // sheetName → headers[]
let _sheetObjCache = {};    // sheetName → Sheet

function openSsById_(id) {
  const key = String(id || '');
  if (!key) throw new Error('spreadsheet id vacío');
  if (_ssCache[key]) return _ssCache[key];
  const ss = SpreadsheetApp.openById(key);
  _ssCache[key] = ss;
  return ss;
}

function authSs_() {
  const id = getMasterSpreadsheetId_();
  if (!id) throw new Error('Master spreadsheet no configurado. Llama a configurarSpreadsheetMaestro(id) primero.');
  return openSsById_(id);
}

function ssActiva_() {
  // Hoja de datos del usuario activo. Si _currentSheetId está resuelto (por
  // _authadmin / bootstrap), abre esa; en caso contrario cae al master.
  if (_currentSheetId) {
    try { return openSsById_(_currentSheetId); }
    catch (e) { /* id inválido: fallback */ }
  }
  return authSs_();
}

function getUserSheet_(name) {
  if (_sheetObjCache[name]) return _sheetObjCache[name];
  const s = ssActiva_().getSheetByName(name);
  if (s) _sheetObjCache[name] = s;
  return s;
}

function invalidateUserDataCache_(name) {
  if (name) {
    delete _userReadCache[name];
    delete _headersCache[name];
    // no borramos _sheetObjCache: el objeto Sheet sigue válido
  } else {
    _userReadCache = {};
    _headersCache = {};
  }
}

// ───────── Estado request-scoped ─────────

let _currentToken = '';
let _currentSheetId = '';

function currentUser_() {
  if (!_currentToken) return '';
  return validarTokenSesion_(_currentToken);
}

function requireUsuario_() {
  const u = currentUser_();
  if (!u) throw new Error('No autenticado');
  return u;
}

function currentRol_() {
  const u = currentUser_();
  if (!u) return '';
  const row = buscarUsuario_(u);
  return row ? String(row.rol || ROLES.BASICO) : '';
}

function requireAdmin_() {
  if (currentRol_() !== ROLES.ADMIN) throw new Error('Solo admin puede realizar esta acción');
}

// Endpoints disponibles aunque el usuario autenticado no tenga hojas
// vinculadas (no son admin). El resto fallan rápido para no filtrar
// datos del spreadsheet equivocado.
const ALWAYS_ALLOWED_FOR_NO_HOJAS = new Set([
  'authStatus', 'loginUsuario', 'logoutUsuario', 'ping',
  'configurarSpreadsheetMaestro', 'bootstrapBase', 'bootstrap',
  'listarMisHojas', 'cambiarHojaActiva', 'cambiarMiContrasena'
]);

function usuarioTieneHojas_(username) {
  if (!username) return false;
  return leerHojasUsuarios_().some(l =>
    String(l.username || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()
  );
}

function _authadmin(token, fnName, ...args) {
  const username = validarTokenSesion_(token);
  if (!username) throw new Error('No autenticado');
  _currentToken = String(token || '').trim();
  // Limpiar cachés de datos de usuario al resolver la hoja (por si el token
  // llega a un request con otra hoja activa).
  _userReadCache = {};
  _headersCache = {};
  _sheetObjCache = {};
  try { _currentSheetId = resolverHojaActivaId_(username); }
  catch (e) { _currentSheetId = ''; }
  const usuario = buscarUsuario_(username);
  const rol = usuario ? String(usuario.rol || ROLES.BASICO) : ROLES.BASICO;
  if (rol !== ROLES.ADMIN && !usuarioTieneHojas_(username) && !ALWAYS_ALLOWED_FOR_NO_HOJAS.has(fnName)) {
    throw new Error('No tienes hojas de cálculo asignadas. Pide al administrador que vincule una hoja.');
  }
  const fn = globalThis[fnName];
  if (typeof fn !== 'function') throw new Error('Función no encontrada: ' + fnName);
  return fn.apply(null, args);
}

// ───────── Auth sheet helpers ─────────

function asegurarAuthHojaGenerica_(nombre) {
  const ss = authSs_();
  const cab = AUTH_SCHEMA[nombre];
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

const _authReadCache = {};
function _authCacheKey_(nombre) { return nombre; }

function leerAuthHojaGenerica_(nombre) {
  if (_authReadCache[_authCacheKey_(nombre)]) return cloneRows_(_authReadCache[_authCacheKey_(nombre)]);
  // Cache entre requests: evita abrir el master spreadsheet en casi todas las
  // llamadas autenticadas (Tokens, Usuarios, HojasUsuarios, Spreadsheets, Config).
  const cached = cacheGetJson_(authSheetCacheKey_(nombre));
  if (cached) {
    _authReadCache[_authCacheKey_(nombre)] = cached;
    return cloneRows_(cached);
  }
  const h = asegurarAuthHojaGenerica_(nombre);
  const valores = h.getDataRange().getValues();
  if (valores.length < 2) {
    _authReadCache[_authCacheKey_(nombre)] = [];
    cachePutJson_(authSheetCacheKey_(nombre), [], CACHE_TTL_AUTH_SHEET_SEC);
    return [];
  }
  const cab = valores[0];
  const rows = valores.slice(1).map(fila => {
    const o = {};
    cab.forEach((k, i) => (o[k] = fila[i]));
    return o;
  });
  _authReadCache[_authCacheKey_(nombre)] = rows;
  cachePutJson_(authSheetCacheKey_(nombre), rows, CACHE_TTL_AUTH_SHEET_SEC);
  return cloneRows_(rows);
}

function escribirAuthHojaGenerica_(nombre, filas) {
  const h = asegurarAuthHojaGenerica_(nombre);
  const cab = AUTH_SCHEMA[nombre];
  h.clearContents();
  const filasSafe = filas || [];
  const matriz = [cab].concat(filasSafe.map(f => cab.map(k => f[k] != null ? f[k] : '')));
  if (matriz.length) h.getRange(1, 1, matriz.length, cab.length).setValues(matriz);
  h.setFrozenRows(1);
  _authReadCache[_authCacheKey_(nombre)] = cloneRows_(filasSafe);
  cachePutJson_(authSheetCacheKey_(nombre), filasSafe, CACHE_TTL_AUTH_SHEET_SEC);
}

function cloneRows_(rows) {
  return (rows || []).map(r => Object.assign({}, r));
}

function invalidarAuthCache_(nombre) {
  if (nombre) {
    delete _authReadCache[_authCacheKey_(nombre)];
    cacheRemove_(authSheetCacheKey_(nombre));
  }
}

function leerUsuariosAuth_() {
  const rows = leerAuthHojaGenerica_('Usuarios');
  return rows.map(normalizarUsuarioAuth_);
}

function normalizarUsuarioAuth_(u) {
  if (!u) return u;
  const rol = String(u.rol || '').trim();
  u.rol = rol || (String(u.username || '').trim().toLowerCase() === DEFAULT_ADMIN_USERNAME ? ROLES.ADMIN : ROLES.BASICO);
  return u;
}

function escribirUsuariosAuth_(filas) {
  const normalizadas = (filas || []).map(f => normalizarUsuarioAuth_(Object.assign({}, f)));
  escribirAuthHojaGenerica_('Usuarios', normalizadas);
}

function leerSpreadsheets_() { return leerAuthHojaGenerica_('Spreadsheets'); }
function escribirSpreadsheets_(filas) { escribirAuthHojaGenerica_('Spreadsheets', filas); }
function leerHojasUsuarios_() { return leerAuthHojaGenerica_('HojasUsuarios'); }
function escribirHojasUsuarios_(filas) { escribirAuthHojaGenerica_('HojasUsuarios', filas); }

function resolverHojaActivaId_(username) {
  username = String(username || '').trim();
  if (!username) return '';
  const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() === username.toLowerCase());
  const defecto = links.find(l => String(l.por_defecto) === 'true' || String(l.por_defecto) === true);
  if (defecto) return String(defecto.spreadsheet_id);
  if (links.length) return String(links[0].spreadsheet_id);
  return '';
}

function listarHojasDelUsuario_(username) {
  const links = leerHojasUsuarios_().filter(l =>
    String(l.username || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()
  );
  const todas = leerSpreadsheets_();
  const porId = {};
  todas.forEach(s => { porId[String(s.spreadsheet_id)] = s; });
  return links.map(l => {
    const meta = porId[String(l.spreadsheet_id)] || {};
    return {
      spreadsheet_id: String(l.spreadsheet_id),
      nombre: meta.nombre || '(sin nombre)',
      descripcion: meta.descripcion || '',
      por_defecto: String(l.por_defecto) === 'true' || l.por_defecto === true
    };
  });
}

function obtenerConfig_(clave) {
  const fila = leerAuthHojaGenerica_('Config').find(r => String(r.clave || '') === clave);
  return fila ? String(fila.valor || '').trim() : '';
}

function guardarConfig_(clave, valor) {
  const filas = leerAuthHojaGenerica_('Config');
  const existente = filas.find(r => String(r.clave || '') === clave);
  if (existente) existente.valor = String(valor || '').trim();
  else filas.push({ clave: clave, valor: String(valor || '').trim() });
  escribirAuthHojaGenerica_('Config', filas);
  return String(valor || '').trim();
}

// ───────── Sesión ─────────

function crearTokenSesion_(username) {
  const token = Utilities.getUuid();
  const usernameNorm = String(username || '').trim();
  const filas = leerAuthHojaGenerica_('Tokens');
  filas.push({ token: token, username: usernameNorm, fecha_creacion: isoAhora_() });
  escribirAuthHojaGenerica_('Tokens', filas);
  // Pre-rellenar caché para que la siguiente request valide sin abrir Sheets.
  cachePutJson_(tokenCacheKey_(token), { u: usernameNorm }, CACHE_TTL_TOKEN_SEC);
  return token;
}

function validarTokenSesion_(token) {
  const t = String(token || '').trim();
  if (!t) return '';
  // 1) Cache entre requests (token válido 30 min → evita leer la hoja Tokens).
  const hit = cacheGetJson_(tokenCacheKey_(t));
  if (hit && hit.u) return String(hit.u);
  // 2) Fallback a la hoja (con caché de hoja + memoria de request).
  const fila = leerAuthHojaGenerica_('Tokens').find(r => String(r.token || '') === t);
  const username = fila ? String(fila.username || '').trim() : '';
  if (username) cachePutJson_(tokenCacheKey_(t), { u: username }, CACHE_TTL_TOKEN_SEC);
  return username;
}

function invalidarTokenSesion_(token) {
  const t = String(token || '').trim();
  if (!t) return;
  // Invalidar caché ANTES de reescribir para que un request concurrente no
  // rehidrate el token revocado desde la hoja vieja todavía cacheada.
  cacheRemove_(tokenCacheKey_(t));
  escribirAuthHojaGenerica_('Tokens', leerAuthHojaGenerica_('Tokens').filter(r => String(r.token || '') !== t));
}

// ───────── Usuarios ─────────

function asegurarUsuarios_() {
  asegurarAuthHojaGenerica_('Usuarios');
  const rows = leerUsuariosAuth_().filter(r => r.username);
  if (rows.length) return;

  // Seed admin por defecto. La contraseña inicial debe cambiarse en el primer login.
  const salt = Utilities.getUuid().replace(/-/g, '');
  escribirUsuariosAuth_([{
    username: DEFAULT_ADMIN_USERNAME,
    password_hash: passwordHash_(DEFAULT_ADMIN_PASSWORD, salt),
    salt: salt,
    rol: ROLES.ADMIN,
    activo: true,
    fecha_creacion: isoAhora_()
  }]);
}

function buscarUsuario_(username) {
  const target = String(username || '').trim().toLowerCase();
  if (!target) return null;
  return leerUsuariosAuth_().find(u => String(u.username || '').trim().toLowerCase() === target) || null;
}

function authStatus(token) {
  const tokenUser = token ? validarTokenSesion_(token) : '';
  if (tokenUser) _currentToken = String(token || '').trim();
  else _currentToken = '';
  const rol = tokenUser ? (buscarUsuario_(tokenUser) || {}).rol || ROLES.BASICO : '';
  return { authenticated: !!tokenUser, user: tokenUser || '', rol: rol || '' };
}

function passwordHashLegacy_(password, salt) {
  // Versión antigua (12000 iter) para migración transparente en el primer login.
  let h = String(salt || '') + '|' + String(password || '');
  for (let i = 0; i < 12000; i++) h = sha256Hex_(h);
  return h;
}

function loginUsuario(username, password) {
  asegurarUsuarios_();
  const user = String(username || '').trim();
  const pass = String(password || '');
  if (!user || !pass) throw new Error('Debes indicar usuario y contraseña');
  const found = buscarUsuario_(user);
  if (!found || String(found.activo) === 'false') throw new Error('Credenciales inválidas');
  const salt = String(found.salt || '');
  const stored = String(found.password_hash || '');
  let hash = passwordHash_(pass, salt);
  if (hash !== stored) {
    // Migración: ¿es un hash generado con las 12000 iteraciones antiguas?
    const legacy = passwordHashLegacy_(pass, salt);
    if (legacy !== stored) throw new Error('Credenciales inválidas');
    // Re-hash con el nuevo coste y persistir (una sola vez por usuario).
    const rows = leerUsuariosAuth_();
    const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
    if (idx >= 0) {
      rows[idx].password_hash = hash; // ya calculado con PASSWORD_ITERATIONS
      escribirUsuariosAuth_(rows);
    }
  }
  const token = crearTokenSesion_(found.username);
  return { ok: true, user: found.username, rol: String(found.rol || ROLES.BASICO), token: token };
}

function logoutUsuario(token) {
  invalidarTokenSesion_(token);
  _currentToken = '';
  return { ok: true };
}

function crearUsuarioAdmin(username, password, rol) {
  requireAdmin_();
  asegurarUsuarios_();
  const user = String(username || '').trim();
  const pass = String(password || '');
  const rolFinal = String(rol || ROLES.BASICO).trim().toLowerCase();
  if (!Object.values(ROLES).includes(rolFinal)) throw new Error('Rol inválido');
  if (!/^[a-zA-Z0-9_.-]{3,40}$/.test(user)) throw new Error('Usuario inválido (3-40, letras, números, _.-)');
  if (pass.length < 8) throw new Error('La contraseña debe tener mínimo 8 caracteres');
  if (buscarUsuario_(user)) throw new Error('Ese usuario ya existe');
  const salt = Utilities.getUuid().replace(/-/g, '');
  const rows = leerUsuariosAuth_();
  rows.push({
    username: user,
    password_hash: passwordHash_(pass, salt),
    salt: salt,
    rol: rolFinal,
    activo: true,
    fecha_creacion: isoAhora_()
  });
  escribirUsuariosAuth_(rows);
  return { ok: true, user: user, rol: rolFinal };
}

function listarUsuariosAdmin() {
  requireAdmin_();
  asegurarUsuarios_();
  return leerUsuariosAuth_().map(u => ({
    username: u.username,
    rol: String(u.rol || ROLES.BASICO),
    activo: String(u.activo) !== 'false',
    fecha_creacion: u.fecha_creacion || ''
  }));
}

function resetearContrasenaAdmin(username, passwordNueva) {
  requireAdmin_();
  asegurarUsuarios_();
  const user = String(username || '').trim();
  const nueva = String(passwordNueva || '');
  if (!user) throw new Error('Debes indicar el usuario');
  if (nueva.length < 8) throw new Error('La nueva contraseña debe tener mínimo 8 caracteres');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  rows[idx].salt = Utilities.getUuid().replace(/-/g, '');
  rows[idx].password_hash = passwordHash_(nueva, rows[idx].salt);
  escribirUsuariosAuth_(rows);
  return { ok: true, user: user };
}

function cambiarRolUsuarioAdmin(username, rol) {
  requireAdmin_();
  asegurarUsuarios_();
  const actor = currentUser_();
  const user = String(username || '').trim();
  const rolFinal = String(rol || ROLES.BASICO).trim().toLowerCase();
  if (!Object.values(ROLES).includes(rolFinal)) throw new Error('Rol inválido');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  if (rolFinal !== ROLES.ADMIN) {
    const adminsRestantes = rows.filter(u => String(u.rol) === ROLES.ADMIN && String(u.username || '').trim().toLowerCase() !== user.toLowerCase()).length;
    const esMismoActor = String(rows[idx].username || '').trim().toLowerCase() === String(actor || '').trim().toLowerCase();
    if (esMismoActor && adminsRestantes === 0) throw new Error('No puedes quitarte el último admin');
  }
  rows[idx].rol = rolFinal;
  escribirUsuariosAuth_(rows);
  return { ok: true, user: user, rol: rolFinal };
}

function eliminarUsuarioAdmin(username) {
  requireAdmin_();
  asegurarUsuarios_();
  const actor = currentUser_();
  const user = String(username || '').trim();
  if (!user) throw new Error('Debes indicar el usuario');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  if (String(rows[idx].username || '').trim().toLowerCase() === String(actor || '').trim().toLowerCase()) {
    throw new Error('No puedes eliminarte a ti mismo');
  }
  if (String(rows[idx].rol) === ROLES.ADMIN) {
    const otrosAdmins = rows.filter(u => String(u.rol) === ROLES.ADMIN && String(u.username || '').trim().toLowerCase() !== user.toLowerCase()).length;
    if (otrosAdmins === 0) throw new Error('No puedes eliminar al último admin');
  }
  rows.splice(idx, 1);
  escribirUsuariosAuth_(rows);
  // Limpiar vinculaciones del usuario eliminado.
  const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() !== user.toLowerCase());
  escribirHojasUsuarios_(links);
  invalidarTokensDeUsuario_(user);
  return { ok: true, user: user };
}

function invalidarTokensDeUsuario_(username) {
  const u = String(username || '').trim().toLowerCase();
  if (!u) return;
  escribirAuthHojaGenerica_('Tokens',
    leerAuthHojaGenerica_('Tokens').filter(r => String(r.username || '').trim().toLowerCase() !== u)
  );
}

function cambiarMiContrasena(passwordActual, passwordNueva) {
  const actor = requireUsuario_();
  asegurarUsuarios_();
  const actual = String(passwordActual || '');
  const nueva = String(passwordNueva || '');
  if (!actual || !nueva) throw new Error('Debes indicar la contraseña actual y la nueva');
  if (nueva.length < 8) throw new Error('La nueva contraseña debe tener mínimo 8 caracteres');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === String(actor).toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  const user = rows[idx];
  if (String(user.activo) === 'false') throw new Error('Usuario inactivo');
  if (passwordHash_(actual, String(user.salt || '')) !== String(user.password_hash || '')) {
    throw new Error('La contraseña actual no es correcta');
  }
  if (passwordHash_(nueva, String(user.salt || '')) === String(user.password_hash || '')) {
    throw new Error('La nueva contraseña debe ser distinta de la actual');
  }
  rows[idx].salt = Utilities.getUuid().replace(/-/g, '');
  rows[idx].password_hash = passwordHash_(nueva, rows[idx].salt);
  escribirUsuariosAuth_(rows);
  return { ok: true, user: actor };
}

// ───────── Spreadsheets (admin) ─────────

function listarSpreadsheetsAdmin() {
  requireAdmin_();
  const sheets = leerSpreadsheets_();
  const links = leerHojasUsuarios_();
  return sheets.map(s => {
    const vinculados = links.filter(l => String(l.spreadsheet_id) === String(s.spreadsheet_id));
    return {
      spreadsheet_id: String(s.spreadsheet_id),
      nombre: String(s.nombre || ''),
      descripcion: String(s.descripcion || ''),
      fecha_alta: s.fecha_alta || '',
      usuarios: vinculados.map(l => String(l.username))
    };
  });
}

function altaSpreadsheetAdmin(spreadsheetId, nombre, descripcion) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  const nom = String(nombre || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!nom) throw new Error('Indica un nombre');
  if (!/^[a-zA-Z0-9_-]{20,}$/.test(id)) throw new Error('El ID no parece un spreadsheet ID de Google');
  // Validación temprana.
  try { SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('No se puede abrir el spreadsheet: ' + (e && e.message || e)); }
  const filas = leerSpreadsheets_();
  if (filas.some(s => String(s.spreadsheet_id) === id)) throw new Error('Ese spreadsheet ya está registrado');
  filas.push({
    spreadsheet_id: id,
    nombre: nom,
    descripcion: String(descripcion || '').trim(),
    fecha_alta: isoAhora_()
  });
  escribirSpreadsheets_(filas);
  return { ok: true, spreadsheet_id: id, nombre: nom };
}

function bajaSpreadsheetAdmin(spreadsheetId) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  const sheets = leerSpreadsheets_();
  if (!sheets.some(s => String(s.spreadsheet_id) === id)) throw new Error('Spreadsheet no registrado');
  const links = leerHojasUsuarios_().filter(l => String(l.spreadsheet_id) !== id);
  escribirHojasUsuarios_(links);
  escribirSpreadsheets_(sheets.filter(s => String(s.spreadsheet_id) !== id));
  return { ok: true };
}

function renombrarSpreadsheetAdmin(spreadsheetId, nombre) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  const nom = String(nombre || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!nom) throw new Error('Indica un nombre');
  const sheets = leerSpreadsheets_();
  const idx = sheets.findIndex(s => String(s.spreadsheet_id) === id);
  if (idx === -1) throw new Error('Spreadsheet no registrado');
  sheets[idx].nombre = nom;
  escribirSpreadsheets_(sheets);
  return { ok: true, spreadsheet_id: id, nombre: nom };
}

function listarVinculacionesAdmin() {
  requireAdmin_();
  const sheets = leerSpreadsheets_();
  const porId = {};
  sheets.forEach(s => { porId[String(s.spreadsheet_id)] = s; });
  const users = leerUsuariosAuth_().map(u => String(u.username || '').trim()).filter(Boolean);
  return users.map(username => {
    const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() === username.toLowerCase());
    return {
      username: username,
      hojas: links.map(l => ({
        spreadsheet_id: String(l.spreadsheet_id),
        nombre: (porId[String(l.spreadsheet_id)] || {}).nombre || '(sin nombre)',
        por_defecto: String(l.por_defecto) === 'true' || l.por_defecto === true
      }))
    };
  });
}

function vincularHojaUsuarioAdmin(username, spreadsheetId, porDefecto) {
  requireAdmin_();
  const user = String(username || '').trim();
  const id = String(spreadsheetId || '').trim();
  if (!user) throw new Error('Indica el usuario');
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!leerUsuariosAuth_().some(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase())) {
    throw new Error('Usuario no existe');
  }
  if (!leerSpreadsheets_().some(s => String(s.spreadsheet_id) === id)) {
    throw new Error('Spreadsheet no registrado; primero añádelo en el directorio');
  }
  const quiereDefecto = porDefecto === true || porDefecto === 'true';
  const links = leerHojasUsuarios_();
  const existe = links.find(l =>
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (existe) throw new Error('Ese usuario ya tiene vinculada esa hoja');
  if (quiereDefecto) {
    links.forEach(l => {
      if (String(l.username || '').trim().toLowerCase() === user.toLowerCase()) l.por_defecto = false;
    });
  } else if (!links.some(l => String(l.username || '').trim().toLowerCase() === user.toLowerCase())) {
    // Primera hoja del usuario → se marca por defecto automáticamente.
    porDefecto = true;
  }
  links.push({
    username: user,
    spreadsheet_id: id,
    por_defecto: porDefecto === true || porDefecto === 'true',
    fecha_alta: isoAhora_()
  });
  escribirHojasUsuarios_(links);
  // Pre-cargar las hojas en el spreadsheet destino si está vacío.
  try { ensureUserSchema_(id); } catch (e) { /* best-effort */ }
  return { ok: true, username: user, spreadsheet_id: id, por_defecto: porDefecto };
}

function desvincularHojaUsuarioAdmin(username, spreadsheetId) {
  requireAdmin_();
  const user = String(username || '').trim();
  const id = String(spreadsheetId || '').trim();
  if (!user || !id) throw new Error('Faltan parámetros');
  const links = leerHojasUsuarios_();
  const restantes = links.filter(l => !(
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id
  ));
  if (restantes.length === links.length) throw new Error('Vinculación no encontrada');
  const eraDefecto = links.find(l =>
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id &&
    (String(l.por_defecto) === 'true' || l.por_defecto === true)
  );
  escribirHojasUsuarios_(restantes);
  if (eraDefecto) {
    const otra = restantes.find(l => String(l.username || '').trim().toLowerCase() === user.toLowerCase());
    if (otra) { otra.por_defecto = true; escribirHojasUsuarios_(restantes); }
  }
  return { ok: true };
}

function setHojaPorDefectoAdmin(username, spreadsheetId) {
  requireAdmin_();
  const user = String(username || '').trim();
  const id = String(spreadsheetId || '').trim();
  if (!user || !id) throw new Error('Faltan parámetros');
  const links = leerHojasUsuarios_();
  const target = links.find(l =>
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (!target) throw new Error('El usuario no tiene vinculada esa hoja');
  links.forEach(l => {
    if (String(l.username || '').trim().toLowerCase() === user.toLowerCase()) l.por_defecto = false;
  });
  target.por_defecto = true;
  escribirHojasUsuarios_(links);
  return { ok: true };
}

function listarMisHojas() {
  const username = requireUsuario_();
  return listarHojasDelUsuario_(username);
}

function cambiarHojaActiva(spreadsheetId) {
  const username = requireUsuario_();
  const id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('Indica la hoja');
  const links = leerHojasUsuarios_();
  const existe = links.some(l =>
    String(l.username || '').trim().toLowerCase() === username.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (!existe) throw new Error('No tienes vinculada esa hoja');
  links.forEach(l => {
    if (String(l.username || '').trim().toLowerCase() === username.toLowerCase()) l.por_defecto = false;
  });
  const target = links.find(l =>
    String(l.username || '').trim().toLowerCase() === username.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (target) {
    target.por_defecto = true;
    escribirHojasUsuarios_(links);
  }
  _currentSheetId = id;
  // Nueva hoja activa → invalidar cachés de datos de usuario
  _userReadCache = {};
  _headersCache = {};
  _sheetObjCache = {};
  delete _ssCache[id]; // forzar reopen limpio si hace falta
  return { ok: true, hojaActivaId: id };
}

// ───────── Bootstrap ─────────

function bootstrapBase() {
  const owner = requireUsuario_();
  asegurarUsuarios_();
  const hojasUsuario = listarHojasDelUsuario_(owner);
  const hojaActivaId = _currentSheetId || resolverHojaActivaId_(owner);

  if (!hojasUsuario.length || !hojaActivaId) {
    _currentSheetId = '';
    return {
      sesion: { user: owner, rol: currentRol_() || ROLES.BASICO },
      version: APP_VERSION,
      hojas: [],
      hojaActivaId: '',
      tasks: [],
      labels: [],
      entries: [],
      activeTimer: null,
      routines: [],
      routineStatus: { today: [], week: [], month: [] },
      routineWeek: [],
      sin_hojas: true,
      sin_datos: true
    };
  }

  _currentSheetId = hojaActivaId;
  ensureUserSchema_(hojaActivaId);
  // ponytail: el frontend solía disparar getRoutines + getRoutineStatus + getDailyRoutineWeekStatus
  // tras bootstrap — 3 round-trips extra en el cold-start (cada uno puede ser 30s en primer hit).
  // Ahora los devolvemos aquí, en la misma ejecución de Apps Script: 8 readRows_ pero 1 viaje.
  return {
    sesion: { user: owner, rol: currentRol_() || ROLES.BASICO },
    version: APP_VERSION,
    hojas: hojasUsuario,
    hojaActivaId: hojaActivaId,
    tasks: getTasks(),
    labels: getLabels(),
    entries: getEntries(),
    activeTimer: getActiveTimer(),
    routines: readRows_('Routines'),
    routineStatus: getRoutineStatus(),
    routineWeek: getDailyRoutineWeekStatus(7, 0)
  };
}

function bootstrap() {
  return bootstrapBase();
}

// ───────── Configuración inicial ─────────

function ping() {
  if (!getMasterSpreadsheetId_()) {
    return { ok: true, configurado: false };
  }
  return { ok: true, configurado: true, version: APP_VERSION };
}

function configurarSpreadsheetMaestro(id, adminUsername, adminPassword) {
  const limpio = String(id || '').trim();
  if (!/^[A-Za-z0-9_-]{20,}$/.test(limpio)) throw new Error('ID de spreadsheet inválido');
  // Validación temprana: el script debe tener acceso de edición al spreadsheet.
  SpreadsheetApp.openById(limpio);
  setMasterSpreadsheetId_(limpio);
  invalidarAuthCache_('Usuarios');
  invalidarAuthCache_('Config');

  // Seed inicial opcional del admin. Si se pasan username/password, se usa
  // esa pareja; si no, cae al seed por defecto (admin / admin1234) sólo si
  // la hoja Usuarios está vacía.
  const user = String(adminUsername || '').trim();
  const pass = String(adminPassword || '');
  const quiereCustom = !!(user || pass);
  if (quiereCustom) {
    if (!/^[a-zA-Z0-9_.-]{3,40}$/.test(user || DEFAULT_ADMIN_USERNAME)) {
      throw new Error('Usuario admin inválido (3-40, letras, números, _.-)');
    }
    if (pass.length < 8) throw new Error('La contraseña admin debe tener mínimo 8 caracteres');
    const salt = Utilities.getUuid().replace(/-/g, '');
    escribirUsuariosAuth_([{
      username: user || DEFAULT_ADMIN_USERNAME,
      password_hash: passwordHash_(pass, salt),
      salt: salt,
      rol: ROLES.ADMIN,
      activo: true,
      fecha_creacion: isoAhora_()
    }]);
  } else {
    // Disparar el seed por defecto si la hoja Usuarios aún no tiene filas.
    asegurarUsuarios_();
  }
  return { ok: true, configurado: true };
}

// ───────── CRUD per-usuario (Tasks / Labels / TimeEntries) ─────────
// Todas estas funciones leen/escriben en el spreadsheet del usuario activo
// (ssActiva_()), no en el master.

function uid_() { return Utilities.getUuid(); }
function iso_() { return new Date().toISOString(); }

// ponytail: one-shot migration — corre una sola vez por spreadsheet al primer bootstrap
// tras el cambio de esquema RoutineCompletions → 3 hojas por periodo. Agrupa filas por
// (period, periodKey), une los routineIds y los escribe en la hoja nueva correspondiente.
// Tras la migración, la hoja antigua se borra. No-op si ya se migró (oldSheet ausente).
function migrateLegacyRoutineCompletions_(ss) {
  const oldSheet = ss.getSheetByName('RoutineCompletions');
  if (!oldSheet) return;
  const data = oldSheet.getDataRange().getValues();
  // Hoja vacía o sólo header: limpiamos sin escribir nada en las nuevas.
  if (data.length < 2) { ss.deleteSheet(oldSheet); return; }
  const headers = data[0];
  const idCol = headers.indexOf('routineId');
  const periodCol = headers.indexOf('period');
  const keyCol = headers.indexOf('periodKey');
  if (idCol < 0 || periodCol < 0 || keyCol < 0) {
    ss.deleteSheet(oldSheet); return;
  }
  const groups = {};
  for (let i = 1; i < data.length; i++) {
    const period = String(data[i][periodCol] || '').trim();
    const key = String(data[i][keyCol] || '').trim();
    const id = String(data[i][idCol] || '').trim();
    if (!period || !key || !id || !ROUTINE_PERIODS.has(period)) continue;
    const k = period + '|' + key;
    if (!groups[k]) groups[k] = { period: period, key: key, ids: [] };
    groups[k].ids.push(id);
  }
  Object.values(groups).forEach(g => {
    const name = completionSheetName_(g.period);
    const newSheet = ss.getSheetByName(name);
    if (!newSheet) return; // se creará al final del loop de SHEETS en el caller
    const lastRow = newSheet.getLastRow();
    const existing = lastRow > 1 ? newSheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
    let rowIdx = -1;
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0]) === g.key) { rowIdx = i; break; }
    }
    const existingIds = rowIdx >= 0 ? String(existing[rowIdx][1] || '').split(',').filter(Boolean) : [];
    const merged = existingIds.concat(g.ids.filter(x => existingIds.indexOf(x) < 0));
    if (rowIdx >= 0) {
      newSheet.getRange(rowIdx + 2, 1, 1, 2).setValues([[g.key, merged.join(',')]]);
    } else {
      newSheet.appendRow([g.key, merged.join(',')]);
    }
  });
  ss.deleteSheet(oldSheet);
}

function ensureUserSchema_(sheetId) {
  const id = String(sheetId || _currentSheetId || '');
  if (!id) throw new Error('No hay hoja activa');
  const ss = openSsById_(id);
  Object.entries(SHEETS).forEach(([name, headers]) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, headers.length).setValues([headers]);
      s.setFrozenRows(1);
    } else {
      // ponytail: backfill any missing trailing columns (e.g. Tasks.completed) — one-shot per sheet on first bootstrap
      const currentCols = s.getLastColumn();
      if (currentCols < headers.length) {
        const newHeaders = headers.slice(currentCols);
        s.getRange(1, currentCols + 1, 1, newHeaders.length).setValues([newHeaders]);
        const lastRow = s.getLastRow();
        if (lastRow > 1) {
          const blank = [newHeaders.map(() => '')];
          s.getRange(2, currentCols + 1, lastRow - 1, newHeaders.length).setValues(
            Array.from({ length: lastRow - 1 }, () => blank[0])
          );
        }
      }
    }
  });
  migrateLegacyRoutineCompletions_(ss);
}

function ensureSchema_() {
  // Compatibilidad con versiones anteriores que sólo tenían una hoja única.
  // Si no hay _currentSheetId, cae al master (modo single-sheet).
  if (!_currentSheetId) {
    // Sin hoja activa: bootstrap asume master en single-sheet; crear las tablas allí.
    const ss = authSs_();
    Object.entries(SHEETS).forEach(([name, headers]) => {
      let s = ss.getSheetByName(name);
      if (!s) s = ss.insertSheet(name);
      if (s.getLastRow() === 0) {
        s.getRange(1, 1, 1, headers.length).setValues([headers]);
        s.setFrozenRows(1);
      } else {
        // ponytail: backfill trailing columns — mismo motivo que ensureUserSchema_
        const currentCols = s.getLastColumn();
        if (currentCols < headers.length) {
          const newHeaders = headers.slice(currentCols);
          s.getRange(1, currentCols + 1, 1, newHeaders.length).setValues([newHeaders]);
          const lastRow = s.getLastRow();
          if (lastRow > 1) {
            s.getRange(2, currentCols + 1, lastRow - 1, newHeaders.length).setValues(
              Array.from({ length: lastRow - 1 }, () => newHeaders.map(() => ''))
            );
          }
        }
      }
    });
    migrateLegacyRoutineCompletions_(ss);
    return;
  }
  ensureUserSchema_(_currentSheetId);
}

function normalizeRow_(obj) {
  Object.keys(obj).forEach(k => {
    const v = obj[k];
    if (v instanceof Date) obj[k] = v.toISOString();
    else if (v === null || v === undefined) obj[k] = '';
  });
  if ('durationMinutes' in obj) obj.durationMinutes = Number(obj.durationMinutes) || 0;
  return obj;
}

function getHeaders_(name) {
  if (_headersCache[name]) return _headersCache[name];
  const s = getUserSheet_(name);
  if (!s || s.getLastRow() < 1) return (SHEETS[name] || []);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  _headersCache[name] = headers;
  return headers;
}

function readRows_(name) {
  if (_userReadCache[name]) return cloneRows_(_userReadCache[name]);
  const s = getUserSheet_(name);
  if (!s || s.getLastRow() <= 1) {
    _userReadCache[name] = [];
    return [];
  }
  const values = s.getDataRange().getValues();
  const headers = values[0];
  _headersCache[name] = headers;
  const rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return normalizeRow_(obj);
    });
  // belt-and-suspenders: strip non-JSON classes
  const clean = JSON.parse(JSON.stringify(rows));
  _userReadCache[name] = clean;
  return cloneRows_(clean);
}

function appendRow_(name, obj) {
  const s = getUserSheet_(name);
  const headers = getHeaders_(name);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  s.appendRow(row);
  // Mantener caché coherente
  if (_userReadCache[name]) {
    const copy = normalizeRow_(Object.assign({}, obj));
    // Asegurar que las fechas/ISO queden como string
    _userReadCache[name].push(JSON.parse(JSON.stringify(copy)));
  } else {
    invalidateUserDataCache_(name);
  }
}

function findRowNum_(name, id) {
  const s = getUserSheet_(name);
  if (!s || s.getLastRow() <= 1) return -1;
  // Preferir caché en memoria si está caliente
  if (_userReadCache[name]) {
    const idx = _userReadCache[name].findIndex(r => String(r.id) === String(id));
    return idx >= 0 ? idx + 2 : -1; // +2 porque fila 1 = header
  }
  const headers = getHeaders_(name);
  const idCol = headers.indexOf('id');
  if (idCol < 0) return -1;
  const last = s.getLastRow();
  // Leer solo la columna id (mucho más barato que getDataRange completo)
  const ids = s.getRange(2, idCol + 1, last, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function updateRow_(name, id, patch) {
  const s = getUserSheet_(name);
  const rowNum = findRowNum_(name, id);
  if (rowNum === -1) throw new Error(`${name} row not found: ${id}`);
  const headers = getHeaders_(name);
  const row = s.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  Object.keys(patch).forEach(k => {
    const col = headers.indexOf(k);
    if (col >= 0) row[col] = patch[k];
    else throw new Error(`${name} no tiene la columna "${k}". Ejecuta bootstrap para migrar el esquema.`);
  });
  s.getRange(rowNum, 1, 1, headers.length).setValues([row]);
  // Actualizar caché en memoria
  if (_userReadCache[name]) {
    const idx = rowNum - 2;
    if (idx >= 0 && idx < _userReadCache[name].length) {
      Object.assign(_userReadCache[name][idx], patch);
      normalizeRow_(_userReadCache[name][idx]);
    }
  }
}

function deleteRow_(name, id) {
  const s = getUserSheet_(name);
  const rowNum = findRowNum_(name, id);
  if (rowNum > 0) {
    s.deleteRow(rowNum);
    if (_userReadCache[name]) {
      const idx = rowNum - 2;
      if (idx >= 0) _userReadCache[name].splice(idx, 1);
    }
  }
}

// --- Bootstrap data ---
function getInitialData() {
  return {
    tasks:  readRows_('Tasks'),
    labels: readRows_('Labels'),
    entries: readRows_('TimeEntries'),
    activeTimer: getActiveTimer()
  };
}

// --- Tasks ---
function getTasks() { return readRows_('Tasks'); }
function createTask(name, labelId) {
  const t = { id: uid_(), name: String(name).trim(), labelId: labelId || '', createdAt: iso_() };
  appendRow_('Tasks', t);
  return getTasks();
}
function updateTask(id, patch) {
  if ('name' in patch) patch.name = String(patch.name).trim();
  if ('labelId' in patch) patch.labelId = patch.labelId || '';
  updateRow_('Tasks', id, patch);
  return getTasks();
}
function deleteTask(id) {
  deleteRow_('Tasks', id);
  return getTasks();
}

// --- Labels ---
function getLabels() { return readRows_('Labels'); }
function nextColor_() {
  const labels = getLabels();
  return PALETTE[labels.length % PALETTE.length];
}
function createLabel(name) {
  const l = { id: uid_(), name: String(name).trim(), color: nextColor_(), createdAt: iso_() };
  appendRow_('Labels', l);
  return getLabels();
}
function updateLabel(id, patch) {
  if ('name' in patch) patch.name = String(patch.name).trim();
  updateRow_('Labels', id, patch);
  return getLabels();
}
function deleteLabel(id) {
  // Limpiar labelId en tareas que lo usaban (usa caché de request)
    const tasks = getTasks().filter(t => t.labelId === id);
    tasks.forEach(t => updateRow_('Tasks', t.id, { labelId: '' }));
  deleteRow_('Labels', id);
  return getLabels();
}

// --- Time entries ---
function getEntries() { return readRows_('TimeEntries'); }
function createEntry(payload) {
  const start = new Date(payload.startTime);
  const end   = new Date(payload.endTime);
  if (!payload.taskId) throw new Error('taskId required');
  if (isNaN(start) || isNaN(end)) throw new Error('invalid start/end time');
  if (end <= start) throw new Error('end must be after start');
  const durationMinutes = Math.max(1, Math.round((end - start) / 60000));
  appendRow_('TimeEntries', {
    id: uid_(),
    taskId: payload.taskId,
    startTime: start,
    endTime: end,
    durationMinutes,
    source: payload.source || 'manual',
    notes: payload.notes || '',
    createdAt: new Date()
  });
  return getEntries();
}
function deleteEntry(id) {
  deleteRow_('TimeEntries', id);
  return getEntries();
}

// ponytail: one-shot rebaser for entries stored under the buggy ISO parser — run from the editor with your local UTC offset (PDT = +7, UTC = 0)
function rebaseEntries(offsetHours) {
  const s = ssActiva_().getSheetByName('TimeEntries');
  if (!s || s.getLastRow() <= 1) return 'No entries';
  const offsetMs = offsetHours * 3600000;
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const startI = headers.indexOf('startTime');
  const endI   = headers.indexOf('endTime');
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    let changed = false;
    if (data[i][startI] instanceof Date) {
      data[i][startI] = new Date(data[i][startI].getTime() + offsetMs).toISOString();
      changed = true;
    }
    if (data[i][endI] instanceof Date) {
      data[i][endI] = new Date(data[i][endI].getTime() + offsetMs).toISOString();
      changed = true;
    }
    if (changed) count++;
  }
  s.getRange(1, 1, data.length, headers.length).setValues(data);
  return `Rebased ${count} entr${count === 1 ? 'y' : 'ies'} by ${offsetHours}h`;
}

// ponytail: diagnostics — run from the editor to inspect the sheet's raw state
function dumpTimeEntries() {
  const s = ssActiva_().getSheetByName('TimeEntries');
  if (!s) return 'No TimeEntries sheet';
  const data = s.getDataRange().getValues();
  return JSON.stringify({
    sheetName: 'TimeEntries',
    lastRow: s.getLastRow(),
    headerRow: data[0],
    sample: data.slice(1, 4),
    types: data[1] ? data[1].map((v, i) => `[${i}] ${data[0][i]}: ${v === null ? 'null' : typeof v} (${v instanceof Date ? 'Date' : String(v).slice(0, 40)})`) : []
  }, null, 2);
}

function resetTimeEntries() {
  const s = ssActiva_().getSheetByName('TimeEntries');
  if (!s) return 'No TimeEntries sheet';
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  return 'Cleared all entries';
}

// --- Routines (per-user) ---
// ponytail: una rutina es SOLO un nombre + periodo (daily/weekly/monthly). Las completions
// de cada periodo viven en su propia hoja (DailyCompletions / WeeklyCompletions / MonthlyCompletions)
// con shape [periodKey, ids], así no necesitamos regenerar la lista cada día/semana/mes.
// La clave natural del periodo se computa en el cliente (yyyy-mm-dd / yyyy-Www / yyyy-mm) para
// que coincida exactamente con la "fecha local" del usuario.
const ROUTINE_PERIODS = new Set(['daily', 'weekly', 'monthly']);

function completionSheetName_(period) {
  return period === 'daily'  ? 'DailyCompletions'
       : period === 'weekly' ? 'WeeklyCompletions'
       :                       'MonthlyCompletions';
}

function periodKeyLocal_(period, when) {
  const d = when ? new Date(when) : new Date();
  const p = n => String(n).padStart(2, '0');
  if (period === 'daily')  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`;
  if (period === 'monthly') return `${d.getFullYear()}-${p(d.getMonth()+1)}`;
  // weekly: ISO week (YYYY-Www). Coincide con el formato que ya usa Stats.
  const thu = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  thu.setDate(d.getDate() + 4 - ((d.getDay() + 6) % 7));
  const year = thu.getFullYear();
  const jan4 = new Date(year, 0, 4);
  const monW1 = new Date(jan4); monW1.setDate(jan4.getDate() - ((jan4.getDay() + 6) % 7));
  const wk = 1 + Math.round((thu - monW1) / (7 * 86400000));
  return `${year}-W${p(wk)}`;
}

// ponytail: Sheets coacciona strings como "2026-08-25" (daily) o "2026-08" (monthly) a
// celdas de fecha, y al releerlas readRows_ las devuelve como Date -> ISO. Eso rompe la
// comparación de periodKey (la UI nunca marca y toggleRoutineCompletion duplica filas).
// Este helper reduce cualquier valor almacenado al formato canónico del periodo.
function normalizeStoredPeriodKey_(value, period) {
  if (value == null) return '';
  let s = String(value).trim();
  if (!s) return '';
  // Si viene como ISO datetime (celda coaccionada a fecha), quédate con la parte de fecha.
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
  if (iso) s = `${iso[1]}-${iso[2]}-${iso[3]}`;
  if (period === 'daily')   return s.slice(0, 10);   // YYYY-MM-DD
  if (period === 'monthly') return s.slice(0, 7);    // YYYY-MM
  return s;                                          // weekly: YYYY-Www (no se coacciona)
}

function getRoutines() { return readRows_('Routines'); }
function createRoutine(name, period) {
  const p = String(period || '').trim();
  if (!ROUTINE_PERIODS.has(p)) throw new Error('Periodo inválido (daily|weekly|monthly)');
  const n = String(name || '').trim();
  if (!n) throw new Error('Indica un nombre');
  // ponytail: dedupe (name, period) — mismo nombre + mismo periodo = mismo concepto, no se duplica.
  // Si el usuario quiere una variante ("Ejercicio AM" vs "Ejercicio PM") que cambie el nombre.
  const dup = readRows_('Routines').some(r =>
    String(r.period || '').trim() === p &&
    String(r.name || '').trim().toLowerCase() === n.toLowerCase()
  );
  if (dup) throw new Error('Ya existe una tarea recurrente con ese nombre en ese periodo');
  appendRow_('Routines', { id: uid_(), name: n, period: p, createdAt: iso_() });
  return getRoutines();
}
function updateRoutine(id, patch) {
  if ('name' in patch) patch.name = String(patch.name || '').trim();
  if ('period' in patch) {
    const p = String(patch.period || '').trim();
    if (!ROUTINE_PERIODS.has(p)) throw new Error('Periodo inválido');
    patch.period = p;
  }
  updateRow_('Routines', id, patch);
  return getRoutines();
}
function deleteRoutine(id) {
  deleteRow_('Routines', id);
  // Limpieza best-effort de completions huérfanas en las tres hojas por periodo:
  // quita routineId del campo ids de cada hoja y borra filas que quedan vacías.
  const target = String(id);
  ['DailyCompletions', 'WeeklyCompletions', 'MonthlyCompletions'].forEach(name => {
    try {
      const sheet = getUserSheet_(name);
      if (!sheet) return;
      const headers = getHeaders_(name);
      if (!headers || headers.length < 2) return;
      const rows = readRows_(name);
      const cleaned = rows
        .map(r => ({ periodKey: r.periodKey, ids: String(r.ids || '').split(',').filter(x => x && x !== target).join(',') }))
        .filter(r => r.ids);
      if (cleaned.length === 0) {
        if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      } else {
        const matriz = [headers].concat(cleaned.map(r => headers.map(h => r[h] != null ? r[h] : '')));
        sheet.getRange(1, 1, matriz.length, headers.length).setValues(matriz);
        if (sheet.getLastRow() > matriz.length) sheet.deleteRows(matriz.length + 1, sheet.getLastRow() - matriz.length);
      }
      invalidateUserDataCache_(name);
    } catch (e) { /* best-effort */ }
  });
  return getRoutines();
}

function toggleRoutineCompletion(routineId, period, periodKey) {
  // El frontend manda el periodKey concreto (today / this week / this month). Las completions
  // viven en tres hojas (una por periodo) con shape [periodKey, ids]. Si routineId ya está
  // en la lista de ese periodKey, lo quitamos; si no, lo añadimos.
  const routine = readRows_('Routines').find(r => String(r.id) === String(routineId));
  if (!routine) throw new Error('Tarea recurrente no encontrada');
  const p = String(period || '').trim();
  const k = String(periodKey || '').trim();
  if (!ROUTINE_PERIODS.has(p)) throw new Error('Periodo inválido');
  if (!k) throw new Error('PeriodKey requerido');

  const sheetName = completionSheetName_(p);
  const sheet = getUserSheet_(sheetName);
  if (!sheet) throw new Error(sheetName + ' no existe; ejecuta bootstrap primero');

  // ponytail: normalizamos el periodKey de cada fila (Sheets pudo coaccionar "2026-08-25"/
  // "2026-08" a fecha al escribirlos) y colapsamos por periodo. Así el findIndex vuelve a
  // acertar, y si ya había filas duplicadas para el mismo periodo se funden en una sola.
  const byKey = new Map();
  readRows_(sheetName).forEach(r => {
    const key = normalizeStoredPeriodKey_(r.periodKey, p);
    if (!key) return;
    if (!byKey.has(key)) byKey.set(key, new Set());
    const set = byKey.get(key);
    String(r.ids || '').split(',').filter(Boolean).forEach(id => set.add(id));
  });
  const rows = Array.from(byKey.entries()).map(([periodKey, set]) => ({ periodKey, ids: Array.from(set).join(',') }));
  const target = String(routineId);
  const idx = rows.findIndex(r => r.periodKey === k);

  let nextRows;
  if (idx >= 0) {
    const ids = String(rows[idx].ids || '').split(',').filter(Boolean);
    const pos = ids.indexOf(target);
    if (pos >= 0) ids.splice(pos, 1); else ids.push(target);
    nextRows = rows.slice();
    nextRows[idx] = { periodKey: k, ids: ids.join(',') };
  } else {
    nextRows = rows.concat([{ periodKey: k, ids: target }]);
  }

  const headers = getHeaders_(sheetName);
  if (nextRows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  } else {
    const matriz = [headers].concat(nextRows.map(r => headers.map(h => r[h] != null ? r[h] : '')));
    const destino = sheet.getRange(1, 1, matriz.length, headers.length);
    // ponytail: forzamos formato de texto ANTES de escribir para que Sheets NO convierta el
    // periodKey ("2026-08-25", "2026-08") a fecha; si no, al releer vuelve como Date y rompe
    // la comparación (la UI nunca marca y se insertan filas duplicadas para el mismo periodo).
    destino.setNumberFormat('@');
    destino.setValues(matriz);
    // ponytail: setValues sólo rellena hasta matriz.length; las filas sobrantes quedan huérfanas
    // y la siguiente lectura las devuelve, así que el "uncheck" reaparece como checked.
    if (sheet.getLastRow() > matriz.length) sheet.deleteRows(matriz.length + 1, sheet.getLastRow() - matriz.length);
  }
  invalidateUserDataCache_(sheetName);
  // ponytail: el frontend ya hace optimistic update + sólo necesita saber el estado canónico
  // del (routine, periodKey) que acaba de tocar. Devolver getRoutineStatus() costaba 3 readRows_
  // extra por toggle (Routines + WeeklyCompletions + MonthlyCompletions siempre, Daily a veces).
  // Aquí ya leí­mos la fila del día/periodo tocado — sabemos el resultado sin volver a leer.
  return {
    routineId: target,
    period: p,
    periodKey: k,
    completed: nextRows[idx >= 0 ? idx : nextRows.length - 1].ids.split(',').filter(String).indexOf(target) >= 0
  };
}

function getRoutineStatus() {
  // Devuelve para cada periodo (today/this week/this month): todas las rutinas del periodo
  // con un flag completed según routineId esté o no en la lista ids del periodKey actual.
  const routines = readRows_('Routines');
  const keyToday = periodKeyLocal_('daily');
  const keyWeek  = periodKeyLocal_('weekly');
  const keyMonth = periodKeyLocal_('monthly');
  const idsByPeriod = {
    daily:   readRows_('DailyCompletions'),
    weekly:  readRows_('WeeklyCompletions'),
    monthly: readRows_('MonthlyCompletions')
  };
  const doneByPeriod = {};
  Object.keys(idsByPeriod).forEach(p => {
    doneByPeriod[p] = new Set();
    const wanted = p === 'daily' ? keyToday : p === 'weekly' ? keyWeek : keyMonth;
    idsByPeriod[p].forEach(row => {
      if (normalizeStoredPeriodKey_(row.periodKey, p) === wanted) {
        String(row.ids || '').split(',').filter(Boolean).forEach(id => doneByPeriod[p].add(id));
      }
    });
  });
  const build = period => {
    const key = period === 'daily' ? keyToday : period === 'weekly' ? keyWeek : keyMonth;
    return routines
      .filter(r => String(r.period) === period)
      .map(r => ({
        routine: r,
        periodKey: key,
        completed: doneByPeriod[period].has(String(r.id))
      }));
  };
  return { today: build('daily'), week: build('weekly'), month: build('monthly') };
}

// ponytail: panel semanal del dashboard + grid N-días de Recurrentes comparten endpoint.
// weeksBack = cuántas semanas ISO hacia atrás desde HOY (clamp 0..52).
// n = nº de días desde el LUNES de esa semana. SIEMPRE devolvemos la semana completa
// (Mon..Sun), NO una ventana deslizante de "últimos N días" — antes hacía que el frontend
// recibiera un bloque Thu..Wed y mapease Wed→domingo porque asumía Mon=índice 0.
// El frontend usa DOW_LABELS_SHORT con lunes=0, así que necesitamos alinear Mon..Sun.
function getDailyRoutineWeekStatus(daysBack, weeksBack) {
  const routines = readRows_('Routines').filter(r => String(r.period) === 'daily');
  const n = Math.max(1, Math.min(60, Number(daysBack) || 7));
  const wb = Math.max(0, Math.min(52, Number(weeksBack) || 0));
  const ref = new Date();
  ref.setDate(ref.getDate() - wb * 7);
  // lunes ISO de la semana que contiene ref (lunes=0). Mismo cálculo que usa el cliente.
  const monday = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate() - ((ref.getDay() + 6) % 7));
  const dayKeys = [];
  for (let i = 0; i < n; i++) {
    const d = new Date(monday); d.setDate(monday.getDate() + i);
    dayKeys.push(periodKeyLocal_('daily', d));
  }
  const byKey = new Map();
  readRows_('DailyCompletions').forEach(r => {
    const k = normalizeStoredPeriodKey_(r.periodKey, 'daily');
    if (!k) return;
    if (!byKey.has(k)) byKey.set(k, new Set());
    String(r.ids || '').split(',').filter(Boolean).forEach(id => byKey.get(k).add(id));
  });
  return routines.map(r => {
    const target = String(r.id);
    const days = dayKeys.map(k => ({ date: k, completed: byKey.has(k) && byKey.get(k).has(target) }));
    return { id: target, name: r.name, days, count: days.filter(d => d.completed).length };
  });
}

// --- Timer (per-user via ActiveTimer sheet in the user's spreadsheet) ---
// Hoja de una sola fila debajo del header: si existe, hay timer activo.
function getActiveTimer() {
  const s = getUserSheet_('ActiveTimer');
  if (!s || s.getLastRow() < 2) return null;
  const row = s.getRange(2, 1, 1, 2).getValues()[0];
  if (!row[0]) return null;
  return { taskId: String(row[0]), startTime: String(row[1]) };
}

function startTimer(taskId) {
  const t = { taskId, startTime: iso_() };
  const s = getUserSheet_('ActiveTimer');
  if (!s) throw new Error('ActiveTimer sheet missing; bootstrap first');
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  s.getRange(2, 1, 1, 2).setValues([[t.taskId, t.startTime]]);
  return t;
}

function stopTimer() {
  const timer = getActiveTimer();
  if (!timer) return null;
  const s = getUserSheet_('ActiveTimer');
  if (s && s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  createEntry({
    taskId: timer.taskId,
    startTime: timer.startTime,
    endTime: iso_(),
    source: 'timer'
  });
  return getEntries();
}

// ───────── API dispatch ─────────

const API_ACTIONS = new Set([
  'ping', 'authStatus', 'loginUsuario', 'logoutUsuario', 'configurarSpreadsheetMaestro',
  'bootstrap', 'bootstrapBase',
  'listarUsuariosAdmin', 'crearUsuarioAdmin', 'resetearContrasenaAdmin',
  'cambiarRolUsuarioAdmin', 'eliminarUsuarioAdmin',
  'listarSpreadsheetsAdmin', 'altaSpreadsheetAdmin', 'bajaSpreadsheetAdmin',
  'renombrarSpreadsheetAdmin',
  'listarVinculacionesAdmin', 'vincularHojaUsuarioAdmin',
  'desvincularHojaUsuarioAdmin', 'setHojaPorDefectoAdmin',
  'cambiarMiContrasena', 'listarMisHojas', 'cambiarHojaActiva',
  'getInitialData', 'getTasks', 'createTask', 'updateTask', 'deleteTask',
  'getLabels', 'createLabel', 'updateLabel', 'deleteLabel',
  'getEntries', 'createEntry', 'deleteEntry',
  'getActiveTimer', 'startTimer', 'stopTimer',
  'getRoutines', 'createRoutine', 'updateRoutine', 'deleteRoutine',
  'toggleRoutineCompletion', 'getRoutineStatus', 'getDailyRoutineWeekStatus'
]);

const API_PUBLIC_ACTIONS = new Set([
  'ping', 'authStatus', 'loginUsuario', 'logoutUsuario', 'configurarSpreadsheetMaestro'
]);

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatchApi_(request) {
  const action = String(request && request.action || '').trim();
  const args = Array.isArray(request && request.args) ? request.args : [];
  if (!action) throw new Error('Missing action');
  if (!API_ACTIONS.has(action)) throw new Error('Acción no permitida: ' + action);
  const fn = globalThis[action];
  if (typeof fn !== 'function') throw new Error('Función no encontrada: ' + action);
  if (API_PUBLIC_ACTIONS.has(action)) return fn.apply(null, args);
  return _authadmin(request.token, action, ...args);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!params.action) {
    return jsonResponse_({ ok: true, data: { service: 'tasktime', version: APP_VERSION } });
  }
  try {
    return jsonResponse_({ ok: true, data: dispatchApi_({
      action: params.action,
      token: String(params.token || ''),
      args: params.args ? JSON.parse(params.args) : []
    }) });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    return jsonResponse_({ ok: true, data: dispatchApi_(body) });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message || err) });
  }
}
