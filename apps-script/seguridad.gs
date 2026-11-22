// ═══════════════════════════════════════════════════════════════════
//  NOVATO · SEGURIDAD Y PROTECCIONES DE ESCRITURA
//  Todo lo que protege al backend: login con PIN validado en el servidor, token
//  firmado, candado contra cargas simultáneas, anti-duplicados y validación de
//  datos. Ver DECISIONES.md ("Seguridad del backend") para el porqué y el rollout.
//
//  Se configura con Script Properties (nunca en el repo):
//    AUTH_SECRET     clave para firmar tokens y salar los PIN (la crea configurarAutenticacion)
//    PIN_<usuario>   hash del PIN de cada socio (lo crea configurarAutenticacion / cambiarPin_)
//    AUTH_REQUIRED   'true' → el backend rechaza todo pedido sin token válido.
//                    Mientras no sea 'true', la app vieja sigue funcionando igual.
// ═══════════════════════════════════════════════════════════════════

var USUARIOS_APP   = ['mati', 'andres', 'lucas', 'santi'];
var USER_NOMBRES   = { mati: 'Mati', andres: 'Andrés', lucas: 'Lucas', santi: 'Santi' };
var TOKEN_HORAS    = 12;
var MAX_INTENTOS   = 5;      // PIN incorrectos antes de bloquear
var BLOQUEO_SEG    = 900;    // 15 minutos
var ACCIONES_ESCRITURA = ['addTransaccion', 'addMovement', 'addTransfer', 'addStockControl'];

function props_() { return PropertiesService.getScriptProperties(); }
function authConfigurada_() { return !!props_().getProperty('AUTH_SECRET'); }
function authRequerida_()   { return props_().getProperty('AUTH_REQUIRED') === 'true'; }

// ── CRIPTO ──────────────────────────────────────────────────────────
function hex_(bytes) {
  return bytes.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function hashPin_(uid, pin) {
  var sal = props_().getProperty('AUTH_SECRET') || '';
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + uid + ':' + pin, Utilities.Charset.UTF_8));
}
function firmar_(texto) {
  var sal = props_().getProperty('AUTH_SECRET') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, sal));
}
// Comparación que no se corta en la primera diferencia (evita filtrar información por tiempo).
function iguales_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var r = 0;
  for (var i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function crearToken_(uid) {
  var exp = Date.now() + TOKEN_HORAS * 3600 * 1000;
  var cuerpo = uid + '.' + exp;
  return cuerpo + '.' + firmar_(cuerpo);
}
// Devuelve el id de usuario si el token es válido y no venció; si no, null.
function verificarToken_(token) {
  if (!token || !authConfigurada_()) return null;
  var partes = String(token).split('.');
  if (partes.length !== 3) return null;
  var cuerpo = partes[0] + '.' + partes[1];
  if (!iguales_(partes[2], firmar_(cuerpo))) return null;
  if (USUARIOS_APP.indexOf(partes[0]) === -1) return null;
  if (!(Number(partes[1]) > Date.now())) return null;
  return partes[0];
}

// ── LOGIN / CAMBIO DE PIN ───────────────────────────────────────────
function intentosFallidos_(uid) {
  return Number(CacheService.getScriptCache().get('fail_' + uid)) || 0;
}
function registrarFallo_(uid) {
  var n = intentosFallidos_(uid) + 1;
  CacheService.getScriptCache().put('fail_' + uid, String(n), BLOQUEO_SEG);
  if (n === MAX_INTENTOS) {
    enviarTelegram('🔒 <b>Acceso bloqueado</b>\n' + MAX_INTENTOS + ' PIN incorrectos seguidos para ' +
                   escapeHtml(USER_NOMBRES[uid] || uid) + '. Se desbloquea solo en 15 minutos.');
  }
  return n;
}

// Comprueba el PIN de un usuario respetando el bloqueo. Devuelve {ok:true} o {ok:false, code, mensaje}.
function comprobarPin_(uid, pin) {
  if (USUARIOS_APP.indexOf(uid) === -1) return { ok: false, code: 'USUARIO', mensaje: 'Usuario desconocido.' };
  if (!authConfigurada_() || !props_().getProperty('PIN_' + uid)) {
    return { ok: false, code: 'AUTH_NOT_CONFIGURED', mensaje: 'La autenticación del servidor todavía no está configurada.' };
  }
  if (intentosFallidos_(uid) >= MAX_INTENTOS) {
    return { ok: false, code: 'BLOQUEADO', mensaje: 'Demasiados intentos. Probá de nuevo en 15 minutos.' };
  }
  if (!/^\d{4}$/.test(String(pin || '')) || !iguales_(hashPin_(uid, pin), props_().getProperty('PIN_' + uid))) {
    var n = registrarFallo_(uid);
    return { ok: false, code: 'PIN_INCORRECTO', mensaje: 'PIN incorrecto.', restantes: Math.max(0, MAX_INTENTOS - n) };
  }
  CacheService.getScriptCache().remove('fail_' + uid);
  return { ok: true };
}

function login_(p) {
  var uid = String(p.userId || '');
  var chk = comprobarPin_(uid, p.pin);
  if (!chk.ok) return chk;
  return { ok: true, token: crearToken_(uid), user: { id: uid, nombre: USER_NOMBRES[uid] } };
}

function cambiarPin_(uid, p) {
  if (!uid) return { ok: false, code: 'AUTH', mensaje: 'Sesión no válida.' };
  var nuevo = String(p.pinNuevo || '');
  if (!/^\d{4}$/.test(nuevo)) return { ok: false, code: 'PIN_FORMATO', mensaje: 'El PIN nuevo debe tener 4 dígitos.' };
  var chk = comprobarPin_(uid, p.pinActual);
  if (!chk.ok) return chk;
  props_().setProperty('PIN_' + uid, hashPin_(uid, nuevo));
  return { ok: true };
}

// ── DESPACHO: autenticación, candado y anti-duplicados ──────────────
// Punto único de entrada de doGet y doPost.
function despachar_(action, p) {
  p = p || {};
  if (action === 'login') return login_(p);

  var uid = verificarToken_(p.token);
  if (authRequerida_() && !uid) return { error: 'Tu sesión venció. Ingresá de nuevo.', code: 'AUTH' };

  // El usuario sale del token, no de lo que diga el cliente. Sin token (modo de
  // transición) sólo se aceptan los nombres de los socios.
  p.user = uid ? USER_NOMBRES[uid] : usuarioConocido_(p.user);

  if (action === 'changePin') return cambiarPin_(uid, p);
  if (ACCIONES_ESCRITURA.indexOf(action) !== -1) return escrituraProtegida_(action, p);
  return ejecutarAccion_(action, p);
}

function usuarioConocido_(nombre) {
  for (var k in USER_NOMBRES) { if (USER_NOMBRES[k] === nombre) return nombre; }
  return '';
}

// Una escritura a la vez (candado de script) y sin repetir lo ya hecho (requestId).
// Dos socios cargando juntos ya no pueden recibir la misma referencia ni pisarse el stock,
// y un reintento por mala señal devuelve el resultado anterior en vez de duplicar la carga.
function escrituraProtegida_(action, p) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    return { error: 'Hay otra carga en curso. Probá de nuevo en unos segundos.', code: 'BUSY' };
  }
  try {
    var rid = p.requestId ? String(p.requestId).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 64) : '';
    var cache = CacheService.getScriptCache();
    if (rid) {
      var previo = cache.get('req_' + rid);
      if (previo) { var r = JSON.parse(previo); r.duplicado = true; return r; }
    }
    var res = ejecutarAccion_(action, p);
    if (rid && res && !res.error) cache.put('req_' + rid, JSON.stringify(res), 21600);
    return res;
  } finally {
    lock.releaseLock();
  }
}

// ── VALIDACIÓN DE DATOS ─────────────────────────────────────────────
// Texto libre: sin caracteres de control, con largo máximo, y sin símbolos al inicio
// que Sheets interpretaría como fórmula (=, +, -, @). Un contacto "=IMPORTXML(...)"
// escrito con setValue se ejecutaría en la planilla; con esto queda como texto inofensivo.
function textoSeguro_(v, max) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
  s = s.replace(/^[=+\-@\s]+/, '');
  if (max && s.length > max) s = s.slice(0, max);
  return s;
}

function numero_(v, nombre, min, max) {
  var n = (v === '' || v == null) ? 0 : Number(v);
  if (!isFinite(n)) throw new Error('Valor inválido en ' + nombre + '.');
  if (n < min || n > max) throw new Error(nombre + ' fuera de rango (' + min + ' a ' + max + ').');
  return n;
}
function entero_(v, nombre, min, max) {
  var n = numero_(v, nombre, min, max);
  if (Math.floor(n) !== n) throw new Error(nombre + ' debe ser un número entero.');
  return n;
}

function fechaValida_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Fecha inválida.');
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) throw new Error('Fecha inválida.');
  var hoy = new Date();
  var min = new Date(hoy.getFullYear() - 8, 0, 1);
  var max = new Date(hoy.getTime() + 31 * 86400000);
  if (d < min || d > max) throw new Error('Fecha fuera de rango.');
  return String(s);
}

function productoValido_(v) {
  v = String(v || '');
  if (!v) return '';
  if (PRODUCTO_APP_TO_SHEET[v]) return v;
  for (var k in PRODUCTO_APP_TO_SHEET) { if (PRODUCTO_APP_TO_SHEET[k] === v) return k; }
  throw new Error('Producto desconocido: ' + textoSeguro_(v, 40));
}
function depositoValido_(v) {
  v = String(v || '');
  if (UBICACIONES.indexOf(v) === -1) throw new Error('Depósito desconocido: ' + textoSeguro_(v, 40));
  return v;
}
function cajaValida_(v) {
  v = String(v || '');
  if (!CAJA_LABELS[v]) throw new Error('Caja desconocida: ' + textoSeguro_(v, 40));
  return v;
}

// Cada función devuelve una copia saneada de los parámetros, con los mismos nombres
// de campo que ya usan addMovement / addTransfer / addStockControl.
function sanearMovimiento_(p) {
  var tipo = p.tipo === 'gasto' ? 'gasto' : (p.tipo === 'cobro' ? 'cobro' : '');
  if (!tipo) throw new Error('Tipo de movimiento inválido.');
  return {
    user: p.user, tipo: tipo,
    monto: numero_(p.monto, 'El monto', 0.01, 1e10),
    caja: cajaValida_(p.caja),
    fecha: fechaValida_(p.fecha),
    detalle: textoSeguro_(p.detalle, 60),
    contacto: textoSeguro_(p.contacto, 80),
    referencia: textoSeguro_(p.referencia, 20)
  };
}

function sanearTransferencia_(p) {
  var desde = depositoValido_(p.desde), hacia = depositoValido_(p.hacia);
  if (desde === hacia) throw new Error('Origen y destino no pueden ser el mismo depósito.');
  return {
    user: p.user,
    fecha: fechaValida_(p.fecha),
    producto: productoValido_(p.producto),
    cantidad: entero_(p.cantidad, 'La cantidad', 1, 100000),
    desde: desde, hacia: hacia,
    notas: textoSeguro_(p.notas, 200)
  };
}
