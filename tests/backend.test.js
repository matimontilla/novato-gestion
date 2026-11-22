// Pruebas de la lógica del backend (seguridad.gs + partes puras de novatobackend.gs)
// con los servicios de Apps Script simulados. Correr con:  npm test
// No toca el Sheet real: lo que depende de la planilla se prueba a mano (ver DECISIONES.md).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

function cargarBackend() {
  const store = {}, cache = {};
  let ahora = Date.now();
  let lockLibre = true;
  const telegram = [];
  const ctx = {
    console,
    Date: class extends Date { constructor(...a) { a.length ? super(...a) : super(ahora); } static now() { return ahora; } },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in store ? store[k] : null),
      setProperty: (k, v) => { store[k] = String(v); },
    }) },
    CacheService: { getScriptCache: () => ({
      get: k => (k in cache ? cache[k] : null),
      put: (k, v) => { cache[k] = String(v); },
      remove: k => { delete cache[k]; },
    }) },
    LockService: { getScriptLock: () => ({
      tryLock: () => lockLibre, releaseLock: () => {},
    }) },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (_a, s) => [...crypto.createHash('sha256').update(s).digest()].map(b => (b > 127 ? b - 256 : b)),
      computeHmacSha256Signature: (t, k) => [...crypto.createHmac('sha256', k).update(t).digest()],
      base64EncodeWebSafe: bytes => Buffer.from(bytes.map(b => b & 0xff)).toString('base64url'),
      getUuid: () => crypto.randomUUID(),
    },
    ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    SpreadsheetApp: {},
    UrlFetchApp: { fetch: () => ({}) },
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  for (const f of ['seguridad.gs', 'novatobackend.gs']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8'), ctx, { filename: f });
  }
  // Doble del teléfono/Telegram y del stock, para no depender del Sheet.
  vm.runInContext('enviarTelegram = function(m){ __tg.push(m); };', Object.assign(ctx, { __tg: telegram }));
  return {
    ctx, store, cache, telegram,
    run: code => vm.runInContext(code, ctx),
    avanzar: ms => { ahora += ms; },
    setLock: v => { lockLibre = v; },
  };
}

function conAuth() {
  const b = cargarBackend();
  b.run('configurarAutenticacionTest_ = function(){ props_().setProperty("AUTH_SECRET","secreto-de-prueba"); USUARIOS_APP.forEach(function(u){ props_().setProperty("PIN_"+u, hashPin_(u,"1234")); }); }');
  b.run('configurarAutenticacionTest_()');
  return b;
}
const J = o => JSON.stringify(o);

test('login correcto devuelve token que verifica al usuario', () => {
  const b = conAuth();
  const r = b.run('login_({userId:"mati", pin:"1234"})');
  assert.equal(r.ok, true);
  assert.equal(b.run(`verificarToken_(${J(r.token)})`), 'mati');
});

test('PIN incorrecto no da token y bloquea a los 5 intentos', () => {
  const b = conAuth();
  for (let i = 0; i < 5; i++) {
    const r = b.run('login_({userId:"lucas", pin:"0000"})');
    assert.equal(r.ok, false);
    assert.equal(r.code, 'PIN_INCORRECTO');
  }
  const bloqueado = b.run('login_({userId:"lucas", pin:"1234"})'); // aun con el PIN bueno
  assert.equal(bloqueado.code, 'BLOQUEADO');
  assert.ok(b.telegram.some(m => m.includes('bloqueado')));
});

test('token alterado, de otro usuario o vencido no vale', () => {
  const b = conAuth();
  const t = b.run('login_({userId:"mati", pin:"1234"})').token;
  const [u, exp, sig] = t.split('.');
  assert.equal(b.run(`verificarToken_(${J('andres.' + exp + '.' + sig)})`), null);
  assert.equal(b.run(`verificarToken_(${J(u + '.' + (Number(exp) + 999999) + '.' + sig)})`), null);
  assert.equal(b.run('verificarToken_("")'), null);
  b.avanzar(13 * 3600 * 1000);
  assert.equal(b.run(`verificarToken_(${J(t)})`), null);
});

test('con AUTH_REQUIRED, un pedido sin token se rechaza; con token pasa y el usuario sale del token', () => {
  const b = conAuth();
  b.run('props_().setProperty("AUTH_REQUIRED","true")');
  b.run('ejecutarAccion_ = function(a,p){ return {ok:true, user:p.user}; }');
  const sin = b.run('despachar_("getData", {})');
  assert.equal(sin.code, 'AUTH');
  const t = b.run('login_({userId:"santi", pin:"1234"})').token;
  const con = b.run(`despachar_("getData", {token:${J(t)}, user:"Mati"})`); // el cliente miente
  assert.equal(con.user, 'Santi');
});

test('sin AUTH_REQUIRED (transición) acepta pedidos sin token pero solo con nombres de socios', () => {
  const b = conAuth();
  b.run('ejecutarAccion_ = function(a,p){ return {ok:true, user:p.user}; }');
  assert.equal(b.run('despachar_("getData", {user:"Lucas"})').user, 'Lucas');
  assert.equal(b.run('despachar_("getData", {user:"Intruso"})').user, '');
});

test('cambiar PIN exige el actual y el nuevo cuenta desde ahí', () => {
  const b = conAuth();
  assert.equal(b.run('cambiarPin_("mati",{pinActual:"9999",pinNuevo:"4321"})').ok, false);
  assert.equal(b.run('cambiarPin_("mati",{pinActual:"1234",pinNuevo:"12"})').code, 'PIN_FORMATO');
  assert.equal(b.run('cambiarPin_("mati",{pinActual:"1234",pinNuevo:"4321"})').ok, true);
  assert.equal(b.run('login_({userId:"mati",pin:"1234"})').ok, false);
  assert.equal(b.run('login_({userId:"mati",pin:"4321"})').ok, true);
});

test('sin configurar, el login avisa para que la app use el PIN local de antes', () => {
  const b = cargarBackend();
  assert.equal(b.run('login_({userId:"mati",pin:"1234"})').code, 'AUTH_NOT_CONFIGURED');
});

test('escritura: el mismo requestId no ejecuta dos veces y devuelve lo anterior', () => {
  const b = conAuth();
  let n = 0;
  b.ctx.__contar = () => ++n;
  b.run('ejecutarAccion_ = function(a,p){ return {ok:true, referencia:"X-00"+__contar()}; }');
  const a = b.run('despachar_("addMovement", {requestId:"abc-123"})');
  const c = b.run('despachar_("addMovement", {requestId:"abc-123"})');
  assert.equal(n, 1);
  assert.equal(a.referencia, 'X-001');
  assert.equal(c.referencia, 'X-001');
  assert.equal(c.duplicado, true);
  b.run('despachar_("addMovement", {requestId:"otro"})');
  assert.equal(n, 2);
});

test('escritura: con el candado ocupado responde BUSY y no ejecuta nada', () => {
  const b = conAuth();
  b.setLock(false);
  b.run('ejecutarAccion_ = function(){ throw new Error("no debería correr"); }');
  assert.equal(b.run('despachar_("addTransfer", {})').code, 'BUSY');
});

test('un error en la escritura no se cachea (se puede reintentar)', () => {
  const b = conAuth();
  b.run('__i=0; ejecutarAccion_ = function(){ __i++; if(__i===1) return {error:"falló"}; return {ok:true}; }');
  assert.ok(b.run('despachar_("addMovement", {requestId:"r1"})').error);
  assert.equal(b.run('despachar_("addMovement", {requestId:"r1"})').ok, true);
});

test('textoSeguro_ neutraliza fórmulas y caracteres de control', () => {
  const b = cargarBackend();
  assert.equal(b.run('textoSeguro_("=IMPORTXML(\\"http://x\\")")'), 'IMPORTXML("http://x")');
  assert.equal(b.run('textoSeguro_("+54 261")'), '54 261');
  assert.equal(b.run('textoSeguro_("  @cmd ")'), 'cmd');
  assert.equal(b.run('textoSeguro_("Organyca\\n\\t")'), 'Organyca');
  assert.equal(b.run('textoSeguro_("abcdef", 3)'), 'abc');
  assert.equal(b.run('textoSeguro_(null)'), '');
});

test('numero_ / entero_ / fechaValida_ rechazan basura', () => {
  const b = cargarBackend();
  assert.throws(() => b.run('numero_("abc","El monto",0,100)'));
  assert.throws(() => b.run('numero_("Infinity","El monto",0,100)'));
  assert.throws(() => b.run('numero_(-5,"El monto",0,100)'));
  assert.throws(() => b.run('entero_(2.5,"Botellas",0,100)'));
  assert.equal(b.run('entero_("12","Botellas",0,100)'), 12);
  assert.throws(() => b.run('fechaValida_("2026-02-31")'));
  assert.throws(() => b.run('fechaValida_("22/11/2026")'));
  assert.throws(() => b.run('fechaValida_("1990-01-01")'));
  assert.equal(b.run('fechaValida_("2026-11-22")'), '2026-11-22');
});

test('sanearMovimiento_ valida caja, tipo y monto positivo', () => {
  const b = cargarBackend();
  const ok = b.run('sanearMovimiento_({user:"Mati",tipo:"cobro",monto:"1500.5",caja:"Pipi (Andrés)",fecha:"2026-11-22",detalle:"=Venta",contacto:"Organyca",referencia:"ORG-001"})');
  assert.equal(ok.monto, 1500.5);
  assert.equal(ok.detalle, 'Venta');
  assert.throws(() => b.run('sanearMovimiento_({tipo:"cobro",monto:"0",caja:"Mati",fecha:"2026-11-22"})'));
  assert.throws(() => b.run('sanearMovimiento_({tipo:"robo",monto:"10",caja:"Mati",fecha:"2026-11-22"})'));
  assert.throws(() => b.run('sanearMovimiento_({tipo:"cobro",monto:"10",caja:"Cuenta fantasma",fecha:"2026-11-22"})'));
});

test('sanearTransferencia_ rechaza mismo depósito, producto raro y cantidad no entera', () => {
  const b = cargarBackend();
  const base = '{user:"Mati",fecha:"2026-11-22",producto:"Malbec 2022",cantidad:"10",desde:"R Peña",hacia:"Pipi",notas:"x"}';
  assert.equal(b.run(`sanearTransferencia_(${base})`).cantidad, 10);
  assert.throws(() => b.run(`sanearTransferencia_(Object.assign(${base},{hacia:"R Peña"}))`));
  assert.throws(() => b.run(`sanearTransferencia_(Object.assign(${base},{producto:"Vino X"}))`));
  assert.throws(() => b.run(`sanearTransferencia_(Object.assign(${base},{cantidad:"2.5"}))`));
  assert.throws(() => b.run(`sanearTransferencia_(Object.assign(${base},{cantidad:"0"}))`));
});

test('validarStockLineas_ suma las líneas del mismo producto/depósito y corta antes de escribir', () => {
  const b = cargarBackend();
  b.run('getStockUbicacion = function(){ return {"Malbec 2022":{"R Peña":10,"Pipi":3}}; }');
  assert.doesNotThrow(() => b.run('validarStockLineas_([{producto:"Malbec 2022",deposito:"R Peña",botellas:6},{producto:"Malbec 2022",deposito:"Pipi",botellas:3}])'));
  assert.throws(() => b.run('validarStockLineas_([{producto:"Malbec 2022",deposito:"R Peña",botellas:6},{producto:"Malbec 2022",deposito:"R Peña",botellas:6}])'), /Stock insuficiente/);
  assert.doesNotThrow(() => b.run('validarStockLineas_([{producto:"",deposito:"",botellas:0,monto:100}])')); // gasto sin botellas
});

test('repartirProporcional cierra exacto y respeta el signo', () => {
  const b = cargarBackend();
  const r = b.run('repartirProporcional(-1000, [1, 1, 1])');
  assert.equal(r.reduce((a, x) => a + x, 0), -1000);
  assert.ok(r.every(x => x <= 0));
  assert.equal(JSON.stringify(b.run('repartirProporcional(500, [0, 0])')), '[500,0]');
});

test('doPost: lee el JSON en texto plano, serializa objetos y responde JSON', () => {
  const b = conAuth();
  b.run('ejecutarAccion_ = function(a,p){ return {ok:true, a:a, lineas:p.lineas, tipo:typeof p.lineas}; }');
  const e = { postData: { contents: JSON.stringify({ action: 'getData', lineas: [{ producto: 'Malbec 2022', botellas: 3 }], x: 'y' }) } };
  b.ctx.__e = e;
  const out = JSON.parse(b.run('doPost(__e)').s);
  assert.equal(out.ok, true);
  assert.equal(out.tipo, 'string');
  assert.equal(JSON.parse(out.lineas)[0].botellas, 3);
  b.ctx.__e = { postData: { contents: 'esto no es json' } };
  assert.ok(JSON.parse(b.run('doPost(__e)').s).error);
});

test('doGet y doPost: un error inesperado vuelve como {error}, no como excepción', () => {
  const b = conAuth();
  b.run('ejecutarAccion_ = function(){ throw new Error("rompió"); }');
  b.ctx.__e = { parameter: { action: 'getData' } };
  assert.equal(JSON.parse(b.run('doGet(__e)').s).error, 'rompió');
});
