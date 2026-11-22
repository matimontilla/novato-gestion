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
      // Simula formatDate para los patrones que usa el backend (zona horaria incluida).
      formatDate: (d, tz, fmt) => {
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', hourCycle: 'h23', minute: '2-digit' }).formatToParts(d).map(x => [x.type, x.value]));
        return fmt.replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day).replace('HH', parts.hour).replace('mm', parts.minute);
      },
    },
    ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    SpreadsheetApp: {},
    UrlFetchApp: { fetch: () => ({}) },
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  for (const f of ['seguridad.gs', 'novatobackend.gs', 'asistente.gs']) {
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

// ── Carga rápida (asistente.gs) ──────────────────────────────────────
function conIa(respuestas) {
  const b = conAuth();
  b.telegram.length = 0;
  b.run('props_().setProperty("ANTHROPIC_API_KEY","sk-ant-PRUEBA")');
  b.ctx.__fetchs = [];
  b.ctx.__resps = respuestas;
  b.ctx.UrlFetchApp = { fetch: (url, opt) => { b.ctx.__fetchs.push({ url, opt }); const r = b.ctx.__resps.shift(); return { getResponseCode: () => r.code, getContentText: () => JSON.stringify(r.body) }; } };
  b.run(`
    getContactosBalance = function(){ return [{nombre:"Organyca",ultimaFecha:2},{nombre:"Rosario",ultimaFecha:1}]; };
    getClientes = function(){ return [{nombre:"Particular"}]; };
    getCategoriasBalance = function(){ return [{detalle:"Venta",tipo:"Ingreso",total:9},{detalle:"Muestra",tipo:"Neutro",total:2}]; };
  `);
  return b;
}
const respTool = input => ({ code: 200, body: { content: [{ type: 'tool_use', name: 'borrador', input }] } });
const tokenDe = (b, uid = 'mati') => b.run(`login_({userId:${J(uid)}, pin:"1234"})`).token;
const asistente = (b, texto, extra = {}) => { const t = tokenDe(b); return b.run(`despachar_("interpretarCarga", ${J({ token: t, texto, ...extra })})`); };

test('IA: sin clave de API la función se apaga con un mensaje claro', () => {
  const b = conAuth();
  const r = asistente(b, 'vendí 3 malbec');
  assert.equal(r.ok, false);
  assert.equal(r.code, 'IA_NO_CONFIGURADA');
});

test('IA: sin sesión del servidor no se puede usar (aunque AUTH_REQUIRED esté apagado)', () => {
  const b = conIa([respTool({ tipo: 'venta' })]);
  const r = b.run('despachar_("interpretarCarga", {texto:"vendí 3 malbec"})');
  assert.equal(r.code, 'IA_SIN_SESION');
  assert.equal(b.ctx.__fetchs.length, 0); // ni siquiera llamó a la API
});

test('IA: una venta con precio unitario: el total lo calcula el código y los nombres se llevan a la lista', () => {
  const b = conIa([respTool({
    tipo: 'venta', detalle: 'venta', contacto: 'organyca', canal: 'distribuidor', fecha: '2026-11-21',
    lineas: [{ producto: 'malbec 2022', deposito: 'r peña', botellas: 12, precio_unitario: 9000 }],
    avisos: [],
  })]);
  const r = asistente(b, 'Vendí 12 Malbec 2022 a Organyca a 9000 cada una');
  assert.equal(r.ok, true);
  const p = r.propuesta;
  assert.equal(p.tipo, 'venta');
  assert.equal(p.contacto, 'Organyca');
  assert.equal(p.contactoEsNuevo, false);
  assert.equal(p.detalle, 'Venta');
  assert.equal(p.canal, 'Distribuidor');
  assert.equal(p.fecha, '2026-11-21');
  assert.equal(p.lineas[0].producto, 'Malbec 2022');
  assert.equal(p.lineas[0].deposito, 'R Peña');
  assert.equal(p.lineas[0].monto, 108000);
});

test('IA: lo inventado o fuera de lista se vacía y se avisa; un cliente nuevo se marca', () => {
  const b = conIa([respTool({
    tipo: 'venta', contacto: 'Vinoteca Nueva', canal: 'Marte', fecha: '2026-13-45',
    lineas: [{ producto: 'Syrah 2019', deposito: 'Luna', botellas: 5 }],
  })]);
  const p = asistente(b, 'dejé 5 syrah en la vinoteca nueva').propuesta;
  assert.equal(p.contactoEsNuevo, true);
  assert.equal(p.canal, '');
  assert.equal(p.fecha, '');
  assert.equal(p.lineas[0].producto, '');
  assert.equal(p.lineas[0].deposito, '');
  assert.equal(p.lineas[0].monto, '');
  assert.ok(p.avisos.length >= 4);
});

test('IA: movimiento de caja y transferencia se normalizan; montos inválidos quedan vacíos', () => {
  const b = conIa([
    respTool({ tipo: 'movimiento_caja', contacto: 'Rosario', movimiento: { tipo: 'cobro', monto: 50000, caja: 'pipi (andrés)' } }),
    respTool({ tipo: 'transferencia', transferencia: { producto: 'blend 2023', cantidad: 20.4, desde: 'R Peña', hacia: 'pipi' } }),
    respTool({ tipo: 'movimiento_caja', movimiento: { tipo: 'gasto', monto: -5, caja: 'Mati' } }),
  ]);
  const m = asistente(b, 'me pagó rosario 50 mil a la caja de pipi').propuesta;
  assert.equal(m.tipo, 'movimiento');
  assert.equal(m.movimiento.tipo, 'cobro');
  assert.equal(m.movimiento.monto, 50000);
  assert.equal(m.movimiento.caja, 'Pipi (Andrés)');
  const t = asistente(b, 'pasé 20 blend de r peña a pipi').propuesta;
  assert.equal(t.tipo, 'transferencia');
  assert.equal(t.transferencia.cantidad, 20);
  assert.equal(t.transferencia.hacia, 'Pipi');
  const neg = asistente(b, 'gasté algo').propuesta;
  assert.equal(neg.movimiento.monto, '');
});

test('IA: las cajas se convierten a botellas en el código (1 caja = 6 botellas), también media caja y mezcla con sueltas', () => {
  const b = conIa([
    respTool({ tipo: 'venta', contacto: 'Organyca', lineas: [
      { producto: 'Malbec 2022', cajas: 2, precio_por_caja: 50000 },
      { producto: 'Blend 2023', cajas: 0.5, botellas: 2, precio_unitario: 9000 },
    ] }),
    respTool({ tipo: 'transferencia', transferencia: { producto: 'Malbec 2022', cajas: 3, desde: 'R Peña', hacia: 'Pipi' } }),
    respTool({ tipo: 'transferencia', transferencia: { producto: 'Malbec 2022', cantidad: 4, cajas: 1, desde: 'R Peña', hacia: 'Pipi' } }),
  ]);
  const v = asistente(b, 'vendí 2 cajas de malbec 2022 a 50 mil y media caja y 2 botellas de blend').propuesta;
  assert.equal(v.lineas[0].botellas, 12);
  assert.equal(v.lineas[0].monto, 100000);   // 2 cajas x 50.000
  assert.equal(v.lineas[1].botellas, 5);     // media caja (3) + 2 sueltas
  assert.equal(v.lineas[1].monto, 45000);    // 5 botellas x 9.000
  assert.equal(asistente(b, 'pasé 3 cajas de r peña a pipi').propuesta.transferencia.cantidad, 18);
  assert.equal(asistente(b, 'pasé 1 caja y 4 botellas').propuesta.transferencia.cantidad, 10);
});

test('controles de stock: el último de cada depósito queda separado y el global es el más reciente', () => {
  const b = cargarBackend();
  b.run('formatDate = function(d){ return "F"+d; }; formatHora = function(h){ return "H"+h; };');
  const d = n => `new Date(${n})`;
  const r = b.run(`agruparControles_([
    ["a","a","U","R Peña","Malbec 2022",100,98,-2,${d(1000)}],
    ["a","a","U","R Peña","Blend 2023",50,50,0,${d(1000)}],
    ["b","b","U","Pipi","Malbec 2022",10,10,0,${d(2000)}],
    ["c","c","U","R Peña","Malbec 2022",95,96,1,${d(3000)}],
    ["x","x","U","Pipi","Malbec 2022",1,1,0,"sin fecha"]
  ])`);
  assert.equal(r.ultimo.deposito, 'R Peña');
  assert.equal(r.ultimo.items.length, 1);
  assert.equal(r.ultimo.items[0].real, 96);
  assert.equal(r.porDeposito['R Peña'].items[0].real, 96);   // el viejo de R Peña (1000) quedó atrás
  assert.equal(r.porDeposito['Pipi'].items.length, 1);       // la fila sin timestamp se ignora
  assert.deepEqual(Object.keys(r.porDeposito).sort(), ['Pipi', 'R Peña']);
});

test('IA: una venta sin líneas pasa a no_entiendo', () => {
  const b = conIa([respTool({ tipo: 'venta', contacto: 'Organyca', lineas: [] })]);
  const p = asistente(b, 'vendí algo a organyca').propuesta;
  assert.equal(p.tipo, 'no_entiendo');
});

test('IA: se manda el catálogo de nombres pero ningún monto ni saldo, y la clave va solo en el header', () => {
  const b = conIa([respTool({ tipo: 'no_entiendo' })]);
  asistente(b, 'hola');
  const { url, opt } = b.ctx.__fetchs[0];
  assert.equal(url, 'https://api.anthropic.com/v1/messages');
  assert.equal(opt.headers['x-api-key'], 'sk-ant-PRUEBA');
  const cuerpo = JSON.parse(opt.payload);
  assert.equal(cuerpo.model, 'claude-haiku-5-5');
  assert.equal(cuerpo.tool_choice.name, 'borrador');
  assert.ok(cuerpo.system.includes('Organyca') && cuerpo.system.includes('Malbec 2022') && cuerpo.system.includes('Pipi (Andrés)'));
  assert.ok(!opt.payload.includes('sk-ant-PRUEBA'));
  assert.ok(!/saldo|precio ref/i.test(cuerpo.system));
});

test('IA: errores de la API se traducen sin filtrar la clave', () => {
  const b = conIa([{ code: 401, body: { error: { message: 'invalid x-api-key sk-ant-PRUEBA' } } }, { code: 429, body: {} }, { code: 200, body: { content: [] } }]);
  const a = asistente(b, 'uno');
  assert.equal(a.code, 'IA_ERROR');
  assert.ok(!JSON.stringify(a).includes('sk-ant'));
  assert.match(asistente(b, 'dos').mensaje, /no está disponible/);
  assert.equal(asistente(b, 'tres').code, 'IA_ERROR');
});

test('IA: frase vacía o demasiado larga se rechaza sin llamar a la API', () => {
  const b = conIa([]);
  assert.equal(asistente(b, '   ').code, 'VACIO');
  assert.equal(asistente(b, 'x'.repeat(501)).code, 'LARGO');
  assert.equal(b.ctx.__fetchs.length, 0);
});

test('IA: límite por usuario por hora', () => {
  const b = conIa(Array.from({ length: 31 }, () => respTool({ tipo: 'no_entiendo' })));
  for (let i = 0; i < 30; i++) assert.equal(asistente(b, 'hola ' + i).ok, true);
  const r = asistente(b, 'una más');
  assert.equal(r.code, 'LIMITE');
  assert.equal(b.ctx.__fetchs.length, 30);
});

test('analíticas: Yuniku (elaboración por encargo) no entra en ventas por cliente/mes/producto pero sí en el margen de su añada', () => {
  const b = cargarBackend();
  const fila = (fecha, detalle, sub, prod, anada, ars, usd, bot, concepto) => {
    const r = new Array(19).fill('');
    r[1] = fecha; r[2] = detalle; r[3] = sub; r[4] = prod; r[5] = anada; r[6] = ars; r[7] = usd; r[10] = 0; r[12] = bot; r[13] = concepto;
    return r;
  };
  const data = [
    fila(new b.ctx.Date(2022, 7, 4), 'Venta', 'Yuniku', 'Chardonnay 2022', '2022', 1000, 10, 0, 'Ingreso'),
    fila(new b.ctx.Date(2023, 4, 17), 'Venta', 'Carlos De Aquin', 'Chardonnay 2022', '2022', 400, 4, 60, 'Ingreso'),
    fila(new b.ctx.Date(2022, 5, 29), 'Insumos', 'Tapones', 'Chardonnay 2022', '2022', -300, -3, '', 'Egreso'),
  ];
  const hoja = { getLastRow: () => data.length + 2, getRange: () => ({ getValues: () => data }) };
  b.ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: () => hoja }) };
  const r = JSON.parse(JSON.stringify(b.run('getAnalytics()')));
  assert.deepEqual(r.ventasCliente.map(c => c.cliente), ['Carlos De Aquin']);
  assert.equal(r.ventasMes.reduce((s, m) => s + m.ars, 0), 400);
  assert.equal(r.ventasProducto[0].ars, 400);
  assert.equal(r.porEncargo.length, 1);
  assert.equal(r.porEncargo[0].cliente, 'Yuniku');
  assert.equal(r.porEncargo[0].ars, 1000);
  const m = r.margenAnada.find(x => x.anada === '2022');
  assert.equal(m.ventasArs, 1400); // el encargo sí cuenta contra los costos de su añada
  assert.equal(m.botellasVendidas, 60);
  assert.equal(r.flujoAnual.reduce((s, a) => s + a.ingresosArs, 0), 1400);
});
