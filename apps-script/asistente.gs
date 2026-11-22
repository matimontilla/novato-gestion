// ═══════════════════════════════════════════════════════════════════
//  NOVATO · ASISTENTE DE CARGA RÁPIDA
//  Convierte una frase ("vendí 12 Malbec 2022 a Organyca a 9000 cada una") en un
//  BORRADOR de formulario. NO escribe nada en la planilla: la app lo muestra, el
//  usuario lo revisa y recién al confirmar pasa por los endpoints de siempre
//  (con toda su validación). Ver DECISIONES.md ("Carga rápida con IA").
//
//  Script Properties (nunca en el repo):
//    ANTHROPIC_API_KEY   clave de la API de Anthropic. Sin ella la función se apaga sola.
//    CLAUDE_MODEL        opcional; por defecto 'claude-haiku-5-5'.
// ═══════════════════════════════════════════════════════════════════

var CLAUDE_MODEL_DEFAULT = 'claude-haiku-5-5';
var CANALES_APP          = ['Vinoteca', 'Distribuidor', 'Restaurant', 'Directo', 'Exportación', 'Otro'];
var BOTELLAS_POR_CAJA    = 6;     // una caja de vino = 6 botellas (dato del negocio)
var IA_MAX_CHARS         = 500;   // largo máximo de la frase
var IA_LIMITE_USUARIO_HORA = 30;  // frases por usuario por hora
var IA_LIMITE_DIA        = 300;   // frases de todos juntos por día (tope de gasto)

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// Cuenta un uso y devuelve false si se pasó de un límite. Usa CacheService (aproximado, suficiente
// como freno de costo; no es contabilidad).
function iaPermitido_(uid) {
  var cache = CacheService.getScriptCache();
  var ahora = new Date();
  var kHora = 'ia_h_' + uid + '_' + Utilities.formatDate(ahora, 'America/Argentina/Mendoza', 'yyyyMMddHH');
  var kDia  = 'ia_d_' + Utilities.formatDate(ahora, 'America/Argentina/Mendoza', 'yyyyMMdd');
  var h = Number(cache.get(kHora)) || 0, d = Number(cache.get(kDia)) || 0;
  if (h >= IA_LIMITE_USUARIO_HORA) return { ok: false, mensaje: 'Llegaste al límite de ' + IA_LIMITE_USUARIO_HORA + ' frases por hora. Probá más tarde o cargalo a mano.' };
  if (d >= IA_LIMITE_DIA)          return { ok: false, mensaje: 'Se alcanzó el tope diario de la carga rápida. Cargalo a mano por hoy.' };
  cache.put(kHora, String(h + 1), 3600);
  cache.put(kDia,  String(d + 1), 86400);
  return { ok: true };
}

// Lo único que se le cuenta a la IA: nombres (productos, depósitos, cajas, categorías, contactos)
// y la fecha de hoy. Nunca montos, saldos ni precios.
function catalogoParaIa_() {
  var contactos = [], vistos = {};
  getContactosBalance().slice(0, 80).forEach(function(c) {
    var k = norm_(c.nombre); if (k && !vistos[k]) { vistos[k] = 1; contactos.push(c.nombre); }
  });
  getClientes().forEach(function(c) {
    var k = norm_(c.nombre); if (k && !vistos[k]) { vistos[k] = 1; contactos.push(c.nombre); }
  });
  return {
    hoy: Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd'),
    productos: Object.keys(PRODUCTO_APP_TO_SHEET),
    depositos: UBICACIONES.slice(),
    cajas: Object.keys(CAJA_LABELS),
    canales: CANALES_APP,
    categorias: getCategoriasBalance().slice(0, 25).map(function(c) { return c.detalle; }),
    contactos: contactos
  };
}

function promptIa_(cat) {
  return [
    'Sos el asistente de carga de Novato, una bodega boutique de Mendoza (Argentina). Los socios te dictan o escriben en español rioplatense una operación y vos la convertís en un borrador para el formulario, llamando SIEMPRE a la herramienta "borrador".',
    '',
    'Fecha de hoy: ' + cat.hoy + '. Resolvé "ayer", "el lunes", etc. a yyyy-mm-dd contra esa fecha. Si no dicen fecha, dejá fecha vacía.',
    '',
    'Tipos de borrador:',
    '- "venta": se entregan botellas o se registra una operación comercial (vendí, le dejé, entregué, muestra, retiro, compra de insumos). Usa "detalle", "contacto" y "lineas".',
    '- "movimiento_caja": entra o sale plata de una caja sin mover botellas (cobré, me pagó, pagué, gasté). Usa "movimiento" y "contacto"; "detalle" si lo dicen.',
    '- "transferencia": mover botellas entre depósitos (pasé 20 Malbec 2022 de R Peña a Pipi). Usa "transferencia".',
    '- "no_entiendo": si no se puede determinar qué operación es. Explicalo en "avisos".',
    '',
    'Reglas:',
    '- Usá SOLO los nombres de estas listas, escritos exactamente igual. "malbec 22" es "Malbec 2022". Si algo no coincide con ninguna lista, dejá el campo vacío y avisalo en "avisos" (en contacto sí podés poner el nombre tal como lo dijeron: puede ser un cliente nuevo).',
    '- NUNCA inventes datos. Lo que no dijeron queda vacío (depósito, caja, canal, fecha, monto). No completes montos por tu cuenta.',
    '- Precios: si dicen "a 9000 cada una" poné precio_unitario=9000 y no calcules el total. Si dicen un total ("por 108 mil") poné monto_total. Los números en pesos argentinos: "108 mil" = 108000, "9 lucas" = 9000.',
    '- "Pipi" puede ser un depósito (botellas) o la caja de Andrés (plata): decidilo según si habla de botellas o de plata.',
    '- Una operación con varios productos va como varias entradas en "lineas".',
    '- Cajas: si dicen "caja(s)", poné el número en "cajas" (media caja = 0.5) y NO lo conviertas a botellas ni lo sumes a "botellas": el código hace la conversión. "botellas" es sólo para botellas sueltas ("2 cajas y 3 botellas" = cajas 2, botellas 3). Lo mismo en transferencias ("cajas" en vez de "cantidad"). Si dicen un precio por caja ("a 50 mil la caja") poné precio_por_caja y no calcules el total.',
    '',
    'Listas válidas:',
    'Productos: ' + cat.productos.join(' | '),
    'Depósitos: ' + cat.depositos.join(' | '),
    'Cajas: ' + cat.cajas.join(' | '),
    'Canales: ' + cat.canales.join(' | '),
    'Categorías (detalle): ' + cat.categorias.join(' | '),
    'Clientes y proveedores conocidos: ' + cat.contactos.join(' | ')
  ].join('\n');
}

function esquemaIa_(cat) {
  var str = { type: 'string' };
  return {
    type: 'object',
    properties: {
      tipo:     { type: 'string', enum: ['venta', 'movimiento_caja', 'transferencia', 'no_entiendo'] },
      detalle:  { type: 'string', description: 'Categoría de la lista, o vacío' },
      contacto: { type: 'string', description: 'Cliente o proveedor' },
      canal:    { type: 'string', description: 'Un canal de la lista, o vacío' },
      fecha:    { type: 'string', description: 'yyyy-mm-dd o vacío' },
      lineas: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            producto:        { type: 'string' },
            deposito:        { type: 'string' },
            botellas:        { type: 'integer', description: 'Botellas sueltas (sin contar las cajas)' },
            cajas:           { type: 'number', description: 'Cantidad de cajas, tal como la dijeron (0.5 = media caja)' },
            precio_unitario: { type: 'number', description: 'Pesos por botella si lo dijeron' },
            precio_por_caja: { type: 'number', description: 'Pesos por caja si lo dijeron' },
            monto_total:     { type: 'number', description: 'Pesos totales de la línea si lo dijeron' }
          },
          required: ['producto']
        }
      },
      movimiento: {
        type: 'object',
        properties: {
          tipo:       { type: 'string', enum: ['cobro', 'gasto'] },
          monto:      { type: 'number' },
          caja:       { type: 'string' },
          referencia: { type: 'string' }
        }
      },
      transferencia: {
        type: 'object',
        properties: {
          producto: { type: 'string' }, cantidad: { type: 'integer', description: 'Botellas sueltas (sin contar las cajas)' },
          cajas: { type: 'number', description: 'Cantidad de cajas, tal como la dijeron (0.5 = media caja)' },
          desde: { type: 'string' }, hacia: { type: 'string' }
        }
      },
      avisos: { type: 'array', items: str, description: 'Dudas o datos que faltan, en una frase corta cada uno' }
    },
    required: ['tipo']
  };
}

// Lleva un nombre dicho por el usuario a uno de la lista (sin importar mayúsculas ni tildes).
// Devuelve '' si no hay coincidencia exacta.
function aLista_(valor, lista) {
  var k = norm_(valor);
  if (!k) return '';
  for (var i = 0; i < lista.length; i++) { if (norm_(lista[i]) === k) return lista[i]; }
  return '';
}

// Normaliza lo que devolvió el modelo: todo pasa por las listas conocidas y por los mismos
// saneadores que usan las cargas reales. Lo que no cuadra se vacía y se avisa.
function normalizarBorrador_(raw, cat) {
  var avisos = Array.isArray(raw.avisos) ? raw.avisos.map(function(a) { return textoSeguro_(a, 160); }).filter(String).slice(0, 5) : [];
  var tipo = ['venta', 'movimiento_caja', 'transferencia', 'no_entiendo'].indexOf(raw.tipo) !== -1 ? raw.tipo : 'no_entiendo';
  var out = { tipo: tipo === 'movimiento_caja' ? 'movimiento' : tipo, avisos: avisos };

  function listaOAviso(valor, lista, nombre) {
    if (!valor) return '';
    var ok = aLista_(valor, lista);
    if (!ok) avisos.push(nombre + ' "' + textoSeguro_(valor, 40) + '" no existe; elegilo a mano.');
    return ok;
  }

  var contacto = textoSeguro_(raw.contacto, 80);
  var exacto = aLista_(contacto, cat.contactos);
  out.contacto = exacto || contacto;
  out.contactoEsNuevo = !!contacto && !exacto;
  if (out.contactoEsNuevo) avisos.push('"' + contacto + '" no está entre tus clientes o proveedores; se cargaría como nuevo.');

  out.detalle = raw.detalle ? listaOAviso(raw.detalle, cat.categorias, 'La categoría') : '';
  out.canal = raw.canal ? listaOAviso(raw.canal, cat.canales, 'El canal') : '';
  try { out.fecha = raw.fecha ? fechaValida_(raw.fecha) : ''; } catch (e) { out.fecha = ''; avisos.push('No entendí la fecha; se usa la de hoy.'); }

  out.lineas = (Array.isArray(raw.lineas) ? raw.lineas : []).slice(0, 20).map(function(l) {
    var sueltas = Math.round(Number(l.botellas));
    if (!isFinite(sueltas) || sueltas < 0) sueltas = 0;
    var cajas = Number(l.cajas);
    if (!isFinite(cajas) || cajas < 0 || cajas > 100000) cajas = 0;
    var botellas = sueltas + Math.round(cajas * BOTELLAS_POR_CAJA); // la conversión caja -> botellas la hace el código
    if (botellas > 100000) botellas = 0;
    var monto = '';
    var pu = Number(l.precio_unitario), mt = Number(l.monto_total), pc = Number(l.precio_por_caja);
    if (isFinite(mt) && mt > 0 && mt <= 1e10) monto = Math.round(mt);
    else if (isFinite(pu) && pu > 0 && botellas > 0 && pu * botellas <= 1e10) monto = Math.round(pu * botellas); // el cálculo lo hace el código, no el modelo
    else if (isFinite(pc) && pc > 0 && cajas > 0 && pc * cajas <= 1e10) monto = Math.round(pc * cajas);
    return {
      producto: listaOAviso(l.producto, cat.productos, 'El producto'),
      deposito: listaOAviso(l.deposito, cat.depositos, 'El depósito'),
      botellas: botellas,
      monto: monto
    };
  });

  var m = raw.movimiento || {};
  var montoMov = Number(m.monto);
  out.movimiento = {
    tipo: m.tipo === 'gasto' ? 'gasto' : (m.tipo === 'cobro' ? 'cobro' : ''),
    monto: (isFinite(montoMov) && montoMov > 0 && montoMov <= 1e10) ? montoMov : '',
    caja: m.caja ? listaOAviso(m.caja, cat.cajas, 'La caja') : '',
    referencia: textoSeguro_(m.referencia, 20)
  };

  var t = raw.transferencia || {};
  var cant = Math.round(Number(t.cantidad));
  if (!isFinite(cant) || cant < 0) cant = 0;
  var cajasT = Number(t.cajas);
  if (isFinite(cajasT) && cajasT > 0 && cajasT <= 100000) cant += Math.round(cajasT * BOTELLAS_POR_CAJA);
  out.transferencia = {
    producto: t.producto ? listaOAviso(t.producto, cat.productos, 'El producto') : '',
    cantidad: (isFinite(cant) && cant > 0 && cant <= 100000) ? cant : '',
    desde: t.desde ? listaOAviso(t.desde, cat.depositos, 'El depósito de origen') : '',
    hacia: t.hacia ? listaOAviso(t.hacia, cat.depositos, 'El depósito de destino') : ''
  };

  if (out.tipo === 'venta' && !out.lineas.length) { out.tipo = 'no_entiendo'; avisos.push('No encontré qué producto o cantidad se entrega.'); }
  return out;
}

// Acción 'interpretarCarga'. Sólo lectura: devuelve un borrador, no escribe nada.
function interpretarCarga_(p, uid) {
  var apiKey = props_().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, code: 'IA_NO_CONFIGURADA', mensaje: 'La carga rápida todavía no está activada en el servidor.' };

  var texto = String(p.texto == null ? '' : p.texto).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!texto) return { ok: false, code: 'VACIO', mensaje: 'Escribí o dictá qué querés cargar.' };
  if (texto.length > IA_MAX_CHARS) return { ok: false, code: 'LARGO', mensaje: 'Es muy largo (máximo ' + IA_MAX_CHARS + ' caracteres). Dividilo en operaciones más cortas.' };

  var permiso = iaPermitido_(uid || p.user || 'anon');
  if (!permiso.ok) return { ok: false, code: 'LIMITE', mensaje: permiso.mensaje };

  var cat = catalogoParaIa_();
  var cuerpo = {
    model: props_().getProperty('CLAUDE_MODEL') || CLAUDE_MODEL_DEFAULT,
    max_tokens: 800,
    system: promptIa_(cat),
    messages: [{ role: 'user', content: texto }],
    tools: [{ name: 'borrador', description: 'Borrador de la operación a cargar en el formulario.', input_schema: esquemaIa_(cat) }],
    tool_choice: { type: 'tool', name: 'borrador' }
  };

  var resp;
  try {
    resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(cuerpo),
      muteHttpExceptions: true
    });
  } catch (e) {
    return { ok: false, code: 'IA_RED', mensaje: 'No pude conectarme con el asistente. Cargalo a mano o probá de nuevo.' };
  }

  var codigo = resp.getResponseCode();
  if (codigo !== 200) {
    // Se registra el código (no la clave ni el texto) para poder diagnosticar desde Ejecuciones.
    console.error('Asistente: la API respondió ' + codigo);
    var msg = codigo === 401 ? 'La clave de la API no es válida. Avisale a Mati.'
            : codigo === 402 || codigo === 429 ? 'El asistente no está disponible en este momento (límite o saldo). Cargalo a mano.'
            : 'El asistente falló. Cargalo a mano o probá de nuevo.';
    return { ok: false, code: 'IA_ERROR', mensaje: msg };
  }

  var data;
  try { data = JSON.parse(resp.getContentText()); } catch (e) { return { ok: false, code: 'IA_ERROR', mensaje: 'Respuesta ilegible del asistente.' }; }
  var bloque = (data.content || []).filter(function(b) { return b.type === 'tool_use' && b.name === 'borrador'; })[0];
  if (!bloque || !bloque.input) return { ok: false, code: 'IA_ERROR', mensaje: 'El asistente no devolvió un borrador. Cargalo a mano.' };

  return { ok: true, propuesta: normalizarBorrador_(bloque.input, cat) };
}
