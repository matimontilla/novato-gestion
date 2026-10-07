// ═══════════════════════════════════════════════════════════════════
//  NOVATO BODEGA · Google Apps Script Backend (lo que usa la app)
//  NO editar en el editor de Apps Script: se despliega solo desde GitHub
//  (apps-script/ en el repo) con la Action deploy-apps-script.yml.
//  Herramientas de mantenimiento para correr a mano: ver herramientas.gs.
// ═══════════════════════════════════════════════════════════════════

// Rango "ancho" para que las fórmulas de reconciliación (BALANCE ↔ CAJA)
// sigan funcionando a medida que se agreguen filas nuevas con el tiempo.
// (Ya no hace falta un límite de fila fijo — ver nota en escribirFormulasBalance)

// Historial de transferencias entre depósitos (NO es la fuente de verdad
// del stock — eso vive directo en STOCK. Esto es sólo un log de auditoría).
var TAB_TRANSF  = 'APP_TRANSFERENCIAS';
var TAB_CONTROL = 'APP_CONTROL_STOCK'; // historial de controles físicos de stock (compartido entre dispositivos)

var UBICACIONES = ['R Peña','Pipi','Lucas','Santi','Mati'];

// Traducción entre las etiquetas cortas que usa la app y los nombres reales
// usados en las pestañas del Sheet.
var PRODUCTO_APP_TO_SHEET = {
  'Malbec 2021':      'Malbec 2021',
  'Malbec 2022':      'Malbec 2022',
  'Blend 2023':       'Blend 2023',
  'Cab. Franc 2021':  'Cabernet Franc 2021',
  'Cab. Franc 2022':  'Cabernet Franc 2022',
  'Chardonnay 2022':  'Chardonnay 2022'
};
var UBIC_SHEET_MAP = { 'R Peña':'R PEÑA', 'Pipi':'PIPI', 'Lucas':'LUCAS', 'Santi':'REZMA', 'Mati':'MATI' };
var CAJA_LABELS = {
  'Empresa (Ludico)': 'LUDICO',
  'Mati':             'MATI',
  'Lucas':            'LUCAS',
  'Pipi (Andrés)':    'PIPI',
  'Santi':            'SANTI'
};
var CLIENTE_PREFIJO = {
  'Angela San Rafael':  'ASR',
  'Bahia Blanca':       'BHB',
  'Carlos De Aquín':    'CDA',
  'Adriana Laos':       'ADL',
  'Chacho Andia':       'CHA',
  'Mosto Divino':       'MOD',
  'Organyca':           'ORG',
  'Particular':         'VPA',
  'Rosario':            'ROS',
  'Santiago MDQ':       'SMD'
};

// ── ROUTER PRINCIPAL ────────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action || 'getData';
  var result;
  try {
    if      (action === 'getData')       result = getData();
    else if (action === 'addTransaccion') result = addTransaccion(e.parameter);
    else if (action === 'addMovement')   result = addMovement(e.parameter);
    else if (action === 'addTransfer')   result = addTransfer(e.parameter);
    else if (action === 'addStockControl') result = addStockControl(e.parameter);
    else if (action === 'getOps')        result = { ops: getRecentOps(20) };
    else if (action === 'getMovimientosCaja') result = { movimientos: getMovimientosCaja(100) };
    else if (action === 'getDetalleOperacion') result = getDetalleOperacion(e.parameter.referencia);
    else if (action === 'getInsumos')    result = getInsumos();
    else if (action === 'getAnalytics')  result = getAnalytics();
    else                                 result = { error: 'Acción desconocida: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET DATA (carga inicial de la app) ──────────────────────────────
function getData() {
  return {
    ok:                 true,
    lastPrice:          getLastSalePrice(),
    ops:                getRecentOps(20),
    stockUbicacion:     getStockUbicacion(),
    ventasPendientes:   getVentasPendientes(),
    comprasPendientes:  getComprasPendientes(),
    operacionesPendientes: getOperacionesPendientes(),
    clientes:           getClientes(),
    ultimoControlStock: getUltimoControlStock(),
    resumenCajas:       getResumenCajas(),
    categorias:         getCategoriasBalance(),
    contactosBalance:   getContactosBalance()
  };
}

// Categorías reales usadas en BALANCE (DETALLE: Venta, Retiros, Muestra, Tapones,
// Flete, etc.) con el signo dominante que tuvieron históricamente (Ingreso/Egreso/
// Neutro si mayormente vienen en blanco, como Retiros y Muestra que mueven botellas
// sin plata de por medio) — así el formulario no tiene que preguntar el signo a mano.
function getCategoriasBalance() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return [];
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 3, lastRow - 2, 5).getValues(); // C..G: DETALLE,SUBDETALLE,PRODUCTO,AÑADA,MONTO $
  var stats = {};
  for (var i = 0; i < data.length; i++) {
    var det = data[i][0];
    if (!det || det === 'TOTALES:') continue;
    var m = data[i][4];
    if (!stats[det]) stats[det] = { pos: 0, neg: 0, total: 0 };
    stats[det].total++;
    if (typeof m === 'number') { if (m > 0) stats[det].pos++; else if (m < 0) stats[det].neg++; }
  }
  var lista = [];
  for (var det2 in stats) {
    var s = stats[det2];
    var tipo = (s.pos === 0 && s.neg === 0) ? 'Neutro' : (s.neg > s.pos ? 'Egreso' : 'Ingreso');
    lista.push({ detalle: det2, tipo: tipo, total: s.total });
  }
  lista.sort(function(a, b) { return b.total - a.total; });
  return lista;
}

// Contactos reales (SUBDETALLE) vistos en BALANCE, con el/los DETALLE con que se
// usó cada uno y su última fecha de uso. El frontend usa esto para: (a) filtrar el
// selector de contacto según el detalle elegido, y (b) ordenar por uso más reciente.
// Devuelve [{nombre, detalles:[...], ultimaFecha:ms}], del más reciente al más viejo.
function getContactosBalance() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return [];
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 2, lastRow - 2, 3).getValues(); // B..D: FECHA, DETALLE, SUBDETALLE
  var mapa = {}; // nombre -> {nombre, detalles:{det:ultimaFechaConEseDetalle}, ultimaFecha}
  for (var i = 0; i < data.length; i++) {
    var fecha = data[i][0], detalle = data[i][1], nombre = data[i][2];
    if (!nombre) continue;
    var ms = (fecha instanceof Date) ? fecha.getTime() : 0;
    if (!mapa[nombre]) mapa[nombre] = { nombre: nombre, detalles: {}, ultimaFecha: 0 };
    var reg = mapa[nombre];
    if (ms > reg.ultimaFecha) reg.ultimaFecha = ms;
    if (detalle) {
      if (!reg.detalles[detalle] || ms > reg.detalles[detalle]) reg.detalles[detalle] = ms;
    }
  }
  var lista = [];
  for (var k in mapa) {
    var r = mapa[k];
    // detalles como lista de {detalle, ultima} para poder filtrar y ordenar en el front
    var dets = [];
    for (var d in r.detalles) dets.push({ detalle: d, ultima: r.detalles[d] });
    lista.push({ nombre: r.nombre, detalles: dets, ultimaFecha: r.ultimaFecha });
  }
  lista.sort(function(a, b) { return b.ultimaFecha - a.ultimaFecha; }); // más reciente primero
  return lista;
}

// Lee el cuadro "CAJAS" (AR$ / USD / CRYPTO por caja) que ya existe al pie de la
// tabla grande en la pestaña CAJA. Busca el bloque por texto ("CAJAS" ... "TOTAL:")
// en vez de por número de fila fijo, para no romperse cuando se insertan filas
// nuevas más arriba (adición de ventas/movimientos corre este bloque hacia abajo).
function getResumenCajas() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  if (!caja) return [];
  var lastRow = caja.getLastRow();
  if (lastRow < 1) return [];
  var data = caja.getRange(1, 2, lastRow, 3).getValues(); // B..D: caja, AR$, USD

  var inicio = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'CAJAS') { inicio = i; break; }
  }
  if (inicio === -1) return [];

  var resultado = [];
  for (var j = inicio + 1; j < data.length; j++) {
    var nombre = data[j][0];
    if (!nombre) continue;
    if (nombre === 'TOTAL:') break;
    if (nombre === 'AR$' || nombre === 'USD') continue; // fila de sub-encabezados
    resultado.push({
      caja: nombre,
      ars:  Number(data[j][1]) || 0,
      usd:  Number(data[j][2]) || 0
    });
  }
  return resultado;
}

// Lista de clientes real (con canal y si está activo/inactivo), directo de la pestaña
// CLIENTES — reemplaza la lista fija que tenía la app, que ya estaba desactualizada
// (le faltaban Yuniku, Gahvino, Nadia, Diego Carino, y varios más).
function getClientes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CLIENTES');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var data = sheet.getRange(3, 2, lastRow - 2, 4).getValues(); // B..E: CLIENTE, CANAL, ULTIMA OP, ESTADO
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var nombre = data[i][0];
    if (!nombre) continue;
    out.push({ nombre: nombre, canal: data[i][1] || '', activo: data[i][3] === 'Activo' });
  }
  return out;
}

// ── ÚLTIMO PRECIO DE VENTA (desde BALANCE) ──────────────────────────
function getLastSalePrice() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var balance = ss.getSheetByName('BALANCE');
  if (!balance) return null;
  var data    = balance.getDataRange().getValues();
  var headers = data[1];
  var iDetalle  = headers.indexOf('DETALLE');
  var iBotellas = headers.indexOf('BOTELLAS');
  var iMonto    = headers.indexOf('MONTO $');
  for (var i = data.length - 1; i >= 2; i--) {
    var row = data[i];
    if (row[iDetalle] === 'Venta' && row[iBotellas] > 0 && row[iMonto] > 0) {
      return Math.round(row[iMonto] / row[iBotellas]);
    }
  }
  return null;
}

// ── FECHA / COTIZACIÓN ───────────────────────────────────────────────
function parseFechaApp(s) {
  var partes = String(s).split('-');
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
}

// getLastRow() no sirve para saber dónde termina la data real: tanto BALANCE como
// CAJA tienen bloques de totales unas filas más abajo (con SUMIF/SUM propios, y algún
// espacio suelto) que hacen que getLastRow() devuelva una fila mucho más lejana que la
// última operación real. Esta función busca la última fila con una FECHA real (un
// objeto Date, no un espacio suelto ni una celda de fórmula) en la columna indicada.
function obtenerUltimaFilaConFecha(sheet, colFecha) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 2;
  var valores = sheet.getRange(3, colFecha, lastRow - 2, 1).getValues();
  for (var i = valores.length - 1; i >= 0; i--) {
    if (valores[i][0] instanceof Date) return 3 + i;
  }
  return 2;
}

// Busca la cotización disponible más cercana (igual o anterior) a una fecha en BLUE_API.
// BLUE_API ya se actualiza sola (confirmado funcionando), así que a diferencia de la
// versión anterior de este script, acá NO hace falta salir a buscar la cotización a una
// API externa como respaldo — alcanza con leer la pestaña.
function getDolarRate(fechaStr) {
  var blue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BLUE_API');
  if (!blue) return null;
  var lastRow = blue.getLastRow();
  if (lastRow < 2) return null;
  var data = blue.getRange(2, 1, lastRow - 1, 3).getValues(); // A:C → day, value_sell, value_buy
  var target = parseFechaApp(fechaStr).getTime();
  var best = null, bestDiff = Infinity;
  for (var i = 0; i < data.length; i++) {
    var d = data[i][0], v = data[i][2]; // C = value_buy, igual que usa la propia fórmula de la hoja
    if (!(d instanceof Date) || v === '' || v === null) continue;
    var diff = target - d.getTime();
    if (diff >= 0 && diff < bestDiff) { bestDiff = diff; best = v; }
  }
  return best;
}

// ── VENTAS → BALANCE ──────────────────────────────────────────────────
function prefijoCliente(nombre) {
  if (CLIENTE_PREFIJO[nombre]) return CLIENTE_PREFIJO[nombre];
  var limpio = String(nombre).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z]/g, '').toUpperCase();
  return (limpio + 'XXX').substring(0, 3);
}

// Busca el producto de la fila de BALANCE que tenga esta REFERENCIA — se usa para
// que un cobro/gasto vinculado herede el producto de la operación original.
// Todas las líneas de BALANCE que comparten una REFERENCIA (una operación puede
// tener varios productos = varias filas). Se usa tanto para heredar el producto en
// CAJA como para repartir proporcionalmente un cobro/gasto entre varias líneas.
function buscarLineasPorReferencia(referencia) {
  if (!referencia) return [];
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 5, lastRow - 2, 8).getValues(); // E..L: PRODUCTO,AÑADA,MONTO$,...,REFERENCIA
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][7] === referencia) out.push({ producto: data[i][0] || '', monto: Math.abs(Number(data[i][2]) || 0) });
  }
  return out;
}

// Reparte un monto total proporcionalmente según una lista de "pesos" (ej. el monto
// de cada línea de la venta original), redondeando al peso y ajustando la ÚLTIMA
// parte para que la suma cierre exacto (sin dejar residuos de centavos sueltos).
function repartirProporcional(total, pesos) {
  var sumaPesos = pesos.reduce(function(a, b) { return a + Math.abs(b); }, 0);
  var signoTotal = total < 0 ? -1 : 1;
  var totalAbs = Math.round(Math.abs(total));

  if (!sumaPesos) {
    // Sin base para prorratear (ej. todas las líneas en 0): todo va a la primera parte
    var partesIguales = pesos.map(function() { return 0; });
    if (partesIguales.length) partesIguales[0] = totalAbs * signoTotal;
    return partesIguales;
  }

  var partes = pesos.map(function(w) { return Math.round(totalAbs * (Math.abs(w) / sumaPesos)); });
  var sumaPartes = partes.reduce(function(a, b) { return a + b; }, 0);
  partes[partes.length - 1] += (totalAbs - sumaPartes); // ajuste de redondeo en la última línea
  return partes.map(function(p) { return p * signoTotal; });
}

function nextReferencia(prefix) {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var lastRow = balance.getLastRow();
  var max = 0;
  if (lastRow >= 3) {
    var refs = balance.getRange(3, 12, lastRow - 2, 1).getValues(); // L = REFERENCIA
    refs.forEach(function(r) {
      var v = String(r[0] || '');
      if (v.indexOf(prefix + '-') === 0) {
        var n = parseInt(v.split('-')[1], 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  return prefix + '-' + ('00' + (max + 1)).slice(-3);
}

// Escribe (o reescribe) las columnas calculadas de una fila de BALANCE: AÑADA, MONTO
// US$ FF/FP, CU $/US$, CONCEPTO, SALDO $/US$, % DIF POR TC y AÑO. Se usa tanto para
// filas nuevas (addTransaccion) como para reparar filas rotas (repararBalance). IMPORTANTE:
// este Sheet usa configuración regional en español → los argumentos de función van
// separados por PUNTO Y COMA (;), no coma. Escribir con comas produce #ERROR!.
// El saldo (O/P) se calcula por fila con SUMIFS filtrando por referencia Y por
// producto. Esto funciona igual para operaciones de un solo producto (donde el filtro
// de producto no cambia nada) y para multi-producto (donde cada línea resta sólo los
// cobros de SU producto, porque addMovement reparte el cobro por producto en CAJA).
// Así cada fila muestra su propio saldo, saldo US$ y rentabilidad por TC, sin doble
// conteo. El parámetro incluirSaldo se mantiene sólo para las filas que legítimamente
// no deben tener saldo (costos prorrateados sin cobro asociado, tipo CI/CO).
function escribirFormulasBalance(row, incluirSaldo) {
  if (incluirSaldo === undefined) incluirSaldo = true;
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  balance.getRange(row, 6).setValue('=RIGHT(E' + row + ';4)'); // F AÑADA
  balance.getRange(row, 8, 1, 4).setValues([[
    '=IFERROR(IF(G' + row + '=0;"";G' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A;BLUE_API!$C$2:$C;;-1));"")', // H MONTO US$ FF
    incluirSaldo ? '=IF(G' + row + '=0;"";H' + row + '+P' + row + ')' : '',                                       // I MONTO US$ FP
    '=IF(OR(G' + row + '=0;E' + row + '="");"";IF(G' + row + '>0;IF(M' + row + '=0;"";G' + row + '/M' + row + ');G' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)))', // J CU $ (venta sin botellas cargadas, ej. operación contable sin entrega física: vacío)
    '=IF(OR(G' + row + '=0;E' + row + '="");"";IF(G' + row + '>0;IF(M' + row + '=0;"";H' + row + '/M' + row + ');H' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)))'  // K CU US$
  ]]);
  balance.getRange(row, 14).setValue('=IF(G' + row + '>0;"Ingreso";(IF(G' + row + '=0;"Movimiento";"Egreso")))'); // N CONCEPTO
  if (incluirSaldo) {
    balance.getRange(row, 15, 1, 3).setValues([[
      '=IF(G' + row + '=0;"";G' + row + '-SUMIFS(CAJA!$F$3:$F;CAJA!$I$3:$I;L' + row + ';CAJA!$E$3:$E;E' + row + '))',   // O SALDO $ (por referencia + producto)
      '=IF(G' + row + '=0;"";-(H' + row + '-SUMIFS(CAJA!$G$3:$G;CAJA!$I$3:$I;L' + row + ';CAJA!$E$3:$E;E' + row + ')))', // P SALDO US$ (por referencia + producto)
      '=IF(G' + row + '=0;"";IF(N' + row + '="Egreso";-P' + row + '/H' + row + ';P' + row + '/H' + row + '))'          // Q % DIF POR TC
    ]]);
  } else {
    balance.getRange(row, 15, 1, 3).setValues([['', '', '']]);
  }
  var cAnio = colBalance_('AÑO'), cDias = colBalance_('DIAS CIERRE');
  if (cAnio) balance.getRange(row, cAnio).setValue('=IF(BALANCE!$B' + row + '="";"";YEAR(BALANCE!$B' + row + '))'); // AÑO
  if (cDias) balance.getRange(row, cDias).setFormula(formulaDiasCierre(row));                                       // DIAS CIERRE
}

// Número de columna de BALANCE según su encabezado (fila 2), o 0 si no existe.
// Se usa para las columnas agregadas al final (AÑO, DEPOSITO, DIAS CIERRE): así se
// pueden reordenar en la planilla sin romper el código. Las columnas A..Q siguen
// usando letras fijas porque las fórmulas las referencian por letra.
var _colsBalance = null;
function colBalance_(nombre) {
  if (!_colsBalance) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
    var hdr = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
    _colsBalance = {};
    for (var i = 0; i < hdr.length; i++) {
      var h = String(hdr[i] || '').trim().toUpperCase();
      if (h && !_colsBalance[h]) _colsBalance[h] = i + 1;
    }
  }
  return _colsBalance[String(nombre).toUpperCase()] || 0;
}

// Días que tardó en cerrarse una operación: desde su fecha (B) hasta el ÚLTIMO cobro o
// pago registrado en CAJA bajo la misma referencia. Se compara sólo el DÍA (INT), sin
// la hora, para no obtener medios días.
//
// Qué cobros cuentan:
//  · Los de la misma referencia Y producto (en multi-producto cada línea mide lo suyo).
//  · Si no hay ninguno con ese producto (cobro cargado con otro producto, o fila sin
//    producto), los de la referencia sola. Antes devolvía 0 y daba -45.000.
// Queda VACÍA si: no hay monto, la fila no lleva saldo, la operación sigue abierta,
// no hay ningún movimiento en CAJA con esa referencia, o es CI/CO (agregados de muchos
// pagos con una fecha promedio: los "días de cierre" no tienen sentido).
// Un valor NEGATIVO es real: el pago se registró antes que la operación (anticipo/seña).
function formulaDiasCierre(row) {
  var r = String(row);
  var ref = 'CAJA!$I$3:$I;L' + r, prod = 'CAJA!$E$3:$E;E' + r;
  return '=IFERROR(IF(OR(G' + r + '=0;O' + r + '="";C' + r + '="CI";C' + r + '="CO");"";' +
           'IF(ROUND(O' + r + ';0)<>0;"";' +
           'IF(COUNTIF(' + ref + ')=0;"";' +
           'INT(IF(AND(E' + r + '<>"";COUNTIFS(' + ref + ';' + prod + ')>0);' +
                 'MAXIFS(CAJA!$B$3:$B;' + ref + ';' + prod + ');' +
                 'MAXIFS(CAJA!$B$3:$B;' + ref + ')))' +   // cierra MAXIFS, IF, INT
           '-INT(B' + r + '))));"")';                  // resta la fecha; cierra IF x3 e IFERROR
}

// Registra la venta en BALANCE (con las mismas fórmulas que usa cualquier fila
// cargada a mano) y descuenta el depósito de origen en STOCK. NO toca CAJA:
// eso sólo pasa cuando se registre el cobro correspondiente.
// ── NOTIFICACIONES POR TELEGRAM ──────────────────────────────────────
// Configuración en Project Settings → Script Properties (NO acá en el código,
// para no dejar el token del bot guardado en el repo de GitHub):
//   TELEGRAM_BOT_TOKEN   → el token que te da BotFather
//   TELEGRAM_CHAT_IDS    → chat_id de cada persona que quiera recibir avisos,
//                          separados por coma (ej: "111111111,222222222")
// Si estas propiedades no están configuradas, enviarTelegram() no hace nada —
// así que es seguro dejarlo desplegado aunque todavía no se haya armado el bot.
function enviarTelegram(mensaje) {
  var props      = PropertiesService.getScriptProperties();
  var token      = props.getProperty('TELEGRAM_BOT_TOKEN');
  var chatIdsStr = props.getProperty('TELEGRAM_CHAT_IDS');
  if (!token || !chatIdsStr) return;

  var chatIds = chatIdsStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  chatIds.forEach(function(chatId) {
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: chatId, text: mensaje, parse_mode: 'HTML' }),
        muteHttpExceptions: true // si Telegram falla, no debe romper la operación real
      });
    } catch (e) {
      // best-effort: un fallo de notificación nunca debe tumbar una venta/cobro/etc.
    }
  });
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// $1.234.567 con separador de miles estilo argentino, sin depender de locale
function formatearMonto(n) {
  var neg = n < 0;
  var entero = String(Math.round(Math.abs(n)));
  var out = '';
  for (var i = 0; i < entero.length; i++) {
    if (i > 0 && (entero.length - i) % 3 === 0) out += '.';
    out += entero[i];
  }
  return (neg ? '-' : '') + out;
}

// Registra una transacción general en BALANCE (Venta, Retiros, Muestra, Ajuste,
// o cualquier categoría real de costo como Tapones/Flete/Elaboracion). Puede tener
// una o varias líneas de producto (p.lineas, JSON: [{producto,deposito,botellas,monto}]),
// todas bajo la MISMA referencia. El signo del monto se determina solo según el
// historial de esa categoría (getCategoriasBalance): Ingreso→positivo, Egreso→negativo,
// Neutro→tal cual se tipeó. Con más de una línea, el saldo de cada fila individual
// queda en blanco — se maneja agregado por referencia (getVentasPendientes/
// getComprasPendientes), no fila por fila, igual que los costos prorrateados.
function addTransaccion(p) {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) throw new Error('No se encontró la pestaña BALANCE');

  var detalle  = p.detalle || 'Venta';
  var contacto = p.contacto || '';
  var lineas   = JSON.parse(p.lineas || '[]');
  if (!lineas.length) throw new Error('No hay líneas de producto para registrar');

  var categorias = getCategoriasBalance();
  var cat  = categorias.filter(function(c) { return c.detalle === detalle; })[0];
  var tipo = cat ? cat.tipo : 'Ingreso'; // categoría nunca vista antes → asumimos Ingreso

  var referencia = nextReferencia(prefijoCliente(contacto || detalle));
  var multiLinea = lineas.length > 1;
  var resumenTelegram = [];

  lineas.forEach(function(linea) {
    var productoSheet = linea.producto ? (PRODUCTO_APP_TO_SHEET[linea.producto] || linea.producto) : '';
    var botellas      = parseInt(linea.botellas) || 0;
    var montoTipeado  = parseFloat(linea.monto) || 0;
    var montoArs;
    if (tipo === 'Egreso') montoArs = -Math.abs(montoTipeado);
    else if (tipo === 'Ingreso') montoArs = Math.abs(montoTipeado);
    else montoArs = montoTipeado; // Neutro: tal cual (Retiros/Muestra suelen quedar en 0)

    var ultimaReal = obtenerUltimaFilaConFecha(balance, 2); // B = FECHA
    balance.insertRowAfter(ultimaReal);
    var row = ultimaReal + 1;

    balance.getRange(row, 2).setValue(parseFechaApp(p.fecha)); // B FECHA
    balance.getRange(row, 3).setValue(detalle);                // C DETALLE
    balance.getRange(row, 4).setValue(contacto);                // D SUBDETALLE
    balance.getRange(row, 5).setValue(productoSheet);           // E PRODUCTO
    balance.getRange(row, 7).setValue(montoArs);                // G MONTO $
    balance.getRange(row, 12).setValue(referencia);             // L REFERENCIA (compartida entre líneas)
    balance.getRange(row, 13).setValue(botellas);               // M BOTELLAS
    balance.getRange(row, 1).setValue(p.user || '');            // A: quién lo cargó
    // S DEPOSITO — de qué depósito salieron las botellas. Sólo tiene sentido si hubo
    // movimiento físico de stock; si no hay botellas, queda vacío.
    var cDep = colBalance_('DEPOSITO');
    if (botellas > 0 && cDep) balance.getRange(row, cDep).setValue(linea.deposito || 'R Peña');

    escribirFormulasBalance(row, true); // saldo por fila (SUMIFS filtra por producto, sin doble conteo aunque sea multi-producto)

    if (productoSheet && botellas > 0) {
      ajustarStockUbicacion(linea.producto, linea.deposito || 'R Peña', null, botellas);
    }

    resumenTelegram.push(
      (productoSheet ? productoSheet : '') +
      (botellas ? ' · ' + botellas + ' bot' : '') +
      (montoArs ? ' · $' + formatearMonto(montoArs) : '')
    );
  });

  var montoTotal = lineas.reduce(function(s, l) { return s + (parseFloat(l.monto) || 0); }, 0);
  enviarTelegram(
    '🍾 <b>' + escapeHtml(detalle) + '</b>' + (multiLinea ? ' (' + lineas.length + ' productos)' : '') + '\n' +
    (contacto ? 'Cliente/Proveedor: ' + escapeHtml(contacto) + '\n' : '') +
    resumenTelegram.map(function(l) { return '· ' + escapeHtml(l); }).join('\n') + '\n' +
    (montoTotal ? 'Total: $' + formatearMonto(montoTotal) + '\n' : '') +
    'Cargado por: ' + escapeHtml(p.user || '-')
  );

  return { ok: true, referencia: referencia };
}

// Ventas con saldo pendiente de cobro — para que la pantalla de Caja pueda
// vincular un cobro a la venta correspondiente y la reconciliación se cierre sola.
// Suma todo lo pagado en CAJA por cada REFERENCIA (una sola pasada sobre CAJA) —
// funciona igual sin importar si esa referencia tiene 1 fila en CAJA o varias
// (operación multi-producto repartida proporcionalmente).
function getTotalPagadoPorReferencia() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  var out = {};
  if (!caja) return out;
  var lastRow = caja.getLastRow();
  if (lastRow < 3) return out;
  var data = caja.getRange(3, 6, lastRow - 2, 4).getValues(); // F..I: MONTO $, MONTO US$, CAJA, REFERENCIA
  for (var i = 0; i < data.length; i++) {
    var ref = data[i][3]; // I REFERENCIA
    if (!ref) continue;
    if (!out[ref]) out[ref] = { ars: 0, usd: 0 };
    out[ref].ars += Number(data[i][0]) || 0; // F MONTO $
    out[ref].usd += Number(data[i][1]) || 0; // G MONTO US$
  }
  return out;
}

// Ventas con saldo pendiente de cobro — agrupadas por REFERENCIA (una operación
// puede tener varios productos = varias filas de BALANCE compartiendo el código).
// El saldo es el TOTAL de la operación menos lo ya cobrado en CAJA bajo esa referencia.
function getVentasPendientes() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return [];
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 2, lastRow - 2, 15).getValues(); // B..P

  var grupos = {}; // referencia -> {cliente, producto:[], fecha, montoArs, montoUsd}
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (r[1] !== 'Venta' || !r[10]) continue; // C DETALLE, L REFERENCIA
    var ref = r[10];
    if (!grupos[ref]) grupos[ref] = { cliente: r[2], producto: [], fecha: r[0], montoArs: 0, montoUsd: 0 };
    grupos[ref].montoArs += Number(r[5]) || 0; // G MONTO $
    grupos[ref].montoUsd += Number(r[6]) || 0; // H MONTO US$ FF
    if (r[3]) grupos[ref].producto.push(r[3]); // E PRODUCTO
  }

  var pagos = getTotalPagadoPorReferencia();
  var out = [];
  for (var ref2 in grupos) {
    var g = grupos[ref2];
    var pagado = pagos[ref2] || { ars: 0, usd: 0 };
    var saldoArs = Math.round(g.montoArs - pagado.ars);
    if (!saldoArs) continue; // ya saldado (o residuo de redondeo)
    out.push({
      referencia: ref2,
      cliente:    g.cliente,
      producto:   g.producto.join(', '),
      fecha:      g.fecha,
      saldoArs:   saldoArs,
      saldoUsd:   Math.round(g.montoUsd - pagado.usd)
    });
  }
  return out;
}

// Compras/costos (Tapones, Flete, Elaboracion, Uva, etc.) con saldo pendiente de
// pago — mismo mecanismo que getVentasPendientes (agrupado por referencia), pero
// del lado "Egreso" en vez de "Ingreso".
function getComprasPendientes() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return [];
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 2, lastRow - 2, 15).getValues(); // B..P

  var grupos = {}; // referencia -> {detalle, proveedor, producto:[], fecha, montoArs, montoUsd}
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (r[12] !== 'Egreso' || !r[10]) continue; // N CONCEPTO, L REFERENCIA
    var ref = r[10];
    if (!grupos[ref]) grupos[ref] = { detalle: r[1], proveedor: r[2] || '', producto: [], fecha: r[0], montoArs: 0, montoUsd: 0 };
    grupos[ref].montoArs += Number(r[5]) || 0; // G MONTO $
    grupos[ref].montoUsd += Number(r[6]) || 0; // H MONTO US$ FF
    if (r[3]) grupos[ref].producto.push(r[3]); // E PRODUCTO
  }

  var pagos = getTotalPagadoPorReferencia();
  var out = [];
  for (var ref2 in grupos) {
    var g = grupos[ref2];
    var pagado = pagos[ref2] || { ars: 0, usd: 0 };
    var saldoArs = Math.round(g.montoArs - pagado.ars);
    if (!saldoArs) continue;
    out.push({
      referencia: ref2,
      detalle:    g.detalle,
      proveedor:  g.proveedor,
      producto:   g.producto.join(', '),
      fecha:      g.fecha,
      saldoArs:   saldoArs,
      saldoUsd:   Math.round(g.montoUsd - pagado.usd)
    });
  }
  return out;
}

// Ventas por cobrar + compras por pagar en una sola lista, para el dashboard.
// Ordenadas de más antigua a más nueva (lo más viejo pendiente primero).
function getOperacionesPendientes() {
  var ventas   = getVentasPendientes().map(function(v) {
    return { tipo:'venta', referencia:v.referencia, contraparte:v.cliente, detalle:'Venta', producto:v.producto, fecha:v.fecha, saldoArs:v.saldoArs, saldoUsd:v.saldoUsd };
  });
  var compras  = getComprasPendientes().map(function(c) {
    return { tipo:'compra', referencia:c.referencia, contraparte:c.proveedor, detalle:c.detalle, producto:c.producto, fecha:c.fecha, saldoArs:c.saldoArs, saldoUsd:c.saldoUsd };
  });
  var todas = ventas.concat(compras);
  todas.sort(function(a, b) {
    var ta = a.fecha instanceof Date ? a.fecha.getTime() : 0;
    var tb = b.fecha instanceof Date ? b.fecha.getTime() : 0;
    return ta - tb;
  });
  return todas;
}

// ── ANALYTICS PARA LA PESTAÑA DATOS ──────────────────────────────────
// Todas las agregaciones para los gráficos, en UNA sola lectura de BALANCE.
// Columnas usadas: B fecha, C detalle, D subdetalle, E producto, F añada, G monto$,
// H montoUS$, K CU US$, M botellas, N concepto.
function getAnalytics() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return {};
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return {};
  var data = balance.getRange(3, 1, lastRow - 2, 19).getValues();

  var ventasMes = {}, ventasCliente = {}, flujoAnio = {}, anadaVentas = {}, anadaCostos = {},
      costoProducto = {}, insumosAnada = {}, ventasProducto = {};

  for (var i = 0; i < data.length; i++) {
    var r        = data[i];
    var fecha    = r[1];
    var detalle  = r[2] || '';
    var contacto = r[3] || '';
    var producto = r[4] || '';
    var anada    = String(r[5] || '').trim();
    var montoArs = Number(r[6]) || 0;
    var montoUsd = Number(r[7]) || 0;
    var cuUsd    = Number(r[10]) || 0;
    var botellas = Number(r[12]) || 0;
    var concepto = r[13] || '';
    if (!(fecha instanceof Date)) continue;
    var anio = fecha.getFullYear();

    // Flujo por año calendario (liquidez): ingresos vs egresos
    if (!flujoAnio[anio]) flujoAnio[anio] = { anio: anio, ingresosArs: 0, egresosArs: 0, ingresosUsd: 0, egresosUsd: 0 };
    if (montoArs > 0) { flujoAnio[anio].ingresosArs += montoArs; flujoAnio[anio].ingresosUsd += montoUsd; }
    else if (montoArs < 0) { flujoAnio[anio].egresosArs += -montoArs; flujoAnio[anio].egresosUsd += -montoUsd; }

    if (detalle === 'Venta') {
      // Ventas por mes
      var mes = anio + '-' + ('0' + (fecha.getMonth() + 1)).slice(-2);
      if (!ventasMes[mes]) ventasMes[mes] = { periodo: mes, anio: anio, ars: 0, usd: 0, botellas: 0 };
      ventasMes[mes].ars += montoArs; ventasMes[mes].usd += montoUsd; ventasMes[mes].botellas += botellas;

      // Ventas por cliente
      var cli = contacto || '(sin nombre)';
      if (!ventasCliente[cli]) ventasCliente[cli] = { cliente: cli, ars: 0, usd: 0, botellas: 0, operaciones: 0 };
      ventasCliente[cli].ars += montoArs; ventasCliente[cli].usd += montoUsd;
      ventasCliente[cli].botellas += botellas; ventasCliente[cli].operaciones++;

      // Ventas por producto (para meses de inventario / margen)
      if (producto) {
        if (!ventasProducto[producto]) ventasProducto[producto] = { producto: producto, ars: 0, usd: 0, botellas: 0 };
        ventasProducto[producto].ars += montoArs; ventasProducto[producto].usd += montoUsd;
        ventasProducto[producto].botellas += botellas;
      }

      // Margen por añada — lado ingresos
      if (anada) {
        if (!anadaVentas[anada]) anadaVentas[anada] = { anada: anada, ars: 0, usd: 0, botellas: 0 };
        anadaVentas[anada].ars += montoArs; anadaVentas[anada].usd += montoUsd; anadaVentas[anada].botellas += botellas;
      }
    }

    if (concepto === 'Egreso') {
      // Margen por añada — lado costos
      if (anada) {
        if (!anadaCostos[anada]) anadaCostos[anada] = { anada: anada, ars: 0, usd: 0 };
        anadaCostos[anada].ars += -montoArs; anadaCostos[anada].usd += -montoUsd;
      }
      // Costo unitario acumulado por producto (suma de CU US$ de cada costo)
      if (producto) {
        if (!costoProducto[producto]) costoProducto[producto] = { producto: producto, cuUsd: 0, totalUsd: 0 };
        costoProducto[producto].cuUsd += Math.abs(cuUsd);
        costoProducto[producto].totalUsd += -montoUsd;
      }
      // Evolución de insumos: costo unitario por tipo de insumo y añada
      if (anada && detalle && Math.abs(cuUsd) > 0) {
        var clave = detalle + '||' + anada;
        if (!insumosAnada[clave]) insumosAnada[clave] = { insumo: detalle, anada: anada, cuUsd: 0 };
        insumosAnada[clave].cuUsd += Math.abs(cuUsd);
      }
    }
  }

  function aLista(obj, orden) {
    var out = [];
    for (var k in obj) out.push(obj[k]);
    if (orden) out.sort(orden);
    return out;
  }

  // Margen por añada: cruzar ventas y costos
  var anadas = {};
  for (var a in anadaVentas) anadas[a] = true;
  for (var b in anadaCostos) anadas[b] = true;
  var margenAnada = [];
  for (var k2 in anadas) {
    var v = anadaVentas[k2] || { ars: 0, usd: 0, botellas: 0 };
    var c = anadaCostos[k2] || { ars: 0, usd: 0 };
    margenAnada.push({
      anada: k2,
      ventasArs: Math.round(v.ars), ventasUsd: Math.round(v.usd), botellasVendidas: v.botellas,
      costosArs: Math.round(c.ars), costosUsd: Math.round(c.usd),
      margenArs: Math.round(v.ars - c.ars), margenUsd: Math.round(v.usd - c.usd)
    });
  }
  margenAnada.sort(function(x, y) { return x.anada < y.anada ? -1 : 1; });

  return {
    ventasMes:     aLista(ventasMes, function(x, y) { return x.periodo < y.periodo ? -1 : 1; }),
    ventasCliente: aLista(ventasCliente, function(x, y) { return y.ars - x.ars; }),
    ventasProducto: aLista(ventasProducto, function(x, y) { return y.ars - x.ars; }),
    flujoAnual:    aLista(flujoAnio, function(x, y) { return x.anio - y.anio; }),
    margenAnada:   margenAnada,
    costoProducto: aLista(costoProducto, function(x, y) { return y.cuUsd - x.cuUsd; }),
    insumosAnada:  aLista(insumosAnada, function(x, y) { return x.anada < y.anada ? -1 : (x.anada > y.anada ? 1 : (y.cuUsd - x.cuUsd)); })
  };
}

// Escribe las fórmulas de conversión ARS<->USD al dólar de hoy en un rango de filas.
function escribirFormulasInsumos(filaInicio, cantidad) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INSUMOS');
  if (!sheet) return;
  var dolarHoy = 'XLOOKUP(MAX(BLUE_API!$A$2:$A);BLUE_API!$A$2:$A;BLUE_API!$C$2:$C)';
  var formulas = [];
  for (var i = 0; i < cantidad; i++) {
    var f = filaInicio + i;
    formulas.push([
      '=IFERROR(IF($B' + f + '="";"";IF($C' + f + '="USD";$B' + f + '*' + dolarHoy + ';$B' + f + '));"")', // E COSTO ARS
      '=IFERROR(IF($B' + f + '="";"";IF($C' + f + '="USD";$B' + f + ';$B' + f + '/' + dolarHoy + '));"")'  // F COSTO USD
    ]);
  }
  sheet.getRange(filaInicio, 5, cantidad, 2).setFormulas(formulas);
}

function getInsumos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INSUMOS');
  if (!sheet) return { insumos: [], totalArs: 0, totalUsd: 0, existe: false };
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { insumos: [], totalArs: 0, totalUsd: 0, existe: true };

  var data = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
  var insumos = [], totalArs = 0, totalUsd = 0;
  var hoy = new Date();
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var nombre = r[0];
    if (!nombre || String(nombre).indexOf('COSTO TOTAL') === 0) continue; // saltear la fila de total
    var activo = String(r[7] || '').toUpperCase() !== 'NO';
    var costoArs = Number(r[4]) || 0;
    var costoUsd = Number(r[5]) || 0;
    var fecha = (r[3] instanceof Date) ? r[3] : null;
    var diasDesde = fecha ? Math.floor((hoy - fecha) / 86400000) : null;
    insumos.push({
      nombre:    nombre,
      costo:     Number(r[1]) || 0,
      moneda:    r[2] || 'ARS',
      fecha:     fecha ? formatDate(fecha) : '',
      diasDesde: diasDesde,
      costoArs:  costoArs,
      costoUsd:  costoUsd,
      proveedor: r[6] || '',
      activo:    activo,
      cargado:   !!(Number(r[1]) || 0)
    });
    if (activo) { totalArs += costoArs; totalUsd += costoUsd; }
  }
  return { insumos: insumos, totalArs: totalArs, totalUsd: totalUsd, existe: true };
}

// ── MOVIMIENTOS DE CAJA → CAJA ────────────────────────────────────────
// p.referencia es opcional: si viene, vincula el cobro/pago con una operación de
// BALANCE (misma REFERENCIA) para que su saldo se actualice solo. Si esa operación
// tiene varios productos, el monto se reparte proporcionalmente entre una fila de
// CAJA por producto (mismo patrón que BALANCE), en vez de una sola fila ambigua.
function addMovement(p) {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  if (!caja) throw new Error('No se encontró la pestaña CAJA');

  var signo      = p.tipo === 'gasto' ? -1 : 1;
  var montoTotal = (parseFloat(p.monto) || 0) * signo;
  var cajaLbl    = CAJA_LABELS[p.caja] || p.caja;

  var lineas = p.referencia ? buscarLineasPorReferencia(p.referencia) : [];
  var montos, productos;
  if (lineas.length > 1) {
    montos    = repartirProporcional(montoTotal, lineas.map(function(l) { return l.monto; }));
    productos = lineas.map(function(l) { return l.producto; });
  } else {
    montos    = [montoTotal];
    productos = [lineas.length === 1 ? lineas[0].producto : ''];
  }

  var ultimaReal = obtenerUltimaFilaConFecha(caja, 2); // B = FECHA
  montos.forEach(function(monto, i) {
    caja.insertRowAfter(ultimaReal);
    var row = ultimaReal + 1;
    ultimaReal = row; // la próxima línea se inserta justo debajo de esta
    caja.getRange(row, 1, 1, 9).setValues([[
      p.user || '',                                          // A: quién lo cargó
      parseFechaApp(p.fecha),                                // B FECHA
      p.detalle || (p.tipo === 'cobro' ? 'Cobro' : 'Gasto'),  // C DETALLE
      p.contacto || '',                                       // D SUBDETALLE
      productos[i] || '',                                     // E PRODUCTO (heredado de la operación vinculada, si hay)
      monto,                                                  // F MONTO $
      '=IFERROR(F' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A;BLUE_API!$C$2:$C;;-1);"")', // G MONTO US$
      cajaLbl,                                                // H CAJA
      p.referencia || ''                                      // I REFERENCIA (opcional → BALANCE)
    ]]);
  });

  enviarTelegram(
    (p.tipo === 'cobro' ? '💵 <b>Cobro</b>\n' : '💸 <b>Gasto</b>\n') +
    'Detalle: ' + escapeHtml(p.detalle || '-') + '\n' +
    'Cliente/Proveedor: ' + escapeHtml(p.contacto || '-') + '\n' +
    (lineas.length > 1 ? 'Repartido entre ' + lineas.length + ' productos\n' : '') +
    'Monto: $' + formatearMonto(montoTotal) + '\n' +
    'Caja: ' + escapeHtml(cajaLbl) + '\n' +
    'Cargado por: ' + escapeHtml(p.user || '-')
  );

  return { ok: true };
}

// ── TRANSFERENCIAS ENTRE DEPÓSITOS → STOCK ───────────────────────────
function addTransfer(p) {
  var cantidad = parseInt(p.cantidad) || 0;
  ajustarStockUbicacion(p.producto, p.desde, p.hacia, cantidad);

  var headers = ['FECHA','PRODUCTO','CANTIDAD','DESDE','HACIA','NOTAS','USUARIO','REGISTRADO'];
  var log = getOrCreateSheet(TAB_TRANSF, headers);
  log.appendRow([p.fecha, p.producto, cantidad, p.desde, p.hacia, p.notas || '', p.user, new Date()]);

  enviarTelegram(
    '🔀 <b>Transferencia de stock</b>\n' +
    cantidad + ' bot ' + escapeHtml(p.producto) + ': ' + escapeHtml(p.desde) + ' → ' + escapeHtml(p.hacia) + '\n' +
    'Cargado por: ' + escapeHtml(p.user || '-')
  );

  return { ok: true };
}

// ── CONTROL FÍSICO DE STOCK → historial compartido entre dispositivos ────
// p.items viene como JSON: [{label, stock, real, diff}, ...] — una fila por producto,
// todas con el mismo REGISTRADO (timestamp) para poder agruparlas como una sesión.
function addStockControl(p) {
  var headers = ['FECHA','HORA','USUARIO','DEPOSITO','PRODUCTO','TEORICO','REAL','DIFERENCIA','REGISTRADO'];
  var sheet   = getOrCreateSheet(TAB_CONTROL, headers);
  var items   = JSON.parse(p.items || '[]');
  var ahora   = new Date();
  items.forEach(function(it) {
    sheet.appendRow([p.fecha, p.hora, p.user, p.deposito || '', it.label, it.stock, it.real, it.diff, ahora]);
    fijarStockUbicacion(it.label, p.deposito, it.real); // corrige STOCK con lo contado de verdad
  });

  var difs = items.filter(function(it) { return it.diff !== 0; });
  var msg  = '📋 <b>Control de stock: ' + escapeHtml(p.deposito) + '</b>\n';
  if (difs.length) {
    msg += '⚠ Diferencias:\n' + difs.map(function(it) {
      return '· ' + escapeHtml(it.label) + ': ' + (it.diff > 0 ? '+' : '') + it.diff + ' bot';
    }).join('\n') + '\n';
  } else {
    msg += '✓ Sin diferencias\n';
  }
  msg += 'Cargado por: ' + escapeHtml(p.user || '-');
  enviarTelegram(msg);

  return { ok: true };
}

// Devuelve el control físico más reciente (agrupando las filas que comparten el
// mismo REGISTRADO), para que cualquier dispositivo vea el mismo "último control".
function getUltimoControlStock() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_CONTROL);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  var maxTs = null;
  for (var i = 0; i < data.length; i++) {
    var ts = data[i][8];
    if (ts instanceof Date && (!maxTs || ts.getTime() > maxTs.getTime())) maxTs = ts;
  }
  if (!maxTs) return null;

  var fecha = '', hora = '', deposito = '', items = [];
  for (var j = 0; j < data.length; j++) {
    var row = data[j];
    if (row[8] instanceof Date && row[8].getTime() === maxTs.getTime()) {
      fecha = formatDate(row[0]); hora = formatHora(row[1]); deposito = row[3];
      items.push({ label: row[4], stock: row[5], real: row[6], diff: row[7] });
    }
  }
  return { fecha: fecha, hora: hora, deposito: deposito, items: items };
}

// Lee el estado actual de stock por producto y depósito directo de STOCK
// (la tabla de depósitos ya existente), traduciendo nombres al vocabulario de la app.
function getStockUbicacion() {
  var stock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STOCK');
  if (!stock) return {};
  var data    = stock.getRange('B12:I18').getValues();
  var headers = data[0];
  var out = {};
  for (var i = 1; i < data.length; i++) {
    var producto = data[i][0];
    if (!producto) continue;
    var appName = null;
    for (var k in PRODUCTO_APP_TO_SHEET) { if (PRODUCTO_APP_TO_SHEET[k] === producto) { appName = k; break; } }
    if (!appName) appName = producto;
    var ubic = {};
    for (var u in UBIC_SHEET_MAP) {
      var col = headers.indexOf(UBIC_SHEET_MAP[u]);
      ubic[u] = col > -1 ? (Number(data[i][col]) || 0) : 0;
    }
    out[appName] = ubic;
  }
  return out;
}

// Mueve stock entre depósitos (o sólo descuenta, si hacia es null — para una venta)
// directo en la tabla de depósitos de STOCK. TOTAL y DIFERENCIA son fórmulas ya
// existentes en esa tabla y no las tocamos: se recalculan solas.
function ajustarStockUbicacion(productoApp, desde, hacia, cantidad) {
  var stock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STOCK');
  if (!stock) throw new Error('No se encontró la pestaña STOCK');
  var data        = stock.getRange('B12:I18').getValues();
  var headers     = data[0];
  var nombreSheet = PRODUCTO_APP_TO_SHEET[productoApp] || productoApp;

  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === nombreSheet) { rowIdx = i; break; }
  }
  if (rowIdx === -1) throw new Error('Producto no encontrado en STOCK: ' + productoApp);

  var sheetRow = 12 + rowIdx;
  if (desde) {
    var colDesde = headers.indexOf(UBIC_SHEET_MAP[desde] || desde);
    if (colDesde > -1) stock.getRange(sheetRow, 2 + colDesde).setValue((Number(data[rowIdx][colDesde]) || 0) - cantidad);
  }
  if (hacia) {
    var colHacia = headers.indexOf(UBIC_SHEET_MAP[hacia] || hacia);
    if (colHacia > -1) stock.getRange(sheetRow, 2 + colHacia).setValue((Number(data[rowIdx][colHacia]) || 0) + cantidad);
  }
}

// Fija el valor REAL contado en un depósito para un producto (a diferencia de
// ajustarStockUbicacion, no suma/resta — pisa el valor con lo contado físicamente).
// La usa el Control de stock para corregir STOCK con lo que se encontró de verdad.
function fijarStockUbicacion(productoApp, deposito, valor) {
  var stock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STOCK');
  if (!stock) throw new Error('No se encontró la pestaña STOCK');
  var data        = stock.getRange('B12:I18').getValues();
  var headers     = data[0];
  var nombreSheet = PRODUCTO_APP_TO_SHEET[productoApp] || productoApp;

  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === nombreSheet) { rowIdx = i; break; }
  }
  if (rowIdx === -1) throw new Error('Producto no encontrado en STOCK: ' + productoApp);

  var col = headers.indexOf(UBIC_SHEET_MAP[deposito] || deposito);
  if (col === -1) throw new Error('Depósito no reconocido: ' + deposito);
  stock.getRange(12 + rowIdx, 2 + col).setValue(valor);
}

// ── OPERACIONES RECIENTES (ventas + movimientos + transferencias) ───
// Ventas y Movimientos se identifican por tener algo cargado en la columna A
// (que hasta ahora estaba sin usar) — ahí guardamos quién lo cargó desde la app,
// así distinguimos "lo cargó la app" de las ~850 filas históricas ya existentes.
function getRecentOps(n) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ops = [];

  var balance = ss.getSheetByName('BALANCE');
  if (balance && balance.getLastRow() > 2) {
    var db = balance.getRange(3, 1, balance.getLastRow() - 2, 13).getValues(); // A..M
    for (var i = db.length - 1; i >= 0; i--) {
      var rb = db[i];
      if (rb[2] !== 'Venta' || !(rb[1] instanceof Date)) continue; // C DETALLE, B FECHA (real, no una fila de totales/vacía)
      ops.push({
        id:    'b_' + i,
        icon:  '🍾',
        desc:  rb[12] + ' bot ' + rb[4] + ' → ' + rb[3], // M botellas, E producto, D cliente
        monto: '$' + Math.round(rb[6]).toLocaleString(), // G monto $
        fecha: formatDate(rb[1]),
        user:  rb[0] || '',
        ts:    rb[1].getTime()
      });
    }
  }

  var caja = ss.getSheetByName('CAJA');
  if (caja && caja.getLastRow() > 2) {
    var dc = caja.getRange(3, 1, caja.getLastRow() - 2, 9).getValues(); // A..I
    for (var j = dc.length - 1; j >= 0; j--) {
      var rc = dc[j];
      if (!rc[0] || (rc[2] !== 'Cobro' && rc[2] !== 'Gasto')) continue; // A usuario, C DETALLE
      ops.push({
        id:    'c_' + j,
        icon:  rc[2] === 'Cobro' ? '💵' : '💸',
        desc:  rc[2] + ': ' + (rc[3] || ''),
        monto: (rc[5] < 0 ? '-' : '') + '$' + Math.abs(Math.round(rc[5])).toLocaleString() + ' (' + rc[7] + ')',
        fecha: formatDate(rc[1]),
        user:  rc[0],
        ts:    rc[1] instanceof Date ? rc[1].getTime() : j
      });
    }
  }

  var st = ss.getSheetByName(TAB_TRANSF);
  if (st && st.getLastRow() > 1) {
    var dt = st.getDataRange().getValues();
    for (var k = dt.length - 1; k >= 1; k--) {
      var rt = dt[k];
      ops.push({
        id:    't_' + k,
        icon:  '🔀',
        desc:  rt[2] + ' bot ' + rt[1] + ': ' + rt[3] + ' → ' + rt[4],
        monto: '—',
        fecha: formatDate(rt[0]),
        user:  rt[6],
        ts:    rt[7] ? new Date(rt[7]).getTime() : k
      });
    }
  }

  ops.sort(function(a, b) { return (b.ts || 0) - (a.ts || 0); });
  return ops.slice(0, n);
}

// Detalle completo de una operación (todas las filas de BALANCE con esa REFERENCIA,
// más todos los cobros/pagos de CAJA vinculados). Para la vista de detalle que se abre
// al tocar una operación pendiente en el dashboard.
function getDetalleOperacion(referencia) {
  if (!referencia) return { lineas: [], pagos: [] };
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Líneas de BALANCE (una por producto)
  var balance = ss.getSheetByName('BALANCE');
  var lineas = [], cabecera = null;
  if (balance) {
    var lastRowB = balance.getLastRow();
    if (lastRowB >= 3) {
      var cDepD = colBalance_('DEPOSITO');
      var dataB = balance.getRange(3, 1, lastRowB - 2, Math.max(17, balance.getLastColumn())).getValues();
      for (var i = 0; i < dataB.length; i++) {
        var r = dataB[i];
        if (r[11] !== referencia) continue; // L REFERENCIA
        var botellas = Number(r[12]) || 0;  // M BOTELLAS
        var montoArs = Number(r[6]) || 0;   // G MONTO $
        lineas.push({
          producto:  r[4] || '',                                    // E
          botellas:  botellas,
          montoArs:  Math.round(montoArs),
          montoUsd:  Math.round(Number(r[7]) || 0),                  // H MONTO US$
          precioUnit: botellas ? Math.round(montoArs / botellas) : null, // precio por botella
          deposito:  cDepD ? (r[cDepD - 1] || '') : '',                // DEPOSITO (por encabezado)
          fila:      i + 3
        });
        if (!cabecera) cabecera = {
          fecha:    formatDate(r[1]),   // B
          detalle:  r[2] || '',         // C
          contacto: r[3] || '',         // D
          user:     r[0] || '',         // A quién la cargó
          fechaMs:  (r[1] instanceof Date) ? r[1].getTime() : null
        };
      }
    }
  }

  // Cobros/pagos de CAJA vinculados a esa referencia
  var caja = ss.getSheetByName('CAJA');
  var pagos = [];
  if (caja) {
    var lastRowC = caja.getLastRow();
    if (lastRowC >= 3) {
      var dataC = caja.getRange(3, 1, lastRowC - 2, 9).getValues(); // A..I
      for (var j = 0; j < dataC.length; j++) {
        var c = dataC[j];
        if (c[8] !== referencia) continue; // I REFERENCIA
        if (!(c[1] instanceof Date)) continue;
        pagos.push({
          fecha:    formatDate(c[1]),          // B
          detalle:  c[2] || '',                // C
          producto: c[4] || '',                // E
          montoArs: Math.round(Number(c[5]) || 0), // F
          montoUsd: Math.round(Number(c[6]) || 0), // G
          caja:     c[7] || '',                // H
          user:     c[0] || '',                // A
          fechaMs:  c[1].getTime()
        });
      }
    }
  }

  // Totales: lo facturado vs lo cobrado/pagado
  var totalArs = 0, totalUsd = 0, totalBotellas = 0;
  lineas.forEach(function(l) { totalArs += l.montoArs; totalUsd += l.montoUsd; totalBotellas += l.botellas; });
  var pagadoArs = 0, pagadoUsd = 0, ultimoPagoMs = null;
  pagos.forEach(function(p) {
    pagadoArs += p.montoArs; pagadoUsd += p.montoUsd;
    if (p.fechaMs && (!ultimoPagoMs || p.fechaMs > ultimoPagoMs)) ultimoPagoMs = p.fechaMs;
  });

  // Días: si está cerrada, cuánto tardó (fecha → último movimiento); si sigue abierta,
  // cuántos días lleva desde que se cargó.
  var saldoArs = Math.round(totalArs - pagadoArs);
  var cerrada = saldoArs === 0 && pagos.length > 0;
  var dias = null;
  if (cabecera && cabecera.fechaMs) {
    var hasta = cerrada ? ultimoPagoMs : new Date().getTime();
    // Diferencia en días calendario de Mendoza, sin la hora (evita medios días)
    var dia = function(ms) { return Number(Utilities.formatDate(new Date(ms), 'America/Argentina/Mendoza', 'yyyyMMdd')); };
    var aFecha = function(n) { return Date.UTC(Math.floor(n / 10000), Math.floor(n / 100) % 100 - 1, n % 100); };
    if (hasta) dias = Math.round((aFecha(dia(hasta)) - aFecha(dia(cabecera.fechaMs))) / 86400000);
  }

  return {
    referencia:   referencia,
    cabecera:     cabecera || {},
    lineas:       lineas,
    pagos:        pagos,
    totalArs:     Math.round(totalArs),
    totalUsd:     Math.round(totalUsd),
    totalBotellas: totalBotellas,
    pagadoArs:    Math.round(pagadoArs),
    pagadoUsd:    Math.round(pagadoUsd),
    saldoArs:     saldoArs,
    saldoUsd:     Math.round(totalUsd - pagadoUsd),
    cerrada:      cerrada,
    dias:         dias
  };
}

// Últimos N movimientos reales de CAJA (cobros/gastos que cargaron los socios), del
// más reciente al más viejo, para el cuadro de control en la pestaña Caja de la app.
// Lee sólo el bloque final de la hoja (no toda), y descarta filas sin fecha real
// (vacías, o el bloque resumen "CAJAS" del pie). Devuelve todos los campos que se
// muestran: fecha, detalle, subdetalle (cliente/proveedor), monto, caja y quién cargó.
function getMovimientosCaja(n) {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  if (!caja) return [];
  var lastRow = caja.getLastRow();
  if (lastRow < 3) return [];

  var data = caja.getRange(3, 1, lastRow - 2, 9).getValues(); // A..I: USER,FECHA,DETALLE,SUBDET,PRODUCTO,MONTO$,MONTOUS$,CAJA,REFERENCIA
  var out = [];
  // Recorrer de abajo hacia arriba (lo más reciente primero) y cortar al juntar N.
  for (var i = data.length - 1; i >= 0 && out.length < n; i--) {
    var r = data[i];
    if (!(r[1] instanceof Date)) continue; // sólo filas con fecha real (evita vacías y el bloque CAJAS)
    var montoNum = Number(r[5]) || 0;      // F MONTO $
    out.push({
      fecha:     formatDate(r[1]),          // B
      detalle:   r[2] || '',                // C
      contacto:  r[3] || '',                // D SUBDETALLE (cliente/proveedor)
      monto:     Math.round(montoNum),       // número crudo; el signo/formato lo pone el front
      caja:      r[7] || '',                // H
      user:      r[0] || '',                // A
      ts:        r[1].getTime()
    });
  }
  return out;
}

// ── UTILIDAD OPCIONAL — correr UNA SOLA VEZ a mano ────────────────────
// Reclasifica los 70 pagos de CAJA bajo CCI22-001 según el criterio real acordado
// (no solo por año, sino por a qué correspondía cada uno):
//  · AFIP/Contadores (cualquier año) → costo operativo de la empresa, no de producción.
//    Pasan a CO (referencia CCI23-001, DETALLE='CO').
//  · Ya etiquetados con Malbec 2023 (el préstamo Semilla + 1 de Impuestos y Débitos) →
//    costo real de esa producción. Pasan a su propia referencia SEM23-001.
//  · Envío → correspondían a envíos de productos de la partida 2022, quedan en CCI22-001.
//  · El resto (costo indirecto real de 2022 + 2 filas ambiguas de ajuste/movimiento
// ── UTILIDAD OPCIONAL — correr UNA VEZ a mano ────────────────────────
// Las filas de BALANCE que reparten un costo indirecto entre varios productos (ej.
// las de CCI22-001: Cabernet Franc/Malbec/Chardonnay 2022) calculan su MONTO $ con
// una fórmula propia: =SUMIF(CAJA!...)*VLOOKUP(...;STOCK...). Esa fórmula quedó con
// rangos de CAJA DISTINTOS entre sí (una desactualizada, las otras más amplias),
// dando resultados inconsistentes entre filas que deberían coincidir. Esto normaliza
// ── ACTUALIZACIÓN DIARIA DE BLUE_API ─────────────────────────────────
// Reemplaza el mecanismo que se usaba antes (armado desde Excel/Claude en Excel, de
// fuente desconocida y que dejó de correr). Usa el endpoint de evolución histórica
// (no sólo "latest") para que, además de agregar la cotización de hoy, rellene
// automáticamente cualquier hueco de fechas que haya quedado sin cargar (por
// ejemplo si el trigger diario falla un día). B=compra (value_buy), C=venta
// (value_sell), mismo orden que ya usan el resto de las fórmulas de la planilla.
function actualizarBlueApi() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BLUE_API');
  if (!sheet) { Logger.log('No encontré la pestaña BLUE_API.'); return; }

  var resp = UrlFetchApp.fetch('https://api.bluelytics.com.ar/v2/evolution.json', { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    Logger.log('Error consultando bluelytics.com.ar: código ' + resp.getResponseCode());
    return;
  }
  var data = JSON.parse(resp.getContentText());
  var blue = Array.isArray(data) ? data.filter(function(d) { return d.source && String(d.source).toLowerCase() === 'blue'; }) : [];
  if (!blue.length) {
    Logger.log('La respuesta no trajo cotizaciones "blue" — revisar formato. Primeros 500 caracteres: ' + resp.getContentText().substring(0, 500));
    return;
  }

  // Reconstrucción completa: el endpoint trae TODO el histórico (desde 2011), que es
  // la misma fuente de la que salió esta tabla. En vez de parchar filas sueltas —
  // frágil y propenso a huecos/desorden — rearmamos la tabla entera limpia:
  // deduplicada por fecha, ordenada descendente (más reciente arriba), fechas puras
  // sin hora, y una sola escritura. Esto arregla de raíz cualquier desorden previo.
  var porFecha = {};
  blue.forEach(function(d) {
    if (d.date && d.value_buy != null && d.value_sell != null) porFecha[d.date] = d; // dedup: última gana
  });

  var fechas = Object.keys(porFecha).sort().reverse(); // descendente (yyyy-MM-dd ordena bien como texto)
  var filas = fechas.map(function(k) {
    var d = porFecha[k];
    var partes = k.split('-');
    // MEDIANOCHE de Mendoza (script y planilla están en America/Argentina/Mendoza).
    // Las fórmulas buscan la cotización con XLOOKUP "exacto o anterior": con la
    // cotización a las 00:00, una operación del día D a cualquier hora encuentra la de D.
    // (Antes se usaba mediodía como parche porque el script estaba en Darwin y la
    // planilla en Los Ángeles; eso corría todas las fechas un día.)
    var fecha = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]), 0, 0, 0);
    return [fecha, d.value_buy, d.value_sell]; // B=compra (value_buy), C=venta (value_sell)
  });

  // Limpiar todo debajo del encabezado y reescribir de una sola vez.
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var viejo = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 3));
    viejo.clearContent();
    viejo.clearFormat(); // saca bandas de color / formato condicional viejo que hacían ver filas "vacías"
  }
  var destino = sheet.getRange(2, 1, filas.length, 3);
  destino.setValues(filas);
  destino.setFontColor('#000000'); // texto negro visible (el formato viejo lo tenía casi blanco)
  destino.setBackground(null);     // fondo uniforme
  sheet.getRange(2, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy'); // columna FECHA como fecha pura, sin hora
  sheet.getRange(1, 1, 1, 3).setValues([['day', 'value_buy', 'value_sell']]); // encabezado correcto: B=compra, C=venta

  Logger.log('Listo — BLUE_API reconstruida: ' + filas.length + ' fechas, de ' + fechas[fechas.length-1] + ' a ' + fechas[0] + '.');
}

// ── HELPERS ──────────────────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Argentina/Mendoza', 'dd/MM/yyyy');
  return String(val);
}

function formatHora(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Argentina/Mendoza', 'HH:mm');
  return String(val);
}
