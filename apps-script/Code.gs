// ═══════════════════════════════════════════════════════════════════
//  NOVATO BODEGA · Google Apps Script Backend
//  Pegar TODO este código en Extensions → Apps Script del Google Sheet
//  Luego: Deploy → Manage deployments → editar el deployment existente
//    → Version: New version → Deploy  (mantiene la misma URL)
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
  'Malbec 2023':      'Malbec 2023',
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

// Contactos reales (SUBDETALLE) vistos en BALANCE — clientes Y proveedores mezclados,
// para el selector de "Cliente/Proveedor" del formulario general de transacciones.
function getContactosBalance() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return [];
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 4, lastRow - 2, 1).getValues(); // D SUBDETALLE
  var vistos = {}, lista = [];
  for (var i = 0; i < data.length; i++) {
    var v = data[i][0];
    if (v && !vistos[v]) { vistos[v] = true; lista.push(v); }
  }
  lista.sort();
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
// incluirSaldo=false se usa para filas que comparten referencia con otras (multi-
// producto, o costos prorrateados tipo CCI22-001): el saldo se maneja agregado por
// referencia en getVentasPendientes/getComprasPendientes, no fila por fila.
function escribirFormulasBalance(row, incluirSaldo) {
  if (incluirSaldo === undefined) incluirSaldo = true;
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  balance.getRange(row, 6).setValue('=RIGHT(E' + row + ';4)'); // F AÑADA
  balance.getRange(row, 8, 1, 4).setValues([[
    '=IF(G' + row + '=0;"";G' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A$4590;BLUE_API!$C$2:$C$4590;;-1))', // H MONTO US$ FF
    incluirSaldo ? '=IF(G' + row + '=0;"";H' + row + '+P' + row + ')' : '',                                       // I MONTO US$ FP
    '=IF(OR(G' + row + '=0;E' + row + '="");"";IF(G' + row + '>0;IF(M' + row + '=0;"";G' + row + '/M' + row + ');G' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)))', // J CU $ (venta sin botellas cargadas, ej. operación contable sin entrega física: vacío)
    '=IF(OR(G' + row + '=0;E' + row + '="");"";IF(G' + row + '>0;IF(M' + row + '=0;"";H' + row + '/M' + row + ');H' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)))'  // K CU US$
  ]]);
  balance.getRange(row, 14).setValue('=IF(G' + row + '>0;"Ingreso";(IF(G' + row + '=0;"Movimiento";"Egreso")))'); // N CONCEPTO
  if (incluirSaldo) {
    balance.getRange(row, 15, 1, 3).setValues([[
      '=IF(G' + row + '=0;"";G' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$F$3:$F))',                     // O SALDO $
      '=IF(G' + row + '=0;"";-(H' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$G$3:$G)))',                  // P SALDO US$
      '=IF(G' + row + '=0;"";IF(N' + row + '="Egreso";-P' + row + '/H' + row + ';P' + row + '/H' + row + '))' // Q % DIF POR TC
    ]]);
  } else {
    balance.getRange(row, 15, 1, 3).setValues([['', '', '']]);
  }
  balance.getRange(row, 18).setValue('=IF(BALANCE!$B' + row + '="";"";YEAR(BALANCE!$B' + row + '))'); // R AÑO
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

// ── UTILIDAD — correr UNA VEZ a mano para averiguar el chat_id de cada persona ──
// Antes de correr esto, pedile a cada uno que le escriba cualquier cosa (ej. "hola")
// al bot desde Telegram. Después elegí esta función en el desplegable y tocá ▶ Run —
// el resultado queda en Ver → Registros (View → Logs).
function obtenerChatIdsTelegram() {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) { Logger.log('Primero configurá TELEGRAM_BOT_TOKEN en Project Settings → Script Properties.'); return; }
  var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates');
  var data = JSON.parse(resp.getContentText());
  if (!data.ok || !data.result.length) { Logger.log('Todavía no hay mensajes. Pedile a cada persona que le escriba algo al bot primero.'); return; }
  data.result.forEach(function(u) {
    var chat = u.message && u.message.chat;
    if (chat) Logger.log('Nombre: ' + (chat.first_name || '') + ' ' + (chat.last_name || '') + ' — chat_id: ' + chat.id);
  });
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

    escribirFormulasBalance(row, !multiLinea);

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
      '=F' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A$4590;BLUE_API!$C$2:$C$4590;;-1)', // G MONTO US$
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

// ── UTILIDAD OPCIONAL — correr UNA VEZ a mano si querés ────────────────
// Reescribe TODAS las fórmulas calculadas de BALANCE (sin importar en qué estado
// estén — rotas, con rango viejo, o directamente bien) usando las mismas plantillas
// que addTransaccion, más la fila de TOTALES (con una fórmula basada en ROW() que no
// depende de ningún número de fila fijo, así nunca vuelve a quedar corta ni corre
// riesgo de auto-referenciarse) y el EGRESOS agregado de STOCK (rango abierto real
// hacia BALANCE). Es un arreglo definitivo — no perjudica nada si se corre de nuevo.
function repararTodasLasFormulas() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var lastRow = balance.getLastRow();
  if (lastRow < 3) { Logger.log('BALANCE está vacío, nada para reparar.'); return; }

  var n = lastRow - 2;
  var fechas      = balance.getRange(3, 2, n, 1).getValues();  // B FECHA — 1 sola llamada para todo el rango
  var conceptos   = balance.getRange(3, 14, n, 1).getValues(); // N — para ubicar la fila 'TOTALES:'
  var referencias = balance.getRange(3, 12, n, 1).getValues(); // L — para detectar costos prorrateados

  // Si una REFERENCIA aparece en más de una fila, es un costo indirecto prorrateado
  // entre varios productos (ej. CCI22-001 repartido entre los 3 vinos de 2022 según
  // su % de producción) — ahí MONTO $ ya es una fórmula que mira a CAJA directamente,
  // así que el SALDO de reconciliación fila-por-fila no tiene sentido y se deja en
  // blanco a propósito. Detectarlo por repetición evita tener que listar códigos a mano.
  var conteoRef = {};
  for (var c = 0; c < n; c++) {
    var ref = referencias[c][0];
    if (ref) conteoRef[ref] = (conteoRef[ref] || 0) + 1;
  }

  var filaTotales = -1;
  var filas = [], filasProrrateadas = [];
  for (var i = 0; i < n; i++) {
    var row = i + 3;
    if (conceptos[i][0] === 'TOTALES:') { filaTotales = row; continue; }
    if (fechas[i][0] instanceof Date) {
      filas.push(row);
      if (referencias[i][0] && conteoRef[referencias[i][0]] > 1) filasProrrateadas.push(row);
    }
  }

  // Agrupar en bloques consecutivos, para escribir cada bloque con UNA sola llamada
  // por grupo de columnas (en vez de una llamada por celda) — esto es lo que importa
  // para la velocidad: pocas llamadas grandes, no miles de llamadas chicas.
  var grupos = [], actual = [];
  for (var g = 0; g < filas.length; g++) {
    if (actual.length === 0 || filas[g] === actual[actual.length - 1] + 1) actual.push(filas[g]);
    else { grupos.push(actual); actual = [filas[g]]; }
  }
  if (actual.length) grupos.push(actual);

  grupos.forEach(function(grupo) {
    var inicio = grupo[0];
    var colF = [], colHK = [], colN = [], colOR = [];
    grupo.forEach(function(row) {
      colF.push(['=RIGHT(E' + row + ';4)']);
      colHK.push([
        '=IF(G' + row + '=0;"";G' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A$4590;BLUE_API!$C$2:$C$4590;;-1))',
        '=IF(G' + row + '=0;"";H' + row + '+P' + row + ')',
        '=IF(OR(G' + row + '=0;E' + row + '="");"";IF(G' + row + '>0;IF(M' + row + '=0;"";G' + row + '/M' + row + ');G' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)))',
        '=IF(OR(G' + row + '=0;E' + row + '="");"";IF(G' + row + '>0;IF(M' + row + '=0;"";H' + row + '/M' + row + ');H' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)))'
      ]);
      colN.push(['=IF(G' + row + '>0;"Ingreso";(IF(G' + row + '=0;"Movimiento";"Egreso")))']);
      colOR.push([
        '=IF(G' + row + '=0;"";G' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$F$3:$F))',
        '=IF(G' + row + '=0;"";-(H' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$G$3:$G)))',
        '=IF(G' + row + '=0;"";IF(N' + row + '="Egreso";-P' + row + '/H' + row + ';P' + row + '/H' + row + '))',
        '=IF(BALANCE!$B' + row + '="";"";YEAR(BALANCE!$B' + row + '))'
      ]);
    });
    balance.getRange(inicio, 6, grupo.length, 1).setValues(colF);
    balance.getRange(inicio, 8, grupo.length, 4).setValues(colHK);
    balance.getRange(inicio, 14, grupo.length, 1).setValues(colN);
    balance.getRange(inicio, 15, grupo.length, 4).setValues(colOR);
  });

  // Ahora sí, dejar en blanco O:Q (SALDO $, SALDO US$, % DIF) de las filas prorrateadas,
  // agrupadas en bloques consecutivos igual que arriba.
  if (filasProrrateadas.length) {
    var gruposProrr = [], actualP = [];
    for (var gp = 0; gp < filasProrrateadas.length; gp++) {
      if (actualP.length === 0 || filasProrrateadas[gp] === actualP[actualP.length - 1] + 1) actualP.push(filasProrrateadas[gp]);
      else { gruposProrr.push(actualP); actualP = [filasProrrateadas[gp]]; }
    }
    if (actualP.length) gruposProrr.push(actualP);
    gruposProrr.forEach(function(grupo) {
      var vacio = grupo.map(function(){ return ['', '', '']; });
      balance.getRange(grupo[0], 15, grupo.length, 3).setValues(vacio); // O:Q
    });
  }

  if (filaTotales > 0) {
    // ROW()-1 apunta siempre a "la fila justo arriba mío", sin importar cuánto se
    // haya desplazado esta fila de TOTALES por inserciones — nunca queda corto ni
    // se auto-referencia.
    balance.getRange(filaTotales, 15).setValue('=SUM(INDIRECT("O3:O"&(ROW()-1)))');
    balance.getRange(filaTotales, 16).setValue('=SUMIF(INDIRECT("O3:O"&(ROW()-1));0;INDIRECT("P3:P"&(ROW()-1)))');
  }

  var stock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STOCK');
  var stockReparado = false;
  if (stock) {
    var colEgresos = [];
    for (var r2 = 4; r2 <= 9; r2++) {
      colEgresos.push(['=SUMIF(BALANCE!$E$2:$E;STOCK!$B' + r2 + ';BALANCE!$M$2:$M)']);
    }
    stock.getRange(4, 6, 6, 1).setValues(colEgresos);
    stockReparado = true;
  }

  Logger.log('Listo — ' + filas.length + ' fila(s) de BALANCE reescritas en ' + grupos.length + ' bloque(s)' +
    (filaTotales > 0 ? ', TOTALES (fila ' + filaTotales + ') arreglado' : ', no encontré la fila TOTALES') +
    (stockReparado ? ', EGRESOS de STOCK reescrito.' : '.'));
}

// ── DESHACER LA VENTA DE PRUEBA (correr UNA SOLA VEZ, después borrar o ignorar) ──
// Específico para el caso puntual: borra la fila 379 de BALANCE (la venta de prueba
// rota, 6 bot Cabernet Franc 2022) y le devuelve esas 6 botellas a R Peña en STOCK,
// ── UTILIDAD OPCIONAL — correr UNA VEZ a mano si querés ────────────────
// Antes del arreglo de addMovement, los movimientos de caja cargados desde la app
// dejaban un número congelado en MONTO US$ (columna G) en vez de la fórmula XLOOKUP
// que usa el resto de esa columna. Esto reescribe la fórmula en cualquier fila de
// CAJA cargada desde la app (columna A no vacía, o sea con usuario) donde G no
// tenga ya una fórmula. No toca filas históricas manuales.
function repararMontoUSCaja() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  var lastRow = caja.getLastRow();
  if (lastRow < 3) { Logger.log('CAJA está vacía, nada para reparar.'); return; }

  var n = lastRow - 2;
  var usuarios  = caja.getRange(3, 1, n, 1).getValues();    // A
  var formulasG = caja.getRange(3, 7, n, 1).getFormulas();  // G — '' si es un valor plano, no una fórmula

  var filas = [];
  for (var i = 0; i < n; i++) {
    if (usuarios[i][0] && !formulasG[i][0]) filas.push(i + 3);
  }

  var grupos = [], actual = [];
  for (var g = 0; g < filas.length; g++) {
    if (actual.length === 0 || filas[g] === actual[actual.length - 1] + 1) actual.push(filas[g]);
    else { grupos.push(actual); actual = [filas[g]]; }
  }
  if (actual.length) grupos.push(actual);

  grupos.forEach(function(grupo) {
    var col = grupo.map(function(row) {
      return ['=F' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A$4590;BLUE_API!$C$2:$C$4590;;-1)'];
    });
    caja.getRange(grupo[0], 7, grupo.length, 1).setValues(col);
  });

  Logger.log('Listo — ' + filas.length + ' fila(s) de CAJA con MONTO US$ reparado.');
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
  var blue = Array.isArray(data) ? data.filter(function(d) { return d.source === 'blue'; }) : [];
  if (!blue.length) {
    Logger.log('La respuesta no trajo cotizaciones "blue" — revisar formato. Primeros 500 caracteres: ' + resp.getContentText().substring(0, 500));
    return;
  }
  blue.sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); }); // ascendente

  var lastRow = sheet.getLastRow();
  var existentes = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r) {
      if (r[0] instanceof Date) existentes[Utilities.formatDate(r[0], 'America/Argentina/Mendoza', 'yyyy-MM-dd')] = true;
    });
  }

  var agregadas = 0;
  blue.forEach(function(d) {
    if (existentes[d.date]) return; // ya está cargada, no duplicar
    var fecha = new Date(d.date + 'T00:00:00');
    sheet.insertRowAfter(1); // siempre después del encabezado — procesando en orden ascendente, el resultado final queda con lo más reciente arriba
    sheet.getRange(2, 1, 1, 3).setValues([[fecha, d.value_buy, d.value_sell]]);
    existentes[d.date] = true;
    agregadas++;
  });

  Logger.log('Listo — ' + agregadas + ' fecha(s) agregada(s) a BLUE_API (incluye cualquier hueco encontrado).');
}

// UTILIDAD — correr UNA VEZ a mano para activar la actualización diaria automática.
// No hace falta tocar el menú de Triggers a mano: esto crea el trigger por código.
// Si ya existe uno para esta función, no lo duplica.
function instalarTriggerBlueApi() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'actualizarBlueApi') {
      Logger.log('Ya existe un trigger para actualizarBlueApi — no se crea otro.');
      return;
    }
  }
  ScriptApp.newTrigger('actualizarBlueApi').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Listo — actualizarBlueApi() va a correr automáticamente todos los días alrededor de las 9am.');
}

// transacción real y el resumen de CAJAS del pie de la hoja. Como no tienen fecha,
// esa fórmula da #DIV/0! y ensucia cualquier búsqueda de errores. Esto las limpia,
// tocando solamente filas donde FECHA está vacía (nunca una fila con datos reales).
function limpiarFilasVaciasCaja() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  var lastRow = caja.getLastRow();
  if (lastRow < 3) { Logger.log('CAJA está vacía.'); return; }

  var n = lastRow - 2;
  var fechas = caja.getRange(3, 2, n, 1).getValues(); // B FECHA — sólo para identificar filas vacías
  var cambios = 0;
  for (var i = 0; i < n; i++) {
    if (fechas[i][0] === '' || fechas[i][0] === null) {
      var celda = caja.getRange(i + 3, 7); // G MONTO US$
      if (celda.getFormula()) { celda.clearContent(); cambios++; }
    }
  }
  Logger.log('Listo — ' + cambios + ' fila(s) vacía(s) de CAJA con fórmula suelta, limpiadas.');
}

// el rango de CAJA en TODAS las fórmulas de este tipo a uno abierto, para que no
// vuelva a pasar. Sólo toca la columna G (MONTO $) y sólo en filas que ya tienen
// este patrón — no afecta ninguna otra fórmula.
function normalizarFormulasProrrateadasCI() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var lastRow = balance.getLastRow();
  if (lastRow < 3) { Logger.log('BALANCE está vacía.'); return; }

  var n = lastRow - 2;
  // getFormulas() se usa sólo para IDENTIFICAR filas candidatas — nunca se vuelve a
  // escribir el array completo de una, así una celda con un valor plano (no fórmula,
  // que getFormulas() siempre devuelve como '' aunque tenga un número cargado) nunca
  // puede terminar pisada por accidente.
  var formulas = balance.getRange(3, 7, n, 1).getFormulas(); // G MONTO $
  var cambios = 0;

  for (var i = 0; i < n; i++) {
    var f = formulas[i][0];
    if (!f || f.indexOf('SUMIF(CAJA') === -1) continue; // sólo filas con este patrón de prorrateo
    var nueva = f
      .replace(/CAJA!\$?I\$?\d+:\$?I\$?\d+/, 'CAJA!$I$3:$I')
      .replace(/CAJA!\$?F\$?\d+:\$?F\$?\d+/, 'CAJA!$F$3:$F');
    if (nueva !== f) {
      balance.getRange(i + 3, 7).setValue(nueva); // escribe SÓLO esta celda puntual
      cambios++;
    }
  }

  Logger.log('Listo — ' + cambios + ' fórmula(s) de MONTO $ (costos prorrateados) normalizadas a rango abierto.');
}

//    de caja) queda sin tocar en CCI22-001.
// Después crea las 2 filas correspondientes en BALANCE (CO sin producto, y Semilla con
// Malbec 2023) para que el saldo de cada referencia reconcilie correctamente.
function reclasificarCostosIndirectos() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  var cajaLastRow = caja.getLastRow();
  if (cajaLastRow < 3) { Logger.log('CAJA está vacía.'); return; }

  var n = cajaLastRow - 2;
  var subdetalles = caja.getRange(3, 4, n, 1).getValues(); // D SUBDETALLE
  var productos   = caja.getRange(3, 5, n, 1).getValues(); // E PRODUCTO
  var refRange    = caja.getRange(3, 9, n, 1);
  var referencias = refRange.getValues();                  // I REFERENCIA
  var detRange    = caja.getRange(3, 3, n, 1);
  var detalles    = detRange.getValues();                  // C DETALLE

  var cambiosCO = 0, cambiosSem = 0, envioConfirmados = 0;
  for (var i = 0; i < n; i++) {
    if (referencias[i][0] !== 'CCI22-001') continue;
    var sub  = subdetalles[i][0];
    var prod = productos[i][0];
    if (sub === 'AFIP' || sub === 'Contadores') {
      referencias[i][0] = 'CCI23-001';
      detalles[i][0] = 'CO';
      cambiosCO++;
    } else if (prod === 'Malbec 2023') {
      referencias[i][0] = 'SEM23-001';
      cambiosSem++;
    } else if (sub === 'Envio') {
      envioConfirmados++; // ya está en CCI22-001, sólo lo confirmamos en el conteo
    }
  }

  if (!cambiosCO && !cambiosSem) { Logger.log('No encontré filas para reclasificar — ¿ya se corrió esto antes?'); return; }

  refRange.setValues(referencias);
  detRange.setValues(detalles);

  // Sumar los montos reales de cada grupo ya reclasificado, para las filas de BALANCE
  var montos = caja.getRange(3, 6, n, 1).getValues(); // F MONTO $
  var totalCO = 0, totalSem = 0;
  for (var j = 0; j < n; j++) {
    if (referencias[j][0] === 'CCI23-001' && detalles[j][0] === 'CO') totalCO += (Number(montos[j][0]) || 0);
    if (referencias[j][0] === 'SEM23-001') totalSem += (Number(montos[j][0]) || 0);
  }

  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var hoy = new Date();

  // Fila CO: sin producto, no se prorratea a ningún vino
  var filaCO = obtenerUltimaFilaConFecha(balance, 2);
  balance.insertRowAfter(filaCO);
  filaCO = filaCO + 1;
  balance.getRange(filaCO, 2).setValue(hoy);
  balance.getRange(filaCO, 3).setValue('CO');
  balance.getRange(filaCO, 4).setValue('Contadores/AFIP 2022-2025 (constitución SAS)');
  balance.getRange(filaCO, 7).setValue(totalCO);
  balance.getRange(filaCO, 12).setValue('CCI23-001');
  balance.getRange(filaCO, 13).setValue(0);
  escribirFormulasBalance(filaCO);

  // Fila Semilla: préstamo usado para financiar producción de Malbec 2023
  var filaSem = obtenerUltimaFilaConFecha(balance, 2);
  balance.insertRowAfter(filaSem);
  filaSem = filaSem + 1;
  balance.getRange(filaSem, 2).setValue(hoy);
  balance.getRange(filaSem, 3).setValue('CI');
  balance.getRange(filaSem, 4).setValue('Semilla');
  balance.getRange(filaSem, 5).setValue('Malbec 2023');
  balance.getRange(filaSem, 7).setValue(totalSem);
  balance.getRange(filaSem, 12).setValue('SEM23-001');
  balance.getRange(filaSem, 13).setValue(0);
  escribirFormulasBalance(filaSem);

  Logger.log('CAJA: ' + cambiosCO + ' fila(s) a CO (CCI23-001), ' + cambiosSem + ' a SEM23-001, ' + envioConfirmados + ' de Envío confirmadas en CCI22-001.');
  Logger.log('BALANCE: fila ' + filaCO + ' (CO, $' + totalCO + ') y fila ' + filaSem + ' (Semilla/Malbec 2023, $' + totalSem + ') creadas.');
}

// como si la venta nunca se hubiera cargado. Verifica que la fila 379 sea realmente
// esa venta antes de tocar nada — si no coincide, no borra nada y avisa por Logger.
function deshacerVentaTest() {
  var balance  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var fila     = 379;
  var detalle  = balance.getRange(fila, 3).getValue();  // C DETALLE
  var producto = balance.getRange(fila, 5).getValue();  // E PRODUCTO
  var botellas = balance.getRange(fila, 13).getValue(); // M BOTELLAS

  if (detalle !== 'Venta' || producto !== 'Cabernet Franc 2022' || botellas !== 6) {
    Logger.log('La fila ' + fila + ' no coincide con la venta de prueba esperada ' +
      '(Venta, Cabernet Franc 2022, 6 bot) — encontré: ' + detalle + ', ' + producto + ', ' + botellas + ' bot. ' +
      'No borré nada, revisá la fila a mano.');
    return;
  }

  balance.deleteRow(fila);
  ajustarStockUbicacion('Cab. Franc 2022', null, 'R Peña', 6); // le devuelve las 6 botellas a R Peña
  Logger.log('Listo — fila ' + fila + ' borrada y 6 botellas de Cabernet Franc 2022 devueltas a R Peña.');
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
