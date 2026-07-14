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
    else if (action === 'addSale')       result = addSale(e.parameter);
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
    clientes:           getClientes(),
    ultimoControlStock: getUltimoControlStock(),
    resumenCajas:       getResumenCajas()
  };
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
// filas nuevas (addSale) como para reparar filas rotas (repararBalance). IMPORTANTE:
// este Sheet usa configuración regional en español → los argumentos de función van
// separados por PUNTO Y COMA (;), no coma. Escribir con comas produce #ERROR!.
function escribirFormulasBalance(row) {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  balance.getRange(row, 6).setValue('=RIGHT(E' + row + ';4)'); // F AÑADA
  balance.getRange(row, 8, 1, 4).setValues([[
    '=G' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A$4590;BLUE_API!$C$2:$C$4590;;-1)', // H MONTO US$ FF
    '=H' + row + '+P' + row,                                                               // I MONTO US$ FP
    '=G' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)',                         // J CU $
    '=H' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)'                          // K CU US$
  ]]);
  balance.getRange(row, 14).setValue('=IF(G' + row + '>0;"Ingreso";(IF(G' + row + '=0;"Movimiento";"Egreso")))'); // N CONCEPTO
  balance.getRange(row, 15, 1, 4).setValues([[
    '=G' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$F$3:$F)',                         // O SALDO $ (rango abierto)
    '=-(H' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$G$3:$G))',                      // P SALDO US$
    '=IF(N' + row + '="Egreso";-P' + row + '/H' + row + ';P' + row + '/H' + row + ')',    // Q % DIF POR TC
    '=IF(BALANCE!$B' + row + '="";"";YEAR(BALANCE!$B' + row + '))'                         // R AÑO
  ]]);
}

// Registra la venta en BALANCE (con las mismas fórmulas que usa cualquier fila
// cargada a mano) y descuenta el depósito de origen en STOCK. NO toca CAJA:
// eso sólo pasa cuando se registre el cobro correspondiente.
function addSale(p) {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) throw new Error('No se encontró la pestaña BALANCE');

  var productoSheet  = PRODUCTO_APP_TO_SHEET[p.producto] || p.producto;
  var rate           = getDolarRate(p.fecha);
  var montoIngresado = parseFloat(p.monto) || 0;
  // MONTO $ en BALANCE siempre es en pesos; si la venta se cargó en USD, la pasamos
  // a pesos con la cotización del día para que las fórmulas de abajo sean consistentes.
  var montoArs = (p.moneda === 'USD' && rate) ? Math.round(montoIngresado * rate) : montoIngresado;

  var cliente    = p.cliente || 'Particular';
  var referencia = nextReferencia(prefijoCliente(cliente));
  var ultimaReal = obtenerUltimaFilaConFecha(balance, 2); // B = FECHA
  balance.insertRowAfter(ultimaReal);
  var row = ultimaReal + 1;

  // Datos que vienen del usuario
  balance.getRange(row, 2).setValue(parseFechaApp(p.fecha));    // B FECHA
  balance.getRange(row, 3).setValue('Venta');                   // C DETALLE
  balance.getRange(row, 4).setValue(cliente);                   // D SUBDETALLE
  balance.getRange(row, 5).setValue(productoSheet);              // E PRODUCTO
  balance.getRange(row, 7).setValue(montoArs);                   // G MONTO $
  balance.getRange(row, 12).setValue(referencia);                // L REFERENCIA
  balance.getRange(row, 13).setValue(parseInt(p.botellas) || 0); // M BOTELLAS
  balance.getRange(row, 1).setValue(p.user || '');               // A: quién lo cargó (columna sin uso hasta ahora)

  // Columnas calculadas — mismas fórmulas que cualquier fila cargada a mano
  escribirFormulasBalance(row);

  // Descuenta del depósito de origen (por defecto R Peña si no se especificó)
  ajustarStockUbicacion(p.producto, p.deposito || 'R Peña', null, parseInt(p.botellas) || 0);

  return { ok: true, referencia: referencia };
}

// Ventas con saldo pendiente de cobro — para que la pantalla de Caja pueda
// vincular un cobro a la venta correspondiente y la reconciliación se cierre sola.
function getVentasPendientes() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  if (!balance) return [];
  var lastRow = balance.getLastRow();
  if (lastRow < 3) return [];
  var data = balance.getRange(3, 2, lastRow - 2, 15).getValues(); // B..P
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (r[1] !== 'Venta') continue;      // C DETALLE
    var saldoArs = r[13];                // O SALDO $
    if (!saldoArs) continue;
    out.push({
      referencia: r[10],                 // L
      cliente:    r[2],                  // D
      producto:   r[3],                  // E
      saldoArs:   Math.round(saldoArs),
      saldoUsd:   Math.round(r[14] || 0) // P
    });
  }
  return out;
}

// ── MOVIMIENTOS DE CAJA → CAJA ────────────────────────────────────────
// p.referencia es opcional: si viene, vincula el cobro/pago con una venta de
// BALANCE (misma REFERENCIA) para que su SALDO se actualice solo.
function addMovement(p) {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  if (!caja) throw new Error('No se encontró la pestaña CAJA');

  var signo   = p.tipo === 'gasto' ? -1 : 1;
  var monto   = (parseFloat(p.monto) || 0) * signo;
  var rate    = getDolarRate(p.fecha);
  var montoUS = rate ? Math.round((monto / rate) * 100) / 100 : '';
  var cajaLbl = CAJA_LABELS[p.caja] || p.caja;

  var ultimaReal = obtenerUltimaFilaConFecha(caja, 2); // B = FECHA
  caja.insertRowAfter(ultimaReal);
  var row = ultimaReal + 1;
  caja.getRange(row, 1, 1, 9).setValues([[
    p.user || '',                              // A: quién lo cargó (columna sin uso hasta ahora)
    parseFechaApp(p.fecha),                    // B FECHA
    p.tipo === 'cobro' ? 'Cobro' : 'Gasto',    // C DETALLE
    p.concepto || '',                          // D SUBDETALLE
    '',                                        // E PRODUCTO
    monto,                                     // F MONTO $
    montoUS,                                   // G MONTO US$
    cajaLbl,                                   // H CAJA
    p.referencia || ''                         // I REFERENCIA (opcional → venta de BALANCE)
  ]]);
  return { ok: true };
}

// ── TRANSFERENCIAS ENTRE DEPÓSITOS → STOCK ───────────────────────────
function addTransfer(p) {
  var cantidad = parseInt(p.cantidad) || 0;
  ajustarStockUbicacion(p.producto, p.desde, p.hacia, cantidad);

  var headers = ['FECHA','PRODUCTO','CANTIDAD','DESDE','HACIA','NOTAS','USUARIO','REGISTRADO'];
  var log = getOrCreateSheet(TAB_TRANSF, headers);
  log.appendRow([p.fecha, p.producto, cantidad, p.desde, p.hacia, p.notas || '', p.user, new Date()]);
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
      fecha = row[0]; hora = row[1]; deposito = row[3];
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
// que addSale, más la fila de TOTALES (con una fórmula basada en ROW() que no
// depende de ningún número de fila fijo, así nunca vuelve a quedar corta ni corre
// riesgo de auto-referenciarse) y el EGRESOS agregado de STOCK (rango abierto real
// hacia BALANCE). Es un arreglo definitivo — no perjudica nada si se corre de nuevo.
function repararTodasLasFormulas() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var lastRow = balance.getLastRow();
  if (lastRow < 3) { Logger.log('BALANCE está vacío, nada para reparar.'); return; }

  var n = lastRow - 2;
  var fechas    = balance.getRange(3, 2, n, 1).getValues();  // B FECHA — 1 sola llamada para todo el rango
  var conceptos = balance.getRange(3, 14, n, 1).getValues(); // N — para ubicar la fila 'TOTALES:'

  var filaTotales = -1;
  var filas = [];
  for (var i = 0; i < n; i++) {
    var row = i + 3;
    if (conceptos[i][0] === 'TOTALES:') { filaTotales = row; continue; }
    if (fechas[i][0] instanceof Date) filas.push(row);
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
        '=G' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A$4590;BLUE_API!$C$2:$C$4590;;-1)',
        '=H' + row + '+P' + row,
        '=G' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)',
        '=H' + row + '/VLOOKUP(E' + row + ';STOCK!$B$3:$G$9;3;FALSE)'
      ]);
      colN.push(['=IF(G' + row + '>0;"Ingreso";(IF(G' + row + '=0;"Movimiento";"Egreso")))']);
      colOR.push([
        '=G' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$F$3:$F)',
        '=-(H' + row + '-SUMIF(CAJA!$I$3:$I;L' + row + ';CAJA!$G$3:$G))',
        '=IF(N' + row + '="Egreso";-P' + row + '/H' + row + ';P' + row + '/H' + row + ')',
        '=IF(BALANCE!$B' + row + '="";"";YEAR(BALANCE!$B' + row + '))'
      ]);
    });
    balance.getRange(inicio, 6, grupo.length, 1).setValues(colF);
    balance.getRange(inicio, 8, grupo.length, 4).setValues(colHK);
    balance.getRange(inicio, 14, grupo.length, 1).setValues(colN);
    balance.getRange(inicio, 15, grupo.length, 4).setValues(colOR);
  });

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
