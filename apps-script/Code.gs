// ═══════════════════════════════════════════════════════════════════
//  NOVATO BODEGA · Google Apps Script Backend
//  Pegar TODO este código en Extensions → Apps Script del Google Sheet
//  Luego: Deploy → Manage deployments → editar el deployment existente
//    → Version: New version → Deploy  (mantiene la misma URL)
// ═══════════════════════════════════════════════════════════════════

// Rango "ancho" para que las fórmulas de reconciliación (BALANCE ↔ CAJA)
// sigan funcionando a medida que se agreguen filas nuevas con el tiempo.
var CAJA_RANGO_FIN = 5000;

// Historial de transferencias entre depósitos (NO es la fuente de verdad
// del stock — eso vive directo en STOCK. Esto es sólo un log de auditoría).
var TAB_TRANSF = 'APP_TRANSFERENCIAS';

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
    ok:               true,
    lastPrice:        getLastSalePrice(),
    ops:              getRecentOps(20),
    stockUbicacion:   getStockUbicacion(),
    ventasPendientes: getVentasPendientes(),
    clientes:         getClientes()
  };
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

// Registra la venta en BALANCE (con las mismas fórmulas que usa cualquier fila
// cargada a mano) y descuenta el depósito de origen en STOCK. NO toca CAJA:
// eso sólo pasa cuando se registre el cobro correspondiente.
function addSale(p) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var balance = ss.getSheetByName('BALANCE');
  if (!balance) throw new Error('No se encontró la pestaña BALANCE');

  var productoSheet  = PRODUCTO_APP_TO_SHEET[p.producto] || p.producto;
  var rate           = getDolarRate(p.fecha);
  var montoIngresado = parseFloat(p.monto) || 0;
  // MONTO $ en BALANCE siempre es en pesos; si la venta se cargó en USD, la pasamos
  // a pesos con la cotización del día para que las fórmulas de abajo sean consistentes.
  var montoArs = (p.moneda === 'USD' && rate) ? Math.round(montoIngresado * rate) : montoIngresado;

  var cliente    = p.cliente || 'Particular';
  var referencia = nextReferencia(prefijoCliente(cliente));
  var row        = balance.getLastRow() + 1;

  balance.getRange(row, 2, 1, 13).setValues([[
    parseFechaApp(p.fecha),                                                       // B FECHA
    'Venta',                                                                      // C DETALLE
    cliente,                                                                      // D SUBDETALLE
    productoSheet,                                                                // E PRODUCTO
    '=RIGHT(E' + row + ',4)',                                                     // F AÑADA
    montoArs,                                                                     // G MONTO $
    '=G' + row + '/XLOOKUP(B' + row + ',BLUE_API!$A$2:$A$4590,BLUE_API!$C$2:$C$4590,,-1)', // H MONTO US$ FF
    '=H' + row + '+P' + row,                                                      // I MONTO US$ FP
    '=G' + row + '/VLOOKUP(E' + row + ',STOCK!$B$3:$G$9,3,FALSE)',                // J CU $
    '=H' + row + '/VLOOKUP(E' + row + ',STOCK!$B$3:$G$9,3,FALSE)',                // K CU US$
    referencia,                                                                   // L REFERENCIA
    parseInt(p.botellas) || 0,                                                    // M BOTELLAS
    '=IF(G' + row + '>0,"Ingreso",(IF(G' + row + '=0,"Movimiento","Egreso")))'    // N CONCEPTO
  ]]);
  balance.getRange(row, 15, 1, 3).setValues([[
    '=G' + row + '-SUMIF(CAJA!$I$3:$I$' + CAJA_RANGO_FIN + ',L' + row + ',CAJA!$F$3:$F$' + CAJA_RANGO_FIN + ')',
    '=-(H' + row + '-SUMIF(CAJA!$I$3:$I$' + CAJA_RANGO_FIN + ',L' + row + ',CAJA!$G$3:$G$' + CAJA_RANGO_FIN + '))',
    '=IF(N' + row + '="Egreso",-P' + row + '/H' + row + ',P' + row + '/H' + row + ')'
  ]]);
  balance.getRange(row, 1).setValue(p.user || ''); // A: quién lo cargó (columna sin uso hasta ahora)

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

  caja.appendRow([
    p.user || '',                              // A: quién lo cargó (columna sin uso hasta ahora)
    parseFechaApp(p.fecha),                    // B FECHA
    p.tipo === 'cobro' ? 'Cobro' : 'Gasto',    // C DETALLE
    p.concepto || '',                          // D SUBDETALLE
    '',                                        // E PRODUCTO
    monto,                                     // F MONTO $
    montoUS,                                   // G MONTO US$
    cajaLbl,                                   // H CAJA
    p.referencia || ''                         // I REFERENCIA (opcional → venta de BALANCE)
  ]);
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
      if (!rb[0] || rb[2] !== 'Venta') continue; // A usuario, C DETALLE
      ops.push({
        id:    'b_' + i,
        icon:  '🍾',
        desc:  rb[12] + ' bot ' + rb[4] + ' → ' + rb[3], // M botellas, E producto, D cliente
        monto: '$' + Math.round(rb[6]).toLocaleString(), // G monto $
        fecha: formatDate(rb[1]),
        user:  rb[0],
        ts:    i
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
        ts:    j
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
// Ensancha el límite de fila (511 → 5000) que usan las fórmulas de reconciliación
// existentes en BALANCE (columnas O y P) para mirar de vuelta a CAJA. Sin esto,
// un cobro cargado hoy contra una venta VIEJA (anterior a la app) no se reflejaría
// en el saldo de esa venta vieja, porque su fórmula sólo mira CAJA hasta la fila 511
// (y CAJA ya tiene 524 filas reales). No toca ninguna otra cosa.
function ampliarRangosReconciliacion() {
  var balance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var lastRow = balance.getLastRow();
  if (lastRow < 3) { Logger.log('BALANCE está vacío, nada para ampliar.'); return; }
  var range     = balance.getRange(3, 15, lastRow - 2, 2); // O:P
  var formulas  = range.getFormulas();
  var patronFin = /\$511\b/g;
  var cambios   = 0;
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] && patronFin.test(formulas[i][j])) {
        formulas[i][j] = formulas[i][j].replace(patronFin, '$' + CAJA_RANGO_FIN);
        cambios++;
      }
      patronFin.lastIndex = 0;
    }
  }
  range.setFormulas(formulas);
  Logger.log('Listo — ' + cambios + ' fórmulas ampliadas a fila ' + CAJA_RANGO_FIN + '.');
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
