// ═══════════════════════════════════════════════════════════════════
//  NOVATO · HERRAMIENTAS DE MANTENIMIENTO
//  Funciones para correr a mano con ▶ Run: diagnósticos, reparaciones y
//  configuración. La app NO las usa. Las de una sola vez que ya se corrieron
//  se borraron (quedan en el historial de GitHub).
//
//  Diagnóstico ............ diagnosticarErrores, diagnosticarZonasHorarias,
//                           diagnosticarCuadroCajas, diagnosticarInsumos,
//                           analizarDiasCierre
//  Reescribir fórmulas ..... reescribirDiasCierre
//  Reparación ............. repararTodasLasFormulas, normalizarRangosAbiertos,
//                           blindarFormulasDolar, repararCuadroCajasArs,
//                           repararFormatosInsumos
//  Configuración .......... instalarTriggerBlueApi, obtenerChatIdsTelegram,
//                           agregarFiltrosEncabezados, ordenarYFormatearColumnasExtra
// ═══════════════════════════════════════════════════════════════════

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
// Lee la hoja INSUMOS para la pantalla de costeo de la app. Devuelve cada insumo con
// su costo en ambas monedas, y los totales de los que están marcados ACTIVO=SI.
// UTILIDAD — repara los formatos de la hoja INSUMOS. Al crearla no se le puso formato
// numérico explícito a la columna B (COSTO x BOT), así que alguna celda pudo quedar con
// formato de fecha heredado y romper el cálculo. Esto fija: B como número (4 decimales,
// porque los costos en USD por botella suelen ser chicos), D como fecha, E/F como número.
// Sólo cambia FORMATOS, nunca valores. Si encuentra un valor que quedó convertido a
// fecha de verdad, lo avisa por Logger para que se vuelva a tipear a mano.
function repararFormatosInsumos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INSUMOS');
  if (!sheet) { Logger.log('No existe la pestaña INSUMOS.'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) { Logger.log('INSUMOS sin datos.'); return; }
  var n = lastRow - 2;

  sheet.getRange(3, 2, n, 1).setNumberFormat('#,##0.0000'); // B COSTO x BOT
  sheet.getRange(3, 4, n, 1).setNumberFormat('dd/mm/yyyy'); // D FECHA PRECIO
  sheet.getRange(3, 5, n, 2).setNumberFormat('#,##0.00');   // E/F convertidos

  // Detectar valores que quedaron convertidos a fecha (el formato no los recupera)
  var valores = sheet.getRange(3, 2, n, 1).getValues();
  var rotos = [];
  for (var i = 0; i < n; i++) {
    if (valores[i][0] instanceof Date) rotos.push('B' + (i + 3));
  }

  SpreadsheetApp.flush();
  if (rotos.length) {
    Logger.log('Formatos corregidos, PERO estas celdas tienen una FECHA guardada en vez de un número: ' +
               rotos.join(', ') + '. Volvé a tipear el precio en esas celdas a mano.');
  } else {
    Logger.log('Listo — formatos de INSUMOS corregidos. Ningún valor quedó convertido a fecha.');
  }
}
// UTILIDAD — diagnóstico de la hoja INSUMOS. Muestra, celda por celda, el valor, el
// TIPO de dato (número, texto, fecha) y el formato aplicado. Sirve para detectar
// celdas que quedaron con formato de fecha y rompen los cálculos.
function diagnosticarInsumos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INSUMOS');
  if (!sheet) { Logger.log('No existe la pestaña INSUMOS.'); return; }
  var lastRow = sheet.getLastRow();
  var valores  = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var formatos = sheet.getRange(2, 1, lastRow - 1, 8).getNumberFormats();
  var formulas = sheet.getRange(2, 1, lastRow - 1, 8).getFormulas();

  var cols = ['A INSUMO', 'B COSTO', 'C MONEDA', 'D FECHA', 'E ARS', 'F USD', 'G PROV', 'H ACTIVO'];
  for (var i = 0; i < valores.length; i++) {
    var fila = i + 2;
    var partes = [];
    for (var c = 0; c < 8; c++) {
      var v = valores[i][c];
      if (v === '' || v === null) continue;
      var tipo = (v instanceof Date) ? 'FECHA' : (typeof v);
      var fmt = formatos[i][c];
      var marca = '';
      // Señalar celdas sospechosas: número con formato de fecha, o fecha donde no va
      if (c === 1 || c === 4 || c === 5) { // B, E, F deberían ser números
        if (v instanceof Date) marca = '  <<< ES UNA FECHA, deberia ser numero';
        else if (fmt && (fmt.indexOf('d') > -1 || fmt.indexOf('y') > -1) && fmt.indexOf('#') === -1) marca = '  <<< FORMATO DE FECHA';
      }
      partes.push(cols[c] + '=' + v + ' [' + tipo + ' fmt:' + fmt + ']' + (formulas[i][c] ? ' F:' + formulas[i][c].substring(0, 40) : '') + marca);
    }
    if (partes.length) Logger.log('fila ' + fila + ' → ' + partes.join('   |   '));
  }
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
        '=IFERROR(IF(G' + row + '=0;"";G' + row + '/XLOOKUP(B' + row + ';BLUE_API!$A$2:$A;BLUE_API!$C$2:$C;;-1));"")',
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
// UTILIDAD — correr UNA VEZ. Las dos filas de BALANCE creadas para CO (Costos SAS,
// ref CCI23-001) y Semilla (Malbec 2023, ref SEM23-001) quedaron con fecha 19/7/2026
// (el día que se crearon), así que su MONTO US$ se valuaba al dólar de hoy. Pero esos
// montos son la suma de decenas de pagos hechos entre 2022 y 2025, a un dólar mucho
// más bajo. Esto les pone la FECHA MEDIANA PONDERADA POR MONTO de sus pagos reales en
// CAJA — la fecha donde se acumuló la mitad de la plata gastada — para que la
// conversión a dólares refleje el tipo de cambio real de la época. Busca las filas
// por su referencia (no por número fijo) y sólo toca esas dos.
// UTILIDAD — correr UNA VEZ. Las fórmulas de BALANCE y CAJA que convierten a dólares
// referencian BLUE_API con un rango de filas FIJO (ej. $A$56:$A$4644). Cuando
// actualizarBlueApi reescribe la tabla, las filas se corren y ese rango queda
// desajustado (se saltea las fechas más recientes → conversión mal). Esto reescribe
// esas referencias a un RANGO ABIERTO ($A$2:$A, $C$2:$C) que nunca se desajusta,
// sin importar cómo cambie BLUE_API. Escribe sólo en las celdas que realmente tienen
// el patrón (nunca sobre un rango completo, para no repetir el borrado accidental).
// UTILIDAD — correr cuando se detecten rangos fijos problemáticos. Las fórmulas que
// referencian OTRA hoja con un rango de filas FIJO (ej. CAJA!$I$3:$I999, o
// BALANCE!$B$3:$B$369, o BLUE_API!$A$56:$A$4644) se rompen cuando esa hoja crece más
// allá del tope, o cuando sus filas se corren. Esto reescribe esos rangos a RANGO
// ABIERTO (ej. CAJA!$I$3:$I), inmune a crecimiento y desplazamiento. Cubre las
// referencias a BLUE_API, CAJA y BALANCE. Escribe sólo en celdas puntuales que
// realmente tienen el patrón (nunca sobre un rango completo).
function normalizarRangosAbiertos() {
  Logger.log('== normalizarRangosAbiertos v2 (BLUE_API + CAJA + BALANCE) ==');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var total = 0;
  var muestras = [];

  // Para cada hoja REFERENCIADA, el patrón que abre su rango. Fila de inicio se
  // respeta ($2 para BLUE_API por el encabezado, $3 para CAJA/BALANCE por sus 2
  // filas de encabezado); sólo se saca el tope fijo del final.
  function abrir(formula) {
    return formula
      // BLUE_API: cualquier fila de inicio → $2 (dedup de encabezado), sin tope
      .replace(/BLUE_API!\$A\$\d+:\$A\$?\d*/g, 'BLUE_API!$A$2:$A')
      .replace(/BLUE_API!\$C\$\d+:\$C\$?\d*/g, 'BLUE_API!$C$2:$C')
      // CAJA y BALANCE: respetar la fila de inicio, sacar sólo el tope
      .replace(/(CAJA!\$[A-Z]+)\$(\d+):\$([A-Z]+)\$?\d+/g, '$1$$$2:$$$3')
      .replace(/(BALANCE!\$[A-Z]+)\$(\d+):\$([A-Z]+)\$?\d+/g, '$1$$$2:$$$3');
  }

  ['BALANCE', 'CAJA', 'STOCK', 'CLIENTES'].forEach(function(nombre) {
    var sheet = ss.getSheetByName(nombre);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;
    var formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
    var diag = false;
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c];
        if (!f) continue;
        // Sólo tocar fórmulas que referencian otra hoja con rango (evita reescribir de más)
        if (f.indexOf('BLUE_API!') === -1 && f.indexOf('CAJA!') === -1 && f.indexOf('BALANCE!') === -1) continue;
        if (!diag) { Logger.log('  [' + nombre + '] primera fórmula cruda: ' + f); diag = true; }
        var nueva = abrir(f);
        if (nueva !== f) {
          sheet.getRange(r + 1, c + 1).setFormula(nueva);
          total++;
          if (muestras.length < 3) muestras.push(nombre + '!' + sheet.getRange(r + 1, c + 1).getA1Notation() + ': ' + f.substring(0, 55) + ' → ' + nueva.substring(0, 55));
        }
      }
    }
  });
  SpreadsheetApp.flush();
  muestras.forEach(function(m) { Logger.log('  ejemplo: ' + m); });
  Logger.log('Listo — ' + total + ' fórmula(s) pasadas a rango abierto (BLUE_API / CAJA / BALANCE).');
}
function blindarFormulasDolar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var total = 0;
  [{ hoja: 'CAJA', col: 7 }, { hoja: 'BALANCE', col: 8 }].forEach(function(o) {
    var sheet = ss.getSheetByName(o.hoja);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    var formulas = sheet.getRange(3, o.col, lastRow - 2, 1).getFormulas();
    for (var i = 0; i < formulas.length; i++) {
      var f = formulas[i][0];
      if (!f || f.indexOf('XLOOKUP') === -1 || f.indexOf('BLUE_API') === -1) continue;
      if (f.indexOf('=IFERROR(') === 0) continue; // ya está blindada
      var nueva = '=IFERROR(' + f.substring(1) + ';"")';
      sheet.getRange(i + 3, o.col).setFormula(nueva); // sólo esta celda puntual
      total++;
    }
  });
  SpreadsheetApp.flush();
  Logger.log('Listo — ' + total + ' fórmula(s) de conversión a dólar blindadas contra #DIV/0!.');
}
// UTILIDAD — diagnóstico del cuadro "CAJAS" del pie de la pestaña CAJA. Muestra, para
// cada celda del bloque, si tiene FÓRMULA (viva, se actualiza sola) o un VALOR pegado
// (congelado, queda desactualizado). Sirve para detectar cuando alguien pegó números
// encima de las fórmulas.
// UTILIDAD — restaura las fórmulas de la columna AR$ del cuadro "CAJAS" del pie de la
// pestaña CAJA. En algún momento se pegaron valores encima de las fórmulas y el cuadro
// dejó de actualizarse (quedó, por ejemplo, LUDICO en 0.0012 — residuo de coma flotante
// de la fórmula original).
//
// IMPORTANTE — sólo toca la columna C (AR$), NUNCA la D (USD):
//   · AR$  = calculado, suma histórica de todos los movimientos de esa caja
//   · USD  = DECLARADO a mano, porque depende del tipo de cambio al que se compraron
//            los dólares, si vinieron de cripto, etc. No es derivable del histórico.
//
// Busca el bloque por texto para no romperse si las filas se corren.
// UTILIDAD — correr UNA VEZ. Prepara el encabezado de la columna S (DEPOSITO) en
// BALANCE, donde a partir de ahora se guarda de qué depósito salieron las botellas
// en cada carga hecha desde la app.
//
// Se agrega AL FINAL (columna S, después de R AÑO) y no insertada al lado de PRODUCTO
// a propósito: insertar una columna en el medio correría F..R, y el código escribe sus
// fórmulas con letras de columna fijas, así que pasaría a escribir en las columnas
// equivocadas. Mover la columna a mano después SÍ es seguro (Sheets ajusta las
// referencias); insertar en el medio, no.
//
// Verifica que S esté libre antes de escribir: si ya hay otra cosa, avisa y no toca nada.
// UTILIDAD — pone filtros en la fila de encabezados de CAJA, BALANCE y STOCK (CLIENTES
// ya los tiene). Si una hoja ya tiene filtro, lo saca y lo vuelve a crear con el rango
// actualizado, así cubre las filas nuevas.
//
// OJO al usarlos: filtrar es inofensivo, pero el menú del filtro también ofrece ORDENAR
// (A-Z, Z-A), y eso reordena las filas FÍSICAMENTE y para todos. En BALANCE eso
// desarmaría el orden cronológico (y la posición en que acomodaste a mano las filas de
// CI/CO). Para explorar sin afectar la hoja, conviene usar Datos → Vistas de filtro,
// que son personales y no cambian el orden real.
function agregarFiltrosEncabezados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Fila de encabezado por hoja (los datos arrancan en la siguiente)
  var config = [
    { hoja: 'BALANCE', filaEncabezado: 2 },
    { hoja: 'CAJA',    filaEncabezado: 2 },
    { hoja: 'STOCK',   filaEncabezado: 2 }
  ];

  config.forEach(function(c) {
    var sheet = ss.getSheetByName(c.hoja);
    if (!sheet) { Logger.log(c.hoja + ': no existe, salteada.'); return; }

    var filtroExistente = sheet.getFilter();
    if (filtroExistente) filtroExistente.remove(); // recrear con el rango actualizado

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= c.filaEncabezado || lastCol < 1) { Logger.log(c.hoja + ': sin datos suficientes.'); return; }

    sheet.getRange(c.filaEncabezado, 1, lastRow - c.filaEncabezado + 1, lastCol).createFilter();
    Logger.log(c.hoja + ': filtro puesto en la fila ' + c.filaEncabezado + ' (hasta fila ' + lastRow + ', ' + lastCol + ' columnas).');
  });

  SpreadsheetApp.flush();
  Logger.log('Listo. Recordá: filtrar es seguro, ORDENAR desde el filtro reordena las filas para todos.');
}
function repararCuadroCajasArs() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  if (!caja) { Logger.log('No hay CAJA.'); return; }
  var lastRow = caja.getLastRow();
  var colB = caja.getRange(1, 2, lastRow, 1).getValues();

  var filaCajas = -1;
  for (var i = 0; i < colB.length; i++) {
    if (colB[i][0] === 'CAJAS') { filaCajas = i + 1; break; }
  }
  if (filaCajas === -1) { Logger.log('No encontré el bloque CAJAS.'); return; }

  // Recorrer desde el encabezado hasta TOTAL:, escribiendo la fórmula de cada caja.
  var primera = null, ultima = null, filaTotal = null;
  for (var f = filaCajas + 1; f <= lastRow; f++) {
    var nombre = colB[f - 1][0];
    if (!nombre) continue;
    if (nombre === 'TOTAL:') { filaTotal = f; break; }
    if (nombre === 'AR$' || nombre === 'USD') continue; // fila de sub-encabezados
    caja.getRange(f, 3).setFormula('=SUMIF(CAJA!$H$3:$H;$B' + f + ';CAJA!$F$3:$F)'); // C AR$
    if (primera === null) primera = f;
    ultima = f;
  }

  if (primera === null) { Logger.log('No encontré filas de cajas para reparar.'); return; }
  if (filaTotal) caja.getRange(filaTotal, 3).setFormula('=SUM(C' + primera + ':C' + ultima + ')');

  SpreadsheetApp.flush();
  Logger.log('Listo — AR$ reparado en filas ' + primera + ' a ' + ultima +
             (filaTotal ? ', total en fila ' + filaTotal : '') + '. La columna USD quedó intacta (es declarada).');
}
function diagnosticarCuadroCajas() {
  var caja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CAJA');
  if (!caja) { Logger.log('No hay CAJA.'); return; }
  var lastRow = caja.getLastRow();
  var colB = caja.getRange(1, 2, lastRow, 1).getValues();

  var inicio = -1;
  for (var i = 0; i < colB.length; i++) {
    if (colB[i][0] === 'CAJAS') { inicio = i + 1; break; } // fila real (1-indexed)
  }
  if (inicio === -1) { Logger.log('No encontré el bloque CAJAS.'); return; }
  Logger.log('Bloque CAJAS empieza en la fila ' + inicio);

  // Mostrar 12 filas desde el encabezado, columnas B..E
  var filas = Math.min(12, lastRow - inicio + 1);
  var formulas = caja.getRange(inicio, 2, filas, 4).getFormulas();
  var valores  = caja.getRange(inicio, 2, filas, 4).getValues();
  for (var r = 0; r < filas; r++) {
    var partes = [];
    for (var c = 0; c < 4; c++) {
      var col = String.fromCharCode(66 + c); // B, C, D, E
      var f = formulas[r][c];
      var v = valores[r][c];
      if (f) partes.push(col + ': [FÓRMULA] ' + f);
      else if (v !== '' && v !== null) partes.push(col + ': [valor] ' + v);
    }
    if (partes.length) Logger.log('  fila ' + (inicio + r) + ' → ' + partes.join('  |  '));
  }
}
function diagnosticarZonasHorarias() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Zona del SCRIPT:   ' + Session.getScriptTimeZone());
  Logger.log('Zona de la PLANILLA: ' + ss.getSpreadsheetTimeZone());

  [['BALANCE', 3], ['BLUE_API', 2]].forEach(function(c) {
    var sh = ss.getSheetByName(c[0]);
    if (!sh) return;
    var desde = c[0] === 'BALANCE' ? Math.max(c[1], obtenerUltimaFilaConFecha(sh, 2) - 4) : c[1];
    var col = c[0] === 'BALANCE' ? 2 : 1;
    var rango = sh.getRange(desde, col, 5, 1);
    var vals = rango.getValues(), vis = rango.getDisplayValues();
    Logger.log('--- ' + c[0] + ' ---');
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (!(v instanceof Date)) continue;
      Logger.log('  fila ' + (desde + i) + ' | UTC ' + v.toISOString() +
                 ' | se ve: ' + vis[i][0] +
                 ' | en Mendoza: ' + Utilities.formatDate(v, 'America/Argentina/Mendoza', 'dd/MM/yyyy HH:mm'));
    }
  });
}
function diagnosticarErrores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ['BALANCE', 'CAJA', 'STOCK', 'CLIENTES', 'BLUE_API', 'DINAMICOS'];
  var errores = ['#REF!', '#ERROR!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '#NULL!'];
  hojas.forEach(function(nombre) {
    var sheet = ss.getSheetByName(nombre);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) { Logger.log(nombre + ': vacía'); return; }
    var valores = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var encontrados = [];
    for (var r = 0; r < valores.length; r++) {
      for (var c = 0; c < valores[r].length; c++) {
        if (errores.indexOf(valores[r][c]) !== -1) {
          encontrados.push(sheet.getRange(r + 1, c + 1).getA1Notation() + '=' + valores[r][c]);
        }
      }
    }
    Logger.log(nombre + ': ' + encontrados.length + ' error(es)' + (encontrados.length ? ' → ' + encontrados.slice(0, 15).join(', ') : ' ✓'));
  });
}

// Reordena las columnas agregadas al final de BALANCE para que queden
// Q (% DIF POR TC) → DIAS CIERRE → AÑO → DEPOSITO, y les copia el formato de la
// columna Q (encabezado, bordes, fuente, ancho) con el formato numérico que
// corresponde a cada una. Se puede correr más de una vez: si ya están en su lugar,
// sólo vuelve a aplicar el formato. El código ubica estas columnas por encabezado,
// así que moverlas no rompe nada.
function ordenarYFormatearColumnasExtra() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  var orden = ['DIAS CIERRE', 'AÑO', 'DEPOSITO'];
  var faltan = orden.filter(function(n) { _colsBalance = null; return !colBalance_(n); });
  if (faltan.length) { Logger.log('No encontré estos encabezados en la fila 2: ' + faltan.join(', ') + '. No toqué nada.'); return; }

  orden.forEach(function(nombre, k) {
    _colsBalance = null;
    var actual = colBalance_(nombre), destino = 18 + k; // R, S, T
    if (actual !== destino) {
      // moveColumns: el índice de destino es la columna ANTES de la cual se inserta
      sh.moveColumns(sh.getRange(1, actual, sh.getMaxRows(), 1), destino > actual ? destino + 1 : destino);
      Logger.log(nombre + ': movida de la columna ' + actual + ' a la ' + destino);
    }
  });
  _colsBalance = null;

  var filas = sh.getMaxRows();
  var molde = sh.getRange(1, 17, filas, 1); // columna Q
  var formatos = { 'DIAS CIERRE': '0', 'AÑO': '0', 'DEPOSITO': '@' };
  orden.forEach(function(nombre) {
    var c = colBalance_(nombre);
    molde.copyTo(sh.getRange(1, c, filas, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(3, c, filas - 2, 1).setNumberFormat(formatos[nombre]);
    sh.setColumnWidth(c, sh.getColumnWidth(17));
  });
  SpreadsheetApp.flush();
  Logger.log('Listo — orden: Q % DIF POR TC | R DIAS CIERRE | S AÑO | T DEPOSITO, con el formato de la tabla.');
}

// Analiza los valores de DIAS CIERRE en BALANCE: distribución general y, para cada
// valor raro (negativo, cero o más de 365 días), el contexto para entender por qué:
// fechas de la operación y de sus cobros/pagos en CAJA, y si la fórmula de saldo de
// esa fila es la vieja (sólo por referencia) o la nueva (referencia + producto).
// No modifica nada.
function analizarDiasCierre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bal = ss.getSheetByName('BALANCE'), caja = ss.getSheetByName('CAJA');
  _colsBalance = null;
  var cDias = colBalance_('DIAS CIERRE');
  if (!cDias) { Logger.log('No hay columna DIAS CIERRE.'); return; }
  var ultB = obtenerUltimaFilaConFecha(bal, 2), ultC = obtenerUltimaFilaConFecha(caja, 2);
  var B = bal.getRange(3, 1, ultB - 2, Math.max(cDias, 15)).getValues();
  var O = bal.getRange(3, 15, ultB - 2, 1).getFormulas();
  var C = caja.getRange(3, 1, ultC - 2, 9).getValues(); // A..I

  // Índice de CAJA por referencia
  var porRef = {};
  C.forEach(function(r) {
    if (!r[8] || !(r[1] instanceof Date)) return;
    (porRef[r[8]] = porRef[r[8]] || []).push({ fecha: r[1], prod: r[4] || '' });
  });
  var f = function(d) { return Utilities.formatDate(d, 'America/Argentina/Mendoza', 'dd/MM/yy'); };

  var vals = [], raros = [];
  for (var i = 0; i < B.length; i++) {
    var d = B[i][cDias - 1];
    if (typeof d !== 'number') continue;
    vals.push(d);
    if (d < 0 || d > 365 || d % 1 !== 0) raros.push(i); // los ceros son normales (contado)
  }
  vals.sort(function(a, b) { return a - b; });
  var med = vals.length ? vals[Math.floor(vals.length / 2)] : null;
  Logger.log('Operaciones con DIAS CIERRE: ' + vals.length + ' | min ' + vals[0] + ' | mediana ' + med + ' | max ' + vals[vals.length - 1]);
  Logger.log('Negativos: ' + vals.filter(function(v) { return v < 0; }).length +
             ' | Cero: ' + vals.filter(function(v) { return v === 0; }).length +
             ' | 1-60: ' + vals.filter(function(v) { return v > 0 && v <= 60; }).length +
             ' | 61-365: ' + vals.filter(function(v) { return v > 60 && v <= 365; }).length +
             ' | >365: ' + vals.filter(function(v) { return v > 365; }).length);
  Logger.log('--- Para revisar (negativos, >365 o con decimales) ---');
  raros.forEach(function(i) {
    var r = B[i], ref = r[11], prod = r[4] || '', movs = porRef[ref] || [];
    var mismoProd = movs.filter(function(m) { return m.prod === prod; });
    var fechas = movs.map(function(m) { return m.fecha; }).sort(function(a, b) { return a - b; });
    Logger.log('fila ' + (i + 3) + ' | ' + r[2] + ' · ' + (r[3] || '') + ' · ' + (prod || '(sin producto)') +
      ' | ref ' + ref + ' | fecha op ' + (r[1] instanceof Date ? f(r[1]) : r[1]) + ' | DIAS ' + r[cDias - 1] +
      ' | saldo: ' + (String(O[i][0]).indexOf('SUMIFS') > -1 ? 'ref+producto' : 'solo ref') +
      ' | CAJA: ' + movs.length + ' mov (' + mismoProd.length + ' con este producto)' +
      (fechas.length ? ', del ' + f(fechas[0]) + ' al ' + f(fechas[fechas.length - 1]) : ''));
  });
}

// Reescribe la fórmula de DIAS CIERRE en todas las filas de BALANCE que tienen
// REFERENCIA, con la versión vigente de formulaDiasCierre(). Usar cuando cambia la
// fórmula. Escribe celda por celda y sólo en filas con referencia; no toca otras
// columnas. Se puede correr las veces que haga falta.
function reescribirDiasCierre() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BALANCE');
  _colsBalance = null;
  var c = colBalance_('DIAS CIERRE');
  if (!c) { Logger.log('No hay columna DIAS CIERRE.'); return; }
  var ult = obtenerUltimaFilaConFecha(sh, 2);
  var refs = sh.getRange(3, 12, ult - 2, 1).getValues(); // L REFERENCIA
  var n = 0;
  for (var i = 0; i < refs.length; i++) {
    if (!refs[i][0]) continue;
    sh.getRange(i + 3, c).setFormula(formulaDiasCierre(i + 3));
    n++;
  }
  SpreadsheetApp.flush();
  Logger.log('Listo — fórmula de DIAS CIERRE reescrita en ' + n + ' fila(s). Corré analizarDiasCierre() para revisar.');
}
