// ============================================
// CONGELAMIENTO DE MESES - v7
// ============================================
// 'panel' = PANEL DE CONTROL (formula = fuente - acumulado)
// 'importrange' = Solapa propia (auto-detecta filas y columnas)
// ============================================

var CONFIG = {
  solapaDestino: 'PANEL DE CONTROL',
  vendedores: [
    {
      nombre: 'TECNO', tipo: 'panel', columna: 'F',
      filas2026: [8,12,16,20,24,28,32,36,40,44,48,52],
      filas2027: [62,66,70,74,78,82,86,90,94,98,102,106],
      fuentes: {
        pesos: { solapa: 'VENTAS TECNO EN PESOS', celda: 'AM70', ref: "'VENTAS TECNO EN PESOS'!AM70" },
        dolares: { solapa: 'VENTA TECNO EN DOLARES', celda: 'AQ284', ref: "'VENTA TECNO EN DOLARES'!AQ284" }
      }
    },
    {
      nombre: 'CREDITOS', tipo: 'panel', columna: 'J',
      filas2026: [8,12,16,20,24,28,32,36,40,44,48,52],
      filas2027: [62,66,70,74,78,82,86,90,94,98,102,106],
      fuentes: {
        pesos: { solapa: 'VENTA CREDITOS EN PESOS', celda: 'AL68', ref: "'VENTA CREDITOS EN PESOS'!AL68" },
        dolares: { solapa: 'VENTA CREDITOS EN DOLARES', celda: 'AL66', ref: "'VENTA CREDITOS EN DOLARES'!AL66" }
      }
    },
    { nombre: 'TINO', tipo: 'importrange', solapa: 'TINO' },
    { nombre: 'OSITO S.R.L.', tipo: 'importrange', solapa: 'OSITO S.R.L.' },
    { nombre: 'PATITO S.A.', tipo: 'importrange', solapa: 'PATITO S.A.' },
    { nombre: 'GONZA', tipo: 'importrange', solapa: 'GONZA' },
    { nombre: 'MELY', tipo: 'importrange', solapa: 'MELY' },
    { nombre: 'LINEA 314', tipo: 'importrange', solapa: 'LINEA 314' },
    { nombre: 'TOBIAS', tipo: 'importrange', solapa: 'TOBIAS' }
  ]
};

var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

// --- AUXILIARES ---

function armarFormula_(ref, acum) {
  if (acum === 0) return '=' + ref;
  return acum > 0 ? '=' + ref + '-' + acum : '=' + ref + '+' + Math.abs(acum);
}

function extraerAcumulado_(formula) {
  if (!formula) return 0;
  var m = formula.match(/-\s*([\d.]+)\s*$/);
  return m ? parseFloat(m[1]) : 0;
}

function sumarCongelado_(destSheet, col, vend, hastaAnio, hastaMes) {
  var acumP = 0, acumD = 0;
  var anios = [[2026, vend.filas2026], [2027, vend.filas2027]];
  for (var a = 0; a < anios.length; a++) {
    var anio = anios[a][0], filas = anios[a][1];
    if (!filas || anio > hastaAnio) continue;
    var tope = (anio < hastaAnio) ? 12 : hastaMes;
    for (var m = 0; m < tope; m++) {
      var vP = destSheet.getRange(col + filas[m]).getValue();
      var vD = destSheet.getRange(col + (filas[m] + 1)).getValue();
      if (typeof vP === 'number') acumP += vP;
      if (typeof vD === 'number') acumD += vD;
    }
  }
  return { pesos: acumP, dolares: acumD };
}

// --- AUTO-DETECCION PARA IMPORTRANGE ---

function detectarGrupos_(sheet) {
  var nombres = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
                 'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  var lastRow = sheet.getLastRow();
  if (lastRow < 12) return [];
  var colA = sheet.getRange('A1:A' + lastRow).getValues();
  var rows = [];
  for (var i = 0; i < colA.length; i++) {
    var val = String(colA[i][0]).trim().toUpperCase();
    if (nombres.indexOf(val) >= 0) rows.push(i + 1);
  }
  var grupos = [], grupo = [];
  for (var i = 0; i < rows.length; i++) {
    if (grupo.length > 0 && rows[i] - grupo[grupo.length - 1] > 3) {
      if (grupo.length === 12) grupos.push(grupo);
      grupo = [];
    }
    grupo.push(rows[i]);
  }
  if (grupo.length === 12) grupos.push(grupo);
  return grupos;
}

function detectarAnio_(sheet, grupo) {
  var startRow = Math.max(1, grupo[0] - 5);
  var numRows = grupo[0] - startRow;
  if (numRows <= 0) return null;
  var lastCol = Math.min(sheet.getLastColumn(), 20);
  var values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  for (var r = values.length - 1; r >= 0; r--) {
    for (var c = 0; c < values[r].length; c++) {
      var val = String(values[r][c]);
      if (val.indexOf('2027') >= 0) return 2027;
      if (val.indexOf('2026') >= 0) return 2026;
    }
  }
  return null;
}

function congelarFila_(sheet, row, lastCol) {
  var range = sheet.getRange(row, 2, 1, lastCol - 1);
  var values = range.getValues()[0];
  var formulas = range.getFormulas()[0];
  var count = 0;
  for (var c = 0; c < values.length; c++) {
    if (formulas[c] || (values[c] !== '' && values[c] !== null)) {
      sheet.getRange(row, c + 2).setValue(values[c]);
      count++;
    }
  }
  return count;
}

function infoFila_(sheet, row, lastCol) {
  var range = sheet.getRange(row, 2, 1, lastCol - 1);
  var values = range.getValues()[0];
  var formulas = range.getFormulas()[0];
  var conFormula = 0, conValor = 0;
  for (var c = 0; c < values.length; c++) {
    if (formulas[c]) conFormula++;
    else if (values[c] !== '' && values[c] !== null && values[c] !== 0) conValor++;
  }
  return { formulas: conFormula, valores: conValor };
}

// --- CONGELAR MES ANTERIOR ---
function congelarMes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var mesC = mesActual - 1, anioC = anioActual;
  if (mesC < 0) { mesC = 11; anioC--; }
  var log = ['CONGELAMIENTO DE MES',
    'Fecha: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Congelando: ' + MESES[mesC] + ' ' + anioC, ''];

  CONFIG.vendedores.forEach(function(vend) {
    log.push('=== ' + vend.nombre + ' ===');

    if (vend.tipo === 'importrange') {
      var sheet = ss.getSheetByName(vend.solapa);
      if (!sheet) { log.push('  SOLAPA "' + vend.solapa + '" NO ENCONTRADA'); return; }
      var grupos = detectarGrupos_(sheet);
      var lastCol = sheet.getLastColumn();
      var found = false;
      for (var g = 0; g < grupos.length; g++) {
        var anio = detectarAnio_(sheet, grupos[g]);
        if (anio !== anioC) continue;
        found = true;
        var row = grupos[g][mesC];
        var count = congelarFila_(sheet, row, lastCol);
        log.push('  Fila ' + row + ': ' + count + ' celdas congeladas');
      }
      if (!found) log.push('  Sin datos para ' + anioC);

    } else {
      var destSheet = ss.getSheetByName(CONFIG.solapaDestino);
      if (!destSheet) { log.push('  PANEL DE CONTROL no encontrado'); return; }
      var filasC = anioC === 2026 ? vend.filas2026 : (anioC === 2027 ? vend.filas2027 : null);
      var filasN = anioActual === 2026 ? vend.filas2026 : (anioActual === 2027 ? vend.filas2027 : null);
      if (!filasC || !filasN) { log.push('  Sin config para ' + anioC); return; }
      var fc = filasC[mesC], fn = filasN[mesActual], col = vend.columna;
      var rP = destSheet.getRange(col + fc), rD = destSheet.getRange(col + (fc + 1));
      var fP = rP.getFormula(), fD = rD.getFormula();
      var vP = rP.getValue(), vD = rD.getValue();
      var prevP = extraerAcumulado_(fP), prevD = extraerAcumulado_(fD);
      if (fP) { rP.setValue(vP); log.push('  PESOS congelado: ' + col + fc + ' = ' + vP); }
      else log.push('  PESOS ya fijo: ' + col + fc + ' = ' + vP);
      if (fD) { rD.setValue(vD); log.push('  DOLARES congelado: ' + col + (fc+1) + ' = ' + vD); }
      else log.push('  DOLARES ya fijo: ' + col + (fc+1) + ' = ' + vD);
      var newP = prevP + (typeof vP === 'number' ? vP : 0);
      var newD = prevD + (typeof vD === 'number' ? vD : 0);
      log.push('  Nuevo acumulado: P=$' + newP + ' D=U$D' + newD);
      destSheet.getRange(col + fn).setFormula(armarFormula_(vend.fuentes.pesos.ref, newP));
      destSheet.getRange(col + (fn+1)).setFormula(armarFormula_(vend.fuentes.dolares.ref, newD));
      log.push('  Nuevo: ' + col + fn + ' = ' + armarFormula_(vend.fuentes.pesos.ref, newP));
      log.push('  Nuevo: ' + col + (fn+1) + ' = ' + armarFormula_(vend.fuentes.dolares.ref, newD));
    }
    log.push('');
  });
  var msg = log.join('\n'); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Congelamiento', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

// --- VISTA PREVIA ---
function vistaPrevia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var mesC = mesActual - 1, anioC = anioActual;
  if (mesC < 0) { mesC = 11; anioC--; }
  var log = ['VISTA PREVIA - NO modifica nada',
    'Congelaria: ' + MESES[mesC] + ' ' + anioC, ''];

  CONFIG.vendedores.forEach(function(vend) {
    log.push('=== ' + vend.nombre + ' ===');

    if (vend.tipo === 'importrange') {
      var sheet = ss.getSheetByName(vend.solapa);
      if (!sheet) { log.push('  SOLAPA "' + vend.solapa + '" NO ENCONTRADA'); return; }
      var grupos = detectarGrupos_(sheet);
      var lastCol = sheet.getLastColumn();
      log.push('  Solapa: ' + vend.solapa + ' (' + grupos.length + ' bloques detectados)');
      for (var g = 0; g < grupos.length; g++) {
        var anio = detectarAnio_(sheet, grupos[g]);
        if (anio !== anioC) continue;
        var row = grupos[g][mesC];
        var info = infoFila_(sheet, row, lastCol);
        log.push('  ' + MESES[mesC] + ' (fila ' + row + '): ' + info.formulas + ' con formula, ' + info.valores + ' con valor fijo');
      }

    } else {
      var destSheet = ss.getSheetByName(CONFIG.solapaDestino);
      if (!destSheet) { log.push('  PANEL DE CONTROL no encontrado'); return; }
      var filasC = anioC === 2026 ? vend.filas2026 : (anioC === 2027 ? vend.filas2027 : null);
      var filasN = anioActual === 2026 ? vend.filas2026 : (anioActual === 2027 ? vend.filas2027 : null);
      if (!filasC || !filasN) { log.push('  Sin config'); return; }
      var fc = filasC[mesC], fn = filasN[mesActual], col = vend.columna;
      var rP = destSheet.getRange(col + fc), rD = destSheet.getRange(col + (fc + 1));
      log.push('  ' + MESES[mesC] + ':');
      log.push('    PESOS ' + col + fc + ' = ' + rP.getValue() + (rP.getFormula() ? ' (formula)' : ' (fijo)'));
      log.push('    DOLARES ' + col + (fc+1) + ' = ' + rD.getValue() + (rD.getFormula() ? ' (formula)' : ' (fijo)'));
      var rNP = destSheet.getRange(col + fn), rND = destSheet.getRange(col + (fn + 1));
      log.push('  ' + MESES[mesActual] + ':');
      log.push('    PESOS ' + col + fn + ' = ' + rNP.getValue() + (rNP.getFormula() ? ' (' + rNP.getFormula() + ')' : ' (vacio/fijo)'));
      log.push('    DOLARES ' + col + (fn+1) + ' = ' + rND.getValue() + (rND.getFormula() ? ' (' + rND.getFormula() + ')' : ' (vacio/fijo)'));
    }
    log.push('');
  });
  SpreadsheetApp.getUi().alert('Vista Previa', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- CONGELAR MES ESPECIFICO ---
function congelarMesEspecifico() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Congelar Mes', 'Mes y anio (ej: 3 2026 para Marzo 2026):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var p = resp.getResponseText().trim().split(/\s+/);
  if (p.length < 2) { ui.alert('Formato: numero_mes anio (ej: 3 2026)'); return; }
  var mes = parseInt(p[0]) - 1, anio = parseInt(p[1]);
  if (isNaN(mes) || mes < 0 || mes > 11 || isNaN(anio)) { ui.alert('Invalido'); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ['CONGELAMIENTO: ' + MESES[mes] + ' ' + anio, ''];

  CONFIG.vendedores.forEach(function(vend) {
    log.push('=== ' + vend.nombre + ' ===');

    if (vend.tipo === 'importrange') {
      var sheet = ss.getSheetByName(vend.solapa);
      if (!sheet) { log.push('  SOLAPA "' + vend.solapa + '" NO ENCONTRADA'); return; }
      var grupos = detectarGrupos_(sheet);
      var lastCol = sheet.getLastColumn();
      var found = false;
      for (var g = 0; g < grupos.length; g++) {
        var anioG = detectarAnio_(sheet, grupos[g]);
        if (anioG !== anio) continue;
        found = true;
        var row = grupos[g][mes];
        var count = congelarFila_(sheet, row, lastCol);
        log.push('  Fila ' + row + ': ' + count + ' celdas congeladas');
      }
      if (!found) log.push('  Sin datos para ' + anio);

    } else {
      var dest = ss.getSheetByName(CONFIG.solapaDestino);
      if (!dest) { log.push('  PANEL DE CONTROL no encontrado'); return; }
      var filas = anio === 2026 ? vend.filas2026 : (anio === 2027 ? vend.filas2027 : null);
      if (!filas) { log.push('  Sin config para ' + anio); return; }
      var f = filas[mes], col = vend.columna;
      var rP = dest.getRange(col + f), rD = dest.getRange(col + (f + 1));
      if (rP.getFormula()) { rP.setValue(rP.getValue()); log.push('  PESOS: ' + col + f + ' congelado = ' + rP.getValue()); }
      else log.push('  PESOS: ya fijo');
      if (rD.getFormula()) { rD.setValue(rD.getValue()); log.push('  DOLARES: ' + col + (f+1) + ' congelado = ' + rD.getValue()); }
      else log.push('  DOLARES: ya fijo');
    }
  });
  ui.alert('Congelamiento', log.join('\n'), ui.ButtonSet.OK);
}

// --- REPARAR MES ACTUAL ---
function repararMesActual() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var log = ['DIAGNOSTICO: ' + MESES[mesActual] + ' ' + anioActual, ''];

  CONFIG.vendedores.forEach(function(vend) {
    log.push('=== ' + vend.nombre + ' ===');

    if (vend.tipo === 'importrange') {
      var sheet = ss.getSheetByName(vend.solapa);
      if (!sheet) { log.push('  SOLAPA "' + vend.solapa + '" NO ENCONTRADA'); return; }
      var grupos = detectarGrupos_(sheet);
      var lastCol = sheet.getLastColumn();
      for (var g = 0; g < grupos.length; g++) {
        var anio = detectarAnio_(sheet, grupos[g]);
        if (anio !== anioActual) continue;
        var row = grupos[g][mesActual];
        var info = infoFila_(sheet, row, lastCol);
        log.push('  Fila ' + row + ': ' + info.formulas + ' con formula, ' + info.valores + ' con valor fijo');
      }

    } else {
      var destSheet = ss.getSheetByName(CONFIG.solapaDestino);
      if (!destSheet) { log.push('  PANEL DE CONTROL no encontrado'); return; }
      var filas = anioActual === 2026 ? vend.filas2026 : (anioActual === 2027 ? vend.filas2027 : null);
      if (!filas) { log.push('  Sin config para ' + anioActual); return; }
      var col = vend.columna;
      var rP = destSheet.getRange(col + filas[mesActual]);
      var rD = destSheet.getRange(col + (filas[mesActual] + 1));
      log.push('  PESOS ' + col + filas[mesActual] + ': ' + rP.getValue() + (rP.getFormula() ? ' (' + rP.getFormula() + ')' : ' (fijo)'));
      log.push('  DOLARES ' + col + (filas[mesActual]+1) + ': ' + rD.getValue() + (rD.getFormula() ? ' (' + rD.getFormula() + ')' : ' (fijo)'));
    }
    log.push('');
  });
  ui.alert('Diagnostico', log.join('\n'), ui.ButtonSet.OK);
}

// --- ARREGLAR ABRIL (TECNO) ---
function arreglarAbril() {
  var dest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.solapaDestino);
  if (!dest) { SpreadsheetApp.getUi().alert('No se encontro PANEL DE CONTROL'); return; }
  dest.getRange('F20').setFormula("='VENTAS TECNO EN PESOS'!AM70-1082250");
  dest.getRange('F21').setFormula("='VENTA TECNO EN DOLARES'!AQ284-30661");
  SpreadsheetApp.getUi().alert('Abril TECNO reparado:\nF20 = fuente pesos - 1082250\nF21 = fuente dolares - 30661');
}

// --- TRIGGER AUTOMATICO ---
function configurarTriggerMensual() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'congelarMes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('congelarMes').timeBased().onMonthDay(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Trigger configurado: dia 1 de cada mes a las 5-6 AM');
}

// --- MENU ---
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Congelamiento')
    .addItem('Vista Previa (sin cambios)', 'vistaPrevia')
    .addItem('Diagnostico Mes Actual', 'repararMesActual')
    .addSeparator()
    .addItem('Congelar Mes Anterior', 'congelarMes')
    .addItem('Congelar Mes Especifico...', 'congelarMesEspecifico')
    .addSeparator()
    .addItem('Arreglar Abril (TECNO)', 'arreglarAbril')
    .addItem('Trigger Automatico', 'configurarTriggerMensual')
    .addToUi();
}
