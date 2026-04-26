// ============================================
// CONGELAMIENTO DE MESES - EXTERNO (Universal)
// ============================================
// Pegar este script en cada planilla EXTERNA de vendedor.
// Detecta automaticamente los meses y columnas.
// Funciona en cualquier planilla sin configurar nada.
// ============================================

var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function detectarGrupos_(sheet) {
  var nombres = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
                 'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  var lastRow = sheet.getLastRow();
  if (lastRow < 12) return [];
  var maxCol = Math.min(sheet.getLastColumn(), 10);
  if (maxCol < 1) return [];
  var mejorCol = -1, mejorCount = 0, mejorData = [];
  for (var col = 1; col <= maxCol; col++) {
    var vals = sheet.getRange(1, col, lastRow, 1).getDisplayValues();
    var found = [];
    for (var i = 0; i < vals.length; i++) {
      var val = vals[i][0].trim().toUpperCase();
      var idx = nombres.indexOf(val);
      if (idx >= 0) found.push({row: i + 1, mes: idx});
    }
    if (found.length > mejorCount) {
      mejorCount = found.length;
      mejorData = found;
      mejorCol = col;
    }
  }
  if (mejorData.length < 12) return [];
  var grupos = [];
  var used = {};
  for (var i = 0; i < mejorData.length; i++) {
    if (mejorData[i].mes !== 0 || used[i]) continue;
    var grupo = [mejorData[i].row];
    var nextMes = 1;
    for (var j = i + 1; j < mejorData.length && nextMes < 12; j++) {
      if (used[j]) continue;
      if (mejorData[j].mes === nextMes) {
        grupo.push(mejorData[j].row);
        nextMes++;
      } else if (mejorData[j].mes === 0) {
        break;
      }
    }
    if (grupo.length === 12) {
      grupos.push(grupo);
      for (var k = i; k < mejorData.length; k++) {
        if (grupo.indexOf(mejorData[k].row) >= 0) used[k] = true;
      }
    }
  }
  return grupos;
}

function detectarAnio_(sheet, grupo) {
  var startRow = Math.max(1, grupo[0] - 10);
  var numRows = grupo[0] - startRow;
  if (numRows <= 0) return null;
  var lastCol = Math.min(sheet.getLastColumn(), 20);
  var values = sheet.getRange(startRow, 1, numRows, lastCol).getDisplayValues();
  for (var r = values.length - 1; r >= 0; r--) {
    for (var c = 0; c < values[r].length; c++) {
      var val = values[r][c];
      if (val.indexOf('2027') >= 0) return 2027;
      if (val.indexOf('2026') >= 0) return 2026;
      if (val.indexOf('2025') >= 0) return 2025;
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
    if (formulas[c]) {
      sheet.getRange(row, c + 2).setValue(values[c]);
      count++;
    }
  }
  return count;
}

function congelarMes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var mesC = mesActual - 1, anioC = anioActual;
  if (mesC < 0) { mesC = 11; anioC--; }
  var log = ['CONGELAMIENTO DE MES',
    'Fecha: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Congelando: ' + MESES[mesC] + ' ' + anioC,
    'Planilla: ' + ss.getName(), ''];

  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var grupos = detectarGrupos_(sheet);
    if (grupos.length === 0) return;
    var lastCol = sheet.getLastColumn();
    var found = false;
    log.push('=== Solapa: ' + sheet.getName() + ' (' + grupos.length + ' bloques) ===');
    for (var g = 0; g < grupos.length; g++) {
      var anio = detectarAnio_(sheet, grupos[g]);
      if (anio !== anioC) continue;
      found = true;
      var row = grupos[g][mesC];
      var count = congelarFila_(sheet, row, lastCol);
      log.push('  Fila ' + row + ': ' + count + ' celdas congeladas');
    }
    if (!found) log.push('  Sin datos para ' + anioC);
    log.push('');
  });

  var msg = log.join('\n'); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Congelamiento', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}

function vistaPrevia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var mesC = mesActual - 1, anioC = anioActual;
  if (mesC < 0) { mesC = 11; anioC--; }
  var log = ['VISTA PREVIA - NO modifica nada',
    'Congelaria: ' + MESES[mesC] + ' ' + anioC,
    'Planilla: ' + ss.getName(), ''];

  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var grupos = detectarGrupos_(sheet);
    if (grupos.length === 0) return;
    var lastCol = sheet.getLastColumn();
    log.push('=== Solapa: ' + sheet.getName() + ' (' + grupos.length + ' bloques) ===');
    for (var g = 0; g < grupos.length; g++) {
      var anio = detectarAnio_(sheet, grupos[g]);
      if (anio !== anioC) continue;
      var row = grupos[g][mesC];
      var range = sheet.getRange(row, 2, 1, lastCol - 1);
      var formulas = range.getFormulas()[0];
      var values = range.getValues()[0];
      var conFormula = 0, conValor = 0;
      for (var c = 0; c < values.length; c++) {
        if (formulas[c]) conFormula++;
        else if (values[c] !== '' && values[c] !== null && values[c] !== 0) conValor++;
      }
      log.push('  ' + MESES[mesC] + ' (fila ' + row + '): ' + conFormula + ' con formula, ' + conValor + ' con valor fijo');
    }
    log.push('');
  });
  SpreadsheetApp.getUi().alert('Vista Previa', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function diagnostico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombres = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
                 'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  var log = ['DIAGNOSTICO - ' + ss.getName(), ''];
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return;
    log.push('=== ' + sheet.getName() + ' (' + lastRow + ' filas, ' + sheet.getLastColumn() + ' col) ===');
    var maxCol = Math.min(sheet.getLastColumn(), 10);
    for (var col = 1; col <= maxCol; col++) {
      var vals = sheet.getRange(1, col, lastRow, 1).getDisplayValues();
      var count = 0, primera = 0;
      for (var i = 0; i < vals.length; i++) {
        var val = vals[i][0].trim().toUpperCase();
        if (nombres.indexOf(val) >= 0) { count++; if (!primera) primera = i + 1; }
      }
      if (count > 0) {
        log.push('  Col ' + String.fromCharCode(64 + col) + ': ' + count + ' meses (1ra fila: ' + primera + ')');
      }
    }
    var grupos = detectarGrupos_(sheet);
    log.push('  Bloques: ' + grupos.length);
    for (var g = 0; g < grupos.length; g++) {
      var anio = detectarAnio_(sheet, grupos[g]);
      log.push('    #' + (g+1) + ': filas ' + grupos[g][0] + '-' + grupos[g][11] + ' (anio: ' + (anio || '?') + ')');
    }
    log.push('');
  });
  SpreadsheetApp.getUi().alert('Diagnostico', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

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

  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var grupos = detectarGrupos_(sheet);
    if (grupos.length === 0) return;
    var lastCol = sheet.getLastColumn();
    log.push('=== Solapa: ' + sheet.getName() + ' ===');
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
  });
  ui.alert('Congelamiento', log.join('\n'), ui.ButtonSet.OK);
}

function configurarTriggerMensual() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'congelarMes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('congelarMes').timeBased().onMonthDay(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Trigger configurado: dia 1 de cada mes a las 5-6 AM');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Congelamiento')
    .addItem('Vista Previa (sin cambios)', 'vistaPrevia')
    .addItem('Diagnostico', 'diagnostico')
    .addSeparator()
    .addItem('Congelar Mes Anterior', 'congelarMes')
    .addItem('Congelar Mes Especifico...', 'congelarMesEspecifico')
    .addSeparator()
    .addItem('Trigger Automatico', 'configurarTriggerMensual')
    .addToUi();
}
