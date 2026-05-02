// ============================================
// CONGELAMIENTO DE MESES - EXTERNO (Universal)
// ============================================
// Pegar este script en cada planilla EXTERNA de vendedor.
// Detecta automaticamente los meses y columnas.
// Solo congela el cuadro principal (izquierdo).
// No toca cuadros a la derecha que tengan meses.
// ============================================

var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
var NOMBRES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
               'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

function limpiarTexto_(txt) {
  var s = String(txt), r = '';
  for (var i = 0; i < s.length; i++) {
    r += s.charCodeAt(i) > 127 ? ' ' : s[i];
  }
  return r.replace(/\s+/g, ' ').trim().toUpperCase();
}

function analizarSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 12 || lastCol < 1) return {grupos: [], limiteCol: lastCol};
  var scanCols = Math.min(lastCol, 20);
  var allValues = sheet.getRange(1, 1, lastRow, scanCols).getDisplayValues();
  var monthsByCol = {};
  var monthRows = [];
  for (var r = 0; r < allValues.length; r++) {
    var foundInRow = false;
    for (var c = 0; c < allValues[r].length; c++) {
      var val = limpiarTexto_(allValues[r][c]);
      var idx = NOMBRES.indexOf(val);
      if (idx >= 0) {
        if (!monthsByCol[c]) monthsByCol[c] = 0;
        monthsByCol[c]++;
        if (!foundInRow) {
          monthRows.push({row: r + 1, mes: idx});
          foundInRow = true;
        }
      }
    }
  }
  if (monthRows.length < 12) return {grupos: [], limiteCol: lastCol};
  var primaryCol = -1;
  for (var r2 = 0; r2 < allValues.length; r2++) {
    for (var c2 = 0; c2 < allValues[r2].length; c2++) {
      var v2 = limpiarTexto_(allValues[r2][c2]);
      if (NOMBRES.indexOf(v2) >= 0) { primaryCol = c2; break; }
    }
    if (primaryCol >= 0) break;
  }
  var limiteCol = lastCol;
  var sortedCols = Object.keys(monthsByCol).map(Number).sort(function(a, b) { return a - b; });
  for (var i = 0; i < sortedCols.length; i++) {
    if (sortedCols[i] > primaryCol && monthsByCol[sortedCols[i]] >= 12) {
      limiteCol = sortedCols[i];
      break;
    }
  }
  var grupos = [];
  var used = {};
  for (var i = 0; i < monthRows.length; i++) {
    if (monthRows[i].mes !== 0 || used[i]) continue;
    var grupo = [monthRows[i].row];
    var nextMes = 1;
    for (var j = i + 1; j < monthRows.length && nextMes < 12; j++) {
      if (used[j]) continue;
      if (monthRows[j].mes === nextMes) {
        grupo.push(monthRows[j].row);
        nextMes++;
      } else if (monthRows[j].mes === 0) {
        break;
      }
    }
    if (grupo.length === 12) {
      grupos.push(grupo);
      for (var k = i; k < monthRows.length; k++) {
        if (grupo.indexOf(monthRows[k].row) >= 0) used[k] = true;
      }
    }
  }
  return {grupos: grupos, limiteCol: limiteCol};
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

function detectarColsExcluidas_(sheet, grupo) {
  var startRow = Math.max(1, grupo[0] - 10);
  var numRows = grupo[0] - startRow;
  if (numRows <= 0) return {};
  var lastCol = Math.min(sheet.getLastColumn(), 40);
  var values = sheet.getRange(startRow, 1, numRows, lastCol).getDisplayValues();
  var excluidas = {};
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var t = limpiarTexto_(values[r][c]);
      if (t.indexOf('INV UNIF') >= 0 || t.indexOf('INVERSION UNIFICAD') >= 0) {
        excluidas[c] = true;
      }
    }
  }
  return excluidas;
}

function congelarFila_(sheet, row, limiteCol, excluidas, nextRow) {
  var range = sheet.getRange(row, 1, 1, limiteCol);
  var values = range.getValues()[0];
  var formulas = range.getFormulas()[0];
  var count = 0, activated = 0;
  for (var c = 0; c < values.length; c++) {
    if (excluidas && excluidas[c]) continue;
    if (formulas[c]) {
      if (nextRow && nextRow > 0) {
        var nextCell = sheet.getRange(nextRow, c + 1);
        if (!nextCell.getFormula()) {
          sheet.getRange(row, c + 1).copyTo(nextCell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
          activated++;
        }
      }
      sheet.getRange(row, c + 1).setValue(values[c]);
      count++;
    }
  }
  return {frozen: count, activated: activated};
}

function repararFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var mesActual = now.getMonth(); // 0-based: May = 4
  var anioActual = now.getFullYear();
  var log = ['REPARACION DE FORMULAS - MES ACTUAL',
    'Fecha: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Reparando: ' + MESES[mesActual] + ' ' + anioActual,
    'Planilla: ' + ss.getName(), ''];

  var totalReparadas = 0;
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var info = analizarSheet_(sheet);
    if (info.grupos.length === 0) return;
    var sheetLog = [];
    var sheetTouched = false;
    for (var g = 0; g < info.grupos.length; g++) {
      var anio = detectarAnio_(sheet, info.grupos[g]);
      if (anio !== anioActual) continue;
      var targetRow = info.grupos[g][mesActual];
      var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
      var exclCols = Object.keys(excl).map(function(c) { return String.fromCharCode(65 + parseInt(c)); });
      if (exclCols.length > 0) sheetLog.push('  Columnas excluidas (INV UNIF): ' + exclCols.join(', '));

      // Check if target row already has formulas
      var targetRange = sheet.getRange(targetRow, 1, 1, info.limiteCol);
      var targetFormulas = targetRange.getFormulas()[0];
      var hasFormulas = false;
      for (var c = 0; c < targetFormulas.length; c++) {
        if (excl[c]) continue;
        if (targetFormulas[c]) { hasFormulas = true; break; }
      }

      if (hasFormulas) {
        sheetLog.push('  ' + MESES[mesActual] + ' (fila ' + targetRow + '): YA tiene formulas, no necesita reparacion');
        continue;
      }

      // Search for a source month that HAS formulas
      // Try months after current first (Jun, Jul, Aug...), then earlier months (Apr, Mar, Feb...)
      var sourceRow = -1;
      var sourceMes = -1;
      // Forward: mesActual+1 to 11
      for (var m = mesActual + 1; m < 12; m++) {
        var candidateRow = info.grupos[g][m];
        var candidateRange = sheet.getRange(candidateRow, 1, 1, info.limiteCol);
        var candidateFormulas = candidateRange.getFormulas()[0];
        var candidateHas = false;
        for (var c2 = 0; c2 < candidateFormulas.length; c2++) {
          if (excl[c2]) continue;
          if (candidateFormulas[c2]) { candidateHas = true; break; }
        }
        if (candidateHas) {
          sourceRow = candidateRow;
          sourceMes = m;
          break;
        }
      }
      // If not found forward, try backward: mesActual-1 to 0
      if (sourceRow < 0) {
        for (var m2 = mesActual - 1; m2 >= 0; m2--) {
          var candidateRow2 = info.grupos[g][m2];
          var candidateRange2 = sheet.getRange(candidateRow2, 1, 1, info.limiteCol);
          var candidateFormulas2 = candidateRange2.getFormulas()[0];
          var candidateHas2 = false;
          for (var c3 = 0; c3 < candidateFormulas2.length; c3++) {
            if (excl[c3]) continue;
            if (candidateFormulas2[c3]) { candidateHas2 = true; break; }
          }
          if (candidateHas2) {
            sourceRow = candidateRow2;
            sourceMes = m2;
            break;
          }
        }
      }

      if (sourceRow < 0) {
        sheetLog.push('  ' + MESES[mesActual] + ' (fila ' + targetRow + '): SIN FORMULAS y no se encontro mes fuente para copiar');
        continue;
      }

      // Copy formulas from source to target
      var sourceRange = sheet.getRange(sourceRow, 1, 1, info.limiteCol);
      var sourceFormulas = sourceRange.getFormulas()[0];
      var copied = 0;
      for (var c4 = 0; c4 < sourceFormulas.length; c4++) {
        if (excl[c4]) continue;
        if (sourceFormulas[c4]) {
          var targetCell = sheet.getRange(targetRow, c4 + 1);
          sheet.getRange(sourceRow, c4 + 1).copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
          copied++;
        }
      }
      totalReparadas += copied;
      sheetTouched = true;
      sheetLog.push('  ' + MESES[mesActual] + ' (fila ' + targetRow + '): ' + copied + ' formulas copiadas desde ' + MESES[sourceMes] + ' (fila ' + sourceRow + ')');
    }
    if (sheetLog.length > 0) {
      log.push('=== Solapa: ' + sheet.getName() + ' ===');
      log = log.concat(sheetLog);
      log.push('');
    }
  });

  log.push('TOTAL: ' + totalReparadas + ' formulas reparadas');
  var msg = log.join('\n'); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Reparacion de Formulas', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
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
    var info = analizarSheet_(sheet);
    if (info.grupos.length === 0) return;
    var found = false;
    var colLetra = String.fromCharCode(65 + info.limiteCol - 1);
    log.push('=== Solapa: ' + sheet.getName() + ' (' + info.grupos.length + ' bloques, congela hasta col ' + colLetra + ') ===');
    for (var g = 0; g < info.grupos.length; g++) {
      var anio = detectarAnio_(sheet, info.grupos[g]);
      if (anio !== anioC) continue;
      found = true;
      var row = info.grupos[g][mesC];
      var nextRow = (mesC < 11) ? info.grupos[g][mesC + 1] : 0;
      var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
      var exclCols = Object.keys(excl).map(function(c){return String.fromCharCode(65+parseInt(c));});
      if (exclCols.length > 0) log.push('  Columnas excluidas (INV UNIF): ' + exclCols.join(', '));
      var result = congelarFila_(sheet, row, info.limiteCol, excl, nextRow);
      log.push('  Fila ' + row + ': ' + result.frozen + ' formulas congeladas' + (result.activated > 0 ? ', ' + result.activated + ' formulas activadas en fila ' + nextRow : ''));
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
    var info = analizarSheet_(sheet);
    if (info.grupos.length === 0) return;
    var colLetra = String.fromCharCode(65 + info.limiteCol - 1);
    log.push('=== Solapa: ' + sheet.getName() + ' (' + info.grupos.length + ' bloques, congela hasta col ' + colLetra + ') ===');
    for (var g = 0; g < info.grupos.length; g++) {
      var anio = detectarAnio_(sheet, info.grupos[g]);
      if (anio !== anioC) continue;
      var row = info.grupos[g][mesC];
      var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
      var exclCols = Object.keys(excl).map(function(c){return String.fromCharCode(65+parseInt(c));});
      if (exclCols.length > 0) log.push('  Columnas excluidas (INV UNIF): ' + exclCols.join(', '));
      var range = sheet.getRange(row, 1, 1, info.limiteCol);
      var formulas = range.getFormulas()[0];
      var values = range.getValues()[0];
      var dispValues = range.getDisplayValues()[0];
      var conFormula = 0, conValor = 0;
      var detalles = [];
      for (var c = 0; c < values.length; c++) {
        if (excl[c]) continue;
        var colL = String.fromCharCode(65 + c);
        if (formulas[c]) {
          conFormula++;
          var dv = dispValues[c] || String(values[c]);
          detalles.push(colL + '=' + dv + ' (formula)');
        } else if (values[c] !== '' && values[c] !== null && values[c] !== 0) {
          conValor++;
        }
      }
      log.push('  ' + MESES[mesC] + ' (fila ' + row + '): ' + conFormula + ' con formula, ' + conValor + ' con valor fijo');
      if (detalles.length > 0) log.push('  Valores a congelar: ' + detalles.join(' | '));
    }
    log.push('');
  });
  SpreadsheetApp.getUi().alert('Vista Previa', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function diagnostico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ['DIAGNOSTICO - ' + ss.getName(), ''];
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 5) return;
    log.push('=== ' + sheet.getName() + ' (' + lastRow + ' filas, ' + lastCol + ' col) ===');
    var scanCols = Math.min(lastCol, 20);
    var allValues = sheet.getRange(1, 1, lastRow, scanCols).getDisplayValues();
    var mesesPorCol = {};
    for (var r = 0; r < allValues.length; r++) {
      for (var c = 0; c < allValues[r].length; c++) {
        var val = limpiarTexto_(allValues[r][c]);
        if (NOMBRES.indexOf(val) >= 0) {
          var colL = String.fromCharCode(65 + c);
          if (!mesesPorCol[colL]) mesesPorCol[colL] = {count: 0, primera: r + 1};
          mesesPorCol[colL].count++;
        }
      }
    }
    for (var col in mesesPorCol) {
      log.push('Col ' + col + ': ' + mesesPorCol[col].count + ' meses (1ra fila: ' + mesesPorCol[col].primera + ')');
    }
    var info = analizarSheet_(sheet);
    log.push('Bloques: ' + info.grupos.length);
    if (info.limiteCol < lastCol) {
      log.push('Limite congelamiento: hasta col ' + String.fromCharCode(65 + info.limiteCol - 1) + ' (no toca cuadros a la derecha)');
    }
    for (var g = 0; g < info.grupos.length; g++) {
      var anio = detectarAnio_(sheet, info.grupos[g]);
      log.push('  #' + (g + 1) + ': filas ' + info.grupos[g][0] + '-' + info.grupos[g][11] + ' (anio: ' + (anio || '?') + ')');
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
    var info = analizarSheet_(sheet);
    if (info.grupos.length === 0) return;
    var colLetra = String.fromCharCode(65 + info.limiteCol - 1);
    log.push('=== Solapa: ' + sheet.getName() + ' (hasta col ' + colLetra + ') ===');
    var found = false;
    for (var g = 0; g < info.grupos.length; g++) {
      var anioG = detectarAnio_(sheet, info.grupos[g]);
      if (anioG !== anio) continue;
      found = true;
      var row = info.grupos[g][mes];
      var nextRow = (mes < 11) ? info.grupos[g][mes + 1] : 0;
      var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
      var exclCols = Object.keys(excl).map(function(c){return String.fromCharCode(65+parseInt(c));});
      if (exclCols.length > 0) log.push('  Columnas excluidas (INV UNIF): ' + exclCols.join(', '));
      var result = congelarFila_(sheet, row, info.limiteCol, excl, nextRow);
      log.push('  Fila ' + row + ': ' + result.frozen + ' formulas congeladas' + (result.activated > 0 ? ', ' + result.activated + ' formulas activadas en fila ' + nextRow : ''));
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
    .addItem('Reparar Formulas Mes Actual', 'repararFormulas')
    .addSeparator()
    .addItem('Congelar Mes Anterior', 'congelarMes')
    .addItem('Congelar Mes Especifico...', 'congelarMesEspecifico')
    .addSeparator()
    .addItem('Trigger Automatico', 'configurarTriggerMensual')
    .addToUi();
}
