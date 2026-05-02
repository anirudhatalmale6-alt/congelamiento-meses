// ============================================
// CONGELAMIENTO MAESTRO - CENTRALIZADO
// ============================================
// Script unico en la planilla central (PANEL DE CONTROL)
// Congela TECNO y CREDITOS localmente + 7 vendedores externos
// ============================================

// ---- CONFIGURACION CENTRAL (TECNO / CREDITOS) ----
var CONFIG_CENTRAL = {
  solapaDestino: 'PANEL DE CONTROL',
  vendedores: [
    {
      nombre: 'TECNO', columna: 'F',
      filas2026: [8,12,16,20,24,28,32,36,40,44,48,52],
      filas2027: [62,66,70,74,78,82,86,90,94,98,102,106],
      fuentes: {
        pesos: { ref: "'VENTAS TECNO EN PESOS'!AM70" },
        dolares: { ref: "'VENTA TECNO EN DOLARES'!AQ284" }
      }
    },
    {
      nombre: 'CREDITOS', columna: 'J',
      filas2026: [8,12,16,20,24,28,32,36,40,44,48,52],
      filas2027: [62,66,70,74,78,82,86,90,94,98,102,106],
      fuentes: {
        pesos: { ref: "'VENTA CREDITOS EN PESOS'!AL68" },
        dolares: { ref: "'VENTA CREDITOS EN DOLARES'!AL66" }
      }
    }
  ]
};

// ---- CONFIGURACION EXTERNOS ----
// IMPORTANTE: Reemplazar los IDs con los reales
var EXTERNOS = [
  { nombre: 'TINO',         id: '1KBusYiaUuD4-rQ-JHTTv6kaH27xHyC4p6IFyRoScimM' },
  { nombre: 'OSITO S.R.L.', id: '1hrDYiUGbfwars04Wx_ZImVrrgLZKIzO6bDk-CxGeX-c' },
  { nombre: 'PATITO S.A.',  id: '1k1Uyphm-df7eN6IyEx77fqOixuROy4p1Cfq8t5tng78' },
  { nombre: 'GONZA',        id: '1DlKcy7lmn0Yr02fGrEUf8FiB-BEltv7eVw5nKuObTag' },
  { nombre: 'MELY',         id: '1_EJbkqX7Xp8ui8QCaNaIMv3SteKTApx9S7xWiCbn2Q4' },
  { nombre: 'LINEA 314',    id: '1kIO9TlRatBTWP5K1sM5KeguFXNB5qzSu80W41mQkbZw' },
  { nombre: 'TOBIAS',       id: '1_jCQkl2fBgsWVH326o6VBLq2tciSK6TZw43NCb0UT8Q' }
];

var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
var NOMBRES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
               'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

// ---- FUNCIONES AUXILIARES ----

function armarFormula_(ref, acum) {
  if (acum === 0) return '=' + ref;
  return acum > 0 ? '=' + ref + '-' + acum : '=' + ref + '+' + Math.abs(acum);
}

function extraerAcumulado_(formula) {
  if (!formula) return 0;
  var m = formula.match(/-\s*([\d.]+)\s*$/);
  return m ? parseFloat(m[1]) : 0;
}

function limpiarTexto_(txt) {
  var s = String(txt), r = '';
  for (var i = 0; i < s.length; i++) {
    r += s.charCodeAt(i) > 127 ? ' ' : s[i];
  }
  return r.replace(/\s+/g, ' ').trim().toUpperCase();
}

// ---- LOGICA EXTERNA (auto-detecta meses en cualquier planilla) ----

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

// ---- CONGELAMIENTO CENTRAL (TECNO / CREDITOS) ----

function congelarCentral_(mesC, anioC, soloPrevia) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destSheet = ss.getSheetByName(CONFIG_CENTRAL.solapaDestino);
  if (!destSheet) return ['ERROR: No se encontro ' + CONFIG_CENTRAL.solapaDestino];
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var log = [];

  CONFIG_CENTRAL.vendedores.forEach(function(vend) {
    log.push('  === ' + vend.nombre + ' ===');
    var filasC = anioC === 2026 ? vend.filas2026 : (anioC === 2027 ? vend.filas2027 : null);
    if (!filasC) { log.push('    Sin config para ' + anioC); return; }
    var fc = filasC[mesC], col = vend.columna;
    var rP = destSheet.getRange(col + fc), rD = destSheet.getRange(col + (fc + 1));
    var fP = rP.getFormula(), fD = rD.getFormula();
    var vP = rP.getValue(), vD = rD.getValue();

    if (soloPrevia) {
      log.push('    PESOS ' + col + fc + ' = ' + vP + (fP ? ' (formula)' : ' (fijo)'));
      log.push('    DOLARES ' + col + (fc+1) + ' = ' + vD + (fD ? ' (formula)' : ' (fijo)'));
    } else {
      var prevP = extraerAcumulado_(fP), prevD = extraerAcumulado_(fD);
      if (fP) { rP.setValue(vP); log.push('    PESOS congelado: ' + col + fc + ' = ' + vP); }
      else log.push('    PESOS ya fijo: ' + col + fc);
      if (fD) { rD.setValue(vD); log.push('    DOLARES congelado: ' + col + (fc+1) + ' = ' + vD); }
      else log.push('    DOLARES ya fijo: ' + col + (fc+1));
      // Actualizar mes siguiente con nueva formula
      var filasN = anioActual === 2026 ? vend.filas2026 : (anioActual === 2027 ? vend.filas2027 : null);
      if (filasN) {
        var fn = filasN[mesActual];
        var newP = prevP + (typeof vP === 'number' ? vP : 0);
        var newD = prevD + (typeof vD === 'number' ? vD : 0);
        destSheet.getRange(col + fn).setFormula(armarFormula_(vend.fuentes.pesos.ref, newP));
        destSheet.getRange(col + (fn+1)).setFormula(armarFormula_(vend.fuentes.dolares.ref, newD));
        log.push('    Nuevo acumulado: P=$' + newP + ' D=U$D' + newD);
      }
    }
  });
  return log;
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

// ---- CONGELAMIENTO GENERICO DE FILA (con activacion de mes siguiente) ----

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

// ---- CONGELAMIENTO EXTERNO (abre planilla por ID) ----

function congelarExterno_(extConfig, mesC, anioC, soloPrevia) {
  var log = [];
  try {
    var ss = SpreadsheetApp.openById(extConfig.id);
    var sheets = ss.getSheets();
    var found = false;
    sheets.forEach(function(sheet) {
      var info = analizarSheet_(sheet);
      if (info.grupos.length === 0) return;
      var colLetra = String.fromCharCode(65 + info.limiteCol - 1);
      for (var g = 0; g < info.grupos.length; g++) {
        var anio = detectarAnio_(sheet, info.grupos[g]);
        if (anio !== anioC) continue;
        found = true;
        var row = info.grupos[g][mesC];
        var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
        var exclCols = Object.keys(excl).map(function(c){return String.fromCharCode(65+parseInt(c));});
        if (exclCols.length > 0) log.push('    Columnas excluidas (INV UNIF): ' + exclCols.join(', '));
        if (soloPrevia) {
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
          log.push('    ' + sheet.getName() + ' fila ' + row + ': ' + conFormula + ' formulas, ' + conValor + ' fijos (hasta col ' + colLetra + ')');
          if (detalles.length > 0) log.push('    Valores a congelar: ' + detalles.join(' | '));
        } else {
          var nextRow = (mesC < 11) ? info.grupos[g][mesC + 1] : 0;
          var result = congelarFila_(sheet, row, info.limiteCol, excl, nextRow);
          log.push('    ' + sheet.getName() + ' fila ' + row + ': ' + result.frozen + ' formulas congeladas' + (result.activated > 0 ? ', ' + result.activated + ' formulas activadas en fila ' + nextRow : ''));
        }
      }
    });
    if (!found) log.push('    Sin datos para ' + MESES[mesC] + ' ' + anioC);
  } catch(e) {
    log.push('    ERROR: ' + e.message);
  }
  return log;
}

// ---- CONGELAMIENTO TABS INDIVIDUALES EN PLANILLA CENTRAL ----
// Escanea TODAS las solapas de la planilla central (excepto PANEL DE CONTROL)
// y aplica congelamiento + activacion de mes siguiente usando deteccion automatica

function congelarTabsCentral_(mesC, anioC, soloPrevia) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var nombre = sheet.getName();
    if (nombre === CONFIG_CENTRAL.solapaDestino) return;
    var info = analizarSheet_(sheet);
    if (info.grupos.length === 0) return;
    var colLetra = String.fromCharCode(65 + info.limiteCol - 1);
    var found = false;
    for (var g = 0; g < info.grupos.length; g++) {
      var anio = detectarAnio_(sheet, info.grupos[g]);
      if (anio !== anioC) continue;
      found = true;
      var row = info.grupos[g][mesC];
      var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
      var exclCols = Object.keys(excl).map(function(c){return String.fromCharCode(65+parseInt(c));});
      if (exclCols.length > 0) log.push('    Columnas excluidas (INV UNIF): ' + exclCols.join(', '));
      if (soloPrevia) {
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
        log.push('    ' + nombre + ' fila ' + row + ': ' + conFormula + ' formulas, ' + conValor + ' fijos (hasta col ' + colLetra + ')');
        if (detalles.length > 0) log.push('    Valores a congelar: ' + detalles.join(' | '));
      } else {
        var nextRow = (mesC < 11) ? info.grupos[g][mesC + 1] : 0;
        var result = congelarFila_(sheet, row, info.limiteCol, excl, nextRow);
        log.push('    ' + nombre + ' fila ' + row + ': ' + result.frozen + ' formulas congeladas' + (result.activated > 0 ? ', ' + result.activated + ' formulas activadas en fila ' + nextRow : ''));
      }
    }
    if (!found && info.grupos.length > 0) log.push('    ' + nombre + ': sin datos para ' + anioC);
  });
  return log;
}

// ---- FUNCIONES PRINCIPALES ----

function vistaPreviaTodo() {
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var mesC = mesActual - 1, anioC = anioActual;
  if (mesC < 0) { mesC = 11; anioC--; }

  var log = ['VISTA PREVIA GLOBAL - NO modifica nada',
    'Fecha: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Congelaria: ' + MESES[mesC] + ' ' + anioC, '',
    '=== PLANILLA CENTRAL (TECNO / CREDITOS) ==='];

  log = log.concat(congelarCentral_(mesC, anioC, true));
  log.push('');

  log.push('=== SOLAPAS INDIVIDUALES CENTRAL ===');
  log = log.concat(congelarTabsCentral_(mesC, anioC, true));
  log.push('');

  EXTERNOS.forEach(function(ext) {
    log.push('=== ' + ext.nombre + ' (externo) ===');
    log = log.concat(congelarExterno_(ext, mesC, anioC, true));
  });

  SpreadsheetApp.getUi().alert('Vista Previa Global', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function congelarTodo() {
  var ui = SpreadsheetApp.getUi();
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();
  var mesC = mesActual - 1, anioC = anioActual;
  if (mesC < 0) { mesC = 11; anioC--; }

  var resp = ui.alert('Confirmar Congelamiento',
    'Se va a congelar ' + MESES[mesC] + ' ' + anioC + ' en TODAS las planillas.\n\nTECNO, CREDITOS + 7 vendedores externos.\n\nContinuar?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var log = ['CONGELAMIENTO GLOBAL',
    'Fecha: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Congelando: ' + MESES[mesC] + ' ' + anioC, '',
    '=== PLANILLA CENTRAL (TECNO / CREDITOS) ==='];

  log = log.concat(congelarCentral_(mesC, anioC, false));
  log.push('');

  log.push('=== SOLAPAS INDIVIDUALES CENTRAL ===');
  log = log.concat(congelarTabsCentral_(mesC, anioC, false));
  log.push('');

  EXTERNOS.forEach(function(ext) {
    log.push('=== ' + ext.nombre + ' (externo) ===');
    log = log.concat(congelarExterno_(ext, mesC, anioC, false));
  });

  var msg = log.join('\n');
  Logger.log(msg);
  ui.alert('Congelamiento Completo', msg, ui.ButtonSet.OK);
}

function congelarMesEspecificoTodo() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Congelar Mes Especifico', 'Mes y anio (ej: 3 2026 para Marzo 2026):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var p = resp.getResponseText().trim().split(/\s+/);
  if (p.length < 2) { ui.alert('Formato: numero_mes anio (ej: 3 2026)'); return; }
  var mes = parseInt(p[0]) - 1, anio = parseInt(p[1]);
  if (isNaN(mes) || mes < 0 || mes > 11 || isNaN(anio)) { ui.alert('Invalido'); return; }

  var confirm = ui.alert('Confirmar',
    'Congelar ' + MESES[mes] + ' ' + anio + ' en TODAS las planillas?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var log = ['CONGELAMIENTO: ' + MESES[mes] + ' ' + anio, '',
    '=== PLANILLA CENTRAL ==='];
  log = log.concat(congelarCentral_(mes, anio, false));
  log.push('');

  log.push('=== SOLAPAS INDIVIDUALES CENTRAL ===');
  log = log.concat(congelarTabsCentral_(mes, anio, false));
  log.push('');

  EXTERNOS.forEach(function(ext) {
    log.push('=== ' + ext.nombre + ' ===');
    log = log.concat(congelarExterno_(ext, mes, anio, false));
  });

  ui.alert('Congelamiento', log.join('\n'), ui.ButtonSet.OK);
}

// ---- REPARACION DE FORMULAS (activa mes actual en todas las planillas) ----

function repararFila_(sheet, targetRow, limiteCol, excluidas, grupo) {
  var targetRange = sheet.getRange(targetRow, 1, 1, limiteCol);
  var targetFormulas = targetRange.getFormulas()[0];
  var hasFormulas = false;
  for (var c = 0; c < targetFormulas.length; c++) {
    if (excluidas && excluidas[c]) continue;
    if (targetFormulas[c]) { hasFormulas = true; break; }
  }
  if (hasFormulas) return {status: 'ok', copied: 0};

  var sourceRow = -1, sourceMes = -1;
  var mesActual = new Date().getMonth();
  for (var m = mesActual + 1; m < 12; m++) {
    var row = grupo[m];
    var formulas = sheet.getRange(row, 1, 1, limiteCol).getFormulas()[0];
    for (var c2 = 0; c2 < formulas.length; c2++) {
      if (excluidas && excluidas[c2]) continue;
      if (formulas[c2]) { sourceRow = row; sourceMes = m; break; }
    }
    if (sourceRow >= 0) break;
  }
  if (sourceRow < 0) {
    for (var m2 = mesActual - 1; m2 >= 0; m2--) {
      var row2 = grupo[m2];
      var formulas2 = sheet.getRange(row2, 1, 1, limiteCol).getFormulas()[0];
      for (var c3 = 0; c3 < formulas2.length; c3++) {
        if (excluidas && excluidas[c3]) continue;
        if (formulas2[c3]) { sourceRow = row2; sourceMes = m2; break; }
      }
      if (sourceRow >= 0) break;
    }
  }
  if (sourceRow < 0) return {status: 'no_source', copied: 0};

  var sourceFormulas = sheet.getRange(sourceRow, 1, 1, limiteCol).getFormulas()[0];
  var copied = 0;
  for (var c4 = 0; c4 < sourceFormulas.length; c4++) {
    if (excluidas && excluidas[c4]) continue;
    if (sourceFormulas[c4]) {
      sheet.getRange(sourceRow, c4 + 1).copyTo(sheet.getRange(targetRow, c4 + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      copied++;
    }
  }
  return {status: 'repaired', copied: copied, sourceMes: sourceMes, sourceRow: sourceRow};
}

function repararTodo() {
  var now = new Date();
  var mesActual = now.getMonth(), anioActual = now.getFullYear();

  var log = ['REPARACION DE FORMULAS - GLOBAL',
    'Fecha: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Reparando: ' + MESES[mesActual] + ' ' + anioActual, ''];

  // 1) Central individual tabs
  log.push('=== SOLAPAS INDIVIDUALES CENTRAL ===');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getName() === CONFIG_CENTRAL.solapaDestino) return;
    var info = analizarSheet_(sheet);
    if (info.grupos.length === 0) return;
    for (var g = 0; g < info.grupos.length; g++) {
      var anio = detectarAnio_(sheet, info.grupos[g]);
      if (anio !== anioActual) continue;
      var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
      var result = repararFila_(sheet, info.grupos[g][mesActual], info.limiteCol, excl, info.grupos[g]);
      if (result.status === 'ok') log.push('  ' + sheet.getName() + ': ya tiene formulas');
      else if (result.status === 'repaired') log.push('  ' + sheet.getName() + ': ' + result.copied + ' formulas copiadas desde ' + MESES[result.sourceMes]);
      else log.push('  ' + sheet.getName() + ': SIN FUENTE para copiar formulas');
    }
  });
  log.push('');

  // 2) External vendors
  EXTERNOS.forEach(function(ext) {
    log.push('=== ' + ext.nombre + ' (externo) ===');
    try {
      var extSS = SpreadsheetApp.openById(ext.id);
      extSS.getSheets().forEach(function(sheet) {
        var info = analizarSheet_(sheet);
        if (info.grupos.length === 0) return;
        for (var g = 0; g < info.grupos.length; g++) {
          var anio = detectarAnio_(sheet, info.grupos[g]);
          if (anio !== anioActual) continue;
          var excl = detectarColsExcluidas_(sheet, info.grupos[g]);
          var result = repararFila_(sheet, info.grupos[g][mesActual], info.limiteCol, excl, info.grupos[g]);
          if (result.status === 'ok') log.push('  ' + sheet.getName() + ': ya tiene formulas');
          else if (result.status === 'repaired') log.push('  ' + sheet.getName() + ': ' + result.copied + ' formulas copiadas desde ' + MESES[result.sourceMes]);
          else log.push('  ' + sheet.getName() + ': SIN FUENTE para copiar formulas');
        }
      });
    } catch(e) {
      log.push('  ERROR: ' + e.message);
    }
  });

  var msg = log.join('\n');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Reparacion de Formulas', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function configurarTriggerMensual() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'congelarTodo') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('congelarTodo').timeBased().onMonthDay(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Trigger configurado: dia 1 de cada mes a las 5-6 AM\nCongela TODO automaticamente (central + externos)');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Congelamiento')
    .addItem('Vista Previa GLOBAL (sin cambios)', 'vistaPreviaTodo')
    .addSeparator()
    .addItem('Reparar Formulas Mes Actual - GLOBAL', 'repararTodo')
    .addSeparator()
    .addItem('Congelar Mes Anterior - TODO', 'congelarTodo')
    .addItem('Congelar Mes Especifico - TODO...', 'congelarMesEspecificoTodo')
    .addSeparator()
    .addItem('Trigger Automatico Mensual', 'configurarTriggerMensual')
    .addToUi();
}
