// ============================================
// DISTRIBUCION DE CHECKBOX - MELY
// ============================================
// Distribuye checkboxes (casillas de verificacion) en 3 solapas:
// CREDITOS, VENTA EN DOLARES, VENTA EN PESOS
// Para cada cliente, coloca checkboxes en los meses correspondientes
// segun fecha de inicio y cantidad de cuotas.
// ============================================

function distribucionDeCheckbox() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = ['creditos', 'VENTA EN DOLARES', 'VENTA EN PESOS'];
  var log = [];

  tabs.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      log.push(tabName + ': solapa no encontrada');
      return;
    }
    var result = procesarSolapa_(sheet);
    log.push(tabName + ': ' + result.count + ' clientes procesados');
  });

  SpreadsheetApp.getUi().alert('Distribucion de Checkbox\n\n' + log.join('\n'));
}

function procesarSolapa_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 10) return {count: 0};

  var headerRows = sheet.getRange(1, 1, 3, lastCol).getValues();

  // Auto-detectar columnas de meses buscando fechas en filas 1-2
  var monthColStart = -1, monthColEnd = -1;
  var baseMes = -1;
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < headerRows[r].length; c++) {
      var val = headerRows[r][c];
      var fecha = null;
      if (val instanceof Date) {
        fecha = val;
      } else if (typeof val === 'string') {
        var m = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (m) fecha = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
      }
      if (fecha && !isNaN(fecha.getTime())) {
        var col1 = c + 1; // 1-based
        if (monthColStart < 0) {
          monthColStart = col1;
          baseMes = fecha.getFullYear() * 12 + (fecha.getMonth() + 1);
        }
        monthColEnd = col1;
      }
    }
    if (monthColStart >= 0) break;
  }

  if (monthColStart < 0) return {count: 0};

  // Auto-detectar columna FECHA y CUOTAS desde fila de encabezados (filas 2-3)
  var colFecha = -1, colCuotas = -1;
  for (var r = 1; r < 3; r++) {
    for (var c = 0; c < Math.min(headerRows[r].length, monthColStart - 1); c++) {
      var txt = String(headerRows[r][c]).trim().toUpperCase();
      if (txt === 'FECHA') colFecha = c;
      if (txt === 'CUOTAS') colCuotas = c;
    }
  }

  if (colFecha < 0 || colCuotas < 0) return {count: 0};

  var allData = sheet.getRange(1, 1, lastRow, Math.max(monthColEnd, colCuotas + 1)).getValues();

  var checkbox = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  var count = 0;

  // Buscar filas de clientes: tienen nombre en col C (index 2) y no son encabezados
  for (var r = 2; r < allData.length - 1; r++) {
    var nombre = allData[r][2]; // col C
    if (!nombre || String(nombre).trim() === '') continue;
    var nomTxt = String(nombre).trim().toUpperCase();
    if (nomTxt === 'CLIENTES' || nomTxt === 'CLIENTE') continue;

    var fechaVal = allData[r][colFecha];
    var cuotasVal = allData[r][colCuotas];

    if (!fechaVal || !cuotasVal) continue;

    var fecha;
    if (fechaVal instanceof Date) {
      fecha = fechaVal;
    } else {
      var parts = String(fechaVal).split('/');
      if (parts.length < 3) continue;
      fecha = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    }
    if (isNaN(fecha.getTime())) continue;

    var cuotas = parseInt(cuotasVal);
    if (isNaN(cuotas) || cuotas <= 0) continue;

    var mesCliente = fecha.getFullYear() * 12 + (fecha.getMonth() + 1);
    var colInicio = monthColStart + (mesCliente - baseMes);
    if (colInicio < monthColStart) colInicio = monthColStart;

    var cbRow = r + 2; // fila de checkbox (1-based: r es 0-based +1 para 1-based +1 para siguiente fila)

    for (var q = 0; q < cuotas; q++) {
      var col = colInicio + q;
      if (col > monthColEnd) break;
      var cell = sheet.getRange(cbRow, col);
      cell.setDataValidation(checkbox);
      var val = cell.getValue();
      if (val === '' || val === null) {
        cell.setValue(false);
      }
    }
    count++;
    r++; // saltar la fila de checkbox
  }

  return {count: count};
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Herramientas')
    .addItem('Distribucion de Checkbox', 'distribucionDeCheckbox')
    .addToUi();
}
