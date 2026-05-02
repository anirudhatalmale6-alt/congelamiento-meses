// ============================================
// DISTRIBUCION DE CHECKBOX - MELY (creditos)
// ============================================
// Distribuye checkboxes (casillas de verificacion) en la solapa "creditos"
// Para cada cliente, segun fecha de inicio y cantidad de cuotas,
// coloca checkboxes en los meses correspondientes.
// ============================================

function distribucionDeCheckbox() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('creditos');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No se encontro la solapa "creditos"');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 5) return;

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Columnas de meses: K(11) = Nov 2025, hasta AJ(36) = Dic 2027
  var COL_INICIO = 11; // columna K (1-based)
  var COL_FIN = 36;    // columna AJ (1-based)
  var BASE_MES = 2025 * 12 + 11; // Nov 2025

  // Columnas de datos del cliente (1-based)
  var COL_FECHA = 4;   // D
  var COL_CUOTAS = 7;  // G

  var checkbox = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  var count = 0;

  // Recorrer filas de clientes (empezando en fila 4, cada 2 filas)
  // Fila 3 = encabezados, fila 4 = primer cliente, fila 5 = checkboxes del primer cliente
  for (var r = 3; r < data.length; r += 2) { // r es indice 0-based, fila 4 = index 3
    var nombre = data[r][2]; // columna C (0-based index 2)
    if (!nombre || String(nombre).trim() === '') continue;

    var fechaVal = data[r][COL_FECHA - 1]; // columna D (0-based index 3)
    var cuotasVal = data[r][COL_CUOTAS - 1]; // columna G (0-based index 6)

    if (!fechaVal || !cuotasVal) continue;

    var fecha;
    if (fechaVal instanceof Date) {
      fecha = fechaVal;
    } else {
      // Formato d/m/yyyy
      var parts = String(fechaVal).split('/');
      if (parts.length < 3) continue;
      fecha = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    }

    var cuotas = parseInt(cuotasVal);
    if (isNaN(cuotas) || cuotas <= 0) continue;

    var mesInicio = fecha.getFullYear() * 12 + (fecha.getMonth() + 1);
    var colInicio = COL_INICIO + (mesInicio - BASE_MES);

    if (colInicio < COL_INICIO) colInicio = COL_INICIO;

    var checkboxRow = r + 2; // fila de checkboxes (1-based), r es 0-based asi que +2

    for (var q = 0; q < cuotas; q++) {
      var col = colInicio + q;
      if (col > COL_FIN) break;

      var cell = sheet.getRange(checkboxRow, col);
      cell.setDataValidation(checkbox);

      // Si la celda esta vacia, inicializar en FALSE
      var val = cell.getValue();
      if (val === '' || val === null) {
        cell.setValue(false);
      }
    }
    count++;
  }

  SpreadsheetApp.getUi().alert('Distribucion de Checkbox completada.\n' + count + ' clientes procesados.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Herramientas')
    .addItem('Distribucion de Checkbox', 'distribucionDeCheckbox')
    .addToUi();
}
