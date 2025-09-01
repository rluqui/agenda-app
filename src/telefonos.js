// Módulo de teléfonos
function obtenerTelefonos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teléfonos");
  return hoja.getDataRange().getValues();
}
