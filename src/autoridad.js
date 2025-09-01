// Módulo de autoridades
function obtenerAutoridades() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Autoridades");
  return hoja.getDataRange().getValues();
}
