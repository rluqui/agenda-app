// Módulo de roles
function obtenerRoles() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roles");
  return hoja.getDataRange().getValues();
}
