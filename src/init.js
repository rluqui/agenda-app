// Inicialización del entorno
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Agenda")
    .addItem("Actualizar datos", "actualizarDatos")
    .addToUi();
}
