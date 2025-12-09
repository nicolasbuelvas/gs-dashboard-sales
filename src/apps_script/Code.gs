function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Sales Dashboard")
    .addItem("Open Dashboard", "showDashboard")
    .addToUi();
}

function showDashboard() {
  var html = HtmlService.createTemplateFromFile("ui").evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Retorna todos los datos del CSV ya cargado en la hoja
 * IMPORTANTE: Debe existir una hoja llamada "SalesData"
 */
function getSalesData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("SalesData");
  const rows = sheet.getDataRange().getValues();

  const headers = rows.shift();

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}