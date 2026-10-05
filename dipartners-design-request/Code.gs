function doPost(e) {
  const SHEET_NAME = "Solicitudes";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Fecha registro","Proyecto","Solicitante","Tipo","Objetivo","Entregable",
      "Mensaje","Público objetivo","Estilo","Referencias","Fecha entrega",
      "Prioridad","Formato","Materiales","Observaciones","Email","Contacto"
    ]);
  }

  const d = JSON.parse(e.postData.contents || "{}");
  sheet.appendRow([
    new Date(), d.proyecto||"", d.solicitante||"", d.tipo||"", d.objetivo||"",
    d.entregable||"", d.mensaje||"", d.publico||"", d.estilo||"", d.referencias||"",
    d.fecha||"", d.prioridad||"", d.formato||"", d.materiales||"",
    d.observaciones||"", d.email||"", d.contacto||""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}