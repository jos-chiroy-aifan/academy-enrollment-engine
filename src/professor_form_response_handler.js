function onProfessorFormSubmit(e) {
  const sheetName = "Professor Registration - Responses";
  const statusColumnName = "Status";
  const statusDefaultValue = "Pending";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusColIndex = headers.indexOf(statusColumnName) + 1;
  if (statusColIndex === 0) {
    throw new Error(`No se encontró la columna "${statusColumnName}" en el sheet "${sheetName}".`);
  }

  const row = e.range.getRow();
  sheet.getRange(row, statusColIndex).setValue(statusDefaultValue);
}
