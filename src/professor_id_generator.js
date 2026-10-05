/************************************
 * CONFIGURACIÓN BASE DE PROFESORES
 ************************************/
const PROFESSORS_SHEET_NAME = 'Professors';
const PROFESSORS_ID_COLUMN = 1;       // A = ProfessorID
const PROFESSORS_FIRST_DATA_ROW = 2;

// Columnas obligatorias en "Professors":
// B First Name, C Last Name, D Birthdate, F Gender, G Email, H Phone, J Consent for Data Processing
const PROFESSORS_REQUIRED_BASE = [2, 3, 4, 6, 7, 8, 10];

/************************************
 * TRIGGER GENERADOR DE ID
 ************************************/
function generateProfessorIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROFESSORS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < PROFESSORS_FIRST_DATA_ROW) return;

  for (let row = PROFESSORS_FIRST_DATA_ROW; row <= lastRow; row++) {
    const id = sheet.getRange(row, PROFESSORS_ID_COLUMN).getValue();
    if (id) continue;

    if (isProfessorRowReady(sheet, row)) {
      sheet.getRange(row, PROFESSORS_ID_COLUMN).setValue(generateNextProfessorId(sheet));
    }
  }
}

/************************************
 * VALIDACIÓN DE FILA COMPLETA
 ************************************/
function isProfessorRowReady(sheet, row) {
  // Validar obligatorios (no vacíos)
  for (let col of PROFESSORS_REQUIRED_BASE) {
    const v = sheet.getRange(row, col).getValue();
    if (v === "" || v === null) return false;
  }

  // Data consent (J) DEBE ser TRUE
  const dataConsent = sheet.getRange(row, 10).getValue();
  if (dataConsent !== true) return false;

  return true;
}

/************************************
 * GENERAR ID AFM-PYYYYXXX
 ************************************/
function generateNextProfessorId(sheet) {
  const year = new Date().getFullYear();
  const prefix = "AFM-P";
  const seqDigits = 3;

  const ids = sheet
    .getRange(PROFESSORS_FIRST_DATA_ROW, PROFESSORS_ID_COLUMN, sheet.getLastRow() - PROFESSORS_FIRST_DATA_ROW + 1, 1)
    .getValues()
    .flat()
    .filter(v => typeof v === "string" && v.startsWith(prefix));

  let next = 1;

  if (ids.length > 0) {
    const lastId = ids[ids.length - 1];

    // AFM-P2026001
    const match = lastId.match(/^AFM-P(\d{4})(\d{3})$/);
    if (match) {
      const lastYear = Number(match[1]);
      const lastSeq = Number(match[2]);
      next = (lastYear === year) ? lastSeq + 1 : 1;
    }
  }

  return `${prefix}${year}${String(next).padStart(seqDigits, "0")}`;
}
