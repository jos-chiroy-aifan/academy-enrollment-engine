/************************************
 * CONFIGURACIÓN BASE DE ESTUDIANTES
 ************************************/
const STUDENTS_SHEET_NAME = 'Students';
const STUDENTS_ID_COLUMN = 1;
const STUDENTS_FIRST_DATA_ROW = 2;

// Columnas obligatorias base:
const STUDENTS_REQUIRED_BASE = [2, 3, 4, 7, 8, 9, 13, 14];
// B Firstname, C Lastname, D Birthdate, G Gender, H Email, I Phone
// M Consent Photos (true/false OK), N Data Consent (must be TRUE)

/************************************
 * TRIGGER GENERADOR DE ID
 ************************************/
function generateStudentIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < STUDENTS_FIRST_DATA_ROW) return;

  for (let row = STUDENTS_FIRST_DATA_ROW; row <= lastRow; row++) {
    const id = sheet.getRange(row, STUDENTS_ID_COLUMN).getValue();
    if (id) continue;

    if (isStudentRowReady(sheet, row)) {
      sheet.getRange(row, STUDENTS_ID_COLUMN).setValue(generateNextStudentId(sheet));
    }
  }
}

/************************************
 * VALIDACIÓN DE FILA COMPLETA
 ************************************/
function isStudentRowReady(sheet, row) {
  // Validar base obligatoria (no vacía)
  for (let col of STUDENTS_REQUIRED_BASE) {
    const v = sheet.getRange(row, col).getValue();
    if (v === "" || v === null) return false;
  }

  // Data consent (N) DEBE ser TRUE
  const dataConsent = sheet.getRange(row, 14).getValue();
  if (dataConsent !== true) return false;

  // ¿Es menor?
  const isMinor = sheet.getRange(row, 6).getValue(); // F Minor (Yes/No)

  if (isMinor === "Yes") {
    // Responsable obligatorio
    const resp = sheet.getRange(row, 10).getValue();     // J
    const rel = sheet.getRange(row, 11).getValue();      // K
    const respPhone = sheet.getRange(row, 12).getValue();// L
    if (resp === "" || rel === "" || respPhone === "") return false;
  }

  return true;
}

/************************************
 * GENERAR ID AFM-SYYYYXXX
 ************************************/
function generateNextStudentId(sheet) {
  const year = new Date().getFullYear();
  const prefix = "AFM-S";
  const seqDigits = 3;

  const ids = sheet
    .getRange(STUDENTS_FIRST_DATA_ROW, STUDENTS_ID_COLUMN, sheet.getLastRow() - STUDENTS_FIRST_DATA_ROW + 1, 1)
    .getValues()
    .flat()
    .filter(v => typeof v === "string" && v.startsWith(prefix));

  let next = 1;

  if (ids.length > 0) {
    const lastId = ids[ids.length - 1];

    // Extrae año y correlativo de forma segura: AFM-S2024001
    const match = lastId.match(/^AFM-S(\d{4})(\d{3})$/);
    if (match) {
      const lastYear = Number(match[1]);
      const lastSeq = Number(match[2]);
      next = (lastYear === year) ? lastSeq + 1 : 1;
    }
  }

  return `${prefix}${year}${String(next).padStart(seqDigits, "0")}`;
}
