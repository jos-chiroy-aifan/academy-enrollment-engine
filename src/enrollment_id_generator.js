/************************************
 * CONFIGURACIÓN BASE DE ENROLLMENTS
 ************************************/
const ENROLLMENTS_SHEET_NAME = 'Student_Enrollments';
const ENROLLMENTS_ID_COLUMN = 1;       // A = EnrollmentID
const ENROLLMENTS_FIRST_DATA_ROW = 2;

// Columnas obligatorias base (según tu estructura):
// B StudentID, C ProgramID, D DayOfWeek, E TimeSlot, F Periods, G AssignedProfessorID, H State
const ENROLLMENTS_REQUIRED_BASE = [2, 3, 4, 5, 6, 7, 8];

/************************************
 * TRIGGER GENERADOR DE ID
 ************************************/
function generateEnrollmentIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ENROLLMENTS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < ENROLLMENTS_FIRST_DATA_ROW) return;

  for (let row = ENROLLMENTS_FIRST_DATA_ROW; row <= lastRow; row++) {
    const id = sheet.getRange(row, ENROLLMENTS_ID_COLUMN).getValue();
    if (id) continue;

    if (isEnrollmentRowReady(sheet, row)) {
      sheet.getRange(row, ENROLLMENTS_ID_COLUMN).setValue(generateNextEnrollmentId(sheet));
    }
  }
}

/************************************
 * VALIDACIÓN DE FILA COMPLETA
 ************************************/
function isEnrollmentRowReady(sheet, row) {
  for (let col of ENROLLMENTS_REQUIRED_BASE) {
    const v = sheet.getRange(row, col).getValue();
    if (v === "" || v === null) return false;
  }

  // Periods debe ser 1 o 2
  const periods = sheet.getRange(row, 6).getValue(); // F
  if (!(periods === 1 || periods === 2)) return false;

  // State recomendado: Active / Inactive
  const state = String(sheet.getRange(row, 8).getValue() || "").trim(); // H
  if (state !== "Active" && state !== "Inactive") return false;

  return true;
}

/************************************
 * GENERAR ID AFM-EYYYYXXX
 ************************************/
function generateNextEnrollmentId(sheet) {
  const year = new Date().getFullYear();
  const prefix = "AFM-E";
  const seqDigits = 3;

  const ids = sheet
    .getRange(ENROLLMENTS_FIRST_DATA_ROW, ENROLLMENTS_ID_COLUMN, sheet.getLastRow() - ENROLLMENTS_FIRST_DATA_ROW + 1, 1)
    .getValues()
    .flat()
    .filter(v => typeof v === "string" && v.startsWith(prefix));

  let next = 1;

  if (ids.length > 0) {
    const lastId = ids[ids.length - 1];

    // AFM-E2026001
    const match = lastId.match(/^AFM-E(\d{4})(\d{3})$/);
    if (match) {
      const lastYear = Number(match[1]);
      const lastSeq = Number(match[2]);
      next = (lastYear === year) ? lastSeq + 1 : 1;
    }
  }

  return `${prefix}${year}${String(next).padStart(seqDigits, "0")}`;
}
