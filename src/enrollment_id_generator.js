/************************************
 * CONFIGURACIÓN BASE DE ENROLLMENTS
 ************************************/
const ENROLLMENTS_SHEET_NAME = 'Student_Enrollments';
const ENROLLMENTS_ID_COLUMN = 1;       // A = EnrollmentID
const ENROLLMENTS_FIRST_DATA_ROW = 2;

// Columnas obligatorias base (estructura actual A-Q):
// B StudentID, D ProgramID, F DayOfWeek, G TimeSlot, H Periods, I AssignedProfessorID, K State
// (C Student, E Program y J Assigned Professor son fórmulas/lookup: no se exigen)
const ENROLLMENTS_REQUIRED_BASE = [2, 4, 6, 7, 8, 9, 11];
const ENROLLMENTS_PERIODS_COLUMN = 8;  // H = Periods
const ENROLLMENTS_STATE_COLUMN = 11;   // K = State

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
  const periods = sheet.getRange(row, ENROLLMENTS_PERIODS_COLUMN).getValue(); // H
  if (!(periods === 1 || periods === 2)) return false;

  // State recomendado: Active / Inactive
  const state = String(sheet.getRange(row, ENROLLMENTS_STATE_COLUMN).getValue() || "").trim(); // K
  if (state !== "Active" && state !== "Inactive") return false;

  return true;
}

/************************************
 * GENERAR ID AFM-EYYYYXXX
 * Usa el MAYOR correlativo del año (igual que la migración automática,
 * getLastEnrollmentSequence_), no el último de la hoja: si las filas
 * están desordenadas, el último no es el mayor y se repetirían IDs.
 ************************************/
function generateNextEnrollmentId(sheet) {
  const year = new Date().getFullYear();
  const prefix = "AFM-E";
  const seqDigits = 3;

  const next = getLastEnrollmentSequence_(sheet); // definido en schedule_form_response_migration.js

  return `${prefix}${year}${String(next).padStart(seqDigits, "0")}`;
}
