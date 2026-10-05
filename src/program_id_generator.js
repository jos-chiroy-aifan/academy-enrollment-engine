/************************************
 * CONFIGURACIÓN BASE DE PROGRAMAS
 ************************************/
const PROGRAM_SHEET_NAME = 'Programs';
const PROGRAM_ID_COLUMN = 1;       // Columna A = ProgramID
const PROGRAM_FIRST_DATA_ROW = 2;  // Fila de inicio de datos

// Columnas requeridas para generar ID
// B = Program Name, C = Category, D = Level, E = State
const PROGRAM_REQUIRED_COLUMNS = {
  name: 2,
  category: 3,
  level: 4,
  state: 5
};

const PROGRAM_ID_PREFIX = 'AFM-PRG';
const PROGRAM_SEQ_DIGITS = 3;      // AFMA-PRG001, AFMA-PRG002, ...

/************************************
 * FUNCIÓN PRINCIPAL
 * Recorre la hoja Programs y genera
 * IDs solo para filas completas y activas
 ************************************/
function generateProgramIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PROGRAM_SHEET_NAME);
  if (!sheet) throw new Error('Hoja "Programs" no encontrada.');

  const lastRow = sheet.getLastRow();
  if (lastRow < PROGRAM_FIRST_DATA_ROW) return;

  for (let row = PROGRAM_FIRST_DATA_ROW; row <= lastRow; row++) {
    const currentId = sheet.getRange(row, PROGRAM_ID_COLUMN).getValue();
    if (currentId) continue; // Ya tiene ProgramID → no tocar

    // Solo generar ID si la fila está completa y en estado Active
    if (isProgramRowReady(sheet, row)) {
      const newId = generateNextProgramId(sheet);
      sheet.getRange(row, PROGRAM_ID_COLUMN).setValue(newId);
    }
  }
}

/************************************
 * VERIFICA SI LA FILA ESTÁ LISTA
 * - Program Name lleno
 * - Category llena
 * - Level lleno
 * - State = "Active"
 ************************************/
function isProgramRowReady(sheet, row) {
  const name = sheet.getRange(row, PROGRAM_REQUIRED_COLUMNS.name).getValue();
  const category = sheet.getRange(row, PROGRAM_REQUIRED_COLUMNS.category).getValue();
  const level = sheet.getRange(row, PROGRAM_REQUIRED_COLUMNS.level).getValue();
  const state = sheet.getRange(row, PROGRAM_REQUIRED_COLUMNS.state).getValue();

  // Todos los campos requeridos deben tener valor
  if (!name || !category || !level || !state) {
    return false;
  }

  // Solo generar ID si el estado es EXACTAMENTE "Active"
  if (state !== 'Active') {
    return false;
  }

  return true;
}

/************************************
 * GENERA EL SIGUIENTE ID DEL TIPO:
 *   AFMA-PRG001
 *   AFMA-PRG002
 *   AFMA-PRG003
 ************************************/
function generateNextProgramId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < PROGRAM_FIRST_DATA_ROW) {
    // Primer programa
    return `${PROGRAM_ID_PREFIX}${String(1).padStart(PROGRAM_SEQ_DIGITS, '0')}`;
  }

  const idValues = sheet
    .getRange(
      PROGRAM_FIRST_DATA_ROW,
      PROGRAM_ID_COLUMN,
      lastRow - PROGRAM_FIRST_DATA_ROW + 1
    )
    .getValues()
    .flat();

  let maxSeq = 0;

  idValues.forEach(id => {
    if (id && typeof id === 'string' && id.startsWith(PROGRAM_ID_PREFIX)) {
      const seqStr = id.slice(PROGRAM_ID_PREFIX.length); // después de "AFMA-PRG"
      const seqNum = parseInt(seqStr, 10);
      if (!isNaN(seqNum) && seqNum > maxSeq) {
        maxSeq = seqNum;
      }
    }
  });

  const nextSeq = maxSeq + 1;
  return `${PROGRAM_ID_PREFIX}${String(nextSeq).padStart(PROGRAM_SEQ_DIGITS, '0')}`;
}
