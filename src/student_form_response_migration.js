/************************************
 * MIGRACIÓN DE ESTUDIANTE (MODULO)
 * - Procesa Datos Personales y Genera Carnet.
 * - Soporta hermanos con el mismo correo mediante verificación de Nombre + Apellido.
 * - Escribe el estado en la columna AB (Status) real.
 ************************************/

function processStudentCreationOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName("Student Registration - Responses");
  const sSheet = ss.getSheetByName("Students");

  const lastRow = rSheet.getLastRow();
  if (lastRow < 2) return null;

  // MAPEO EXACTO SEGÚN EL SCHEMA REAL (Base 1 para getRange, Base 0 para arrays)
  const COL = {
    timestamp: 1,      // A (idx 0)
    firstName: 2,      // B (idx 1)
    lastName: 3,       // C (idx 2)
    gender: 4,         // D (idx 3)
    email: 5,          // E (idx 4)
    phone: 6,          // F (idx 5)
    day: 7,            // G (idx 6)
    month: 8,          // H (idx 7)
    year: 9,           // I (idx 8)
    minor: 10,         // J (idx 9)
    respName: 11,      // K (idx 10)
    relationship: 12,  // L (idx 11)
    respPhone: 13,     // M (idx 12)
    consentPhotos: 14, // N (idx 13)
    terms: 15,         // O (idx 14)
    program: 16,       // P (idx 15)
    status: 28         // AB (idx 27) -> ¡Columna 28 real, no 27!
  };

  const row = rSheet.getRange(lastRow, 1, 1, rSheet.getLastColumn()).getValues()[0];

  const currentStatus = String(row[COL.status - 1] || "").trim();
  if (currentStatus === "Migrated" || currentStatus === "Approved") return null;

  const firstName = String(row[COL.firstName - 1] || "").trim();
  const lastName = String(row[COL.lastName - 1] || "").trim();
  const email = String(row[COL.email - 1] || "").trim();

  // 1. REVISAR SI EL ESTUDIANTE YA EXISTE POR LLAVE COMPUESTA (Nombre + Apellido + Correo)
  const existingStudent = findStudentByEmailAndName_(sSheet, email, firstName, lastName);
  if (existingStudent) {
    rSheet.getRange(lastRow, COL.status).setValue("Approved");
    Logger.log(`ℹ️ Estudiante ya registrado previamente: ${existingStudent.id} (${existingStudent.fullName})`);
    return existingStudent;
  }

  // 2. GENERACIÓN DEL CARNET (ID)
  const nextSequence = getLastSequence(sSheet);
  const currentYear = new Date().getFullYear();
  const newStudentId = `AFM-S${currentYear}${String(nextSequence).padStart(3, "0")}`;

  // 3. PROCESAMIENTO DE DATOS PERSONALES
  const birthdate = formatBirthdate(row[COL.day - 1], row[COL.month - 1], row[COL.year - 1]);
  const genderMap = { "Masculino": "Male", "Femenino": "Female", "Prefiero no decirlo": "Prefer not to say" };
  const gender = genderMap[row[COL.gender - 1]] || "Prefer not to say";

  // Mapeo exacto a las 17 columnas de "Students" (A a Q)
  const studentRow = [
    newStudentId,                                                   // A: StudentID
    firstName,                                                      // B: First Name
    lastName,                                                       // C: Last Name
    birthdate,                                                      // D: Birthdate
    "",                                                             // E: Age (Fórmula)
    row[COL.minor - 1] === "Sí" ? "Yes" : "No",                     // F: Minor
    gender,                                                         // G: Gender
    email,                                                          // H: Email
    String(row[COL.phone - 1] || ""),                               // I: Phone
    row[COL.respName - 1] || "-",                                   // J: Responsible Name
    row[COL.relationship - 1] || "-",                               // K: Relationship
    row[COL.respPhone - 1] || "-",                                  // L: Responsible Phone
    String(row[COL.consentPhotos - 1]).toLowerCase().includes("sí"),// M: Consent for Photos
    String(row[COL.terms - 1]).toLowerCase().includes("acepto"),    // N: Consent for Data Processing
    "Active",                                                       // O: State
    "Form",                                                         // P: Source
    ""                                                              // Q: Notes
  ];

  const sTargetRow = getFirstRealEmptyRow(sSheet, "A");
  sSheet.getRange(sTargetRow, 1, 1, studentRow.length).setValues([studentRow]);
  sSheet.getRange(sTargetRow, 5).setFormula(`=IF(D${sTargetRow}="","",DATEDIF(D${sTargetRow},TODAY(),"Y"))`);

  // Escribir en la columna AB (Status) real
  rSheet.getRange(lastRow, COL.status).setValue("Approved");
  Logger.log(`✅ Estudiante nuevo creado: ${newStudentId} - ${firstName} ${lastName}`);

  return {
    id: newStudentId,
    fullName: `${firstName} ${lastName}`.trim()
  };
}

function getLastSequence(sheet) {
  const currentYear = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxSeq = 0;
  ids.forEach(id => {
    if (String(id).startsWith(`AFM-S${currentYear}`)) {
      const seq = Number(String(id).slice(-3));
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  return maxSeq + 1;
}

function formatBirthdate(day, monthText, year) {
  const months = { "Enero": 0, "Febrero": 1, "Marzo": 2, "Abril": 3, "Mayo": 4, "Junio": 5, "Julio": 6, "Agosto": 7, "Septiembre": 8, "Octubre": 9, "Noviembre": 10, "Diciembre": 11 };
  return new Date(year, months[monthText] || 0, day);
}

function getFirstRealEmptyRow(sheet, columnLetter) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const values = sheet.getRange(columnLetter + "2:" + columnLetter + lastRow).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "") return i + 2;
  }
  return lastRow + 1;
}