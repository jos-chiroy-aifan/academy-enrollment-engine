/****************************************************************************
 * MÓDULO INDEPENDIENTE DE REINTENTOS Y DIAGNÓSTICO DE ESTUDIANTES
 * - No modifica el flujo automático actual.
 * - Permite reintentar una fila específica o todas las marcadas como "Approved".
 * - Si el alumno fue borrado de 'Students', lo vuelve a crear sin fallar.
 * - Muestra en consola el motivo exacto de cualquier rechazo.
 * - Soporta Llave Compuesta (Nombre + Apellido + Correo) para hermanos.
 * - Apunta a la columna real AB (índice 27) para el Status.
 ****************************************************************************/

const RETRY_CONFIG = {
  MAX_ROOMS: Number(PropertiesService.getScriptProperties().getProperty("MAX_PHYSICAL_ROOMS")) || 6, // Capacidad máxima de salones físicos simultáneos
  PRIORITY_PROF_ID: PropertiesService.getScriptProperties().getProperty("PRIORITY_PROFESSOR_ID") || "AFM-P2026001"
};

/**
 * OPCIÓN 0: Reintentar una fila fija específica (Muy útil para depurar sin tocar Sheets).
 */
function retrySpecificRow() {
  const rowToRetry = 93; // 👈 Cambia este número por la fila que quieras probar
  Logger.log(`🚀 Iniciando reintento manual para la fila fija: ${rowToRetry}`);
  processSingleStudentRowRetry_(rowToRetry);
}

/**
 * OPCIÓN 1: Reintentar ÚNICAMENTE la fila que tienes seleccionada actualmente en Sheets.
 */
function retrySelectedStudentRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "Student Registration - Responses") {
    Logger.log("⚠️ Debes estar ubicado en la hoja 'Student Registration - Responses'.");
    return;
  }
  const rowNumber = sheet.getActiveCell().getRow();
  if (rowNumber < 2) {
    Logger.log("⚠️ Selecciona una fila de datos válida (fila 2 en adelante).");
    return;
  }
  Logger.log(`🚀 Iniciando reintento manual para la fila seleccionada: ${rowNumber}`);
  processSingleStudentRowRetry_(rowNumber);
}

/**
 * OPCIÓN 2: Reintentar TODAS las filas que digan "Approved" en la hoja de respuestas.
 */
function retryAllApprovedStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName("Student Registration - Responses");
  const lastRow = rSheet.getLastRow();
  const lastCol = rSheet.getLastColumn();

  if (lastRow < 2) {
    Logger.log("⚠️ No hay filas que procesar.");
    return;
  }

  const headers = rSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Búsqueda flexible que ignora espacios accidentales
  const findCol = (name, fallback) => {
    const cleanTarget = name.toLowerCase().replace(/[\s_\-]/g, '');
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).toLowerCase().replace(/[\s_\-]/g, '') === cleanTarget) return i;
    }
    return fallback;
  };

  const statusCol = findCol("Status", 27); // Columna AB = idx 27 real

  const statusValues = rSheet.getRange(2, statusCol + 1, lastRow - 1, 1).getValues();
  let count = 0;

  for (let i = 0; i < statusValues.length; i++) {
    const currentStatus = String(statusValues[i][0]).trim();
    if (currentStatus === "Approved") {
      const rowNum = i + 2;
      Logger.log(`\n======================================================`);
      Logger.log(`🔄 Procesando fila ${rowNum} con estatus 'Approved'...`);
      processSingleStudentRowRetry_(rowNum);
      count++;
    }
  }

  Logger.log(`\n🏁 Reintento masivo finalizado. Total procesadas: ${count}`);
}

/**
 * NÚCLEO DEL REINTENTO PARA UNA FILA INDIVIDUAL
 */
function processSingleStudentRowRetry_(rowNumber) {
  const log = createRunLogger_("REINTENTO");
  log.set("fila", rowNumber);

  let outcome;
  try {
    outcome = runSingleStudentRowRetry_(rowNumber, log);
  } catch (err) {
    log.error("Excepción no controlada. El reintento se detuvo: revisa cuál fue el último PASO registrado arriba.", err);
    log.result("FALLO", `Excepción: ${err && err.message ? err.message : err}`);
    throw err;
  }
  log.result(outcome.result, outcome.detail);
}

function runSingleStudentRowRetry_(rowNumber, log) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rSheet = ss.getSheetByName("Student Registration - Responses");
  const sSheet = ss.getSheetByName("Students");
  const eSheet = ss.getSheetByName("Student_Enrollments");
  const progSheet = ss.getSheetByName("Programs");
  const profSheet = ss.getSheetByName("Professors");
  const ppSheet = ss.getSheetByName("Professor_Programs");
  const paSheet = ss.getSheetByName("Professor_Availability");

  const lastCol = rSheet.getLastColumn();
  const headers = rSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Buscador de columnas insensible a mayúsculas, tildes o espacios extras
  const findCol = (name, fallback) => {
    const cleanTarget = name.toLowerCase().replace(/[\s_\-]/g, '');
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).toLowerCase().replace(/[\s_\-]/g, '') === cleanTarget) return i;
    }
    return fallback;
  };

  const I = {
    timestamp: 0,
    firstName: findCol("Nombre(s)", 1),          // Col B (idx 1)
    lastName: findCol("Apellido(s)", 2),          // Col C (idx 2)
    email: findCol("Correo electrónico", 4),      // Col E (idx 4)
    programName: findCol("Programa", 15),         // Col P (idx 15)
    status: findCol("Status", 27)                 // Col AB (idx 27 real)
  };

  const rowValues = rSheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const email = String(rowValues[I.email] || "").trim();
  const firstName = String(rowValues[I.firstName] || "").trim();
  const lastName = String(rowValues[I.lastName] || "").trim();
  const programName = String(rowValues[I.programName] || "").trim();

  const currentStatus = String(rowValues[I.status] || "").trim();
  log.info(`Alumno: "${firstName} ${lastName}" | Correo: ${email} | Programa: ${programName} | Estatus al iniciar: "${currentStatus || "(vacío)"}"`);

  // -------------------------------------------------------------
  // PASO 1: OBTENER O RE-CREAR ESTUDIANTE (LLAVE COMPUESTA)
  // -------------------------------------------------------------
  let studentData = findStudentByEmailAndName_(sSheet, email, firstName, lastName);

  if (!studentData) {
    log.info("PASO 1: el alumno no existe en 'Students' (posiblemente borrado). Creándolo nuevamente...");
    studentData = createStudentFromRowData_(sSheet, rowValues);
    if (!studentData) {
      log.error("PASO 1: falló la recreación del estudiante.");
      rSheet.getRange(rowNumber, I.status + 1).setValue("Error: Cannot Recreate Student");
      return { result: "FALLO", detail: "PASO 1 | Error: Cannot Recreate Student" };
    }
    log.info(`✅ PASO 1: alumno recreado con éxito: ID ${studentData.id} (${studentData.fullName})`);
  } else {
    log.info(`✅ PASO 1: alumno encontrado en base de datos: ID ${studentData.id} (${studentData.fullName})`);
  }
  log.set("alumno", studentData.id);

  // -------------------------------------------------------------
  // PASO 2: VALIDAR PROGRAMA Y PARSEAR HORARIO
  // -------------------------------------------------------------
  const programInfo = getProgramInfoExtended_(progSheet, programName);
  if (programInfo.id === "UNKNOWN") {
    log.error(`PASO 2: programa "${programName}" no existe en la hoja 'Programs'. Revisa tildes.`);
    rSheet.getRange(rowNumber, I.status + 1).setValue("Pending Program Error");
    return { result: "FALLO", detail: `PASO 2 | Pending Program Error | programa "${programName}"` };
  }

  const scheduleRaw = pickScheduleValueByProgram_(rowValues, headers, programName);
  const parsed = parseSchedule_(scheduleRaw);
  if (!parsed || !parsed.dayOfWeek || !parsed.timeSlot) {
    log.error(`PASO 2: formato de horario inválido. Texto leído: "${scheduleRaw}"`);
    rSheet.getRange(rowNumber, I.status + 1).setValue("Pending Schedule Error");
    return { result: "FALLO", detail: `PASO 2 | Pending Schedule Error | horario "${scheduleRaw}"` };
  }

  const dayKey = parsed.dayOfWeek;
  const timeSlot = parsed.timeSlot;
  const periods = 1;
  log.info(`✅ PASO 2: programa "${programName}" (${programInfo.id}) | horario leído: "${scheduleRaw}" → ${dayKey} ${timeSlot}`);

  // -------------------------------------------------------------
  // GUARDIA: ¿el alumno YA tiene esta misma inscripción activa?
  // Sin esto, el profesor aparece ocupado por la propia clase del
  // alumno y el reintento marca un falso "Pending Capacity Full".
  // -------------------------------------------------------------
  const studentEnrollments = findActiveEnrollmentsForStudent_(eSheet, studentData.id);
  log.info(`Inscripciones activas previas del alumno: ${studentEnrollments.length ? studentEnrollments.map(formatEnrollment_).join(" ; ") : "ninguna"}`);
  const existingEnrollment = studentEnrollments.find(x => x.programId === programInfo.id && x.dayKey === dayKey && x.timeSlot === timeSlot);
  if (existingEnrollment) {
    log.warn(`El alumno YA está inscrito en este programa y horario (${existingEnrollment.enrollmentId}). No se crea otra inscripción ni se envían correos. Estatus: "${currentStatus || "(vacío)"}" → "Migrated"`);
    rSheet.getRange(rowNumber, I.status + 1).setValue("Migrated");
    return { result: "OMITIDO", detail: `Ya inscrito: ${existingEnrollment.enrollmentId}. Fila marcada Migrated sin crear nada` };
  }

  // -------------------------------------------------------------
  // PASO 3: VALIDACIÓN DE CAPACIDAD FÍSICA (SALONES)
  // -------------------------------------------------------------
  const professorsActive = buildActiveProfessors_(profSheet);
  const professorPrograms = buildProfessorPrograms_(ppSheet, professorsActive);
  const availability = buildAvailability_(paSheet, professorsActive);
  const timeSlotOrder = buildTimeSlotOrder_(paSheet);
  const nextTimeSlotFn = (slot) => {
    const i = timeSlotOrder.indexOf(slot);
    return (i >= 0 && i < timeSlotOrder.length - 1) ? timeSlotOrder[i + 1] : null;
  };

  const usage = buildActiveEnrollmentUsage_(eSheet, nextTimeSlotFn);

  const slotKey = `${dayKey}|${timeSlot}`;
  const occupiedRooms = usage.roomsBySlot?.[slotKey] || 0;

  log.info(`PASO 3: salones ocupados en ${slotKey}: ${occupiedRooms}/${RETRY_CONFIG.MAX_ROOMS}`);

  if (occupiedRooms >= RETRY_CONFIG.MAX_ROOMS) {
    log.error(`PASO 3 (Salones Llenos): ya se alcanzó el límite de ${RETRY_CONFIG.MAX_ROOMS} en ${slotKey}. Ocupan el bloque: ${listEnrollmentsAtSlot_(eSheet, dayKey, timeSlot).map(formatEnrollment_).join(" ; ")}`);
    rSheet.getRange(rowNumber, I.status + 1).setValue("Pending Capacity Full");
    return { result: "FALLO", detail: `PASO 3 | Pending Capacity Full | motivo: salones llenos (${occupiedRooms}/${RETRY_CONFIG.MAX_ROOMS}) en ${slotKey}` };
  }

  // -------------------------------------------------------------
  // PASO 4: BUSCAR PROFESOR DISPONIBLE
  // -------------------------------------------------------------
  const candidates = getCandidateProfessorsWithCapacity_({
    programId: programInfo.id,
    dayKey: dayKey,
    timeSlot: timeSlot,
    periods: periods,
    nextTimeSlotFn: nextTimeSlotFn,
    professorPrograms: professorPrograms,
    availability: availability,
    usage: usage,
  });

  logCandidateDiagnosis_(log, { programId: programInfo.id, dayKey, timeSlot, professorPrograms, availability, usage });

  if (candidates.length === 0) {
    const occupants = listEnrollmentsAtSlot_(eSheet, dayKey, timeSlot);
    log.error(`PASO 4 (Sin Profesores): ningún profesor activo para "${programName}" tiene libre el bloque ${slotKey}. Clases activas en ese bloque: ${occupants.length ? occupants.map(formatEnrollment_).join(" ; ") : "ninguna"}`);
    rSheet.getRange(rowNumber, I.status + 1).setValue("Pending Capacity Full");
    return { result: "FALLO", detail: "PASO 4 | Pending Capacity Full | motivo: ningún profesor libre" };
  }


  let assignedProfessorId;
  if (candidates.includes(RETRY_CONFIG.PRIORITY_PROF_ID)) {
    assignedProfessorId = RETRY_CONFIG.PRIORITY_PROF_ID;
  } else {
    assignedProfessorId = candidates[Math.floor(Math.random() * candidates.length)];
  }

  const professorFullName = getProfessorFullName_(profSheet, assignedProfessorId);
  log.info(`✅ PASO 4: profesor asignado -> ${assignedProfessorId} (${professorFullName}) | regla: ${assignedProfessorId === RETRY_CONFIG.PRIORITY_PROF_ID ? "prioridad" : `aleatorio entre ${candidates.length} candidato(s)`}`);

  // -------------------------------------------------------------
  // PASO 5: CREAR INSCRIPCIÓN (ENROLLMENT)
  // -------------------------------------------------------------
  const now = new Date();
  const currentYear = now.getFullYear();
  const nextEnrSequence = getLastEnrollmentSequence_(eSheet);
  const enrollmentId = `AFM-E${currentYear}${String(nextEnrSequence).padStart(3, "0")}`;
  const payDay = rowValues[I.timestamp];

  // Mapeo exacto de las 17 columnas (A a Q) para Student_Enrollments
  const newEnrollmentRow = [
    enrollmentId,
    studentData.id,
    studentData.fullName,
    programInfo.id,
    programName,
    dayKey,
    timeSlot,
    periods,
    assignedProfessorId,
    professorFullName,
    "Active",
    "Form-Retry",
    now,
    now,
    programInfo.price,
    payDay,
    "Reintento manual ejecutado"
  ];

  const eTargetRow = getFirstRealEmptyRowMain_(eSheet, "A");
  eSheet.getRange(eTargetRow, 1, 1, newEnrollmentRow.length).setValues([newEnrollmentRow]);
  log.info(`✅ PASO 5: inscripción registrada en 'Student_Enrollments' fila ${eTargetRow} con ID ${enrollmentId}`);

  // Marcar como Migrated en la columna AB (idx 27)
  rSheet.getRange(rowNumber, I.status + 1).setValue("Migrated");
  log.info(`Estatus: "${currentStatus || "(vacío)"}" → "Migrated"`);

  // -------------------------------------------------------------
  // PASO 6: ENVÍO DE NOTIFICACIONES
  // -------------------------------------------------------------
  log.info(`PASO 6: enviando correos | cuota de correos restante hoy: ${MailApp.getRemainingDailyQuota()}`);
  try {
    if (typeof notifyProfessorNewStudent_ === 'function') {
      const sentTo = notifyProfessorNewStudent_({
        professorId: assignedProfessorId,
        studentId: studentData.id,
        programId: programInfo.id,
        dayKey: dayKey,
        timeSlot: timeSlot,
        periods: periods
      });
      if (sentTo) log.info(`PASO 6: correo entregado a Gmail para el profesor ${assignedProfessorId} (${sentTo}).`);
      else log.warn(`PASO 6: NO se envió correo al profesor ${assignedProfessorId}: no tiene correo o no está Active en 'Professors'.`);
    } else {
      log.warn("PASO 6: no existe notifyProfessorNewStudent_; no se notificó al profesor.");
    }
  } catch (err) { log.error("PASO 6: error enviando correo al profesor", err); }

  try {
    if (typeof sendScheduleConfirmationEmail_ === 'function') {
      sendScheduleConfirmationEmail_({
        email: email,
        studentId: studentData.id,
        firstName: firstName,
        lastName: lastName,
        programName: programName,
        dayOfWeek: dayKey,
        timeSlot: timeSlot,
        periods: periods,
        professorName: professorFullName,
        notes: ""
      });
      log.info(`PASO 6: correo de confirmación entregado a Gmail para el alumno (${email}).`);
    } else {
      log.warn("PASO 6: no existe sendScheduleConfirmationEmail_; no se notificó al alumno.");
    }
  } catch (err) { log.error("PASO 6: error enviando correo al alumno", err); }

  return { result: "OK", detail: `${enrollmentId} | ${studentData.id} | ${programName} | ${dayKey} ${timeSlot} | prof ${assignedProfessorId}` };
}

/**
 * HELPER: Crea al estudiante en caso de haber sido eliminado de 'Students'
 */
function createStudentFromRowData_(sSheet, row) {
  const COL = {
    firstName: 2, lastName: 3, gender: 4, email: 5, phone: 6,
    day: 7, month: 8, year: 9, minor: 10,
    respName: 11, relationship: 12, respPhone: 13,
    consentPhotos: 14, terms: 15
  };

  const nextSequence = getLastSequence(sSheet);
  const currentYear = new Date().getFullYear();
  const newStudentId = `AFM-S${currentYear}${String(nextSequence).padStart(3, "0")}`;

  const birthdate = formatBirthdate(row[COL.day - 1], row[COL.month - 1], row[COL.year - 1]);
  const genderMap = { "Masculino": "Male", "Femenino": "Female", "Prefiero no decirlo": "Prefer not to say" };
  const gender = genderMap[row[COL.gender - 1]] || "Prefer not to say";

  // Mapeo exacto de las 17 columnas de 'Students' (A a Q)
  const studentRow = [
    newStudentId,
    String(row[COL.firstName - 1] || "").trim(),
    String(row[COL.lastName - 1] || "").trim(),
    birthdate,
    "", // Age formula
    row[COL.minor - 1] === "Sí" ? "Yes" : "No",
    gender,
    String(row[COL.email - 1] || "").trim(),
    String(row[COL.phone - 1] || "").trim(),
    row[COL.respName - 1] || "-",
    row[COL.relationship - 1] || "-",
    row[COL.respPhone - 1] || "-",
    String(row[COL.consentPhotos - 1]).toLowerCase().includes("sí"),
    String(row[COL.terms - 1]).toLowerCase().includes("acepto"),
    "Active",
    "Form-Retry",
    ""
  ];

  const sTargetRow = getFirstRealEmptyRow(sSheet, "A");
  sSheet.getRange(sTargetRow, 1, 1, studentRow.length).setValues([studentRow]);
  sSheet.getRange(sTargetRow, 5).setFormula(`=IF(D${sTargetRow}="","",DATEDIF(D${sTargetRow},TODAY(),"Y"))`);

  return {
    id: newStudentId,
    fullName: `${row[COL.firstName - 1]} ${row[COL.lastName - 1]}`.trim()
  };
}