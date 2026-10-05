/****************************************************************************
 * MIGRACIÓN PRINCIPAL (SCRIPT 2)
 * 1. Llama a la creación de estudiante (o lo recupera si ya existe por Llave Compuesta).
 * 2. Valida capacidad de salones físicos (máximo 6 simultáneos).
 * 3. Asigna Profesor sin doble descuento y Horario.
 * 4. Crea la inscripción con Costos y Nombres Completos (Mapeo A-Q exacto).
 ****************************************************************************/

function processUnifiedMainMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- HOJAS REQUERIDAS ---
  const rSheet = ss.getSheetByName("Student Registration - Responses");
  const sSheet = ss.getSheetByName("Students");
  const eSheet = ss.getSheetByName("Student_Enrollments");
  const progSheet = ss.getSheetByName("Programs");
  const profSheet = ss.getSheetByName("Professors");
  const ppSheet = ss.getSheetByName("Professor_Programs");
  const paSheet = ss.getSheetByName("Professor_Availability");

  if (!rSheet || !sSheet || !eSheet || !progSheet || !profSheet || !ppSheet || !paSheet) {
    throw new Error("Faltan hojas requeridas en la base de datos.");
  }

  const lastRow = rSheet.getLastRow();
  const lastCol = rSheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log("⚠️ No hay filas de datos suficientes en 'Student Registration - Responses'.");
    return;
  }

  // --- MAPEO DE COLUMNAS (Para leer la respuesta sin fallos por espacios) ---
  const headers = rSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const findCol = (name, fallback) => {
    const cleanTarget = name.toLowerCase().replace(/[\s_\-]/g, '');
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).toLowerCase().replace(/[\s_\-]/g, '') === cleanTarget) return i;
    }
    return fallback;
  };

  const I = {
    timestamp: 0,
    firstName: findCol("Nombre(s)", 1),
    lastName: findCol("Apellido(s)", 2),
    email: findCol("Correo electrónico", 4),
    programName: findCol("Programa", 15),
    status: findCol("Status", 27) // Columna AB (idx 27 real)
  };

  const rowValues = rSheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
  const email = String(rowValues[I.email] || "").trim();
  const firstName = String(rowValues[I.firstName] || "").trim();
  const lastName = String(rowValues[I.lastName] || "").trim();
  const currentStatus = String(rowValues[I.status] || "").trim();

  Logger.log(`ℹ️ Procesando última fila (${lastRow}) | Alumno: "${firstName} ${lastName}" | Correo: "${email}" | Estatus actual: "${currentStatus}"`);

  // Si ya está terminado, salir
  if (currentStatus === "Migrated") {
    Logger.log(`ℹ️ La fila ${lastRow} ya se encuentra como 'Migrated'. No hay nada que procesar.`);
    return;
  }

  // =================================================================
  // PASO 1: OBTENER ESTUDIANTE (Llamando al Script 1 o Llave Compuesta)
  // =================================================================
  let studentData = null;

  // Intentamos crear el estudiante
  // Si devuelve objeto, es nuevo. Si devuelve null/undefined, es que ya estaba "Approved".
  studentData = processStudentCreationOnly();

  // Si processStudentCreationOnly no devolvió nada (porque ya estaba Approved),
  // buscamos al estudiante por LLAVE COMPUESTA (Email + Nombre + Apellido) para no pisar hermanos.
  if (!studentData) {
    studentData = findStudentByEmailAndName_(sSheet, email, firstName, lastName);

    if (!studentData) {
      Logger.log(`❌ ERROR PASO 1: Estudiante "${firstName} ${lastName}" (${email}) figuraba como Approved pero no se encontró en la hoja 'Students'.`);
      rSheet.getRange(lastRow, I.status + 1).setValue("Error: Student Approved but Not Found");
      return;
    }
  }

  const studentId = studentData.id;
  const studentFullName = studentData.fullName;
  Logger.log(`✅ PASO 1 Exitoso: Estudiante listo -> ID: ${studentId} | Nombre: ${studentFullName}`);

  // =================================================================
  // PASO 2: PREPARAR DATOS DE PROGRAMA Y HORARIO
  // =================================================================
  const programName = String(rowValues[I.programName] || "").trim();

  // Usamos la función auxiliar para traer ID y PRECIO
  const programInfo = getProgramInfoExtended_(progSheet, programName);

  if (programInfo.id === "UNKNOWN") {
    Logger.log(`❌ ERROR PASO 2: Programa no encontrado en catálogo -> "${programName}". Verifique tildes o espacios exactos en 'Programs'.`);
    rSheet.getRange(lastRow, I.status + 1).setValue("Pending Program Error");
    return;
  }

  // Detectar Horario en columnas Q-Z (idx 16 a 25)
  const scheduleRaw = pickScheduleValueByProgram_(rowValues, headers, programName);
  const parsed = parseSchedule_(scheduleRaw);

  if (!parsed) {
    Logger.log(`❌ ERROR PASO 2: No se pudo parsear el horario. Texto leído en celda: "${scheduleRaw}"`);
    rSheet.getRange(lastRow, I.status + 1).setValue("Pending Schedule Error");
    return;
  }

  const dayKey = parsed.dayOfWeek;
  const timeSlot = parsed.timeSlot;
  const periods = 1; // Por defecto 1 periodo
  Logger.log(`✅ PASO 2 Exitoso: Programa ID: ${programInfo.id} | Día normalizado: ${dayKey} | Horario: ${timeSlot}`);

  // =================================================================
  // PASO 3: LÓGICA DE ASIGNACIÓN
  // =================================================================

  // Construir índices
  const professorsActive = buildActiveProfessors_(profSheet);
  const professorPrograms = buildProfessorPrograms_(ppSheet, professorsActive);
  const availability = buildAvailability_(paSheet, professorsActive);
  const timeSlotOrder = buildTimeSlotOrder_(paSheet);

  const nextTimeSlotFn = (slot) => {
    const i = timeSlotOrder.indexOf(slot);
    return (i >= 0 && i < timeSlotOrder.length - 1) ? timeSlotOrder[i + 1] : null;
  };

  const usage = buildActiveEnrollmentUsage_(eSheet, nextTimeSlotFn);

  // VALIDACIÓN DE SALONES FÍSICOS (MÁXIMO: Script Property MAX_PHYSICAL_ROOMS, por defecto 6)
  const maxPhysicalRooms = Number(PropertiesService.getScriptProperties().getProperty("MAX_PHYSICAL_ROOMS")) || 6;
  const roomUsageKey = `${dayKey}|${timeSlot}`;
  const occupiedRooms = usage.roomsBySlot?.[roomUsageKey] || 0;
  if (occupiedRooms >= maxPhysicalRooms) {
    Logger.log(`❌ ERROR PASO 3 (Salones Llenos): En el bloque ${roomUsageKey} ya hay ${occupiedRooms}/${maxPhysicalRooms} salones ocupados.`);
    rSheet.getRange(lastRow, I.status + 1).setValue("Pending Capacity Full");
    return;
  }

  // Buscar Candidatos Libres
  let candidates = getCandidateProfessorsWithCapacity_({
    programId: programInfo.id,
    dayKey: dayKey,
    timeSlot: timeSlot,
    periods: periods,
    nextTimeSlotFn: nextTimeSlotFn,
    professorPrograms: professorPrograms,
    availability: availability,
    usage: usage,
  });

  if (candidates.length === 0) {
    Logger.log(`❌ ERROR PASO 3 (Sin Profesores): No hay profesores activos disponibles para el programa "${programName}" (${programInfo.id}) en ${roomUsageKey}.`);
    rSheet.getRange(lastRow, I.status + 1).setValue("Pending Capacity Full");
    return;
  }

  // Prioridad
  const priorityProfId = PropertiesService.getScriptProperties().getProperty("PRIORITY_PROFESSOR_ID") || "AFM-P2026001";
  let assignedProfessorId;

  if (candidates.includes(priorityProfId)) {
    assignedProfessorId = priorityProfId;
  } else {
    assignedProfessorId = candidates[Math.floor(Math.random() * candidates.length)];
  }

  // Obtener Nombre del Profesor
  const professorFullName = getProfessorFullName_(profSheet, assignedProfessorId);
  Logger.log(`✅ PASO 3 Exitoso: Profesor asignado -> ${assignedProfessorId} (${professorFullName})`);

  // =================================================================
  // PASO 4: CREAR INSCRIPCIÓN (Enrollment)
  // =================================================================
  const now = new Date();
  const currentYear = new Date().getFullYear();
  const nextEnrSequence = getLastEnrollmentSequence_(eSheet);
  const enrollmentId = `AFM-E${currentYear}${String(nextEnrSequence).padStart(3, "0")}`;
  const payDay = rowValues[I.timestamp]; // Fecha del form

  // Mapeo exacto de 17 columnas (A a Q) para Student_Enrollments
  const newEnrollmentRow = [
    enrollmentId,       // A: EnrollmentID
    studentId,          // B: StudentID
    studentFullName,    // C: Student
    programInfo.id,     // D: ProgramID
    programName,        // E: Program
    dayKey,             // F: DayOfWeek
    timeSlot,           // G: TimeSlot
    periods,            // H: Periods
    assignedProfessorId,// I: AssignedProfessorID
    professorFullName,  // J: Assigned Professor
    "Active",           // K: State
    "Form",             // L: Source
    now,                // M: CreatedAt
    now,                // N: LastUpdatedAt
    programInfo.price,  // O: Monthly payment
    payDay,             // P: Payday
    ""                  // Q: Notes
  ];

  const eTargetRow = getFirstRealEmptyRowMain_(eSheet, "A");
  eSheet.getRange(eTargetRow, 1, 1, newEnrollmentRow.length).setValues([newEnrollmentRow]);
  Logger.log(`✅ PASO 4 Exitoso: Inscripción creada en 'Student_Enrollments' fila ${eTargetRow} con ID ${enrollmentId}`);

  // Actualizar uso en memoria
  incUsage_(usage, programInfo.id, dayKey, timeSlot, assignedProfessorId);

  // =================================================================
  // PASO 5: FINALIZAR
  // =================================================================

  // Marcamos como "Migrated" en la columna AB (idx 27)
  rSheet.getRange(lastRow, I.status + 1).setValue("Migrated");
  Logger.log(`🎉 Proceso completado exitosamente para la fila ${lastRow}. Estatus marcado como 'Migrated'.`);

  // Notificaciones
  try {
    if (typeof notifyProfessorNewStudent_ === 'function') {
      notifyProfessorNewStudent_({
        professorId: assignedProfessorId,
        studentId: studentId,
        programId: programInfo.id,
        dayKey: dayKey,
        timeSlot: timeSlot,
        periods: periods
      });
    }
  } catch (err) { Logger.log("Error correo profe: " + err); }

  try {
    if (typeof sendScheduleConfirmationEmail_ === 'function') {
      sendScheduleConfirmationEmail_({
        email: email,
        studentId: studentId,
        firstName: firstName,
        lastName: lastName,
        programName: programName,
        dayOfWeek: dayKey,
        timeSlot: timeSlot,
        periods: periods,
        professorName: professorFullName,
        notes: ""
      });
    }
  } catch (err) { Logger.log("Error correo alumno: " + err); }
}


/**
 * -------------------------------------------------------------------------
 * FUNCIONES AUXILIARES
 * -------------------------------------------------------------------------
 */

// Recupera ID y Precio del programa
function getProgramInfoExtended_(sheet, programName) {
  const last = sheet.getLastRow();
  if (last < 2) return { id: "UNKNOWN", price: 0 };
  const data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const normTarget = String(programName || "").trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    // Col A=ID (0), Col B=Name (1), Col F=Price (5)
    if (String(data[i][1]).trim().toLowerCase() === normTarget) {
      return { id: String(data[i][0]).trim(), price: Number(data[i][5]) || 0 };
    }
  }
  return { id: "UNKNOWN", price: 0 };
}

// Busca el horario en columnas Q a Z
function pickScheduleValueByProgram_(row, headers, programName) {
  const norm = (s) => String(s || "").trim().toLowerCase().replace(/[\s_\-]/g, '');
  const colName = norm(`Horarios - ${programName}`);

  for (let i = 0; i < headers.length; i++) {
    if (norm(headers[i]) === colName && row[i]) {
      return String(row[i]).trim();
    }
  }

  // Fallback: Busca entre columnas Q y Z (índices 16 a 25)
  for (let k = 16; k <= 25 && k < row.length; k++) {
    const val = String(row[k] || "");
    if ((val.includes("-") || val.includes("a")) && val.includes(":")) {
      return val.trim();
    }
  }
  return "";
}

function parseSchedule_(raw) {
  if (!raw) return null;
  const parts = String(raw).trim().split("-");
  if (parts.length < 2) return null;

  const day = normalizeDayKey_(parts[0]);
  const rawTime = parts.slice(1).join("-").trim();
  const time = normalizeTimeSlot_(rawTime);

  return { dayOfWeek: day, timeSlot: time };
}

// Recupera estudiante por Llave Compuesta: Email + Nombre + Apellido
function findStudentByEmailAndName_(sSheet, emailTarget, firstNameTarget, lastNameTarget) {
  const last = sSheet.getLastRow();
  if (last < 2) return null;

  const data = sSheet.getRange(2, 1, last - 1, sSheet.getLastColumn()).getValues();
  const norm = (s) => String(s || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  const tEmail = norm(emailTarget);
  const tFirst = norm(firstNameTarget);
  const tLast = norm(lastNameTarget);

  for (const r of data) {
    // Mapeo exacto de Students: Col A=ID (0), Col B=First (1), Col C=Last (2), Col H=Email (7)
    const rId = String(r[0]).trim();
    const rFirst = norm(r[1]);
    const rLast = norm(r[2]);
    const rEmail = norm(r[7]);

    if (rEmail === tEmail && rFirst === tFirst && rLast === tLast) {
      return {
        id: rId,
        fullName: `${r[1]} ${r[2]}`.trim()
      };
    }
  }
  return null;
}

function getProfessorFullName_(sheet, profId) {
  const last = sheet.getLastRow();
  if (last < 2) return profId;
  const data = sheet.getRange(2, 1, last - 1, 3).getValues();
  for (const r of data) {
    if (String(r[0]).trim() === profId) return `${r[1]} ${r[2]}`.trim();
  }
  return profId;
}

function getFirstRealEmptyRowMain_(sheet, colLetter) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const values = sheet.getRange(colLetter + "2:" + colLetter + lastRow).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "") return i + 2;
  }
  return lastRow + 1;
}

// --- FUNCIONES DE DISPONIBILIDAD Y OCUPACIÓN ---

function buildActiveProfessors_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return new Set();
  const vals = sheet.getRange(2, 1, last - 1, 11).getValues();
  const set = new Set();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][10]).trim() === "Active") set.add(String(vals[i][0]).trim());
  }
  return set;
}

function buildProfessorPrograms_(sheet, activeProfs) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const vals = sheet.getRange(2, 1, last - 1, 4).getValues();
  const out = {};
  for (let i = 0; i < vals.length; i++) {
    const pId = String(vals[i][1]).trim();
    const progId = String(vals[i][2]).trim();
    const st = String(vals[i][3]).trim();
    if (st === "Active" && activeProfs.has(pId)) {
      if (!out[progId]) out[progId] = new Set();
      out[progId].add(pId);
    }
  }
  return out;
}

function buildAvailability_(sheet, activeProfs) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const vals = sheet.getRange(2, 1, last - 1, 5).getValues();
  const out = {};
  for (let i = 0; i < vals.length; i++) {
    const pId = String(vals[i][1]).trim();
    const dayKey = normalizeDayKey_(vals[i][2]);
    const slot = normalizeTimeSlot_(vals[i][3]);
    const st = String(vals[i][4]).trim();

    if (st === "Active" && activeProfs.has(pId) && dayKey && slot) {
      if (!out[pId]) out[pId] = {};
      if (!out[pId][dayKey]) out[pId][dayKey] = new Set();
      out[pId][dayKey].add(slot);
    }
  }
  return out;
}

function buildTimeSlotOrder_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const vals = sheet.getRange(2, 4, last - 1, 1).getValues().flat();
  const uniq = [...new Set(vals.map(v => normalizeTimeSlot_(v)).filter(Boolean))];
  const toMin = (t) => {
    const p = t.split("-")[0].split(":");
    return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  };
  return uniq.sort((a, b) => toMin(a) - toMin(b));
}

function buildActiveEnrollmentUsage_(sheet, nextTimeSlotFn) {
  const last = sheet.getLastRow();
  const usage = {
    roomsBySlot: {},
    busyProfs: new Set()
  };
  if (last < 2) return usage;

  const vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (const r of vals) {
    if (String(r[10]).trim() !== "Active") continue; // Col K: State

    const dayKey = normalizeDayKey_(r[5]);           // Col F: DayOfWeek
    const slot = normalizeTimeSlot_(r[6]);           // Col G: TimeSlot
    const profId = String(r[8]).trim();              // Col I: AssignedProfessorID

    if (dayKey && slot) {
      const key = `${r[3]}|${dayKey}|${slot}`;
      usage[key] = (usage[key] || 0) + 1;

      // Conteo de salones físicos ocupados en este día y bloque
      const slotKey = `${dayKey}|${slot}`;
      usage.roomsBySlot[slotKey] = (usage.roomsBySlot[slotKey] || 0) + 1;

      // Registro de qué profesor está dando clase a esa hora
      if (profId) {
        usage.busyProfs.add(`${profId}|${slotKey}`);
      }
    }
  }
  return usage;
}

function getCandidateProfessorsWithCapacity_(opts) {
  const { programId, dayKey, timeSlot, professorPrograms, availability, usage } = opts;
  const profs = professorPrograms[programId];
  if (!profs) return [];

  const slotKey = `${dayKey}|${timeSlot}`;
  const candidates = [];

  for (const pId of profs) {
    const isAvail = availability[pId]?.[dayKey]?.has(timeSlot);
    const isBusy = usage.busyProfs?.has(`${pId}|${slotKey}`);

    // Si da la materia, tiene disponibilidad y NO tiene clase asignada a esa hora:
    if (isAvail && !isBusy) {
      candidates.push(pId);
    }
  }

  // candidates ya contiene únicamente a los profesores con capacidad real disponible
  return candidates;
}

function incUsage_(usage, programId, dayKey, timeSlot, profId) {
  const k = `${programId}|${dayKey}|${timeSlot}`;
  usage[k] = (usage[k] || 0) + 1;

  const slotKey = `${dayKey}|${timeSlot}`;
  if (!usage.roomsBySlot) usage.roomsBySlot = {};
  usage.roomsBySlot[slotKey] = (usage.roomsBySlot[slotKey] || 0) + 1;

  if (profId && usage.busyProfs) {
    usage.busyProfs.add(`${profId}|${slotKey}`);
  }
}

function getLastEnrollmentSequence_(sheet) {
  const currentYear = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxSeq = 0;

  ids.forEach(id => {
    if (String(id).startsWith(`AFM-E${currentYear}`)) {
      const seq = Number(String(id).slice(-3));
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });

  return maxSeq + 1;
}