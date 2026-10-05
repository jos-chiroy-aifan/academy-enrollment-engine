/************************************
 * schedules_form_sync.gs
 * - Actualiza el Form (dropdowns) con:
 * 1) Programas con disponibilidad real
 * 2) Horarios disponibles por programa
 * 3) Descuenta horarios ya ocupados (Student_Enrollments)
 * 4) Limita la oferta a máximo 6 salones físicos simultáneos
 * 5) Mantiene la navegación por secciones activa
 ************************************/

const SCHEDULES_FORM_ID = PropertiesService.getScriptProperties().getProperty("UNIFIED_FORM_ID");

// Nombres de las hojas de cálculo (Database)
const SHEET_PROGRAMS = "Programs";
const SHEET_PROFESSORS = "Professors";
const SHEET_PROFESSOR_PROGRAMS = "Professor_Programs";
const SHEET_PROFESSOR_AVAILABILITY = "Professor_Availability";
const SHEET_STUDENT_ENROLLMENTS = "Student_Enrollments";

// Capacidad máxima de aulas físicas
const MAX_PHYSICAL_ROOMS = Number(PropertiesService.getScriptProperties().getProperty("MAX_PHYSICAL_ROOMS")) || 6;

// Título de la pregunta principal de programas
const FORM_Q_PROGRAM_TITLE = "Programa";

// Sección de respaldo (Fallback) en caso de error de ruta
const FORM_FALLBACK_SECTION_TITLE = "Autorizaciones"; 

/**
 * FUNCIÓN PRINCIPAL DE SINCRONIZACIÓN
 */
function schedulesFormSync() {
  try {
    sync_programs_to_professors_form();
    // Pausa preventiva para liberar el bloqueo de Drive tras guardar cambios
    Utilities.sleep(1500);
  } catch (e) {
    Logger.log("Aviso: No se pudo ejecutar sync_programs_to_professors_form. " + e);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const programsSheet = ss.getSheetByName(SHEET_PROGRAMS);
  const professorsSheet = ss.getSheetByName(SHEET_PROFESSORS);
  const professorProgramsSheet = ss.getSheetByName(SHEET_PROFESSOR_PROGRAMS);
  const availabilitySheet = ss.getSheetByName(SHEET_PROFESSOR_AVAILABILITY);
  const enrollmentsSheet = ss.getSheetByName(SHEET_STUDENT_ENROLLMENTS);

  if (!programsSheet || !professorsSheet || !professorProgramsSheet || !availabilitySheet) {
    throw new Error("Falta alguna sheet requerida (Programs / Professors / Professor_Programs / Professor_Availability).");
  }

  // 1. OBTENER DATOS ACTIVOS Y DISPONIBILIDAD
  const programs = getActivePrograms_(programsSheet);
  const activeProfessorIds = getActiveProfessorIds_(professorsSheet);
  const profPrograms = getProfessorProgramsActive_(professorProgramsSheet, activeProfessorIds);
  const availability = getActiveAvailability_(availabilitySheet, activeProfessorIds);

  const timeSlotOrder = buildTimeSlotOrder_(availabilitySheet);
  const nextTimeSlot = (slot) => {
    const i = timeSlotOrder.indexOf(slot);
    return (i >= 0 && i < timeSlotOrder.length - 1) ? timeSlotOrder[i + 1] : null;
  };

  const { busyByProfessor, roomUsage } = buildEnrollmentUsageIndexes_(enrollmentsSheet, nextTimeSlot);

  // 2. CALCULAR CUPOS REALES CONSIDERANDO TOPE DE SALONES
  const availableSchedulesByProgramName = buildAvailableSchedulesByProgramName_({
    programs,
    profPrograms,
    availability,
    busyByProfessor,
    roomUsage
  });

  const programsWithAvailability = programs
    .map(p => p.name)
    .filter(name => (availableSchedulesByProgramName[name] || []).length > 0);

  // 3. ACTUALIZAR GOOGLE FORM (Con reintento seguro para evitar timeouts)
  const form = openFormWithRetry_(assertScriptProperty_("UNIFIED_FORM_ID", SCHEDULES_FORM_ID));

  const dynamicSectionMap = {};
  programs.forEach(p => {
    dynamicSectionMap[p.name] = `Horario del Programa - ${p.name}`;
  });

  setDropdownChoicesByTitle_(form, FORM_Q_PROGRAM_TITLE, programsWithAvailability.length ? programsWithAvailability : ["(Sin programas disponibles)"]);
  setProgramRouting_(form, FORM_Q_PROGRAM_TITLE, programsWithAvailability, dynamicSectionMap, FORM_FALLBACK_SECTION_TITLE);

  // 4. ACTUALIZAR PREGUNTAS DE HORARIOS DINÁMICAMENTE
  for (const programName of Object.keys(availableSchedulesByProgramName)) {
    const qTitle = `Horarios - ${programName}`;
    const choices = availableSchedulesByProgramName[programName] || [];
    setDropdownChoicesByTitle_(form, qTitle, choices.length ? choices : ["(Sin cupos disponibles por ahora)"]);
  }
  
  Logger.log("✅ Sincronización finalizada con éxito.");
}

/************************************
 * FUNCIONES AUXILIARES: LECTURA DE DATOS
 ************************************/

function getActivePrograms_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();

  const idxId = findHeaderIndex_(headers, ["ProgramID", "ID"]);
  const idxName = findHeaderIndex_(headers, ["Program Name", "Nombre", "Programa"]);
  const idxState = findHeaderIndex_(headers, ["State", "Status", "Estado"]);

  const out = [];
  for (const r of values) {
    const pId = String(r[idxId >= 0 ? idxId : 0] || "").trim();
    const name = String(r[idxName >= 0 ? idxName : 1] || "").trim();
    const state = String(r[idxState >= 0 ? idxState : 4] || "").trim();
    if (state.toLowerCase() === "active" && pId && name) {
      out.push({ programId: pId, name: name });
    }
  }
  return out;
}

function getActiveProfessorIds_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return new Set();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();

  const idxId = findHeaderIndex_(headers, ["ProfessorID", "ID"]);
  const idxState = findHeaderIndex_(headers, ["State", "Status", "Estado"]);

  const set = new Set();
  for (const r of values) {
    const id = String(r[idxId >= 0 ? idxId : 0] || "").trim();
    const st = String(r[idxState >= 0 ? idxState : 10] || "").trim();
    if (id && st.toLowerCase() === "active") set.add(id);
  }
  return set;
}

function getProfessorProgramsActive_(sheet, activeProfessorIds) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const map = {}; 
  for (const r of values) {
    const profId = String(r[1]).trim();
    const progId = String(r[2]).trim();
    const state = String(r[3]).trim();
    if (state.toLowerCase() === "active" && activeProfessorIds.has(profId)) {
      if (!map[progId]) map[progId] = new Set();
      map[progId].add(profId);
    }
  }
  return map;
}

function getActiveAvailability_(sheet, activeProfessorIds) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const out = {}; 
  for (const r of values) {
    const profId = String(r[1]).trim();
    const dayKey = normalizeDayKey_(r[2]);
    const slot = normalizeTimeSlot_(r[3]);
    const state = String(r[4]).trim();
    if (state.toLowerCase() === "active" && activeProfessorIds.has(profId) && dayKey && slot) {
      if (!out[profId]) out[profId] = {};
      if (!out[profId][dayKey]) out[profId][dayKey] = new Set();
      out[profId][dayKey].add(slot);
    }
  }
  return out;
}

function buildTimeSlotOrder_(availabilitySheet) {
  const last = availabilitySheet.getLastRow();
  if (last < 2) return [];
  const values = availabilitySheet.getRange(2, 4, last - 1, 1).getValues().flat();
  const uniq = Array.from(new Set(values.map(v => String(v || "").trim()).filter(Boolean)));
  uniq.sort((a, b) => {
    const sa = a.split("-")[0].trim();
    const sb = b.split("-")[0].trim();
    return timeToMinutes_(sa) - timeToMinutes_(sb);
  });
  return uniq;
}

/************************************
 * FUNCIONES AUXILIARES: OCUPACIÓN
 ************************************/

function buildEnrollmentUsageIndexes_(enrollmentsSheet, nextTimeSlotFn) {
  const busyByProfessor = {}; 
  const roomUsage = {};

  if (!enrollmentsSheet) return { busyByProfessor, roomUsage };
  const last = enrollmentsSheet.getLastRow();
  if (last < 2) return { busyByProfessor, roomUsage };

  const headers = enrollmentsSheet.getRange(1, 1, 1, enrollmentsSheet.getLastColumn()).getValues()[0];

  // Búsqueda flexible de columnas con fallbacks a tus columnas A-K
  const cProf = findHeaderIndex_(headers, ["AssignedProfessorID", "ProfessorID", "AssignedProfessor", "Profesor"], 8);
  const cDay = findHeaderIndex_(headers, ["DayOfWeek", "Day", "Dia"], 5);
  const cSlot = findHeaderIndex_(headers, ["TimeSlot", "Slot", "Horario"], 6);
  const cState = findHeaderIndex_(headers, ["State", "Status", "Estado"], 10);

  const values = enrollmentsSheet.getRange(2, 1, last - 1, enrollmentsSheet.getLastColumn()).getValues();
  for (const r of values) {
    const stateVal = String(r[cState] || "").trim().toLowerCase();
    if (stateVal !== "active") continue; // Descarta alumnos inactivos

    const pId = String(r[cProf] || "").trim();
    const dKey = normalizeDayKey_(r[cDay]);
    const slot = normalizeTimeSlot_(r[cSlot]);

    if (dKey && slot) {
      const roomKey = `${dKey}|${slot}`;
      roomUsage[roomKey] = (roomUsage[roomKey] || 0) + 1;

      if (pId) {
        if (!busyByProfessor[pId]) busyByProfessor[pId] = {};
        if (!busyByProfessor[pId][dKey]) busyByProfessor[pId][dKey] = new Set();
        busyByProfessor[pId][dKey].add(slot);
      }
    }
  }
  return { busyByProfessor, roomUsage };
}

function buildAvailableSchedulesByProgramName_(ctx) {
  const { programs, profPrograms, availability, busyByProfessor, roomUsage } = ctx;
  const out = {};
  programs.forEach(p => (out[p.name] = []));

  const loggedFullRooms = new Set();

  for (const p of programs) {
    const profs = profPrograms[p.programId];
    if (!profs) continue;
    const comboHasCapacity = new Set();

    for (const profId of profs) {
      const profAvail = availability[profId];
      if (!profAvail) continue;

      for (const dayKey of Object.keys(profAvail)) {
        for (const slot of profAvail[dayKey]) {
          const roomKey = `${dayKey}|${slot}`;
          const currentOccupancy = roomUsage[roomKey] || 0;

          // Si ya hay MAX_PHYSICAL_ROOMS o más salones ocupados en la sede
          if (currentOccupancy >= MAX_PHYSICAL_ROOMS) {
            if (!loggedFullRooms.has(roomKey)) {
              Logger.log(`🚨 [CAPACIDAD LLENA] Horario bloqueado por salones agotados -> Día: ${dayKey}, Bloque: ${slot}, Ocupados: ${currentOccupancy}/${MAX_PHYSICAL_ROOMS}`);
              loggedFullRooms.add(roomKey);
            }
            continue;
          }

          // Si el profesor ya tiene alumno a esa hora, no puede tomar otro
          const busySet = busyByProfessor[profId]?.[dayKey];
          if (busySet && busySet.has(slot)) {
            continue;
          }

          comboHasCapacity.add(roomKey);
        }
      }
    }

    out[p.name] = Array.from(comboHasCapacity).map(k => {
      const [d, s] = k.split("|");
      return formatScheduleLabelEs_(d, s);
    }).sort(scheduleSort_);
  }
  return out;
}

/************************************
 * FUNCIONES AUXILIARES: FORMULARIO
 ************************************/

function setDropdownChoicesByTitle_(form, title, choices) {
  const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  const targetNorm = norm(title);

  const items = form.getItems();
  for (const it of items) {
    if (norm(it.getTitle()) === targetNorm) {
      const type = it.getType();
      if (type === FormApp.ItemType.LIST) {
        it.asListItem().setChoiceValues(choices);
        return;
      } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        it.asMultipleChoiceItem().setChoiceValues(choices);
        return;
      }
    }
  }
  Logger.log(`⚠️ No se encontró la pregunta "${title}" en el formulario.`);
}

function setProgramRouting_(form, programQuestionTitle, validPrograms, sectionTitleByProgram, fallbackSectionTitle) {
  const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  const targetNorm = norm(programQuestionTitle);

  const items = form.getItems();
  let programItem = null;
  for (const it of items) {
    if (norm(it.getTitle()) === targetNorm) {
      programItem = it;
      break;
    }
  }
  if (!programItem) return;

  const pageByTitle = {};
  form.getItems(FormApp.ItemType.PAGE_BREAK).forEach(it => {
    pageByTitle[norm(it.asPageBreakItem().getTitle())] = it.asPageBreakItem();
  });

  const fallbackPage = pageByTitle[norm(fallbackSectionTitle)];
  const choices = (validPrograms || []).map(pName => {
    const sectionTitle = sectionTitleByProgram[pName];
    const targetPage = sectionTitle ? pageByTitle[norm(sectionTitle)] : null;
    const page = targetPage || fallbackPage;
    
    if (programItem.getType() === FormApp.ItemType.LIST) {
      const list = programItem.asListItem();
      return page ? list.createChoice(pName, page) : list.createChoice(pName);
    } else {
      const mc = programItem.asMultipleChoiceItem();
      return page ? mc.createChoice(pName, page) : mc.createChoice(pName);
    }
  });

  if (choices.length > 0) {
    if (programItem.getType() === FormApp.ItemType.LIST) {
      programItem.asListItem().setChoices(choices);
    } else if (programItem.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
      programItem.asMultipleChoiceItem().setChoices(choices);
    }
  }
}

function openFormWithRetry_(formId, maxAttempts = 3) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return FormApp.openById(formId);
    } catch (err) {
      if (attempt === maxAttempts) throw err;
      Utilities.sleep(1500 * attempt);
    }
  }
}

function findHeaderIndex_(headers, candidateNames, fallback = -1) {
  const norm = (s) => String(s || "").toLowerCase().replace(/[\s_\-]/g, '').normalize("NFD").replace(/[\u0300-\u036f]/g, '');
  const cleanCandidates = candidateNames.map(norm);

  for (let i = 0; i < headers.length; i++) {
    const cleanHeader = norm(headers[i]);
    if (cleanCandidates.includes(cleanHeader)) return i;
  }
  return fallback;
}

/************************************
 * FUNCIONES AUXILIARES: TIEMPO Y FORMATO
 ************************************/

function formatScheduleLabelEs_(dayKey, timeSlot) {
  const map = { 
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles", 
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo" 
  };
  
  let [start, end] = String(timeSlot).split('-');
  
  if (start && start.startsWith("0")) start = start.substring(1);
  if (end && end.startsWith("0")) end = end.substring(1);
  
  const finalTime = (start && end) ? `${start} a ${end}` : String(timeSlot).replace("-", " a ");

  return `${map[dayKey] || dayKey} - ${finalTime}`;
}

function scheduleSort_(a, b) {
  const dayOrder = { "Lunes": 1, "Martes": 2, "Miércoles": 3, "Jueves": 4, "Viernes": 5, "Sábado": 6, "Domingo": 7 };
  const oa = dayOrder[a.split(" - ")[0]] || 99;
  const ob = dayOrder[b.split(" - ")[0]] || 99;
  if (oa !== ob) return oa - ob;
  return timeToMinutes_(a.split(" - ")[1]) - timeToMinutes_(b.split(" - ")[1]);
}

function timeToMinutes_(t) {
  const clean = String(t).split("a")[0].trim();
  const parts = clean.split(":").map(Number);
  return (parts[0] * 60) + (parts[1] || 0);
}