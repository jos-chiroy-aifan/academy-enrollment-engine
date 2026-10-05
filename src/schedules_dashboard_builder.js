/****************************************************************
 * [ADM] SCHEDULES DASHBOARD (MOTOR EXACTO Y GRUPOS)
 * - Cruce de datos exacto mediante Normalización de Tiempos.
 * - Soporte para CLASES GRUPALES (Apila alumnos en la misma celda).
 * - Pinta Verde (Libre), Amarillo (Ocupado), Rojo (No trabaja).
 ****************************************************************/

// --- CONEXIONES ---
const DASHBOARD_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty("DASHBOARD_SPREADSHEET_ID"); // <-- TU VIEW

// Sheets en DB
const DB_SHEET_PROFESSORS   = "Professors";
const DB_SHEET_STUDENTS     = "Students";
const DB_SHEET_PROGRAMS     = "Programs";
const DB_SHEET_AVAILABILITY = "Professor_Availability";
const DB_SHEET_ENROLLMENTS  = "Student_Enrollments";

// Dashboard sheets
const DASH_GENERAL_SHEET = "GENERAL";
const DASH_PROF_PREFIX   = "Horario - ";

// Config horario
const START_HOUR  = 8;   // 08:00
const END_HOUR    = 18;  // 18:00 (fin)
const PERIOD_MIN  = 40;

// Días
const DAYS = [
  { key: "Monday",    label: "Lunes" },
  { key: "Tuesday",   label: "Martes" },
  { key: "Wednesday", label: "Miércoles" },
  { key: "Thursday",  label: "Jueves" },
  { key: "Friday",    label: "Viernes" },
  { key: "Saturday",  label: "Sábado" },
];

// Paleta de Colores
const C_RED_BG      = "#FCE8E6"; 
const C_RED_TXT     = "#EAA19B"; 
const C_GREEN_BG    = "#CEEAD6"; 
const C_GREEN_TXT   = "#137333";
const C_YELLOW_BG   = "#FFF2CC"; 
const C_YELLOW_TXT  = "#B05E52"; 
const C_HEADER_BG   = "#202124"; 
const C_HEADER_TXT  = "#FFFFFF"; 
const C_TITLE_BG    = "#F1F3F4"; 
const C_BORDER      = "#DADCE0"; 

/**
 * ✅ PUNTO DE ENTRADA PRINCIPAL
 */
function buildSchedulesDashboard() {
  // Ahora usamos getActiveSpreadsheet porque el script vive en la DB
  const db = SpreadsheetApp.getActiveSpreadsheet();
  const dash = SpreadsheetApp.openById(assertScriptProperty_("DASHBOARD_SPREADSHEET_ID", DASHBOARD_SPREADSHEET_ID));

  const profSheet   = db.getSheetByName(DB_SHEET_PROFESSORS);
  const studSheet   = db.getSheetByName(DB_SHEET_STUDENTS);
  const progSheet   = db.getSheetByName(DB_SHEET_PROGRAMS);
  const availSheet  = db.getSheetByName(DB_SHEET_AVAILABILITY);
  const enrollSheet = db.getSheetByName(DB_SHEET_ENROLLMENTS);

  if (!profSheet || !studSheet || !progSheet || !availSheet || !enrollSheet) {
    throw new Error("Faltan hojas en la Base de Datos.");
  }

  const timeSlots = buildFixedTimeSlots_(); 

  const professors   = readActiveProfessors_(profSheet); 
  const studentsById = readStudentsIndex_(studSheet);    
  const programsById = readProgramsIndex_(progSheet);    

  const profNameById = {};
  professors.forEach(p => (profNameById[p.id] = p.name));

  const availabilityByProf = readAvailability_(availSheet); 
  const enrollments        = readActiveEnrollments_(enrollSheet);  

  const classesByProf = buildClassesByProfessor_(enrollments, studentsById, programsById, timeSlots);
  const classesBySlot = buildClassesBySlotMixed_(enrollments, studentsById, programsById, profNameById, timeSlots);

  renderGeneralSheetMixed_(dash, timeSlots, classesBySlot);
  renderProfessorSheetsColored_(dash, professors, timeSlots, availabilityByProf, classesByProf);

  SpreadsheetApp.flush();
}

/* =========================================================
 * RENDER: PROFESORES INDIVIDUALES
 * ========================================================= */
function renderProfessorSheetsColored_(dash, professors, timeSlots, availabilityByProf, classesByProf) {
  cleanupOldProfessorSheets_(dash, professors);

  for (const prof of professors) {
    const name = `${DASH_PROF_PREFIX}${prof.name}`;
    const sh = getOrCreateSheet_(dash, name);
    sh.clear({ contentsOnly: true });
    sh.clearFormats();

    sh.getRange(1, 1).setBackground(C_TITLE_BG); 
    sh.getRange(1, 2, 1, DAYS.length).merge();
    sh.getRange(1, 2)
      .setValue(`HORARIO – ${prof.name.toUpperCase()}`)
      .setFontFamily("Montserrat").setFontWeight("bold").setFontSize(16)
      .setBackground(C_TITLE_BG)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.setRowHeight(1, 40);

    sh.getRange(2, 2, 1, DAYS.length).merge();
    sh.getRange(2, 2)
      .setValue("🟩 Libre (Disponible)   |   🟥 No Laboral (No disponible)   |   🟨 Ocupado (Grupal o Individual)")
      .setFontSize(10).setFontColor("#555")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    const header = ["Hora", ...DAYS.map(d => d.label)];
    sh.getRange(4, 1, 1, header.length)
      .setValues([header])
      .setBackground(C_HEADER_BG).setFontColor(C_HEADER_TXT)
      .setFontFamily("Montserrat").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.setRowHeight(4, 30);

    const rows = timeSlots.length;
    const cols = DAYS.length;
    const bgColors = [];
    const fontColors = [];
    const values = [];
    const fontWeights = [];
    const timeLabels = []; 

    const profAvail = availabilityByProf[prof.id] || new Set();
    const profClasses = classesByProf[prof.id] || {};

    for (let r = 0; r < rows; r++) {
      const slotObj = timeSlots[r];
      timeLabels.push([slotObj.display]); 

      const rowBg = [];
      const rowFc = [];
      const rowVal = [];
      const rowFw = [];

      for (let c = 0; c < cols; c++) {
        const dayKey = DAYS[c].key;
        const matrixKey = `${dayKey}|${slotObj.key}`; 

        if (profClasses[matrixKey]) {
          rowBg.push(C_YELLOW_BG);
          rowFc.push(C_YELLOW_TXT);
          rowVal.push(profClasses[matrixKey].text); // Aquí ya viene agrupado
          rowFw.push("bold");
        } 
        else if (profAvail.has(matrixKey)) {
          rowBg.push(C_GREEN_BG);
          rowFc.push(C_GREEN_TXT);
          rowVal.push("Libre");
          rowFw.push("normal");
        } 
        else {
          rowBg.push(C_RED_BG);
          rowFc.push(C_RED_TXT);
          rowVal.push("-");
          rowFw.push("normal");
        }
      }
      bgColors.push(rowBg);
      fontColors.push(rowFc);
      values.push(rowVal);
      fontWeights.push(rowFw);
    }

    const bodyStart = 5;
    
    sh.getRange(bodyStart, 1, rows, 1)
      .setValues(timeLabels)
      .setFontWeight("bold").setFontFamily("Montserrat")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    const dataRange = sh.getRange(bodyStart, 2, rows, cols);
    dataRange.setBackgrounds(bgColors);
    dataRange.setFontColors(fontColors);
    dataRange.setValues(values);
    dataRange.setFontWeights(fontWeights);
    
    sh.setFrozenRows(4);
    sh.setFrozenColumns(1); 
    sh.setColumnWidth(1, 110);
    for (let c = 2; c <= 1 + cols; c++) sh.setColumnWidth(c, 240);
    
    // Celdas más altas (90) para que quepan bien las clases grupales
    sh.setRowHeights(bodyStart, rows, 90); 
    
    dataRange.setVerticalAlignment("middle").setHorizontalAlignment("center").setWrap(true).setFontSize(10);

    sh.getRange(4, 1, 1 + rows, 1 + cols)
      .setBorder(true, true, true, true, true, true, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
}

/* =========================================================
 * RENDER: GENERAL (mezclado)
 * ========================================================= */
function renderGeneralSheetMixed_(dash, timeSlots, classesBySlot) {
  const sh = getOrCreateSheet_(dash, DASH_GENERAL_SHEET);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  sh.getRange(1, 1).setBackground(C_TITLE_BG);
  sh.getRange(1, 2, 1, DAYS.length).merge();
  sh.getRange(1, 2).setValue("HORARIO GENERAL (VISTA GLOBAL)")
    .setFontFamily("Montserrat").setFontWeight("bold").setFontSize(16).setBackground(C_TITLE_BG)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 40);

  const legend = ["Leyenda de uso:", "🟨 Ocupado", "Formato: [Profesor] · Alumno · Programa", ""];
  sh.getRange(2, 2, 1, 4).setValues([legend]); 
  sh.getRange(2, 2, 1, 4).setFontSize(10).setFontColor("#555");
  sh.getRange(2, 2).setFontWeight("bold");

  const header = ["Hora", ...DAYS.map(d => d.label)];
  sh.getRange(4, 1, 1, header.length).setValues([header])
    .setBackground(C_HEADER_BG).setFontColor(C_HEADER_TXT)
    .setFontFamily("Montserrat").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(4, 30);

  const rows = timeSlots.length;
  const cols = DAYS.length;
  const bg = [];
  const vals = [];
  const timeLabels = [];

  for (let r = 0; r < rows; r++) {
    const slotObj = timeSlots[r];
    timeLabels.push([slotObj.display]);
    const rowBg = [];
    const rowVal = [];

    for (let c = 0; c < cols; c++) {
      const key = `${DAYS[c].key}|${slotObj.key}`;
      const entries = classesBySlot[key] || [];
      if (entries.length) {
        rowBg.push(C_YELLOW_BG);
        rowVal.push(entries.join("\n\n"));
      } else {
        rowBg.push((r % 2 === 0) ? "#FFFFFF" : "#F8F9FA"); 
        rowVal.push("");
      }
    }
    bg.push(rowBg);
    vals.push(rowVal);
  }

  const bodyStart = 5;
  sh.getRange(bodyStart, 1, rows, 1).setValues(timeLabels).setFontWeight("bold").setFontFamily("Montserrat").setHorizontalAlignment("center").setVerticalAlignment("middle");
  const dataRange = sh.getRange(bodyStart, 2, rows, cols);
  dataRange.setBackgrounds(bg).setValues(vals).setVerticalAlignment("middle").setHorizontalAlignment("center").setWrap(true).setFontSize(9);
  
  sh.setFrozenRows(4);
  sh.setFrozenColumns(1); 
  sh.setColumnWidth(1, 110);
  for (let c = 2; c <= 1 + cols; c++) sh.setColumnWidth(c, 260);
  sh.setRowHeights(bodyStart, rows, 90);

  sh.getRange(4, 1, 1 + rows, 1 + cols).setBorder(true, true, true, true, true, true, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

/* =========================================================
 * PREPARADORES DE CLASES (AGRUPACIÓN DE ALUMNOS)
 * ========================================================= */
function buildClassesByProfessor_(enrollments, studentsById, programsById, timeSlots) {
  const slotIndex = {};
  timeSlots.forEach((s, i) => (slotIndex[s.key] = i)); 
  const out = {}; 

  for (const e of enrollments) {
    const studentName = studentsById[e.studentId]?.name || e.studentId;
    const programName = programsById[e.programId] || e.programId;
    
    const slotsToFill = [e.timeSlotKey]; 
    
    if (e.periods === 2) {
      const idx = slotIndex[e.timeSlotKey];
      if (idx !== undefined && idx + 1 < timeSlots.length) {
        slotsToFill.push(timeSlots[idx + 1].key);
      }
    }

    // AQUI ESTÁ LA MAGIA DE LA AGRUPACIÓN
    for (const slotKey of slotsToFill) {
      const matrixKey = `${e.dayKey}|${slotKey}`;
      
      if (!out[e.profId]) out[e.profId] = {};
      
      // Si no existe arreglo para este horario, lo creamos
      if (!out[e.profId][matrixKey]) {
        out[e.profId][matrixKey] = { list: [] };
      }

      // Agregamos al alumno a la lista de ese horario
      out[e.profId][matrixKey].list.push(`👤 ${studentName}\n🎵 ${programName}`);
    }
  }

  // Convertimos las listas en un solo bloque de texto separado por guiones
  for (const profId in out) {
    for (const matrixKey in out[profId]) {
      out[profId][matrixKey].text = out[profId][matrixKey].list.join("\n------\n");
    }
  }

  return out;
}

function buildClassesBySlotMixed_(enrollments, studentsById, programsById, profNameById, timeSlots) {
  const slotIndex = {};
  timeSlots.forEach((s, i) => (slotIndex[s.key] = i));
  const out = {}; 

  for (const e of enrollments) {
    const studentName = studentsById[e.studentId]?.name || e.studentId;
    const programName = programsById[e.programId] || e.programId;
    const profName = profNameById[e.profId] || e.profId;
    
    const slotsToFill = [e.timeSlotKey];
    if (e.periods === 2) {
      const idx = slotIndex[e.timeSlotKey];
      if (idx !== undefined && idx + 1 < timeSlots.length) slotsToFill.push(timeSlots[idx + 1].key);
    }

    for (const slotKey of slotsToFill) {
      const matrixKey = `${e.dayKey}|${slotKey}`;
      if (!out[matrixKey]) out[matrixKey] = [];
      out[matrixKey].push(`[${profName}]\n👤 ${studentName} - ${programName}`);
    }
  }
  for (const k of Object.keys(out)) {
    out[k].sort((a, b) => a.localeCompare(b));
  }
  return out;
}

/* =========================================================
 * READERS DE LA DB
 * ========================================================= */
function getColIndex_(sheet, colName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(colName);
}

function readActiveProfessors_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const cId = getColIndex_(sheet, "ProfessorID");
  const cFn = getColIndex_(sheet, "First Name");
  const cLn = getColIndex_(sheet, "Last Name");
  const cState = getColIndex_(sheet, "State");

  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const out = [];

  for (const r of values) {
    const id = String(r[cId] || "").trim();
    if (!id || String(r[cState] || "").trim() !== "Active") continue;
    out.push({ id, name: `${String(r[cFn] || "").trim()} ${String(r[cLn] || "").trim()}`.trim() || id });
  }
  return out;
}

function readAvailability_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const cProf = getColIndex_(sheet, "ProfessorID");
  const cDay = getColIndex_(sheet, "DayOfWeek");
  const cSlot = getColIndex_(sheet, "TimeSlot");
  const cState = getColIndex_(sheet, "State");

  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const out = {}; 
  
  for (const r of values) {
    const profId = String(r[cProf] || "").trim();
    const dayKey = normalizeDayKey_(r[cDay]);
    const rawSlot = String(r[cSlot] || "").trim();
    const state = String(r[cState] || "").trim();
    
    if (!profId || !dayKey || !rawSlot || state !== "Active") continue;
    
    if (!out[profId]) out[profId] = new Set();
    
    out[profId].add(`${dayKey}|${normalizeTimeSlot_(rawSlot)}`);
  }
  return out;
}

function readActiveEnrollments_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const col = (n) => getColIndex_(sheet, n);
  
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const out = [];

  for (const r of values) {
    if (String(r[col("State")] || "").trim() !== "Active") continue;
    const dayKey = normalizeDayKey_(r[col("DayOfWeek")]);
    const rawSlot = String(r[col("TimeSlot")] || "").trim();
    
    if (dayKey && rawSlot) {
      out.push({
        studentId: String(r[col("StudentID")] || "").trim(),
        programId: String(r[col("ProgramID")] || "").trim(),
        profId: String(r[col("AssignedProfessorID")] || "").trim(),
        dayKey: dayKey,
        timeSlotKey: normalizeTimeSlot_(rawSlot), 
        periods: Number(r[col("Periods")] || 1)
      });
    }
  }
  return out;
}

function readStudentsIndex_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const cId = getColIndex_(sheet, "StudentID");
  const cFn = getColIndex_(sheet, "First Name");
  const cLn = getColIndex_(sheet, "Last Name");

  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const out = {};
  for (const r of values) {
    const id = String(r[cId] || "").trim();
    if (id) out[id] = { name: `${String(r[cFn] || "").trim()} ${String(r[cLn] || "").trim()}`.trim() || id };
  }
  return out;
}

function readProgramsIndex_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const values = sheet.getRange(2, 1, last - 1, 2).getValues();
  const out = {};
  for (const r of values) {
    if (r[0] && r[1]) out[String(r[0]).trim()] = String(r[1]).trim();
  }
  return out;
}

/* =========================================================
 * HELPERS MENORES
 * ========================================================= */
function buildFixedTimeSlots_() {
  const slots = [];
  let totalMin = START_HOUR * 60;
  const endMin = END_HOUR * 60;
  while (totalMin + PERIOD_MIN <= endMin) {
    const start = `${String(Math.floor(totalMin / 60)).padStart(2, "0")}:${String(totalMin % 60).padStart(2, "0")}`;
    const end = `${String(Math.floor((totalMin + PERIOD_MIN) / 60)).padStart(2, "0")}:${String((totalMin + PERIOD_MIN) % 60).padStart(2, "0")}`;
    
    slots.push({
      display: `${start} - ${end}`, 
      key: `${start}-${end}`        
    });
    totalMin += PERIOD_MIN;
  }
  return slots;
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function cleanupOldProfessorSheets_(dash, professors) {
  const allowed = new Set(professors.map(p => `${DASH_PROF_PREFIX}${p.name}`));
  for (const sh of dash.getSheets()) {
    const n = sh.getName();
    if (n.startsWith(DASH_PROF_PREFIX) && !allowed.has(n)) dash.deleteSheet(sh);
  }
}