function buildActiveProgramMap_(progSheet) {
  const last = progSheet.getLastRow();
  if (last < 2) return {};
  const data = progSheet.getRange(2, 1, last - 1, progSheet.getLastColumn()).getValues();
  const map = {};
  for (const r of data) {
    const programId = r[0];
    const programName = r[1];
    const state = r[4];
    if (state === "Active" && programId && programName) {
      map[String(programName).trim()] = String(programId).trim();
    }
  }
  return map;
}

function splitMulti_(value) {
  if (!value) return [];
  return String(value).split(",").map(s => s.trim()).filter(Boolean);
}

function getExistingPairs_(sheet, colA, colB) {
  const last = sheet.getLastRow();
  const set = new Set();
  if (last < 2) return set;
  const data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (const r of data) {
    const a = r[colA - 1];
    const b = r[colB - 1];
    if (a && b) set.add(`${a}__${b}`);
  }
  return set;
}

function getExistingTriples_(sheet, colA, colB, colC) {
  const last = sheet.getLastRow();
  const set = new Set();
  if (last < 2) return set;
  const data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (const r of data) {
    const a = r[colA - 1];
    const b = r[colB - 1];
    const c = r[colC - 1];
    if (a && b && c) set.add(`${a}__${b}__${c}`);
  }
  return set;
}

// ✅ VALIDA QUE UNA SCRIPT PROPERTY OBLIGATORIA ESTÉ CONFIGURADA
function assertScriptProperty_(key, value) {
  if (!value) {
    throw new Error(`Falta configurar la Script Property "${key}" (Configuración del proyecto > Propiedades del script).`);
  }
  return value;
}

function generateLinkId_(prefix) {
  const year = new Date().getFullYear();
  const rand = Math.random().toString(36).substring(2, 7).toUpperCase();
  return `${prefix}-${year}${rand}`;
}

function dayToSpanish_(day) {
  const map = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles", 
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo",
    "MONDAY": "Lunes", "TUESDAY": "Martes", "WEDNESDAY": "Miércoles",
    "THURSDAY": "Jueves", "FRIDAY": "Viernes", "SATURDAY": "Sábado", "SUNDAY": "Domingo",
    "Lunes": "Lunes", "Martes": "Martes", "Miércoles": "Miércoles",
    "Jueves": "Jueves", "Viernes": "Viernes", "Sábado": "Sábado", "Domingo": "Domingo"
  };
  return map[String(day).trim()] || String(day).trim(); 
}

// ✅ NORMALIZADOR DE DÍAS (Ahora con Domingo)
function normalizeDayKey_(day) {
  if (!day) return null;
  const d = String(day).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (d.includes("lun") || d.includes("mon")) return "Monday";
  if (d.includes("mar") || d.includes("tue")) return "Tuesday";
  if (d.includes("mie") || d.includes("wed")) return "Wednesday";
  if (d.includes("jue") || d.includes("thu")) return "Thursday";
  if (d.includes("vie") || d.includes("fri")) return "Friday";
  if (d.includes("sab") || d.includes("sat")) return "Saturday";
  if (d.includes("dom") || d.includes("sun")) return "Sunday";
  return null;
}

// ✅ EL ÚNICO NORMALIZADOR DE HORAS (Adiós a los errores de espacios)
function normalizeTimeSlot_(str) {
  if (!str) return "";
  try {
    const clean = String(str).replace(/\s+/g, '').replace('a', '-'); 
    const parts = clean.split('-');
    if (parts.length !== 2) return clean;
    const fmt = (t) => {
      let [h, m] = t.split(':');
      return `${String(h).padStart(2, '0')}:${String(m || "00").padStart(2, '0')}`;
    };
    return `${fmt(parts[0])}-${fmt(parts[1])}`; 
  } catch (e) {
    return String(str).trim();
  }
}

function dumpDatabaseSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const result = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    if (lastCol === 0) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const sample = lastRow >= 2 ? sheet.getRange(2, 1, 1, lastCol).getValues()[0] : [];

    const cols = headers.map((h, i) => {
      // Letra de columna aproximada (A, B, C...)
      const letter = i < 26 ? String.fromCharCode(65 + i) : `A${String.fromCharCode(65 + (i - 26))}`;
      const val = sample[i] !== undefined && sample[i] !== "" ? ` -> ej: "${String(sample[i]).slice(0, 25)}"` : "";
      return `  [${letter} | idx ${i}] "${h}"${val}`;
    });

    result.push(`📋 HOJA: "${name}" (${lastRow} filas)\n` + cols.join("\n"));
  });

  Logger.log("\n\n" + result.join("\n\n"));
}