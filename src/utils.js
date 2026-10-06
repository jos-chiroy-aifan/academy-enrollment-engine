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

// ✅ ¿La fila tiene datos reales en alguna de estas columnas (base 1)?
// Sirve para distinguir una fila "plantilla" (solo fórmulas/casillas) de una con información.
function rowHasData_(sheet, row, cols) {
  return cols.some(c => {
    const v = sheet.getRange(row, c).getValue();
    return v !== "" && v !== null;
  });
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

/****************************************************************
 * LOGGER DE EJECUCIÓN
 * - Cada línea lleva el nombre del proceso y un ID corto de ejecución,
 *   así se distinguen dos ejecuciones simultáneas en el panel "Ejecuciones".
 * - info / warn / error usan console.* para que el nivel (ícono) se vea en el panel.
 * - result() cierra la ejecución con una línea "RESULTADO" fácil de buscar.
 ****************************************************************/
function createRunLogger_(scope) {
  const runId = Utilities.getUuid().slice(0, 8);
  const startedAt = Date.now();
  const ctx = {};
  let warnings = 0;
  let errors = 0;

  const prefix = () => {
    const extra = Object.keys(ctx).map(k => `${k}=${ctx[k]}`).join(" ");
    return `[${scope} #${runId}${extra ? " " + extra : ""}]`;
  };

  return {
    runId,
    set(key, value) { ctx[key] = value; },
    info(msg) { console.info(`${prefix()} ${msg}`); },
    warn(msg) { warnings++; console.warn(`${prefix()} ⚠️ ${msg}`); },
    error(msg, err) {
      errors++;
      const detail = err ? `\n${err.stack || err}` : "";
      console.error(`${prefix()} ❌ ${msg}${detail}`);
    },
    // status: "OK" | "OMITIDO" | "FALLO"
    result(status, detail) {
      const ms = Date.now() - startedAt;
      const line = `${prefix()} 🏁 RESULTADO: ${status} | ${ms} ms | advertencias=${warnings} errores=${errors}${detail ? " | " + detail : ""}`;
      if (status === "FALLO") console.error(line);
      else if (status === "OMITIDO" || warnings > 0) console.warn(line);
      else console.info(line);
    }
  };
}

/****************************************************************
 * RADIOGRAFÍA DE LA BASE DE DATOS (para documentar la estructura)
 * Por hoja: filas de datos, columnas, filas congeladas.
 * Por columna: letra/índice, encabezado exacto, tipo de dato, llenado,
 * fórmulas, listas desplegables y, si NO es dato personal, los valores
 * posibles o el "formato" de los textos (ej. "9:99 - 9:99" en horarios).
 * Las columnas con datos personales solo muestran tipo y llenado:
 * nunca se imprimen nombres, correos, teléfonos ni fechas de nacimiento.
 ****************************************************************/
function dumpDatabaseSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_DISTINCT = 12; // Con más valores distintos se muestran formatos en vez de valores
  const MAX_LEN = 40;

  const PII_HEADER = /(nombre|name|apellido|correo|e-?mail|tel[eé]fono|phone|nacimiento|birth|responsable|responsible|parentesco|relationship|trayectoria|notes|notas|^student$|^professor$|assigned (student|professor)$)/i;
  const isPersonal = (header) => PII_HEADER.test(header) && !/program/i.test(header);

  const colLetter = (i) => {
    let n = i + 1, s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };
  const cut = (v) => { const t = String(v); return t.length > MAX_LEN ? t.slice(0, MAX_LEN) + "…" : t; };
  const typeOf = (v) => {
    if (v instanceof Date) return "Fecha";
    if (typeof v === "number") return "Número";
    if (typeof v === "boolean") return "Booleano";
    return "Texto";
  };
  // Formato de un texto: dígitos -> 9, palabras -> a (conserva espacios y símbolos)
  const shapeOf = (v) => cut(String(v).replace(/[0-9]/g, "9").replace(/[A-Za-zÀ-ÿ]+/g, "a"));
  const fmtDate = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  const topList = (map, limit) => [...map.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([k, n]) => `"${k}" (${n})`)
    .join(", ");

  Logger.log(`🗂️ BASE DE DATOS: "${ss.getName()}" | zona horaria: ${Session.getScriptTimeZone()} | hojas: ${ss.getSheets().length}`);

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const dataRows = Math.max(lastRow - 1, 0);

    const lines = [`📋 HOJA "${name}" | filas de datos: ${dataRows} | columnas: ${lastCol} | filas congeladas: ${sheet.getFrozenRows()}${sheet.isSheetHidden() ? " | OCULTA" : ""}`];

    if (lastCol === 0) {
      lines.push("  (hoja vacía)");
      Logger.log(lines.join("\n"));
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const values = dataRows > 0 ? sheet.getRange(2, 1, dataRows, lastCol).getValues() : [];
    const formulas = dataRows > 0 ? sheet.getRange(2, 1, dataRows, lastCol).getFormulas() : [];
    const validations = dataRows > 0 ? sheet.getRange(2, 1, 1, lastCol).getDataValidations()[0] : [];

    headers.forEach((h, c) => {
      const header = String(h);
      const personal = isPersonal(header.trim());
      const parts = [`  [${colLetter(c)} | idx ${c}] "${header}"`];

      if (header === "") parts.push("⚠ sin encabezado");
      else if (header !== header.trim()) parts.push("⚠ encabezado con espacios al inicio/final");

      const types = new Map();
      const distinct = new Map();
      const shapes = new Map();
      let filled = 0, nMin = null, nMax = null, dMin = null, dMax = null, formulaCount = 0, firstFormula = "";

      for (let r = 0; r < values.length; r++) {
        if (formulas[r][c]) {
          formulaCount++;
          if (!firstFormula) firstFormula = formulas[r][c];
        }
        const v = values[r][c];
        if (v === "" || v === null) continue;
        filled++;
        const t = typeOf(v);
        types.set(t, (types.get(t) || 0) + 1);
        if (t === "Número") { nMin = nMin === null ? v : Math.min(nMin, v); nMax = nMax === null ? v : Math.max(nMax, v); }
        if (t === "Fecha") { dMin = !dMin || v < dMin ? v : dMin; dMax = !dMax || v > dMax ? v : dMax; }
        if (!personal) {
          const key = t === "Fecha" ? fmtDate(v) : cut(v);
          distinct.set(key, (distinct.get(key) || 0) + 1);
          if (t === "Texto") shapes.set(shapeOf(v), (shapes.get(shapeOf(v)) || 0) + 1);
        }
      }

      parts.push(types.size === 0 ? "vacía"
        : types.size === 1 ? [...types.keys()][0]
        : "MIXTO: " + [...types.entries()].map(([t, n]) => `${t} ${n}`).join(", "));
      parts.push(`lleno ${filled}/${dataRows}`);

      if (formulaCount > 0) parts.push(`fórmula en ${formulaCount} filas (ej. ${cut(firstFormula)})`);

      const rule = validations[c];
      if (rule) {
        const type = rule.getCriteriaType();
        if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          parts.push(`lista desplegable: ${rule.getCriteriaValues()[0].map(cut).join(" / ")}`);
        } else if (type === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
          parts.push("casilla de verificación");
        } else {
          parts.push(`validación: ${type}`);
        }
      }

      if (personal) {
        parts.push("🔒 dato personal (sin muestra)");
      } else if (filled > 0) {
        if (distinct.size <= MAX_DISTINCT) {
          parts.push(`valores: ${topList(distinct, MAX_DISTINCT)}`);
        } else {
          parts.push(`${distinct.size} valores distintos`);
          if (nMin !== null) parts.push(`rango ${nMin} → ${nMax}`);
          if (dMin) parts.push(`rango ${fmtDate(dMin)} → ${fmtDate(dMax)}`);
          if (shapes.size > 0) parts.push(`formatos: ${topList(shapes, 4)}`);
        }
      }

      lines.push(parts.join(" | "));
    });

    // Una entrada de log por hoja para que ninguna se corte
    Logger.log(lines.join("\n"));
  });
}