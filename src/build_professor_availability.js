function upsertProfessorAvailabilityFromResponse_(paSheet, professorId, byDay) {
  // Estructura Professor_Availability:
  // A: ProfessorAvailabilityID | B: ProfessorID | C: DayOfWeek | D: TimeSlot
  // E: State | F: Source | G: LastUpdatedAt | H: Professor Name | I: Programs

  const now = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Helpers para obtener nombre y programas del profesor
  const profName = getProfessorNameHelper_(ss, professorId);
  const profPrograms = getProfessorProgramsHelper_(ss, professorId);

  // 1. Obtener todos los slots solicitados en el form
  const requestedSlots = new Set();
  for (const [dayOfWeek, rawSlots] of byDay) {
    const slots = splitMulti_(rawSlots);
    for (const slot of slots) {
      requestedSlots.add(`${dayOfWeek}__${slot}`);
    }
  }

  // 2. Leer lo que ya existe en la DB
  const lastRow = paSheet.getLastRow();

  if (lastRow < 2) {
    for (const reqKey of requestedSlots) {
      const [d, s] = reqKey.split("__");
      paSheet.appendRow([generateLinkId_("AFM-PA"), professorId, d, s, "Active", "Form", now, profName, profPrograms]);
    }
    return;
  }

  const data = paSheet.getRange(2, 1, lastRow - 1, Math.max(paSheet.getLastColumn(), 7)).getValues();
  const existingRowsMap = new Map();

  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === professorId) {
      const key = `${data[i][2]}__${data[i][3]}`;
      existingRowsMap.set(key, { row: i + 2, state: data[i][4] });
    }
  }

  // 3. Activar y Desactivar
  for (const [key, info] of existingRowsMap.entries()) {
    if (requestedSlots.has(key)) {
      if (info.state !== "Active") {
        paSheet.getRange(info.row, 5).setValue("Active");
      }
      paSheet.getRange(info.row, 7).setValue(now);
      paSheet.getRange(info.row, 8, 1, 2).setValues([[profName, profPrograms]]);

      requestedSlots.delete(key);
    } else {
      if (info.state !== "Inactive") {
        paSheet.getRange(info.row, 5).setValue("Inactive");
        paSheet.getRange(info.row, 7).setValue(now);
      }
    }
  }

  // 4. Agregar nuevos slots
  for (const reqKey of requestedSlots) {
    const [dayOfWeek, slot] = reqKey.split("__");
    paSheet.appendRow([
      generateLinkId_("AFM-PA"),
      professorId,
      dayOfWeek,
      slot,
      "Active",
      "Form",
      now,
      profName,
      profPrograms
    ]);
  }
}

function getProfessorNameHelper_(ss, professorId) {
  const sheet = ss.getSheetByName("Professors");
  if (!sheet) return "-";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === professorId) {
      return `${data[i][1]} ${data[i][2]}`.trim();
    }
  }
  return "-";
}

function getProfessorProgramsHelper_(ss, professorId) {
  const ppSheet = ss.getSheetByName("Professor_Programs");
  const progSheet = ss.getSheetByName("Programs");
  if (!ppSheet || !progSheet) return "-";

  const progMap = {};
  const progData = progSheet.getDataRange().getValues();
  for (let i = 1; i < progData.length; i++) {
    progMap[String(progData[i][0]).trim()] = String(progData[i][1]).trim();
  }

  const ppData = ppSheet.getDataRange().getValues();
  const programs = [];
  for (let i = 1; i < ppData.length; i++) {
    if (String(ppData[i][1]).trim() === professorId && String(ppData[i][3]).trim() === "Active") {
      const pName = progMap[String(ppData[i][2]).trim()];
      if (pName && !programs.includes(pName)) programs.push(pName);
    }
  }
  return programs.join(", ") || "-";
}

/**
 * Rellena las columnas H (Nombre) e I (Programas) en 'Professor_Availability'
 * Ejecutar una vez para actualizar los registros existentes.
 */
function enrichProfessorAvailabilitySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paSheet = ss.getSheetByName("Professor_Availability");
  const pSheet = ss.getSheetByName("Professors");
  const ppSheet = ss.getSheetByName("Professor_Programs");
  const progSheet = ss.getSheetByName("Programs");

  if (!paSheet || !pSheet || !ppSheet || !progSheet) {
    throw new Error("Faltan hojas requeridas.");
  }

  // 1. Mapear Nombres de Profesores (Professors: A=ID, B=First, C=Last)
  const profNameMap = {};
  const profData = pSheet.getDataRange().getValues();
  for (let i = 1; i < profData.length; i++) {
    const id = String(profData[i][0]).trim();
    const fullName = `${profData[i][1]} ${profData[i][2]}`.trim();
    if (id) profNameMap[id] = fullName;
  }

  // 2. Mapear Nombres de Programas (Programs: A=ID, B=Name)
  const progNameMap = {};
  const progData = progSheet.getDataRange().getValues();
  for (let i = 1; i < progData.length; i++) {
    const id = String(progData[i][0]).trim();
    const name = String(progData[i][1]).trim();
    if (id) progNameMap[id] = name;
  }

  // 3. Mapear Programas por Profesor (Professor_Programs: B=ProfID, C=ProgID, D=State)
  const profProgramsMap = {};
  const ppData = ppSheet.getDataRange().getValues();
  for (let i = 1; i < ppData.length; i++) {
    const pId = String(ppData[i][1]).trim();
    const progId = String(ppData[i][2]).trim();
    const state = String(ppData[i][3]).trim();

    if (state === "Active" && pId && progId) {
      if (!profProgramsMap[pId]) profProgramsMap[pId] = [];
      const progName = progNameMap[progId] || progId;
      if (!profProgramsMap[pId].includes(progName)) {
        profProgramsMap[pId].push(progName);
      }
    }
  }

  // 4. Preparar encabezados en H1 e I1
  paSheet.getRange(1, 8).setValue("Professor Name");
  paSheet.getRange(1, 9).setValue("Programs");

  // 5. Construir los datos para las filas existentes
  const lastRow = paSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay filas de disponibilidad.");
    return;
  }

  const paData = paSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Leemos solo Col B (ProfessorID)
  const updateBatch = [];

  for (let i = 0; i < paData.length; i++) {
    const pId = String(paData[i][0]).trim();
    const profName = profNameMap[pId] || "-";
    const programs = (profProgramsMap[pId] || []).join(", ") || "-";

    updateBatch.push([profName, programs]);
  }

  // 6. Escribir de golpe en columnas H e I (Cols 8 y 9)
  paSheet.getRange(2, 8, updateBatch.length, 2).setValues(updateBatch);
  Logger.log(`✅ Se actualizaron ${updateBatch.length} filas en 'Professor_Availability' con Nombre y Programas.`);
}