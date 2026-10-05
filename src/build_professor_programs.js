function upsertProfessorProgramsFromResponse_(ppSheet, professorId, programNames, programNameToId) {
  // Estructura Professor_Programs:
  // A ProfessorProgramID | B ProfessorID | C ProgramID | D State | E Notes

  // 1. Obtener IDs de programas solicitados en el formulario
  const requestedProgramIds = new Set();
  for (const name of programNames) {
    const pId = programNameToId[name];
    if (pId) requestedProgramIds.add(pId);
  }

  // 2. Buscar las filas que YA le pertenecen a este profesor
  const lastRow = ppSheet.getLastRow();
  
  // Si la hoja está totalmente vacía, pasamos directo a insertar
  if (lastRow < 2) {
    for (const pId of requestedProgramIds) {
      ppSheet.appendRow([generateLinkId_("AFM-PP"), professorId, pId, "Active", ""]);
    }
    return;
  }

  const data = ppSheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Leemos de A a D
  const existingRowsMap = new Map(); // Guardaremos ProgramID -> { row, state }

  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === professorId) { // Si la Col B coincide
      const progId = data[i][2]; // Col C
      existingRowsMap.set(progId, { row: i + 2, state: data[i][3] }); // Col D
    }
  }

  // 3. Activar, Desactivar o Ignorar
  for (const [pId, info] of existingRowsMap.entries()) {
    if (requestedProgramIds.has(pId)) {
      // Lo pidió y ya existía -> Aseguramos que diga "Active"
      if (info.state !== "Active") {
        ppSheet.getRange(info.row, 4).setValue("Active");
      }
      // Lo borramos del Set para saber que ya lo procesamos
      requestedProgramIds.delete(pId);
    } else {
      // Ya lo tenía en la DB, pero NO lo pidió en el nuevo form -> "Inactive"
      if (info.state !== "Inactive") {
        ppSheet.getRange(info.row, 4).setValue("Inactive");
      }
    }
  }

  // 4. Agregar los que quedaron en el Set (son totalmente nuevos)
  for (const pId of requestedProgramIds) {
    ppSheet.appendRow([
      generateLinkId_("AFM-PP"), // Tu generador de IDs
      professorId,
      pId,
      "Active",
      ""
    ]);
  }
}