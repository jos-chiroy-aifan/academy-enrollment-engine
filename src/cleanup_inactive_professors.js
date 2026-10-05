/****************************************************************
 * DEPURADOR DE ASIGNACIONES DE PROFESORES INACTIVOS
 * - Lee la hoja Professors para detectar quién NO está 'Active'.
 * - Actualiza Professor_Programs y Professor_Availability a 'Inactive'.
 * - No modifica la lógica existente de los otros scripts.
 ****************************************************************/

function cleanUpInactiveProfessorsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const profSheet = ss.getSheetByName("Professors");
  const ppSheet = ss.getSheetByName("Professor_Programs");
  const paSheet = ss.getSheetByName("Professor_Availability");

  if (!profSheet || !ppSheet || !paSheet) {
    throw new Error("Faltan hojas requeridas (Professors, Professor_Programs o Professor_Availability).");
  }

  // 1. Obtener el conjunto de profesores ACTIVOS
  const profData = profSheet.getDataRange().getValues();
  const activeProfIds = new Set();
  const allKnownProfIds = new Set();

  // Col A (0): ProfessorID, Col K (10): State
  for (let i = 1; i < profData.length; i++) {
    const profId = String(profData[i][0]).trim();
    const state = String(profData[i][10]).trim();
    if (profId) {
      allKnownProfIds.add(profId);
      if (state === "Active") {
        activeProfIds.add(profId);
      }
    }
  }

  Logger.log(`Profesores activos encontrados: ${activeProfIds.size}`);

  // 2. Depurar Professor_Programs (Col B: ProfessorID [idx 1], Col D: State [idx 3])
  const ppLastRow = ppSheet.getLastRow();
  let ppUpdatedCount = 0;

  if (ppLastRow >= 2) {
    const ppRange = ppSheet.getRange(2, 1, ppLastRow - 1, 4);
    const ppValues = ppRange.getValues();

    for (let i = 0; i < ppValues.length; i++) {
      const profId = String(ppValues[i][1]).trim();
      const currentState = String(ppValues[i][3]).trim();

      // Si el profesor NO está activo (o ya no existe en Professors) y su fila aún dice Active
      if (!activeProfIds.has(profId) && currentState === "Active") {
        ppSheet.getRange(i + 2, 4).setValue("Inactive");
        ppUpdatedCount++;
      }
    }
  }
  Logger.log(`Professor_Programs: ${ppUpdatedCount} registros cambiados a Inactive.`);

  // 3. Depurar Professor_Availability (Col B: ProfessorID [idx 1], Col E: State [idx 4])
  const paLastRow = paSheet.getLastRow();
  let paUpdatedCount = 0;

  if (paLastRow >= 2) {
    const paRange = paSheet.getRange(2, 1, paLastRow - 1, 5);
    const paValues = paRange.getValues();

    for (let i = 0; i < paValues.length; i++) {
      const profId = String(paValues[i][1]).trim();
      const currentState = String(paValues[i][4]).trim();

      // Si el profesor NO está activo y su disponibilidad aún dice Active
      if (!activeProfIds.has(profId) && currentState === "Active") {
        paSheet.getRange(i + 2, 5).setValue("Inactive");
        paUpdatedCount++;
      }
    }
  }
  Logger.log(`Professor_Availability: ${paUpdatedCount} registros cambiados a Inactive.`);

  SpreadsheetApp.flush();
  Logger.log("✅ Depuración finalizada con éxito.");
}