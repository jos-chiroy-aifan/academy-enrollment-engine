function migrateApprovedProfessorResponses() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rSheet = ss.getSheetByName("Professor Registration - Responses");
  const pSheet = ss.getSheetByName("Professors");

  const progSheet = ss.getSheetByName("Programs");
  const ppSheet = ss.getSheetByName("Professor_Programs");
  const paSheet = ss.getSheetByName("Professor_Availability");

  const lastRow = rSheet.getLastRow();
  const lastCol = rSheet.getLastColumn();
  if (lastRow < 2) return;

  // Headers dinámicos
  const headers = rSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = (name) => {
    const i = headers.indexOf(name);
    if (i === -1) throw new Error(`No se encontró la columna "${name}" en "Professor Registration - Responses".`);
    return i;
  };

  const I = {
    firstName: idx("Nombre(s)"),
    lastName: idx("Apellido(s)"),
    birthdate: idx("Fecha de Nacimiento"),
    gender: idx("Género"),
    email: idx("Correo electrónico"),
    phone: idx("Número de teléfono"),
    program: idx("Programa"),
    mon: idx("Horarios Disponibles - Lunes"),
    tue: idx("Horarios Disponibles - Martes"),
    wed: idx("Horarios Disponibles - Miércoles"),
    thu: idx("Horarios Disponibles - Jueves"),
    fri: idx("Horarios Disponibles - Viernes"),
    sat: idx("Horarios Disponibles - Sábado"),
    sun: idx("Horarios Disponibles - Domingo"),
    consentImage: idx("Autorización de uso de imagen"),
    terms: idx("Términos de participación y tratamiento de datos"),
    status: idx("Status"),
    lastProcessedAt: idx("LastProcessedAt"),
  };

  const data = rSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Map: Program Name -> ProgramID (solo activos)
  const programNameToId = buildActiveProgramMap_(progSheet);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentStatus = row[I.status];

    // ✅ FILTRO: Acepta "Approved" (Nuevo) o "Update" (Actualización)
    if (currentStatus !== "Approved" && currentStatus !== "Update") continue;

    let professorId = "";

    // -----------------------------------------------------
    // LÓGICA 1: PROFESOR NUEVO ("Approved")
    // -----------------------------------------------------
    if (currentStatus === "Approved") {
      // Normalizar género
      const genderMap = {
        "Masculino": "Male",
        "Femenino": "Female",
        "Prefiero no decirlo": "Prefer not to say",
      };
      const gender = genderMap[row[I.gender]] || "Prefer not to say";

      // Consentimientos
      const consentPhotos = String(row[I.consentImage]).toLowerCase().includes("sí");
      const termsOk = String(row[I.terms]).toLowerCase().includes("acepto");
      const birthdate = row[I.birthdate];

      const newRow = [
        "",                 // ProfessorID (se genera luego)
        row[I.firstName],   // First Name
        row[I.lastName],    // Last Name
        birthdate,          // Birthdate
        "",                 // Age (formula)
        gender,             // Gender
        row[I.email],       // Email
        row[I.phone],       // Phone
        consentPhotos,      // Consent for Photos
        termsOk,            // Consent for Data Processing
        "Active",           // State
        ""                  // Notes
      ];

      pSheet.appendRow(newRow);
      const profRow = pSheet.getLastRow();

      // Age formula
      pSheet.getRange(profRow, 5).setFormula(`=IF(D${profRow}="","",DATEDIF(D${profRow},TODAY(),"Y"))`);

      // Generar IDs
      generateProfessorIds();

      // Leer el ProfessorID ya asignado
      professorId = pSheet.getRange(profRow, 1).getValue();
      if (!professorId) throw new Error("No se pudo obtener ProfessorID después de generateProfessorIds().");

      // Si por alguna razón el profe no está Active, no lo consideramos
      const profState = pSheet.getRange(profRow, 11).getValue(); // K State
      if (profState !== "Active") {
        rSheet.getRange(i + 2, I.status + 1).setValue("Migrated");
        rSheet.getRange(i + 2, I.lastProcessedAt + 1).setValue(new Date());
        continue;
      }
    } 
    // -----------------------------------------------------
    // LÓGICA 2: ACTUALIZAR PROFESOR EXISTENTE ("Update")
    // -----------------------------------------------------
    else if (currentStatus === "Update") {
      const emailToSearch = String(row[I.email]).trim().toLowerCase();
      professorId = buscarProfesorPorEmail_(pSheet, emailToSearch);
      
      if (!professorId) {
        rSheet.getRange(i + 2, I.status + 1).setValue("Error: No Encontrado");
        rSheet.getRange(i + 2, I.lastProcessedAt + 1).setValue(new Date());
        continue;
      }
    }

    // -----------------------------------------------------
    // RECONSTRUCCIÓN DE PROGRAMAS Y HORARIOS
    // -----------------------------------------------------
    const programNames = splitMulti_(row[I.program]);
    upsertProfessorProgramsFromResponse_(ppSheet, professorId, programNames, programNameToId);

    const byDay = [
      ["Monday", row[I.mon]],
      ["Tuesday", row[I.tue]],
      ["Wednesday", row[I.wed]],
      ["Thursday", row[I.thu]],
      ["Friday", row[I.fri]],
      ["Saturday", row[I.sat]],
      ["Sunday", row[I.sun]],
    ];
    upsertProfessorAvailabilityFromResponse_(paSheet, professorId, byDay);

    // -----------------------------------------------------
    // MARCAR MIGRADO Y ENVIAR CORREO
    // -----------------------------------------------------
    rSheet.getRange(i + 2, I.lastProcessedAt + 1).setValue(new Date());

    if (currentStatus === "Approved") {
      rSheet.getRange(i + 2, I.status + 1).setValue("Migrated");
      
      // Enviar correo SOLO si es nuevo
      sendProfessorRegistrationEmail_({
        email: row[I.email],
        professorId,
        firstName: row[I.firstName],
        lastName: row[I.lastName],
        birthdate: row[I.birthdate],
        gender: row[I.gender], // NOTA: Aquí tenías una variable 'gender' que si era update iba a fallar, la saqué directo del row para evitar el error.
        phone: row[I.phone],
        programs: programNames
      });
    } else {
      // Si fue Update, solo cambiamos la etiqueta
      rSheet.getRange(i + 2, I.status + 1).setValue("Updated");
    }
  }
}

// ✅ NUEVO HELPER: Busca al profe por correo
function buscarProfesorPorEmail_(pSheet, email) {
  const data = pSheet.getDataRange().getValues();
  // Asumiendo Col A (idx 0) = ID, Col G (idx 6) = Email, Col K (idx 10) = State
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]).trim().toLowerCase() === email && String(data[i][10]) === "Active") {
      return data[i][0];
    }
  }
  return null;
}