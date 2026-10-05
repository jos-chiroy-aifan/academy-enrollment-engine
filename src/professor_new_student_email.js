/************************************
 * PROFESSOR EMAIL - NEW STUDENT ENROLLMENT
 * Template: "Nueva Asignación" (AFM)
 ************************************/

const ACADEMY_NAME = "Academia de Formación Musical y Artística";
const SHEET_PROFESSORS2 = "Professors";
const SHEET_PROGRAMS2 = "Programs";

/**
 * Sends an email to the professor when a new student is assigned.
 * @param {Object} payload { professorId, studentId, programId, dayKey, timeSlot, periods }
 */
function notifyProfessorNewStudent_(payload) {
  if (!payload || !payload.professorId) throw new Error("notifyProfessorNewStudent_: missing professorId");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName(SHEET_PROFESSORS2);
  const progSheet = ss.getSheetByName(SHEET_PROGRAMS2);

  if (!pSheet) throw new Error(`Sheet "${SHEET_PROFESSORS2}" not found.`);
  if (!progSheet) throw new Error(`Sheet "${SHEET_PROGRAMS2}" not found.`);

  const prof = getProfessorContactById_(pSheet, payload.professorId);
  if (!prof || !prof.email) {
    Logger.log(`⚠️ Could not send email: ProfessorID ${payload.professorId} has no email or is not Active.`);
    return;
  }

  const programName = getProgramNameById_(progSheet, payload.programId) || String(payload.programId || "").trim();

  const subject = `Nueva asignación | ${ACADEMY_NAME}`;
  const htmlBody = buildProfessorNewStudentEmailHtml_({
    academyName: ACADEMY_NAME,
    professorName: prof.fullName || "Profesor(a)",
    professorId: payload.professorId,
    studentId: String(payload.studentId || "").trim(),
    programName,
    dayEs: dayToSpanish_(payload.dayKey),
    timeSlot: String(payload.timeSlot || "").trim(),
    periods: Number(payload.periods || 1)
  });

  MailApp.sendEmail({
    to: String(prof.email).trim(),
    subject,
    htmlBody,
    name: ACADEMY_NAME
  });
}

/************************************
 * LOOKUPS
 ************************************/

/**
 * Retrieves the professor's email and full name by ID from the Professors sheet.
 */
function getProfessorContactById_(pSheet, professorId) {
  const last = pSheet.getLastRow();
  if (last < 2) return null;

  const headers = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];

  const cId = headers.indexOf("ProfessorID");
  const cFirst = headers.indexOf("First Name");
  const cLast = headers.indexOf("Last Name");
  const cEmail = headers.indexOf("Email");
  const cState = headers.indexOf("State");

  if (cId === -1 || cEmail === -1) {
    throw new Error(`"${SHEET_PROFESSORS2}" must have headers: ProfessorID, Email (and optionally First Name, Last Name, State).`);
  }

  const values = pSheet.getRange(2, 1, last - 1, pSheet.getLastColumn()).getValues();
  const target = String(professorId).trim();

  for (const r of values) {
    const id = String(r[cId] || "").trim();
    if (id !== target) continue;

    const state = (cState >= 0) ? String(r[cState] || "").trim() : "Active";
    if (state && state !== "Active") return null;

    const first = (cFirst >= 0) ? String(r[cFirst] || "").trim() : "";
    const lastName = (cLast >= 0) ? String(r[cLast] || "").trim() : "";
    const email = String(r[cEmail] || "").trim();

    return { email, fullName: `${first} ${lastName}`.trim() };
  }

  return null;
}

/**
 * Retrieves the program name by ID from the Programs sheet.
 */
function getProgramNameById_(progSheet, programId) {
  const last = progSheet.getLastRow();
  if (last < 2) return null;

  // Programs: A ProgramID, B Program Name
  const values = progSheet.getRange(2, 1, last - 1, 2).getValues();
  const target = String(programId || "").trim();

  for (const r of values) {
    const id = String(r[0] || "").trim();
    const name = String(r[1] || "").trim();
    if (id && id === target) return name;
  }
  return null;
}

/************************************
 * HTML GENERATION
 ************************************/

/**
 * Generates the HTML body for the professor notification email.
 */
function buildProfessorNewStudentEmailHtml_(d) {
  const esc = (v) => htmlEscape_(v);

  const slotPretty = esc(String(d.timeSlot || "").replace(" - ", " a "));
  const periodsTxt = (Number(d.periods || 1) === 2) ? "2 períodos" : "1 período";

  // Dynamic rows
  const rows = [
    row_("Carnet", esc(d.studentId)),
    row_("Programa", esc(d.programName)),
    row_("Día", esc(d.dayEs)),
    row_("Horario", slotPretty),
    row_("Duración", esc(periodsTxt)),
  ].join("");

  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Nuevo Aprendiz Asignado</title>
  <style>
    /* FONTS */
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;0,700;1,400&family=Lora:wght@400;500&family=Montserrat:wght@500;700&display=swap');

    body {
      margin: 0; padding: 0;
      background-color: #EBE9E4; /* Stone background */
      font-family: 'Lora', serif;
      color: #1A1A1A;
    }
    .wrapper { width: 100%; background-color: #EBE9E4; padding: 40px 0; }
    
    /* CARD */
    .card {
      max-width: 600px; margin: 0 auto;
      background-color: #FFFFFF;
      box-shadow: 0 25px 50px rgba(0,0,0,0.08);
      overflow: hidden;
    }

    /* HEADER */
    .header { 
      background-color: #FFFFFF;
      padding: 50px 30px 10px 30px; 
      text-align: center; 
    }
    
    .brand {
      font-family: 'Montserrat', sans-serif; 
      font-size: 10px; 
      letter-spacing: 3px; 
      text-transform: uppercase; 
      color: #111; 
      margin: 0; display: block; font-weight: 700;
    }
    
    /* CONTENT */
    .content { 
      background-color: #FFFFFF;
      padding: 20px 50px 50px 50px; 
    }

    /* TITLES */
    .chapter-container {
      text-align: center;
      margin-bottom: 40px;
      color: #B05E52; 
    }

    .chapter-main {
      font-family: 'Cormorant Garamond', serif; 
      font-size: 26px;
      font-weight: 700; 
      margin: 0;
      line-height: 1;
      text-transform: uppercase; 
      letter-spacing: 3px;
    }

    .chapter-sub {
      font-family: 'Cormorant Garamond', serif;
      font-size: 22px; 
      font-style: italic; 
      font-weight: 400;
      display: block;
      margin-top: 10px;
      letter-spacing: 1px;
      color: #B05E52; 
      opacity: 0.9;
    }

    /* INTRO TEXT */
    .intro-text {
      text-align: center;
      font-size: 15px;
      line-height: 1.6;
      color: #444;
      margin-bottom: 30px;
      padding: 0 10px;
    }
    
    /* DATA BOX */
    .data-box {
      background-color: #F9F8F6; 
      border-top: 4px solid #B05E52; 
      border-bottom: 1px solid #EAE8E0; 
      margin: 0 0 30px 0; 
      padding: 25px 30px; 
      border-radius: 2px;
    }

    .data-table { width: 100%; border-collapse: collapse; }
    
    /* Rows without bottom border as requested */
    .data-row td { 
      padding: 12px 0; 
      vertical-align: middle;
      border-bottom: none; 
    }

    /* LABELS & VALUES */
    .label {
      width: 30%; 
      font-family: 'Montserrat', sans-serif; 
      font-size: 9px;
      color: #888; 
      text-transform: uppercase; 
      letter-spacing: 2px;
      padding-top: 3px;
    }
    
    .value {
      font-family: 'Lora', serif; 
      font-size: 15px; 
      color: #222; 
      font-weight: 400; 
    }

    /* FOOTER */
    .footer {
      background-color: #F7F5F2; 
      border-top: 1px solid #EAE8E0;
      padding: 40px 20px; 
      text-align: center;
    }

    .footer-copy {
      font-family: 'Montserrat', sans-serif; font-size: 9px;
      color: #BBB; letter-spacing: 1px; margin: 0;
    }
    
    .secret-code {
      display: block;
      font-family: 'Montserrat', sans-serif;
      font-size: 8px;
      letter-spacing: 4px;
      margin-top: 14px;
      opacity: 0.18;
      color: #CFCAC2;
      user-select: none;
    }

    /* MOBILE */
    @media only screen and (max-width: 600px) {
      .card { width: 100%; box-shadow: none; margin: 0; }
      .wrapper { padding: 0; }
      .header, .footer { padding: 30px 20px; }
      .content { padding: 20px 20px 40px 20px; }
      .data-box { padding: 20px; }
      .value { display: block; padding-top: 2px; margin-bottom: 8px; }
      .label { display: block; width: 100%; margin-bottom: 0px; }
    }
  </style>
</head>
<body>

  <div class="wrapper">
    <div class="card">
      
      <div class="header">
        <span class="brand">${esc(d.academyName)}</span>
      </div>

      <div class="content">
        
        <div class="chapter-container">
            <h1 class="chapter-main">NUEVA ASIGNACIÓN</h1>
            <span class="chapter-sub">Nuevo Aprendiz</span>
        </div>

        <div class="intro-text">
          Estimado(a) Mtro(a). <strong>${esc(d.professorName)}</strong>.<br>
          Un nuevo estudiante ha sido inscrito en tu programa.
        </div>

        <div class="data-box">
          <table class="data-table">
            ${rows}
          </table>
        </div>

        <div class="intro-text" style="font-size: 13px; color: #666; margin-bottom: 0;">
          Si existe algún conflicto con este horario, por favor responder a este correo.
        </div>

      </div>

      <div class="footer">
        <p class="footer-copy">
          &copy; 2025 AFM
        </p>
        <span class="secret-code">L1602</span>
      </div>

    </div>
  </div>

</body>
</html>
  `.trim();

  function row_(label, value) {
    return `
      <tr class="data-row">
        <td class="label">${label}</td>
        <td class="value">${value}</td>
      </tr>
    `;
  }
}

/************************************
 * UTILS
 ************************************/

/**
 * Basic HTML escape to prevent injection.
 */
function htmlEscape_(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}