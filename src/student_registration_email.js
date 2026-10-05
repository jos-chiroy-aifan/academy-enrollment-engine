/************************************
 * STUDENT EMAIL - SCHEDULE LINK
 * Academia de Formación Musical y Artística
 * Diseño: "Primer Capítulo" (Atelier V3) - EDITADO
 ************************************/

const SCHEDULES_FORM_PUBLIC_URL = PropertiesService.getScriptProperties().getProperty("SCHEDULES_FORM_PUBLIC_URL");

/**
 * Envía la carta de bienvenida "Primer Capítulo" al alumno.
 *
 * @param {Object} student  Datos: {email, studentId, firstName, lastName, birthdate, gender, phone}
 * @param {string} [scheduleUrl] URL del form (opcional)
 */
function sendStudentScheduleEmail_(student, scheduleUrl = SCHEDULES_FORM_PUBLIC_URL) {
  if (!student || !student.email) throw new Error("sendStudentScheduleEmail_: falta student.email");

  const academyName = "Academia de Formación Musical y Artística";
  const subject = `Registro Confirmado | ${academyName}`;
  const htmlBody = buildStudentScheduleEmailHtml_(student, scheduleUrl);

  MailApp.sendEmail({
    to: String(student.email).trim(),
    subject,
    htmlBody,
    name: academyName
  });
}

/**
 * Genera el HTML del correo con el diseño "Primer Capítulo".
 * Editado:
 * - Botón más bonito en desktop (CSS + inline styles)
 * - L1602 mucho más oculto (marca fantasma)
 * - Fallback link menos llamativo
 */
function buildStudentScheduleEmailHtml_(student, scheduleUrl) {

  // --- HELPERS INTERNOS ---
  const esc = (v) => htmlEscape_(v);
  const fmtDate = (d) => {
    try {
      if (!d) return "";
      const dt = (d instanceof Date) ? d : new Date(d);
      if (isNaN(dt.getTime())) return String(d);
      return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } catch (_) {
      return String(d || "");
    }
  };

  // --- PREPARACIÓN DE DATOS ---
  const fullName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
  const studentId = student.studentId || student.carnet || "Pendiente";
  const phone = student.phone || "";
  const birthdate = student.birthdate ? fmtDate(student.birthdate) : "";
  const gender = student.gender || "";

  // --- CONSTRUCCIÓN DE FILAS DE LA FICHA ---
  const rows = [];
  if (studentId) rows.push(row_("Carnet", esc(studentId)));
  if (fullName) rows.push(row_("Nombre", esc(fullName)));
  if (birthdate) rows.push(row_("Fecha", esc(birthdate)));
  if (gender) rows.push(row_("Género", esc(gender)));
  if (phone) rows.push(row_("Teléfono", esc(phone)));
  if (student.email) rows.push(row_("Correo", esc(student.email)));

  const summaryTable = rows.length
    ? `<table class="data-table" width="100%" cellspacing="0" cellpadding="0" border="0">${rows.join("")}</table>`
    : "";

  // --- TEMPLATE HTML ---
  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Primer Capítulo</title>

  <style>
    /* FUENTES */
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;0,700;1,400&family=Lora:wght@400;500&family=Montserrat:wght@500;700&display=swap');

    body {
      margin: 0; padding: 0;
      background-color: #EBE9E4; /* Fondo exterior (Piedra) */
      font-family: 'Lora', serif;
      color: #1A1A1A;
    }

    .wrapper { width: 100%; background-color: #EBE9E4; padding: 40px 0; }

    /* TARJETA */
    .card {
      max-width: 600px; margin: 0 auto;
      background-color: #FFFFFF;
      box-shadow: 0 25px 50px rgba(0,0,0,0.08);
      overflow: hidden;
    }

    /* --- HEADER (BLANCO) --- */
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
      margin: 0;
      display: block;
      font-weight: 700;
    }

    /* --- CONTENIDO --- */
    .content {
      background-color: #FFFFFF;
      padding: 20px 50px 50px 50px;
    }

    /* TÍTULO NARRATIVO */
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

    /* CAJA DE DATOS (FICHA) */
    .data-box {
      background-color: #F9F8F6; /* Beige suave */
      border-top: 4px solid #B05E52; /* Acento Terracota */
      border-bottom: 1px solid #EAE8E0;
      margin: 0 0 40px 0;
      padding: 25px 30px;
      border-radius: 2px;
    }

    /* ETIQUETAS */
    .label {
      width: 30%;
      font-family: 'Montserrat', sans-serif;
      font-size: 9px;
      color: #888;
      text-transform: uppercase;
      letter-spacing: 2px;
      padding-top: 3px;
    }

    /* VALORES */
    .value {
      font-family: 'Lora', serif;
      font-size: 15px;
      color: #222;
      font-weight: 400;
    }

    /* BOTÓN */
    .btn-wrapper { text-align: center; margin-bottom: 20px; }

    .instruction {
      font-family: 'Montserrat', sans-serif; font-size: 9px;
      text-transform: uppercase; letter-spacing: 1px; color: #AAA;
      margin-bottom: 15px; display: block;
    }

    /* Botón mejorado para desktop */
    .btn {
      display: inline-block;
      background-color: #111;
      color: #fff !important;
      text-decoration: none;
      padding: 16px 48px;
      font-family: 'Montserrat', sans-serif;
      font-size: 11px;
      letter-spacing: 3px;
      text-transform: uppercase;
      border-radius: 6px;
      box-shadow: 0 10px 28px rgba(0,0,0,0.22);
      border: 1px solid rgba(255,255,255,0.08);
    }

    .btn:hover { background-color: #B05E52; }

    /* LINK TEXTO (fallback) */
    .link-text {
      text-align: center;
      font-size: 10px;
      color: rgba(0,0,0,0.45);
      margin-top: 26px;
      font-family: 'Montserrat', sans-serif;
    }

    .link-text a {
      color: #B05E52;
      text-decoration: none;
      border-bottom: 1px solid rgba(176,94,82,0.25);
    }

    /* --- FOOTER (ARENA) --- */
    .footer {
      background-color: #F7F5F2;
      border-top: 1px solid #EAE8E0;
      padding: 40px 20px;
      text-align: center;
    }

    /* FRASE */
    .art-quote {
      font-family: 'Cormorant Garamond', serif;
      font-size: 18px;
      font-style: italic;
      color: #555;
      margin-bottom: 20px;
      padding: 0 20px;
    }

    .footer-copy {
      font-family: 'Montserrat', sans-serif; font-size: 9px;
      color: #BBB; letter-spacing: 1px; margin: 0;
    }

    /* L1602 más oculto */
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
        <span class="brand">Academia de Formación Musical y Artística</span>
      </div>

      <div class="content">
        <div class="chapter-container">
          <h1 class="chapter-main">PRIMER CAPÍTULO</h1>
          <span class="chapter-sub">Artista</span>
        </div>

        <div class="data-box">
          ${summaryTable}
        </div>

        <div class="btn-wrapper">
          <span class="instruction">Siguiente paso</span>

          <!-- Inline styles para que Gmail lo respete en desktop -->
          <a href="${esc(scheduleUrl)}"
             class="btn"
             style="
               display:inline-block;
               background:#111;
               color:#fff !important;
               text-decoration:none;
               padding:16px 48px;
               font-family:Montserrat, Arial, sans-serif;
               font-size:11px;
               letter-spacing:3px;
               text-transform:uppercase;
               border-radius:6px;
               box-shadow:0 10px 28px rgba(0,0,0,0.22);
               border:1px solid rgba(255,255,255,0.08);
             ">
            Definir Horario
          </a>
        </div>

        <div class="link-text">
          Si el botón no funciona:<br>
          <a href="${esc(scheduleUrl)}">${esc(scheduleUrl)}</a>
        </div>
      </div>

      <div class="footer">
        <div class="art-quote">
          "Ser aprendiz es el origen de todo arte."
        </div>

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

  // Función auxiliar para crear filas de la tabla
  function row_(label, value) {
    return `
    <tr class="data-row">
      <td class="label">${label}</td>
      <td class="value">${value}</td>
    </tr>`;
  }
}

/**
 * Escape básico para prevenir inyección HTML
 */
function htmlEscape_(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}