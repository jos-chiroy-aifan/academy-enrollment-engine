/************************************
 * SCHEDULE CONFIRMATION EMAIL
 * Academia de Formación Musical y Artística
 * Diseño: "Segundo Capítulo" (Atelier V3)
 ************************************/

const AFM_ACADEMY_NAME = "Academia de Formación Musical y Artística";

/**
 * Envía correo de confirmación de horario ("Segundo Capítulo").
 *
 * @param {Object} payload
 *  - email (string) requerido
 *  - studentId (string) requerido
 *  - firstName (string)
 *  - lastName (string)
 *  - programName (string) requerido
 *  - dayOfWeek (string) requerido  (puede venir "Monday" o "Lunes")
 *  - timeSlot (string) requerido   (ej: "8:00 - 8:40" o label)
 *  - periods (number) opcional (1/2)
 *  - professorName (string) opcional
 *  - notes (string) opcional
 */
function sendScheduleConfirmationEmail_(payload) {
  if (!payload || !payload.email) throw new Error("sendScheduleConfirmationEmail_: falta payload.email");
  if (!payload.studentId) throw new Error("sendScheduleConfirmationEmail_: falta payload.studentId");
  if (!payload.programName) throw new Error("sendScheduleConfirmationEmail_: falta payload.programName");
  if (!payload.dayOfWeek) throw new Error("sendScheduleConfirmationEmail_: falta payload.dayOfWeek");
  if (!payload.timeSlot) throw new Error("sendScheduleConfirmationEmail_: falta payload.timeSlot");

  const subject = `Horario Confirmado | ${AFM_ACADEMY_NAME}`;
  const htmlBody = buildScheduleConfirmationEmailHtml_(payload);

  MailApp.sendEmail({
    to: String(payload.email).trim(),
    subject,
    htmlBody,
    name: AFM_ACADEMY_NAME
  });
}

/**
 * HTML "Segundo Capítulo".
 * - Mismo look&feel que "Primer Capítulo"
 * - Título: SEGUNDO CAPÍTULO / Arte
 * - Incluye resumen del horario confirmado
 */
function buildScheduleConfirmationEmailHtml_(p) {
  const esc = (v) => htmlEscape_(v);

  const fullName = `${p.firstName || ""} ${p.lastName || ""}`.trim();
  const studentId = p.studentId || "—";
  const programName = p.programName || "—";

  // Normalizamos día a español bonito (si te llega Monday, lo vuelve Lunes)
  const dayEs = dayToSpanish_(p.dayOfWeek);
  const timeLabel = normalizeTimeLabelEs_(p.timeSlot); // "8:00 - 8:40" -> "8:00 a 8:40" si aplica

  const periods = Number(p.periods || 1);
  const periodsLabel = (periods === 2) ? "2 períodos continuos" : "1 período";

  const rows = [];
  rows.push(row_("Carnet", esc(studentId)));
  if (fullName) rows.push(row_("Nombre", esc(fullName)));
  rows.push(row_("Programa", esc(programName)));
  rows.push(row_("Horario", esc(`${dayEs} - ${timeLabel}`)));
  rows.push(row_("Duración", esc(periodsLabel)));

  if (p.professorName) rows.push(row_("Profesor(a)", esc(p.professorName)));
  if (p.email) rows.push(row_("Correo", esc(p.email)));
  if (p.notes) rows.push(row_("Notas", esc(p.notes)));

  const summaryTable = `<table class="data-table" width="100%" cellspacing="0" cellpadding="0" border="0">${rows.join("")}</table>`;

  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Segundo Capítulo</title>

  <style>
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;0,700;1,400&family=Lora:wght@400;500&family=Montserrat:wght@500;700&display=swap');

    body {
      margin: 0; padding: 0;
      background-color: #EBE9E4;
      font-family: 'Lora', serif;
      color: #1A1A1A;
    }

    .wrapper { width: 100%; background-color: #EBE9E4; padding: 40px 0; }

    .card {
      max-width: 600px; margin: 0 auto;
      background-color: #FFFFFF;
      box-shadow: 0 25px 50px rgba(0,0,0,0.08);
      overflow: hidden;
    }

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

    .content {
      background-color: #FFFFFF;
      padding: 20px 50px 50px 50px;
    }

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

    .data-box {
      background-color: #F9F8F6;
      border-top: 4px solid #B05E52;
      border-bottom: 1px solid #EAE8E0;
      margin: 0 0 18px 0;
      padding: 25px 30px;
      border-radius: 2px;
    }

    .label {
      width: 30%;
      font-family: 'Montserrat', sans-serif;
      font-size: 9px;
      color: #888;
      text-transform: uppercase;
      letter-spacing: 2px;
      padding-top: 3px;
      vertical-align: top;
    }

    .value {
      font-family: 'Lora', serif;
      font-size: 15px;
      color: #222;
      font-weight: 400;
      padding-bottom: 8px;
    }

    .note {
      text-align: center;
      font-family: 'Montserrat', sans-serif;
      font-size: 10px;
      letter-spacing: 0.6px;
      color: rgba(0,0,0,0.55);
      margin: 26px 0 0 0;
      line-height: 1.6;
    }

    .footer {
      background-color: #F7F5F2;
      border-top: 1px solid #EAE8E0;
      padding: 40px 20px;
      text-align: center;
    }

    .art-quote {
      font-family: 'Cormorant Garamond', serif;
      font-size: 18px;
      font-style: italic;
      color: #555;
      margin-bottom: 20px;
      padding: 0 20px;
    }

    .footer-copy {
      font-family: 'Montserrat', sans-serif;
      font-size: 9px;
      color: #BBB;
      letter-spacing: 1px;
      margin: 0;
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
          <h1 class="chapter-main">SEGUNDO CAPÍTULO</h1>
          <span class="chapter-sub">Arte</span>
        </div>

        <div class="data-box">
          ${summaryTable}
        </div>

        <p class="note">
          Tu horario ha sido registrado correctamente.<br/>
          Si necesitas hacer un cambio, por favor responde a este correo o comunícate con la administración.
        </p>
      </div>

      <div class="footer">
        <div class="art-quote">
          "El arte comienza cuando el tiempo encuentra su forma."
        </div>

        <p class="footer-copy">&copy; 2025 AFM</p>
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
    </tr>`;
  }
}

/**
 * Convierte timeSlot a label ES si viene en formato "8:00 - 8:40".
 * Si ya viene como "8:00 a 8:40" o "Lunes - 8:00 a 8:40", lo deja limpio.
 */
function normalizeTimeLabelEs_(timeSlot) {
  const t = String(timeSlot || "").trim();
  if (!t) return "";

  // Si viene "Lunes - 8:00 a 8:40", nos quedamos con la parte de horas
  const maybeParts = t.split(" - ");
  const looksLikeDay = maybeParts.length >= 2 && /[A-Za-zÁÉÍÓÚáéíóúñÑ]/.test(maybeParts[0]);
  const timePart = looksLikeDay ? maybeParts.slice(1).join(" - ").trim() : t;

  // "8:00 - 8:40" -> "8:00 a 8:40"
  return timePart.replace(/\s*-\s*/g, " a ").replace(/\s{2,}/g, " ").trim();
}

/**
 * Normaliza día a "Monday..Saturday" (Primera letra mayúscula)
 */
function normalizeDayKey_(day) {
  const d = String(day || "").trim().toLowerCase();

  // Español
  if (d === "lunes") return "Monday";
  if (d === "martes") return "Tuesday";
  if (d === "miércoles" || d === "miercoles") return "Wednesday";
  if (d === "jueves") return "Thursday";
  if (d === "viernes") return "Friday";
  if (d === "sábado" || d === "sabado") return "Saturday";
  if (d === "domingo") return "Sunday";

  // Inglés
  if (d === "monday" || d === "mon") return "Monday";
  if (d === "tuesday" || d === "tue" || d === "tues") return "Tuesday";
  if (d === "wednesday" || d === "wed") return "Wednesday";
  if (d === "thursday" || d === "thu" || d === "thur" || d === "thurs") return "Thursday";
  if (d === "friday" || d === "fri") return "Friday";
  if (d === "saturday" || d === "sat") return "Saturday";
  if (d === "sunday" || d === "sun") return "Sunday";

  const up = String(day || "").trim().toUpperCase();
  const allowed = new Set(["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]);
  if (allowed.has(up)) return up.charAt(0) + up.slice(1).toLowerCase();

  return String(day || "").trim();
}

/**
 * DayKey -> Español
 */
function dayToSpanish_(dayKey) {
  const k = normalizeDayKey_(dayKey);
  const map = {
    "Monday": "Lunes",
    "Tuesday": "Martes",
    "Wednesday": "Miércoles",
    "Thursday": "Jueves",
    "Friday": "Viernes",
    "Saturday": "Sábado",
    "Sunday": "Domingo",
  };
  return map[k] || String(dayKey || "").trim();
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