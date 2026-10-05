/************************************
 * PROFESSOR EMAIL - REGISTRATION CONFIRMED
 * Academia de Formación Musical y Artística
 * Estilo: Credencial Docente (Atelier V3 - Final)
 ************************************/

const ACADEMY_NAME2 = "Academia de Formación Musical y Artística";

/**
 * Envía correo de confirmación al Profesor.
 *
 * @param {Object} prof Datos: {email, professorId, firstName, lastName, birthdate, gender, phone, programs[], notes}
 */
function sendProfessorRegistrationEmail_(prof) {
  if (!prof || !prof.email) throw new Error("sendProfessorRegistrationEmail_: falta prof.email");

  const subject = `Acreditación Docente | ${ACADEMY_NAME2}`;
  const htmlBody = buildProfessorRegistrationEmailHtml_(prof);

  MailApp.sendEmail({
    to: String(prof.email).trim(),
    subject,
    htmlBody,
    name: ACADEMY_NAME2,
  });
}

/**
 * Genera el HTML del correo para Profesores.
 * Mantiene la estética "Atelier V3" con el L1602 fantasma.
 */
function buildProfessorRegistrationEmailHtml_(prof) {
  
  // --- HELPERS ---
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

  // --- DATOS ---
  const fullName = `${prof.firstName || ""} ${prof.lastName || ""}`.trim();
  const professorId = prof.professorId || "Pendiente";
  const phone = prof.phone || "";
  const birthdate = prof.birthdate ? fmtDate(prof.birthdate) : "";
  const gender = prof.gender || "";
  const programs = Array.isArray(prof.programs) ? prof.programs.filter(Boolean).join(", ") : "";
  
  // --- FILAS DE LA FICHA ---
  const rows = [];
  rows.push(row_("Código", esc(professorId)));
  if (fullName) rows.push(row_("Nombre", esc(fullName)));
  if (birthdate) rows.push(row_("Fecha", esc(birthdate)));
  if (gender) rows.push(row_("Género", esc(gender)));
  if (phone) rows.push(row_("Teléfono", esc(phone)));
  if (prof.email) rows.push(row_("Correo", esc(prof.email)));
  if (programs) rows.push(row_("Programas", esc(programs)));

  const summaryTable = `<table class="data-table" width="100%" cellspacing="0" cellpadding="0" border="0">${rows.join("")}</table>`;

  // --- INTRO Y NOTAS ---
  const intro = 
    `Es un honor sumar tu experiencia a nuestro equipo.<br>` +
    `Tu perfil ha sido habilitado para la gestión académica.`;

  const extraNote = prof.notes 
    ? `<div class="notes-container">Nota: ${esc(prof.notes)}</div>` 
    : "";

  // --- TEMPLATE HTML ---
  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Acreditación Docente</title>
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

    /* HEADER (BLANCO) */
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

    /* CONTENIDO */
    .content { 
      background-color: #FFFFFF;
      padding: 20px 50px 50px 50px; 
    }

    /* TÍTULO JERÁRQUICO */
    .title-block {
      text-align: center;
      margin-bottom: 35px;
      color: #B05E52; 
    }

    .main-title {
      font-family: 'Cormorant Garamond', serif; 
      font-size: 26px;
      font-weight: 700; 
      margin: 0;
      line-height: 1;
      text-transform: uppercase; 
      letter-spacing: 3px;
    }

    .sub-title {
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

    /* TEXTO INTRODUCTORIO */
    .intro-text {
      text-align: center;
      font-size: 15px;
      line-height: 1.6;
      color: #444;
      margin-bottom: 30px;
      padding: 0 20px;
    }
    
    /* FICHA TÉCNICA (Beige + Terracota) */
    .data-box {
      background-color: #F9F8F6; 
      border-top: 4px solid #B05E52; 
      border-bottom: 1px solid #EAE8E0; 
      margin: 0 0 30px 0; 
      padding: 25px 30px; 
      border-radius: 2px;
    }

    .data-table { width: 100%; border-collapse: collapse; }
    .data-row td { 
      padding: 12px 0; 
      vertical-align: middle;
      border-bottom: 1px solid #EAE8E0; 
    }
    .data-row:last-child td { border-bottom: none; }

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

    /* NOTAS */
    .notes-container {
      font-size: 13px; color: #666; font-style: italic; 
      background: #FAFAFA; border: 1px solid #EEE; padding: 15px; text-align: center;
      margin-top: 20px;
    }

    /* FOOTER (ARENA) */
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
      font-family: 'Montserrat', sans-serif; font-size: 9px;
      color: #BBB; letter-spacing: 1px; margin: 0;
    }
    
    /* L1602 FANTASMA (Actualizado según tu cambio) */
    .secret-code {
      display: block;
      font-family: 'Montserrat', sans-serif;
      font-size: 8px;
      letter-spacing: 4px;
      margin-top: 14px;
      opacity: 0.18;      /* Opacidad baja para efecto marca de agua */
      color: #CFCAC2;     /* Color sutil */
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
        <span class="brand">${esc(ACADEMY_NAME2)}</span>
      </div>

      <div class="content">
        
        <div class="title-block">
            <h1 class="main-title">ACREDITACIÓN DOCENTE</h1>
            <span class="sub-title">Claustro Académico</span>
        </div>

        <div class="intro-text">
          ${intro}
        </div>

        <div class="data-box">
          ${summaryTable}
        </div>

        ${extraNote}

      </div>

      <div class="footer">
        
        <div class="art-quote">
          "Enseñar también es una forma de arte."
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

  function row_(label, value) {
    return `
    <tr class="data-row">
      <td class="label">${label}</td>
      <td class="value">${value}</td>
    </tr>`;
  }
}

/** Escape básico */
function htmlEscape_(v) {
  return String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}