/****************************************************************
 * FINANCIAL REPORT EMAIL TEMPLATE (VISTA)
 * Diseño: "Primer Capítulo" (Atelier V3) - Adaptado a Finanzas
 * Estilos: Beige, Fuentes Serif, Tablas limpias.
 * Firma: L1602
 ****************************************************************/

/**
 * Genera el HTML completo del reporte financiero.
 * * @param {Array} overdue Listado de objetos {name, amount, dateObj}
 * @param {Array} thisWeek Listado de objetos {name, amount, dateObj}
 * @param {Date} startDate Inicio de semana
 * @param {Date} endDate Fin de semana
 */
function buildFinancialReportEmailHtml_(overdue, thisWeek, startDate, endDate) {

  // Helpers de formato
  const fmt = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM");
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  
  const rangeStr = `${fmt(startDate)} - ${fmt(endDate)}`;
  
  // Calcular Totales para el resumen
  const totalOverdue = overdue.reduce((sum, item) => sum + item.amount, 0);
  const totalWeek = thisWeek.reduce((sum, item) => sum + item.amount, 0);

  // --- CONSTRUCCIÓN DEL CONTENIDO INTERNO ---
  let contentHtml = "";

  // 1. SECCIÓN VENCIDOS (Rojo Terracota)
  if (overdue.length > 0) {
    contentHtml += `
      <div class="section-title" style="color:#B05E52;">⚠ ACUMULADO VENCIDO</div>
      ${buildFinanceTable_(overdue)}
      <div class="total-row">Total Vencido: Q${totalOverdue}</div>
      <div class="separator"></div>
    `;
  } else {
    contentHtml += `
      <div class="clean-state">✨ No hay deuda acumulada de semanas anteriores.</div>
      <div class="separator"></div>
    `;
  }

  // 2. SECCIÓN ESTA SEMANA (Oscuro)
  if (thisWeek.length > 0) {
    contentHtml += `
      <div class="section-title" style="color:#1A1A1A;">📅 COBROS DE ESTA SEMANA</div>
      ${buildFinanceTable_(thisWeek)}
      <div class="total-row">Total Semana: Q${totalWeek}</div>
    `;
  } else {
    contentHtml += `
      <div class="clean-state">No hay cobros programados para esta semana.</div>
    `;
  }

  // --- TEMPLATE PRINCIPAL (HTML + CSS INLINE) ---
  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <title>Reporte Financiero</title>
  <style>
    /* FUENTES */
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;0,700;1,400&family=Lora:wght@400;700&family=Montserrat:wght@500;700&display=swap');
    
    body { margin:0; padding:0; background-color:#EBE9E4; font-family:'Lora', serif; color:#1A1A1A; }
    .wrapper { width:100%; background-color:#EBE9E4; padding:40px 0; }
    
    /* TARJETA */
    .card { max-width:600px; margin:0 auto; background-color:#FFFFFF; box-shadow:0 25px 50px rgba(0,0,0,0.08); overflow:hidden; }
    
    /* HEADER */
    .header { background-color:#FFFFFF; padding:50px 30px 20px 30px; text-align:center; }
    .brand { font-family:'Montserrat', sans-serif; font-size:10px; letter-spacing:3px; text-transform:uppercase; font-weight:700; display:block; color:#111; }
    
    /* TITULO */
    .chapter-container { text-align:center; margin-bottom:30px; color:#B05E52; }
    .chapter-main { font-family:'Cormorant Garamond', serif; font-size:26px; font-weight:700; margin:0; line-height:1; text-transform:uppercase; letter-spacing:3px; }
    .chapter-sub { font-family:'Cormorant Garamond', serif; font-size:18px; font-style:italic; margin-top:5px; display:block; color:#555; }
    
    .content { padding:0 50px 50px 50px; }
    
    /* ESTILOS DE TABLA FINANCIERA */
    .section-title { font-family:'Montserrat', sans-serif; font-size:10px; letter-spacing:2px; font-weight:700; margin-bottom:15px; text-transform:uppercase; }
    
    /* Tabla simple con CSS inline para soporte en Gmail */
    .finance-table { width:100%; border-collapse:collapse; margin-bottom:10px; font-size:13px; }
    .row-item { border-bottom:1px solid #EAE8E0; }
    .col-name { padding:8px 0; text-align:left; color:#111; width:60%; }
    .col-date { padding:8px 0; text-align:center; color:#888; font-size:11px; width:20%; }
    .col-amount { padding:8px 0; text-align:right; font-family:'Montserrat', sans-serif; font-weight:700; color:#B05E52; width:20%; }
    
    .total-row { text-align:right; font-family:'Montserrat', sans-serif; font-size:11px; font-weight:700; margin-bottom:20px; color:#111; letter-spacing:1px; }
    .separator { height:1px; background-color:transparent; margin:30px 0; border-top:1px dashed #CCC; }
    .clean-state { font-style:italic; color:#888; font-size:13px; margin-bottom:20px; }

    /* FOOTER */
    .footer { background-color:#F7F5F2; border-top:1px solid #EAE8E0; padding:40px 20px; text-align:center; }
    .footer-copy { font-family:'Montserrat', sans-serif; font-size:9px; color:#BBB; letter-spacing:1px; margin:0; }
    
    /* FIRMA L1602 */
    .secret-code { display:block; font-family:'Montserrat', sans-serif; font-size:8px; letter-spacing:4px; margin-top:14px; opacity:0.18; color:#CFCAC2; user-select:none; }
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
          <h1 class="chapter-main">ESTADO FINANCIERO</h1>
          <span class="chapter-sub">Semana: ${rangeStr}</span>
        </div>

        ${contentHtml}
        
      </div>

      <div class="footer">
        <p class="footer-copy">&copy; 2025 AFM Internal System</p>
        <span class="secret-code">L1602</span>
      </div>
    </div>
  </div>
</body>
</html>
  `.trim();
}

/**
 * Helper interno para crear las filas de la tabla HTML
 */
function buildFinanceTable_(list) {
  let rows = "";
  // Formato interno de fecha
  const fmt = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM");
  
  list.forEach(item => {
    rows += `
      <tr class="row-item">
        <td class="col-name">${item.name}</td>
        <td class="col-date">${fmt(item.dateObj)}</td>
        <td class="col-amount">Q${item.amount}</td>
      </tr>
    `;
  });
  
  return `<table class="finance-table">${rows}</table>`;
}