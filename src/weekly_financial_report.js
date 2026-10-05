/****************************************************************
 * REPORTE SEMANAL DE FINANZAS (LÓGICA)
 * 1. Conecta con la DB de Pagos.
 * 2. Filtra deudas vencidas y de esta semana.
 * 3. Llama al generador de HTML (en el otro archivo).
 * 4. Envía el correo a la lista fija.
 ****************************************************************/

// --- CONFIGURACIÓN (Script Properties) ---
// Global: también lo usa generateMonthlyBills() en monthly_payments_generator.js
const ID_DB_FINANZAS = PropertiesService.getScriptProperties().getProperty("DB_FINANCE_SPREADSHEET_ID");

// LISTA DE CORREOS
// Se configura en la Script Property FINANCE_REPORT_EMAILS, separados por coma.
const FINANCE_REPORT_EMAILS = (PropertiesService.getScriptProperties().getProperty("FINANCE_REPORT_EMAILS") || "").split(",").map(e => e.trim()).filter(Boolean);

function sendWeeklyCollectionReport() {
  const ss = SpreadsheetApp.openById(assertScriptProperty_("DB_FINANCE_SPREADSHEET_ID", ID_DB_FINANZAS));
  const sheet = ss.getSheetByName("Payments");
  const data = sheet.getDataRange().getValues();

  // --- FECHAS ---
  const today = new Date();
  today.setHours(0,0,0,0);
  
  // Calcular el Domingo de esta semana
  const sunday = new Date(today);
  sunday.setDate(today.getDate() + 6);
  sunday.setHours(23,59,59,999);

  // Listas para clasificar
  let overdueList = [];  // Vencidos (Anteriores a hoy)
  let thisWeekList = []; // Esta semana (Hoy -> Domingo)

  // --- RECORRIDO DE DATOS (Saltando header) ---
  // Indices: C=Name(2), F=Total(5), G=Pagado(6), H=Status(7), I=PayDay(8)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const student = row[2];
    const total = Number(row[5]) || 0;
    const paid = Number(row[6]) || 0;
    const pendingAmount = total - paid; // Saldo Real
    const status = String(row[7]).trim();
    
    // Validar fecha
    let payDay = null;
    if (row[8] && row[8] instanceof Date) {
      payDay = row[8];
    }

    // FILTRO: Solo procesamos si hay deuda pendiente y fecha válida
    // Ignoramos "Paid" y "Cancelled"
    if (status !== "Paid" && status !== "Cancelled" && pendingAmount > 0 && payDay) {
      
      const item = { 
        name: student, 
        amount: pendingAmount, 
        dateObj: payDay 
      };

      if (payDay < today) {
        overdueList.push(item);
      } 
      else if (payDay >= today && payDay <= sunday) {
        thisWeekList.push(item);
      }
    }
  }

  // Si no hay nada que cobrar, no enviamos nada (o podrías descomentar para enviar aviso de "Todo al día")
  if (overdueList.length === 0 && thisWeekList.length === 0) {
    Logger.log("Nada que reportar hoy.");
    return; 
  }

  // --- GENERACIÓN DEL HTML (Llamada al otro archivo) ---
  // Le pasamos las listas limpias y las fechas para el título
  const htmlBody = buildFinancialReportEmailHtml_(overdueList, thisWeekList, today, sunday);

  // --- ENVÍO ---
  const recipients = assertScriptProperty_("FINANCE_REPORT_EMAILS", FINANCE_REPORT_EMAILS.join(","));
  
  MailApp.sendEmail({
    to: recipients,
    subject: "📜 Reporte Semanal | Finanzas AFM",
    htmlBody: htmlBody,
    name: "Finanzas AFM"
  });
  
  Logger.log(`Reporte enviado a: ${recipients}`);
}