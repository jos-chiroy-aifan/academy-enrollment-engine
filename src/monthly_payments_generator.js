/****************************************************************
 * GENERADOR DE DEUDA MENSUAL (AGRUPADO + SIN NOTAS)
 * 1. Unifica todas las clases de un alumno en un solo cobro.
 * 2. Incluye columna para Abonos (AmountPaid).
 * 3. SIN columna de notas.
 ****************************************************************/

function generateMonthlyBills() {
  // --- IDs (Script Properties) ---
  // ID_DB_FINANZAS es global (definido en weekly_financial_report.js)
  const ID_DB_GENERAL = PropertiesService.getScriptProperties().getProperty("DB_GENERAL_SPREADSHEET_ID");

  const ssGeneral = SpreadsheetApp.openById(assertScriptProperty_("DB_GENERAL_SPREADSHEET_ID", ID_DB_GENERAL));
  const ssFinance = SpreadsheetApp.openById(assertScriptProperty_("DB_FINANCE_SPREADSHEET_ID", ID_DB_FINANZAS));
  
  const sheetEnr = ssGeneral.getSheetByName("Student_Enrollments");
  const sheetPay = ssFinance.getSheetByName("Payments");

  // 1. OBTENER DATOS
  const data = sheetEnr.getDataRange().getValues();
  const studentsBill = {}; // Objeto para agrupar
  
  // 2. PREPARAR FECHAS
  const now = new Date();
  const currentYear = now.getFullYear();
  const yearShort = String(currentYear).slice(-2); // "26"
  const monthNum = String(now.getMonth() + 1).padStart(2, '0'); // "02"
  
  const monthsEs = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const currentMonthTxt = monthsEs[now.getMonth()]; 
  const lastDayOfThisMonth = new Date(currentYear, now.getMonth() + 1, 0).getDate();

  // 3. AGRUPAR DEUDA POR ALUMNO
  // Indices: B=StudentID(1), C=Name(2), K=State(10), O=Price(14), P=PayDay_Original(15)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const stId = row[1];
    const stName = row[2];
    const state = row[10];
    const price = Number(row[14]) || 0;
    
    // Solo procesar activos
    if (String(state).trim() === "Active") {
      
      // Iniciar objeto del alumno si no existe
      if (!studentsBill[stId]) {
        // Calcular día de corte (del primer registro activo que encontremos)
        let billingDay = 1;
        if (row[15] && row[15] instanceof Date) {
          billingDay = row[15].getDate();
        }

        studentsBill[stId] = {
          name: stName,
          totalAmount: 0,
          billingDay: billingDay
        };
      }

      // Sumar al total
      studentsBill[stId].totalAmount += price;
    }
  }

  // 4. GENERAR FILAS EN PAGOS
  const newRows = [];

  for (const stId in studentsBill) {
    const studentData = studentsBill[stId];
    
    // --- ID ÚNICO POR ALUMNO (P2602-101) ---
    // Extraemos solo los números del ID del estudiante
    const suffix = stId.replace(/[^0-9]/g, '');
    const billId = `P${yearShort}${monthNum}-${suffix}`;

    // --- PROTECCIÓN: ¿YA EXISTE? ---
    if (!paymentExists_(sheetPay, billId)) {
      
      // Ajuste Fin de Mes
      let finalDay = studentData.billingDay;
      if (finalDay > lastDayOfThisMonth) finalDay = lastDayOfThisMonth;
      const specificDueDate = new Date(currentYear, now.getMonth(), finalDay);

      newRows.push([
        billId,                     // A: PaymentID
        stId,                       // B: StudentID
        studentData.name,           // C: StudentName
        currentMonthTxt,            // D: BillMonth
        currentYear,                // E: BillYear
        studentData.totalAmount,    // F: Amount (DEUDA TOTAL)
        0,                          // G: AmountPaid (ABONO INICIAL 0)
        "Pending",                  // H: Status
        specificDueDate,            // I: PayDay
        "",                         // J: PaidDate
        ""                          // K: Method
      ]);
    }
  }

  // 5. GUARDAR
  if (newRows.length > 0) {
    const lastRow = sheetPay.getLastRow();
    sheetPay.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    Logger.log(`✅ Se generaron ${newRows.length} cobros unificados.`);
  } else {
    Logger.log("✅ Todo al día.");
  }
}

function paymentExists_(sheet, idToCheck) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === idToCheck) return true;
  }
  return false;
}