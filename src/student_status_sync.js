/****************************************************************
 * SINCRONIZACIÓN DE ESTADO (ESTUDIANTE -> INSCRIPCIONES)
 * Nombre: syncStudentStatusOnEdit
 * Acción: Al cambiar el estado en "Students", actualiza "Student_Enrollments".
 ****************************************************************/

function syncStudentStatusOnEdit(e) {
  // Verificación de seguridad por si e viene nulo (si lo corres manual)
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  // 1. FILTRO: Solo actuamos si se edita la hoja "Students"
  if (sheet.getName() === "Students") {
    
    // 2. FILTRO: Solo actuamos si se edita la columna "State" (Columna S = Índice 19)
    if (range.getColumn() === 15) {
      
      const newStatus = e.value; // El valor nuevo (ej: "Inactive")
      const row = range.getRow();
      
      // Evitar cabeceras
      if (row < 2) return;

      // Obtenemos el ID del alumno de esa fila (Columna A = Índice 1)
      const studentId = sheet.getRange(row, 1).getValue();

      // Si tenemos un ID válido y un estado, ejecutamos la sincronización
      if (studentId && newStatus) {
        updateEnrollmentStatus_(studentId, newStatus);
      }
    }
  }
}

/**
 * Función auxiliar que recorre las inscripciones y actualiza el estado.
 * Busca el StudentID en la Columna B y actualiza la Columna K.
 */
function updateEnrollmentStatus_(studentId, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEnr = ss.getSheetByName("Student_Enrollments");
  
  if (!sheetEnr) return;

  const data = sheetEnr.getDataRange().getValues();
  let count = 0;
  
  // Recorrer filas (Empezamos en 1 para saltar encabezado)
  for (let i = 1; i < data.length; i++) {
    // Columna B (StudentID) es el índice [1] en el array (0, 1, 2...)
    const rowId = String(data[i][1]); 
    
    // Si el ID coincide
    if (rowId === String(studentId)) {
      
      // Verificamos si el estado actual (Columna K, índice [10]) es diferente
      if (String(data[i][10]) !== status) {
        
        // Actualizamos la Columna K (que es la columna 11 en la hoja)
        // i + 1 porque el array empieza en 0 pero las filas de hoja en 1
        sheetEnr.getRange(i + 1, 11).setValue(status);
        count++;
      }
    }
  }
  
  // Notificación visual (opcional)
  if (count > 0) {
    // Usamos try-catch por si el trigger 'on edit' no tiene permisos de UI
    try {
      ss.toast(`🔄 Se actualizaron ${count} clases a "${status}".`);
    } catch(e) {}
  }
}