/************************************
 * programs_form_sync.gs
 * Sincronización de Programas Activos
 ************************************/

/**
 * CONFIGURACIÓN
 */
const PROGRAMS_SHEET_NAME = 'Programs';

const PROGRAMS_COL = {
  program_name: 2, // Columna B
  state: 5         // Columna E
};

const PROFESSORS_FORM = {
  form_id: PropertiesService.getScriptProperties().getProperty("UNIFIED_FORM_ID"),
  question_title: 'Programa' 
};

/**
 * Sincroniza programas ACTIVE hacia el Formulario Unificado
 * Configura la navegación automática según el nombre del programa
 */
function sync_programs_to_professors_form() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PROGRAMS_SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`No existe la hoja: ${PROGRAMS_SHEET_NAME}`);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  // 1. OBTENER PROGRAMAS ACTIVOS DESDE LA DB
  const activePrograms = [];
  for (let i = 1; i < data.length; i++) {
    const programName = data[i][PROGRAMS_COL.program_name - 1];
    const state = data[i][PROGRAMS_COL.state - 1];

    if (state === 'Active' && programName) {
      activePrograms.push(String(programName).trim());
    }
  }

  activePrograms.sort((a, b) => a.localeCompare(b));

  // 2. ABRIR FORMULARIO Y MAPEARY SECCIONES
  const form = FormApp.openById(assertScriptProperty_("UNIFIED_FORM_ID", PROFESSORS_FORM.form_id));
  
  // Indexar todas las secciones (PageBreaks) para habilitar navegación
  const pages = form.getItems(FormApp.ItemType.PAGE_BREAK);
  const pageMap = {};
  pages.forEach(p => {
    pageMap[p.getTitle()] = p.asPageBreakItem();
  });

  const items = form.getItems();
  let found = false;

  for (const item of items) {
    if (item.getTitle() === PROFESSORS_FORM.question_title) {
      const type = item.getType();
      const listItem = item.asListItem();

      // 3. CONSTRUCCIÓN DE OPCIONES CON NAVEGACIÓN DINÁMICA
      const choices = activePrograms.map(p => {
        // Formato esperado de sección: "Horario del Programa - [Nombre]"
        const sectionTitle = `Horario del Programa - ${p}`; 
        const targetPage = pageMap[sectionTitle];
        
        if (targetPage) {
          // Asigna la ruta de navegación a la sección correspondiente
          return listItem.createChoice(p, targetPage);
        } else {
          Logger.log(`⚠️ No se encontró sección para: ${p}`);
          return listItem.createChoice(p);
        }
      });

      // 4. ACTUALIZAR PREGUNTA (LIST O MULTIPLE_CHOICE)
      if (type === FormApp.ItemType.LIST) {
        listItem.setChoices(choices);
        found = true;
      } 
      else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        item.asMultipleChoiceItem().setChoices(choices);
        found = true;
      }
      
      if (found) break;
    }
  }

  if (!found) {
    throw new Error(`No se encontró la pregunta "${PROFESSORS_FORM.question_title}".`);
  }
  
  Logger.log("Sincronización exitosa.");
}