/**
 * @fileoverview Archivo principal de ECD GESTIÓN OS.
 * Contiene la función maestra que orquesta la creación y configuración del sistema completo.
 */

/**
 * Función maestra para crear o regenerar el sistema ECD GESTIÓN OS completo.
 * Borra las hojas existentes, crea las nuevas, aplica estilos y configura todo.
 */
function runFullSystemBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = Object.values(SHEET_NAMES);

  // Confirmación del usuario para evitar borrado accidental
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmación', '¿Deseas generar el sistema ECD GESTIÓN OS? Esto borrará las hojas existentes con los mismos nombres.', ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert('Operación cancelada.');
    return;
  }

  ss.toast('Iniciando construcción del sistema...', 'ECD GESTIÓN OS', -1);

  // 1. Borrar hojas existentes para una regeneración limpia
  const existingSheets = ss.getSheets().map(s => s.getName());
  allSheetNames.forEach(name => {
    if (existingSheets.includes(name)) {
      ss.deleteSheet(ss.getSheetByName(name));
    }
  });

  // 2. Crear todas las hojas en el orden correcto
  allSheetNames.forEach(name => {
    ss.insertSheet(name);
  });

  // Eliminar la hoja "Sheet1" inicial si existe
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  // 3. Aplicar limpieza y estilos base a cada hoja
  allSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      cleanSheet(sheet);
    }
  });

  // 4. Llamar al constructor principal que maneja todas las hojas
  buildAllSheets(ss);

  // 5. Proteger rangos críticos (ejemplo: fórmulas)
  allSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(p => p.remove()); // Limpiar protecciones antiguas

    // Proteger celdas con fórmulas
    const dataRange = sheet.getDataRange();
    const formulas = dataRange.getFormulas();
    for(let i=0; i < formulas.length; i++) {
      for(let j=0; j < formulas[i].length; j++) {
        if(formulas[i][j]) {
          const cell = sheet.getRange(i+1, j+1);
          const protection = cell.protect().setDescription('Celda con fórmula');
                  // Asegurarse que solo el dueño pueda editar, no otros usuarios
          const me = Session.getEffectiveUser();
          protection.addEditor(me);
          protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== me.getEmail()));
        }
      }
    }
  });

  ss.toast('Sistema construido. Aplicando toques finales...', 'ECD GESTIÓN OS', 10);

  // 6. Activar la hoja HOME al finalizar
  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.HOME));

  ss.toast('¡Sistema ECD GESTIÓN OS generado correctamente!', 'COMPLETADO', 10);
  ui.alert('¡Éxito!', 'El sistema ECD GESTIÓN OS se ha generado y configurado correctamente.', ui.ButtonSet.OK);
}
