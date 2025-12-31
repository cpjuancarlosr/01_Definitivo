/**
 * @fileoverview Archivo principal de ECD GESTIÓN OS.
 * Contiene la función maestra y el objeto global de la hoja de cálculo.
 */

// GLOBAL SPREADSHEET OBJECT - accessible by all script files
const ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Función maestra para crear o regenerar el sistema ECD GESTIÓN OS completo.
 * Borra las hojas existentes, crea las nuevas, aplica estilos y configura todo.
 */
function runFullSystemBuild() {
  const allSheetNames = Object.values(SHEET_NAMES);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmación', '¿Deseas generar el sistema ECD GESTIÓN OS? Esto borrará las hojas existentes con los mismos nombres.', ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert('Operación cancelada.');
    return;
  }

  ss.toast('Iniciando construcción del sistema...', 'ECD GESTIÓN OS', -1);

  const tempSheet = ss.insertSheet('temp_placeholder_' + new Date().getTime());

  const existingSheets = ss.getSheets();
  existingSheets.forEach(sheet => {
    if (allSheetNames.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
    }
  });

  allSheetNames.forEach(name => {
    ss.insertSheet(name);
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  allSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      cleanSheet(sheet);
    }
  });

  buildAllSheets();

  allSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(p => p.remove());

    const dataRange = sheet.getDataRange();
    const formulas = dataRange.getFormulas();
    for(let i=0; i < formulas.length; i++) {
      for(let j=0; j < formulas[i].length; j++) {
        if(formulas[i][j]) {
          const cell = sheet.getRange(i+1, j+1);
          cell.protect().setDescription('Celda con fórmula');
        }
      }
    }
  });

  ss.toast('Sistema construido. Aplicando toques finales...', 'ECD GESTIÓN OS', 5);

  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.HOME));
  ss.deleteSheet(tempSheet);

  ss.toast('¡Sistema ECD GESTIÓN OS generado correctamente!', 'COMPLETADO', 5);
  ui.alert('¡Éxito!', 'El sistema ECD GESTIÓN OS se ha generado y configurado correctamente.', ui.ButtonSet.OK);
}
