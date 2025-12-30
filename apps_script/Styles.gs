/**
 * @fileoverview Motor de Estilos y UI/UX para ECD GESTIÓN OS.
 * Aplica la estética "Industrial Modern" a las hojas de cálculo, garantizando
 * una apariencia coherente y profesional tipo software.
 */

/**
 * Aplica un formato de encabezado de sección a un rango.
 * Fondo negro, texto blanco, negrita, centrado.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function applyHeaderStyle(range) {
  range
    .setBackground(PALETTE.BLACK)
    .setFontColor(PALETTE.WHITE)
    .setFontFamily(FONT_FAMILY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

/**
 * Aplica un formato de sub-encabezado a un rango.
 * Fondo gris de línea, texto negro, negrita.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function applySubHeaderStyle(range) {
  range
    .setBackground(PALETTE.LINE_GRAY)
    .setFontColor(PALETTE.BLACK)
    .setFontFamily(FONT_FAMILY)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
}

/**
 * Aplica un formato de celda de input a un rango.
 * Fondo gris claro, alineación izquierda.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function applyInputStyle(range) {
  range
    .setBackground(PALETTE.INPUT_GRAY)
    .setFontFamily(FONT_FAMILY)
    .setHorizontalAlignment('left');
}

/**
 * Aplica un formato de celda de resultado o KPI a un rango.
 * Texto color acento, negrita.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function applyResultStyle(range) {
  range
    .setFontColor(PALETTE.ACCENT)
    .setFontFamily(FONT_FAMILY)
    .setFontWeight('bold');
}

/**
 * Limpia el formato de una hoja completa y aplica una base limpia.
 * Elimina todas las cuadrículas, establece el fondo blanco y la fuente por defecto.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a limpiar.
 */
function cleanSheet(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRows, maxCols)
    .setBackground(PALETTE.WHITE)
    .setFontColor(PALETTE.BLACK)
    .setFontFamily(FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight('normal')
    .setWrap(false)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
  sheet.setGridLines(false);
}

/**
 * Crea una línea separadora debajo de una fila específica.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja donde se creará la línea.
 * @param {number} row El número de la fila debajo de la cual se dibujará la línea.
 * @param {number=} optStartCol Columna inicial (opcional).
 * @param {number=} optEndCol Columna final (opcional).
 */
function addBottomBorder(sheet, row, optStartCol, optEndCol) {
  const startCol = optStartCol || 1;
  const numCols = (optEndCol || sheet.getMaxColumns()) - startCol + 1;
  const range = sheet.getRange(row, startCol, 1, numCols);
  range.setBorder(null, null, true, null, null, null, PALETTE.LINE_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
