/**
 * @fileoverview Crea y gestiona el menú personalizado para ECD GESTIÓN OS.
 * Se asegura de que el usuario pueda (re)generar el sistema fácilmente.
 */

/**
 * Se ejecuta cuando la hoja de cálculo se abre.
 * Agrega el menú "ECD OS" a la UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_OPTIONS.mainTitle)
    .addItem(MENU_OPTIONS.regenerate, 'runFullSystemBuild') // Cambiado para llamar al orquestador principal
    .addToUi();
}
