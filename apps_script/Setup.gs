/**
 * @fileoverview Contiene las constantes y configuraciones globales del sistema ECD GESTIÓN OS.
 * Este archivo centraliza todos los parámetros clave para facilitar el mantenimiento y la escalabilidad.
 */

// Paleta de colores "Industrial Modern"
const PALETTE = {
  BLACK: '#0F0F0F',
  WHITE: '#FFFFFF',
  INPUT_GRAY: '#F5F5F5',
  LINE_GRAY: '#DADADA',
  ACCENT: '#7DAA00'
};

// Nombres de las hojas de cálculo
const SHEET_NAMES = {
  HOME: '00_HOME_Dashboard',
  TOC: '01_Tabla_de_Contenidos',
  PASSWORDS: '02_Control_Contraseñas',
  CLIENTS: '03_Clientes',
  MONTHLY_OBLIGATIONS: '04_Obligaciones_Mensuales',
  CALENDAR: '05_Calendario_Fiscal_y_Tareas',
  NOTES: '06_Notas_Estrategicas',
  LINKS: '07_Links_Recurrentes',
  PERSONAL_FINANCE: '08_Finanzas_Personales',
  ISR_ANNUAL: '09_ISR_PM_Anual',
  IVA_ANNUAL: '10_IVA_PM_Anual',
  EXECUTIVE_SUMMARY: '11_Resumen_Ejecutivo_Fiscal',
  CONFIG: '99_Config_Sistema'
};

// Tipografía estándar
const FONT_FAMILY = 'Arial';

// Opciones del menú personalizado
const MENU_OPTIONS = {
  mainTitle: 'ECD OS',
  regenerate: 'Regenerar Sistema Completo',
  // Futuras opciones:
  // clientMode: 'Activar Modo Cliente',
  // monthlyReset: 'Ejecutar Cierre Mensual'
};
