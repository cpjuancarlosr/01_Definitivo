/**
 * @fileoverview Módulo para construir la estructura de cada hoja en ECD GESTIÓN OS.
 * Cada función es responsable de configurar una hoja específica, aplicando la estética,
 * la estructura y las fórmulas de software requeridas.
 */

// --- BUILDERS PRINCIPALES ---

function buildAllSheets(ss) {
  buildHomeDashboard(ss);
  buildTocSheet(ss);
  buildPasswordsSheet(ss);
  buildClientsSheet(ss);
  buildMonthlyObligationsSheet(ss);
  buildCalendarSheet(ss);
  buildNotesSheet(ss);
  buildLinksSheet(ss);
  buildPersonalFinanceSheet(ss);
  buildIsrAnnualSheet(ss); // <-- This is the function being fixed
  buildIvaAnnualSheet(ss);
  buildExecutiveSummarySheet(ss);
  buildConfigSheet(ss);
}


// --- BUILDERS INDIVIDUALES ---

function buildHomeDashboard(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.HOME);
  sheet.getRange('B2').setValue('ECD GESTIÓN OS').setFontSize(24).setFontWeight('bold');
  sheet.getRange('B3').setValue('Bienvenido al sistema operativo de tu negocio.').setFontSize(12);
  sheet.setColumnWidth(1, 20);
}

function buildTocSheet(ss) {
    const sheet = ss.getSheetByName(SHEET_NAMES.TOC);
    sheet.setColumnWidths(1, 1, 20);
    sheet.setColumnWidths(2, 1, 300);
    sheet.setColumnWidths(3, 1, 400);

    const sections = [
        { title: 'OPERACIÓN', items: [SHEET_NAMES.CLIENTS, SHEET_NAMES.MONTHLY_OBLIGATIONS, SHEET_NAMES.CALENDAR] },
        { title: 'ESTRATEGIA', items: [SHEET_NAMES.NOTES, SHEET_NAMES.LINKS] },
        { title: 'FINANZAS', items: [SHEET_NAMES.ISR_ANNUAL, SHEET_NAMES.IVA_ANNUAL, SHEET_NAMES.EXECUTIVE_SUMMARY, SHEET_NAMES.PERSONAL_FINANCE] },
        { title: 'SISTEMA', items: [SHEET_NAMES.PASSWORDS, SHEET_NAMES.CONFIG] }
    ];

    let row = 2;
    applyHeaderStyle(sheet.getRange(row, 2, 1, 2).setValues([['MÓDULO', 'DESCRIPCIÓN']]));
    row++;

    sections.forEach(section => {
        const range = sheet.getRange(row, 2, 1, 2);
        applySubHeaderStyle(range.merge().setValues([[section.title]]));
        row++;
        section.items.forEach(item => {
            const gid = ss.getSheetByName(item).getSheetId();
            const formula = `=HYPERLINK("#gid=${gid}"; "${item}")`;
            sheet.getRange(row, 2).setFormula(formula).setFontWeight('bold');
            row++;
        });
    });
}

function buildPasswordsSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.PASSWORDS);
  const headers = ['CLIENTE', 'SERVICIO / PLATAFORMA', 'USUARIO', 'CONTRASEÑA', 'NOTA'];
  sheet.getRange('B2').setValue('Control de Contraseñas Interno');
  applyHeaderStyle(sheet.getRange('B3:F3').setValues([headers]));
  sheet.setColumnWidths(2, 5, 180);
  applyInputStyle(sheet.getRange('B4:F100'));
}

function buildClientsSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  const headers = ['ID CLIENTE', 'NOMBRE COMERCIAL', 'RAZÓN SOCIAL', 'RFC', 'RÉGIMEN FISCAL', 'ESTATUS FISCAL', 'OBSERVACIONES'];
  sheet.getRange('B2').setValue('Base Maestra de Clientes');
  applyHeaderStyle(sheet.getRange('B3:H3').setValues([headers]));
  sheet.setColumnWidths(2, 7, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 250);
  applyInputStyle(sheet.getRange('B4:H100'));
}

function buildMonthlyObligationsSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.MONTHLY_OBLIGATIONS);
  sheet.getRange('B2').setValue('Control de Obligaciones Mensuales');
  const headers = ['CLIENTE', 'PERIODO', 'ISR', 'IVA', 'DIOT', 'IMSS', 'ISN', 'ESTATUS GENERAL'];
  applyHeaderStyle(sheet.getRange('B3:I3').setValues([headers]));
  sheet.setColumnWidths(2, 8, 120);
  sheet.setColumnWidth(2, 200);
  applyInputStyle(sheet.getRange('B4:I100'));
}

function buildCalendarSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CALENDAR);
  sheet.getRange('B2').setValue('Calendario Fiscal y Tareas');
  const headers = ['FECHA VENCIMIENTO', 'CLIENTE', 'TAREA / OBLIGACIÓN', 'ESTATUS', 'RESPONSABLE'];
  applyHeaderStyle(sheet.getRange('B3:F3').setValues([headers]));
  sheet.setColumnWidths(2, 5, 150);
  sheet.setColumnWidth(4, 250);
  applyInputStyle(sheet.getRange('B4:F100'));
}

function buildNotesSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.NOTES);
  sheet.getRange('B2').setValue('Bitácora de Notas Estratégicas');
  const headers = ['FECHA', 'CLIENTE / TEMA', 'NOTA ESTRATÉGICA'];
  applyHeaderStyle(sheet.getRange('B3:D3').setValues([headers]));
  sheet.setColumnWidths(2, 2, 120);
  sheet.setColumnWidth(4, 600);
  applyInputStyle(sheet.getRange('B4:D100'));
}

function buildLinksSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.LINKS);
  sheet.getRange('B2').setValue('Links a Portales Recurrentes');
  const headers = ['CATEGORÍA', 'NOMBRE DEL SITIO', 'URL'];
  applyHeaderStyle(sheet.getRange('B3:D3').setValues([headers]));
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 400);
  applyInputStyle(sheet.getRange('B4:D100'));
}

function buildPersonalFinanceSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.PERSONAL_FINANCE);
  sheet.getRange('B2').setValue('Control de Finanzas Personales');
  const headers = ['FECHA', 'CATEGORÍA', 'DESCRIPCIÓN', 'INGRESO', 'EGRESO', 'SALDO'];
  applyHeaderStyle(sheet.getRange('B3:G3').setValues([headers]));
  sheet.getRange('G4').setFormula('=SUM(D4-E4)');
  sheet.getRange('G5:G100').setFormula('=$G4+SUM($D$5:D5)-SUM($E$5:E5)');
  sheet.setColumnWidths(2, 7, 120);
  sheet.setColumnWidth(4, 250);
  applyInputStyle(sheet.getRange('B4:F100'));
}

function buildConfigSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  sheet.getRange('B2').setValue('Configuración del Sistema');
  const headers = [['PARÁMETRO', 'VALOR']];
  const params = [
    ['Tasa ISR PM', 0.30],
    ['Tasa IVA General', 0.16],
    ['Coeficiente de Utilidad PM', 0.15],
  ];
  applyHeaderStyle(sheet.getRange('B3:C3').setValues(headers));
  const paramRange = sheet.getRange('B4:C6');
  paramRange.setValues(params);
  paramRange.getCell(1,2).setNumberFormat('0.00%');
  paramRange.getCell(2,2).setNumberFormat('0.00%');
  paramRange.getCell(3,2).setNumberFormat('0.00%');

  applyInputStyle(sheet.getRange('C4:C6'));
  sheet.setColumnWidths(2, 2, 250);
}

function buildIsrAnnualSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ISR_ANNUAL);
  sheet.getRange('B2').setValue('Cálculo Anual de ISR (Persona Moral)');

  const headers = ['MES', 'INGRESOS NOMINALES (Mes)', 'INGRESOS ACUMULADOS', 'COEFICIENTE UTILIDAD', 'UTILIDAD FISCAL ESTIMADA', 'ISR CAUSADO (Acumulado)', 'PAGOS PROV. ANTERIORES', 'PAGO PROVISIONAL DEL MES'];
  applyHeaderStyle(sheet.getRange('B4:I4').setValues([headers]));

  const meses = [['Enero'], ['Febrero'], ['Marzo'], ['Abril'], ['Mayo'], ['Junio'], ['Julio'], ['Agosto'], ['Septiembre'], ['Octubre'], ['Noviembre'], ['Diciembre']];
  sheet.getRange('B5:B16').setValues(meses);

  // Fórmulas CORREGIDAS
  sheet.getRange('C5').setFormula('=SUM(C5)'); // Acumulado de Enero es el mismo
  sheet.getRange('C6:C16').setFormula('=C5+SUM(C$6:C6)');
  sheet.getRange('D5:D16').setFormula(`=VLOOKUP("Coeficiente de Utilidad PM", '${SHEET_NAMES.CONFIG}'!B4:C6, 2, FALSE)`);
  sheet.getRange('E5:E16').setFormula('=C5*D5');
  sheet.getRange('F5:F16').setFormula(`=E5*VLOOKUP("Tasa ISR PM", '${SHEET_NAMES.CONFIG}'!B4:C6, 2, FALSE)`);
  sheet.getRange('H5').setValue(0); // Enero no tiene pago anterior
  sheet.getRange('H6:H16').setFormula('=SUM(I$5:I5)');
  sheet.getRange('I5:I16').setFormula('=G5-H5');

  // Estilos
  sheet.setColumnWidths(2, 9, 140);
  applyInputStyle(sheet.getRange('C5:C16')); // Ingresos del mes es input
  sheet.getRange('D5:I16').setNumberFormat('"$"#,##0.00'); // Formato de moneda
  applyResultStyle(sheet.getRange('I5:I16'));
}


function buildIvaAnnualSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.IVA_ANNUAL);
  sheet.getRange('B2').setValue('Control Anual de IVA (Persona Moral)');

  const headers = ['MES', 'IVA TRASLADADO (Cobrado)', 'IVA ACREDITABLE (Pagado)', 'IVA RETENIDO', 'IVA A CARGO / FAVOR'];
  applyHeaderStyle(sheet.getRange('B4:F4').setValues([headers]));

  const meses = [['Enero'], ['Febrero'], ['Marzo'], ['Abril'], ['Mayo'], ['Junio'], ['Julio'], ['Agosto'], ['Septiembre'], ['Octubre'], ['Noviembre'], ['Diciembre']];
  sheet.getRange('B5:B16').setValues(meses);

  // Fórmula
  sheet.getRange('F5:F16').setFormula('=C5-D5-E5');

  sheet.setColumnWidths(2, 6, 160);
  applyInputStyle(sheet.getRange('C5:E16'));
  applyResultStyle(sheet.getRange('F5:F16'));
}

function buildExecutiveSummarySheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.EXECUTIVE_SUMMARY);
  sheet.getRange('B2').setValue('Resumen Ejecutivo Fiscal Anual');
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 150);

  const kpis = [
    ['ISR TOTAL ANUAL CAUSADO'],
    ['IVA NETO ANUAL (A CARGO / FAVOR)'],
    ['PAGOS PROVISIONALES TOTALES'],
    ['CARGA FISCAL TOTAL']
  ];

  sheet.getRange('B4:B7').setValues(kpis).setFontWeight('bold');

  // Fórmulas del Resumen
  sheet.getRange('C4').setFormula(`=MAX('${SHEET_NAMES.ISR_ANNUAL}'!G5:G16)`);
  sheet.getRange('C5').setFormula(`=SUM('${SHEET_NAMES.IVA_ANNUAL}'!F5:F16)`);
  sheet.getRange('C6').setFormula(`=SUM('${SHEET_NAMES.ISR_ANNUAL}'!I5:I16)`);
  sheet.getRange('C7').setFormula('=SUM(C4:C6)');

  applyResultStyle(sheet.getRange('C4:C7').setNumberFormat('"$"#,##0.00'));
  addBottomBorder(sheet, 8, 2, 3);
}
