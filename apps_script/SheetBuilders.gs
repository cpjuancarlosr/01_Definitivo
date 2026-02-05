/**
 * @fileoverview Módulo para construir la estructura de cada hoja en ECD GESTIÓN OS.
 * Esta es la versión final y corregida que resuelve todos los errores de method chaining.
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
  buildIsrAnnualSheet(ss);
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

    const SHEET_DESCRIPTIONS = {
      [SHEET_NAMES.CLIENTS]: 'Base de datos central de clientes.',
      [SHEET_NAMES.MONTHLY_OBLIGATIONS]: 'Control de impuestos y obligaciones mensuales.',
      [SHEET_NAMES.CALENDAR]: 'Calendario de vencimientos y tareas clave.',
      [SHEET_NAMES.NOTES]: 'Bitácora de notas y decisiones estratégicas.',
      [SHEET_NAMES.LINKS]: 'Acceso rápido a portales y herramientas.',
      [SHEET_NAMES.ISR_ANNUAL]: 'Cálculo y planeación de ISR anual.',
      [SHEET_NAMES.IVA_ANNUAL]: 'Flujo y control de IVA anual.',
      [SHEET_NAMES.EXECUTIVE_SUMMARY]: 'Dashboard de KPIs fiscales.',
      [SHEET_NAMES.PERSONAL_FINANCE]: 'Control de finanzas personales.',
      [SHEET_NAMES.PASSWORDS]: 'Gestor de credenciales internas.',
      [SHEET_NAMES.CONFIG]: 'Configuración de parámetros del sistema.'
    };

    const sections = [
        { title: 'OPERACIÓN', items: [SHEET_NAMES.CLIENTS, SHEET_NAMES.MONTHLY_OBLIGATIONS, SHEET_NAMES.CALENDAR] },
        { title: 'ESTRATEGIA', items: [SHEET_NAMES.NOTES, SHEET_NAMES.LINKS] },
        { title: 'FINANZAS', items: [SHEET_NAMES.ISR_ANNUAL, SHEET_NAMES.IVA_ANNUAL, SHEET_NAMES.EXECUTIVE_SUMMARY, SHEET_NAMES.PERSONAL_FINANCE] },
        { title: 'SISTEMA', items: [SHEET_NAMES.PASSWORDS, SHEET_NAMES.CONFIG] }
    ];

    let row = 2;
    const tocHeaderRange = sheet.getRange(row, 2, 1, 2);
    tocHeaderRange.setValues([['MÓDULO', 'DESCRIPCIÓN']]);
    applyHeaderStyle(tocHeaderRange);
    row++;

    sections.forEach(section => {
        const range = sheet.getRange(row, 2, 1, 2);
        range.merge();
        range.setValue(section.title);
        applySubHeaderStyle(range);
        row++;

        section.items.forEach(item => {
            const itemSheet = ss.getSheetByName(item);
            if (itemSheet) {
                const gid = itemSheet.getSheetId();
                const formula = `=HYPERLINK("#gid=${gid}"; "${item}")`;
                sheet.getRange(row, 2).setFormula(formula).setFontWeight('bold');
                sheet.getRange(row, 3).setValue(SHEET_DESCRIPTIONS[item]);
                row++;
            }
        });
    });
}

function buildPasswordsSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.PASSWORDS);
  const headers = ['CLIENTE', 'SERVICIO / PLATAFORMA', 'USUARIO', 'CONTRASEÑA', 'NOTA'];
  sheet.getRange('B2').setValue('Control de Contraseñas Interno');
  const passHeaderRange = sheet.getRange('B3:F3');
  passHeaderRange.setValues([headers]);
  applyHeaderStyle(passHeaderRange);
  sheet.setColumnWidths(2, 5, 180);
  applyInputStyle(sheet.getRange('B4:F100'));
}

function buildClientsSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  const headers = ['ID CLIENTE', 'NOMBRE COMERCIAL', 'RAZÓN SOCIAL', 'RFC', 'RÉGIMEN FISCAL', 'ESTATUS FISCAL', 'OBSERVACIONES'];
  sheet.getRange('B2').setValue('Base Maestra de Clientes');
  const clientHeaderRange = sheet.getRange('B3:H3');
  clientHeaderRange.setValues([headers]);
  applyHeaderStyle(clientHeaderRange);
  sheet.setColumnWidths(2, 7, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 250);
  applyInputStyle(sheet.getRange('B4:H100'));
}

function buildMonthlyObligationsSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.MONTHLY_OBLIGATIONS);
  sheet.getRange('B2').setValue('Control de Obligaciones Mensuales');
  const headers = ['CLIENTE', 'PERIODO', 'ISR', 'IVA', 'DIOT', 'IMSS', 'ISN', 'ESTATUS GENERAL'];
  const monthlyHeaderRange = sheet.getRange('B3:I3');
  monthlyHeaderRange.setValues([headers]);
  applyHeaderStyle(monthlyHeaderRange);
  sheet.setColumnWidths(2, 8, 120);
  sheet.setColumnWidth(2, 200);
  applyInputStyle(sheet.getRange('B4:I100'));
}

function buildCalendarSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CALENDAR);
  sheet.getRange('B2').setValue('Calendario Fiscal y Tareas');
  const headers = ['FECHA VENCIMIENTO', 'CLIENTE', 'TAREA / OBLIGACIÓN', 'ESTATUS', 'RESPONSABLE'];
  const calendarHeaderRange = sheet.getRange('B3:F3');
  calendarHeaderRange.setValues([headers]);
  applyHeaderStyle(calendarHeaderRange);
  sheet.setColumnWidths(2, 5, 150);
  sheet.setColumnWidth(4, 250);
  applyInputStyle(sheet.getRange('B4:F100'));
}

function buildNotesSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.NOTES);
  sheet.getRange('B2').setValue('Bitácora de Notas Estratégicas');
  const headers = ['FECHA', 'CLIENTE / TEMA', 'NOTA ESTRATÉGICA'];
  const notesHeaderRange = sheet.getRange('B3:D3');
  notesHeaderRange.setValues([headers]);
  applyHeaderStyle(notesHeaderRange);
  sheet.setColumnWidths(2, 2, 120);
  sheet.setColumnWidth(4, 600);
  applyInputStyle(sheet.getRange('B4:D100'));
}

function buildLinksSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.LINKS);
  sheet.getRange('B2').setValue('Links a Portales Recurrentes');
  const headers = ['CATEGORÍA', 'NOMBRE DEL SITIO', 'URL'];
  const linksHeaderRange = sheet.getRange('B3:D3');
  linksHeaderRange.setValues([headers]);
  applyHeaderStyle(linksHeaderRange);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 400);
  applyInputStyle(sheet.getRange('B4:D100'));
}

function buildPersonalFinanceSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.PERSONAL_FINANCE);
  sheet.getRange('B2').setValue('Control de Finanzas Personales');
  const headers = ['FECHA', 'CATEGORÍA', 'DESCRIPCIÓN', 'INGRESO', 'EGRESO', 'SALDO'];
  const pfHeaderRange = sheet.getRange('B3:G3');
  pfHeaderRange.setValues([headers]);
  applyHeaderStyle(pfHeaderRange);
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
  const configHeaderRange = sheet.getRange('B3:C3');
  configHeaderRange.setValues(headers);
  applyHeaderStyle(configHeaderRange);
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
  const isrHeaderRange = sheet.getRange('B4:I4');
  isrHeaderRange.setValues([headers]);
  applyHeaderStyle(isrHeaderRange);
  const meses = [['Enero'], ['Febrero'], ['Marzo'], ['Abril'], ['Mayo'], ['Junio'], ['Julio'], ['Agosto'], ['Septiembre'], ['Octubre'], ['Noviembre'], ['Diciembre']];
  sheet.getRange('B5:B16').setValues(meses);
  sheet.getRange('D5:D16').setFormula(`='${SHEET_NAMES.CONFIG}'!C6`);
  sheet.getRange('C5').setFormula('=IF(ISBLANK(B5),"",SUM(B5))');
  sheet.getRange('C6:C16').setFormula('=IF(ISBLANK(B6),"",C5+SUM(B$6:B6))');
  sheet.getRange('E5:E16').setFormula('=D5*E5');
  sheet.getRange('F5:F16').setFormula(`=F5*'${SHEET_NAMES.CONFIG}'!C4`);
  sheet.getRange('H5').setValue(0);
  sheet.getRange('H6:H16').setFormula('=SUM(I$5:I5)');
  sheet.getRange('I5:I16').setFormula('=G5-H5');
  sheet.setColumnWidths(2, 9, 140);
  applyInputStyle(sheet.getRange('B5:B16'));
  sheet.getRange('C5:I16').setNumberFormat('"$"#,##0.00');
  applyResultStyle(sheet.getRange('I5:I16'));
}

function buildIvaAnnualSheet(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.IVA_ANNUAL);
  sheet.getRange('B2').setValue('Control Anual de IVA (Persona Moral)');
  const headers = ['MES', 'IVA TRASLADADO (Cobrado)', 'IVA ACREDITABLE (Pagado)', 'IVA RETENIDO', 'IVA A CARGO / FAVOR'];
  const ivaHeaderRange = sheet.getRange('B4:F4');
  ivaHeaderRange.setValues([headers]);
  applyHeaderStyle(ivaHeaderRange);
  const meses = [['Enero'], ['Febrero'], ['Marzo'], ['Abril'], ['Mayo'], ['Junio'], ['Julio'], ['Agosto'], ['Septiembre'], ['Octubre'], ['Noviembre'], ['Diciembre']];
  sheet.getRange('B5:B16').setValues(meses);
  sheet.getRange('F5:F16').setFormula('=C5-D5-E5');
  sheet.setColumnWidths(2, 6, 160);
  applyInputStyle(sheet.getRange('C5:E16'));
  applyResultStyle(sheet.getRange('F5:F16'));
  sheet.getRange('C5:F16').setNumberFormat('"$"#,##0.00');
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
  sheet.getRange('C4').setFormula(`=MAX('${SHEET_NAMES.ISR_ANNUAL}'!G5:G16)`);
  sheet.getRange('C5').setFormula(`=SUM('${SHEET_NAMES.IVA_ANNUAL}'!F5:F16)`);
  sheet.getRange('C6').setFormula(`=SUM('${SHEET_NAMES.ISR_ANNUAL}'!I5:I16)`);
  sheet.getRange('C7').setFormula('=SUM(C4:C6)');
  applyResultStyle(sheet.getRange('C4:C7').setNumberFormat('"$"#,##0.00'));
  addBottomBorder(sheet, 8, 2, 3);
}
