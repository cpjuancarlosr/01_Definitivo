function onOpen(){
  SpreadsheetApp.getUi().createMenu('Jefatura Contable')
    .addItem('Cargar XML…','JC_showPicker')
    .addItem('Generar Previa','JC_buildPreview')
    .addItem('Emitir Pólizas','JC_emitirPolizas')
    .addSeparator()
    .addItem('Importar Banco (CSV/PDF)','JC_importBanco')
    .addItem('Conciliación','JC_runConciliacion')
    .addSeparator()
    .addItem('Refrescar Reporte','JC_refresh')
    .addToUi();
}

function JC_parseFilesById_(ids){
  const sh = ss_('CFDI');
  const rows = [];
  ids.forEach(id=>{
    const file = DriveApp.getFileById(id);
    if (file.getMimeType().indexOf('xml')===-1) return;
    const xml = XmlService.parse(file.getBlob().getDataAsString('UTF-8'));
    const d = JC_extractCFDI_(xml.getRootElement());
    d.ArchivoXML_ID = id;
    rows.push(JC_toRow_(d));
  });
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

function JC_extractCFDI_(root){
  const ns = {
    cfdi: XmlService.getNamespace('cfdi','http://www.sat.gob.mx/cfd/4'),
    tfd:  XmlService.getNamespace('tfd','http://www.sat.gob.mx/TimbreFiscalDigital'),
    p20:  XmlService.getNamespace('pago20','http://www.sat.gob.mx/Pagos20')
  };
  const g = (el,att)=> el?.getAttribute(att)?.getValue()||'';
  const comp = root;
  const em = comp.getChild('Emisor',ns.cfdi); const re = comp.getChild('Receptor',ns.cfdi);
  const tim = comp.getChild('Complemento',ns.cfdi)?.getChild('TimbreFiscalDigital',ns.tfd);
  const conceptos = comp.getChild('Conceptos',ns.cfdi)?.getChildren('Concepto',ns.cfdi)||[];
  const c0 = conceptos[0];
  const data = {
    Tipo: g(comp,'TipoDeComprobante'),
    UUID: tim? g(tim,'UUID'): '',
    Serie: g(comp,'Serie'), Folio: g(comp,'Folio'), Fecha: g(comp,'Fecha'),
    RFC_Emisor: g(em,'Rfc'), Nombre_Emisor: g(em,'Nombre'),
    RFC_Receptor: g(re,'Rfc'), Nombre_Receptor: g(re,'Nombre'),
    UsoCFDI: g(re,'UsoCFDI'), Método: g(comp,'MetodoPago'), Forma_Pago: g(comp,'FormaPago'),
    Moneda: g(comp,'Moneda')||'MXN', Tipo_Cambio: g(comp,'TipoCambio')||'',
    Subtotal: +g(comp,'SubTotal')||0, Descuento: +g(comp,'Descuento')||0, Total: +g(comp,'Total')||0,
    Concepto_Principal: c0? (g(c0,'ClaveProdServ')+': '+(g(c0,'Descripcion')||'')) : ''
  };
  // Impuestos compactos
  const imp = comp.getChild('Impuestos',ns.cfdi);
  if (imp){
    const tras = imp.getChild('Traslados',ns.cfdi)?.getChildren('Traslado',ns.cfdi)||[];
    tras.forEach(t=>{
      const tasa = +g(t,'TasaOCuota')||0; const impi = +g(t,'Importe')||0;
      if (Math.abs(tasa-0.16)<1e-6) data.IVA_16 = (data.IVA_16||0)+impi;
      else if (Math.abs(tasa-0.08)<1e-6) data.IVA_08 = (data.IVA_08||0)+impi;
      else data.IVA_00 = (data.IVA_00||0)+impi;
    });
    const rets = imp.getChild('Retenciones',ns.cfdi)?.getChildren('Retencion',ns.cfdi)||[];
    rets.forEach(r=>{ const im = g(r,'Impuesto'); const impi = +g(r,'Importe')||0;
      if (im==='001') data.Ret_ISR = (data.Ret_ISR||0)+impi;
      if (im==='002') data.Ret_IVA = (data.Ret_IVA||0)+impi; });
  }
  // Complemento de pagos 2.0 → marcar Tipo=P si aplica (detalle a futuro)
  const pagos = comp.getChild('Complemento',ns.cfdi)?.getChild('Pagos',ns.p20);
  if (pagos) data.Tipo = 'P';
  return data;
}

function JC_toRow_(d){
  return [d.Tipo,d.UUID,d.Fecha, d.Fecha? Utilities.formatDate(new Date(d.Fecha), Session.getScriptTimeZone(),'yyyy-mm'):'', d.Serie,d.Folio,
    d.RFC_Emisor,d.Nombre_Emisor,d.RFC_Receptor,d.Nombre_Receptor,
    d.UsoCFDI,d.Método,d.Forma_Pago,d.Moneda,d.Tipo_Cambio,
    d.Concepto_Principal,
    d.Subtotal,d.Descuento,d.IVA_16||0,d.IVA_08||0,d.IVA_00||0,d.Ret_ISR||0,d.Ret_IVA||0,d.IEPS||0,d.Total,
    '', '', // UUID_Relacionado, Tipo_Relación (se llenan al ligar NC/Pagos)
    '', '', d.ArchivoXML_ID, // Link_XML, Link_PDF, ArchivoXML_ID
    'No','No','Sí' // Con_Póliza, Conciliada, Incluida_Reportes
  ];
}

function ss_(n){ return SpreadsheetApp.getActive().getSheetByName(n); }

function JC_showPicker() {
  const html = HtmlService.createHtmlOutputFromFile('Picker')
    .setWidth(900).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Subir XML a la carpeta del periodo');
}

function JC_handlePickedFiles_(fileIds) {
  // Se llama desde Picker.html (google.picker), tras subir/seleccionar archivos.
  if (!fileIds || !fileIds.length) return;
  JC_parseFilesById_(fileIds);
}

// --- Placeholders for functions in the menu ---
function JC_buildPreview() {}
function JC_emitirPolizas() {}
function JC_importBanco() {}
function JC_runConciliacion() {}
function JC_refresh() {}