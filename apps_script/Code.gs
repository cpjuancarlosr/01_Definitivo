/* eslint-disable no-var */
/**
 * Jefatura Contable · Apps Script helpers.
 *
 * Mejora respecto al boceto original:
 *  - Soporta CFDI 3.3 y 4.0 con detección automática de namespace.
 *  - Evita duplicados por UUID antes de insertar.
 *  - Lee complementos de pago 1.0 y 2.0 para poblar la hoja "Pagos".
 *  - Escribe por lotes y centraliza utilidades (hojas, índices, logs).
 *  - Agrega bitácora básica para seguimiento.
 */

var JC = (function () {
  'use strict';

  /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet?} */
  var cachedSs = null;
  var cachedHeaders = {};

  var SHEET_NAMES = {
    SETUP: 'Setup',
    CFDI: 'CFDI',
    CONCEPTOS: 'Conceptos',
    PAGOS: 'Pagos',
    BITACORA: 'Bitácora'
  };

  /** Columnas esperadas en las hojas. Ajustar si la estructura cambia. */
  var HEADERS = {};
  HEADERS[SHEET_NAMES.CFDI] = [
      'ID',
      'Tipo',
      'UUID',
      'Serie',
      'Folio',
      'Fecha',
      'Periodo',
      'RFC_Emisor',
      'Nombre_Emisor',
      'RFC_Receptor',
      'Nombre_Receptor',
      'UsoCFDI',
      'Método',
      'Forma_Pago',
      'Moneda',
      'Tipo_Cambio',
      'Subtotal',
      'Descuento',
      'IVA_Trasladado_16',
      'IVA_Trasladado_08',
      'IVA_Trasladado_00',
      'IEPS',
      'Ret_ISR',
      'Ret_IVA',
      'Total',
      'Cancelado',
      'Motivo_Cancelacion',
      'ArchivoXML_ID',
      'ArchivoPDF_ID',
      'Link_XML',
      'Link_PDF',
      'Con_Póliza',
      'Incluida_Reportes',
      'Conciliada'
    ];
  HEADERS[SHEET_NAMES.CONCEPTOS] = [
      'UUID',
      'Renglon',
      'ClaveProdServ',
      'ClaveUnidad',
      'Unidad',
      'Cantidad',
      'Descripción',
      'Valor_Unitario',
      'Importe',
      'Descuento',
      'Tasa_IVA',
      'IVA_Importe',
      'IEPS_Tasa',
      'IEPS_Importe'
    ];
  HEADERS[SHEET_NAMES.PAGOS] = [
      'UUID_P',
      'Fecha_Pago',
      'Forma_Pago',
      'Monto',
      'Moneda',
      'Tipo_Cambio',
      'Parcialidad',
      'UUID_I_origen',
      'Serie_I',
      'Folio_I',
      'Saldo_Anterior',
      'Importe_Pagado',
      'Saldo_Ins',
      'Cuenta_Banco'
    ];
  HEADERS[SHEET_NAMES.BITACORA] = [
      'FechaHora',
      'Usuario',
      'Acción',
      'Entidad',
      'Referencia',
      'Detalle'
    ];

  function getSpreadsheet() {
    if (!cachedSs) {
      cachedSs = SpreadsheetApp.getActive();
    }
    return cachedSs;
  }

  function getSheet(name) {
    var sheet = getSpreadsheet().getSheetByName(name);
    if (!sheet) {
      throw new Error('No se encontró la hoja "' + name + '".');
    }
    return sheet;
  }

  function getSetup() {
    var sh = getSheet(SHEET_NAMES.SETUP);
    var get = function (row, col) {
      return sh.getRange(row, col).getValue();
    };
    return {
      rfc: String(get(2, 2) || ''),
      modoIVA: String(get(11, 2) || ''),
      folderPeriodoId: String(get(15, 2) || ''),
      folderXmlId: String(get(16, 2) || ''),
      folderPdfId: String(get(17, 2) || '')
    };
  }

  function toast(message, title, seconds) {
    getSpreadsheet().toast(message, title || 'Jefatura Contable', seconds || 5);
  }

  function toRow(obj, header) {
    return header.map(function (key) {
      var value = obj.hasOwnProperty(key) ? obj[key] : '';
      if (value === undefined || value === null) {
        return '';
      }
      return value;
    });
  }

  function appendRows(sheetName, rows) {
    if (!rows.length) {
      return;
    }
    var sheet = getSheet(sheetName);
    var header = getHeader(sheetName);
    var values = rows.map(function (row) {
      return toRow(row, header);
    });
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, values.length, header.length).setValues(values);
  }

  function getHeader(sheetName) {
    if (!cachedHeaders[sheetName]) {
      var sheet = getSheet(sheetName);
      var lastColumn = sheet.getLastColumn();
      if (!lastColumn || sheet.getLastRow() === 0) {
        var defaultHeader = HEADERS[sheetName];
        if (!defaultHeader) {
          throw new Error('No hay encabezado definido para ' + sheetName);
        }
        sheet.getRange(1, 1, 1, defaultHeader.length).setValues([defaultHeader]);
        lastColumn = defaultHeader.length;
      }
      var headerRange = sheet.getRange(1, 1, 1, lastColumn);
      cachedHeaders[sheetName] = headerRange.getValues()[0];
    }
    return cachedHeaders[sheetName];
  }

  function buildIndex(sheetName, keyColumnName) {
    var header = getHeader(sheetName);
    var colIndex = header.indexOf(keyColumnName);
    if (colIndex === -1) {
      throw new Error('No se encontró la columna "' + keyColumnName + '" en ' + sheetName);
    }
    var sheet = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {};
    }
    var values = sheet
      .getRange(2, colIndex + 1, lastRow - 1, 1)
      .getValues()
      .map(function (row) {
        return String(row[0] || '');
      });
    var index = {};
    values.forEach(function (value) {
      if (!value) {
        return;
      }
      index[value] = true;
    });
    return index;
  }

  function normalizeNumber(value) {
    if (value === null || value === undefined || value === '') {
      return 0;
    }
    var num = Number(value);
    if (isNaN(num)) {
      return 0;
    }
    return Number(num.toFixed(2));
  }

  function isoToDateString(isoString) {
    if (!isoString) {
      return '';
    }
    var date = new Date(isoString);
    if (isNaN(date.getTime())) {
      return '';
    }
    return Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd"T"HH:mm:ss'
    );
  }

  function calcPeriodo(isoString) {
    if (!isoString) {
      return '';
    }
    var date = new Date(isoString);
    if (isNaN(date.getTime())) {
      return '';
    }
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
  }

  function isXmlMime(file) {
    var mime = file.getMimeType();
    return mime === MimeType.XML || /xml/i.test(mime) || /text\//i.test(mime);
  }

  function parseXml(file) {
    var content = file.getBlob().getDataAsString('UTF-8');
    return XmlService.parse(content);
  }

  function getNamespace(root, prefix, uriCandidates) {
    if (!root) {
      return null;
    }
    var namespace = root.getNamespace(prefix);
    if (namespace && namespace.getURI()) {
      return namespace;
    }
    if (!uriCandidates || !uriCandidates.length) {
      return null;
    }
    for (var i = 0; i < uriCandidates.length; i++) {
      var candidate = XmlService.getNamespace(prefix, uriCandidates[i]);
      if (candidate) {
        return candidate;
      }
    }
    return null;
  }

  function attr(element, name) {
    if (!element) {
      return '';
    }
    var attribute = element.getAttribute(name);
    return attribute ? attribute.getValue() : '';
  }

  function sumImpuestosTrasladados(traslados, target, nsCfdi) {
    if (!traslados) {
      return target;
    }
    traslados.forEach(function (tras) {
      var tasa = Number(attr(tras, 'TasaOCuota') || attr(tras, 'TasaOCuota'));
      var impuesto = attr(tras, 'Impuesto');
      var importe = normalizeNumber(attr(tras, 'Importe'));
      if (!importe) {
        return;
      }
      if (impuesto === '002') {
        if (Math.abs(tasa - 0.16) < 1e-6) {
          target.IVA_Trasladado_16 += importe;
        } else if (Math.abs(tasa - 0.08) < 1e-6) {
          target.IVA_Trasladado_08 += importe;
        } else if (Math.abs(tasa) < 1e-6) {
          target.IVA_Trasladado_00 += importe;
        }
      } else if (impuesto === '003') {
        target.IEPS += importe;
      }
    });
    return target;
  }

  function sumImpuestosRetenidos(retenciones) {
    var result = { Ret_ISR: 0, Ret_IVA: 0 };
    if (!retenciones) {
      return result;
    }
    retenciones.forEach(function (ret) {
      var impuesto = attr(ret, 'Impuesto');
      var importe = normalizeNumber(attr(ret, 'Importe'));
      if (impuesto === '001') {
        result.Ret_ISR += importe;
      } else if (impuesto === '002') {
        result.Ret_IVA += importe;
      }
    });
    return result;
  }

  function extractConcepts(conceptos, nsCfdi) {
    if (!conceptos || !conceptos.length) {
      return [];
    }
    return conceptos.map(function (concepto) {
      var impuestos = concepto.getChild('Impuestos', nsCfdi);
      var traslados = impuestos && impuestos.getChild('Traslados', nsCfdi);
      var traslado = traslados && traslados.getChildren('Traslado', nsCfdi);
      var ivaImporte = 0;
      var ivaTasa = '';
      if (traslado && traslado.length) {
        var iva = traslado[0];
        ivaTasa = attr(iva, 'TasaOCuota') || attr(iva, 'Tasa') || '';
        ivaImporte = normalizeNumber(attr(iva, 'Importe'));
      }
      return {
        ClaveProdServ: attr(concepto, 'ClaveProdServ'),
        ClaveUnidad: attr(concepto, 'ClaveUnidad'),
        Unidad: attr(concepto, 'Unidad'),
        Cantidad: normalizeNumber(attr(concepto, 'Cantidad')),
        Descripción: attr(concepto, 'Descripcion') || attr(concepto, 'Descripcion'),
        Valor_Unitario: normalizeNumber(attr(concepto, 'ValorUnitario')),
        Importe: normalizeNumber(attr(concepto, 'Importe')),
        Descuento: normalizeNumber(attr(concepto, 'Descuento')),
        Tasa_IVA: ivaTasa,
        IVA_Importe: ivaImporte,
        IEPS_Tasa: '',
        IEPS_Importe: 0
      };
    });
  }

  function extractPagos(complemento, namespaces) {
    var pagos = [];
    if (!complemento) {
      return pagos;
    }
    var pagos20 = complemento.getChild('Pagos', namespaces.pagos20);
    if (pagos20) {
      pagos = pagos.concat(extractPagosDocto(pagos20, namespaces.pagos20));
    }
    var pagos10 = complemento.getChild('Pagos', namespaces.pagos10);
    if (pagos10) {
      pagos = pagos.concat(extractPagosDocto(pagos10, namespaces.pagos10));
    }
    return pagos;
  }

  function extractPagosDocto(pagosElement, namespace) {
    var pagoNodes = pagosElement.getChildren('Pago', namespace) || [];
    var pagos = [];
    pagoNodes.forEach(function (pago) {
      var cuentaBanco = attr(pago, 'CuentaOrdenante') || attr(pago, 'CuentaBeneficiario');
      var doctos = pago.getChildren('DoctoRelacionado', namespace) || [];
      if (!doctos.length) {
        pagos.push({
          Fecha_Pago: isoToDateString(attr(pago, 'FechaPago')),
          Forma_Pago: attr(pago, 'FormaDePagoP') || attr(pago, 'FormaDePagoP'),
          Monto: normalizeNumber(attr(pago, 'Monto')),
          Moneda: attr(pago, 'MonedaP'),
          Tipo_Cambio: attr(pago, 'TipoCambioP'),
          Parcialidad: '',
          UUID_I_origen: '',
          Serie_I: '',
          Folio_I: '',
          Saldo_Anterior: 0,
          Importe_Pagado: 0,
          Saldo_Ins: 0,
          Cuenta_Banco: cuentaBanco
        });
        return;
      }
      doctos.forEach(function (docto) {
        pagos.push({
          Fecha_Pago: isoToDateString(attr(pago, 'FechaPago')),
          Forma_Pago: attr(pago, 'FormaDePagoP') || attr(pago, 'FormaDePagoP'),
          Monto: normalizeNumber(attr(pago, 'Monto')),
          Moneda: attr(pago, 'MonedaP'),
          Tipo_Cambio: attr(pago, 'TipoCambioP'),
          Parcialidad: attr(docto, 'NumParcialidad'),
          UUID_I_origen: attr(docto, 'IdDocumento'),
          Serie_I: attr(docto, 'Serie'),
          Folio_I: attr(docto, 'Folio'),
          Saldo_Anterior: normalizeNumber(attr(docto, 'ImpSaldoAnt')),
          Importe_Pagado: normalizeNumber(attr(docto, 'ImpPagado')),
          Saldo_Ins: normalizeNumber(attr(docto, 'ImpSaldoInsoluto')),
          Cuenta_Banco: cuentaBanco
        });
      });
    });
    return pagos;
  }

  function parseCFDI(doc, fileId) {
    var root = doc.getRootElement();
    var namespaces = {
      cfdi: getNamespace(root, 'cfdi', [
        'http://www.sat.gob.mx/cfd/4',
        'http://www.sat.gob.mx/cfd/3'
      ]),
      tfd: getNamespace(root, 'tfd', ['http://www.sat.gob.mx/TimbreFiscalDigital']),
      pagos20: XmlService.getNamespace('pago20', 'http://www.sat.gob.mx/Pagos20'),
      pagos10: XmlService.getNamespace('pago10', 'http://www.sat.gob.mx/Pagos')
    };

    var comprobante = root;
    var tipo = attr(comprobante, 'TipoDeComprobante');
    var emisor = comprobante.getChild('Emisor', namespaces.cfdi);
    var receptor = comprobante.getChild('Receptor', namespaces.cfdi);
    var complemento = comprobante.getChild('Complemento', namespaces.cfdi);
    var timbre = complemento && complemento.getChild('TimbreFiscalDigital', namespaces.tfd);
    var uuid = timbre ? attr(timbre, 'UUID') : '';

    var conceptosParent = comprobante.getChild('Conceptos', namespaces.cfdi);
    var conceptosNodes = conceptosParent
      ? conceptosParent.getChildren('Concepto', namespaces.cfdi)
      : [];

    var impuestos = comprobante.getChild('Impuestos', namespaces.cfdi);
    var trasladosParent = impuestos && impuestos.getChild('Traslados', namespaces.cfdi);
    var traslados = trasladosParent
      ? trasladosParent.getChildren('Traslado', namespaces.cfdi)
      : [];
    var retencionesParent = impuestos && impuestos.getChild('Retenciones', namespaces.cfdi);
    var retenciones = retencionesParent
      ? retencionesParent.getChildren('Retencion', namespaces.cfdi)
      : [];

    var cfdiRecord = {
      Tipo: tipo,
      UUID: uuid,
      Serie: attr(comprobante, 'Serie'),
      Folio: attr(comprobante, 'Folio'),
      Fecha: isoToDateString(attr(comprobante, 'Fecha')),
      Periodo: calcPeriodo(attr(comprobante, 'Fecha')),
      RFC_Emisor: attr(emisor, 'Rfc'),
      Nombre_Emisor: attr(emisor, 'Nombre'),
      RFC_Receptor: attr(receptor, 'Rfc'),
      Nombre_Receptor: attr(receptor, 'Nombre'),
      UsoCFDI: attr(receptor, 'UsoCFDI') || attr(receptor, 'UsoCFDI'),
      'Método': attr(comprobante, 'MetodoPago') || attr(comprobante, 'MetodoPago'),
      Forma_Pago: attr(comprobante, 'FormaPago'),
      Moneda: attr(comprobante, 'Moneda') || 'MXN',
      Tipo_Cambio: attr(comprobante, 'TipoCambio'),
      Subtotal: normalizeNumber(attr(comprobante, 'SubTotal')),
      Descuento: normalizeNumber(attr(comprobante, 'Descuento')),
      IVA_Trasladado_16: 0,
      IVA_Trasladado_08: 0,
      IVA_Trasladado_00: 0,
      IEPS: 0,
      Ret_ISR: 0,
      Ret_IVA: 0,
      Total: normalizeNumber(attr(comprobante, 'Total')),
      Cancelado: 'Vigente',
      Motivo_Cancelacion: '',
      ArchivoXML_ID: fileId,
      ArchivoPDF_ID: '',
      Link_XML: fileId
        ? 'https://drive.google.com/file/d/' + fileId
        : '',
      Link_PDF: '',
      'Con_Póliza': 'No',
      Incluida_Reportes: 'Sí',
      Conciliada: 'No'
    };

    sumImpuestosTrasladados(traslados, cfdiRecord, namespaces.cfdi);
    var ret = sumImpuestosRetenidos(retenciones);
    cfdiRecord.Ret_ISR = normalizeNumber(ret.Ret_ISR);
    cfdiRecord.Ret_IVA = normalizeNumber(ret.Ret_IVA);

    var conceptos = extractConcepts(conceptosNodes, namespaces.cfdi);
    var pagos = extractPagos(complemento, namespaces);

    if (cfdiRecord.Tipo === 'P' && !pagos.length) {
      pagos.push({
        Fecha_Pago: cfdiRecord.Fecha,
        Forma_Pago: cfdiRecord.Forma_Pago,
        Monto: cfdiRecord.Total,
        Moneda: cfdiRecord.Moneda,
        Tipo_Cambio: cfdiRecord.Tipo_Cambio,
        Parcialidad: '',
        UUID_I_origen: '',
        Serie_I: '',
        Folio_I: '',
        Saldo_Anterior: 0,
        Importe_Pagado: cfdiRecord.Total,
        Saldo_Ins: 0,
        Cuenta_Banco: ''
      });
    }

    return {
      cfdi: cfdiRecord,
      conceptos: conceptos,
      pagos: pagos
    };
  }

  function logBitacora(action, entity, reference, detail) {
    try {
      appendRows(SHEET_NAMES.BITACORA, [
        {
          FechaHora: Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd HH:mm:ss'
          ),
          Usuario: Session.getActiveUser().getEmail() || 'Desconocido',
          Acción: action,
          Entidad: entity,
          Referencia: reference,
          Detalle: detail || ''
        }
      ]);
    } catch (error) {
      Logger.log('Bitácora no disponible: ' + error);
    }
  }

  function parseFiles(ids) {
    if (!ids || !ids.length) {
      toast('No se recibieron archivos para procesar.', 'Carga de XML');
      return;
    }

    var start = new Date();
    var uuidIndex = buildIndex(SHEET_NAMES.CFDI, 'UUID');
    var cfdiRows = [];
    var conceptosRows = [];
    var pagosRows = [];
    var duplicados = [];
    var procesados = 0;

    ids.forEach(function (id) {
      try {
        var file = DriveApp.getFileById(id);
        if (!isXmlMime(file)) {
          return;
        }
        var doc = parseXml(file);
        var parsed = parseCFDI(doc, id);
        if (!parsed.cfdi.UUID) {
          throw new Error('El archivo no contiene TimbreFiscalDigital UUID.');
        }
        if (uuidIndex[parsed.cfdi.UUID]) {
          duplicados.push(parsed.cfdi.UUID);
          return;
        }
        uuidIndex[parsed.cfdi.UUID] = true;
        cfdiRows.push(parsed.cfdi);
        parsed.conceptos.forEach(function (concepto, index) {
          conceptosRows.push(
            Object.assign(
              {
                UUID: parsed.cfdi.UUID,
                Renglon: index + 1
              },
              concepto
            )
          );
        });
        parsed.pagos.forEach(function (pago) {
          pagosRows.push(
            Object.assign(
              {
                UUID_P: parsed.cfdi.UUID
              },
              pago
            )
          );
        });
        procesados += 1;
      } catch (error) {
        Logger.log('Error al procesar archivo ' + id + ': ' + error);
        logBitacora('Error', 'CFDI', id, String(error));
      }
    });

    appendRows(SHEET_NAMES.CFDI, cfdiRows);
    appendRows(SHEET_NAMES.CONCEPTOS, conceptosRows);
    appendRows(SHEET_NAMES.PAGOS, pagosRows);

    var elapsed = (new Date().getTime() - start.getTime()) / 1000;
    var message =
      'CFDI nuevos: ' + cfdiRows.length +
      '\nConceptos agregados: ' + conceptosRows.length +
      '\nPagos agregados: ' + pagosRows.length +
      '\nTiempo: ' + elapsed.toFixed(2) + ' s';
    if (duplicados.length) {
      message += '\nUUID duplicados omitidos: ' + duplicados.join(', ');
    }
    toast(message, 'Procesamiento completado', Math.min(10, 5 + cfdiRows.length / 2));
    logBitacora('Importación XML', 'CFDI', procesados + ' archivos', message);
  }

  return {
    SHEET_NAMES: SHEET_NAMES,
    getSpreadsheet: getSpreadsheet,
    getSetup: getSetup,
    parseFiles: parseFiles,
    toast: toast
  };
})();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Jefatura Contable')
    .addItem('Cargar XML desde equipo…', 'JC_showPicker')
    .addItem('Procesar carpeta del periodo', 'JC_parseFolderPeriodo')
    .addSeparator()
    .addItem('Generar Previa de póliza', 'JC_buildPreview')
    .addItem('Emitir pólizas seleccionadas', 'JC_emitirPolizas')
    .addSeparator()
    .addItem('Importar estado de cuenta (PDF/CSV)', 'JC_importEstadoCuenta')
    .addItem('Conciliación sugerida', 'JC_runConciliacion')
    .addSeparator()
    .addItem('Exportar pólizas a PDF', 'JC_exportPolizasPDF')
    .addSeparator()
    .addItem('Bloquear/Desbloquear periodo', 'JC_toggleCierre')
    .addToUi();
}

function JC_showPicker() {
  var html = HtmlService.createTemplateFromFile('Picker')
    .evaluate()
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Subir XML a la carpeta del periodo');
}

function JC_handlePickedFiles_(fileIds) {
  if (!fileIds || !fileIds.length) {
    JC.toast('No se seleccionaron archivos.', 'Carga de XML');
    return;
  }
  JC.parseFiles(fileIds);
}

function JC_parseFolderPeriodo() {
  var setup = JC.getSetup();
  if (!setup.folderXmlId) {
    JC.toast('Configura el ID de la carpeta XML en Setup!B16.', 'Configuración incompleta');
    return;
  }
  var folder;
  try {
    folder = DriveApp.getFolderById(setup.folderXmlId);
  } catch (error) {
    JC.toast('No fue posible acceder a la carpeta XML. Verifica permisos.', 'Error');
    return;
  }
  var ids = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    ids.push(files.next().getId());
  }
  if (!ids.length) {
    JC.toast('La carpeta XML está vacía.', 'Procesar carpeta');
    return;
  }
  JC.parseFiles(ids);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Placeholders para funciones existentes en el proyecto original.
 * Se dejan vacíos para conservar el menú sin errores de referencia.
 */
function JC_buildPreview() {}
function JC_emitirPolizas() {}
function JC_importEstadoCuenta() {}
function JC_runConciliacion() {}
function JC_exportPolizasPDF() {}
function JC_toggleCierre() {}
