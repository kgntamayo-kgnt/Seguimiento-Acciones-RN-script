/**
 * =========================================================================
 * SISTEMA DE GESTIÓN Y CONTROL REGULATORIO - REACTOR IAN-R1
 * SERVICIO GEOLÓGICO COLOMBIANO (SGC) - MINISTERIO DE MINAS Y ENERGÍA (MME)
 * Licencia de Operación: LO-MME-013-2025-M2 | Resoluciones 18-1475 y 18-1434
 * =========================================================================
 * Archivo de backend para Google Apps Script (Web App + Google Sheets)
 */

const SHEET_COMMUNICATIONS = 'Comunicaciones';
const SHEET_ACTIONS = 'Actuaciones';

/**
 * Función principal doGet para servir la Web App con el motor HtmlService
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistema de Gestión y Control Regulatorio - Reactor IAN-R1')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Obtiene la hoja de cálculo activa o busca por ID configurado en Script Properties
 */
function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('SPREADSHEET_ID');
  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (err) {
      Logger.log('Error abriendo spreadsheet por ID: ' + err);
    }
  }

  // Si no está vinculado ni configurado, crea un nuevo Google Sheet automáticamente
  var newSs = SpreadsheetApp.create('Control Regulatorio Reactor IAN-R1 (MME - SGC)');
  scriptProperties.setProperty('SPREADSHEET_ID', newSs.getId());
  return newSs;
}

/**
 * Inicializa las pestañas y encabezados en Google Sheets si están vacías
 */
function initSheetsIfNeeded(ss) {
  if (!ss) ss = getSpreadsheet();

  // 1. Hoja Comunicaciones
  var commSheet = ss.getSheetByName(SHEET_COMMUNICATIONS);
  if (!commSheet) {
    commSheet = ss.insertSheet(SHEET_COMMUNICATIONS);
    var commHeaders = [
      'ID', 'Tipo', 'Fecha', 'Radicado MME', 'Radicado SGC', 
      'Asunto', 'Remitente', 'Destinatario', 'Documentos', 
      'Actuaciones Requeridas', 'Fecha Límite', 'Estado', 'Notas'
    ];
    commSheet.getRange(1, 1, 1, commHeaders.length).setValues([commHeaders]);
    commSheet.getRange(1, 1, 1, commHeaders.length)
      .setBackground('#1e293b')
      .setFontColor('#f8fafc')
      .setFontWeight('bold');
    commSheet.setFrozenRows(1);

    // Cargar datos iniciales del Reactor IAN-R1
    var initialCommRows = [
      [
        'COMM-2026-001', 'ENTRADA', '2026-09-02', '1-2026-041551', '2026-800-010844-2',
        'Anexo instructivo IN-TNU-MRN-016 (Actualización de instructivo de control de acceso)',
        'Jimmy Alejandro Muñoz Rocha (SGC)', 'Juan Pablo Parra (MME - Grupo de Asuntos Nucleares)',
        'IN-TNU-MRN-016, F-TNU-MRN-067',
        'Recepción y cotejo del documento. Incorporar al expediente y verificar correspondencia técnica con versión 09.',
        '2026-09-30', 'CERRADO', 'Verificado cumplimiento del Art. 39 Res. 18-1475. Se generó oficio oficial de respuesta MME.'
      ],
      [
        'COMM-2026-002', 'ENTRADA', '2026-08-12', '1-2026-040542', '2026-840-009828-2',
        'Novedades sobre fuentes de referencia Sr-90 y Am-241 para detector de aerosoles',
        'Jimmy Alejandro Muñoz Rocha (SGC)', 'Juan Pablo Parra (MME - Grupo de Asuntos Nucleares)',
        'IN-TNU-MRN-011, F-TNU-MRN-060, F-TNU-MRN-063',
        '1. Recibir nuevas fuentes compatibles del proveedor. 2. Calibración in situ del detector de aerosoles. 3. Actualizar instructivos y programas de calibración. 4. Remitir dossier final al MME.',
        '2026-11-15', 'PENDIENTE', 'Operatividad y puesta en servicio formalmente suspendida hasta contar con fuentes compatibles (Res. 18-1434, Arts. 49 y 50).'
      ],
      [
        'COMM-2026-003', 'SALIDA', '2026-08-01', '1-2026-035120', '2026-800-009210-2',
        'Remisión Informe Semestral de Operación y Funcionamiento (I Semestre 2026)',
        'Jimmy Alejandro Muñoz Rocha (SGC)', 'Ministerio de Minas y Energía',
        'Informe Semestral 2026-I',
        'Cotejo de horas de funcionamiento, dosimetría y novedades operativas.',
        '2026-08-15', 'CERRADO', 'Cumplimiento estricto del Art. 21, Numeral 1 de la Resolución 18-1475 de 2004.'
      ]
    ];
    commSheet.getRange(2, 1, initialCommRows.length, commHeaders.length).setValues(initialCommRows);
    commSheet.autoResizeColumns(1, commHeaders.length);
  }

  // 2. Hoja Actuaciones Técnicas
  var actSheet = ss.getSheetByName(SHEET_ACTIONS);
  if (!actSheet) {
    actSheet = ss.insertSheet(SHEET_ACTIONS);
    var actHeaders = [
      'ID', 'Título', 'Documento', 'Responsable', 
      'Fecha Límite', 'Prioridad', 'Estado', 'Descripción'
    ];
    actSheet.getRange(1, 1, 1, actHeaders.length).setValues([actHeaders]);
    actSheet.getRange(1, 1, 1, actHeaders.length)
      .setBackground('#1e293b')
      .setFontColor('#f8fafc')
      .setFontWeight('bold');
    actSheet.setFrozenRows(1);

    // Cargar actuaciones iniciales
    var initialActRows = [
      [
        'ACT-001', 'Calibración in situ del detector de aerosoles con nuevas fuentes Sr-90 / Am-241',
        'IN-TNU-MRN-011 / F-TNU-MRN-060', 'Andrea Sánchez (OPR)', '2026-11-15', 'ALTA', 'PENDIENTE',
        'Coordinar con fabricante internacional la recepción de las fuentes exentas de geometría compatible y ejecutar la calibración in situ para levantar la suspensión del detector.'
      ],
      [
        'ACT-002', 'Actualización y estandarización procedimental (IN-TNU-MRN-011, F-TNU-MRN-060 y 063)',
        'IN-TNU-MRN-011', 'Edgar Mauricio López (Supervisor)', '2026-12-01', 'ALTA', 'PENDIENTE',
        'Formalizar en los procedimientos del reactor los métodos de verificación periódica y calibración de aerosoles y remitirlos al MME conforme al Art. 39 de la Res. 18-1475.'
      ],
      [
        'ACT-003', 'Elaboración y radicación del Informe Anual 2026 (primeros 15 días hábiles Enero 2027)',
        'Informe Anual IAN-R1', 'Jimmy Alejandro Muñoz (Director)', '2027-01-22', 'MEDIA', 'PENDIENTE',
        'Compilar datos de operación, reentrenamiento, vigilancia dosimétrica, comités de seguridad y programa de envejecimiento (Art. 21, Num. 2, Res. 18-1475).'
      ],
      [
        'ACT-004', 'Seguimiento a la segunda inspección semestral regulatoria de 2026',
        'Acta de Inspección MME', 'Juan Pablo Parra (MME)', '2026-12-15', 'ALTA', 'PENDIENTE',
        'Coordinación y atención a la comisión de inspección del Grupo de Asuntos Nucleares (Art. 46, Res. 18-1475).'
      ]
    ];
    actSheet.getRange(2, 1, initialActRows.length, actHeaders.length).setValues(initialActRows);
    actSheet.autoResizeColumns(1, actHeaders.length);
  }
}

/**
 * Lee todas las comunicaciones y actuaciones técnicas desde Google Sheets
 * Se invoca desde el frontend mediante google.script.run.getAppData()
 */
function getAppData() {
  try {
    var ss = getSpreadsheet();
    initSheetsIfNeeded(ss);

    // Leer Comunicaciones
    var commSheet = ss.getSheetByName(SHEET_COMMUNICATIONS);
    var commData = [];
    var lastCommRow = commSheet.getLastRow();
    if (lastCommRow > 1) {
      var commValues = commSheet.getRange(2, 1, lastCommRow - 1, 13).getValues();
      for (var i = 0; i < commValues.length; i++) {
        var row = commValues[i];
        if (!row[0]) continue;
        var docs = [];
        if (row[8]) {
          docs = String(row[8]).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
        }
        var dateStr = '';
        if (row[2] instanceof Date) {
          dateStr = Utilities.formatDate(row[2], Session.getScriptTimeZone() || 'America/Bogota', 'yyyy-MM-dd');
        } else {
          dateStr = String(row[2] || '');
        }

        var deadlineStr = '';
        if (row[10] instanceof Date) {
          deadlineStr = Utilities.formatDate(row[10], Session.getScriptTimeZone() || 'America/Bogota', 'yyyy-MM-dd');
        } else {
          deadlineStr = String(row[10] || '');
        }

        commData.push({
          id: String(row[0]),
          type: String(row[1] || 'ENTRADA'),
          date: dateStr,
          radMME: String(row[3] || ''),
          radSGC: String(row[4] || ''),
          subject: String(row[5] || ''),
          sender: String(row[6] || ''),
          receiver: String(row[7] || ''),
          docs: docs,
          actionsReq: String(row[9] || ''),
          deadline: deadlineStr,
          status: String(row[11] || 'PENDIENTE'),
          notes: String(row[12] || '')
        });
      }
    }

    // Leer Actuaciones Técnicas
    var actSheet = ss.getSheetByName(SHEET_ACTIONS);
    var actData = [];
    var lastActRow = actSheet.getLastRow();
    if (lastActRow > 1) {
      var actValues = actSheet.getRange(2, 1, lastActRow - 1, 8).getValues();
      for (var j = 0; j < actValues.length; j++) {
        var actRow = actValues[j];
        if (!actRow[0]) continue;
        var actDeadline = '';
        if (actRow[4] instanceof Date) {
          actDeadline = Utilities.formatDate(actRow[4], Session.getScriptTimeZone() || 'America/Bogota', 'yyyy-MM-dd');
        } else {
          actDeadline = String(actRow[4] || '');
        }

        actData.push({
          id: String(actRow[0]),
          title: String(actRow[1] || ''),
          doc: String(actRow[2] || ''),
          owner: String(actRow[3] || ''),
          deadline: actDeadline,
          priority: String(actRow[5] || 'ALTA'),
          status: String(actRow[6] || 'PENDIENTE'),
          desc: String(actRow[7] || '')
        });
      }
    }

    return {
      success: true,
      sheetUrl: ss.getUrl(),
      spreadsheetId: ss.getId(),
      communications: commData,
      technicalActions: actData
    };
  } catch (error) {
    Logger.log('Error en getAppData: ' + error);
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}

/**
 * Guarda o actualiza una Comunicación en Google Sheets
 */
function saveCommunication(comm) {
  try {
    var ss = getSpreadsheet();
    initSheetsIfNeeded(ss);
    var sheet = ss.getSheetByName(SHEET_COMMUNICATIONS);
    var lastRow = sheet.getLastRow();
    var docsString = Array.isArray(comm.docs) ? comm.docs.join(', ') : (comm.docs || '');
    
    var rowData = [
      comm.id || ('COMM-' + new Date().getTime()),
      comm.type || 'ENTRADA',
      comm.date || '',
      comm.radMME || '',
      comm.radSGC || '',
      comm.subject || '',
      comm.sender || '',
      comm.receiver || '',
      docsString,
      comm.actionsReq || '',
      comm.deadline || '',
      comm.status || 'PENDIENTE',
      comm.notes || ''
    ];

    var foundRow = -1;
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(comm.id)) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, id: rowData[0] };
  } catch (err) {
    Logger.log('Error en saveCommunication: ' + err);
    throw new Error('Error al guardar comunicación en Sheets: ' + err.message);
  }
}

/**
 * Elimina una Comunicación por ID
 */
function deleteCommunication(commId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_COMMUNICATIONS);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(commId)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: true, notFound: true };
  } catch (err) {
    throw new Error('Error al eliminar comunicación en Sheets: ' + err.message);
  }
}

/**
 * Guarda o actualiza una Actuación Técnica en Google Sheets
 */
function saveTechnicalAction(act) {
  try {
    var ss = getSpreadsheet();
    initSheetsIfNeeded(ss);
    var sheet = ss.getSheetByName(SHEET_ACTIONS);
    var lastRow = sheet.getLastRow();

    var rowData = [
      act.id || ('ACT-' + new Date().getTime()),
      act.title || '',
      act.doc || '',
      act.owner || '',
      act.deadline || '',
      act.priority || 'ALTA',
      act.status || 'PENDIENTE',
      act.desc || ''
    ];

    var foundRow = -1;
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(act.id)) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, id: rowData[0] };
  } catch (err) {
    Logger.log('Error en saveTechnicalAction: ' + err);
    throw new Error('Error al guardar actuación en Sheets: ' + err.message);
  }
}

/**
 * Actualiza únicamente el estado de una actuación (ej. PENDIENTE / COMPLETADO)
 */
function updateActionStatus(actId, newStatus) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ACTIONS);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'No data' };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(actId)) {
        sheet.getRange(i + 2, 7).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, notFound: true };
  } catch (err) {
    throw new Error('Error al actualizar estado en Sheets: ' + err.message);
  }
}

/**
 * Elimina una Actuación Técnica por ID
 */
function deleteTechnicalAction(actId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ACTIONS);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(actId)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: true, notFound: true };
  } catch (err) {
    throw new Error('Error al eliminar actuación en Sheets: ' + err.message);
  }
}

/**
 * Servicio proxy de Gemini AI dentro de Google Apps Script usando UrlFetchApp
 * Utiliza la clave GEMINI_API_KEY configurada en Script Properties de Apps Script.
 */
function callGeminiApiGAS(payload) {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var apiKey = scriptProperties.getProperty('GEMINI_API_KEY');

    if (!apiKey) {
      return {
        error: 'Para usar Gemini AI en Google Apps Script, configura la propiedad GEMINI_API_KEY en: Configuración del proyecto > Propiedades de la secuencia de comandos.'
      };
    }

    var model = 'gemini-3.8-flash';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

    var body = {
      contents: payload.contents
    };

    if (payload.systemInstruction) {
      body.systemInstruction = payload.systemInstruction;
    }

    if (payload.generationConfig) {
      body.generationConfig = payload.generationConfig;
    }

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log('Error Gemini API (' + responseCode + '): ' + responseText);
      return { error: 'Error Gemini HTTP ' + responseCode + ': ' + responseText };
    }

    var json = JSON.parse(responseText);
    return json;
  } catch (err) {
    Logger.log('Excepción en callGeminiApiGAS: ' + err);
    return { error: 'Excepción en Google Apps Script: ' + err.message };
  }
}

/**
 * Función utilitaria para inicializar manualmente las hojas desde el editor de Apps Script
 */
function setupInitialSheets() {
  var ss = getSpreadsheet();
  initSheetsIfNeeded(ss);
  Logger.log('Hojas configuradas correctamente en: ' + ss.getUrl());
}
