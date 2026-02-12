/**
 * SISTEMA DE RADICACIÓN DE SOLICITUDES - GOOGLE APPS SCRIPT
 * 
 * Este código debe ser copiado en Google Apps Script:
 * 1. Ir a https://script.google.com
 * 2. Crear un nuevo proyecto
 * 3. Pegar este código
 * 4. Crear una nueva hoja de cálculo de Google Sheets
 * 5. Copiar el ID de la hoja (está en la URL)
 * 6. Reemplazar SPREADSHEET_ID con tu ID
 * 7. Implementar como aplicación web
 */

// ============================================
// CONFIGURACIÓN
// ============================================

const SPREADSHEET_ID = 'TU_SPREADSHEET_ID_AQUI';
const SHEET_NAME = 'Solicitudes';

// Columnas de la hoja de cálculo
const COLUMNS = {
  NUMERO_RADICADO: 0,
  FECHA_SOLICITUD: 1,
  NOMBRE_RADICADOR: 2,
  ID_RADICADOR: 3,
  SEDE: 4,
  EQUIPO_REVISAR: 5,
  PRIORIDAD: 6,
  CATEGORIA: 7,
  DETALLE_SOLICITUD: 8,
  AFECTACION: 9,
  ESTADO: 10,
  GESTOR: 11,
  RESPUESTA: 12,
  FECHA_ACTUALIZACION: 13,
  METRICAS: 14
};

// ============================================
// FUNCIÓN PRINCIPAL - doPost
// ============================================

/**
 * Maneja las peticiones POST desde el formulario web
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.action === 'update') {
      // Actualizar solicitud existente
      return updateSolicitud(data);
    } else {
      // Crear nueva solicitud
      return createSolicitud(data);
    }
    
  } catch (error) {
    Logger.log('Error en doPost: ' + error);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// FUNCIÓN PRINCIPAL - doGet
// ============================================

/**
 * Maneja las peticiones GET para obtener datos
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'getAll') {
      return getAllSolicitudes();
    } else if (action === 'getByRadicado') {
      return getSolicitudByRadicado(e.parameter.radicado);
    } else {
      return getAllSolicitudes();
    }
    
  } catch (error) {
    Logger.log('Error en doGet: ' + error);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// FUNCIONES DE CREACIÓN
// ============================================

/**
 * Crea una nueva solicitud en Google Sheets
 */
function createSolicitud(data) {
  const sheet = getOrCreateSheet();
  
  // Preparar fila de datos
  const rowData = [
    data.numeroRadicado,
    data.fechaSolicitud,
    data.nombreRadicador,
    data.idRadicador,
    data.sede,
    data.equipoRevisar,
    data.prioridad,
    data.categoria,
    data.detalleSolicitud,
    data.afectacion,
    data.estado,
    data.gestor,
    data.respuesta || '',
    new Date().toISOString(),
    JSON.stringify(initializeMetrics())
  ];
  
  // Agregar fila a la hoja
  sheet.appendRow(rowData);
  
  // Aplicar formato
  const lastRow = sheet.getLastRow();
  formatRow(sheet, lastRow, data);
  
  // Registrar en el log
  logAction('CREATE', data.numeroRadicado, data.nombreRadicador);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      radicado: data.numeroRadicado
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Actualiza una solicitud existente
 */
function updateSolicitud(data) {
  const sheet = getOrCreateSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Buscar la fila con el número de radicado
  for (let i = 1; i < values.length; i++) {
    if (values[i][COLUMNS.NUMERO_RADICADO] === data.numeroRadicado) {
      // Actualizar estado y respuesta
      sheet.getRange(i + 1, COLUMNS.ESTADO + 1).setValue(data.estado);
      sheet.getRange(i + 1, COLUMNS.RESPUESTA + 1).setValue(data.respuesta);
      sheet.getRange(i + 1, COLUMNS.FECHA_ACTUALIZACION + 1).setValue(new Date().toISOString());
      
      // Actualizar métricas
      const currentMetrics = JSON.parse(values[i][COLUMNS.METRICAS] || '{}');
      const updatedMetrics = updateMetrics(currentMetrics, data.estado);
      sheet.getRange(i + 1, COLUMNS.METRICAS + 1).setValue(JSON.stringify(updatedMetrics));
      
      // Aplicar formato según el nuevo estado
      formatRowByStatus(sheet, i + 1, data.estado);
      
      // Registrar en el log
      logAction('UPDATE', data.numeroRadicado, data.estado);
      
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          radicado: data.numeroRadicado
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // Si no se encuentra el radicado
  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: 'Radicado no encontrado'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// FUNCIONES DE CONSULTA
// ============================================

/**
 * Obtiene todas las solicitudes
 */
function getAllSolicitudes() {
  const sheet = getOrCreateSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Saltar la fila de encabezados
  const solicitudes = [];
  for (let i = 1; i < values.length; i++) {
    solicitudes.push({
      numeroRadicado: values[i][COLUMNS.NUMERO_RADICADO],
      fechaSolicitud: values[i][COLUMNS.FECHA_SOLICITUD],
      nombreRadicador: values[i][COLUMNS.NOMBRE_RADICADOR],
      idRadicador: values[i][COLUMNS.ID_RADICADOR],
      sede: values[i][COLUMNS.SEDE],
      equipoRevisar: values[i][COLUMNS.EQUIPO_REVISAR],
      prioridad: values[i][COLUMNS.PRIORIDAD],
      categoria: values[i][COLUMNS.CATEGORIA],
      detalleSolicitud: values[i][COLUMNS.DETALLE_SOLICITUD],
      afectacion: values[i][COLUMNS.AFECTACION],
      estado: values[i][COLUMNS.ESTADO],
      gestor: values[i][COLUMNS.GESTOR],
      respuesta: values[i][COLUMNS.RESPUESTA],
      fechaActualizacion: values[i][COLUMNS.FECHA_ACTUALIZACION],
      metricas: values[i][COLUMNS.METRICAS] ? JSON.parse(values[i][COLUMNS.METRICAS]) : {}
    });
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(solicitudes))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Obtiene una solicitud específica por número de radicado
 */
function getSolicitudByRadicado(radicado) {
  const sheet = getOrCreateSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][COLUMNS.NUMERO_RADICADO] === radicado) {
      const solicitud = {
        numeroRadicado: values[i][COLUMNS.NUMERO_RADICADO],
        fechaSolicitud: values[i][COLUMNS.FECHA_SOLICITUD],
        nombreRadicador: values[i][COLUMNS.NOMBRE_RADICADOR],
        idRadicador: values[i][COLUMNS.ID_RADICADOR],
        sede: values[i][COLUMNS.SEDE],
        equipoRevisar: values[i][COLUMNS.EQUIPO_REVISAR],
        prioridad: values[i][COLUMNS.PRIORIDAD],
        categoria: values[i][COLUMNS.CATEGORIA],
        detalleSolicitud: values[i][COLUMNS.DETALLE_SOLICITUD],
        afectacion: values[i][COLUMNS.AFECTACION],
        estado: values[i][COLUMNS.ESTADO],
        gestor: values[i][COLUMNS.GESTOR],
        respuesta: values[i][COLUMNS.RESPUESTA],
        fechaActualizacion: values[i][COLUMNS.FECHA_ACTUALIZACION],
        metricas: values[i][COLUMNS.METRICAS] ? JSON.parse(values[i][COLUMNS.METRICAS]) : {}
      };
      
      return ContentService
        .createTextOutput(JSON.stringify(solicitud))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: 'Radicado no encontrado'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// FUNCIONES DE MÉTRICAS
// ============================================

/**
 * Inicializa las métricas de calidad de software
 * Basado en: https://innevo.com/blog/metricas-de-calidad-del-software
 */
function initializeMetrics() {
  return {
    tiempoCreacion: new Date().toISOString(),
    tiempoResolucion: null,
    cambiosEstado: [],
    tiempoRespuesta: null,
    eficienciaResolucion: null,
    retrabajos: 0,
    cumplimientoSLA: null
  };
}

/**
 * Actualiza las métricas cuando cambia el estado
 */
function updateMetrics(currentMetrics, newStatus) {
  const now = new Date().toISOString();
  
  // Registrar cambio de estado
  if (!currentMetrics.cambiosEstado) {
    currentMetrics.cambiosEstado = [];
  }
  currentMetrics.cambiosEstado.push({
    estado: newStatus,
    fecha: now
  });
  
  // Si se completa, calcular tiempo de resolución
  if (newStatus === 'Completado' && !currentMetrics.tiempoResolucion) {
    const creacion = new Date(currentMetrics.tiempoCreacion);
    const resolucion = new Date(now);
    currentMetrics.tiempoResolucion = (resolucion - creacion) / (1000 * 60 * 60); // Horas
  }
  
  // Contar retrabajos (veces que vuelve a pendiente o reasignado)
  if (newStatus === 'Reasignado' || newStatus === 'Pendiente') {
    currentMetrics.retrabajos = (currentMetrics.retrabajos || 0) + 1;
  }
  
  return currentMetrics;
}

/**
 * Calcula las métricas agregadas del sistema
 */
function calculateSystemMetrics() {
  const sheet = getOrCreateSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  let totalSolicitudes = values.length - 1;
  let completadas = 0;
  let tiemposResolucion = [];
  let porPrioridad = { 'Baja': 0, 'Media': 0, 'Alta': 0, 'Crítica': 0 };
  let porEstado = { 'Pendiente': 0, 'En Proceso': 0, 'Completado': 0, 'Reasignado': 0 };
  
  for (let i = 1; i < values.length; i++) {
    const estado = values[i][COLUMNS.ESTADO];
    const prioridad = values[i][COLUMNS.PRIORIDAD];
    const metricas = values[i][COLUMNS.METRICAS] ? JSON.parse(values[i][COLUMNS.METRICAS]) : {};
    
    // Contar por estado
    if (porEstado.hasOwnProperty(estado)) {
      porEstado[estado]++;
    }
    
    // Contar por prioridad
    if (porPrioridad.hasOwnProperty(prioridad)) {
      porPrioridad[prioridad]++;
    }
    
    // Tiempo de resolución
    if (estado === 'Completado' && metricas.tiempoResolucion) {
      completadas++;
      tiemposResolucion.push(metricas.tiempoResolucion);
    }
  }
  
  const promedioResolucion = tiemposResolucion.length > 0 
    ? tiemposResolucion.reduce((a, b) => a + b, 0) / tiemposResolucion.length 
    : 0;
  
  return {
    totalSolicitudes,
    completadas,
    tasaCompletitud: (completadas / totalSolicitudes * 100).toFixed(2),
    promedioResolucion: promedioResolucion.toFixed(2),
    porPrioridad,
    porEstado
  };
}

// ============================================
// FUNCIONES AUXILIARES
// ============================================

/**
 * Obtiene o crea la hoja de cálculo
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    initializeSheet(sheet);
  }
  
  return sheet;
}

/**
 * Inicializa la hoja con encabezados y formato
 */
function initializeSheet(sheet) {
  const headers = [
    'Número Radicado',
    'Fecha Solicitud',
    'Nombre Radicador',
    'ID Radicador',
    'Sede',
    'Equipo a Revisar',
    'Prioridad',
    'Categoría',
    'Detalle Solicitud',
    'Afectación',
    'Estado',
    'Gestor',
    'Respuesta',
    'Fecha Actualización',
    'Métricas'
  ];
  
  sheet.appendRow(headers);
  
  // Formato de encabezados
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0A2540');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Ajustar anchos de columna
  sheet.setColumnWidth(1, 180); // Número Radicado
  sheet.setColumnWidth(2, 120); // Fecha
  sheet.setColumnWidth(3, 200); // Nombre
  sheet.setColumnWidth(4, 120); // ID
  sheet.setColumnWidth(5, 120); // Sede
  sheet.setColumnWidth(6, 120); // Equipo
  sheet.setColumnWidth(7, 100); // Prioridad
  sheet.setColumnWidth(8, 100); // Categoría
  sheet.setColumnWidth(9, 400); // Detalle
  sheet.setColumnWidth(10, 100); // Afectación
  sheet.setColumnWidth(11, 120); // Estado
  sheet.setColumnWidth(12, 120); // Gestor
  sheet.setColumnWidth(13, 400); // Respuesta
  sheet.setColumnWidth(14, 180); // Fecha Actualización
  sheet.setColumnWidth(15, 300); // Métricas
  
  // Congelar fila de encabezados
  sheet.setFrozenRows(1);
}

/**
 * Aplica formato a una fila según los datos
 */
function formatRow(sheet, row, data) {
  // Formato según prioridad
  const prioridadCell = sheet.getRange(row, COLUMNS.PRIORIDAD + 1);
  switch(data.prioridad) {
    case 'Crítica':
      prioridadCell.setBackground('#FF4757');
      prioridadCell.setFontColor('#FFFFFF');
      break;
    case 'Alta':
      prioridadCell.setBackground('#F97316');
      prioridadCell.setFontColor('#FFFFFF');
      break;
    case 'Media':
      prioridadCell.setBackground('#FFB020');
      prioridadCell.setFontColor('#FFFFFF');
      break;
    case 'Baja':
      prioridadCell.setBackground('#10B981');
      prioridadCell.setFontColor('#FFFFFF');
      break;
  }
  
  // Formato según estado
  formatRowByStatus(sheet, row, data.estado);
}

/**
 * Aplica formato según el estado
 */
function formatRowByStatus(sheet, row, estado) {
  const estadoCell = sheet.getRange(row, COLUMNS.ESTADO + 1);
  
  switch(estado) {
    case 'Completado':
      estadoCell.setBackground('#10B981');
      estadoCell.setFontColor('#FFFFFF');
      break;
    case 'En Proceso':
      estadoCell.setBackground('#3B82F6');
      estadoCell.setFontColor('#FFFFFF');
      break;
    case 'Pendiente':
      estadoCell.setBackground('#FFB020');
      estadoCell.setFontColor('#FFFFFF');
      break;
    case 'Reasignado':
      estadoCell.setBackground('#8B5CF6');
      estadoCell.setFontColor('#FFFFFF');
      break;
  }
  
  estadoCell.setFontWeight('bold');
  estadoCell.setHorizontalAlignment('center');
}

/**
 * Registra acciones en el log
 */
function logAction(action, radicado, details) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let logSheet = ss.getSheetByName('Log');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('Log');
    logSheet.appendRow(['Timestamp', 'Acción', 'Radicado', 'Detalles']);
  }
  
  logSheet.appendRow([
    new Date().toISOString(),
    action,
    radicado,
    details
  ]);
}

/**
 * Genera reporte de métricas (función auxiliar)
 */
function generateMetricsReport() {
  const metrics = calculateSystemMetrics();
  Logger.log('Métricas del Sistema:');
  Logger.log(JSON.stringify(metrics, null, 2));
  return metrics;
}
