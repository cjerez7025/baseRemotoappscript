/**
 * MÓDULO 3: GESTIÓN DE PROGRESO
 * Funciones para manejar el progreso del procesamiento
 */

/**
 * Obtiene el estado actual del progreso
 */
function obtenerProgreso() {
  try {
    var props = PropertiesService.getScriptProperties();
    var progresoJSON = props.getProperty('PROGRESO_ACTUAL');
    
    if (!progresoJSON) {
      return {
        etapa: 0,
        total: 13,
        mensaje: 'Iniciando...',
        porcentaje: 0,
        ejecutivosCreados: 0,
        ejecutivosTotal: 0
      };
    }
    
    return JSON.parse(progresoJSON);
  } catch (e) {
    Logger.log('Error en obtenerProgreso: ' + e.toString());
    return {
      etapa: 0,
      total: 13,
      mensaje: 'Error al obtener progreso',
      porcentaje: 0,
      ejecutivosCreados: 0,
      ejecutivosTotal: 0
    };
  }
}

/**
 * Actualiza el progreso del procesamiento
 */
function setProgreso(etapa, mensaje, porcentaje, ejecutivosCreados, ejecutivosTotal) {
  try {
    var progreso = {
      etapa: etapa || 0,
      total: 13,
      mensaje: mensaje || '',
      porcentaje: Math.min(100, Math.max(0, porcentaje || 0)),
      ejecutivosCreados: ejecutivosCreados || 0,
      ejecutivosTotal: ejecutivosTotal || 0,
      timestamp: new Date().getTime()
    };
    
    var props = PropertiesService.getScriptProperties();
    props.setProperty('PROGRESO_ACTUAL', JSON.stringify(progreso));
    
    SpreadsheetApp.flush();
    Logger.log('Progreso: Etapa ' + etapa + ' - ' + mensaje + ' (' + porcentaje + '%)');
  } catch (e) {
    Logger.log('Error actualizando progreso: ' + e.toString());
  }
}