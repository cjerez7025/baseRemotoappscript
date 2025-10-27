/**
 * MÓDULO 7: PROCESO PRINCIPAL
 * Orquesta todo el flujo de procesamiento
 */

/**
 * Inicia el proceso completo de ejecutivos
 */
function procesarEjecutivos() {
  try {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('PROGRESO_ACTUAL');
    
    var html = HtmlService.createHtmlOutputFromFile('ProgressUI')
      .setWidth(620)
      .setHeight(520);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Procesamiento de Ejecutivos');
    
    Utilities.sleep(500);
    ejecutarProcesoCompleto();
    
  } catch (e) {
    Logger.log('Error en procesarEjecutivos: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error', 'No se pudo iniciar: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Ejecuta el proceso completo paso a paso
 */
function ejecutarProcesoCompleto() {
  var ui = SpreadsheetApp.getUi();
  var ss = null;
  var ejecutivosArray = [];
  
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ETAPA 1: Detección (5%)
    setProgreso(1, '🔍 Buscando hoja BBDD_*_REMOTO*...', 5, 0, 0);
    Utilities.sleep(500);
    
    var hojaOrigen = obtenerHojaOrigen(ss);
    if (!hojaOrigen) {
      throw new Error('No se encontró hoja BBDD_*_REMOTO*');
    }
    
    // ETAPA 2: Lectura (10%)
    setProgreso(2, '📊 Leyendo datos...', 10, 0, 0);
    Utilities.sleep(500);
    
    var datos = hojaOrigen.getDataRange().getValues();
    var encabezados = datos[0];
    var filasDatos = datos.slice(1);
    
    // ETAPA 3: Análisis (15%)
    setProgreso(3, '👥 Identificando ejecutivos...', 15, 0, 0);
    Utilities.sleep(500);
    
    var ejecutivosPorNombre = agruparPorEjecutivo(filasDatos, encabezados);
    ejecutivosArray = Object.keys(ejecutivosPorNombre);
    
    if (ejecutivosArray.length === 0) {
      throw new Error('No se encontraron ejecutivos');
    }
    
    // ETAPA 4: Validación (20%)
    setProgreso(4, '✓ Validando ' + ejecutivosArray.length + ' ejecutivos...', 20, 0, ejecutivosArray.length);
    Utilities.sleep(500);
    
    var hojas = ss.getSheets();
    var alertasEjecutivos = validarEjecutivosEnBase(ejecutivosPorNombre, hojas);
    
    // ETAPA 5: Creación (20-50%)
    for (var i = 0; i < ejecutivosArray.length; i++) {
      var nombreEjecutivo = ejecutivosArray[i];
      var porcentaje = 20 + ((i + 1) / ejecutivosArray.length) * 30;
      
      setProgreso(5, '📄 Creando: ' + nombreEjecutivo.replace(/_/g, ' '), 
                  Math.round(porcentaje), i + 1, ejecutivosArray.length);
      
      try {
        crearHojaEjecutivo(ss, nombreEjecutivo, ejecutivosPorNombre[nombreEjecutivo], encabezados);
      } catch (e) {
        Logger.log('Error creando ' + nombreEjecutivo + ': ' + e.toString());
      }
      
      Utilities.sleep(200);
    }
    
    // ETAPA 6: Protección (60%)
    setProgreso(6, '🔒 Aplicando protección...', 60, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(500);
    
    try {
      aplicarProteccionTodasLasHojas(ss);
    } catch (e) {
      Logger.log('Error en protección: ' + e.toString());
    }
    
    // ETAPA 7: BBDD_REPORTE (70%)
    setProgreso(7, '🗃️ Generando BBDD_REPORTE...', 70, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(500);
    
    try {
      crearOActualizarReporteAutomatico(ss);
    } catch (e) {
      throw new Error('Error crítico en BBDD_REPORTE: ' + e.toString());
    }
    
    // ETAPA 8: RESUMEN (80%)
    setProgreso(8, '📈 Generando RESUMEN...', 80, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(500);
    
    try {
      generarResumenAutomatico(ss);
    } catch (e) {
      Logger.log('Error en RESUMEN: ' + e.toString());
    }
    
    // ETAPA 9: LLAMADAS (85%)
    setProgreso(9, '📞 Creando LLAMADAS...', 85, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(500);
    
    try {
      crearTablaLlamadas();
    } catch (e) {
      Logger.log('Error en LLAMADAS: ' + e.toString());
    }
    
    // ETAPA 10: PRODUCTIVIDAD (90%)
    setProgreso(10, '📊 Creando PRODUCTIVIDAD...', 90, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(500);
    
    try {
      crearHojaProductividad();
    } catch (e) {
      Logger.log('Error en PRODUCTIVIDAD: ' + e.toString());
    }
    
    // ETAPA 11: Ordenar (95%)
    setProgreso(11, '🗂️ Ordenando hojas...', 95, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(500);
    
    try {
      ordenarHojasPorGrupo();
    } catch (e) {
      Logger.log('Error ordenando: ' + e.toString());
    }
    
    // ETAPA 12: Finalización (100%)
    setProgreso(12, '✅ Proceso completado', 100, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(1500);
    
    // Mensaje de éxito
    var msg = '✅ PROCESAMIENTO EXITOSO\n\n';
    msg += '📊 Ejecutivos: ' + ejecutivosArray.length + '\n';
    msg += '🔒 Protección aplicada\n';
    msg += '📋 BBDD_REPORTE creada\n';
    msg += '📈 RESUMEN generado\n';
    msg += '📞 LLAMADAS creada\n';
    msg += '📊 PRODUCTIVIDAD creada\n';
    msg += '🗂️ Hojas ordenadas';
    
    if (alertasEjecutivos.hojasHuerfanas.length > 0) {
      msg += '\n\n⚠️ Hojas sin ejecutivos: ' + alertasEjecutivos.hojasHuerfanas.length;
    }
    
    if (alertasEjecutivos.ejecutivosNuevos.length > 0) {
      msg += '\n✨ Ejecutivos nuevos: ' + alertasEjecutivos.ejecutivosNuevos.length;
    }
    
    Utilities.sleep(2000);
    ui.alert('✅ Completado', msg, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    
    var mensajeError = 'Error: ' + error.message;
    if (ejecutivosArray.length > 0) {
      mensajeError += '\n\nEjecutivos: ' + ejecutivosArray.length;
    }
    
    setProgreso(0, '❌ ' + mensajeError, 0, 0, ejecutivosArray.length);
    Utilities.sleep(1000);
    
    ui.alert('❌ Error', error.message + '\n\nRevisa los logs', ui.ButtonSet.OK);
  }
}