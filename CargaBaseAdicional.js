/**
 * MÓDULO 9: CARGA DE BASE ADICIONAL - VERSIÓN CON PERFILAMIENTO
 * Permite cargar y distribuir datos desde un archivo Excel adicional
 * sin afectar la distribución inicial
 * 
 * NUEVAS CARACTERÍSTICAS:
 * - Registro automático de perfiles en CONFIG_PERFILES
 * - Detección de ejecutivos nuevos y existentes
 * - Actualización de roles durante la carga adicional
 */

// Etapas del proceso de carga adicional
const ETAPAS_CARGA_ADICIONAL = [
  { id: 1, nombre: 'Validación', descripcion: 'Verificando archivo...', icono: '🔍', porcentaje: 0 },
  { id: 2, nombre: 'Lectura', descripcion: 'Leyendo datos del archivo...', icono: '📖', porcentaje: 10 },
  { id: 3, nombre: 'Validación Estructura', descripcion: 'Verificando columnas...', icono: '✅', porcentaje: 20 },
  { id: 4, nombre: 'Agrupación', descripcion: 'Agrupando por ejecutivo...', icono: '👥', porcentaje: 30 },
  { id: 5, nombre: 'Preparación', descripcion: 'Preparando distribución...', icono: '⚙️', porcentaje: 40 },
  { id: 6, nombre: 'Distribución', descripcion: 'Distribuyendo datos...', icono: '📊', porcentaje: 50 },
  { id: 7, nombre: 'Limpieza', descripcion: 'Eliminando filas en blanco...', icono: '🧹', porcentaje: 70 },
  { id: 8, nombre: 'BBDD_REPORTE', descripcion: 'Actualizando BBDD_REPORTE...', icono: '📋', porcentaje: 75 },
  { id: 9, nombre: 'RESUMEN', descripcion: 'Actualizando RESUMEN...', icono: '📈', porcentaje: 85 },
  { id: 10, nombre: 'LLAMADAS', descripcion: 'Actualizando LLAMADAS...', icono: '📞', porcentaje: 90 },
  { id: 11, nombre: 'PRODUCTIVIDAD', descripcion: 'Actualizando PRODUCTIVIDAD...', icono: '💼', porcentaje: 93 },
  { id: 12, nombre: 'CONFIG_PERFILES', descripcion: 'Actualizando perfiles...', icono: '👥', porcentaje: 95 },
  { id: 13, nombre: 'Ordenamiento', descripcion: 'Ordenando hojas...', icono: '🗂️', porcentaje: 97 },
  { id: 14, nombre: 'Finalización', descripcion: 'Proceso completado', icono: '✅', porcentaje: 100 }
];

function setProgresoCargaAdicional(etapaId, mensaje, porcentaje, actual, total) {
  var cache = CacheService.getUserCache();
  var etapa = null;
  for (var i = 0; i < ETAPAS_CARGA_ADICIONAL.length; i++) {
    if (ETAPAS_CARGA_ADICIONAL[i].id === etapaId) {
      etapa = ETAPAS_CARGA_ADICIONAL[i];
    }
  }
  var progreso = {
    etapa: etapaId,
    nombreEtapa: etapa ? etapa.nombre : 'Procesando',
    icono: etapa ? etapa.icono : '⚙️',
    mensaje: mensaje,
    porcentaje: porcentaje,
    actual: actual,
    total: total,
    timestamp: new Date().getTime()
  };
  cache.put('progresoCargaAdicional', JSON.stringify(progreso), 600);
  Logger.log('Progreso: ' + porcentaje + '% - ' + mensaje);
}

function getProgresoCargaAdicional() {
  var cache = CacheService.getUserCache();
  var datos = cache.get('progresoCargaAdicional');
  if (datos) {
    return JSON.parse(datos);
  }
  return {
    etapa: 0,
    nombreEtapa: 'Iniciando',
    icono: '⚡',
    mensaje: 'Preparando proceso...',
    porcentaje: 0,
    actual: 0,
    total: 0
  };
}

function mostrarVentanaProgresoCargaAdicional() {
  var html = HtmlService.createHtmlOutputFromFile('VentanaProgresoCargaAdicional')
    .setWidth(620)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '📤 Cargando Base Adicional');
}

function cargarYDistribuirDesdeExcel() {
  try {
    var ui = SpreadsheetApp.getUi();
    var respuesta = ui.alert(
      '📤 Cargar Base Adicional',
      'Por favor, asegúrate de que:\n\n' +
      '1. El archivo Excel esté en tu Google Drive\n' +
      '2. Hayas abierto el Excel con Google Sheets\n' +
      '3. Tenga la misma estructura que la base original\n' +
      '4. Incluya la columna EJECUTIVO\n\n' +
      '¿Deseas continuar?',
      ui.ButtonSet.YES_NO
    );
    if (respuesta !== ui.Button.YES) {
      return;
    }
    var inputResponse = ui.prompt(
      '📁 ID del Archivo',
      'Ingresa el ID o URL del archivo en Google Drive:\n\n' +
      'Ejemplo:\n' +
      'https://docs.google.com/spreadsheets/d/1ABC123.../edit\n' +
      'o solo: 1ABC123...',
      ui.ButtonSet.OK_CANCEL
    );
    if (inputResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    var fileId = extraerFileId(inputResponse.getResponseText());
    if (!fileId) {
      ui.alert('❌ Error', 'ID de archivo inválido. Por favor verifica la URL o ID.', ui.ButtonSet.OK);
      return;
    }
    mostrarVentanaProgresoCargaAdicional();
    procesarCargaAdicional(fileId);
  } catch (error) {
    Logger.log('Error en cargarYDistribuirDesdeExcel: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', 'Error inesperado: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function procesarCargaAdicional(fileId) {
  try {
    setProgresoCargaAdicional(1, 'Validando acceso al archivo...', 0, 0, 1);
    Utilities.sleep(500);
    
    var file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (e) {
      setProgresoCargaAdicional(1, '❌ No se pudo acceder al archivo', 0, 0, 1);
      throw new Error('No se pudo acceder al archivo. Verifica que tengas permisos.');
    }
    
    var spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.open(file);
    } catch (e) {
      setProgresoCargaAdicional(1, '❌ El archivo no es una hoja de cálculo', 0, 0, 1);
      throw new Error('El archivo no es una hoja de cálculo de Google Sheets.');
    }
    
    setProgresoCargaAdicional(2, 'Leyendo datos del archivo...', 10, 0, 1);
    Utilities.sleep(500);
    
    var hojaOrigen = spreadsheet.getSheets()[0];
    if (!hojaOrigen || hojaOrigen.getLastRow() < 2) {
      setProgresoCargaAdicional(2, '❌ El archivo no contiene datos', 10, 0, 1);
      throw new Error('El archivo no contiene datos válidos.');
    }
    
    var datos = hojaOrigen.getDataRange().getValues();
    var encabezados = datos[0];
    var filasDatos = datos.slice(1);
    
    setProgresoCargaAdicional(3, 'Verificando columna EJECUTIVO...', 20, 0, 1);
    Utilities.sleep(500);
    
    var ejecutivoIndex = -1;
    for (var i = 0; i < encabezados.length; i++) {
      var encabezado = encabezados[i].toString().toUpperCase();
      if (encabezado.indexOf('EJECUTIVO') !== -1 || 
          encabezado.indexOf('VENDEDOR') !== -1 || 
          encabezado.indexOf('AGENTE') !== -1) {
        ejecutivoIndex = i;
      }
    }
    
    if (ejecutivoIndex === -1) {
      setProgresoCargaAdicional(3, '❌ No se encontró columna EJECUTIVO', 20, 0, 1);
      throw new Error('No se encontró la columna EJECUTIVO en el archivo.');
    }
    
    setProgresoCargaAdicional(4, 'Agrupando datos por ejecutivo...', 30, 0, filasDatos.length);
    Utilities.sleep(500);
    
    var ejecutivosPorNombre = {};
    var totalRegistros = 0;
    
    for (var j = 0; j < filasDatos.length; j++) {
      if (j % 100 === 0) {
        setProgresoCargaAdicional(4, 'Procesando registro ' + (j + 1) + ' de ' + filasDatos.length, 30, j, filasDatos.length);
      }
      
      var filaVacia = true;
      for (var k = 0; k < filasDatos[j].length; k++) {
        if (filasDatos[j][k] && filasDatos[j][k].toString().trim() !== '') {
          filaVacia = false;
        }
      }
      if (filaVacia) continue;
      
      var nombreEjecutivo = filasDatos[j][ejecutivoIndex];
      if (nombreEjecutivo && nombreEjecutivo.toString().trim() !== '') {
        var nombreFormateado = formatearNombreEjecutivo(nombreEjecutivo.toString());
        if (!ejecutivosPorNombre[nombreFormateado]) {
          ejecutivosPorNombre[nombreFormateado] = [];
        }
        ejecutivosPorNombre[nombreFormateado].push(filasDatos[j]);
        totalRegistros++;
      }
    }
    
    var ejecutivosArray = Object.keys(ejecutivosPorNombre);
    if (ejecutivosArray.length === 0) {
      setProgresoCargaAdicional(4, '❌ No se encontraron ejecutivos', 30, 0, 1);
      throw new Error('No se encontraron ejecutivos válidos en el archivo.');
    }
    
    Logger.log('Ejecutivos encontrados: ' + ejecutivosArray.length);
    for (var e = 0; e < ejecutivosArray.length; e++) {
      Logger.log('  - ' + ejecutivosArray[e] + ': ' + ejecutivosPorNombre[ejecutivosArray[e]].length + ' registros');
    }
    
    setProgresoCargaAdicional(5, 'Preparando distribución de ' + totalRegistros + ' registros...', 40, 0, ejecutivosArray.length);
    Utilities.sleep(500);
    
    setProgresoCargaAdicional(5, 'Verificando duplicados...', 42, 0, 1);
    Utilities.sleep(500);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultadoValidacion = validarDuplicados(ss, ejecutivosPorNombre, encabezados);
    
    if (resultadoValidacion.tieneDuplicados) {
      var mensajeDuplicados = '⚠️ SE ENCONTRARON REGISTROS DUPLICADOS\n\n';
      mensajeDuplicados += 'Total duplicados: ' + resultadoValidacion.totalDuplicados + '\n\n';
      mensajeDuplicados += 'Detalle:\n';
      var detalleResumido = resultadoValidacion.detalle.slice(0, 10);
      for (var d = 0; d < detalleResumido.length; d++) {
        mensajeDuplicados += '• ' + detalleResumido[d] + '\n';
      }
      if (resultadoValidacion.detalle.length > 10) {
        mensajeDuplicados += '\n...y ' + (resultadoValidacion.detalle.length - 10) + ' más\n';
      }
      mensajeDuplicados += '\n¿Deseas continuar ignorando los duplicados?';
      
      var ui = SpreadsheetApp.getUi();
      var respuesta = ui.alert('⚠️ Duplicados Encontrados', mensajeDuplicados, ui.ButtonSet.YES_NO);
      if (respuesta !== ui.Button.YES) {
        setProgresoCargaAdicional(5, '❌ Proceso cancelado por duplicados', 42, 0, 1);
        throw new Error('Proceso cancelado: Se encontraron registros duplicados');
      }
      ejecutivosPorNombre = resultadoValidacion.datosSinDuplicados;
      ejecutivosArray = Object.keys(ejecutivosPorNombre);
      totalRegistros = resultadoValidacion.totalSinDuplicados;
      setProgresoCargaAdicional(5, 'Duplicados eliminados. Continuando con ' + totalRegistros + ' registros únicos...', 45, 0, 1);
      Utilities.sleep(1000);
    } else {
      setProgresoCargaAdicional(5, '✅ No se encontraron duplicados', 45, 0, 1);
      Utilities.sleep(500);
    }
    
    var agregados = 0;
    var nuevos = 0;
    var actualizados = 0;
    var errores = 0;
    
    var hojasExistentes = ss.getSheets();
    var nombresHojas = {};
    for (var h = 0; h < hojasExistentes.length; h++) {
      var nombreHoja = hojasExistentes[h].getName();
      nombresHojas[nombreHoja] = hojasExistentes[h];
    }
    
    Logger.log('=== INICIANDO DISTRIBUCIÓN ===');
    Logger.log('Total de ejecutivos a procesar: ' + ejecutivosArray.length);
    
    for (var n = 0; n < ejecutivosArray.length; n++) {
      var nombreEjecutivo = ejecutivosArray[n];
      var datosEjecutivo = ejecutivosPorNombre[nombreEjecutivo];
      var porcentajeDistribucion = 50 + Math.floor((n / ejecutivosArray.length) * 25);
      setProgresoCargaAdicional(6, 'Distribuyendo datos para: ' + nombreEjecutivo.replace(/_/g, ' '), porcentajeDistribucion, n + 1, ejecutivosArray.length);
      Utilities.sleep(200);
      
      try {
        var hojaEjecutivo = nombresHojas[nombreEjecutivo];
        if (!hojaEjecutivo) {
          var nombreSinGuion = nombreEjecutivo.replace(/_/g, ' ');
          hojaEjecutivo = nombresHojas[nombreSinGuion];
        }
        
        if (hojaEjecutivo) {
          Logger.log('Actualizando hoja existente: ' + nombreEjecutivo);
          var ultimaFila = hojaEjecutivo.getLastRow();
          var encabezadosHoja = hojaEjecutivo.getRange(1, 1, 1, hojaEjecutivo.getLastColumn()).getValues()[0];
          var numColumnasHoja = encabezadosHoja.length;
          var numColsOriginales = determinarColumnasOriginales(encabezadosHoja);
          
          var colEjecutivoEnHoja = -1;
          for (var buscarEjecHoja = 0; buscarEjecHoja < encabezadosHoja.length; buscarEjecHoja++) {
            var encHojaUpper = encabezadosHoja[buscarEjecHoja].toString().toUpperCase();
            if (encHojaUpper.indexOf('EJECUTIVO') !== -1 || 
                encHojaUpper.indexOf('VENDEDOR') !== -1 || 
                encHojaUpper.indexOf('AGENTE') !== -1) {
              colEjecutivoEnHoja = buscarEjecHoja;
              break;
            }
          }
          
          var datosExpandidos = [];
          for (var r = 0; r < datosEjecutivo.length; r++) {
            var fila = datosEjecutivo[r].slice(0, Math.min(datosEjecutivo[r].length, numColsOriginales));
            while (fila.length < numColsOriginales) {
              fila.push('');
            }
            
            if (colEjecutivoEnHoja >= 0 && colEjecutivoEnHoja < fila.length) {
              fila[colEjecutivoEnHoja] = nombreEjecutivo.replace(/_/g, ' ').toUpperCase();
            }
            
            fila = fila.concat([
              '',
              '',
              '',
              'Sin Gestión',
              'Sin Gestión',
              '',
              '',
              ''
            ]);
            datosExpandidos.push(fila);
          }
          hojaEjecutivo.getRange(ultimaFila + 1, 1, datosExpandidos.length, numColumnasHoja).setValues(datosExpandidos);
          aplicarValidacionesYFormulas(hojaEjecutivo, encabezadosHoja, datosExpandidos.length);
          agregados += datosEjecutivo.length;
          actualizados++;
          Logger.log('✓ Agregados ' + datosEjecutivo.length + ' registros a hoja existente');
        } else {
          Logger.log('=== CREANDO NUEVA HOJA PARA EJECUTIVO NUEVO ===');
          Logger.log('Ejecutivo: ' + nombreEjecutivo);
          Logger.log('Número de registros: ' + datosEjecutivo.length);
          
          try {
            crearHojaEjecutivo(ss, nombreEjecutivo, datosEjecutivo, encabezados);
            
            nuevos++;
            agregados += datosEjecutivo.length;
            Logger.log('✅ ' + nombreEjecutivo + ' creada con ' + datosEjecutivo.length + ' registros');
          } catch (crearError) {
            Logger.log('❌ ERROR al crear hoja: ' + crearError.toString());
            throw crearError;
          }
          
          Logger.log('');
        }
      } catch (e) {
        Logger.log('❌ Error procesando ' + nombreEjecutivo + ': ' + e.toString());
        errores++;
      }
    }
    
    Logger.log('=== DISTRIBUCIÓN COMPLETADA ===');
    Logger.log('Registros agregados: ' + agregados);
    Logger.log('Hojas actualizadas: ' + actualizados);
    Logger.log('Hojas nuevas: ' + nuevos);
    Logger.log('Errores: ' + errores);
    
    setProgresoCargaAdicional(7, 'Eliminando filas en blanco...', 70, 0, 1);
    Utilities.sleep(500);
    Logger.log('=== ELIMINANDO FILAS EN BLANCO ===');
    try {
      var resultadoLimpieza = eliminarFilasEnBlancoTodasLasHojas();
      Logger.log('✓ Limpieza completada: ' + resultadoLimpieza.totalFilasEliminadas + ' filas eliminadas');
    } catch (e) {
      Logger.log('❌ Error en limpieza de filas: ' + e.toString());
    }
    
    Logger.log('=== ACTUALIZANDO HOJAS DEL SISTEMA ===');
    
    setProgresoCargaAdicional(8, 'Actualizando BBDD_REPORTE...', 75, 0, 1);
    Utilities.sleep(500);
    try {
      crearOActualizarReporteAutomatico(ss);
      Logger.log('✓ BBDD_REPORTE actualizado');
    } catch (e) {
      Logger.log('❌ Error actualizando BBDD_REPORTE: ' + e.toString());
    }
    
    setProgresoCargaAdicional(9, 'Generando RESUMEN...', 85, 0, 1);
    Utilities.sleep(500);
    try {
      generarResumenAutomatico(ss);
      Logger.log('✓ RESUMEN generado');
    } catch (e) {
      Logger.log('❌ Error generando RESUMEN: ' + e.toString());
    }
    
    setProgresoCargaAdicional(10, 'Actualizando LLAMADAS...', 90, 0, 1);
    Utilities.sleep(500);
    try {
      crearTablaLlamadas();
      Logger.log('✓ LLAMADAS actualizada');
    } catch (e) {
      Logger.log('❌ Error actualizando LLAMADAS: ' + e.toString());
    }
    
    setProgresoCargaAdicional(11, 'Actualizando PRODUCTIVIDAD...', 93, 0, 1);
    Utilities.sleep(500);
    try {
      crearHojaProductividad();
      Logger.log('✓ PRODUCTIVIDAD actualizada');
    } catch (e) {
      Logger.log('❌ Error actualizando PRODUCTIVIDAD: ' + e.toString());
    }
    
    setProgresoCargaAdicional(12, 'Actualizando CONFIG_PERFILES...', 95, 0, 1);
    Utilities.sleep(500);
    Logger.log('=== REGISTRANDO PERFILES EN CONFIG_PERFILES ===');
    try {
      var hojasActuales = ss.getSheets();
      var ejecutivosCreados = [];
      
      for (var k = 0; k < hojasActuales.length; k++) {
        var nombreHoja = hojasActuales[k].getName();
        
        var esExcluida = false;
        for (var m = 0; m < HOJAS_EXCLUIDAS.length; m++) {
          if (nombreHoja.indexOf(HOJAS_EXCLUIDAS[m]) !== -1) {
            esExcluida = true;
            break;
          }
        }
        
        if (esExcluida) continue;
        if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) continue;
        
        if (hojasActuales[k].getLastRow() > 1) {
          ejecutivosCreados.push(nombreHoja);
        }
      }
      
      Logger.log('Hojas de ejecutivos detectadas: ' + ejecutivosCreados.length);
      
      if (ejecutivosCreados.length > 0) {
        var resultadoPerfiles = registrarEjecutivosEnConfig(ejecutivosCreados);
        Logger.log('✓ Perfiles registrados: ' + resultadoPerfiles.nuevos + ' nuevos, ' + 
                   resultadoPerfiles.actualizados + ' actualizados');
        
        // Limpiar cualquier entrada de CONFIG_PERFILES
        limpiarConfigPerfilesDeListaEjecutivos();
        
        // Ocultar la hoja CONFIG_PERFILES
        ocultarConfigPerfiles();
      }
    } catch (errorPerfil) {
      Logger.log('⚠️ Error registrando perfiles (no crítico): ' + errorPerfil.toString());
    }
    
    setProgresoCargaAdicional(13, 'Ordenando hojas...', 97, 0, 1);
    Utilities.sleep(500);
    try {
      ordenarHojasPorGrupo();
      Logger.log('✓ Hojas ordenadas');
    } catch (e) {
      Logger.log('❌ Error ordenando hojas: ' + e.toString());
    }
    
    Logger.log('=== PROCESO COMPLETADO EXITOSAMENTE ===');
    var mensajeFinal = '✅ COMPLETADO\n\n';
    mensajeFinal += '📊 Registros: ' + agregados + '\n';
    mensajeFinal += '👥 Actualizados: ' + actualizados + '\n';
    mensajeFinal += '✨ Nuevos: ' + nuevos + '\n';
    mensajeFinal += '🔄 CONFIG_PERFILES actualizado';
    if (errores > 0) {
      mensajeFinal += '\n⚠️ Errores: ' + errores;
    }
    setProgresoCargaAdicional(14, mensajeFinal, 100, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(2000);
  } catch (error) {
    Logger.log('❌ Error en procesarCargaAdicional: ' + error.toString());
    setProgresoCargaAdicional(1, '❌ Error: ' + error.message, 0, 0, 1);
    throw error;
  }
}

function validarDuplicados(ss, ejecutivosPorNombre, encabezados) {
  var resultado = {
    tieneDuplicados: false,
    totalDuplicados: 0,
    detalle: [],
    datosSinDuplicados: {},
    totalSinDuplicados: 0
  };
  var colIdentificacion = -1;
  var nombreColIdentificacion = '';
  var columnasId = ['RUT', 'RUT_CLIENTE', 'ID', 'IDENTIFICACION', 'DNI', 'CEDULA'];
  for (var i = 0; i < encabezados.length; i++) {
    var enc = encabezados[i].toString().toUpperCase().trim();
    for (var j = 0; j < columnasId.length; j++) {
      if (enc.indexOf(columnasId[j]) !== -1) {
        colIdentificacion = i;
        nombreColIdentificacion = enc;
        break;
      }
    }
    if (colIdentificacion !== -1) break;
  }
  if (colIdentificacion === -1) {
    Logger.log('No se encontró columna de identificación para validar duplicados');
    resultado.datosSinDuplicados = ejecutivosPorNombre;
    resultado.totalSinDuplicados = contarTotalRegistros(ejecutivosPorNombre);
    return resultado;
  }
  Logger.log('Validando duplicados usando columna: ' + nombreColIdentificacion);
  var rutosExistentes = {};
  var hojas = ss.getSheets();
  for (var h = 0; h < hojas.length; h++) {
    var hoja = hojas[h];
    var nombreHoja = hoja.getName();
    if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) continue;
    var esExcluida = false;
    for (var x = 0; x < HOJAS_EXCLUIDAS.length; x++) {
      if (nombreHoja.indexOf(HOJAS_EXCLUIDAS[x]) !== -1) {
        esExcluida = true;
        break;
      }
    }
    if (esExcluida || hoja.getLastRow() <= 1) continue;
    try {
      var encHoja = hoja.getRange(1, 1, 1, Math.min(hoja.getLastColumn(), 20)).getValues()[0];
      var colIdEnHoja = -1;
      for (var y = 0; y < encHoja.length; y++) {
        if (encHoja[y].toString().toUpperCase().trim() === nombreColIdentificacion) {
          colIdEnHoja = y;
          break;
        }
      }
      if (colIdEnHoja !== -1 && hoja.getLastRow() > 1) {
        var datosHoja = hoja.getRange(2, colIdEnHoja + 1, hoja.getLastRow() - 1, 1).getValues();
        for (var z = 0; z < datosHoja.length; z++) {
          var rut = limpiarRut(datosHoja[z][0]);
          if (rut) {
            rutosExistentes[rut] = true;
          }
        }
      }
    } catch (errorHoja) {
      Logger.log('Error leyendo hoja ' + nombreHoja + ': ' + errorHoja.toString());
    }
  }
  Logger.log('RUTs existentes en el sistema: ' + Object.keys(rutosExistentes).length);
  var ejecutivosLimpios = {};
  var totalDuplicados = 0;
  var detalleDuplicados = [];
  for (var ejecutivo in ejecutivosPorNombre) {
    ejecutivosLimpios[ejecutivo] = [];
    var datos = ejecutivosPorNombre[ejecutivo];
    for (var d = 0; d < datos.length; d++) {
      var rutNuevo = limpiarRut(datos[d][colIdentificacion]);
      if (rutNuevo && rutosExistentes[rutNuevo]) {
        totalDuplicados++;
        detalleDuplicados.push(ejecutivo + ': RUT ' + rutNuevo);
      } else {
        ejecutivosLimpios[ejecutivo].push(datos[d]);
        if (rutNuevo) {
          rutosExistentes[rutNuevo] = true;
        }
      }
    }
  }
  resultado.tieneDuplicados = totalDuplicados > 0;
  resultado.totalDuplicados = totalDuplicados;
  resultado.detalle = detalleDuplicados;
  resultado.datosSinDuplicados = ejecutivosLimpios;
  resultado.totalSinDuplicados = contarTotalRegistros(ejecutivosLimpios);
  Logger.log('Validación completada:');
  Logger.log('- Duplicados encontrados: ' + resultado.totalDuplicados);
  Logger.log('- Registros únicos: ' + resultado.totalSinDuplicados);
  return resultado;
}

function limpiarRut(rut) {
  if (!rut) return '';
  return rut.toString().toUpperCase().replace(/[^0-9K]/g, '');
}

function contarTotalRegistros(ejecutivosPorNombre) {
  var total = 0;
  for (var ejecutivo in ejecutivosPorNombre) {
    total += ejecutivosPorNombre[ejecutivo].length;
  }
  return total;
}

function determinarColumnasOriginales(encabezados) {
  var numColsOriginales = encabezados.length;
  for (var i = 0; i < encabezados.length; i++) {
    var enc = encabezados[i].toString().trim();
    for (var j = 0; j < COLUMNAS_NUEVAS.length; j++) {
      if (enc === COLUMNAS_NUEVAS[j]) {
        return i;
      }
    }
  }
  return numColsOriginales;
}

function buscarHojaBBDD(ss) {
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    var nombre = hojas[i].getName();
    if (/^BBDD_.*_REMOTO/i.test(nombre)) {
      return hojas[i];
    }
  }
  return null;
}

function extraerFileId(input) {
  if (!input) return null;
  input = input.trim();
  if (input.indexOf('drive.google.com') !== -1 || input.indexOf('docs.google.com') !== -1) {
    var match = input.match(/[-\w]{25,}/);
    return match ? match[0] : null;
  }
  if (input.length >= 25 && input.match(/^[-\w]+$/)) {
    return input;
  }
  return null;
}

function eliminarFilasEnBlancoTodasLasHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ss.getSheets();
  var totalFilasEliminadas = 0;
  var hojasLimpiadas = 0;
  
  Logger.log('=== INICIANDO LIMPIEZA DE FILAS EN BLANCO ===');
  
  for (var i = 0; i < hojas.length; i++) {
    var hoja = hojas[i];
    var nombreHoja = hoja.getName();
    
    var esExcluida = false;
    var hojasExcluidas = ['BBDD_REPORTE', 'RESUMEN', 'LLAMADAS', 'PRODUCTIVIDAD', 'CONFIG_PERFILES', 'CONFIGURACION', 'PLANTILLA'];
    for (var j = 0; j < hojasExcluidas.length; j++) {
      if (nombreHoja.indexOf(hojasExcluidas[j]) !== -1) {
        esExcluida = true;
        break;
      }
    }
    
    if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) {
      esExcluida = true;
    }
    
    if (esExcluida) continue;
    
    if (hoja.getLastRow() > 1) {
      try {
        var ultimaFila = hoja.getLastRow();
        var datos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
        var filasAEliminar = [];
        
        for (var fila = 0; fila < datos.length; fila++) {
          var filaVacia = true;
          for (var col = 0; col < datos[fila].length; col++) {
            if (datos[fila][col] !== '' && datos[fila][col] !== null) {
              filaVacia = false;
              break;
            }
          }
          if (filaVacia) {
            filasAEliminar.push(fila + 2);
          }
        }
        
        if (filasAEliminar.length > 0) {
          for (var idx = filasAEliminar.length - 1; idx >= 0; idx--) {
            hoja.deleteRow(filasAEliminar[idx]);
          }
          totalFilasEliminadas += filasAEliminar.length;
          hojasLimpiadas++;
          Logger.log('✓ ' + nombreHoja + ': ' + filasAEliminar.length + ' filas eliminadas');
        }
      } catch (errorHoja) {
        Logger.log('⚠️ Error limpiando ' + nombreHoja + ': ' + errorHoja.toString());
      }
    }
  }
  
  Logger.log('=== LIMPIEZA COMPLETADA ===');
  Logger.log('Total filas eliminadas: ' + totalFilasEliminadas);
  Logger.log('Hojas limpiadas: ' + hojasLimpiadas);
  
  return {
    totalFilasEliminadas: totalFilasEliminadas,
    hojasLimpiadas: hojasLimpiadas
  };
}

function limpiarFilasEnBlancoManual() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.alert(
    '🧹 Limpiar Filas en Blanco',
    '¿Deseas eliminar todas las filas completamente vacías de las hojas de ejecutivos?\n\n' +
    'Esta acción:\n' +
    '✓ Eliminará filas totalmente vacías\n' +
    '✓ Mantendrá todas las filas con datos\n' +
    '✓ No afectará BBDD_REPORTE ni hojas del sistema\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta !== ui.Button.YES) {
    return;
  }
  
  ui.alert('🔄 Procesando...', 'Eliminando filas en blanco. Esto puede tardar unos momentos.', ui.ButtonSet.OK);
  
  try {
    var resultado = eliminarFilasEnBlancoTodasLasHojas();
    
    var mensaje = '✅ Limpieza completada\n\n';
    mensaje += '🗑️ Filas eliminadas: ' + resultado.totalFilasEliminadas + '\n';
    mensaje += '📋 Hojas procesadas: ' + resultado.hojasLimpiadas;
    
    ui.alert('✅ Completado', mensaje, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Ocurrió un error durante la limpieza:\n\n' + error.message, ui.ButtonSet.OK);
  }
}