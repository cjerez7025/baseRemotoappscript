/**
 * MÓDULO 9: CARGA DE BASE ADICIONAL
 * Permite cargar y distribuir datos desde un archivo Excel adicional
 * sin afectar la distribución inicial
 */

// Etapas del proceso de carga adicional
const ETAPAS_CARGA_ADICIONAL = [
  { id: 1, nombre: 'Validación', descripcion: 'Verificando archivo...', icono: '🔍', porcentaje: 0 },
  { id: 2, nombre: 'Lectura', descripcion: 'Leyendo datos del archivo...', icono: '📖', porcentaje: 10 },
  { id: 3, nombre: 'Validación Estructura', descripcion: 'Verificando columnas...', icono: '✅', porcentaje: 20 },
  { id: 4, nombre: 'Agrupación', descripcion: 'Agrupando por ejecutivo...', icono: '👥', porcentaje: 30 },
  { id: 5, nombre: 'Preparación', descripcion: 'Preparando distribución...', icono: '⚙️', porcentaje: 40 },
  { id: 6, nombre: 'Distribución', descripcion: 'Distribuyendo datos...', icono: '📊', porcentaje: 50 },
  { id: 7, nombre: 'BBDD_REPORTE', descripcion: 'Actualizando BBDD_REPORTE...', icono: '📋', porcentaje: 75 },
  { id: 8, nombre: 'RESUMEN', descripcion: 'Actualizando RESUMEN...', icono: '📈', porcentaje: 85 },
  { id: 9, nombre: 'LLAMADAS', descripcion: 'Actualizando LLAMADAS...', icono: '📞', porcentaje: 90 },
  { id: 10, nombre: 'PRODUCTIVIDAD', descripcion: 'Actualizando PRODUCTIVIDAD...', icono: '💼', porcentaje: 95 },
  { id: 11, nombre: 'Finalización', descripcion: 'Proceso completado', icono: '✅', porcentaje: 100 }
];

/**
 * Actualiza el progreso en la ventana modal
 */
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

/**
 * Obtiene el progreso actual
 */
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

/**
 * Muestra la ventana de progreso
 */
function mostrarVentanaProgresoCargaAdicional() {
  var html = HtmlService.createHtmlOutputFromFile('VentanaProgresoCargaAdicional')
    .setWidth(620)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '📤 Cargando Base Adicional');
}

/**
 * Carga y distribuye datos desde un archivo Excel subido
 */
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

/**
 * Procesa la carga adicional de datos
 */
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
    
    setProgresoCargaAdicional(5, 'Preparando distribución de ' + totalRegistros + ' registros...', 40, 0, ejecutivosArray.length);
    Utilities.sleep(500);
    
    // VALIDAR DUPLICADOS
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
      
      // Eliminar duplicados del objeto ejecutivosPorNombre
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
          var ultimaFila = hojaEjecutivo.getLastRow();
          var encabezadosHoja = hojaEjecutivo.getRange(1, 1, 1, hojaEjecutivo.getLastColumn()).getValues()[0];
          var numColumnasHoja = encabezadosHoja.length;
          
          var numColsOriginales = determinarColumnasOriginales(encabezadosHoja);
          
          var datosExpandidos = [];
          for (var r = 0; r < datosEjecutivo.length; r++) {
            var fila = datosEjecutivo[r].slice(0, Math.min(datosEjecutivo[r].length, numColsOriginales));
            
            while (fila.length < numColsOriginales) {
              fila.push('');
            }
            
            fila = fila.concat(['', '', '', 'Sin Gestión', 'Sin Gestión', '', '', '']);
            datosExpandidos.push(fila);
          }
          
          hojaEjecutivo.getRange(ultimaFila + 1, 1, datosExpandidos.length, numColumnasHoja).setValues(datosExpandidos);
          aplicarValidacionesYFormulas(hojaEjecutivo, encabezadosHoja, datosExpandidos.length);
          
          agregados += datosEjecutivo.length;
          actualizados++;
        } else {
          Logger.log('Creando nueva hoja para: ' + nombreEjecutivo);
          
          var hojaBBDD = buscarHojaBBDD(ss);
          var encabezadosCompletos = [];
          
          if (hojaBBDD) {
            var numColsBBDD = hojaBBDD.getLastColumn();
            encabezadosCompletos = hojaBBDD.getRange(1, 1, 1, numColsBBDD).getValues()[0];
          } else {
            encabezadosCompletos = encabezados.concat(COLUMNAS_NUEVAS);
          }
          
          var nuevaHoja = ss.insertSheet(nombreEjecutivo);
          
          nuevaHoja.getRange(1, 1, 1, encabezadosCompletos.length)
            .setValues([encabezadosCompletos])
            .setBackground(COLORES.HEADER_REPORTE)
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setFontColor('white');
          
          var numColsOriginales = encabezados.length;
          var datosExpandidos = [];
          
          for (var r = 0; r < datosEjecutivo.length; r++) {
            var fila = datosEjecutivo[r].slice(0, numColsOriginales);
            
            while (fila.length < numColsOriginales) {
              fila.push('');
            }
            
            fila = fila.concat(['', '', '', 'Sin Gestión', 'Sin Gestión', '', '', '']);
            datosExpandidos.push(fila);
          }
          
          if (datosExpandidos.length > 0) {
            nuevaHoja.getRange(2, 1, datosExpandidos.length, encabezadosCompletos.length).setValues(datosExpandidos);
          }
          
          aplicarValidacionesYFormulas(nuevaHoja, encabezadosCompletos, datosExpandidos.length);
          
          nuevaHoja.setFrozenRows(1);
          nuevaHoja.autoResizeColumns(1, encabezadosCompletos.length);
          
          if (datosExpandidos.length > 0) {
            nuevaHoja.getRange(1, 1, datosExpandidos.length + 1, encabezadosCompletos.length).createFilter();
          }
          
          nuevos++;
          agregados += datosEjecutivo.length;
        }
      } catch (e) {
        Logger.log('Error procesando ' + nombreEjecutivo + ': ' + e.toString());
        errores++;
      }
    }
    
    setProgresoCargaAdicional(7, 'Actualizando BBDD_REPORTE...', 75, 0, 1);
    Utilities.sleep(500);
    
    try {
      crearOActualizarReporteAutomatico(ss);
    } catch (e) {
      Logger.log('Error actualizando BBDD_REPORTE: ' + e.toString());
    }
    
    setProgresoCargaAdicional(8, 'Generando RESUMEN...', 85, 0, 1);
    Utilities.sleep(500);
    
    try {
      generarResumenAutomatico(ss);
    } catch (e) {
      Logger.log('Error generando RESUMEN: ' + e.toString());
    }
    
    setProgresoCargaAdicional(9, 'Actualizando LLAMADAS...', 90, 0, 1);
    Utilities.sleep(500);
    
    try {
      crearTablaLlamadas();
    } catch (e) {
      Logger.log('Error actualizando LLAMADAS: ' + e.toString());
    }
    
    setProgresoCargaAdicional(10, 'Actualizando PRODUCTIVIDAD...', 95, 0, 1);
    Utilities.sleep(500);
    
    try {
      crearHojaProductividad();
    } catch (e) {
      Logger.log('Error actualizando PRODUCTIVIDAD: ' + e.toString());
    }
    
    var mensajeFinal = '✅ COMPLETADO\n\n';
    mensajeFinal += '📊 Registros: ' + agregados + '\n';
    mensajeFinal += '👥 Actualizados: ' + actualizados + '\n';
    mensajeFinal += '✨ Nuevos: ' + nuevos;
    
    if (errores > 0) {
      mensajeFinal += '\n⚠️ Errores: ' + errores;
    }
    
    setProgresoCargaAdicional(11, mensajeFinal, 100, ejecutivosArray.length, ejecutivosArray.length);
    Utilities.sleep(2000);
    
  } catch (error) {
    Logger.log('Error en procesarCargaAdicional: ' + error.toString());
    setProgresoCargaAdicional(1, '❌ Error: ' + error.message, 0, 0, 1);
    throw error;
  }
}

/**
 * Valida si hay registros duplicados comparando con datos existentes
 */
function validarDuplicados(ss, ejecutivosPorNombre, encabezados) {
  var resultado = {
    tieneDuplicados: false,
    totalDuplicados: 0,
    detalle: [],
    datosSinDuplicados: {},
    totalSinDuplicados: 0
  };
  
  // Buscar columna de identificación (RUT o ID)
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
  
  // Obtener todos los RUTs/IDs existentes en las hojas de ejecutivos
  var rutosExistentes = {};
  var hojas = ss.getSheets();
  
  for (var h = 0; h < hojas.length; h++) {
    var hoja = hojas[h];
    var nombreHoja = hoja.getName();
    
    // Saltar hojas especiales
    if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) continue;
    
    var esExcluida = false;
    for (var k = 0; k < HOJAS_EXCLUIDAS.length; k++) {
      if (nombreHoja.indexOf(HOJAS_EXCLUIDAS[k]) !== -1) {
        esExcluida = true;
        break;
      }
    }
    
    if (esExcluida || hoja.getLastRow() < 2) continue;
    
    try {
      var encHoja = hoja.getRange(1, 1, 1, Math.min(hoja.getLastColumn(), 30)).getValues()[0];
      var colRutHoja = -1;
      
      for (var m = 0; m < encHoja.length; m++) {
        var encStr = encHoja[m].toString().toUpperCase().trim();
        if (encStr === nombreColIdentificacion) {
          colRutHoja = m;
          break;
        }
      }
      
      if (colRutHoja === -1) continue;
      
      var numFilas = hoja.getLastRow() - 1;
      if (numFilas > 0) {
        var ruts = hoja.getRange(2, colRutHoja + 1, numFilas, 1).getValues();
        
        for (var n = 0; n < ruts.length; n++) {
          if (ruts[n][0] && ruts[n][0].toString().trim() !== '') {
            var rutLimpio = limpiarRut(ruts[n][0].toString());
            rutosExistentes[rutLimpio] = nombreHoja;
          }
        }
      }
    } catch (e) {
      Logger.log('Error leyendo hoja ' + nombreHoja + ': ' + e.toString());
    }
  }
  
  Logger.log('RUTs/IDs existentes encontrados: ' + Object.keys(rutosExistentes).length);
  
  // Validar duplicados en los datos nuevos
  var ejecutivosLimpios = {};
  
  for (var ejecutivo in ejecutivosPorNombre) {
    var datosEjecutivo = ejecutivosPorNombre[ejecutivo];
    var datosSinDuplicados = [];
    
    for (var p = 0; p < datosEjecutivo.length; p++) {
      var fila = datosEjecutivo[p];
      var identificacion = fila[colIdentificacion];
      
      if (identificacion && identificacion.toString().trim() !== '') {
        var idLimpio = limpiarRut(identificacion.toString());
        
        if (rutosExistentes[idLimpio]) {
          resultado.tieneDuplicados = true;
          resultado.totalDuplicados++;
          resultado.detalle.push(
            identificacion + ' (Ejecutivo: ' + ejecutivo.replace(/_/g, ' ') + 
            ' - Ya existe en: ' + rutosExistentes[idLimpio] + ')'
          );
        } else {
          datosSinDuplicados.push(fila);
          rutosExistentes[idLimpio] = ejecutivo;
        }
      } else {
        datosSinDuplicados.push(fila);
      }
    }
    
    if (datosSinDuplicados.length > 0) {
      ejecutivosLimpios[ejecutivo] = datosSinDuplicados;
    }
  }
  
  resultado.datosSinDuplicados = ejecutivosLimpios;
  resultado.totalSinDuplicados = contarTotalRegistros(ejecutivosLimpios);
  
  Logger.log('Validación completada:');
  Logger.log('- Duplicados encontrados: ' + resultado.totalDuplicados);
  Logger.log('- Registros únicos: ' + resultado.totalSinDuplicados);
  
  return resultado;
}

/**
 * Limpia RUT o identificación para comparación
 */
function limpiarRut(rut) {
  if (!rut) return '';
  return rut.toString().toUpperCase().replace(/[^0-9K]/g, '');
}

/**
 * Cuenta el total de registros en el objeto ejecutivosPorNombre
 */
function contarTotalRegistros(ejecutivosPorNombre) {
  var total = 0;
  for (var ejecutivo in ejecutivosPorNombre) {
    total += ejecutivosPorNombre[ejecutivo].length;
  }
  return total;
}

/**
 * Determina el número de columnas originales (antes de las de gestión)
 */
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

/**
 * Busca la hoja BBDD_*_REMOTO* para obtener estructura
 */
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

/**
 * Extrae el ID del archivo desde URL o ID directo
 */
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