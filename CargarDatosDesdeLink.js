/**
 * MÓDULO: CARGA INICIAL DESDE LINK
 * Copia datos desde Google Sheets externo y ejecuta distribución completa
 * VERSIÓN MEJORADA: Incluye eliminación robusta de hojas de ejecutivos
 */

// Etapas del proceso de carga inicial
const ETAPAS_CARGA_INICIAL = [
  { id: 1, nombre: 'Validación', descripcion: 'Verificando archivo...', icono: '🔍', porcentaje: 0 },
  { id: 2, nombre: 'Lectura', descripcion: 'Leyendo datos del archivo...', icono: '📖', porcentaje: 10 },
  { id: 3, nombre: 'Búsqueda BBDD', descripcion: 'Buscando hoja BBDD_*_REMOTO*...', icono: '🎯', porcentaje: 20 },
  { id: 4, nombre: 'Limpieza', descripcion: 'Limpiando datos anteriores...', icono: '🧹', porcentaje: 30 },
  { id: 5, nombre: 'Copia', descripcion: 'Copiando datos a BBDD...', icono: '📋', porcentaje: 40 },
  { id: 6, nombre: 'Formato', descripcion: 'Aplicando formato...', icono: '🎨', porcentaje: 50 },
  { id: 7, nombre: 'Eliminación Hojas', descripcion: 'Eliminando hojas anteriores...', icono: '🗑️', porcentaje: 55 },
  { id: 8, nombre: 'Distribución', descripcion: 'Ejecutando distribución...', icono: '🚀', porcentaje: 60 },
  { id: 9, nombre: 'Finalización', descripcion: 'Proceso completado', icono: '✅', porcentaje: 100 }
];

/**
 * Actualiza el progreso de carga inicial
 */
function setProgresoCargaInicial(etapaId, mensaje, porcentaje, actual, total) {
  var cache = CacheService.getUserCache();
  
  var etapa = null;
  for (var i = 0; i < ETAPAS_CARGA_INICIAL.length; i++) {
    if (ETAPAS_CARGA_INICIAL[i].id === etapaId) {
      etapa = ETAPAS_CARGA_INICIAL[i];
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
  
  cache.put('progresoCargaInicial', JSON.stringify(progreso), 600);
  Logger.log('Progreso: ' + porcentaje + '% - ' + mensaje);
}

/**
 * Obtiene el progreso actual de carga inicial
 */
function getProgresoCargaInicial() {
  var cache = CacheService.getUserCache();
  var datos = cache.get('progresoCargaInicial');
  
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
function mostrarVentanaProgresoCargaInicial() {
  var html = HtmlService.createHtmlOutputFromFile('VentanaProgresoCargaInicial')
    .setWidth(620)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 Carga Inicial de Datos');
}

/**
 * Carga datos desde un Google Sheets externo a la hoja BBDD_*_REMOTO*
 * y ejecuta la distribución automáticamente
 */
function cargarDatosDesdeLink() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Confirmar acción
    var respuesta = ui.alert(
      '📥 Carga Inicial (Copiar y Distribuir)',
      'Esta función:\n\n' +
      '1. Copiará datos desde un Google Sheets externo\n' +
      '2. Los pegará en tu hoja BBDD_*_REMOTO*\n' +
      '3. Ejecutará la distribución completa\n\n' +
      '⚠️ Los datos actuales en BBDD serán reemplazados\n\n' +
      '¿Deseas continuar?',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) {
      return;
    }
    
    // Solicitar link
    var inputResponse = ui.prompt(
      '🔗 Link del Google Sheets',
      'Ingresa el ID o URL del Google Sheets:\n\n' +
      'Ejemplo:\n' +
      'https://docs.google.com/spreadsheets/d/1ABC123.../edit',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (inputResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    var fileId = extraerFileId(inputResponse.getResponseText());
    
    if (!fileId) {
      ui.alert('❌ Error', 'ID de archivo inválido.', ui.ButtonSet.OK);
      return;
    }
    
    // Mostrar ventana de progreso
    mostrarVentanaProgresoCargaInicial();
    
    // Ejecutar proceso con progreso
    procesarCargaInicialConProgreso(fileId);
    
  } catch (error) {
    Logger.log('Error en cargarDatosDesdeLink: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', 'Error inesperado: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Procesa la carga inicial con ventana de progreso
 */
function procesarCargaInicialConProgreso(fileId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ETAPA 1: Validación (0%)
    setProgresoCargaInicial(1, 'Validando acceso al archivo...', 0, 0, 1);
    Utilities.sleep(500);
    
    var file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (e) {
      setProgresoCargaInicial(1, '❌ No se pudo acceder al archivo', 0, 0, 1);
      throw new Error('No se pudo acceder al archivo. Verifica permisos.');
    }
    
    var spreadsheetExterno;
    try {
      spreadsheetExterno = SpreadsheetApp.open(file);
    } catch (e) {
      setProgresoCargaInicial(1, '❌ El archivo no es una hoja de cálculo', 0, 0, 1);
      throw new Error('El archivo no es una hoja de cálculo de Google Sheets.');
    }
    
    // ETAPA 2: Lectura (10%)
    setProgresoCargaInicial(2, 'Leyendo datos del archivo externo...', 10, 0, 1);
    Utilities.sleep(500);
    
    var hojaOrigen = spreadsheetExterno.getSheets()[0];
    
    if (!hojaOrigen || hojaOrigen.getLastRow() < 2) {
      setProgresoCargaInicial(2, '❌ El archivo no contiene datos', 10, 0, 1);
      throw new Error('El archivo no contiene datos válidos.');
    }
    
    var ultimaFilaOrigen = hojaOrigen.getLastRow();
    var ultimaColOrigen = hojaOrigen.getLastColumn();
    var datosCompletos = hojaOrigen.getDataRange().getValues();
    
    Logger.log('Datos leídos: ' + ultimaFilaOrigen + ' filas x ' + ultimaColOrigen + ' columnas');
    
    // ETAPA 3: Buscar hoja BBDD (20%)
    setProgresoCargaInicial(3, 'Buscando hoja BBDD_*_REMOTO*...', 20, 0, 1);
    Utilities.sleep(500);
    
    var hojaBBDD = buscarHojaBBDD(ss);
    
    if (!hojaBBDD) {
      setProgresoCargaInicial(3, '❌ No se encontró hoja BBDD_*_REMOTO*', 20, 0, 1);
      throw new Error('No se encontró ninguna hoja con patrón BBDD_*_REMOTO*');
    }
    
    var nombreBBDD = hojaBBDD.getName();
    Logger.log('Hoja BBDD encontrada: ' + nombreBBDD);
    
    // ETAPA 4: Limpieza (30%)
    setProgresoCargaInicial(4, 'Limpiando datos anteriores de ' + nombreBBDD + '...', 30, 0, 1);
    Utilities.sleep(500);
    
    var ultimaFilaBBDD = hojaBBDD.getLastRow();
    
    // Eliminar filtro existente si hay
    var filtroExistente = hojaBBDD.getFilter();
    if (filtroExistente) {
      filtroExistente.remove();
      Logger.log('Filtro existente eliminado');
    }
    
    if (ultimaFilaBBDD > 1) {
      hojaBBDD.getRange(2, 1, ultimaFilaBBDD - 1, hojaBBDD.getLastColumn()).clear();
    }
    
    Logger.log('Datos anteriores limpiados');
    
    // ETAPA 5: Copia (40%)
    var encabezados = datosCompletos[0];
    var datosSinEncabezado = datosCompletos.slice(1);
    
    setProgresoCargaInicial(5, 'Copiando ' + datosSinEncabezado.length + ' registros...', 40, 0, datosSinEncabezado.length);
    Utilities.sleep(500);
    
    // Copiar encabezados
    hojaBBDD.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    
    // Copiar datos
    if (datosSinEncabezado.length > 0) {
      hojaBBDD.getRange(2, 1, datosSinEncabezado.length, encabezados.length).setValues(datosSinEncabezado);
    }
    
    Logger.log('Datos copiados: ' + datosSinEncabezado.length + ' registros');
    
    // ETAPA 6: Formato (50%)
    setProgresoCargaInicial(6, 'Aplicando formato a ' + nombreBBDD + '...', 50, 0, 1);
    Utilities.sleep(500);
    
    hojaBBDD.getRange(1, 1, 1, encabezados.length)
      .setBackground(COLORES.HEADER_REPORTE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setFontColor('white');
    
    hojaBBDD.setFrozenRows(1);
    hojaBBDD.autoResizeColumns(1, encabezados.length);
    
    // Aplicar filtro
    if (datosSinEncabezado.length > 0) {
      var rangoTotal = hojaBBDD.getRange(1, 1, datosSinEncabezado.length + 1, encabezados.length);
      rangoTotal.createFilter();
    }
    
    Logger.log('Formato aplicado');
    
    // ETAPA 7: Eliminación de hojas anteriores (55%)
    setProgresoCargaInicial(7, 'Eliminando hojas de ejecutivos anteriores...', 55, 0, 1);
    Utilities.sleep(500);
    
    try {
      var hojasEliminadas = eliminarHojasEjecutivosAnteriores(ss);
      Logger.log('Hojas eliminadas: ' + hojasEliminadas);
    } catch (e) {
      Logger.log('Error eliminando hojas: ' + e.toString());
      // No es crítico, continuar
    }
    
    // ETAPA 8: Distribución (60%)
    setProgresoCargaInicial(8, 'Ejecutando distribución por ejecutivos...', 60, 0, 1);
    Utilities.sleep(1000);
    
    // Cerrar ventana de carga inicial antes de la distribución
    var htmlCerrar = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(htmlCerrar, 'Cerrando...');
    Utilities.sleep(500);
    
    try {
      procesarEjecutivos();
      
      Logger.log('=== PROCESO COMPLETADO ===');
      
      // Mensaje final
      var ui = SpreadsheetApp.getUi();
      ui.alert('✅ Proceso Completado', 
        '📊 Datos copiados y distribuidos exitosamente\n\n' +
        'Revisa las hojas generadas:\n' +
        '• BBDD_REPORTE\n' +
        '• RESUMEN\n' +
        '• LLAMADAS\n' +
        '• PRODUCTIVIDAD', 
        ui.ButtonSet.OK);
      
    } catch (e) {
      Logger.log('Error en distribución: ' + e.toString());
      throw new Error('Error al distribuir: ' + e.message);
    }
    
  } catch (error) {
    Logger.log('Error en procesarCargaInicialConProgreso: ' + error.toString());
    setProgresoCargaInicial(1, '❌ Error: ' + error.message, 0, 0, 1);
    throw error;
  }
}

/**
 * Elimina todas las hojas de ejecutivos anteriores
 * VERSIÓN MEJORADA: Elimina TODAS las hojas que no sean especiales
 */
function eliminarHojasEjecutivosAnteriores(ss) {
  try {
    var hojas = ss.getSheets();
    var hojasAEliminar = [];
    
    Logger.log('=== INICIANDO ELIMINACIÓN DE HOJAS ===');
    Logger.log('Total de hojas en el spreadsheet: ' + hojas.length);
    
    // Lista de hojas que NO se deben eliminar
    var hojasProtegidas = [
      /^BBDD_.*_REMOTO/i,  // Hoja base de datos remota
    ];
    
    // Identificar hojas a eliminar
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      
      // 1. Verificar si es hoja protegida (BBDD_*_REMOTO*)
      var esProtegida = false;
      for (var p = 0; p < hojasProtegidas.length; p++) {
        if (hojasProtegidas[p].test(nombre)) {
          esProtegida = true;
          Logger.log('✓ Saltando (protegida): ' + nombre);
          break;
        }
      }
      
      if (esProtegida) {
        continue;
      }
      
      // 2. Verificar si está en lista de hojas excluidas (RESUMEN, LLAMADAS, etc.)
      var esExcluida = false;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.toUpperCase().indexOf(HOJAS_EXCLUIDAS[j].toUpperCase()) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida) {
        Logger.log('✓ Saltando (excluida): ' + nombre);
        continue;
      }
      
      // 3. Si llegó aquí, es una hoja de ejecutivo (o desconocida) → ELIMINAR
      // No importa si tiene datos o no, se elimina
      Logger.log('🗑️ Marcada para eliminar: ' + nombre);
      hojasAEliminar.push(hojas[i]);
    }
    
    // Eliminar hojas identificadas
    Logger.log('=== ELIMINANDO HOJAS ===');
    Logger.log('Total hojas a eliminar: ' + hojasAEliminar.length);
    
    if (hojasAEliminar.length === 0) {
      Logger.log('No hay hojas de ejecutivos para eliminar');
      return 0;
    }
    
    var eliminadas = 0;
    for (var n = 0; n < hojasAEliminar.length; n++) {
      try {
        var nombreHoja = hojasAEliminar[n].getName();
        ss.deleteSheet(hojasAEliminar[n]);
        eliminadas++;
        Logger.log('✅ Eliminada: ' + nombreHoja);
      } catch (e) {
        Logger.log('❌ Error eliminando ' + nombreHoja + ': ' + e.toString());
      }
    }
    
    Logger.log('=== RESUMEN ===');
    Logger.log('Hojas eliminadas exitosamente: ' + eliminadas + '/' + hojasAEliminar.length);
    
    return eliminadas;
    
  } catch (error) {
    Logger.log('ERROR CRÍTICO en eliminarHojasEjecutivosAnteriores: ' + error.toString());
    throw error;
  }
}

/**
 * Busca la hoja BBDD_*_REMOTO*
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
  
  // Si es una URL
  if (input.indexOf('drive.google.com') !== -1 || input.indexOf('docs.google.com') !== -1) {
    var match = input.match(/[-\w]{25,}/);
    return match ? match[0] : null;
  }
  
  // Si es un ID directo
  if (input.length >= 25 && input.match(/^[-\w]+$/)) {
    return input;
  }
  
  return null;
}