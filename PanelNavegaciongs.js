/**
 * MÓDULO: PANEL DE NAVEGACIÓN
 * Panel lateral para navegación rápida entre hojas
 * Versión: 1.0
 */

/**
 * Muestra el panel de navegación lateral
 */
function mostrarPanelNavegacion() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('PanelNavegacion')
      .setTitle('🗂️ Navegación')
      .setWidth(340);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
  } catch (error) {
    Logger.log('Error mostrando panel: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'No se pudo abrir el panel de navegación.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Obtiene la lista de hojas disponibles organizadas por tipo
 * OPTIMIZADO: Procesamiento más rápido
 * @return {Object} Objeto con arrays de hojas por categoría
 */
function obtenerHojasDisponibles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    
    var resultado = {
      baseDatos: [],
      gestion: [],
      ejecutivos: [],
      otras: []
    };
    
    // Hojas excluidas del listado
    var hojasExcluidas = ['Sheet1', 'Hoja 1', 'Hoja1'];
    
    // OPTIMIZACIÓN: Procesar solo una vez
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      
      // Saltar hojas excluidas
      if (hojasExcluidas.indexOf(nombre) !== -1) {
        continue;
      }
      
      // Clasificar hojas (sin verificar estructura interna para más velocidad)
      if (nombre.indexOf('BBDD') !== -1) {
        resultado.baseDatos.push(nombre);
      } else if (nombre === 'RESUMEN' || nombre === 'PRODUCTIVIDAD' || nombre === 'LLAMADAS') {
        resultado.gestion.push(nombre);
      } else if (nombre.indexOf('BBDD') === -1 && 
                 nombre !== 'RESUMEN' && 
                 nombre !== 'PRODUCTIVIDAD' && 
                 nombre !== 'LLAMADAS' &&
                 nombre !== 'CONFIGURACION') {
        // Asumir que es hoja de ejecutivo si no es ninguna de las anteriores
        resultado.ejecutivos.push(nombre);
      } else {
        resultado.otras.push(nombre);
      }
    }
    
    // Ordenar ejecutivos alfabéticamente
    resultado.ejecutivos.sort();
    
    Logger.log('Hojas cargadas rápidamente - Ejecutivos: ' + resultado.ejecutivos.length);
    return resultado;
    
  } catch (error) {
    Logger.log('Error obteniendo hojas: ' + error.toString());
    return {
      baseDatos: [],
      gestion: [],
      ejecutivos: [],
      otras: []
    };
  }
}

/**
 * Determina si una hoja es de un ejecutivo
 * @param {Sheet} hoja - La hoja a verificar
 * @return {boolean} true si es hoja de ejecutivo
 */
function esHojaEjecutivo(hoja) {
  try {
    // Verificar que tenga al menos 2 filas
    if (hoja.getLastRow() < 2) {
      return false;
    }
    
    // Obtener encabezados (máximo 20 columnas para optimizar)
    var numCols = Math.min(hoja.getLastColumn(), 20);
    var encabezados = hoja.getRange(1, 1, 1, numCols).getValues()[0];
    
    // Columnas características de hojas de ejecutivo
    var columnasRequeridas = ['FECHA_LLAMADA', 'ESTADO', 'SUB_ESTADO', 'NOTA_EJECUTIVO'];
    
    var encontradas = 0;
    for (var i = 0; i < columnasRequeridas.length; i++) {
      for (var j = 0; j < encabezados.length; j++) {
        if (encabezados[j] && encabezados[j].toString().toUpperCase() === columnasRequeridas[i]) {
          encontradas++;
          break;
        }
      }
    }
    
    // Debe tener al menos 2 de las columnas requeridas
    return encontradas >= 2;
    
  } catch (error) {
    Logger.log('Error verificando hoja: ' + error.toString());
    return false;
  }
}

/**
 * Activa (navega a) una hoja específica
 * @param {string} nombreHoja - Nombre de la hoja a activar
 */
function activarHoja(nombreHoja) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(nombreHoja);
    
    if (hoja) {
      ss.setActiveSheet(hoja);
      Logger.log('Navegado a: ' + nombreHoja);
      return true;
    } else {
      Logger.log('Hoja no encontrada: ' + nombreHoja);
      return false;
    }
    
  } catch (error) {
    Logger.log('Error activando hoja: ' + error.toString());
    return false;
  }
}

/**
 * Verifica si el usuario actual es supervisor
 * @return {boolean} true si es supervisor
 */
function esUsuarioSupervisor() {
  try {
    var email = Session.getActiveUser().getEmail();
    
    // Lista de emails de supervisores (CONFIGURAR AQUÍ)
    var supervisores = [
      'supervisor1@empresa.com',
      'supervisor2@empresa.com',
      'admin@empresa.com'
    ];
    
    // Verificar si el email está en la lista
    var esSupervisor = supervisores.indexOf(email.toLowerCase()) !== -1;
    
    Logger.log('Usuario: ' + email + ' - Es supervisor: ' + esSupervisor);
    return esSupervisor;
    
  } catch (error) {
    Logger.log('Error verificando permisos: ' + error.toString());
    return false; // Por defecto, no es supervisor
  }
}

/**
 * Ejecuta diagnóstico de hojas
 */
function diagnosticarHojas() {
  try {
    // Verificar si existe la función de diagnóstico
    if (typeof verificarHojasEjecutivos === 'function') {
      verificarHojasEjecutivos();
      SpreadsheetApp.getUi().alert('✓', 'Diagnóstico completado. Revisa el registro (Logs).', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('⚠️', 'Función de diagnóstico no disponible.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('Error en diagnóstico: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'No se pudo ejecutar el diagnóstico.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Ordena las hojas automáticamente
 */
function ordenarHojasAutomaticamente() {
  try {
    // Verificar si existe la función de ordenamiento
    if (typeof ordenarHojasAutomaticamente2024 === 'function') {
      ordenarHojasAutomaticamente2024();
      SpreadsheetApp.getUi().alert('✓', 'Hojas ordenadas correctamente.', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('⚠️', 'Función de ordenamiento no disponible.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('Error ordenando hojas: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'No se pudieron ordenar las hojas.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// NOTA: La función onOpen() está en Menu.gs
// Este archivo NO tiene onOpen() para evitar conflictos