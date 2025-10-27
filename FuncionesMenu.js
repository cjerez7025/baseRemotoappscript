/**
 * MÓDULO 8: FUNCIONES LLAMADAS DESDE EL MENÚ
 * Funciones auxiliares que se invocan desde Menu.gs
 */

/**
 * Actualiza las validaciones en todas las hojas de ejecutivos
 */
function actualizarValidaciones() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var actualizadas = 0;
    
    for (var i = 0; i < hojas.length; i++) {
      var hoja = hojas[i];
      var nombre = hoja.getName();
      
      // Saltar hojas especiales
      if (/^BBDD_.*_REMOTO/i.test(nombre)) continue;
      
      var esExcluida = false;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.indexOf(HOJAS_EXCLUIDAS[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida || hoja.getLastRow() < 2) continue;
      
      try {
        var ultimaCol = hoja.getLastColumn();
        var encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
        var numeroFilas = hoja.getLastRow() - 1;
        
        aplicarValidacionesYFormulas(hoja, encabezados, numeroFilas);
        actualizadas++;
        
      } catch (e) {
        Logger.log('Error en ' + nombre + ': ' + e.toString());
      }
    }
    
    Logger.log('Validaciones actualizadas en ' + actualizadas + ' hojas');
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Aplica fórmulas de ESTADO_COMPROMISO en todas las hojas
 */
function aplicarFormulasEstadoCompromiso() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var actualizadas = 0;
    
    for (var i = 0; i < hojas.length; i++) {
      var hoja = hojas[i];
      var nombre = hoja.getName();
      
      if (/^BBDD_.*_REMOTO/i.test(nombre)) continue;
      
      var esExcluida = false;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.indexOf(HOJAS_EXCLUIDAS[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida || hoja.getLastRow() < 2) continue;
      
      try {
        var ultimaCol = hoja.getLastColumn();
        var encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
        
        var idxFechaCompromiso = encabezados.indexOf('FECHA_COMPROMISO');
        var idxEstadoCompromiso = encabezados.indexOf('ESTADO_COMPROMISO');
        
        if (idxFechaCompromiso !== -1 && idxEstadoCompromiso !== -1) {
          var numeroFilas = hoja.getLastRow() - 1;
          var col = columnNumberToLetter(idxFechaCompromiso + 1);
          var formulas = [];
          
          for (var k = 2; k <= numeroFilas + 1; k++) {
            formulas.push([
              '=IF(ISBLANK(' + col + k + '),"SIN_COMPROMISO",IF(' + col + k + '=TODAY(),"LLAMAR_HOY",IF(' + col + k + '<TODAY(),"COMPROMISO_VENCIDO","COMPROMISO_FUTURO")))'
            ]);
          }
          
          hoja.getRange(2, idxEstadoCompromiso + 1, numeroFilas, 1).setFormulas(formulas);
          actualizadas++;
        }
        
      } catch (e) {
        Logger.log('Error en ' + nombre + ': ' + e.toString());
      }
    }
    
    Logger.log('Fórmulas aplicadas en ' + actualizadas + ' hojas');
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Verifica la estructura de las hojas de ejecutivos
 */
function verificarHojasEjecutivos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var reporte = {
      total: 0,
      validas: 0,
      invalidas: 0,
      detalles: []
    };
    
    for (var i = 0; i < hojas.length; i++) {
      var hoja = hojas[i];
      var nombre = hoja.getName();
      
      if (/^BBDD_.*_REMOTO/i.test(nombre)) continue;
      
      var esExcluida = false;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.indexOf(HOJAS_EXCLUIDAS[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida) continue;
      
      reporte.total++;
      
      try {
        if (hoja.getLastRow() < 2) {
          reporte.invalidas++;
          reporte.detalles.push(nombre + ': Sin datos');
          continue;
        }
        
        var ultimaCol = hoja.getLastColumn();
        var encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
        
        var columnasRequeridas = ['FECHA_LLAMADA', 'ESTADO', 'SUB_ESTADO'];
        var faltantes = [];
        
        for (var k = 0; k < columnasRequeridas.length; k++) {
          if (encabezados.indexOf(columnasRequeridas[k]) === -1) {
            faltantes.push(columnasRequeridas[k]);
          }
        }
        
        if (faltantes.length > 0) {
          reporte.invalidas++;
          reporte.detalles.push(nombre + ': Faltan columnas - ' + faltantes.join(', '));
        } else {
          reporte.validas++;
        }
        
      } catch (e) {
        reporte.invalidas++;
        reporte.detalles.push(nombre + ': Error - ' + e.toString());
      }
    }
    
    var mensaje = 'VERIFICACIÓN DE HOJAS\n\n';
    mensaje += 'Total hojas: ' + reporte.total + '\n';
    mensaje += 'Válidas: ' + reporte.validas + '\n';
    mensaje += 'Con problemas: ' + reporte.invalidas + '\n\n';
    
    if (reporte.detalles.length > 0) {
      mensaje += 'DETALLES:\n';
      for (var m = 0; m < Math.min(10, reporte.detalles.length); m++) {
        mensaje += '• ' + reporte.detalles[m] + '\n';
      }
      if (reporte.detalles.length > 10) {
        mensaje += '...y ' + (reporte.detalles.length - 10) + ' más';
      }
    }
    
    SpreadsheetApp.getUi().alert('Verificación Completada', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
    
    Logger.log('Verificación: ' + reporte.validas + '/' + reporte.total + ' hojas válidas');
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Alias de crearOActualizarReporteAutomatico para compatibilidad con menú
 */
function crearHojaReporte() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    crearOActualizarReporteAutomatico(ss);
    Logger.log('Hoja BBDD_REPORTE creada');
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Alias de crearOActualizarReporteAutomatico para compatibilidad con menú
 */
function actualizarReporte() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    crearOActualizarReporteAutomatico(ss);
    Logger.log('Reporte actualizado');
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Limpia todas las hojas de ejecutivos (elimina datos de gestión)
 */
function limpiarHojasEjecutivos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var limpiadas = 0;
    
    for (var i = 0; i < hojas.length; i++) {
      var hoja = hojas[i];
      var nombre = hoja.getName();
      
      if (/^BBDD_.*_REMOTO/i.test(nombre)) continue;
      
      var esExcluida = false;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.indexOf(HOJAS_EXCLUIDAS[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida) continue;
      
      try {
        var ultimaCol = hoja.getLastColumn();
        var encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
        
        // Buscar columnas de gestión
        var columnasFechaLlamada = encabezados.indexOf('FECHA_LLAMADA');
        
        if (columnasFechaLlamada !== -1 && hoja.getLastRow() > 1) {
          var numeroFilas = hoja.getLastRow() - 1;
          var numColsGestion = ultimaCol - columnasFechaLlamada;
          
          // Limpiar solo las columnas de gestión
          var rango = hoja.getRange(2, columnasFechaLlamada + 1, numeroFilas, numColsGestion);
          rango.clearContent();
          
          limpiadas++;
        }
        
      } catch (e) {
        Logger.log('Error limpiando ' + nombre + ': ' + e.toString());
      }
    }
    
    Logger.log('Hojas limpiadas: ' + limpiadas);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}