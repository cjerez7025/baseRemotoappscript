/**
 * MÓDULO 5: PROTECCIÓN DE DATOS
 * Funciones para proteger columnas originales
 */

/**
 * Protege columnas originales - NADIE puede editar
 */
function protegerColumnasOriginales(hoja, numCols) {
  try {
    if (!hoja || typeof hoja.getName !== 'function') {
      Logger.log('ERROR: Hoja inválida');
      return false;
    }
    
    var nombreHoja = hoja.getName();
    
    if (!numCols || numCols <= 0) {
      Logger.log('ERROR: numCols inválido');
      return false;
    }
    
    var ultima = hoja.getLastRow();
    if (ultima < 2) {
      Logger.log('SKIP: Sin datos suficientes');
      return false;
    }
    
    var rango = hoja.getRange(1, 1, ultima, numCols);
    var rangoNotacion = rango.getA1Notation();
    
    // Eliminar protecciones previas del mismo rango
    var protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = protecciones.length - 1; i >= 0; i--) {
      try {
        var rangoProtegido = protecciones[i].getRange();
        if (rangoProtegido && rangoProtegido.getA1Notation() === rangoNotacion) {
          protecciones[i].remove();
          Logger.log('Protección previa eliminada');
        }
      } catch (e) {
        Logger.log('Error verificando protección: ' + e.toString());
      }
    }
    
    // Crear nueva protección
    var proteccion = rango.protect();
    if (!proteccion) {
      Logger.log('ERROR: No se pudo crear protección');
      return false;
    }
    
    proteccion.setDescription('🔒 DATOS ORIGINALES - NO EDITAR - PROTEGIDO');
    proteccion.setWarningOnly(false);
    
    // Eliminar TODOS los editores
    var editores = proteccion.getEditors();
    if (editores && editores.length > 0) {
      proteccion.removeEditors(editores);
    }
    
    Logger.log('Protección aplicada: ' + nombreHoja + ' - ' + rangoNotacion);
    return true;
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    return false;
  }
}

/**
 * Aplica protección a todas las hojas de ejecutivos
 */
function aplicarProteccionTodasLasHojas(ss) {
  try {
    var hojas = ss.getSheets();
    var protegidas = 0;
    var saltadas = 0;
    var errores = 0;
    
    Logger.log('===== PROTECCIÓN DE HOJAS =====');
    
    for (var i = 0; i < hojas.length; i++) {
      var hoja = hojas[i];
      var nombre = hoja.getName();
      
      // Saltar hojas especiales
      if (/^BBDD_.*_REMOTO/i.test(nombre)) {
        saltadas++;
        continue;
      }
      
      var esExcluida = false;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.toUpperCase().indexOf(HOJAS_EXCLUIDAS[j].toUpperCase()) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida || hoja.getLastRow() < 2) {
        saltadas++;
        continue;
      }
      
      try {
        var ultimaCol = hoja.getLastColumn();
        var encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
        var numColsOriginales = ultimaCol;
        
        // Buscar primera columna nueva
        for (var k = 0; k < encabezados.length; k++) {
          var encabezado = encabezados[k] ? encabezados[k].toString().trim() : '';
          
          for (var m = 0; m < COLUMNAS_NUEVAS.length; m++) {
            if (encabezado === COLUMNAS_NUEVAS[m]) {
              numColsOriginales = k;
              break;
            }
          }
          
          if (numColsOriginales < ultimaCol) break;
        }
        
        if (numColsOriginales > 0) {
          if (protegerColumnasOriginales(hoja, numColsOriginales)) {
            protegidas++;
          } else {
            errores++;
          }
        }
        
      } catch (error) {
        Logger.log('Error en ' + nombre + ': ' + error.toString());
        errores++;
      }
    }
    
    Logger.log('Protegidas: ' + protegidas + ', Saltadas: ' + saltadas + ', Errores: ' + errores);
    
    return {
      protegidas: protegidas,
      saltadas: saltadas,
      errores: errores,
      total: hojas.length
    };
    
  } catch (error) {
    Logger.log('Error crítico: ' + error.toString());
    throw error;
  }
}

/**
 * Ejecuta protección en la hoja actual
 */
function ejecutarProteccionHojaActual() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getActiveSheet();
    
    if (!hoja) {
      SpreadsheetApp.getUi().alert('Error', 'No hay hoja activa', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var nombreHoja = hoja.getName();
    
    // Verificar que no sea hoja especial
    if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) {
      SpreadsheetApp.getUi().alert('Error', 'No se puede proteger la hoja REMOTO', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    for (var i = 0; i < HOJAS_EXCLUIDAS.length; i++) {
      if (nombreHoja.toUpperCase().indexOf(HOJAS_EXCLUIDAS[i]) !== -1) {
        SpreadsheetApp.getUi().alert('Error', 'Esta hoja no debe ser protegida', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
    }
    
    var ultimaCol = hoja.getLastColumn();
    if (ultimaCol === 0) {
      SpreadsheetApp.getUi().alert('Error', 'La hoja no tiene columnas', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
    var numColsOriginales = ultimaCol;
    
    for (var k = 0; k < encabezados.length; k++) {
      var encabezado = encabezados[k] ? encabezados[k].toString().trim() : '';
      for (var m = 0; m < COLUMNAS_NUEVAS.length; m++) {
        if (encabezado === COLUMNAS_NUEVAS[m]) {
          numColsOriginales = k;
          break;
        }
      }
      if (numColsOriginales < ultimaCol) break;
    }
    
    if (numColsOriginales === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No se detectaron columnas originales', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (protegerColumnasOriginales(hoja, numColsOriginales)) {
      SpreadsheetApp.getUi().alert('✅ Éxito', 'Protección aplicada:\nHoja: ' + nombreHoja + '\nColumnas: A-' + columnNumberToLetter(numColsOriginales), SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('❌ Error', 'No se pudo aplicar la protección', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Verifica el estado de protección de la hoja actual
 */
function verificarProteccion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getActiveSheet();
  var nombre = hoja.getName();
  var protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  
  if (protecciones.length === 0) {
    SpreadsheetApp.getUi().alert('Sin protección', 'Esta hoja NO tiene protecciones', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var mensaje = 'Hoja: ' + nombre + '\nProtecciones: ' + protecciones.length + '\n\n';
  
  for (var k = 0; k < protecciones.length; k++) {
    var prot = protecciones[k];
    mensaje += 'Protección ' + (k + 1) + ':\n';
    mensaje += 'Rango: ' + prot.getRange().getA1Notation() + '\n';
    mensaje += 'Solo advertencia: ' + (prot.isWarningOnly() ? 'SÍ' : 'NO') + '\n';
    mensaje += 'Editores: ' + prot.getEditors().length + '\n\n';
  }
  
  SpreadsheetApp.getUi().alert('Estado de Protección', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Elimina todas las protecciones de la hoja actual
 */
function eliminarTodasLasProtecciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getActiveSheet();
  var nombre = hoja.getName();
  var ui = SpreadsheetApp.getUi();
  
  var respuesta = ui.alert('Confirmar', '¿Eliminar TODAS las protecciones de "' + nombre + '"?', ui.ButtonSet.YES_NO);
  
  if (respuesta !== ui.Button.YES) return;
  
  var protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var eliminadas = 0;
  
  for (var i = 0; i < protecciones.length; i++) {
    protecciones[i].remove();
    eliminadas++;
  }
  
  ui.alert('Completado', 'Se eliminaron ' + eliminadas + ' protecciones', ui.ButtonSet.OK);
}