/**
 * ========================================
 * GENERADOR DE FÓRMULAS CON DETECCIÓN DE IDIOMA
 * ========================================
 * 
 * Detecta automáticamente si Google Sheets está en inglés o español
 * y genera la fórmula correspondiente
 */

/**
 * Detecta el idioma Y separador del spreadsheet mediante pruebas
 * @return {object} {funciones: 'es'|'en', separador: ','|';'}
 */
function detectarConfiguracionCompleta() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getActiveSheet();
    var celdaPrueba = hoja.getRange('ZZ1');
    
    // Probar 4 combinaciones posibles
    var combinaciones = [
      {funciones: 'es', separador: ';', formula: '=SI(VERDADERO;"OK";"ERROR")', nombre: 'Español puro'},
      {funciones: 'en', separador: ';', formula: '=IF(TRUE;"OK";"ERROR")', nombre: 'Inglés con ; (híbrido)'},
      {funciones: 'en', separador: ',', formula: '=IF(TRUE,"OK","ERROR")', nombre: 'Inglés puro'},
      {funciones: 'es', separador: ',', formula: '=SI(VERDADERO,"OK","ERROR")', nombre: 'Español con , (raro)'}
    ];
    
    for (var i = 0; i < combinaciones.length; i++) {
      var config = combinaciones[i];
      
      celdaPrueba.clearContent();
      celdaPrueba.setFormula(config.formula);
      SpreadsheetApp.flush();
      
      var resultado = celdaPrueba.getValue();
      
      Logger.log('Prueba ' + (i+1) + ': ' + config.nombre);
      Logger.log('  Fórmula: ' + config.formula);
      Logger.log('  Resultado: ' + resultado);
      
      if (resultado === 'OK') {
        celdaPrueba.clearContent();
        Logger.log('  ✓ CONFIGURACIÓN DETECTADA: ' + config.nombre);
        return {
          funciones: config.funciones,
          separador: config.separador,
          nombre: config.nombre
        };
      }
    }
    
    // Si ninguna funcionó, usar por defecto
    celdaPrueba.clearContent();
    Logger.log('No se detectó configuración, usando inglés por defecto');
    return {funciones: 'en', separador: ',', nombre: 'Inglés por defecto'};
    
  } catch (error) {
    Logger.log('Error detectando configuración: ' + error.toString());
    return {funciones: 'en', separador: ',', nombre: 'Inglés por defecto (error)'};
  }
}

/**
 * Genera la fórmula de ESTADO_COMPROMISO según la configuración
 * @param {string} colLetra - Letra de la columna FECHA_COMPROMISO
 * @param {number} fila - Número de fila
 * @param {object} config - {funciones: 'es'|'en', separador: ','|';'}
 * @return {string} Fórmula completa
 */
function generarFormulaCompromiso(colLetra, fila, config) {
  var sep = config.separador;
  var q = '"'; // Comillas siempre iguales
  
  if (config.funciones === 'es') {
    // ESPAÑOL: SI, ESBLANCO, HOY
    return '=SI(ESBLANCO(' + colLetra + fila + ')' + sep + q + 'SIN_COMPROMISO' + q + sep + 'SI(' + colLetra + fila + '=HOY()' + sep + q + 'LLAMAR_HOY' + q + sep + 'SI(' + colLetra + fila + '<HOY()' + sep + q + 'COMPROMISO_VENCIDO' + q + sep + q + 'COMPROMISO_FUTURO' + q + ')))';
  } else {
    // INGLÉS: IF, ISBLANK, TODAY
    return '=IF(ISBLANK(' + colLetra + fila + ')' + sep + q + 'SIN_COMPROMISO' + q + sep + 'IF(' + colLetra + fila + '=TODAY()' + sep + q + 'LLAMAR_HOY' + q + sep + 'IF(' + colLetra + fila + '<TODAY()' + sep + q + 'COMPROMISO_VENCIDO' + q + sep + q + 'COMPROMISO_FUTURO' + q + ')))';
  }
}

/**
 * Prueba la detección completa de configuración
 */
function probarDeteccionIdioma() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var locale = ss.getSpreadsheetLocale();
    
    // Detectar configuración completa
    var config = detectarConfiguracionCompleta();
    
    // Generar fórmula de ejemplo
    var formulaEjemplo = generarFormulaCompromiso('O', 2, config);
    
    var mensaje = '🔍 DETECCIÓN DE CONFIGURACIÓN\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Locale: ' + locale + '\n';
    mensaje += 'Configuración: ' + config.nombre + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Funciones: ' + (config.funciones === 'es' ? 'ESPAÑOL (SI, ESBLANCO, HOY)' : 'INGLÉS (IF, ISBLANK, TODAY)') + '\n';
    mensaje += 'Separador: ' + (config.separador === ';' ? 'punto y coma (;)' : 'coma (,)') + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Fórmula que se usará:\n\n';
    mensaje += formulaEjemplo + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += '¿Probar en una celda?';
    
    Logger.log('Locale: ' + locale);
    Logger.log('Configuración: ' + config.nombre);
    Logger.log('Fórmula: ' + formulaEjemplo);
    
    var respuesta = ui.alert('Detección de Configuración', mensaje, ui.ButtonSet.YES_NO);
    
    if (respuesta === ui.Button.YES) {
      aplicarFormulaConDeteccionAuto();
    }
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Aplica la fórmula en la celda seleccionada (con detección automática)
 */
function aplicarFormulaConDeteccionAuto() {
  try {
    var ui = SpreadsheetApp.getUi();
    var hoja = SpreadsheetApp.getActiveSheet();
    var celda = hoja.getActiveCell();
    var fila = celda.getRow();
    
    // Detectar configuración
    var config = detectarConfiguracionCompleta();
    
    // Buscar FECHA_COMPROMISO
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var idxFechaCompromiso = encabezados.indexOf('FECHA_COMPROMISO');
    
    if (idxFechaCompromiso === -1) {
      ui.alert('Error', 'No se encontró la columna FECHA_COMPROMISO', ui.ButtonSet.OK);
      return;
    }
    
    var colLetra = columnNumberToLetter(idxFechaCompromiso + 1);
    
    // Generar fórmula según configuración
    var formula = generarFormulaCompromiso(colLetra, fila, config);
    
    // Limpiar y aplicar
    celda.clearContent();
    celda.clearFormat();
    celda.setNumberFormat('General');
    celda.setFormula(formula);
    SpreadsheetApp.flush();
    
    var resultado = celda.getValue();
    var formulaResultante = celda.getFormula();
    
    var mensaje = '✅ FÓRMULA APLICADA\n\n';
    mensaje += 'Configuración: ' + config.nombre + '\n\n';
    mensaje += 'Celda: ' + celda.getA1Notation() + '\n\n';
    mensaje += 'Fórmula:\n' + formulaResultante + '\n\n';
    mensaje += 'Resultado: ' + resultado + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (resultado && resultado !== '#NAME?' && resultado !== '#ERROR!') {
      mensaje += '✅ ¡FUNCIONA CORRECTAMENTE!\n\n';
      mensaje += '¿Aplicar a TODAS las hojas?';
      
      var confirmar = ui.alert('Éxito', mensaje, ui.ButtonSet.YES_NO);
      
      if (confirmar === ui.Button.YES) {
        repararTodasConDeteccionAuto();
      }
      
    } else {
      mensaje += '❌ Error: ' + resultado + '\n\n';
      mensaje += 'La configuración detectada no funcionó.\n';
      mensaje += 'Por favor reporta este caso.';
      ui.alert('Error en Fórmula', mensaje, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Repara TODAS las hojas con detección automática de configuración
 */
function repararTodasConDeteccionAuto() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Detectar configuración completa
    var config = detectarConfiguracionCompleta();
    
    Logger.log('=== REPARACIÓN CON CONFIGURACIÓN: ' + config.nombre + ' ===');
    
    var hojas = ss.getSheets();
    var reparadas = 0;
    var errores = 0;
    var detalles = [];
    
    for (var i = 0; i < hojas.length; i++) {
      var hoja = hojas[i];
      var nombre = hoja.getName();
      
      // Saltar hojas del sistema
      if (/^BBDD_.*_REMOTO/i.test(nombre)) continue;
      
      var esExcluida = false;
      var hojasExcluidas = ['BBDD_REPORTE', 'RESUMEN', 'LLAMADAS', 'PRODUCTIVIDAD', 'CONFIG_PERFILES'];
      for (var j = 0; j < hojasExcluidas.length; j++) {
        if (nombre.indexOf(hojasExcluidas[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida || hoja.getLastRow() < 2) continue;
      
      try {
        var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
        var idxFechaCompromiso = encabezados.indexOf('FECHA_COMPROMISO');
        var idxEstadoCompromiso = encabezados.indexOf('ESTADO_COMPROMISO');
        
        if (idxFechaCompromiso !== -1 && idxEstadoCompromiso !== -1) {
          var numeroFilas = hoja.getLastRow() - 1;
          var colLetra = columnNumberToLetter(idxFechaCompromiso + 1);
          
          var rangoEstado = hoja.getRange(2, idxEstadoCompromiso + 1, numeroFilas, 1);
          rangoEstado.clearContent();
          rangoEstado.clearFormat();
          rangoEstado.setNumberFormat('General');
          
          var formulas = [];
          for (var k = 2; k <= numeroFilas + 1; k++) {
            var f = generarFormulaCompromiso(colLetra, k, config);
            formulas.push([f]);
          }
          
          rangoEstado.setFormulas(formulas);
          SpreadsheetApp.flush();
          
          reparadas++;
          detalles.push('✅ ' + nombre + ' (' + numeroFilas + ' filas)');
          Logger.log('✓ ' + nombre + ' (' + numeroFilas + ' filas)');
        }
        
      } catch (e) {
        errores++;
        detalles.push('❌ ' + nombre + ': ' + e.message);
        Logger.log('✗ ' + nombre + ': ' + e.message);
      }
    }
    
    var mensaje = '✅ REPARACIÓN COMPLETADA\n\n';
    mensaje += 'Configuración: ' + config.nombre + '\n\n';
    mensaje += 'Hojas reparadas: ' + reparadas + '\n';
    mensaje += 'Errores: ' + errores + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Detalles:\n\n';
    
    for (var m = 0; m < Math.min(detalles.length, 10); m++) {
      mensaje += detalles[m] + '\n';
    }
    
    if (detalles.length > 10) {
      mensaje += '\n... y ' + (detalles.length - 10) + ' más';
    }
    
    Logger.log('=== COMPLETADO ===');
    Logger.log('Reparadas: ' + reparadas);
    Logger.log('Errores: ' + errores);
    
    ui.alert('Completado', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function columnNumberToLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}