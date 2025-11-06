/**
 * ========================================
 * GENERADOR DE FÓRMULAS CON DETECCIÓN DE IDIOMA
 * ========================================
 * 
 * Detecta automáticamente si Google Sheets está en inglés o español
 * y genera la fórmula correspondiente
 */

/**
 * Detecta el idioma del spreadsheet
 * @return {string} 'es' o 'en'
 */
function detectarIdiomaSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var locale = ss.getSpreadsheetLocale();
    
    Logger.log('Locale detectado: ' + locale);
    
    // Español: es, es_ES, es_MX, es_CL, etc.
    if (locale.indexOf('es') === 0) {
      return 'es';
    }
    
    // Por defecto: inglés
    return 'en';
    
  } catch (error) {
    Logger.log('Error detectando idioma: ' + error.toString());
    // Si falla, intentar con una fórmula de prueba
    return detectarIdiomaPorPrueba();
  }
}

/**
 * Detecta idioma mediante prueba de fórmula
 * @return {string} 'es' o 'en'
 */
function detectarIdiomaPorPrueba() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getActiveSheet();
    
    // Crear una celda temporal
    var celdaPrueba = hoja.getRange('ZZ1');
    
    // Probar fórmula en español
    celdaPrueba.setFormula('=SI(VERDADERO;"OK";"ERROR")');
    SpreadsheetApp.flush();
    
    var resultado = celdaPrueba.getValue();
    celdaPrueba.clearContent();
    
    if (resultado === 'OK') {
      Logger.log('Idioma detectado por prueba: ESPAÑOL');
      return 'es';
    }
    
    Logger.log('Idioma detectado por prueba: INGLÉS');
    return 'en';
    
  } catch (error) {
    Logger.log('Error en prueba de idioma: ' + error.toString());
    return 'en'; // Por defecto inglés
  }
}

/**
 * Genera la fórmula de ESTADO_COMPROMISO según el idioma
 * @param {string} colLetra - Letra de la columna FECHA_COMPROMISO
 * @param {number} fila - Número de fila
 * @param {string} idioma - 'es' o 'en'
 * @return {string} Fórmula completa
 */
function generarFormulaCompromiso(colLetra, fila, idioma) {
  if (idioma === 'es') {
    // ESPAÑOL: SI, ESBLANCO, HOY, punto y coma
    return '=SI(ESBLANCO(' + colLetra + fila + ');"SIN_COMPROMISO";SI(' + colLetra + fila + '=HOY();"LLAMAR_HOY";SI(' + colLetra + fila + '<HOY();"COMPROMISO_VENCIDO";"COMPROMISO_FUTURO")))';
  } else {
    // INGLÉS: IF, ISBLANK, TODAY, coma
    return '=IF(ISBLANK(' + colLetra + fila + '),"SIN_COMPROMISO",IF(' + colLetra + fila + '=TODAY(),"LLAMAR_HOY",IF(' + colLetra + fila + '<TODAY(),"COMPROMISO_VENCIDO","COMPROMISO_FUTURO")))';
  }
}

/**
 * Prueba la detección de idioma y genera una fórmula
 */
function probarDeteccionIdioma() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Detectar idioma
    var idioma = detectarIdiomaSpreadsheet();
    var locale = ss.getSpreadsheetLocale();
    
    // Generar fórmula de ejemplo
    var formulaEjemplo = generarFormulaCompromiso('O', 2, idioma);
    
    var mensaje = '🔍 DETECCIÓN DE IDIOMA\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Locale: ' + locale + '\n';
    mensaje += 'Idioma detectado: ' + (idioma === 'es' ? 'ESPAÑOL' : 'INGLÉS') + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Fórmula que se usará:\n\n';
    mensaje += formulaEjemplo + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (idioma === 'es') {
      mensaje += '✅ Funciones: SI, ESBLANCO, HOY\n';
      mensaje += '✅ Separador: ; (punto y coma)';
    } else {
      mensaje += '✅ Funciones: IF, ISBLANK, TODAY\n';
      mensaje += '✅ Separador: , (coma)';
    }
    
    Logger.log('Locale: ' + locale);
    Logger.log('Idioma: ' + idioma);
    Logger.log('Fórmula: ' + formulaEjemplo);
    
    ui.alert('Detección de Idioma', mensaje, ui.ButtonSet.OK);
    
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
    
    // Detectar idioma
    var idioma = detectarIdiomaSpreadsheet();
    
    // Buscar FECHA_COMPROMISO
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var idxFechaCompromiso = encabezados.indexOf('FECHA_COMPROMISO');
    
    if (idxFechaCompromiso === -1) {
      ui.alert('Error', 'No se encontró la columna FECHA_COMPROMISO', ui.ButtonSet.OK);
      return;
    }
    
    var colLetra = columnNumberToLetter(idxFechaCompromiso + 1);
    
    // Generar fórmula según idioma
    var formula = generarFormulaCompromiso(colLetra, fila, idioma);
    
    // Limpiar y aplicar
    celda.clearContent();
    celda.clearFormat();
    celda.setNumberFormat('General');
    celda.setFormula(formula);
    SpreadsheetApp.flush();
    
    var resultado = celda.getValue();
    
    var mensaje = '✅ FÓRMULA APLICADA\n\n';
    mensaje += 'Idioma: ' + (idioma === 'es' ? 'ESPAÑOL' : 'INGLÉS') + '\n\n';
    mensaje += 'Celda: ' + celda.getA1Notation() + '\n';
    mensaje += 'Resultado: ' + resultado + '\n\n';
    
    if (resultado && resultado !== '#NAME?' && resultado !== '#ERROR!') {
      mensaje += '✅ Funciona correctamente\n\n';
      mensaje += '¿Aplicar a TODAS las hojas?';
      
      var confirmar = ui.alert('Éxito', mensaje, ui.ButtonSet.YES_NO);
      
      if (confirmar === ui.Button.YES) {
        repararTodasConDeteccionAuto();
      }
      
    } else {
      mensaje += '❌ Error: ' + resultado + '\n\n';
      mensaje += 'Por favor reporta este error.';
      ui.alert('Error en Fórmula', mensaje, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Repara TODAS las hojas con detección automática de idioma
 */
function repararTodasConDeteccionAuto() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Detectar idioma
    var idioma = detectarIdiomaSpreadsheet();
    
    Logger.log('=== REPARACIÓN CON IDIOMA: ' + idioma + ' ===');
    
    var hojas = ss.getSheets();
    var reparadas = 0;
    var errores = 0;
    
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
            var f = generarFormulaCompromiso(colLetra, k, idioma);
            formulas.push([f]);
          }
          
          rangoEstado.setFormulas(formulas);
          SpreadsheetApp.flush();
          
          reparadas++;
          Logger.log('✓ ' + nombre + ' (' + numeroFilas + ' filas)');
        }
        
      } catch (e) {
        errores++;
        Logger.log('✗ ' + nombre + ': ' + e.message);
      }
    }
    
    var mensaje = '✅ REPARACIÓN COMPLETADA\n\n';
    mensaje += 'Idioma usado: ' + (idioma === 'es' ? 'ESPAÑOL' : 'INGLÉS') + '\n\n';
    mensaje += 'Hojas reparadas: ' + reparadas + '\n';
    mensaje += 'Errores: ' + errores;
    
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