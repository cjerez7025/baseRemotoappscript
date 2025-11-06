/**
 * ========================================
 * PRUEBA DE FÓRMULA - DIAGNÓSTICO
 * ========================================
 * 
 * Función para probar la fórmula en UNA celda
 * antes de aplicarla masivamente
 */

/**
 * Prueba la fórmula en la celda actualmente seleccionada
 */
function probarFormulaEnCeldaActual() {
  try {
    var ui = SpreadsheetApp.getUi();
    var hoja = SpreadsheetApp.getActiveSheet();
    var celda = hoja.getActiveCell();
    var fila = celda.getRow();
    var col = celda.getColumn();
    
    // Buscar FECHA_COMPROMISO
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var idxFechaCompromiso = encabezados.indexOf('FECHA_COMPROMISO');
    
    if (idxFechaCompromiso === -1) {
      ui.alert('Error', 'No se encontró la columna FECHA_COMPROMISO en esta hoja', ui.ButtonSet.OK);
      return;
    }
    
    var colLetra = columnNumberToLetter(idxFechaCompromiso + 1);
    
    // Generar fórmula con caracteres ASCII básicos
    var formula = '=SI(ESBLANCO(' + colLetra + fila + ');"SIN_COMPROMISO";SI(' + colLetra + fila + '=HOY();"LLAMAR_HOY";SI(' + colLetra + fila + '<HOY();"COMPROMISO_VENCIDO";"COMPROMISO_FUTURO")))';
    
    // Limpiar celda
    celda.clearContent();
    celda.clearFormat();
    celda.setNumberFormat('General');
    
    // Aplicar fórmula
    celda.setFormula(formula);
    SpreadsheetApp.flush();
    
    // Verificar resultado
    var valor = celda.getValue();
    var mostrarFormula = celda.getFormula();
    
    var mensaje = '✅ FÓRMULA APLICADA\n\n';
    mensaje += 'Celda: ' + celda.getA1Notation() + '\n\n';
    mensaje += 'Fórmula generada:\n' + mostrarFormula + '\n\n';
    mensaje += 'Resultado: ' + valor + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (valor && valor !== '#ERROR!' && valor !== '#NAME?') {
      mensaje += '✅ La fórmula funciona correctamente\n\n';
      mensaje += 'Puedes aplicarla a todas las hojas.';
    } else {
      mensaje += '❌ La fórmula tiene un error\n\n';
      mensaje += 'Error: ' + valor;
    }
    
    Logger.log('Fórmula generada: ' + formula);
    Logger.log('Fórmula en celda: ' + mostrarFormula);
    Logger.log('Resultado: ' + valor);
    
    ui.alert('Prueba de Fórmula', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error en prueba: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Función auxiliar: Convierte número de columna a letra
 */
function columnNumberToLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * Muestra los códigos ASCII de los caracteres de la fórmula
 * Para diagnóstico de caracteres invisibles
 */
function diagnosticarCaracteresFormula() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Fórmula de prueba
    var formula = '=SI(ESBLANCO(O2);"SIN_COMPROMISO";SI(O2=HOY();"LLAMAR_HOY";SI(O2<HOY();"COMPROMISO_VENCIDO";"COMPROMISO_FUTURO")))';
    
    var mensaje = 'ANÁLISIS DE CARACTERES\n\n';
    mensaje += 'Fórmula: ' + formula + '\n\n';
    mensaje += 'Longitud: ' + formula.length + ' caracteres\n\n';
    mensaje += 'Códigos ASCII de caracteres especiales:\n\n';
    
    // Buscar comillas
    for (var i = 0; i < formula.length; i++) {
      var char = formula.charAt(i);
      var code = formula.charCodeAt(i);
      
      // Solo mostrar caracteres especiales
      if (char === '"' || char === ';' || char === '=' || char === '(' || char === ')') {
        mensaje += char + ' → ASCII: ' + code;
        
        // Verificar si es el carácter correcto
        if (char === '"' && code !== 34) mensaje += ' ⚠️ INCORRECTO (debe ser 34)';
        if (char === ';' && code !== 59) mensaje += ' ⚠️ INCORRECTO (debe ser 59)';
        
        mensaje += '\n';
      }
    }
    
    mensaje += '\n━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += 'Caracteres correctos:\n';
    mensaje += '• Comillas: " (ASCII 34)\n';
    mensaje += '• Punto y coma: ; (ASCII 59)\n';
    mensaje += '• Igual: = (ASCII 61)\n';
    mensaje += '• Paréntesis: ( ) (ASCII 40, 41)';
    
    Logger.log(mensaje);
    ui.alert('Diagnóstico de Caracteres', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}