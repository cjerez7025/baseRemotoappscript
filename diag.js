/**
 * DIAGNÓSTICO: Comparación exacta byte por byte de nombres
 */
function diagnosticoNombresExactos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prodSheet = ss.getSheetByName('PRODUCTIVIDAD');
  var bddSheet = ss.getSheetByName('BBDD_REPORTE');
  
  if (!prodSheet || !bddSheet) {
    Logger.log('❌ Faltan hojas');
    return;
  }
  
  Logger.log('========================================');
  Logger.log('COMPARACIÓN EXACTA DE NOMBRES');
  Logger.log('========================================');
  Logger.log('');
  
  // Nombres problemáticos
  var nombresProblema = [
    'ALICIA VALBUENA',
    'CAROLINA RETAMAL',
    'MARIA JOSE AZOCAR',
    'MARIA LAGOS',
    'MARTA CERDA'
  ];
  
  var nombresCorrecto = [
    'CATALINA MIRANDA',
    'CLAUDIA ARANCIBIA'
  ];
  
  // 1. Analizar PRODUCTIVIDAD
  Logger.log('=== NOMBRES EN PRODUCTIVIDAD ===');
  var nombresProd = prodSheet.getRange('A2:A15').getValues();
  var mapaProductividad = {};
  
  for (var i = 0; i < nombresProd.length; i++) {
    if (nombresProd[i][0]) {
      var nombre = nombresProd[i][0].toString();
      mapaProductividad[nombre] = true;
      
      var fila = i + 2;
      var esProblema = nombresProblema.indexOf(nombre) !== -1;
      var esCorrecto = nombresCorrecto.indexOf(nombre) !== -1;
      
      if (esProblema || esCorrecto) {
        Logger.log('');
        Logger.log('Fila ' + fila + ': "' + nombre + '"');
        Logger.log('  Longitud: ' + nombre.length + ' caracteres');
        Logger.log('  Bytes: [' + nombre.split('').map(function(c) { return c.charCodeAt(0); }).join(', ') + ']');
        Logger.log('  Tipo: ' + (esProblema ? '❌ PROBLEMA' : '✅ CORRECTO'));
      }
    }
  }
  
  Logger.log('');
  Logger.log('=== NOMBRES EN BBDD_REPORTE (primeras apariciones) ===');
  
  var datosBBDD = bddSheet.getDataRange().getValues();
  var headers = datosBBDD[0];
  var ejecutivoIndex = headers.indexOf('EJECUTIVO');
  
  var nombresBBDDEncontrados = {};
  
  for (var j = 1; j < datosBBDD.length; j++) {
    var nombreBBDD = datosBBDD[j][ejecutivoIndex];
    if (nombreBBDD && !nombresBBDDEncontrados[nombreBBDD]) {
      nombresBBDDEncontrados[nombreBBDD] = j + 1; // Guardar fila
      
      var nombre = nombreBBDD.toString();
      var esProblema = nombresProblema.indexOf(nombre) !== -1;
      var esCorrecto = nombresCorrecto.indexOf(nombre) !== -1;
      
      if (esProblema || esCorrecto) {
        Logger.log('');
        Logger.log('Fila ' + (j + 1) + ': "' + nombre + '"');
        Logger.log('  Longitud: ' + nombre.length + ' caracteres');
        Logger.log('  Bytes: [' + nombre.split('').map(function(c) { return c.charCodeAt(0); }).join(', ') + ']');
        Logger.log('  Tipo: ' + (esProblema ? '❌ PROBLEMA' : '✅ CORRECTO'));
      }
    }
  }
  
  Logger.log('');
  Logger.log('=== COMPARACIÓN DIRECTA ===');
  
  for (var k = 0; k < nombresProblema.length; k++) {
    var nombreBuscar = nombresProblema[k];
    
    Logger.log('');
    Logger.log('Buscando: "' + nombreBuscar + '"');
    
    var enProductividad = mapaProductividad[nombreBuscar] === true;
    var enBBDD = nombresBBDDEncontrados[nombreBuscar] !== undefined;
    
    Logger.log('  En PRODUCTIVIDAD: ' + (enProductividad ? '✓ SÍ' : '✗ NO'));
    Logger.log('  En BBDD_REPORTE: ' + (enBBDD ? '✓ SÍ (fila ' + nombresBBDDEncontrados[nombreBuscar] + ')' : '✗ NO'));
    
    if (enProductividad && enBBDD) {
      Logger.log('  ✓ NOMBRES COINCIDEN - ¿Por qué la fórmula no funciona?');
      
      // Verificar si hay diferencia invisible
      for (var prod in mapaProductividad) {
        if (prod === nombreBuscar) {
          for (var bbdd in nombresBBDDEncontrados) {
            if (bbdd === nombreBuscar) {
              if (prod === bbdd) {
                Logger.log('  ✓ Los strings son idénticos (===)');
              } else {
                Logger.log('  ❌ Los strings NO son idénticos aunque parezcan iguales');
                Logger.log('    PRODUCTIVIDAD: ' + prod.split('').map(function(c) { return c.charCodeAt(0); }).join(','));
                Logger.log('    BBDD_REPORTE:  ' + bbdd.split('').map(function(c) { return c.charCodeAt(0); }).join(','));
              }
            }
          }
        }
      }
    }
  }
  
  Logger.log('');
  Logger.log('=== VERIFICACIÓN DE FÓRMULAS ===');
  
  // Verificar las fórmulas de las filas problema vs correctas
  var formulaProblema = prodSheet.getRange('F2').getFormula(); // ALICIA VALBUENA
  var formulaCorrecta = prodSheet.getRange('F4').getFormula(); // CATALINA MIRANDA
  
  Logger.log('');
  Logger.log('Fórmula fila 2 (ALICIA VALBUENA - PROBLEMA):');
  Logger.log('  ' + formulaProblema);
  Logger.log('');
  Logger.log('Fórmula fila 4 (CATALINA MIRANDA - CORRECTO):');
  Logger.log('  ' + formulaCorrecta);
  
  Logger.log('');
  Logger.log('========================================');
  
  SpreadsheetApp.getUi().alert(
    '🔍 Diagnóstico Completado',
    'Revisa el log para ver la comparación byte por byte de los nombres.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}