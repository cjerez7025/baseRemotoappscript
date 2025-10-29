/**
 * DIAGNÓSTICO: Por qué PRODUCTIVIDAD muestra ceros para ejecutivas nuevas
 */
function diagnosticarProductividadCeros() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = ss.getSheetByName('BBDD_REPORTE');
    var prodSheet = ss.getSheetByName('PRODUCTIVIDAD');
    
    if (!bddSheet || !prodSheet) {
      Logger.log('❌ Faltan hojas necesarias');
      return;
    }
    
    Logger.log('============================================');
    Logger.log('   DIAGNÓSTICO PRODUCTIVIDAD - CEROS');
    Logger.log('============================================');
    Logger.log('');
    
    // 1. VERIFICAR NOMBRES EN PRODUCTIVIDAD
    Logger.log('=== 1. NOMBRES EN PRODUCTIVIDAD (Tabla 1) ===');
    var nombresProductividad = prodSheet.getRange('A2:A15').getValues();
    var ejecutivosProductividad = [];
    for (var i = 0; i < nombresProductividad.length; i++) {
      if (nombresProductividad[i][0]) {
        ejecutivosProductividad.push(nombresProductividad[i][0].toString().trim());
        Logger.log((i+1) + '. "' + nombresProductividad[i][0] + '"');
      }
    }
    Logger.log('Total ejecutivos en PRODUCTIVIDAD: ' + ejecutivosProductividad.length);
    Logger.log('');
    
    // 2. VERIFICAR NOMBRES EN BBDD_REPORTE
    Logger.log('=== 2. NOMBRES ÚNICOS EN BBDD_REPORTE ===');
    var datosBBDD = bddSheet.getDataRange().getValues();
    var headers = datosBBDD[0];
    var ejecutivoIndex = headers.indexOf('EJECUTIVO');
    
    if (ejecutivoIndex === -1) {
      Logger.log('❌ No se encontró columna EJECUTIVO en BBDD_REPORTE');
      return;
    }
    
    Logger.log('✓ Columna EJECUTIVO está en posición: ' + (ejecutivoIndex + 1));
    
    var ejecutivosUnicos = {};
    for (var j = 1; j < datosBBDD.length; j++) {
      var nombre = datosBBDD[j][ejecutivoIndex];
      if (nombre && nombre.toString().trim() !== '') {
        var nombreLimpio = nombre.toString().trim();
        if (!ejecutivosUnicos[nombreLimpio]) {
          ejecutivosUnicos[nombreLimpio] = 0;
        }
        ejecutivosUnicos[nombreLimpio]++;
      }
    }
    
    var listaEjecutivos = Object.keys(ejecutivosUnicos).sort();
    Logger.log('Total ejecutivos únicos en BBDD_REPORTE: ' + listaEjecutivos.length);
    Logger.log('');
    for (var k = 0; k < listaEjecutivos.length; k++) {
      Logger.log((k+1) + '. "' + listaEjecutivos[k] + '" → ' + ejecutivosUnicos[listaEjecutivos[k]] + ' registros');
    }
    Logger.log('');
    
    // 3. COMPARAR NOMBRES
    Logger.log('=== 3. COMPARACIÓN DE NOMBRES ===');
    Logger.log('Buscando diferencias entre PRODUCTIVIDAD y BBDD_REPORTE...');
    Logger.log('');
    
    for (var m = 0; m < ejecutivosProductividad.length; m++) {
      var nombreProd = ejecutivosProductividad[m];
      var encontrado = false;
      var registros = 0;
      
      for (var n = 0; n < listaEjecutivos.length; n++) {
        if (nombreProd === listaEjecutivos[n]) {
          encontrado = true;
          registros = ejecutivosUnicos[listaEjecutivos[n]];
          break;
        }
      }
      
      if (encontrado) {
        Logger.log('✓ "' + nombreProd + '" → ENCONTRADO (' + registros + ' registros)');
      } else {
        Logger.log('❌ "' + nombreProd + '" → NO ENCONTRADO EN BBDD_REPORTE');
        
        // Buscar similares
        Logger.log('   Buscando nombres similares...');
        for (var p = 0; p < listaEjecutivos.length; p++) {
          var nombreBBDD = listaEjecutivos[p];
          var similar = false;
          
          // Comparar ignorando mayúsculas/minúsculas
          if (nombreProd.toUpperCase() === nombreBBDD.toUpperCase()) {
            similar = true;
            Logger.log('   ⚠️ Encontrado con diferente mayúscula: "' + nombreBBDD + '"');
          }
          
          // Comparar con guiones vs espacios
          var nombreProdSinGuion = nombreProd.replace(/_/g, ' ');
          if (nombreProdSinGuion === nombreBBDD) {
            similar = true;
            Logger.log('   ⚠️ Encontrado con espacios: "' + nombreBBDD + '"');
          }
          
          // Comparar sin espacios ni guiones
          var prodLimpio = nombreProd.replace(/[_\s]/g, '').toUpperCase();
          var bbddLimpio = nombreBBDD.replace(/[_\s]/g, '').toUpperCase();
          if (prodLimpio === bbddLimpio) {
            similar = true;
            Logger.log('   ⚠️ Encontrado (diferente formato): "' + nombreBBDD + '"');
          }
        }
      }
    }
    Logger.log('');
    
    // 4. VERIFICAR HOJAS INDIVIDUALES
    Logger.log('=== 4. VERIFICACIÓN DE HOJAS INDIVIDUALES ===');
    var hojasProblema = ['ALICIA_VALBUENA', 'CAROLINA_RETAMAL', 'MARIA_JOSE_AZOCAR', 'MARIA_LAGOS', 'MARTA_CERDA'];
    
    for (var q = 0; q < hojasProblema.length; q++) {
      var nombreHoja = hojasProblema[q];
      var hoja = ss.getSheetByName(nombreHoja);
      
      if (hoja) {
        var ultimaFila = hoja.getLastRow();
        Logger.log('✓ Hoja "' + nombreHoja + '" existe → ' + (ultimaFila - 1) + ' registros');
        
        if (ultimaFila > 1) {
          // Leer un dato de ejemplo
          var ejemplo = hoja.getRange(2, 1, 1, Math.min(3, hoja.getLastColumn())).getValues()[0];
          Logger.log('   Ejemplo fila 2: ' + ejemplo.join(' | '));
        }
      } else {
        Logger.log('❌ Hoja "' + nombreHoja + '" NO EXISTE');
        
        // Buscar con espacios
        var nombreConEspacio = nombreHoja.replace(/_/g, ' ');
        var hojaEspacio = ss.getSheetByName(nombreConEspacio);
        if (hojaEspacio) {
          Logger.log('   ⚠️ Pero existe como: "' + nombreConEspacio + '"');
        }
      }
    }
    Logger.log('');
    
    // 5. VERIFICAR FÓRMULA EN BBDD_REPORTE
    Logger.log('=== 5. FÓRMULA DE BBDD_REPORTE ===');
    var formulaBBDD = bddSheet.getRange('A2').getFormula();
    Logger.log('Fórmula en A2:');
    Logger.log(formulaBBDD);
    Logger.log('');
    
    // Verificar si las hojas problema están en la fórmula
    for (var r = 0; r < hojasProblema.length; r++) {
      if (formulaBBDD.indexOf(hojasProblema[r]) !== -1) {
        Logger.log('✓ "' + hojasProblema[r] + '" está en la fórmula');
      } else {
        Logger.log('❌ "' + hojasProblema[r] + '" NO está en la fórmula');
      }
    }
    Logger.log('');
    
    // 6. VERIFICAR UNA FÓRMULA EN PRODUCTIVIDAD
    Logger.log('=== 6. EJEMPLO DE FÓRMULA EN PRODUCTIVIDAD ===');
    var formulaProd = prodSheet.getRange('B2').getFormula();
    Logger.log('Fórmula en B2 (primera ejecutiva, columna Sin Gestión):');
    Logger.log(formulaProd);
    Logger.log('');
    
    Logger.log('============================================');
    Logger.log('        FIN DEL DIAGNÓSTICO');
    Logger.log('============================================');
    
    // Mostrar resumen
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '🔍 Diagnóstico Completado',
      'Revisa el log (Ver > Registros de ejecución) para ver el análisis completo.\n\n' +
      'Se analizaron:\n' +
      '• Nombres en PRODUCTIVIDAD\n' +
      '• Nombres en BBDD_REPORTE\n' +
      '• Comparación entre ambos\n' +
      '• Estado de hojas individuales\n' +
      '• Fórmulas',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    Logger.log(error.stack);
  }
}