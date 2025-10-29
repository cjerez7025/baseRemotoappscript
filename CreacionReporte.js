/**
 * MÓDULO 6: CREACIÓN DE BBDD_REPORTE - MEJORADO
 * Consolida datos de TODAS las hojas de ejecutivos
 * Incluye hojas de carga inicial Y distribución adicional
 */

/**
 * Crea o actualiza la hoja BBDD_REPORTE
 * MEJORA: Detecta y consolida TODAS las hojas de ejecutivos sin importar cuándo fueron creadas
 */
function crearOActualizarReporteAutomatico(ss) {
  try {
    Logger.log('=== INICIANDO CREACIÓN/ACTUALIZACIÓN DE BBDD_REPORTE ===');
    
    // Eliminar hoja existente
    var existe = ss.getSheetByName('BBDD_REPORTE');
    if (existe) {
      Logger.log('Eliminando BBDD_REPORTE existente');
      ss.deleteSheet(existe);
    }
    
    var reporte = ss.insertSheet('BBDD_REPORTE');
    var hojas = ss.getSheets();
    var ejecutivos = [];
    
    Logger.log('Total de hojas en el spreadsheet: ' + hojas.length);
    
    // MEJORA: Identificar TODAS las hojas de ejecutivos existentes
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      
      // Excluir hoja origen (BBDD_*_REMOTO)
      var esOrigen = /^BBDD_.*_REMOTO/i.test(nombre);
      
      // Verificar si es hoja excluida (hojas del sistema)
      var esExcluida = esOrigen;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.indexOf(HOJAS_EXCLUIDAS[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      // Si no está excluida y tiene datos, verificar si es hoja de ejecutivo
      if (!esExcluida && hojas[i].getLastRow() > 1) {
        try {
          // Obtener encabezados (primeras 20 columnas para evitar exceso)
          var enc = hojas[i].getRange(1, 1, 1, Math.min(hojas[i].getLastColumn(), 20)).getValues()[0];
          
          // MEJORA: Verificar múltiples columnas características de hojas de ejecutivos
          var requisitos = [
            'FECHA_LLAMADA', 
            'ESTADO_COMPROMISO', 
            'SUB_ESTADO', 
            'NOTA_EJECUTIVO',
            'ESTADO',
            'EJECUTIVO'
          ];
          
          // Si tiene al menos una de estas columnas, es hoja de ejecutivo
          var esHojaEjecutivo = false;
          for (var k = 0; k < requisitos.length; k++) {
            if (enc.indexOf(requisitos[k]) !== -1) {
              esHojaEjecutivo = true;
              break;
            }
          }
          
          if (esHojaEjecutivo) {
            ejecutivos.push(hojas[i]);
            Logger.log('✓ Hoja detectada: ' + nombre + ' (tiene ' + hojas[i].getLastRow() + ' filas)');
          }
          
        } catch (error) {
          Logger.log('Error verificando hoja ' + nombre + ': ' + error.toString());
          continue;
        }
      }
    }
    
    Logger.log('=== HOJAS DE EJECUTIVOS DETECTADAS: ' + ejecutivos.length + ' ===');
    
    if (ejecutivos.length === 0) {
      Logger.log('❌ No hay hojas de ejecutivos para crear BBDD_REPORTE');
      throw new Error('No hay hojas de ejecutivos para crear BBDD_REPORTE');
    }
    
    // Mostrar detalle de hojas incluidas
    for (var m = 0; m < ejecutivos.length; m++) {
      Logger.log('  → ' + ejecutivos[m].getName());
    }
    
    // Obtener encabezados de la primera hoja con más columnas
    // MEJORA: Usar la hoja con más columnas para asegurar capturar todos los encabezados
    var hojaReferencia = ejecutivos[0];
    var maxColumnas = hojaReferencia.getLastColumn();
    
    for (var n = 1; n < ejecutivos.length; n++) {
      if (ejecutivos[n].getLastColumn() > maxColumnas) {
        hojaReferencia = ejecutivos[n];
        maxColumnas = ejecutivos[n].getLastColumn();
      }
    }
    
    Logger.log('Hoja de referencia para encabezados: ' + hojaReferencia.getName() + ' (' + maxColumnas + ' columnas)');
    
    var encabezados = hojaReferencia.getRange(1, 1, 1, maxColumnas).getValues()[0];
    var ultimaCol = columnNumberToLetter(encabezados.length);
    
    // Escribir encabezados
    reporte.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    reporte.getRange(1, 1, 1, encabezados.length)
      .setBackground(COLORES.HEADER_REPORTE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setFontColor('white');
    
    // MEJORA: Crear fórmula de consolidación incluyendo TODAS las hojas detectadas
    var nombres = [];
    for (var p = 0; p < ejecutivos.length; p++) {
      var nombreHoja = ejecutivos[p].getName();
      // Escapar nombres de hojas que contengan caracteres especiales
      if (nombreHoja.indexOf("'") !== -1) {
        nombreHoja = nombreHoja.replace(/'/g, "''");
      }
      nombres.push("'" + nombreHoja + "'!A2:" + ultimaCol);
    }
    
    //var formula = '={' + nombres.join(';') + '}';
    var formula = '=SORT({' + nombres.join(';') + '};3;TRUE)';
    
    Logger.log('Fórmula generada con ' + nombres.length + ' referencias de hojas');
    Logger.log('Longitud de fórmula: ' + formula.length + ' caracteres');
    
    // Aplicar fórmula
    try {
      reporte.getRange('A2').setFormula(formula);
      SpreadsheetApp.flush();
      Utilities.sleep(2000); // MEJORA: Mayor tiempo de espera para procesamiento
      
      var valor = reporte.getRange('A2').getDisplayValue();
      if (valor && valor.indexOf('#ERROR') === -1 && valor.indexOf('#REF') === -1) {
        var totalFilas = reporte.getLastRow() - 1; // Restar encabezado
        Logger.log('✓ BBDD_REPORTE creado correctamente con ' + totalFilas + ' registros consolidados');
      } else {
        Logger.log('⚠️ Advertencia: Posible error en fórmula. Valor en A2: ' + valor);
        throw new Error('Error en fórmula de consolidación: ' + valor);
      }
    } catch (error) {
      Logger.log('❌ Error aplicando fórmula: ' + error.toString());
      throw error;
    }
    
    // Aplicar filtro
    Utilities.sleep(1000);
    var ultimaFila = reporte.getLastRow();
    
    Logger.log('Aplicando filtro a ' + ultimaFila + ' filas');
    
    if (ultimaFila > 1) {
      try {
        reporte.getRange(1, 1, ultimaFila, encabezados.length).createFilter();
        Logger.log('✓ Filtro aplicado correctamente');
      } catch (filterError) {
        Logger.log('⚠️ No se pudo aplicar filtro: ' + filterError.toString());
      }
    }
    
    // Formato final
    reporte.autoResizeColumns(1, Math.min(encabezados.length, 26)); // Limitar a 26 columnas (A-Z)
    reporte.setFrozenRows(1);
    
    Logger.log('=== BBDD_REPORTE COMPLETADO ===');
    Logger.log('✓ ' + ejecutivos.length + ' hojas consolidadas');
    Logger.log('✓ ' + (ultimaFila - 1) + ' registros totales');
    Logger.log('✓ ' + encabezados.length + ' columnas');
    
  } catch (error) {
    Logger.log('❌ ERROR CRÍTICO en crearOActualizarReporteAutomatico: ' + error.toString());
    throw error;
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