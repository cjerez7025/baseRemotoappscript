/**
 * MÓDULO 6: CREACIÓN DE BBDD_REPORTE
 * Consolida datos de todas las hojas de ejecutivos
 */

/**
 * Crea o actualiza la hoja BBDD_REPORTE
 */
function crearOActualizarReporteAutomatico(ss) {
  try {
    // Eliminar hoja existente
    var existe = ss.getSheetByName('BBDD_REPORTE');
    if (existe) ss.deleteSheet(existe);
    
    var reporte = ss.insertSheet('BBDD_REPORTE');
    var hojas = ss.getSheets();
    var ejecutivos = [];
    
    // Identificar hojas de ejecutivos
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      var esOrigen = /^BBDD_.*_REMOTO/i.test(nombre);
      
      var esExcluida = esOrigen;
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.indexOf(HOJAS_EXCLUIDAS[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (!esExcluida && hojas[i].getLastRow() > 1) {
        try {
          var enc = hojas[i].getRange(1, 1, 1, Math.min(hojas[i].getLastColumn(), 20)).getValues()[0];
          var requisitos = ['FECHA_LLAMADA', 'ESTADO_COMPROMISO', 'SUB_ESTADO', 'NOTA_EJECUTIVO'];
          
          for (var k = 0; k < requisitos.length; k++) {
            if (enc.indexOf(requisitos[k]) !== -1) {
              ejecutivos.push(hojas[i]);
              break;
            }
          }
        } catch (error) {
          continue;
        }
      }
    }
    
    if (ejecutivos.length === 0) {
      Logger.log('No hay hojas de ejecutivos');
      throw new Error('No hay hojas de ejecutivos para crear BBDD_REPORTE');
    }
    
    // Obtener encabezados de la primera hoja
    var primera = ejecutivos[0];
    var encabezados = primera.getRange(1, 1, 1, primera.getLastColumn()).getValues()[0];
    var ultimaCol = columnNumberToLetter(encabezados.length);
    
    // Escribir encabezados
    reporte.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    reporte.getRange(1, 1, 1, encabezados.length)
      .setBackground(COLORES.HEADER_REPORTE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setFontColor('white');
    
    // Crear fórmula de consolidación
    var nombres = [];
    for (var m = 0; m < ejecutivos.length; m++) {
      nombres.push("'" + ejecutivos[m].getName() + "'!A2:" + ultimaCol);
    }
    
    var formula = '={' + nombres.join(';') + '}';
    
    // Aplicar fórmula
    try {
      reporte.getRange('A2').setFormula(formula);
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      
      var valor = reporte.getRange('A2').getDisplayValue();
      if (valor && valor.indexOf('#ERROR') === -1) {
        Logger.log('BBDD_REPORTE creado correctamente');
      } else {
        throw new Error('Error en fórmula de consolidación');
      }
    } catch (error) {
      Logger.log('Error aplicando fórmula: ' + error.toString());
      throw error;
    }
    
    // Aplicar filtro
    Utilities.sleep(500);
    var ultimaFila = reporte.getLastRow();
    if (ultimaFila > 1) {
      reporte.getRange(1, 1, ultimaFila, encabezados.length).createFilter();
    }
    
    // Formato final
    reporte.autoResizeColumns(1, encabezados.length);
    reporte.setFrozenRows(1);
    
    Logger.log('BBDD_REPORTE completado con ' + ejecutivos.length + ' hojas');
    
  } catch (error) {
    Logger.log('Error creando reporte: ' + error);
    throw error;
  }
}