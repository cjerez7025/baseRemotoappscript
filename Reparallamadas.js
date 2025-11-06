/**
 * ARCHIVO COMPLETO: Reparallamadas.js
 * Gestión de la hoja LLAMADAS con métricas y gráfico
 * VERSIÓN CORREGIDA - Elimina automáticamente la hoja si existe
 */

console.log('✓ Reparallamadas.js cargado');

function crearTablaLlamadas() {
  console.log('>>> Iniciando crearTablaLlamadas');
  
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      console.log('Error: No existe BBDD_REPORTE');
      return;
    }
    
    // ELIMINAR SI EXISTE (sin preguntar)
    var llamadasSheet = spreadsheet.getSheetByName('LLAMADAS');
    if (llamadasSheet) {
      console.log('⚠️ Hoja LLAMADAS existe. Eliminando...');
      try {
        spreadsheet.deleteSheet(llamadasSheet);
        console.log('✓ Hoja anterior eliminada');
      } catch (deleteError) {
        llamadasSheet.clear();
        console.log('✓ Hoja limpiada');
      }
    }
    
    llamadasSheet = spreadsheet.insertSheet('LLAMADAS');
    console.log('✓ Nueva hoja LLAMADAS creada');
    
    var datos = bddSheet.getDataRange().getValues();
    var headers = datos[0];
    
    var ejecutivoIndex = headers.indexOf('EJECUTIVO');
    var fechaIndex = headers.indexOf('FECHA_LLAMADA');
    
    if (ejecutivoIndex === -1 || fechaIndex === -1) {
      console.log('Error: Columnas no encontradas');
      return;
    }
    
    var ejecutivosSet = new Set();
    var fechasSet = new Set();
    
    for (var i = 1; i < datos.length; i++) {
      var ejecutivo = datos[i][ejecutivoIndex];
      var fecha = datos[i][fechaIndex];
      
      if (ejecutivo) {
        var ejec = ejecutivo.toString().trim();
        if (ejec) ejecutivosSet.add(ejec);
      }
      
      if (fecha) {
        var fech = fecha.toString().trim();
        if (fech) fechasSet.add(fech);
      }
    }
    
    var ejecutivos = Array.from(ejecutivosSet).sort();
    var fechas = Array.from(fechasSet).sort(function(a, b) { return new Date(a) - new Date(b); });
    
    var hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    var fechaHoyStr = formatearFecha(hoy);
    
    if (ejecutivos.length === 0) {
      console.log('No hay ejecutivos');
      llamadasSheet.getRange(1, 1).setValue('CUENTA de rut_cliente');
      llamadasSheet.getRange(1, 2).setValue('FECHA_LLAMADA');
      llamadasSheet.getRange(2, 1).setValue('EJECUTIVO');
      llamadasSheet.getRange(2, 2).setFormula('=TRANSPOSE(UNIQUE(FILTER(BBDD_REPORTE!$N$2:$N;BBDD_REPORTE!$N$2:$N<>"")))');
      llamadasSheet.getRange(2, 5).setValue('Suma total');
      llamadasSheet.getRange(1, 1, 2, 5)
        .setBackground('#4472C4')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      llamadasSheet.autoResizeColumns(1, 5);
      return;
    }
    
    // ESTRUCTURA BÁSICA
    llamadasSheet.getRange(1, 1).setValue('CUENTA de rut_cliente');
    llamadasSheet.getRange(1, 2).setValue('FECHA_LLAMADA');
    llamadasSheet.getRange(2, 1).setValue('EJECUTIVO');
    llamadasSheet.getRange(2, 2).setFormula('=TRANSPOSE(UNIQUE(FILTER(BBDD_REPORTE!$N$2:$N;BBDD_REPORTE!$N$2:$N<>"")))');
    
    var numColumnasConDatos = fechas.length > 0 ? fechas.length : 0;
    var numColumnasBuffer = 2;
    var totalColumnasFormulas = numColumnasConDatos + numColumnasBuffer;
    var colSumaHeader = 2 + totalColumnasFormulas;
    
    llamadasSheet.getRange(2, colSumaHeader).setValue('Suma total');
    
    var maxCols = Math.max(colSumaHeader, 15);
    llamadasSheet.getRange(3, 1, 1, maxCols).setBackground('#E8E8E8');
    
    // EJECUTIVOS
    var filaEjecutivos = 4;
    for (var i = 0; i < ejecutivos.length; i++) {
      var fila = filaEjecutivos + i;
      llamadasSheet.getRange(fila, 1).setValue(ejecutivos[i]);
      
      for (var j = 0; j < totalColumnasFormulas; j++) {
        var columna = j + 2;
        var letraCol = columnNumberToLetter(columna);
        var formula = '=COUNTIFS(BBDD_REPORTE!$C$2:$C;$A' + fila + ';BBDD_REPORTE!$N$2:$N;' + letraCol + '$2)';
        llamadasSheet.getRange(fila, columna).setFormula(formula);
      }
      
      var letraFirst = 'B';
      var letraLast = columnNumberToLetter(2 + totalColumnasFormulas - 1);
      var formulaSum = '=SUM(' + letraFirst + fila + ':' + letraLast + fila + ')';
      llamadasSheet.getRange(fila, colSumaHeader).setFormula(formulaSum);
    }
    
    // TOTALES
    var filaTotal = ejecutivos.length + filaEjecutivos;
    llamadasSheet.getRange(filaTotal, 1).setValue('Suma total');
    
    for (var j = 0; j < totalColumnasFormulas; j++) {
      var columna = j + 2;
      var letraCol = columnNumberToLetter(columna);
      var formula = '=SUM(' + letraCol + filaEjecutivos + ':' + letraCol + (ejecutivos.length + filaEjecutivos - 1) + ')';
      llamadasSheet.getRange(filaTotal, columna).setFormula(formula);
    }
    
    var letraFirst = 'B';
    var letraLast = columnNumberToLetter(2 + totalColumnasFormulas - 1);
    llamadasSheet.getRange(filaTotal, colSumaHeader)
      .setFormula('=SUM(' + letraFirst + filaTotal + ':' + letraLast + filaTotal + ')');
    
    // FORMATO
    var maxColsFormat = Math.max(colSumaHeader, 28);
    
    llamadasSheet.getRange(1, 1, 2, maxColsFormat)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    llamadasSheet.getRange(filaTotal, 1, 1, maxColsFormat)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    llamadasSheet.getRange(1, 1, filaTotal, maxColsFormat)
      .setBorder(true, true, true, true, true, true);
    
    if (ejecutivos.length > 0) {
      llamadasSheet.getRange(filaEjecutivos, 2, ejecutivos.length, totalColumnasFormulas)
        .setHorizontalAlignment('center');
    }
    
    llamadasSheet.autoResizeColumn(1);
    
    if (numColumnasConDatos > 0) {
      for (var col = 2; col <= 2 + numColumnasConDatos - 1; col++) {
        llamadasSheet.autoResizeColumn(col);
      }
    }
    
    var anchoBuffer = 100;
    var colInicioBuffer = 2 + numColumnasConDatos;
    var colFinBuffer = colInicioBuffer + numColumnasBuffer - 1;
    
    for (var col = colInicioBuffer; col <= colFinBuffer; col++) {
      llamadasSheet.setColumnWidth(col, anchoBuffer);
    }
    
    llamadasSheet.autoResizeColumn(colSumaHeader);
    
    // RESALTAR HOY
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    
    var fechasGeneradas = llamadasSheet.getRange(2, 2, 1, totalColumnasFormulas).getValues()[0];
    var columnaHoy = -1;
    
    for (var j = 0; j < fechasGeneradas.length; j++) {
      if (fechasGeneradas[j]) {
        var fechaCell = fechasGeneradas[j];
        var fechaStr = '';
        
        if (fechaCell instanceof Date) {
          fechaStr = formatearFecha(fechaCell);
        } else {
          fechaStr = fechaCell.toString().trim();
        }
        
        if (fechaStr === fechaHoyStr) {
          columnaHoy = j + 2;
          break;
        }
      }
    }
    
    if (columnaHoy !== -1) {
      llamadasSheet.getRange(1, columnaHoy, 2, 1)
        .setBackground('#FFD966')
        .setFontWeight('bold');
      
      if (ejecutivos.length > 0) {
        llamadasSheet.getRange(filaEjecutivos, columnaHoy, ejecutivos.length, 1)
          .setBackground('#FFF2CC');
      }
      
      llamadasSheet.getRange(filaTotal, columnaHoy)
        .setBackground('#FFD966')
        .setFontWeight('bold');
      
      console.log('✓ Columna HOY resaltada');
    }
    
    console.log('✓ Tabla LLAMADAS creada');
    console.log('Ejecutivos: ' + ejecutivos.length);
    console.log('Fechas: ' + fechas.length);
    
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    
    crearGraficoProgresoLlamadas();
    
    console.log('✅ Proceso completo');
    
  } catch (error) {
    console.error('❌ Error: ' + error.message);
    console.error(error.stack);
    throw error;
  }
}

function crearGraficoProgresoLlamadas() {
  console.log('>>> Iniciando crearGraficoProgresoLlamadas');
  
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var llamadasSheet = spreadsheet.getSheetByName('LLAMADAS');
    var bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!llamadasSheet || !bddSheet) {
      console.log('Error: Hojas no encontradas');
      return;
    }
    
    var datos = llamadasSheet.getDataRange().getValues();
    var colSumaTotal = -1;
    
    for (var col = 0; col < datos[1].length; col++) {
      if (datos[1][col] === 'Suma total') {
        colSumaTotal = col;
        break;
      }
    }
    
    if (colSumaTotal === -1) {
      console.log('Error: Columna Suma total no encontrada');
      return;
    }
    
    var ejecutivos = [];
    var filaInicio = 4;
    var filaUltimoEjecutivo = filaInicio;
    
    for (var i = filaInicio - 1; i < datos.length; i++) {
      var ejecutivo = datos[i][0];
      if (ejecutivo && ejecutivo !== 'Suma total') {
        ejecutivos.push({
          nombre: ejecutivo,
          fila: i + 1
        });
        filaUltimoEjecutivo = i + 1;
      } else if (ejecutivo === 'Suma total') {
        break;
      }
    }
    
    if (ejecutivos.length === 0) {
      console.log('No hay ejecutivos');
      return;
    }
    
    var headers = bddSheet.getRange(1, 1, 1, bddSheet.getLastColumn()).getValues()[0];
    var ejecutivoIndex = headers.indexOf('EJECUTIVO');
    
    if (ejecutivoIndex === -1) {
      console.log('Error: Columna EJECUTIVO no encontrada');
      return;
    }
    
    var colEjecutivoBBDD = columnNumberToLetter(ejecutivoIndex + 1);
    var filaInicioMetricas = filaUltimoEjecutivo + 3;
    var colInicioMetricas = 1;
    
    llamadasSheet.getRange(filaInicioMetricas, colInicioMetricas)
      .setValue('📊 PROGRESO DE LLAMADAS POR EJECUTIVO')
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#F3F3F3');
    
    var filaEncabezadosMetricas = filaInicioMetricas + 2;
    var encabezadosMetricas = ['EJECUTIVO', 'META', 'LLAMADAS', '% AVANCE', 'PENDIENTE'];
    
    llamadasSheet.getRange(filaEncabezadosMetricas, colInicioMetricas, 1, encabezadosMetricas.length)
      .setValues([encabezadosMetricas])
      .setBackground('#34A853')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    for (var i = 0; i < ejecutivos.length; i++) {
      var fila = filaEncabezadosMetricas + 1 + i;
      var ejecutivo = ejecutivos[i];
      var letraSumaTotal = columnNumberToLetter(colSumaTotal + 1);
      
      llamadasSheet.getRange(fila, colInicioMetricas).setValue(ejecutivo.nombre);
      
      var formulaMeta = '=COUNTIF(BBDD_REPORTE!$' + colEjecutivoBBDD + '$2:$' + colEjecutivoBBDD + ';A' + fila + ')';
      llamadasSheet.getRange(fila, colInicioMetricas + 1).setFormula(formulaMeta);
      
      var formulaLlamadas = '=' + letraSumaTotal + ejecutivo.fila;
      llamadasSheet.getRange(fila, colInicioMetricas + 2).setFormula(formulaLlamadas);
      
      var formulaAvance = '=IFERROR(C' + fila + '/B' + fila + ';0)';
      llamadasSheet.getRange(fila, colInicioMetricas + 3).setFormula(formulaAvance);
      llamadasSheet.getRange(fila, colInicioMetricas + 3).setNumberFormat('0.0%');
      
      var formulaPendiente = '=B' + fila + '-C' + fila;
      llamadasSheet.getRange(fila, colInicioMetricas + 4).setFormula(formulaPendiente);
    }
    
    var filaTotalMetricas = filaEncabezadosMetricas + 1 + ejecutivos.length;
    llamadasSheet.getRange(filaTotalMetricas, colInicioMetricas).setValue('TOTAL');
    
    var columnasASumar = [
      {offset: 1, letra: 'B'},
      {offset: 2, letra: 'C'},
      {offset: 4, letra: 'E'}
    ];
    
    for (var i = 0; i < columnasASumar.length; i++) {
      var col = columnasASumar[i];
      llamadasSheet.getRange(filaTotalMetricas, colInicioMetricas + col.offset)
        .setFormula('=SUM(' + col.letra + (filaEncabezadosMetricas + 1) + ':' + col.letra + (filaTotalMetricas - 1) + ')');
    }
    
    var formulaAvanceTotal = '=IFERROR(C' + filaTotalMetricas + '/B' + filaTotalMetricas + ';0)';
    llamadasSheet.getRange(filaTotalMetricas, colInicioMetricas + 3)
      .setFormula(formulaAvanceTotal)
      .setNumberFormat('0.0%');
    
    llamadasSheet.getRange(filaTotalMetricas, colInicioMetricas, 1, encabezadosMetricas.length)
      .setBackground('#34A853')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    llamadasSheet.getRange(filaEncabezadosMetricas, colInicioMetricas, ejecutivos.length + 2, encabezadosMetricas.length)
      .setBorder(true, true, true, true, true, true);
    
    llamadasSheet.getRange(filaEncabezadosMetricas + 1, colInicioMetricas + 1, ejecutivos.length, 4)
      .setHorizontalAlignment('center');
    
    for (var col = colInicioMetricas; col < colInicioMetricas + encabezadosMetricas.length; col++) {
      llamadasSheet.autoResizeColumn(col);
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    
    var charts = llamadasSheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
      llamadasSheet.removeChart(charts[i]);
    }
    
    var filaGrafico = filaTotalMetricas + 2;
    var numFilasDatos = ejecutivos.length;
    
    try {
      var rangoGrafico = llamadasSheet.getRange(filaEncabezadosMetricas, 1, numFilasDatos + 1, 5);
      
      // ⭐ CORRECCIÓN APLICADA: Removido 'bottom: 120' de chartArea
      var chart = llamadasSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(rangoGrafico)
        .setPosition(filaGrafico, colInicioMetricas, 0, 0)
        .setNumHeaders(1)
        .setOption('useFirstColumnAsDomain', true)
        .setOption('title', 'LLAMADAS y PENDIENTE')
        .setOption('width', 1000)
        .setOption('height', 500)
        .setOption('colors', ['#34A853', '#4285F4', '#FBBC04', '#EA4335'])
        .setOption('legend', {position: 'top', textStyle: {fontSize: 12, bold: true}})
        .setOption('hAxis', {title: 'EJECUTIVO', textStyle: {fontSize: 9}, slantedText: true, slantedTextAngle: 45})
        .setOption('vAxis', {title: 'Cantidad', minValue: 0, textStyle: {fontSize: 11}})
        .setOption('chartArea', {width: '70%', height: '60%', left: 100, top: 70})
        .setOption('bar', {groupWidth: '60%'})
        .setOption('titleTextStyle', {fontSize: 16, bold: true})
        .build();
      
      llamadasSheet.insertChart(chart);
      console.log('✓ Gráfico creado');
      
    } catch (errorChart) {
      console.error('Error creando gráfico: ' + errorChart.message);
    }
    
    console.log('✓ Gráfico completado');
    
  } catch (error) {
    console.error('❌ Error: ' + error.message);
    throw error;
  }
}

function formatearFecha(fecha) {
  if (!(fecha instanceof Date)) return '';
  var dia = String(fecha.getDate()).padStart(2, '0');
  var mes = String(fecha.getMonth() + 1).padStart(2, '0');
  var año = fecha.getFullYear();
  return dia + '-' + mes + '-' + año;
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