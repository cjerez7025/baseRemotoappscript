/**
 * ARCHIVO: Reparallamadas.js
 * Gestión de la hoja LLAMADAS con métricas y gráfico de progreso
 * VERSIÓN CORREGIDA - Gráfico de columnas verticales agrupadas FUNCIONAL
 */

console.log('✓ Reparallamadas.js cargado');

/**
 * Crea tabla LLAMADAS con UNIQUE+TRANSPOSE en N2 para fechas dinámicas
 */
function crearTablaLlamadas() {
  console.log('>>> Iniciando crearTablaLlamadas');
  
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      console.log('Error: No existe BBDD_REPORTE');
      return;
    }
    
    var llamadasSheet = spreadsheet.getSheetByName('LLAMADAS');
    if (llamadasSheet) {
      spreadsheet.deleteSheet(llamadasSheet);
    }
    
    llamadasSheet = spreadsheet.insertSheet('LLAMADAS');
    
    var datos = bddSheet.getDataRange().getValues();
    var headers = datos[0];
    
    var ejecutivoIndex = headers.indexOf('EJECUTIVO');
    var fechaIndex = headers.indexOf('FECHA_LLAMADA');
    
    if (ejecutivoIndex === -1 || fechaIndex === -1) {
      console.log('Error: Columnas no encontradas');
      return;
    }
    
    // Recopilar ejecutivos y fechas únicas
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
    
    // Obtener fecha actual (sin hora)
    var hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    var fechaHoyStr = formatearFecha(hoy);
    
    if (ejecutivos.length === 0) {
      console.log('No hay ejecutivos en BBDD_REPORTE');
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
      console.log('✓ Tabla LLAMADAS creada (sin datos)');
      return;
    }
    
    // FILA 1: Encabezados principales
    llamadasSheet.getRange(1, 1).setValue('CUENTA de rut_cliente');
    llamadasSheet.getRange(1, 2).setValue('FECHA_LLAMADA');
    
    // FILA 2: EJECUTIVO + Fechas
    llamadasSheet.getRange(2, 1).setValue('EJECUTIVO');
    
    var formulaUnique = '=TRANSPOSE(UNIQUE(FILTER(BBDD_REPORTE!$N$2:$N;BBDD_REPORTE!$N$2:$N<>"")))';
    llamadasSheet.getRange(2, 2).setFormula(formulaUnique);
    
    var numColumnasConDatos = fechas.length > 0 ? fechas.length : 0;
    var numColumnasBuffer = 2;
    var totalColumnasFormulas = numColumnasConDatos + numColumnasBuffer;
    
    var colSumaHeader = 2 + totalColumnasFormulas;
    llamadasSheet.getRange(2, colSumaHeader).setValue('Suma total');
    
    // FILA 3: Separador
    var maxCols = Math.max(colSumaHeader, 15);
    llamadasSheet.getRange(3, 1, 1, maxCols).setBackground('#E8E8E8');
    
    // FILAS DE DATOS
    for (var i = 0; i < ejecutivos.length; i++) {
      var fila = i + 4;
      var ejecutivo = ejecutivos[i];
      
      llamadasSheet.getRange(fila, 1).setValue(ejecutivo);
      
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
    
    // FILA DE TOTALES
    var filaTotal = ejecutivos.length + 4;
    llamadasSheet.getRange(filaTotal, 1).setValue('Suma total');
    
    for (var j = 0; j < totalColumnasFormulas; j++) {
      var columna = j + 2;
      var letraCol = columnNumberToLetter(columna);
      var formula = '=SUM(' + letraCol + '4:' + letraCol + (ejecutivos.length + 3) + ')';
      llamadasSheet.getRange(filaTotal, columna).setFormula(formula);
    }
    
    var letraFirst = 'B';
    var letraLast = columnNumberToLetter(2 + totalColumnasFormulas - 1);
    llamadasSheet.getRange(filaTotal, colSumaHeader)
      .setFormula('=SUM(' + letraFirst + filaTotal + ':' + letraLast + filaTotal + ')');
    
    // FORMATO BASE
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
      llamadasSheet.getRange(4, 2, ejecutivos.length, totalColumnasFormulas)
        .setHorizontalAlignment('center');
    }
    
    // AUTO-AJUSTAR COLUMNAS
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
    
    console.log('✓ Columnas buffer con ancho fijo: ' + anchoBuffer + 'px');
    
    // RESALTAR COLUMNA DE FECHA ACTUAL
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
        llamadasSheet.getRange(4, columnaHoy, ejecutivos.length, 1)
          .setBackground('#FFF2CC');
      }
      
      llamadasSheet.getRange(filaTotal, columnaHoy)
        .setBackground('#FFD966')
        .setFontWeight('bold');
      
      console.log('✓ Columna de fecha actual resaltada');
    }
    
    console.log('✓ Tabla LLAMADAS creada');
    console.log('Ejecutivos: ' + ejecutivos.length);
    console.log('Fechas: ' + fechas.length);
    
    // CREAR GRÁFICO
    console.log('Generando tabla de métricas y gráfico...');
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    
    crearGraficoProgresoLlamadas();
    
    console.log('✅ Proceso completo finalizado');
    
  } catch (error) {
    console.error('❌ Error en crearTablaLlamadas: ' + error.message);
    console.error(error.stack);
    throw error;
  }
}

/**
 * Crea tabla de métricas y gráfico de columnas agrupadas
 * CORREGIDO: Usa rangos separados correctamente
 */
function crearGraficoProgresoLlamadas() {
  console.log('>>> Iniciando crearGraficoProgresoLlamadas');
  
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var llamadasSheet = spreadsheet.getSheetByName('LLAMADAS');
    var bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!llamadasSheet || !bddSheet) {
      console.log('Error: No existen las hojas necesarias');
      return;
    }
    
    var datos = llamadasSheet.getDataRange().getValues();
    
    // Encontrar columna "Suma total"
    var colSumaTotal = -1;
    for (var col = 0; col < datos[1].length; col++) {
      if (datos[1][col] === 'Suma total') {
        colSumaTotal = col;
        break;
      }
    }
    
    if (colSumaTotal === -1) {
      console.log('Error: No se encontró columna Suma total');
      return;
    }
    
    // Recopilar ejecutivos
    var ejecutivos = [];
    var filaInicio = 3;
    var filaUltimoEjecutivo = filaInicio;
    
    for (var i = filaInicio; i < datos.length; i++) {
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
      console.log('No hay ejecutivos en la tabla');
      return;
    }
    
    // Obtener columna EJECUTIVO de BBDD_REPORTE
    var headers = bddSheet.getRange(1, 1, 1, bddSheet.getLastColumn()).getValues()[0];
    var ejecutivoIndex = headers.indexOf('EJECUTIVO');
    
    if (ejecutivoIndex === -1) {
      console.log('Error: No se encontró columna EJECUTIVO en BBDD_REPORTE');
      return;
    }
    
    var colEjecutivoBBDD = columnNumberToLetter(ejecutivoIndex + 1);
    
    // UBICACIÓN: DEBAJO DE LA TABLA
    var filaInicioMetricas = filaUltimoEjecutivo + 3;
    var colInicioMetricas = 1;
    
    // Título
    llamadasSheet.getRange(filaInicioMetricas, colInicioMetricas)
      .setValue('📊 PROGRESO DE LLAMADAS POR EJECUTIVO')
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#F3F3F3');
    
    // CREAR TABLA DE MÉTRICAS
    var filaEncabezadosMetricas = filaInicioMetricas + 2;
    
    var encabezadosMetricas = ['EJECUTIVO', 'META', 'LLAMADAS', '% AVANCE', 'PENDIENTE'];
    
    llamadasSheet.getRange(filaEncabezadosMetricas, colInicioMetricas, 1, encabezadosMetricas.length)
      .setValues([encabezadosMetricas])
      .setBackground('#34A853')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Datos de la tabla
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
    
    // Fila de TOTALES
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
    
    // ESPERAR A QUE SE CALCULEN LAS FÓRMULAS
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    
    // LEER VALORES CALCULADOS PARA DEBUG
    console.log('=== VERIFICANDO DATOS PARA EL GRÁFICO ===');
    var datosMetricas = llamadasSheet.getRange(filaEncabezadosMetricas, 1, ejecutivos.length + 1, 5).getValues();
    for (var i = 0; i < Math.min(3, datosMetricas.length); i++) {
      console.log('Fila ' + i + ': ' + JSON.stringify(datosMetricas[i]));
    }
    
    // ELIMINAR GRÁFICOS ANTERIORES
    var charts = llamadasSheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
      llamadasSheet.removeChart(charts[i]);
    }
    console.log('Gráficos anteriores eliminados: ' + charts.length);
    
    // CREAR GRÁFICO DE COLUMNAS AGRUPADAS
    var filaGrafico = filaTotalMetricas + 2;
    
    // Crear rango de datos para el gráfico (sin incluir fila TOTAL)
    var numFilasDatos = ejecutivos.length;
    var filaPrimerDato = filaEncabezadosMetricas + 1;
    
    console.log('Creando gráfico con ' + numFilasDatos + ' ejecutivos');
    console.log('Fila encabezados: ' + filaEncabezadosMetricas);
    console.log('Fila primer dato: ' + filaPrimerDato);
    
    // MÉTODO ALTERNATIVO: Crear DataTable manualmente
    try {
      // Leer los datos calculados
      var rangoDatos = llamadasSheet.getRange(filaEncabezadosMetricas, 1, numFilasDatos + 1, 5);
      var valoresDatos = rangoDatos.getValues();
      
      console.log('Datos leídos para gráfico: ' + valoresDatos.length + ' filas');
      
      // Crear el gráfico con rango único
      var rangoGrafico = llamadasSheet.getRange(
        filaEncabezadosMetricas, 
        1, 
        numFilasDatos + 1, 
        5
      );
      
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
        .setOption('legend', {
          position: 'top',
          textStyle: {fontSize: 12, bold: true}
        })
        .setOption('hAxis', {
          title: 'EJECUTIVO',
          textStyle: {fontSize: 9},
          slantedText: true,
          slantedTextAngle: 45
        })
        .setOption('vAxis', {
          title: 'Cantidad',
          minValue: 0,
          textStyle: {fontSize: 11}
        })
        .setOption('chartArea', {
          width: '70%',
          height: '55%',
          left: 100,
          top: 70,
          bottom: 120
        })
        .setOption('bar', {groupWidth: '60%'})
        .setOption('titleTextStyle', {
          fontSize: 16,
          bold: true
        })
        .setOption('series', {
          1: {  // META (columna B, índice 1)
            type: 'line',
            lineWidth: 2,
            pointSize: 5,
            color: '#34A853'
          },
          2: {  // LLAMADAS (columna C, índice 2)
            type: 'bars',
            color: '#4285F4'
          },
          4: {  // PENDIENTE (columna E, índice 4)
            type: 'bars',
            color: '#EA4335'
          }
        })
        .build();
      
      llamadasSheet.insertChart(chart);
      console.log('✓ Gráfico insertado exitosamente');
      
    } catch (errorChart) {
      console.error('Error creando gráfico: ' + errorChart.message);
      
      // PLAN B: Crear gráfico con rangos específicos
      console.log('Intentando método alternativo...');
      
      var chartAlt = llamadasSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(llamadasSheet.getRange(filaEncabezadosMetricas, 1, numFilasDatos + 1, 1))  // EJECUTIVO
        .addRange(llamadasSheet.getRange(filaEncabezadosMetricas, 3, numFilasDatos + 1, 1))  // LLAMADAS
        .addRange(llamadasSheet.getRange(filaEncabezadosMetricas, 5, numFilasDatos + 1, 1))  // PENDIENTE
        .setPosition(filaGrafico, colInicioMetricas, 0, 0)
        .setNumHeaders(1)
        .setOption('title', 'LLAMADAS y PENDIENTE')
        .setOption('width', 1000)
        .setOption('height', 500)
        .setOption('colors', ['#4285F4', '#EA4335'])
        .setOption('hAxis', {
          slantedText: true,
          slantedTextAngle: 45
        })
        .setOption('vAxis', {
          minValue: 0
        })
        .setOption('chartArea', {
          width: '75%',
          height: '60%'
        })
        .build();
      
      llamadasSheet.insertChart(chartAlt);
      console.log('✓ Gráfico alternativo insertado');
    }
    
    console.log('✓ Gráfico creado exitosamente');
    console.log('Tabla en fila: ' + filaEncabezadosMetricas);
    console.log('Gráfico en fila: ' + filaGrafico);
    
  } catch (error) {
    console.error('❌ Error en crearGraficoProgresoLlamadas: ' + error.message);
    console.error(error.stack);
    throw error;
  }
}

/**
 * Formatea una fecha a string DD-MM-YYYY
 */
function formatearFecha(fecha) {
  if (!(fecha instanceof Date)) return '';
  
  var dia = String(fecha.getDate()).padStart(2, '0');
  var mes = String(fecha.getMonth() + 1).padStart(2, '0');
  var año = fecha.getFullYear();
  
  return dia + '-' + mes + '-' + año;
}

/**
 * Convierte número de columna a letra (A, B, C, ..., Z, AA, AB, etc.)
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