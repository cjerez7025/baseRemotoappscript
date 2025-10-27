/**
 * Crea tabla LLAMADAS con UNIQUE+TRANSPOSE en N2 para fechas dinámicas
 * Las fechas comienzan desde fila 2 de BBDD_REPORTE (sin encabezado)
 * MODIFICADO: Crea la tabla incluso sin fechas registradas
 * BUFFER: 2 columnas extras CON FÓRMULAS para expansión automática de UNIQUE
 */
function crearTablaLlamadas() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      console.log('Error: No existe BBDD_REPORTE');
      return;
    }
    
    let llamadasSheet = spreadsheet.getSheetByName('LLAMADAS');
    if (llamadasSheet) {
      spreadsheet.deleteSheet(llamadasSheet);
    }
    
    llamadasSheet = spreadsheet.insertSheet('LLAMADAS');
    
    const datos = bddSheet.getDataRange().getValues();
    const headers = datos[0];
    
    const ejecutivoIndex = headers.indexOf('EJECUTIVO');
    const fechaIndex = headers.indexOf('FECHA_LLAMADA');
    
    if (ejecutivoIndex === -1 || fechaIndex === -1) {
      console.log('Error: Columnas no encontradas');
      return;
    }
    
    // Recopilar ejecutivos y fechas únicas
    let ejecutivosSet = new Set();
    let fechasSet = new Set();
    
    for (let i = 1; i < datos.length; i++) {
      const ejecutivo = datos[i][ejecutivoIndex];
      const fecha = datos[i][fechaIndex];
      
      if (ejecutivo) {
        const ejec = ejecutivo.toString().trim();
        if (ejec) ejecutivosSet.add(ejec);
      }
      
      if (fecha) {
        const fech = fecha.toString().trim();
        if (fech) fechasSet.add(fech);
      }
    }
    
    const ejecutivos = Array.from(ejecutivosSet).sort();
    const fechas = Array.from(fechasSet).sort((a, b) => new Date(a) - new Date(b));
    
    // CAMBIO PRINCIPAL: Validar solo ejecutivos, las fechas pueden estar vacías
    if (ejecutivos.length === 0) {
      console.log('No hay ejecutivos en BBDD_REPORTE');
      // Crear tabla básica vacía con buffer
      llamadasSheet.getRange(1, 1).setValue('CUENTA de rut_cliente');
      llamadasSheet.getRange(1, 2).setValue('FECHA_LLAMADA');
      llamadasSheet.getRange(2, 1).setValue('EJECUTIVO');
      llamadasSheet.getRange(2, 2).setFormula(`=TRANSPOSE(UNIQUE(FILTER(BBDD_REPORTE!$N$2:$N;BBDD_REPORTE!$N$2:$N<>"")))`);
      // Buffer de 2 columnas + Suma total en columna 5
      llamadasSheet.getRange(2, 5).setValue('Suma total');
      
      // Formato de encabezados
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
    
    // FILA 2: EJECUTIVO + Fechas + Encabezado Suma total
    llamadasSheet.getRange(2, 1).setValue('EJECUTIVO');
    
    // Fórmula UNIQUE+TRANSPOSE que comienza desde fila 2 de BBDD_REPORTE
    const formulaUnique = `=TRANSPOSE(UNIQUE(FILTER(BBDD_REPORTE!$N$2:$N;BBDD_REPORTE!$N$2:$N<>"")))`;
    llamadasSheet.getRange(2, 2).setFormula(formulaUnique);
    
    // Calcular cuántas columnas ocupan las fechas actuales + 2 de buffer
    const numColumnasConDatos = fechas.length > 0 ? fechas.length : 0;
    const numColumnasBuffer = 2;
    const totalColumnasFormulas = numColumnasConDatos + numColumnasBuffer;
    
    // Encabezado "Suma total" después de todas las columnas (datos + buffer)
    const colSumaHeader = 2 + totalColumnasFormulas; // Columna A (1) + B (inicio fechas) + todas las columnas
    llamadasSheet.getRange(2, colSumaHeader).setValue('Suma total');
    
    // FILA 3: Separador
    const maxCols = Math.max(colSumaHeader, 15);
    llamadasSheet.getRange(3, 1, 1, maxCols).setBackground('#E8E8E8');
    
    // FILAS DE DATOS: Ejecutivos + Fórmulas COUNTIFS (incluyendo buffer)
    for (let i = 0; i < ejecutivos.length; i++) {
      const fila = i + 4;
      const ejecutivo = ejecutivos[i];
      
      // Nombre ejecutivo
      llamadasSheet.getRange(fila, 1).setValue(ejecutivo);
      
      // Crear fórmulas para TODAS las columnas de fechas + buffer
      for (let j = 0; j < totalColumnasFormulas; j++) {
        const columna = j + 2; // Empieza en columna B
        const letraCol = columnNumberToLetter(columna);
        
        // Fórmula COUNTIFS que compara con la fecha de fila 2
        const formula = `=COUNTIFS(BBDD_REPORTE!$C$2:$C;$A${fila};BBDD_REPORTE!$N$2:$N;${letraCol}$2)`;
        
        llamadasSheet.getRange(fila, columna).setFormula(formula);
      }
      
      // Suma total por ejecutivo
      const letraFirst = 'B';
      const letraLast = columnNumberToLetter(2 + totalColumnasFormulas - 1);
      const formulaSum = `=SUM(${letraFirst}${fila}:${letraLast}${fila})`;
      llamadasSheet.getRange(fila, colSumaHeader).setFormula(formulaSum);
    }
    
    // FILA DE TOTALES
    const filaTotal = ejecutivos.length + 4;
    llamadasSheet.getRange(filaTotal, 1).setValue('Suma total');
    
    // Totales por cada columna de fecha (incluyendo buffer)
    for (let j = 0; j < totalColumnasFormulas; j++) {
      const columna = j + 2;
      const letraCol = columnNumberToLetter(columna);
      const formula = `=SUM(${letraCol}4:${letraCol}${ejecutivos.length + 3})`;
      llamadasSheet.getRange(filaTotal, columna).setFormula(formula);
    }
    
    // Gran total
    const letraFirst = 'B';
    const letraLast = columnNumberToLetter(2 + totalColumnasFormulas - 1);
    llamadasSheet.getRange(filaTotal, colSumaHeader)
      .setFormula(`=SUM(${letraFirst}${filaTotal}:${letraLast}${filaTotal})`);
    
    // FORMATO
    const maxColsFormat = Math.max(colSumaHeader, 28);
    
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
    
    llamadasSheet.autoResizeColumns(1, maxColsFormat);
    
    console.log('✓ Tabla LLAMADAS creada');
    console.log('Ejecutivos: ' + ejecutivos.length);
    console.log('Fechas actuales: ' + fechas.length);
    console.log('Columnas con fórmulas (datos + buffer): ' + totalColumnasFormulas);
    console.log('Buffer: 2 columnas extras CON FÓRMULAS para expansión automática');
    
  } catch (error) {
    console.error('Error: ' + error.message);
  }
}

/**
 * Convierte número de columna a letra (A, B, C, ..., Z, AA, AB, etc.)
 */
function columnNumberToLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}