/**
 * Crea la hoja PRODUCTIVIDAD con 4 tablas de análisis
 * Basado en datos de BBDD_REPORTE
 * TABLA 2 referencia a TABLA 1 para calcular GESTIONADOS
 */
function crearHojaProductividad() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      console.log('Error: No existe BBDD_REPORTE');
      return;
    }
    
    // Eliminar hoja existente si existe
    let prodSheet = spreadsheet.getSheetByName('PRODUCTIVIDAD');
    if (prodSheet) {
      spreadsheet.deleteSheet(prodSheet);
    }
    
    // Crear nueva hoja
    prodSheet = spreadsheet.insertSheet('PRODUCTIVIDAD');
    
    const datos = bddSheet.getDataRange().getValues();
    const headers = datos[0];
    
    // Obtener índices de columnas
    const ejecutivoIndex = headers.indexOf('EJECUTIVO');
    const estadoIndex = headers.indexOf('ESTADO');
    const clinicaIndex = headers.findIndex(h => /CLINICA|CLINIC|CENTRO/i.test(h));
    
    if (ejecutivoIndex === -1 || estadoIndex === -1) {
      console.log('Error: Columnas necesarias no encontradas');
      return;
    }
    
    // Recopilar datos únicos
    const ejecutivos = new Set();
    const clinicas = new Set();
    const estados = ['Cerrado', 'En Gestión', 'Interesado', 'No Contactado', 'Sin Gestión', 'Sin Interés'];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][ejecutivoIndex]) {
        ejecutivos.add(datos[i][ejecutivoIndex].toString().trim());
      }
      if (clinicaIndex !== -1 && datos[i][clinicaIndex]) {
        clinicas.add(datos[i][clinicaIndex].toString().trim());
      }
    }
    
    const listaEjecutivos = Array.from(ejecutivos).sort();
    const listaClinicas = Array.from(clinicas).sort();
    
    // TABLA 1: Ejecutivos por Estado (Fila 1, Columna A)
    const filaFinTabla1 = crearTablaEjecutivoPorEstado(prodSheet, listaEjecutivos, estados, 1, 1);
    
    // TABLA 2: Métricas por Ejecutivo (Después de TABLA 1 + 3 filas de espacio)
    const filaInicioTabla2 = filaFinTabla1 + 3;
    crearTablaMetricasEjecutivo(prodSheet, listaEjecutivos, estados, filaInicioTabla2, 1);
    
    // TABLA 3: Clínica por Estado
    if (listaClinicas.length > 0) {
      crearTablaClinicaPorEstado(prodSheet, listaClinicas, estados, 38, 1);
      crearTablaMetricasClinica(prodSheet, listaClinicas, 52, 1);
    }
    
    // Ajustar anchos de columna
    prodSheet.autoResizeColumns(1, 15);
    
    console.log('✓ Hoja PRODUCTIVIDAD creada exitosamente');
    
  } catch (error) {
    console.error('Error creando PRODUCTIVIDAD:', error.message);
  }
}

/**
 * TABLA 1: Ejecutivos por Estado
 */
function crearTablaEjecutivoPorEstado(sheet, ejecutivos, estados, filaInicio, colInicio) {
  sheet.getRange(filaInicio, colInicio).setValue('EJECUTIVO');
  
  estados.forEach((estado, index) => {
    sheet.getRange(filaInicio, colInicio + 1 + index).setValue(estado);
  });
  
  sheet.getRange(filaInicio, colInicio + estados.length + 1).setValue('Suma total');
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, estados.length + 2);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  ejecutivos.forEach((ejecutivo, index) => {
    const fila = filaInicio + 1 + index;
    const letraEjecutivo = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(ejecutivo);
    
    estados.forEach((estado, estadoIndex) => {
      const col = colInicio + 1 + estadoIndex;
      const letraEstado = columnNumberToLetter(colInicio + 1 + estadoIndex);
      const formula = '=COUNTIFS(BBDD_REPORTE!$C$2:$C;' + letraEjecutivo + fila + ';BBDD_REPORTE!$P$2:$P;' + letraEstado + filaInicio + ')';
      sheet.getRange(fila, col).setFormula(formula);
    });
    
    const colSuma = colInicio + estados.length + 1;
    const letraInicio = columnNumberToLetter(colInicio + 1);
    const letraFin = columnNumberToLetter(colInicio + estados.length);
    sheet.getRange(fila, colSuma).setFormula('=SUM(' + letraInicio + fila + ':' + letraFin + fila + ')');
  });
  
  const filaTotales = filaInicio + ejecutivos.length + 1;
  sheet.getRange(filaTotales, colInicio).setValue('TOTALES');
  
  estados.forEach((estado, estadoIndex) => {
    const col = colInicio + 1 + estadoIndex;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotales, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + ejecutivos.length) + ')');
  });
  
  const colSuma = colInicio + estados.length + 1;
  const letraInicio = columnNumberToLetter(colInicio + 1);
  const letraFin = columnNumberToLetter(colInicio + estados.length);
  sheet.getRange(filaTotales, colSuma).setFormula('=SUM(' + letraInicio + filaTotales + ':' + letraFin + filaTotales + ')');
  
  const rangoTotales = sheet.getRange(filaTotales, colInicio, 1, estados.length + 2);
  rangoTotales.setBackground('#4472C4');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, ejecutivos.length, estados.length + 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, ejecutivos.length + 2, estados.length + 2).setBorder(true, true, true, true, true, true);
  
  return filaTotales;
}

/**
 * TABLA 2: Métricas por Ejecutivo
 */
function crearTablaMetricasEjecutivo(sheet, ejecutivos, estados, filaInicio, colInicio) {
  const encabezados = ['EJECUTIVO', 'GESTIONADOS', 'META', 'AVANCE', 'CONTACTADOS', '% CONTACTADO', 'INTERESADO', '% INTERESADO', 'CERRADO', 'RENDIMIENTO'];
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  const filaInicioTabla1 = 2;
  
  ejecutivos.forEach((ejecutivo, index) => {
    const fila = filaInicio + 1 + index;
    const letraEjecutivo = columnNumberToLetter(colInicio);
    const filaTabla1 = filaInicioTabla1 + index;
    
    sheet.getRange(fila, colInicio).setValue(ejecutivo);
    
    // GESTIONADOS - Suma simple sin SI
    const formulaGestionados = '=B' + filaTabla1 + '+C' + filaTabla1 + '+D' + filaTabla1 + '+E' + filaTabla1 + '+G' + filaTabla1;
    sheet.getRange(fila, colInicio + 1).setFormula(formulaGestionados);
    
    // META
    sheet.getRange(fila, colInicio + 2).setFormula('=COUNTIF(BBDD_REPORTE!$C$2:$C;' + letraEjecutivo + fila + ')');
    
    // AVANCE
    const letraGestionado = columnNumberToLetter(colInicio + 1);
    const letraMeta = columnNumberToLetter(colInicio + 2);
    sheet.getRange(fila, colInicio + 3).setFormula('=' + letraGestionado + fila + '/' + letraMeta + fila);
    sheet.getRange(fila, colInicio + 3).setNumberFormat('0%');
    
    // CONTACTADOS
    sheet.getRange(fila, colInicio + 4).setFormula('=B' + filaTabla1 + '+C' + filaTabla1 + '+D' + filaTabla1);
    
    // % CONTACTADO
    const letraContactado = columnNumberToLetter(colInicio + 4);
    sheet.getRange(fila, colInicio + 5).setFormula('=' + letraContactado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 5).setNumberFormat('0%');
    
    // INTERESADO
    sheet.getRange(fila, colInicio + 6).setFormula('=D' + filaTabla1);
    
    // % INTERESADO
    const letraInteresado = columnNumberToLetter(colInicio + 6);
    sheet.getRange(fila, colInicio + 7).setFormula('=' + letraInteresado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 7).setNumberFormat('0%');
    
    // CERRADO
    sheet.getRange(fila, colInicio + 8).setFormula('=B' + filaTabla1);
    
    // RENDIMIENTO
    const letraCerrado = columnNumberToLetter(colInicio + 8);
    sheet.getRange(fila, colInicio + 9).setFormula('=' + letraCerrado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 9).setNumberFormat('0.00%');
  });
  
  const filaTotal = filaInicio + ejecutivos.length + 1;
  sheet.getRange(filaTotal, colInicio).setValue('Total General');
  
  [1, 4, 6, 8].forEach(offset => {
    const col = colInicio + offset;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotal, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + ejecutivos.length) + ')');
  });
  
  const rangoTotal = sheet.getRange(filaTotal, colInicio, 1, encabezados.length);
  rangoTotal.setBackground('#4472C4');
  rangoTotal.setFontColor('white');
  rangoTotal.setFontWeight('bold');
  rangoTotal.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, ejecutivos.length, encabezados.length - 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, ejecutivos.length + 2, encabezados.length).setBorder(true, true, true, true, true, true);
}

/**
 * TABLA 3: Clínica por Estado
 */
function crearTablaClinicaPorEstado(sheet, clinicas, estados, filaInicio, colInicio) {
  sheet.getRange(filaInicio, colInicio).setValue('CLINICA');
  
  estados.forEach((estado, index) => {
    sheet.getRange(filaInicio, colInicio + 1 + index).setValue(estado);
  });
  
  sheet.getRange(filaInicio, colInicio + estados.length + 1).setValue('Suma total');
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, estados.length + 2);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  const bddSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BBDD_REPORTE');
  const headers = bddSheet.getRange(1, 1, 1, bddSheet.getLastColumn()).getValues()[0];
  const clinicaCol = headers.findIndex(h => /CLINICA|CLINIC|CENTRO/i.test(h));
  const clinicaLetra = columnNumberToLetter(clinicaCol + 1);
  
  clinicas.forEach((clinica, index) => {
    const fila = filaInicio + 1 + index;
    const letraClinica = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(clinica);
    
    estados.forEach((estado, estadoIndex) => {
      const col = colInicio + 1 + estadoIndex;
      const letraEstado = columnNumberToLetter(colInicio + 1 + estadoIndex);
      const formula = '=COUNTIFS(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ';BBDD_REPORTE!$P$2:$P;' + letraEstado + filaInicio + ')';
      sheet.getRange(fila, col).setFormula(formula);
    });
    
    const colSuma = colInicio + estados.length + 1;
    const letraInicio = columnNumberToLetter(colInicio + 1);
    const letraFin = columnNumberToLetter(colInicio + estados.length);
    sheet.getRange(fila, colSuma).setFormula('=SUM(' + letraInicio + fila + ':' + letraFin + fila + ')');
  });
  
  const filaTotales = filaInicio + clinicas.length + 1;
  sheet.getRange(filaTotales, colInicio).setValue('TOTALES');
  
  estados.forEach((estado, estadoIndex) => {
    const col = colInicio + 1 + estadoIndex;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotales, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + clinicas.length) + ')');
  });
  
  const colSuma = colInicio + estados.length + 1;
  const letraInicio = columnNumberToLetter(colInicio + 1);
  const letraFin = columnNumberToLetter(colInicio + estados.length);
  sheet.getRange(filaTotales, colSuma).setFormula('=SUM(' + letraInicio + filaTotales + ':' + letraFin + filaTotales + ')');
  
  const rangoTotales = sheet.getRange(filaTotales, colInicio, 1, estados.length + 2);
  rangoTotales.setBackground('#4472C4');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, clinicas.length, estados.length + 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, clinicas.length + 2, estados.length + 2).setBorder(true, true, true, true, true, true);
}

/**
 * TABLA 4: Métricas por Clínica
 */
function crearTablaMetricasClinica(sheet, clinicas, filaInicio, colInicio) {
  const encabezados = ['CLINICA', 'GESTIONADO', 'META', 'AVANCE', 'CONTACTADO', '% CONTACTADO', 'INTERESADO', '% INTERESADO', 'CERRADO', 'RENDIMIENTO'];
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  const bddSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BBDD_REPORTE');
  const headers = bddSheet.getRange(1, 1, 1, bddSheet.getLastColumn()).getValues()[0];
  const clinicaCol = headers.findIndex(h => /CLINICA|CLINIC|CENTRO/i.test(h));
  const clinicaLetra = columnNumberToLetter(clinicaCol + 1);
  
  clinicas.forEach((clinica, index) => {
    const fila = filaInicio + 1 + index;
    const letraClinica = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(clinica);
    
    sheet.getRange(fila, colInicio + 1).setFormula('=COUNTIF(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ')');
    sheet.getRange(fila, colInicio + 2).setValue(1);
    
    const letraGestionado = columnNumberToLetter(colInicio + 1);
    const letraMeta = columnNumberToLetter(colInicio + 2);
    sheet.getRange(fila, colInicio + 3).setFormula('=' + letraGestionado + fila + '/' + letraMeta + fila);
    
    sheet.getRange(fila, colInicio + 4).setFormula('=COUNTIFS(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ';BBDD_REPORTE!$P$2:$P;"Cerrado")+COUNTIFS(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ';BBDD_REPORTE!$P$2:$P;"En Gestión")+COUNTIFS(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ';BBDD_REPORTE!$P$2:$P;"Interesado")');
    
    const letraContactado = columnNumberToLetter(colInicio + 4);
    sheet.getRange(fila, colInicio + 5).setFormula('=' + letraContactado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 5).setNumberFormat('0%');
    
    sheet.getRange(fila, colInicio + 6).setFormula('=COUNTIFS(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ';BBDD_REPORTE!$P$2:$P;"Interesado")');
    
    const letraInteresado = columnNumberToLetter(colInicio + 6);
    sheet.getRange(fila, colInicio + 7).setFormula('=' + letraInteresado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 7).setNumberFormat('0%');
    
    sheet.getRange(fila, colInicio + 8).setFormula('=COUNTIFS(BBDD_REPORTE!' + clinicaLetra + '$2:' + clinicaLetra + ';' + letraClinica + fila + ';BBDD_REPORTE!$P$2:$P;"Cerrado")');
    
    const letraCerrado = columnNumberToLetter(colInicio + 8);
    sheet.getRange(fila, colInicio + 9).setFormula('=' + letraCerrado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 9).setNumberFormat('0.00%');
  });
  
  const filaTotal = filaInicio + clinicas.length + 1;
  sheet.getRange(filaTotal, colInicio).setValue('Total General');
  
  [1, 4, 6, 8].forEach(offset => {
    const col = colInicio + offset;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotal, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + clinicas.length) + ')');
  });
  
  const rangoTotal = sheet.getRange(filaTotal, colInicio, 1, encabezados.length);
  rangoTotal.setBackground('#4472C4');
  rangoTotal.setFontColor('white');
  rangoTotal.setFontWeight('bold');
  rangoTotal.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, clinicas.length, encabezados.length - 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, clinicas.length + 2, encabezados.length).setBorder(true, true, true, true, true, true);
}

/**
 * Convierte número de columna a letra
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