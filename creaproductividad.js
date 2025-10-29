/**
 * ARCHIVO COMPLETO: creaproductividad.js - VERSIÓN FINAL CORREGIDA
 * CORRECCIONES APLICADAS:
 * 1. Rangos explícitos ($2:$10000) en lugar de rangos abiertos
 * 2. Referencias dinámicas de columnas (colRefs)
 * 3. GESTIONADO en Tabla 2 referencia a Tabla 1
 * 4. Todas las fórmulas usan columnas detectadas dinámicamente
 */

function crearHojaProductividad() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      console.log('Error: No existe BBDD_REPORTE');
      return;
    }
    
    let prodSheet = spreadsheet.getSheetByName('PRODUCTIVIDAD');
    if (prodSheet) {
      spreadsheet.deleteSheet(prodSheet);
    }
    
    prodSheet = spreadsheet.insertSheet('PRODUCTIVIDAD');
    
    const datos = bddSheet.getDataRange().getValues();
    const headers = datos[0];
    
    console.log('Detectando columnas en BBDD_REPORTE...');
    const ejecutivoIndex = headers.indexOf('EJECUTIVO');
    const estadoIndex = headers.indexOf('ESTADO');
    const clinicaIndex = headers.findIndex(h => /CLINICA|CLINIC|CENTRO/i.test(h));
    
    if (ejecutivoIndex === -1) {
      console.log('Error: No se encontró columna EJECUTIVO');
      return;
    }
    
    if (estadoIndex === -1) {
      console.log('Error: No se encontró columna ESTADO');
      return;
    }
    
    console.log('Columna EJECUTIVO: ' + columnNumberToLetter(ejecutivoIndex + 1));
    console.log('Columna ESTADO: ' + columnNumberToLetter(estadoIndex + 1));
    if (clinicaIndex !== -1) {
      console.log('Columna CLINICA: ' + columnNumberToLetter(clinicaIndex + 1));
    }
    
    // Crear objeto con referencias dinámicas de columnas
    const colRefs = {
      ejecutivo: columnNumberToLetter(ejecutivoIndex + 1),
      estado: columnNumberToLetter(estadoIndex + 1),
      clinica: clinicaIndex !== -1 ? columnNumberToLetter(clinicaIndex + 1) : null
    };
    
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
    
    console.log('Ejecutivos únicos: ' + listaEjecutivos.length);
    console.log('Clínicas únicas: ' + listaClinicas.length);
    
    // TABLA 1: Ejecutivos por Estado
    const filaFinTabla1 = crearTablaEjecutivoPorEstado(prodSheet, listaEjecutivos, estados, 1, 1, colRefs);
    
    // TABLA 2: Métricas por Ejecutivo (referencia a Tabla 1)
    const filaInicioTabla2 = filaFinTabla1 + 3;
    crearTablaMetricasEjecutivo(prodSheet, listaEjecutivos, estados, filaInicioTabla2, 1, colRefs);
    
    // TABLA 3 y 4: Clínicas (si existen)
    if (listaClinicas.length > 0 && colRefs.clinica) {
      crearTablaClinicaPorEstado(prodSheet, listaClinicas, estados, 38, 1, colRefs);
      crearTablaMetricasClinica(prodSheet, listaClinicas, 52, 1, colRefs);
    }
    
    prodSheet.autoResizeColumns(1, 15);
    console.log('Hoja PRODUCTIVIDAD creada exitosamente');
    
  } catch (error) {
    console.error('Error creando PRODUCTIVIDAD:', error.message);
    console.error(error.stack);
  }
}

/**
 * TABLA 1: Ejecutivos por Estado
 */
function crearTablaEjecutivoPorEstado(sheet, ejecutivos, estados, filaInicio, colInicio, colRefs) {
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
      
      const formula = '=COUNTIFS(BBDD_REPORTE!$' + colRefs.ejecutivo + '$2:$' + colRefs.ejecutivo + '$10000;' + 
                      letraEjecutivo + fila + ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;' + 
                      letraEstado + filaInicio + ')';
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
 * CORRECCIÓN: GESTIONADO referencia a suma de columnas de Tabla 1
 */
function crearTablaMetricasEjecutivo(sheet, ejecutivos, estados, filaInicio, colInicio, colRefs) {
  const encabezados = ['EJECUTIVO', 'GESTIONADO', 'META', 'AVANCE', 'CONTACTADO', '% CONTACTADO', 'INTERESADO', '% INTERESADO', 'CERRADO', 'RENDIMIENTO'];
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  const filaInicioTabla1 = 2; // Tabla 1 empieza en fila 2 (después del encabezado en fila 1)
  
  ejecutivos.forEach((ejecutivo, index) => {
    const fila = filaInicio + 1 + index;
    const filaTabla1 = filaInicioTabla1 + index;
    const letraEjecutivo = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(ejecutivo);
    
    // GESTIONADO: Suma de columnas B+C+D+E+G de Tabla 1
    // B=Cerrado, C=En Gestión, D=Interesado, E=No Contactado, G=Sin Interés
    const formulaGestionados = '=B' + filaTabla1 + '+C' + filaTabla1 + '+D' + filaTabla1 + '+E' + filaTabla1 + '+G' + filaTabla1;
    sheet.getRange(fila, colInicio + 1).setFormula(formulaGestionados);
    
    // META
    sheet.getRange(fila, colInicio + 2).setFormula('=COUNTIF(BBDD_REPORTE!$' + colRefs.ejecutivo + '$2:$' + colRefs.ejecutivo + '$10000;' + letraEjecutivo + fila + ')');
    
    // AVANCE
    const letraGestionado = columnNumberToLetter(colInicio + 1);
    const letraMeta = columnNumberToLetter(colInicio + 2);
    sheet.getRange(fila, colInicio + 3).setFormula('=' + letraGestionado + fila + '/' + letraMeta + fila);
    sheet.getRange(fila, colInicio + 3).setNumberFormat('0%');
    
    // CONTACTADO: Suma B+C+D de Tabla 1 (Cerrado + En Gestión + Interesado)
    sheet.getRange(fila, colInicio + 4).setFormula('=B' + filaTabla1 + '+C' + filaTabla1 + '+D' + filaTabla1);
    
    // % CONTACTADO
    const letraContactado = columnNumberToLetter(colInicio + 4);
    sheet.getRange(fila, colInicio + 5).setFormula('=' + letraContactado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 5).setNumberFormat('0%');
    
    // INTERESADO: Columna D de Tabla 1
    sheet.getRange(fila, colInicio + 6).setFormula('=D' + filaTabla1);
    
    // % INTERESADO
    const letraInteresado = columnNumberToLetter(colInicio + 6);
    sheet.getRange(fila, colInicio + 7).setFormula('=' + letraInteresado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 7).setNumberFormat('0%');
    
    // CERRADO: Columna B de Tabla 1
    sheet.getRange(fila, colInicio + 8).setFormula('=B' + filaTabla1);
    
    // RENDIMIENTO
    const letraCerrado = columnNumberToLetter(colInicio + 8);
    sheet.getRange(fila, colInicio + 9).setFormula('=' + letraCerrado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 9).setNumberFormat('0%');
  });
  
  // Fila de TOTALES
  const filaTotal = filaInicio + ejecutivos.length + 1;
  sheet.getRange(filaTotal, colInicio).setValue('Total General');
  
  // Sumar columnas GESTIONADO, CONTACTADO, INTERESADO, CERRADO
  [1, 4, 6, 8].forEach(offset => {
    const col = colInicio + offset;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotal, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + ejecutivos.length) + ')');
  });
  
  // Formato final
  const rangoTotales = sheet.getRange(filaTotal, colInicio, 1, encabezados.length);
  rangoTotales.setBackground('#4472C4');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, ejecutivos.length, encabezados.length - 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, ejecutivos.length + 2, encabezados.length).setBorder(true, true, true, true, true, true);
}

/**
 * TABLA 3: Clínicas por Estado
 */
function crearTablaClinicaPorEstado(sheet, clinicas, estados, filaInicio, colInicio, colRefs) {
  if (!colRefs.clinica) return;
  
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
  
  clinicas.forEach((clinica, index) => {
    const fila = filaInicio + 1 + index;
    const letraClinica = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(clinica);
    
    estados.forEach((estado, estadoIndex) => {
      const col = colInicio + 1 + estadoIndex;
      const letraEstado = columnNumberToLetter(colInicio + 1 + estadoIndex);
      
      const formula = '=COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + 
                      letraClinica + fila + ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;' + 
                      letraEstado + filaInicio + ')';
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
function crearTablaMetricasClinica(sheet, clinicas, filaInicio, colInicio, colRefs) {
  if (!colRefs.clinica) return;
  
  const encabezados = ['CLINICA', 'GESTIONADO', 'META', 'AVANCE', 'CONTACTADO', '% CONTACTADO', 'INTERESADO', '% INTERESADO', 'CERRADO', 'RENDIMIENTO'];
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  clinicas.forEach((clinica, index) => {
    const fila = filaInicio + 1 + index;
    const letraClinica = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(clinica);
    
    // GESTIONADO
    sheet.getRange(fila, colInicio + 1).setFormula('=COUNTIF(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + letraClinica + fila + ')');
    
    // META
    sheet.getRange(fila, colInicio + 2).setValue(1);
    
    // AVANCE
    const letraGestionado = columnNumberToLetter(colInicio + 1);
    const letraMeta = columnNumberToLetter(colInicio + 2);
    sheet.getRange(fila, colInicio + 3).setFormula('=' + letraGestionado + fila + '/' + letraMeta + fila);
    sheet.getRange(fila, colInicio + 3).setNumberFormat('0%');
    
    // CONTACTADO
    sheet.getRange(fila, colInicio + 4).setFormula(
      '=COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + letraClinica + fila + 
      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;"Cerrado")+' +
      'COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + letraClinica + fila + 
      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;"En Gestión")+' +
      'COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + letraClinica + fila + 
      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;"Interesado")'
    );
    
    // % CONTACTADO
    const letraContactado = columnNumberToLetter(colInicio + 4);
    sheet.getRange(fila, colInicio + 5).setFormula('=' + letraContactado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 5).setNumberFormat('0%');
    
    // INTERESADO
    sheet.getRange(fila, colInicio + 6).setFormula(
      '=COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + letraClinica + fila + 
      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;"Interesado")'
    );
    
    // % INTERESADO
    const letraInteresado = columnNumberToLetter(colInicio + 6);
    sheet.getRange(fila, colInicio + 7).setFormula('=' + letraInteresado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 7).setNumberFormat('0%');
    
    // CERRADO
    sheet.getRange(fila, colInicio + 8).setFormula(
      '=COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + '$10000;' + letraClinica + fila + 
      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + '$10000;"Cerrado")'
    );
    
    // RENDIMIENTO
    const letraCerrado = columnNumberToLetter(colInicio + 8);
    sheet.getRange(fila, colInicio + 9).setFormula('=' + letraCerrado + fila + '/' + letraGestionado + fila);
    sheet.getRange(fila, colInicio + 9).setNumberFormat('0%');
  });
  
  sheet.getRange(filaInicio + 1, colInicio + 1, clinicas.length, encabezados.length - 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, clinicas.length + 1, encabezados.length).setBorder(true, true, true, true, true, true);
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