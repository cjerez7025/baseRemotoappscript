/**
 * CREAPRODUCTIVIDAD.JS - VERSIÓN CORREGIDA
 * Sin errores de sintaxis
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
    
    const ejecutivoIndex = headers.indexOf('EJECUTIVO');
    const estadoIndex = headers.indexOf('ESTADO');
    const clinicaIndex = headers.findIndex(h => /CLINICA|CLINIC|CENTRO/i.test(h));
    
    if (ejecutivoIndex === -1 || estadoIndex === -1) {
      console.log('Error: Columnas requeridas no encontradas');
      return;
    }
    
    const colRefs = {
      ejecutivo: columnNumberToLetter(ejecutivoIndex + 1),
      estado: columnNumberToLetter(estadoIndex + 1),
      clinica: clinicaIndex !== -1 ? columnNumberToLetter(clinicaIndex + 1) : null
    };
    
    const ejecutivosSet = new Set();
    for (let i = 1; i < datos.length; i++) {
      const ejecutivo = datos[i][ejecutivoIndex];
      if (ejecutivo && ejecutivo.toString().trim() !== '') {
        ejecutivosSet.add(ejecutivo.toString().trim());
      }
    }
    const ejecutivos = Array.from(ejecutivosSet).sort();
    
    const estadosSet = new Set();
    for (let i = 1; i < datos.length; i++) {
      const estado = datos[i][estadoIndex];
      if (estado && estado.toString().trim() !== '') {
        estadosSet.add(estado.toString().trim());
      }
    }
    const estados = Array.from(estadosSet).sort();
    
    let filaActual = 2;
    
    const filaTotalTabla1 = crearTablaConteoEstado(prodSheet, ejecutivos, estados, filaActual, 1, colRefs);
    
    filaActual = filaTotalTabla1 + 3;
    const filaTotalTabla2 = crearTablaMetricasEjecutivo(prodSheet, ejecutivos, estados, filaActual, 1, colRefs);
    
    if (colRefs.clinica) {
      const clinicasSet = new Set();
      for (let i = 1; i < datos.length; i++) {
        const clinica = datos[i][clinicaIndex];
        if (clinica && clinica.toString().trim() !== '') {
          clinicasSet.add(clinica.toString().trim());
        }
      }
      const clinicas = Array.from(clinicasSet).sort();
      
      filaActual = filaTotalTabla2 + 3;
      crearTablaEstadosClinica(prodSheet, clinicas, estados, filaActual, 1, colRefs);
      
      filaActual = filaActual + clinicas.length + 4;
      crearTablaMetricasClinica(prodSheet, clinicas, filaActual, 1, colRefs);
    }
    
    prodSheet.autoResizeColumns(1, 10);
    
    SpreadsheetApp.getUi().alert('Éxito', 'Hoja PRODUCTIVIDAD creada correctamente', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.log('Error: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function crearTablaConteoEstado(sheet, ejecutivos, estados, filaInicio, colInicio, colRefs) {
  const estadosOrdenados = ['Cerrado', 'En Gestión', 'Interesado', 'No Contactado', 'Sin Gestión', 'Sin Interés'];
  const encabezados = ['EJECUTIVO'].concat(estadosOrdenados).concat(['Suma total']);
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#FFD700');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  ejecutivos.forEach((ejecutivo, index) => {
    const fila = filaInicio + 1 + index;
    const letraEjecutivo = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(ejecutivo);
    
    estadosOrdenados.forEach((estado, estadoIndex) => {
      const col = colInicio + 1 + estadoIndex;
      const formula = '=COUNTIFS(BBDD_REPORTE!$' + colRefs.ejecutivo + '$2:$' + colRefs.ejecutivo + 
                      ';' + letraEjecutivo + fila + 
                      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + ';"' + estado + '")';
      sheet.getRange(fila, col).setFormula(formula);
    });
    
    const colSuma = colInicio + estadosOrdenados.length + 1;
    const letraFirst = columnNumberToLetter(colInicio + 1);
    const letraLast = columnNumberToLetter(colInicio + estadosOrdenados.length);
    sheet.getRange(fila, colSuma).setFormula('=SUM(' + letraFirst + fila + ':' + letraLast + fila + ')');
  });
  
  const filaTotales = filaInicio + ejecutivos.length + 1;
  sheet.getRange(filaTotales, colInicio).setValue('Total General');
  
  for (let i = 1; i <= estadosOrdenados.length + 1; i++) {
    const col = colInicio + i;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotales, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + ejecutivos.length) + ')');
  }
  
  const rangoTotales = sheet.getRange(filaTotales, colInicio, 1, estadosOrdenados.length + 2);
  rangoTotales.setBackground('#4472C4');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, ejecutivos.length, estadosOrdenados.length + 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, ejecutivos.length + 2, estadosOrdenados.length + 2).setBorder(true, true, true, true, true, true);
  
  return filaTotales;
}

function crearTablaMetricasEjecutivo(sheet, ejecutivos, estados, filaInicio, colInicio, colRefs) {
  const encabezados = ['EJECUTIVO', 'GESTIONADO', 'META', 'AVANCE', 'CONTACTADO', 
                       '% CONTACTADO', 'INTERESADO', '% INTERESADO', 'CERRADO', 'RENDIMIENTO'];
  const estadosOrdenados = ['Cerrado', 'En Gestión', 'Interesado', 'No Contactado', 'Sin Gestión', 'Sin Interés'];
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#4472C4');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  ejecutivos.forEach((ejecutivo, index) => {
    const fila = filaInicio + 1 + index;
    const letraEj = columnNumberToLetter(colInicio);
    const colEj = colRefs.ejecutivo;
    const colEst = colRefs.estado;
    
    // GESTIONADO
    let formula = '=COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
                  ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"En Gestión")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Interesado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"No Contactado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Sin Interés")';
    sheet.getRange(fila, colInicio + 1).setFormula(formula);
    
    // META
    const letraSuma = columnNumberToLetter(colInicio + estadosOrdenados.length + 1);
    sheet.getRange(fila, colInicio + 2).setFormula('=' + letraSuma + (index + 3));
    
    // AVANCE
    const letraGest = columnNumberToLetter(colInicio + 1);
    const letraMeta = columnNumberToLetter(colInicio + 2);
    sheet.getRange(fila, colInicio + 3).setFormula('=IF(' + letraMeta + fila + '=0;0%;' + letraGest + fila + '/' + letraMeta + fila + ')');
    
    // CONTACTADO
    formula = '=COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
              ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"En Gestión")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Interesado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Sin Interés")';
    sheet.getRange(fila, colInicio + 4).setFormula(formula);
    
    // % CONTACTADO
    const letraCont = columnNumberToLetter(colInicio + 4);
    sheet.getRange(fila, colInicio + 5).setFormula('=IF(' + letraMeta + fila + '=0;0%;' + letraCont + fila + '/' + letraMeta + fila + ')');
    
    // INTERESADO
    formula = '=COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
              ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Interesado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    sheet.getRange(fila, colInicio + 6).setFormula(formula);
    
    // % INTERESADO
    const letraInt = columnNumberToLetter(colInicio + 6);
    sheet.getRange(fila, colInicio + 7).setFormula('=IF(' + letraCont + fila + '=0;0%;' + letraInt + fila + '/' + letraCont + fila + ')');
    
    // CERRADO
    formula = '=COUNTIFS(BBDD_REPORTE!$' + colEj + '$2:$' + colEj + ';' + letraEj + fila + 
              ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    sheet.getRange(fila, colInicio + 8).setFormula(formula);
    
    // RENDIMIENTO
    const letraCerr = columnNumberToLetter(colInicio + 8);
    sheet.getRange(fila, colInicio + 9).setFormula('=IF(' + letraCont + fila + '=0;0%;' + letraCerr + fila + '/' + letraCont + fila + ')');
  });
  
  const filaTotales = filaInicio + ejecutivos.length + 1;
  sheet.getRange(filaTotales, colInicio).setValue('Total General');
  
  [1, 2, 4, 6, 8].forEach(colOffset => {
    const col = colInicio + colOffset;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotales, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + ejecutivos.length) + ')');
  });
  
  const letraGestT = columnNumberToLetter(colInicio + 1);
  const letraMetaT = columnNumberToLetter(colInicio + 2);
  const letraContT = columnNumberToLetter(colInicio + 4);
  const letraIntT = columnNumberToLetter(colInicio + 6);
  const letraCerrT = columnNumberToLetter(colInicio + 8);
  
  sheet.getRange(filaTotales, colInicio + 3).setFormula('=IF(' + letraMetaT + filaTotales + '=0;0%;' + letraGestT + filaTotales + '/' + letraMetaT + filaTotales + ')');
  sheet.getRange(filaTotales, colInicio + 5).setFormula('=IF(' + letraMetaT + filaTotales + '=0;0%;' + letraContT + filaTotales + '/' + letraMetaT + filaTotales + ')');
  sheet.getRange(filaTotales, colInicio + 7).setFormula('=IF(' + letraContT + filaTotales + '=0;0%;' + letraIntT + filaTotales + '/' + letraContT + filaTotales + ')');
  sheet.getRange(filaTotales, colInicio + 9).setFormula('=IF(' + letraContT + filaTotales + '=0;0%;' + letraCerrT + filaTotales + '/' + letraContT + filaTotales + ')');
  
  const rangoTotales = sheet.getRange(filaTotales, colInicio, 1, 10);
  rangoTotales.setBackground('#4472C4');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 3, ejecutivos.length + 1, 1).setNumberFormat('0%');
  sheet.getRange(filaInicio + 1, colInicio + 5, ejecutivos.length + 1, 1).setNumberFormat('0%');
  sheet.getRange(filaInicio + 1, colInicio + 7, ejecutivos.length + 1, 1).setNumberFormat('0%');
  sheet.getRange(filaInicio + 1, colInicio + 9, ejecutivos.length + 1, 1).setNumberFormat('0%');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, ejecutivos.length, 9).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, ejecutivos.length + 2, 10).setBorder(true, true, true, true, true, true);
  
  return filaTotales;
}

function crearTablaEstadosClinica(sheet, clinicas, estados, filaInicio, colInicio, colRefs) {
  if (!colRefs.clinica) return;
  
  const estadosOrdenados = ['Cerrado', 'En Gestión', 'Interesado', 'No Contactado', 'Sin Gestión', 'Sin Interés'];
  const encabezados = ['CLINICA'].concat(estadosOrdenados).concat(['Suma total']);
  
  encabezados.forEach((encabezado, index) => {
    sheet.getRange(filaInicio, colInicio + index).setValue(encabezado);
  });
  
  const rangoEncabezados = sheet.getRange(filaInicio, colInicio, 1, encabezados.length);
  rangoEncabezados.setBackground('#00B0F0');
  rangoEncabezados.setFontColor('white');
  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setHorizontalAlignment('center');
  
  clinicas.forEach((clinica, index) => {
    const fila = filaInicio + 1 + index;
    const letraCli = columnNumberToLetter(colInicio);
    
    sheet.getRange(fila, colInicio).setValue(clinica);
    
    estadosOrdenados.forEach((estado, estadoIndex) => {
      const col = colInicio + 1 + estadoIndex;
      const formula = '=COUNTIFS(BBDD_REPORTE!$' + colRefs.clinica + '$2:$' + colRefs.clinica + 
                      ';' + letraCli + fila + 
                      ';BBDD_REPORTE!$' + colRefs.estado + '$2:$' + colRefs.estado + ';"' + estado + '")';
      sheet.getRange(fila, col).setFormula(formula);
    });
    
    const colSuma = colInicio + estadosOrdenados.length + 1;
    const letraFirst = columnNumberToLetter(colInicio + 1);
    const letraLast = columnNumberToLetter(colInicio + estadosOrdenados.length);
    sheet.getRange(fila, colSuma).setFormula('=SUM(' + letraFirst + fila + ':' + letraLast + fila + ')');
  });
  
  const filaTotales = filaInicio + clinicas.length + 1;
  sheet.getRange(filaTotales, colInicio).setValue('Total General');
  
  for (let i = 1; i <= estadosOrdenados.length + 1; i++) {
    const col = colInicio + i;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotales, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + clinicas.length) + ')');
  }
  
  const rangoTotales = sheet.getRange(filaTotales, colInicio, 1, estadosOrdenados.length + 2);
  rangoTotales.setBackground('#00B0F0');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, clinicas.length, estadosOrdenados.length + 1).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, clinicas.length + 2, estadosOrdenados.length + 2).setBorder(true, true, true, true, true, true);
}

function crearTablaMetricasClinica(sheet, clinicas, filaInicio, colInicio, colRefs) {
  if (!colRefs.clinica) return;
  
  const encabezados = ['CLINICA', 'GESTIONADO', 'META', 'AVANCE', 'CONTACTADO', 
                       '% CONTACTADO', 'INTERESADO', '% INTERESADO', 'CERRADO', 'RENDIMIENTO'];
  
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
    const letraCli = columnNumberToLetter(colInicio);
    const colCli = colRefs.clinica;
    const colEst = colRefs.estado;
    
    // GESTIONADO
    let formula = '=COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
                  ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"En Gestión")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Interesado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"No Contactado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Sin Interés")';
    sheet.getRange(fila, colInicio + 1).setFormula(formula);
    
    // META
    formula = '=COUNTIF(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + ')';
    sheet.getRange(fila, colInicio + 2).setFormula(formula);
    
    // AVANCE
    const letraGest = columnNumberToLetter(colInicio + 1);
    const letraMeta = columnNumberToLetter(colInicio + 2);
    sheet.getRange(fila, colInicio + 3).setFormula('=IF(' + letraMeta + fila + '=0;0%;' + letraGest + fila + '/' + letraMeta + fila + ')');
    
    // CONTACTADO
    formula = '=COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
              ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"En Gestión")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Interesado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Sin Interés")';
    sheet.getRange(fila, colInicio + 4).setFormula(formula);
    
    // % CONTACTADO
    const letraCont = columnNumberToLetter(colInicio + 4);
    sheet.getRange(fila, colInicio + 5).setFormula('=IF(' + letraMeta + fila + '=0;0%;' + letraCont + fila + '/' + letraMeta + fila + ')');
    
    // INTERESADO
    formula = '=COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
              ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Interesado")';
    formula += '+COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
               ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    sheet.getRange(fila, colInicio + 6).setFormula(formula);
    
    // % INTERESADO
    const letraInt = columnNumberToLetter(colInicio + 6);
    sheet.getRange(fila, colInicio + 7).setFormula('=IF(' + letraCont + fila + '=0;0%;' + letraInt + fila + '/' + letraCont + fila + ')');
    
    // CERRADO
    formula = '=COUNTIFS(BBDD_REPORTE!$' + colCli + '$2:$' + colCli + ';' + letraCli + fila + 
              ';BBDD_REPORTE!$' + colEst + '$2:$' + colEst + ';"Cerrado")';
    sheet.getRange(fila, colInicio + 8).setFormula(formula);
    
    // RENDIMIENTO
    const letraCerr = columnNumberToLetter(colInicio + 8);
    sheet.getRange(fila, colInicio + 9).setFormula('=IF(' + letraCont + fila + '=0;0%;' + letraCerr + fila + '/' + letraCont + fila + ')');
  });
  
  const filaTotales = filaInicio + clinicas.length + 1;
  sheet.getRange(filaTotales, colInicio).setValue('Total General');
  
  [1, 2, 4, 6, 8].forEach(colOffset => {
    const col = colInicio + colOffset;
    const letraCol = columnNumberToLetter(col);
    sheet.getRange(filaTotales, col).setFormula('=SUM(' + letraCol + (filaInicio + 1) + ':' + letraCol + (filaInicio + clinicas.length) + ')');
  });
  
  const letraGestT = columnNumberToLetter(colInicio + 1);
  const letraMetaT = columnNumberToLetter(colInicio + 2);
  const letraContT = columnNumberToLetter(colInicio + 4);
  const letraIntT = columnNumberToLetter(colInicio + 6);
  const letraCerrT = columnNumberToLetter(colInicio + 8);
  
  sheet.getRange(filaTotales, colInicio + 3).setFormula('=IF(' + letraMetaT + filaTotales + '=0;0%;' + letraGestT + filaTotales + '/' + letraMetaT + filaTotales + ')');
  sheet.getRange(filaTotales, colInicio + 5).setFormula('=IF(' + letraMetaT + filaTotales + '=0;0%;' + letraContT + filaTotales + '/' + letraMetaT + filaTotales + ')');
  sheet.getRange(filaTotales, colInicio + 7).setFormula('=IF(' + letraContT + filaTotales + '=0;0%;' + letraIntT + filaTotales + '/' + letraContT + filaTotales + ')');
  sheet.getRange(filaTotales, colInicio + 9).setFormula('=IF(' + letraContT + filaTotales + '=0;0%;' + letraCerrT + filaTotales + '/' + letraContT + filaTotales + ')');
  
  const rangoTotales = sheet.getRange(filaTotales, colInicio, 1, 10);
  rangoTotales.setBackground('#4472C4');
  rangoTotales.setFontColor('white');
  rangoTotales.setFontWeight('bold');
  rangoTotales.setHorizontalAlignment('center');
  
  sheet.getRange(filaInicio + 1, colInicio + 3, clinicas.length + 1, 1).setNumberFormat('0%');
  sheet.getRange(filaInicio + 1, colInicio + 5, clinicas.length + 1, 1).setNumberFormat('0%');
  sheet.getRange(filaInicio + 1, colInicio + 7, clinicas.length + 1, 1).setNumberFormat('0%');
  sheet.getRange(filaInicio + 1, colInicio + 9, clinicas.length + 1, 1).setNumberFormat('0%');
  
  sheet.getRange(filaInicio + 1, colInicio + 1, clinicas.length, 9).setHorizontalAlignment('center');
  sheet.getRange(filaInicio, colInicio, clinicas.length + 2, 10).setBorder(true, true, true, true, true, true);
}

function columnNumberToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}