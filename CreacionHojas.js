/**
 * MÓDULO 4: CREACIÓN DE HOJAS DE EJECUTIVO
 * Funciones para crear y configurar hojas individuales
 */

/**
 * Crea una hoja de ejecutivo con datos y configuración
 */
function crearHojaEjecutivo(ss, nombreEjecutivo, datos, encabezadosOriginales) {
  try {
    var existe = ss.getSheetByName(nombreEjecutivo);
    if (existe) ss.deleteSheet(existe);
    
    var hoja = ss.insertSheet(nombreEjecutivo);
    var expandidos = encabezadosOriginales.concat(COLUMNAS_NUEVAS);
    
    // Escribir encabezados
    hoja.getRange(1, 1, 1, expandidos.length).setValues([expandidos]);
    hoja.getRange(1, 1, 1, expandidos.length)
      .setBackground(COLORES.HEADER_EJECUTIVO)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Escribir datos
    if (datos.length > 0) {
      var datosExp = [];
      for (var i = 0; i < datos.length; i++) {
        datosExp.push(datos[i].concat(['', '', '', 'Sin Gestión', 'Sin Gestión', '', '', '']));
      }
      hoja.getRange(2, 1, datosExp.length, expandidos.length).setValues(datosExp);
    }
    
    // Aplicar configuraciones
    aplicarValidacionesYFormulas(hoja, expandidos, datos.length);
    protegerColumnasOriginales(hoja, encabezadosOriginales.length);
    
    // Crear filtro
    if (datos.length > 0) {
      hoja.getRange(1, 1, datos.length + 1, expandidos.length).createFilter();
    }
    
    // Auto-ajustar columnas
    hoja.autoResizeColumns(1, expandidos.length);
    
  } catch (error) {
    Logger.log('Error creando hoja ' + nombreEjecutivo + ': ' + error);
    throw error;
  }
}

/**
 * Aplica validaciones de datos y fórmulas a una hoja
 */
function aplicarValidacionesYFormulas(hoja, encabezados, numeroFilas) {
  if (numeroFilas === 0) return;
  
  try {
    var idx = {
      fechaLlamada: encabezados.indexOf('FECHA_LLAMADA'),
      fechaCompromiso: encabezados.indexOf('FECHA_COMPROMISO'),
      estado: encabezados.indexOf('ESTADO'),
      subEstado: encabezados.indexOf('SUB_ESTADO'),
      estadoCompromiso: encabezados.indexOf('ESTADO_COMPROMISO')
    };
    
    // Validación de fecha de llamada
    if (idx.fechaLlamada !== -1) {
      hoja.getRange(2, idx.fechaLlamada + 1, numeroFilas, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation()
          .requireDate()
          .setAllowInvalid(false)
          .build());
    }
    
    // Validación de fecha de compromiso
    if (idx.fechaCompromiso !== -1) {
      hoja.getRange(2, idx.fechaCompromiso + 1, numeroFilas, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation()
          .requireDate()
          .setAllowInvalid(false)
          .build());
    }
    
    // Validación de estado
    if (idx.estado !== -1) {
      hoja.getRange(2, idx.estado + 1, numeroFilas, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInList(ESTADOS_GESTION)
          .setAllowInvalid(false)
          .build());
    }
    
    // Validación de sub-estado
    if (idx.subEstado !== -1) {
      hoja.getRange(2, idx.subEstado + 1, numeroFilas, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInList(SUB_ESTADOS_GESTION)
          .setAllowInvalid(false)
          .build());
    }
    
    // Fórmula de estado de compromiso
    if (idx.estadoCompromiso !== -1 && idx.fechaCompromiso !== -1) {
      var col = columnNumberToLetter(idx.fechaCompromiso + 1);
      var formulas = [];
      for (var i = 2; i <= numeroFilas + 1; i++) {
        formulas.push([
          '=IF(ISBLANK(' + col + i + '),"SIN_COMPROMISO",IF(' + col + i + '=TODAY(),"LLAMAR_HOY",IF(' + col + i + '<TODAY(),"COMPROMISO_VENCIDO","COMPROMISO_FUTURO")))'
        ]);
      }
      hoja.getRange(2, idx.estadoCompromiso + 1, numeroFilas, 1).setFormulas(formulas);
    }
  } catch (e) {
    Logger.log('Error en validaciones: ' + e.toString());
  }
}

/**
 * Trigger automático para actualizar estado de compromiso al editar fecha
 */
function onEdit(e) {
  try {
    var hoja = e.source.getActiveSheet();
    if (/^BBDD_.*_REMOTO/i.test(hoja.getName())) return;
    
    var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var idxC = enc.indexOf('FECHA_COMPROMISO');
    var idxE = enc.indexOf('ESTADO_COMPROMISO');
    
    if (idxC !== -1 && idxE !== -1 && e.range.getColumn() === idxC + 1) {
      var fila = e.range.getRow();
      var col = columnNumberToLetter(idxC + 1);
      var formula = '=IF(ISBLANK(' + col + fila + '),"SIN_COMPROMISO",IF(' + col + fila + '=TODAY(),"LLAMAR_HOY",IF(' + col + fila + '<TODAY(),"COMPROMISO_VENCIDO","COMPROMISO_FUTURO")))';
      hoja.getRange(fila, idxE + 1).setFormula(formula);
    }
  } catch (error) {
    Logger.log('Error en onEdit: ' + error);
  }
}