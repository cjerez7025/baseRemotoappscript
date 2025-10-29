/**
 * MÓDULO: ELIMINACIÓN DE FILAS EN BLANCO
 * Elimina filas completamente vacías de las hojas de ejecutivos después de distribución
 * para evitar que se consoliden en BBDD_REPORTE
 */

/**
 * FUNCIÓN PRINCIPAL: Eliminar filas en blanco de todas las hojas después de distribuir
 * Se debe ejecutar después de la distribución de datos
 */
function eliminarFilasBlancasPostDistribucion() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojas = spreadsheet.getSheets();
    let hojasLimpiadas = 0;
    let filasEliminadas = 0;
    
    Logger.log('=== INICIANDO LIMPIEZA DE FILAS EN BLANCO ===');
    
    for (let i = 0; i < hojas.length; i++) {
      const hoja = hojas[i];
      const nombreHoja = hoja.getName();
      
      // Excluir hojas especiales
      if (nombreHoja === 'BBDD_REPORTE' || 
          nombreHoja === 'RESUMEN' || 
          nombreHoja === 'LLAMADAS' || 
          nombreHoja === 'PRODUCTIVIDAD' ||
          /REMOTO/i.test(nombreHoja)) {
        continue;
      }
      
      // Solo procesar hojas de ejecutivos que tengan datos
      if (hoja.getLastRow() > 1) {
        const eliminadas = eliminarFilasBlancasDeHojaOptimizado(hoja);
        if (eliminadas > 0) {
          hojasLimpiadas++;
          filasEliminadas += eliminadas;
          Logger.log(`✓ ${nombreHoja}: ${eliminadas} filas eliminadas`);
        }
      }
    }
    
    Logger.log(`\n=== LIMPIEZA COMPLETADA ===`);
    Logger.log(`   - Hojas procesadas: ${hojasLimpiadas}`);
    Logger.log(`   - Total filas eliminadas: ${filasEliminadas}`);
    
    return {
      exito: true,
      hojasLimpiadas: hojasLimpiadas,
      filasEliminadas: filasEliminadas
    };
    
  } catch (error) {
    Logger.log('ERROR en limpieza de filas: ' + error.message);
    return {
      exito: false,
      error: error.message
    };
  }
}

/**
 * Elimina filas completamente vacías de una hoja específica
 * @param {Sheet} hoja - La hoja de cálculo a limpiar
 * @returns {number} - Número de filas eliminadas
 */
function eliminarFilasBlancasDeHoja(hoja) {
  try {
    const ultimaFila = hoja.getLastRow();
    const ultimaColumna = hoja.getLastColumn();
    
    if (ultimaFila <= 1 || ultimaColumna === 0) {
      return 0; // No hay datos para procesar
    }
    
    // Obtener todos los datos (excepto encabezado)
    const datos = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
    const filasAEliminar = [];
    
    // Identificar filas completamente vacías
    for (let i = 0; i < datos.length; i++) {
      const fila = datos[i];
      const estaVacia = fila.every(celda => {
        return celda === '' || celda === null || celda === undefined;
      });
      
      if (estaVacia) {
        filasAEliminar.push(i + 2); // +2 porque: +1 por índice base 0, +1 por encabezado
      }
    }
    
    // Eliminar filas de abajo hacia arriba para no afectar los índices
    if (filasAEliminar.length > 0) {
      for (let i = filasAEliminar.length - 1; i >= 0; i--) {
        hoja.deleteRow(filasAEliminar[i]);
      }
    }
    
    return filasAEliminar.length;
    
  } catch (error) {
    Logger.log(`Error limpiando hoja ${hoja.getName()}: ${error.message}`);
    return 0;
  }
}

/**
 * Versión optimizada: Elimina bloques de filas vacías consecutivas
 * Más eficiente para hojas con muchas filas vacías
 */
function eliminarFilasBlancasDeHojaOptimizado(hoja) {
  try {
    const ultimaFila = hoja.getLastRow();
    const ultimaColumna = hoja.getLastColumn();
    
    if (ultimaFila <= 1 || ultimaColumna === 0) {
      return 0;
    }
    
    const datos = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
    const bloquesAEliminar = []; // [{inicio, cantidad}]
    let bloqueActual = null;
    let totalFilasEliminadas = 0;
    
    // Identificar bloques de filas vacías
    for (let i = 0; i < datos.length; i++) {
      const fila = datos[i];
      const estaVacia = fila.every(celda => celda === '' || celda === null || celda === undefined);
      
      if (estaVacia) {
        if (bloqueActual === null) {
          // Iniciar nuevo bloque
          bloqueActual = {
            inicio: i + 2, // +2 por índice y encabezado
            cantidad: 1
          };
        } else {
          // Extender bloque actual
          bloqueActual.cantidad++;
        }
      } else {
        // Fila con datos, cerrar bloque si existe
        if (bloqueActual !== null) {
          bloquesAEliminar.push(bloqueActual);
          bloqueActual = null;
        }
      }
    }
    
    // Agregar último bloque si existe
    if (bloqueActual !== null) {
      bloquesAEliminar.push(bloqueActual);
    }
    
    // Eliminar bloques de abajo hacia arriba
    for (let i = bloquesAEliminar.length - 1; i >= 0; i--) {
      const bloque = bloquesAEliminar[i];
      hoja.deleteRows(bloque.inicio, bloque.cantidad);
      totalFilasEliminadas += bloque.cantidad;
    }
    
    return totalFilasEliminadas;
    
  } catch (error) {
    Logger.log(`Error limpiando hoja ${hoja.getName()}: ${error.message}`);
    return 0;
  }
}

/**
 * INTEGRACIÓN: Función modificada de distribución que incluye limpieza automática
 * Esta debe reemplazar o complementar tu función de distribución existente
 */
function distribuirYLimpiar() {
  try {
    // 1. Ejecutar distribución normal (tu función existente)
    procesarEjecutivos(); // Esta función ahora incluye la limpieza automáticamente
    
    // 2. Esperar a que termine la distribución
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    
    // 3. Limpiar filas en blanco
    const resultado = eliminarFilasBlancasPostDistribucion();
    
    // 4. Regenerar BBDD_REPORTE para reflejar los cambios
    if (resultado.exito && resultado.filasEliminadas > 0) {
      Logger.log('Regenerando BBDD_REPORTE sin filas vacías...');
      crearOActualizarReporteAutomatico(SpreadsheetApp.getActiveSpreadsheet());
    }
    
    return resultado;
    
  } catch (error) {
    Logger.log('Error en distribuirYLimpiar: ' + error.message);
    throw error;
  }
}



/**
 * FUNCIÓN DE DIAGNÓSTICO: Identifica filas vacías sin eliminarlas
 * Útil para verificar antes de ejecutar la limpieza real
 */
function identificarFilasVaciasEnTodasLasHojas() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = spreadsheet.getSheets();
  const reporte = [];
  
  Logger.log('\n=== DIAGNÓSTICO DE FILAS VACÍAS ===\n');
  
  for (let i = 0; i < hojas.length; i++) {
    const hoja = hojas[i];
    const nombreHoja = hoja.getName();
    
    // Excluir hojas especiales
    if (nombreHoja === 'BBDD_REPORTE' || 
        nombreHoja === 'RESUMEN' || 
        nombreHoja === 'LLAMADAS' || 
        nombreHoja === 'PRODUCTIVIDAD' ||
        /REMOTO/i.test(nombreHoja)) {
      continue;
    }
    
    if (hoja.getLastRow() > 1) {
      const ultimaFila = hoja.getLastRow();
      const ultimaColumna = hoja.getLastColumn();
      const datos = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
      
      let filasVacias = 0;
      datos.forEach(fila => {
        if (fila.every(celda => celda === '' || celda === null)) {
          filasVacias++;
        }
      });
      
      if (filasVacias > 0) {
        reporte.push({
          hoja: nombreHoja,
          totalFilas: ultimaFila - 1,
          filasVacias: filasVacias,
          porcentaje: ((filasVacias / (ultimaFila - 1)) * 100).toFixed(1) + '%'
        });
      }
    }
  }
  
  Logger.log('📊 REPORTE DE FILAS VACÍAS:\n');
  
  if (reporte.length === 0) {
    Logger.log('✅ No se encontraron filas vacías en ninguna hoja');
  } else {
    let totalVacias = 0;
    reporte.forEach(item => {
      Logger.log(`\n${item.hoja}:`);
      Logger.log(`  - Total filas: ${item.totalFilas}`);
      Logger.log(`  - Filas vacías: ${item.filasVacias} (${item.porcentaje})`);
      totalVacias += item.filasVacias;
    });
    Logger.log(`\n📈 TOTAL FILAS VACÍAS: ${totalVacias}`);
    Logger.log(`📄 HOJAS AFECTADAS: ${reporte.length}`);
  }
  
  // Mostrar resultado en UI
  const ui = SpreadsheetApp.getUi();
  if (reporte.length === 0) {
    ui.alert('✅ Diagnóstico Completo', 
      'No se encontraron filas vacías en las hojas de ejecutivos.', 
      ui.ButtonSet.OK);
  } else {
    let mensaje = '📊 Se encontraron filas vacías:\n\n';
    reporte.forEach(item => {
      mensaje += `${item.hoja}: ${item.filasVacias} filas (${item.porcentaje})\n`;
    });
    mensaje += `\nTotal: ${reporte.reduce((sum, item) => sum + item.filasVacias, 0)} filas vacías`;
    mensaje += `\n\n¿Deseas eliminarlas?`;
    
    const respuesta = ui.alert('🔍 Diagnóstico Completo', mensaje, ui.ButtonSet.YES_NO);
    
    if (respuesta === ui.Button.YES) {
      const resultado = eliminarFilasBlancasPostDistribucion();
      if (resultado.exito) {
        ui.alert('✅ Limpieza Completada', 
          `Se eliminaron ${resultado.filasEliminadas} filas vacías de ${resultado.hojasLimpiadas} hojas.`, 
          ui.ButtonSet.OK);
      }
    }
  }
  
  return reporte;
}