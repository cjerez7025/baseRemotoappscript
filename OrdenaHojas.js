/**
 * Ordena las hojas del spreadsheet en grupos lógicos:
 * 1. Hoja BBDD_*_REMOTO* (base de datos original)
 * 2. Hojas de Gestión (RESUMEN, LLAMADAS, PRODUCTIVIDAD, etc.)
 * 3. Hoja BBDD_REPORTE
 * 4. Hojas de Ejecutivos (ordenadas alfabéticamente)
 */
function ordenarHojasPorGrupo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojas = spreadsheet.getSheets();
    
    // Definir grupos de hojas
    const hojasGestion = [];
    const hojaReporte = [];
    const hojasEjecutivos = [];
    const hojaRemoto = [];
    const hojasOtras = [];
    
    // Orden específico para hojas de gestión
    const ordenGestion = [
      'RESUMEN',
      'LLAMADAS',
      'PRODUCTIVIDAD',
      'DASHBOARD',
      'TOTALES',
      'GRAFICOS',
      'CONFIGURACION',
      'Datos desplegables'
    ];
    
    // Clasificar todas las hojas
    hojas.forEach(hoja => {
      const nombre = hoja.getName();
      
      // Detectar hoja BBDD_*_REMOTO*
      if (/^BBDD_.*_REMOTO/i.test(nombre)) {
        hojaRemoto.push(hoja);
      }
      // Detectar BBDD_REPORTE
      else if (nombre === 'BBDD_REPORTE') {
        hojaReporte.push(hoja);
      }
      // Detectar hojas de gestión
      else if (ordenGestion.includes(nombre)) {
        hojasGestion.push(hoja);
      }
      // Detectar hojas de ejecutivos (tienen columnas específicas)
      else if (esHojaEjecutivo(hoja)) {
        hojasEjecutivos.push(hoja);
      }
      // Otras hojas (Sheet1, Hoja1, etc.)
      else {
        hojasOtras.push(hoja);
      }
    });
    
    // Ordenar hojas de gestión según el orden definido
    hojasGestion.sort((a, b) => {
      const indexA = ordenGestion.indexOf(a.getName());
      const indexB = ordenGestion.indexOf(b.getName());
      return indexA - indexB;
    });
    
    // Ordenar hojas de ejecutivos alfabéticamente
    hojasEjecutivos.sort((a, b) => {
      return a.getName().localeCompare(b.getName());
    });
    
    // Construir el orden final - BASE DE DATOS PRIMERO
    const ordenFinal = [
      ...hojaRemoto,
      ...hojasGestion,
      ...hojaReporte,
      ...hojasEjecutivos,
      ...hojasOtras
    ];
    
    // Aplicar el nuevo orden
    let posicion = 0;
    ordenFinal.forEach(hoja => {
      spreadsheet.setActiveSheet(hoja);
      spreadsheet.moveActiveSheet(posicion + 1); // Las posiciones empiezan en 1
      posicion++;
    });
    
    // Activar la primera hoja
    if (ordenFinal.length > 0) {
      spreadsheet.setActiveSheet(ordenFinal[0]);
    }
    
    // Logs para consola
    console.log('✓ Hojas ordenadas correctamente');
    console.log('🗄️ Base de Datos Original:', hojaRemoto.length);
    console.log('📊 Hojas de Gestión:', hojasGestion.length);
    console.log('📋 BBDD_REPORTE:', hojaReporte.length);
    console.log('👥 Hojas de Ejecutivos:', hojasEjecutivos.length);
    console.log('📄 Otras hojas:', hojasOtras.length);
    console.log('Orden: Base Original → Gestión → Reporte → Ejecutivos (A-Z)');
    
    return true;
    
  } catch (error) {
    console.error('Error ordenando hojas:', error);
    throw error;
  }
}

/**
 * Determina si una hoja es de un ejecutivo
 * Verifica que tenga las columnas características de hojas de ejecutivos
 */
function esHojaEjecutivo(hoja) {
  try {
    // Verificar que tenga al menos 2 filas
    if (hoja.getLastRow() < 2) {
      return false;
    }
    
    // Obtener encabezados
    const encabezados = hoja.getRange(1, 1, 1, Math.min(hoja.getLastColumn(), 20)).getValues()[0];
    
    // Columnas que identifican una hoja de ejecutivo
    const columnasEjecutivo = ['FECHA_LLAMADA', 'ESTADO', 'SUB_ESTADO', 'NOTA_EJECUTIVO'];
    
    // Verificar que tenga al menos 2 de estas columnas
    const columnasEncontradas = columnasEjecutivo.filter(col => encabezados.includes(col));
    
    return columnasEncontradas.length >= 2;
    
  } catch (error) {
    return false;
  }
}

