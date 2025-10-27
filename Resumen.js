/**
 * ARCHIVO: Resumen.gs
 * FUNCIONALIDADES DE RESUMEN Y DISTRIBUCIÓN
 */

/**
 * Verifica si una fila está vacía (todos los valores están vacíos)
 */
function esFilaVacia(row) {
  return row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * Genera automáticamente la hoja RESUMEN con distribución
 */
function generarResumenAutomatico(spreadsheet) {
  try {
    const bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      console.log('No se puede generar resumen: no existe BBDD_REPORTE');
      return;
    }
    
    let summarySheet = spreadsheet.getSheetByName('RESUMEN');
    
    // Crear la hoja RESUMEN si no existe
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet('RESUMEN');
    } else {
      summarySheet.clear();
    }
    
    // Lista oficial de sub-estados con su orden
    const subEstadosOrdenados = [
      'Ventas',
      'Volver a llamar',
      'En gestión',
      'Sin motivo',
      'Problema económico',
      'No cumple requisitos',
      'Mala Experiencia',
      'Cuenta con seguro',
      'Prefiere competencia',
      'No contesta',
      'Teléfono erróneo',
      'Apagado',
      'Ya habia sido llamado',
      'Sin Gestión'
    ];

    // Definir las combinaciones válidas de ESTADO y SUB_ESTADO
    const combinacionesValidas = {
      'Cerrado': ['Sin Gestión'],
      'En Gestión': ['Volver a llamar'],
      'No Contactado': ['Apagado', 'No contesta', 'Teléfono erróneo'],
      'Interesado': ['Volver a llamar'],
      'Sin Gestión': ['Sin Gestión'],
      'Sin Interés': ['Cuenta con seguro', 'No cumple requisitos', 'Prefiere competencia', 
                      'Problema económico', 'Sin motivo', 'Ya habia sido llamado']
    };

    // Obtener datos
    const data = bddSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices
    const subEstadoIndex = headers.indexOf('SUB_ESTADO');
    const fechaLlamadaIndex = headers.indexOf('FECHA_LLAMADA');
    const estadoIndex = headers.indexOf('ESTADO');
    const ejecutivoIndex = headers.indexOf('EJECUTIVO');
    
    if (subEstadoIndex === -1 || fechaLlamadaIndex === -1 || estadoIndex === -1) {
      console.log('No se encontraron las columnas necesarias para el resumen');
      return;
    }
    
    // Filtrar filas vacías y contar registros válidos
    const datosValidos = data.slice(1).filter(row => !esFilaVacia(row));
    const totalRegistros = datosValidos.length;
    
    // Inicializar conteos
    let conteos = {
      registrosCargados: totalRegistros,
      subEstados: {},
      enGestion: 0,
      ventas: 0
    };
    
    // Procesar datos válidos para el resumen
    for (let i = 0; i < datosValidos.length; i++) {
      const row = datosValidos[i];
      const subEstado = row[subEstadoIndex];
      const estado = row[estadoIndex];
      
      // Contar ventas (estado Cerrado)
      if (estado === 'Cerrado') {
        conteos.ventas++;
      }
      
      if (subEstado && subEstado.toString().trim() !== '') {
        // Contabilizar para "En gestión"
        if (estado === 'En Gestión' && subEstado === 'Volver a llamar') {
          conteos.enGestion++;
        } else {
          // Contabilizar otros sub-estados
          conteos.subEstados[subEstado] = (conteos.subEstados[subEstado] || 0) + 1;
        }
      }
    }
    
    // Configurar formato inicial
    summarySheet.getRange('A1').setValue('ENERO');
    summarySheet.getRange('A2').setValue('Registros CARGADOS');
    summarySheet.getRange('A3').setValue('MES DE GESTION');
    
    // Escribir registros cargados
    summarySheet.getRange(2, 2).setValue(totalRegistros);
    
    // Escribir datos en el orden especificado
    let row = 4;
    subEstadosOrdenados.forEach(subEstado => {
      summarySheet.getRange(row, 1).setValue(subEstado);
      
      // Asignar el valor correspondiente según el tipo de fila
      let valor;
      if (subEstado === 'Ventas') {
        valor = conteos.ventas;
        // Aplicar fondo verde a la fila de ventas
        summarySheet.getRange(row, 1, 1, 2).setBackground('#90EE90');
      } else if (subEstado === 'En gestión') {
        valor = conteos.enGestion;
      } else {
        valor = conteos.subEstados[subEstado] || 0;
      }
      
      summarySheet.getRange(row, 2).setValue(valor);
      row++;
    });
    
    // Obtener dimensiones del resumen
    const lastRow = summarySheet.getLastRow();
    const lastCol = summarySheet.getLastColumn();
    
    // Aplicar formato al resumen
    const range = summarySheet.getRange(1, 1, lastRow, lastCol);
    range.setBorder(true, true, true, true, true, true);
    
    // Formato de encabezados del resumen
    const headerRange = summarySheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground('#00B0F0');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Auto-ajustar columnas del resumen
    summarySheet.autoResizeColumns(1, lastCol);

    // Procesar datos válidos para la tabla de distribución con validación
    let distribucion = {};
    let errores = {};
    
    for (let i = 0; i < datosValidos.length; i++) {
      const estado = datosValidos[i][estadoIndex];
      const subEstado = datosValidos[i][subEstadoIndex];
      const ejecutivo = datosValidos[i][ejecutivoIndex] || 'Sin Ejecutivo';
      
      // Validar que estado y subEstado no estén vacíos
      if (estado && estado.toString().trim() !== '' && 
          subEstado && subEstado.toString().trim() !== '') {
        const key = `${estado}|${subEstado}`;
        distribucion[key] = (distribucion[key] || 0) + 1;
        
        // Verificar si la combinación es válida
        const esValida = combinacionesValidas[estado] && 
                        combinacionesValidas[estado].includes(subEstado);
        
        if (!esValida) {
          if (!errores[key]) {
            errores[key] = {
              count: 0,
              ejecutivos: new Set()
            };
          }
          errores[key].count++;
          errores[key].ejecutivos.add(ejecutivo);
        }
      }
    }

    // Convertir la distribución a un array ordenado con validación
    let distribucionArray = Object.entries(distribucion).map(([key, count]) => {
      const [estado, subEstado] = key.split('|');
      const error = errores[key];
      return [
        estado, 
        subEstado, 
        count,
        error ? 'INCORRECTO' : 'CORRECTO',
        error ? Array.from(error.ejecutivos).join(', ') : ''
      ];
    }).sort((a, b) => {
      if (a[0] < b[0]) return -1;
      if (a[0] > b[0]) return 1;
      return a[1] < b[1] ? -1 : 1;
    });

    // Agregar título de la distribución
    const distStartRow = lastRow + 4;
    summarySheet.getRange(distStartRow - 1, 1).setValue('DISTRIBUCIÓN POR ESTADO Y SUB-ESTADO');
    summarySheet.getRange(distStartRow - 1, 1).setFontWeight('bold');

    // Escribir encabezados de la distribución
    summarySheet.getRange(distStartRow, 1, 1, 5).setValues([
      ['ESTADO', 'SUB_ESTADO', 'CUENTA de rut_cliente', 'VALIDACIÓN', 'EJECUTIVOS CON ERROR']
    ]);
    summarySheet.getRange(distStartRow, 1, 1, 5).setFontWeight('bold');

    // Escribir datos de la distribución
    if (distribucionArray.length > 0) {
      const distRange = summarySheet.getRange(distStartRow + 1, 1, distribucionArray.length, 5);
      distRange.setValues(distribucionArray);
      
      // Aplicar formato condicional
      distribucionArray.forEach((row, index) => {
        if (row[3] === 'INCORRECTO') {
          summarySheet.getRange(distStartRow + 1 + index, 1, 1, 5)
            .setBackground('#FFB6C1'); // Color rosa claro para errores
        }
      });
    }

    // Dar formato a la tabla de distribución
    const distRange = summarySheet.getRange(distStartRow, 1, distribucionArray.length + 1, 5);
    distRange.applyRowBanding();
    distRange.setBorder(true, true, true, true, true, true);

    // Auto-ajustar columnas para toda la hoja
    summarySheet.autoResizeColumns(1, 5);

    console.log('Hoja RESUMEN generada exitosamente');
    console.log(`Total de registros válidos procesados: ${totalRegistros}`);
    
  } catch (error) {
    console.error('Error generando resumen automático:', error);
  }
}

/**
 * Muestra una notificación tipo toast en la esquina inferior
 */
function mostrarNotificacion(mensaje) {
  console.log(mensaje);
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { margin: 0; padding: 0; font-family: Arial, sans-serif; }
      .toast {
        background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
        color: white;
        padding: 12px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        font-size: 13px;
        position: fixed;
        bottom: 20px;
        left: 20px;
        z-index: 10000;
        animation: slideIn 0.4s ease-out;
      }
      @keyframes slideIn {
        from {
          transform: translateX(-400px);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
    </style>
    <div class="toast">✓ ${mensaje}</div>
    <script>
      setTimeout(() => google.script.host.close(), 2500);
    </script>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '');
}

/**
 * Función individual para generar resumen (compatible con el código original)
 */
function generateSummary() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      SpreadsheetApp.getUi().alert('Error: No se encontró la hoja BBDD_REPORTE. Debe procesarse primero los ejecutivos.');
      return;
    }
    
    generarResumenAutomatico(spreadsheet);
    mostrarNotificacion('Resumen y distribución generados exitosamente');
    
  } catch (error) {
    console.error('Error en generateSummary:', error);
  }
}

/**
 * Función legacy para inicializar (mantiene compatibilidad)
 * Se ejecuta automáticamente al cargar la planilla
 */
function initializeSheet() {
  try {
    onOpen();
    generarResumenAutomatico(SpreadsheetApp.getActiveSpreadsheet());
    mostrarNotificacion('Resumen generado al cargar');
  } catch (error) {
    console.error('Error en initializeSheet:', error);
  }
}