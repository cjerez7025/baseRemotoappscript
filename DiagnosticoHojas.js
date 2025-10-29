/**
 * SCRIPT DE DIAGNÓSTICO
 * Para identificar por qué las hojas nuevas no se incluyen en BBDD_REPORTE
 */

function diagnosticarHojasEjecutivos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    
    Logger.log('============================================');
    Logger.log('    DIAGNÓSTICO DE HOJAS DE EJECUTIVOS');
    Logger.log('============================================');
    Logger.log('Total de hojas: ' + hojas.length);
    Logger.log('');
    
    var hojasEjecutivos = [];
    var hojasExcluidas = [];
    var hojasSinDatos = [];
    var hojasOtras = [];
    
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      var numFilas = hojas[i].getLastRow();
      var numCols = hojas[i].getLastColumn();
      
      Logger.log('-------------------------------------------');
      Logger.log('HOJA #' + (i + 1) + ': "' + nombre + '"');
      Logger.log('Filas: ' + numFilas + ' | Columnas: ' + numCols);
      
      // 1. Verificar si es BBDD_*_REMOTO
      if (/^BBDD_.*_REMOTO/i.test(nombre)) {
        Logger.log('TIPO: Base de datos origen (BBDD_*_REMOTO)');
        Logger.log('ESTADO: EXCLUIDA');
        hojasExcluidas.push({
          nombre: nombre,
          razon: 'BBDD_*_REMOTO'
        });
        Logger.log('');
        continue;
      }
      
      // 2. Verificar si está en HOJAS_EXCLUIDAS
      var esExcluida = false;
      var razonExclusion = '';
      
      for (var j = 0; j < HOJAS_EXCLUIDAS.length; j++) {
        if (nombre.toUpperCase().indexOf(HOJAS_EXCLUIDAS[j].toUpperCase()) !== -1) {
          esExcluida = true;
          razonExclusion = 'Lista HOJAS_EXCLUIDAS: ' + HOJAS_EXCLUIDAS[j];
          break;
        }
      }
      
      if (esExcluida) {
        Logger.log('TIPO: Hoja del sistema');
        Logger.log('ESTADO: EXCLUIDA (' + razonExclusion + ')');
        hojasExcluidas.push({
          nombre: nombre,
          razon: razonExclusion
        });
        Logger.log('');
        continue;
      }
      
      // 3. Verificar que tenga datos
      if (numFilas <= 1) {
        Logger.log('TIPO: Sin datos');
        Logger.log('ESTADO: EXCLUIDA (solo ' + numFilas + ' fila)');
        hojasSinDatos.push(nombre);
        Logger.log('');
        continue;
      }
      
      // 4. Analizar encabezados
      try {
        var numColsALeer = Math.min(numCols, 30);
        var encabezados = hojas[i].getRange(1, 1, 1, numColsALeer).getValues()[0];
        
        var encStr = [];
        for (var k = 0; k < encabezados.length; k++) {
          if (encabezados[k]) {
            encStr.push(encabezados[k].toString().trim());
          }
        }
        
        Logger.log('ENCABEZADOS (' + encStr.length + '):');
        Logger.log('  ' + encStr.join(', '));
        Logger.log('');
        
        // Buscar columnas clave
        var columnasClave = [
          'FECHA_LLAMADA',
          'ESTADO_COMPROMISO',
          'SUB_ESTADO',
          'NOTA_EJECUTIVO',
          'ESTADO',
          'EJECUTIVO',
          'FECHA_COMPROMISO'
        ];
        
        var encontradas = [];
        for (var m = 0; m < columnasClave.length; m++) {
          for (var n = 0; n < encStr.length; n++) {
            if (encStr[n].toUpperCase() === columnasClave[m].toUpperCase()) {
              encontradas.push(columnasClave[m]);
              break;
            }
          }
        }
        
        Logger.log('COLUMNAS CLAVE ENCONTRADAS (' + encontradas.length + '):');
        if (encontradas.length > 0) {
          Logger.log('  ✓ ' + encontradas.join(', '));
        } else {
          Logger.log('  ✗ Ninguna');
        }
        Logger.log('');
        
        // Decisión final
        if (encontradas.length > 0) {
          Logger.log('TIPO: Hoja de ejecutivo');
          Logger.log('ESTADO: ✅ INCLUIDA EN BBDD_REPORTE');
          hojasEjecutivos.push({
            nombre: nombre,
            filas: numFilas,
            columnas: numCols,
            columnasEncontradas: encontradas
          });
        } else {
          Logger.log('TIPO: Hoja desconocida (sin columnas clave)');
          Logger.log('ESTADO: ❌ EXCLUIDA');
          hojasOtras.push({
            nombre: nombre,
            filas: numFilas,
            columnas: numCols
          });
        }
        
      } catch (error) {
        Logger.log('ERROR al analizar encabezados: ' + error.toString());
        hojasOtras.push({
          nombre: nombre,
          error: error.toString()
        });
      }
      
      Logger.log('');
    }
    
    // RESUMEN FINAL
    Logger.log('============================================');
    Logger.log('              RESUMEN FINAL');
    Logger.log('============================================');
    Logger.log('');
    Logger.log('📊 HOJAS DE EJECUTIVOS: ' + hojasEjecutivos.length);
    for (var p = 0; p < hojasEjecutivos.length; p++) {
      Logger.log('  ' + (p + 1) + '. ' + hojasEjecutivos[p].nombre + 
                 ' (' + (hojasEjecutivos[p].filas - 1) + ' registros)');
    }
    Logger.log('');
    
    Logger.log('🚫 HOJAS EXCLUIDAS: ' + hojasExcluidas.length);
    for (var q = 0; q < Math.min(hojasExcluidas.length, 10); q++) {
      Logger.log('  • ' + hojasExcluidas[q].nombre + ' - ' + hojasExcluidas[q].razon);
    }
    if (hojasExcluidas.length > 10) {
      Logger.log('  ... y ' + (hojasExcluidas.length - 10) + ' más');
    }
    Logger.log('');
    
    Logger.log('📭 HOJAS SIN DATOS: ' + hojasSinDatos.length);
    for (var r = 0; r < Math.min(hojasSinDatos.length, 5); r++) {
      Logger.log('  • ' + hojasSinDatos[r]);
    }
    if (hojasSinDatos.length > 5) {
      Logger.log('  ... y ' + (hojasSinDatos.length - 5) + ' más');
    }
    Logger.log('');
    
    Logger.log('❓ HOJAS OTRAS: ' + hojasOtras.length);
    for (var s = 0; s < Math.min(hojasOtras.length, 5); s++) {
      Logger.log('  • ' + hojasOtras[s].nombre);
    }
    if (hojasOtras.length > 5) {
      Logger.log('  ... y ' + (hojasOtras.length - 5) + ' más');
    }
    Logger.log('');
    
    Logger.log('============================================');
    Logger.log('TOTAL: ' + hojas.length + ' hojas analizadas');
    Logger.log('============================================');
    
    // Mostrar mensaje al usuario
    var ui = SpreadsheetApp.getUi();
    var mensaje = '📊 DIAGNÓSTICO COMPLETADO\n\n';
    mensaje += 'Hojas de ejecutivos detectadas: ' + hojasEjecutivos.length + '\n';
    mensaje += 'Hojas excluidas: ' + hojasExcluidas.length + '\n';
    mensaje += 'Hojas sin datos: ' + hojasSinDatos.length + '\n';
    mensaje += 'Hojas otras: ' + hojasOtras.length + '\n\n';
    mensaje += 'Revisa el log (Ver > Registros de ejecución) para detalles.';
    
    ui.alert('Diagnóstico Completado', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('ERROR CRÍTICO: ' + error.toString());
    Logger.log(error.stack);
  }
}