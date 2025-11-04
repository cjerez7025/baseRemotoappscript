/**
 * ========================================
 * MÓDULO: SISTEMA DE PERFILAMIENTO
 * ========================================
 * 
 * Gestiona los perfiles de usuarios (SUPERVISOR / EJECUTIVO)
 * Crea y mantiene la hoja CONFIG_PERFILES
 * 
 * FUNCIONALIDADES:
 * - Crear estructura CONFIG_PERFILES
 * - Registrar ejecutivos automáticamente
 * - Generar emails a partir de nombres
 * - Asignar roles por defecto
 * 
 * AUTOR: Sistema de Gestión de Llamadas
 * VERSIÓN: 1.0
 */

/**
 * Crea o actualiza la hoja CONFIG_PERFILES con la estructura necesaria
 */
function crearOActualizarConfigPerfiles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    // Si no existe, crearla
    if (!configSheet) {
      Logger.log('Creando hoja CONFIG_PERFILES...');
      configSheet = ss.insertSheet('CONFIG_PERFILES');
      
      // Configurar encabezados
      var encabezados = ['NOMBRE', 'EMAIL', 'ROL', 'HOJA_ASIGNADA', 'FECHA_CREACION', 'ULTIMA_ACTUALIZACION'];
      configSheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
      
      // Formato de encabezados
      var headerRange = configSheet.getRange(1, 1, 1, encabezados.length);
      headerRange.setBackground('#4CAF50');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // Ajustar anchos de columnas
      configSheet.setColumnWidth(1, 200); // NOMBRE
      configSheet.setColumnWidth(2, 250); // EMAIL
      configSheet.setColumnWidth(3, 120); // ROL
      configSheet.setColumnWidth(4, 200); // HOJA_ASIGNADA
      configSheet.setColumnWidth(5, 150); // FECHA_CREACION
      configSheet.setColumnWidth(6, 180); // ULTIMA_ACTUALIZACION
      
      // Congelar primera fila
      configSheet.setFrozenRows(1);
      
      // OCULTAR LA HOJA
      configSheet.hideSheet();
      
      Logger.log('✓ Hoja CONFIG_PERFILES creada exitosamente y ocultada');
    } else {
      Logger.log('Hoja CONFIG_PERFILES ya existe');
    }
    
    return configSheet;
    
  } catch (error) {
    Logger.log('ERROR en crearOActualizarConfigPerfiles: ' + error.toString());
    throw error;
  }
}

/**
 * Genera un email corporativo a partir del nombre del ejecutivo
 * Formato: nombre.apellido@empresa.com
 */
function generarEmailDesdeNombre(nombreCompleto) {
  try {
    // Limpiar el nombre
    var nombreLimpio = nombreCompleto.toString().trim();
    
    // Remover guiones bajos y reemplazar por espacios
    nombreLimpio = nombreLimpio.replace(/_/g, ' ');
    
    // Convertir a minúsculas
    nombreLimpio = nombreLimpio.toLowerCase();
    
    // Remover acentos
    nombreLimpio = nombreLimpio
      .replace(/á/g, 'a')
      .replace(/é/g, 'e')
      .replace(/í/g, 'i')
      .replace(/ó/g, 'o')
      .replace(/ú/g, 'u')
      .replace(/ñ/g, 'n');
    
    // Dividir en palabras
    var palabras = nombreLimpio.split(/\s+/);
    
    // Si tiene al menos 2 palabras (nombre y apellido)
    if (palabras.length >= 2) {
      // Tomar la primera palabra como nombre y la segunda como apellido
      return palabras[0] + '.' + palabras[1] + '@empresa.com';
    } else if (palabras.length === 1) {
      // Si solo tiene una palabra, usar esa
      return palabras[0] + '@empresa.com';
    }
    
    return nombreLimpio.replace(/\s+/g, '.') + '@empresa.com';
    
  } catch (error) {
    Logger.log('Error generando email: ' + error.toString());
    return nombreCompleto.toLowerCase().replace(/\s+/g, '.') + '@empresa.com';
  }
}

/**
 * Registra o actualiza ejecutivos en CONFIG_PERFILES
 * @param {Array} nombresEjecutivos - Array con nombres de hojas de ejecutivos
 * @return {Object} Resultado con contadores de nuevos y actualizados
 */
function registrarEjecutivosEnConfig(nombresEjecutivos) {
  try {
    Logger.log('=== REGISTRANDO EJECUTIVOS EN CONFIG_PERFILES ===');
    
    // Crear o obtener la hoja de configuración
    var configSheet = crearOActualizarConfigPerfiles();
    
    var resultado = {
      nuevos: 0,
      actualizados: 0,
      errores: 0
    };
    
    // Obtener datos existentes
    var ultimaFila = configSheet.getLastRow();
    var datosExistentes = {};
    
    if (ultimaFila > 1) {
      var rangoDatos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
      for (var i = 0; i < rangoDatos.length; i++) {
        var nombre = rangoDatos[i][0];
        if (nombre) {
          datosExistentes[nombre.toString()] = {
            fila: i + 2,
            email: rangoDatos[i][1],
            rol: rangoDatos[i][2],
            hoja: rangoDatos[i][3]
          };
        }
      }
    }
    
    Logger.log('Registros existentes: ' + Object.keys(datosExistentes).length);
    Logger.log('Ejecutivos a procesar: ' + nombresEjecutivos.length);
    
    var ahora = new Date();
    
    // Procesar cada ejecutivo
    for (var j = 0; j < nombresEjecutivos.length; j++) {
      try {
        var nombreHoja = nombresEjecutivos[j];
        var nombreEjecutivo = nombreHoja.replace(/_/g, ' ').toUpperCase();
        
        if (datosExistentes[nombreEjecutivo]) {
          // ACTUALIZAR: El ejecutivo ya existe
          var filaExistente = datosExistentes[nombreEjecutivo].fila;
          
          // Solo actualizar ULTIMA_ACTUALIZACION
          configSheet.getRange(filaExistente, 6).setValue(ahora);
          
          resultado.actualizados++;
          Logger.log('✓ Actualizado: ' + nombreEjecutivo);
          
        } else {
          // NUEVO: Agregar el ejecutivo
          var email = generarEmailDesdeNombre(nombreEjecutivo);
          
          var nuevaFila = [
            nombreEjecutivo,           // NOMBRE
            email,                      // EMAIL
            'EJECUTIVO',                // ROL (por defecto)
            nombreHoja,                 // HOJA_ASIGNADA
            ahora,                      // FECHA_CREACION
            ahora                       // ULTIMA_ACTUALIZACION
          ];
          
          configSheet.appendRow(nuevaFila);
          
          resultado.nuevos++;
          Logger.log('✓ Nuevo: ' + nombreEjecutivo + ' (' + email + ')');
        }
        
      } catch (errorEjec) {
        Logger.log('⚠️ Error procesando ' + nombreEjecutivo + ': ' + errorEjec.toString());
        resultado.errores++;
      }
    }
    
    // Aplicar formato a las nuevas filas
    if (resultado.nuevos > 0) {
      var rangoTotal = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 6);
      rangoTotal.setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
      
      // Alternar colores de filas
      var numFilas = configSheet.getLastRow() - 1;
      for (var k = 0; k < numFilas; k++) {
        var fila = k + 2;
        var color = (k % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
        configSheet.getRange(fila, 1, 1, 6).setBackground(color);
      }
      
      // Centrar columnas ROL y fechas
      configSheet.getRange(2, 3, numFilas, 1).setHorizontalAlignment('center'); // ROL
      configSheet.getRange(2, 5, numFilas, 2).setHorizontalAlignment('center'); // Fechas
    }
    
    Logger.log('=== RESUMEN REGISTRO DE PERFILES ===');
    Logger.log('Nuevos: ' + resultado.nuevos);
    Logger.log('Actualizados: ' + resultado.actualizados);
    Logger.log('Errores: ' + resultado.errores);
    
    return resultado;
    
  } catch (error) {
    Logger.log('ERROR CRÍTICO en registrarEjecutivosEnConfig: ' + error.toString());
    throw error;
  }
}

/**
 * Obtiene el rol de un usuario desde CONFIG_PERFILES
 * @param {string} email - Email del usuario
 * @return {string} 'SUPERVISOR' o 'EJECUTIVO' o 'NO_ENCONTRADO'
 */
function obtenerRolUsuario(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      Logger.log('CONFIG_PERFILES no existe');
      return 'NO_ENCONTRADO';
    }
    
    var ultimaFila = configSheet.getLastRow();
    if (ultimaFila < 2) {
      return 'NO_ENCONTRADO';
    }
    
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 3).getValues();
    
    for (var i = 0; i < datos.length; i++) {
      var emailRegistrado = datos[i][1];
      if (emailRegistrado && emailRegistrado.toString().toLowerCase() === email.toLowerCase()) {
        var rol = datos[i][2] || 'EJECUTIVO'; // Columna ROL
        // NORMALIZAR: Convertir a mayúsculas para evitar problemas
        return rol.toString().toUpperCase();
      }
    }
    
    return 'NO_ENCONTRADO';
    
  } catch (error) {
    Logger.log('Error obteniendo rol: ' + error.toString());
    return 'NO_ENCONTRADO';
  }
}

/**
 * Obtiene la hoja asignada a un ejecutivo
 * @param {string} email - Email del usuario
 * @return {string|null} Nombre de la hoja asignada o null
 */
function obtenerHojaAsignada(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      return null;
    }
    
    var ultimaFila = configSheet.getLastRow();
    if (ultimaFila < 2) {
      return null;
    }
    
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    
    for (var i = 0; i < datos.length; i++) {
      var emailRegistrado = datos[i][1];
      if (emailRegistrado && emailRegistrado.toString().toLowerCase() === email.toLowerCase()) {
        return datos[i][3]; // Columna HOJA_ASIGNADA
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('Error obteniendo hoja asignada: ' + error.toString());
    return null;
  }
}

/**
 * Función auxiliar para crear CONFIG_PERFILES manualmente desde el menú
 */
function crearConfigPerfilesManual() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var respuesta = ui.alert(
      '👥 Crear CONFIG_PERFILES',
      '¿Deseas crear/actualizar la hoja de configuración de perfiles?\n\n' +
      'Esta acción detectará automáticamente todos los ejecutivos actuales.',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) {
      return;
    }
    
    // Crear la hoja
    crearOActualizarConfigPerfiles();
    
    // Detectar ejecutivos actuales
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var ejecutivosDetectados = [];
    
    var hojasExcluidas = ['BBDD_REPORTE', 'RESUMEN', 'LLAMADAS', 'PRODUCTIVIDAD', 'CONFIG_PERFILES', 'CONFIGURACION'];
    
    for (var i = 0; i < hojas.length; i++) {
      var nombreHoja = hojas[i].getName();
      
      // Excluir hojas del sistema
      var esExcluida = false;
      for (var j = 0; j < hojasExcluidas.length; j++) {
        if (nombreHoja.indexOf(hojasExcluidas[j]) !== -1) {
          esExcluida = true;
          break;
        }
      }
      
      if (esExcluida) continue;
      if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) continue;
      
      // Esta es una hoja de ejecutivo
      if (hojas[i].getLastRow() > 1) {
        ejecutivosDetectados.push(nombreHoja);
      }
    }
    
    if (ejecutivosDetectados.length > 0) {
      var resultado = registrarEjecutivosEnConfig(ejecutivosDetectados);
      
      // Limpiar cualquier entrada de CONFIG_PERFILES que se haya colado
      limpiarConfigPerfilesDeListaEjecutivos();
      
      var mensaje = '✅ CONFIG_PERFILES actualizado\n\n';
      mensaje += '👥 Ejecutivos detectados: ' + ejecutivosDetectados.length + '\n';
      mensaje += '✨ Nuevos registros: ' + resultado.nuevos + '\n';
      mensaje += '🔄 Actualizados: ' + resultado.actualizados;
      
      if (resultado.errores > 0) {
        mensaje += '\n⚠️ Errores: ' + resultado.errores;
      }
      
      ui.alert('✅ Completado', mensaje, ui.ButtonSet.OK);
      
      // Mostrar la hoja temporalmente
      var configSheet = ss.getSheetByName('CONFIG_PERFILES');
      if (configSheet && configSheet.isSheetHidden()) {
        configSheet.showSheet();
      }
      ss.setActiveSheet(configSheet);
      
    } else {
      ui.alert('⚠️ Advertencia', 'No se detectaron hojas de ejecutivos con datos.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Error al crear CONFIG_PERFILES:\n\n' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Limpia cualquier registro de CONFIG_PERFILES que se haya registrado como ejecutivo
 */
function limpiarConfigPerfilesDeListaEjecutivos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) return;
    
    var ultimaFila = configSheet.getLastRow();
    if (ultimaFila < 2) return;
    
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    
    // Buscar y eliminar fila de CONFIG_PERFILES
    for (var i = datos.length - 1; i >= 0; i--) {
      var nombre = datos[i][0];
      var hojaAsignada = datos[i][3];
      
      if (nombre && (nombre.toString().toUpperCase().indexOf('CONFIG_PERFILES') !== -1 ||
                     hojaAsignada && hojaAsignada.toString().toUpperCase().indexOf('CONFIG_PERFILES') !== -1)) {
        configSheet.deleteRow(i + 2); // +2 por el encabezado y el índice 0
        Logger.log('✓ Eliminada fila incorrecta de CONFIG_PERFILES');
      }
    }
    
  } catch (error) {
    Logger.log('Error limpiando CONFIG_PERFILES de la lista: ' + error.toString());
  }
}

/**
 * Muestra la hoja CONFIG_PERFILES (para supervisores)
 */
function mostrarConfigPerfiles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      SpreadsheetApp.getUi().alert('⚠️ Advertencia', 'La hoja CONFIG_PERFILES no existe.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (configSheet.isSheetHidden()) {
      configSheet.showSheet();
    }
    
    ss.setActiveSheet(configSheet);
    SpreadsheetApp.getUi().alert('✅ CONFIG_PERFILES', 'Hoja mostrada correctamente.', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Error al mostrar CONFIG_PERFILES:\n\n' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Oculta la hoja CONFIG_PERFILES
 */
function ocultarConfigPerfiles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      return;
    }
    
    if (!configSheet.isSheetHidden()) {
      configSheet.hideSheet();
      Logger.log('✓ CONFIG_PERFILES ocultada');
    }
    
  } catch (error) {
    Logger.log('Error ocultando CONFIG_PERFILES: ' + error.toString());
  }
}