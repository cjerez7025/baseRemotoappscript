/**
 * ========================================
 * DIAGNÓSTICO ESPECÍFICO PARA LORENA
 * ========================================
 * Para identificar exactamente por qué no ve el menú correcto
 */

/**
 * Diagnóstico completo del problema de menú
 */
function diagnosticarProblemaLorena() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var mensaje = '🔍 DIAGNÓSTICO COMPLETO\n';
    mensaje += '═══════════════════════════════\n\n';
    
    // PASO 1: Obtener email
    Logger.log('=== PASO 1: OBTENCIÓN DE EMAIL ===');
    var email = obtenerEmailUsuarioRobusto();
    
    mensaje += '📧 PASO 1: Email detectado\n';
    mensaje += '   Email: ' + (email || 'NO DETECTADO') + '\n';
    mensaje += '   Longitud: ' + (email ? email.length : 0) + ' caracteres\n';
    mensaje += '   En minúsculas: ' + (email ? email.toLowerCase() : 'N/A') + '\n\n';
    
    Logger.log('Email detectado: "' + email + '"');
    
    if (!email) {
      mensaje += '❌ PROBLEMA: No se puede obtener email\n';
      ui.alert('❌ Error Crítico', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    // PASO 2: Verificar CONFIG_PERFILES
    Logger.log('=== PASO 2: VERIFICACIÓN DE CONFIG_PERFILES ===');
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    mensaje += '📋 PASO 2: CONFIG_PERFILES\n';
    
    if (!configSheet) {
      mensaje += '   ❌ NO EXISTE\n\n';
      mensaje += '🔧 SOLUCIÓN:\n';
      mensaje += '   Ejecuta "🔄 Actualizar CONFIG_PERFILES"\n';
      ui.alert('❌ Error', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    mensaje += '   ✅ Existe\n';
    
    var ultimaFila = configSheet.getLastRow();
    mensaje += '   Total usuarios: ' + (ultimaFila - 1) + '\n\n';
    
    Logger.log('CONFIG_PERFILES existe con ' + (ultimaFila - 1) + ' usuarios');
    
    // PASO 3: Buscar usuario en CONFIG_PERFILES
    Logger.log('=== PASO 3: BÚSQUEDA EN CONFIG_PERFILES ===');
    
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    var encontrado = false;
    var filaEncontrada = -1;
    var rolEncontrado = '';
    var emailEncontrado = '';
    
    mensaje += '🔎 PASO 3: Búsqueda de usuario\n';
    mensaje += '   Buscando: ' + email.toLowerCase() + '\n\n';
    
    for (var i = 0; i < datos.length; i++) {
      var nombre = datos[i][0];
      var emailFila = datos[i][1];
      var rol = datos[i][2];
      var hoja = datos[i][3];
      
      if (!emailFila) continue;
      
      var emailFilaLimpio = emailFila.toString().trim().toLowerCase();
      var emailUsuarioLimpio = email.trim().toLowerCase();
      
      Logger.log('Comparando:');
      Logger.log('  - Usuario: "' + emailUsuarioLimpio + '"');
      Logger.log('  - Fila ' + (i+2) + ': "' + emailFilaLimpio + '"');
      Logger.log('  - ¿Coinciden? ' + (emailFilaLimpio === emailUsuarioLimpio));
      
      if (emailFilaLimpio === emailUsuarioLimpio) {
        encontrado = true;
        filaEncontrada = i + 2;
        rolEncontrado = rol;
        emailEncontrado = emailFila;
        
        mensaje += '   ✅ ¡ENCONTRADO!\n';
        mensaje += '   Fila: ' + filaEncontrada + '\n';
        mensaje += '   Nombre: ' + nombre + '\n';
        mensaje += '   Email registrado: ' + emailFila + '\n';
        mensaje += '   Rol: ' + rol + '\n';
        mensaje += '   Hoja: ' + (hoja || 'No asignada') + '\n\n';
        
        Logger.log('✓ Usuario encontrado en fila ' + filaEncontrada + ' como ' + rol);
        break;
      }
    }
    
    if (!encontrado) {
      mensaje += '   ❌ NO ENCONTRADO\n\n';
      mensaje += '🔧 SOLUCIÓN:\n';
      mensaje += '   1. Agrega manualmente a CONFIG_PERFILES:\n';
      mensaje += '      Email: ' + email + '\n';
      mensaje += '      Rol: SUPERVISOR o EJECUTIVO\n';
      mensaje += '   2. O ejecuta "🔄 Sincronizar Usuarios"\n';
      
      Logger.log('❌ Usuario NO encontrado en CONFIG_PERFILES');
      ui.alert('❌ Usuario No Registrado', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    // PASO 4: Verificar función obtenerRolUsuario
    Logger.log('=== PASO 4: VERIFICACIÓN DE obtenerRolUsuario ===');
    
    var rolDevuelto = obtenerRolUsuario(email);
    
    mensaje += '🎯 PASO 4: Función obtenerRolUsuario\n';
    mensaje += '   Email enviado: ' + email + '\n';
    mensaje += '   Rol devuelto: ' + rolDevuelto + '\n\n';
    
    Logger.log('obtenerRolUsuario("' + email + '") devuelve: "' + rolDevuelto + '"');
    
    // PASO 5: Análisis del problema
    Logger.log('=== PASO 5: ANÁLISIS ===');
    
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n';
    mensaje += '📊 RESUMEN DEL ANÁLISIS\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (rolDevuelto === rolEncontrado) {
      mensaje += '✅ TODO CORRECTO\n\n';
      mensaje += 'Deberías ver el menú de ' + rolDevuelto + '\n\n';
      mensaje += '🔄 SOLUCIONES:\n';
      mensaje += '1. Recarga la página (F5)\n';
      mensaje += '2. Cierra y vuelve a abrir el archivo\n';
      mensaje += '3. Si el problema persiste, limpia caché del navegador\n';
    } else {
      mensaje += '❌ INCONSISTENCIA DETECTADA\n\n';
      mensaje += 'Rol en CONFIG_PERFILES: ' + rolEncontrado + '\n';
      mensaje += 'Rol devuelto por función: ' + rolDevuelto + '\n\n';
      mensaje += '🔧 SOLUCIÓN:\n';
      mensaje += 'Hay un problema en la función obtenerRolUsuario()\n';
      mensaje += 'Contacta al administrador del sistema\n';
    }
    
    ui.alert('🔍 Diagnóstico Completo', mensaje, ui.ButtonSet.OK);
    
    // Guardar log completo
    Logger.log('=== FIN DEL DIAGNÓSTICO ===');
    Logger.log('Usuario: ' + email);
    Logger.log('Encontrado en fila: ' + filaEncontrada);
    Logger.log('Rol esperado: ' + rolEncontrado);
    Logger.log('Rol devuelto: ' + rolDevuelto);
    
  } catch (error) {
    Logger.log('Error en diagnóstico: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'Error ejecutando diagnóstico:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Diagnóstico ultra simplificado
 */
function diagnosticoRapidoLorena() {
  try {
    var email = obtenerEmailUsuarioRobusto();
    var rol = obtenerRolUsuario(email);
    
    var mensaje = '📧 Email: ' + email + '\n';
    mensaje += '👔 Rol: ' + rol + '\n\n';
    
    if (rol === 'SUPERVISOR') {
      mensaje += '✅ Deberías ver menú de SUPERVISOR\n\n';
      mensaje += 'Si no lo ves, recarga (F5)';
    } else if (rol === 'EJECUTIVO') {
      mensaje += '✅ Deberías ver menú de EJECUTIVO\n\n';
      mensaje += 'Si no lo ves, recarga (F5)';
    } else {
      mensaje += '❌ No tienes rol asignado\n\n';
      mensaje += 'Contacta al supervisor';
    }
    
    SpreadsheetApp.getUi().alert('Diagnóstico Rápido', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reparar problema de Lorena específicamente
 */
function repararProblemaLorena() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var respuesta = ui.alert(
      '🔧 Reparar Problema',
      '¿Deseas buscar y reparar el registro de lorenasotomayor75@gmail.com en CONFIG_PERFILES?',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) return;
    
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      ui.alert('❌ Error', 'CONFIG_PERFILES no existe', ui.ButtonSet.OK);
      return;
    }
    
    var emailBuscado = 'lorenasotomayor75@gmail.com';
    var ultimaFila = configSheet.getLastRow();
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    
    var encontrado = false;
    var filaEncontrada = -1;
    
    for (var i = 0; i < datos.length; i++) {
      var emailFila = datos[i][1];
      
      if (emailFila && emailFila.toString().trim().toLowerCase() === emailBuscado.toLowerCase()) {
        encontrado = true;
        filaEncontrada = i + 2;
        break;
      }
    }
    
    if (encontrado) {
      var mensaje = '✅ Usuario encontrado en fila ' + filaEncontrada + '\n\n';
      mensaje += 'Email: ' + emailBuscado + '\n';
      mensaje += 'Rol actual: ' + datos[filaEncontrada - 2][2] + '\n\n';
      mensaje += '¿Qué deseas hacer?\n';
      mensaje += '1. Cambiar a SUPERVISOR\n';
      mensaje += '2. Cambiar a EJECUTIVO\n';
      mensaje += '3. Cancelar';
      
      var accion = ui.prompt('Reparar Usuario', mensaje, ui.ButtonSet.OK_CANCEL);
      
      if (accion.getSelectedButton() === ui.Button.OK) {
        var opcion = accion.getResponseText();
        
        if (opcion === '1') {
          configSheet.getRange(filaEncontrada, 3).setValue('SUPERVISOR');
          ui.alert('✅ Actualizado', 'Usuario configurado como SUPERVISOR.\n\nRecarga el archivo (F5)', ui.ButtonSet.OK);
        } else if (opcion === '2') {
          configSheet.getRange(filaEncontrada, 3).setValue('EJECUTIVO');
          ui.alert('✅ Actualizado', 'Usuario configurado como EJECUTIVO.\n\nRecarga el archivo (F5)', ui.ButtonSet.OK);
        }
      }
      
    } else {
      ui.alert(
        '❌ No Encontrado',
        'El usuario ' + emailBuscado + ' NO está en CONFIG_PERFILES.\n\n' +
        'Agrégalo manualmente o ejecuta "🔄 Sincronizar Usuarios"',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}