/**
 * ========================================
 * DIAGNÓSTICO AVANZADO DE CONFIG_PERFILES
 * ========================================
 * 
 * Detecta problemas de coincidencia de emails
 */

/**
 * Diagnóstico completo del sistema de perfiles
 */
function diagnosticarConfigPerfilesCompleto() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Obtener email del usuario
    var emailUsuario = obtenerEmailUsuarioRobusto();
    
    var mensaje = '🔍 DIAGNÓSTICO DE CONFIG_PERFILES\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (!emailUsuario) {
      mensaje += '❌ No se pudo obtener email del usuario';
      ui.alert('Error', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    mensaje += '📧 Email detectado:\n' + emailUsuario + '\n\n';
    mensaje += 'Longitud: ' + emailUsuario.length + ' caracteres\n';
    mensaje += 'En minúsculas: ' + emailUsuario.toLowerCase() + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    // 2. Verificar CONFIG_PERFILES
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      mensaje += '❌ CONFIG_PERFILES no existe';
      ui.alert('Error', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    var ultimaFila = configSheet.getLastRow();
    
    if (ultimaFila < 2) {
      mensaje += '❌ CONFIG_PERFILES está vacía';
      ui.alert('Error', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    mensaje += '📋 CONFIG_PERFILES encontrada\n';
    mensaje += 'Total de usuarios: ' + (ultimaFila - 1) + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    // 3. Leer todos los datos
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 3).getValues();
    
    mensaje += '🔎 BÚSQUEDA EN CONFIG_PERFILES:\n\n';
    
    var encontrado = false;
    
    for (var i = 0; i < datos.length; i++) {
      var nombre = datos[i][0];
      var emailRegistrado = datos[i][1];
      var rol = datos[i][2];
      
      // Comparación detallada
      if (emailRegistrado) {
        var emailRegLower = emailRegistrado.toString().toLowerCase().trim();
        var emailUsuLower = emailUsuario.toLowerCase().trim();
        
        var coincide = emailRegLower === emailUsuLower;
        
        if (coincide) {
          encontrado = true;
          mensaje += '✅ ¡COINCIDENCIA ENCONTRADA!\n\n';
          mensaje += 'Fila ' + (i + 2) + ':\n';
          mensaje += '  Nombre: ' + nombre + '\n';
          mensaje += '  Email: ' + emailRegistrado + '\n';
          mensaje += '  Rol: ' + rol + '\n\n';
        } else {
          // Mostrar solo si es similar (para debug)
          if (emailRegistrado.toString().indexOf(emailUsuario.substring(0, 10)) !== -1) {
            mensaje += '⚠️ Email similar (NO coincide):\n';
            mensaje += '  Fila ' + (i + 2) + ': ' + emailRegistrado + '\n';
            mensaje += '  Longitud: ' + emailRegistrado.toString().length + '\n\n';
          }
        }
      }
    }
    
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (!encontrado) {
      mensaje += '❌ TU EMAIL NO FUE ENCONTRADO\n\n';
      mensaje += 'Posibles causas:\n';
      mensaje += '• Email con espacios en blanco\n';
      mensaje += '• Caracteres invisibles\n';
      mensaje += '• Diferencia en mayúsculas\n\n';
      mensaje += '💡 SOLUCIÓN:\n';
      mensaje += 'Ejecuta: repararEmailsEnConfigPerfiles';
    } else {
      mensaje += '✅ TU EMAIL ESTÁ REGISTRADO\n\n';
      mensaje += 'Si no ves el menú correcto:\n';
      mensaje += '1. Recarga el archivo (F5)\n';
      mensaje += '2. Verifica que Perfiles.js esté actualizado';
    }
    
    Logger.log('=== DIAGNÓSTICO ===');
    Logger.log('Email usuario: [' + emailUsuario + ']');
    Logger.log('Encontrado: ' + encontrado);
    
    ui.alert('Diagnóstico de Perfiles', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Repara emails en CONFIG_PERFILES
 * Elimina espacios y normaliza
 */
function repararEmailsEnConfigPerfiles() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      ui.alert('Error', 'CONFIG_PERFILES no existe', ui.ButtonSet.OK);
      return;
    }
    
    var confirmar = ui.alert(
      '🔧 Reparar Emails',
      '¿Deseas limpiar y normalizar todos los emails en CONFIG_PERFILES?\n\n' +
      'Esto eliminará espacios en blanco y normalizará los emails.',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmar !== ui.Button.YES) {
      return;
    }
    
    var ultimaFila = configSheet.getLastRow();
    
    if (ultimaFila < 2) {
      ui.alert('Aviso', 'No hay datos para reparar', ui.ButtonSet.OK);
      return;
    }
    
    var reparados = 0;
    
    // Leer columna de emails (columna 2)
    var emails = configSheet.getRange(2, 2, ultimaFila - 1, 1).getValues();
    
    for (var i = 0; i < emails.length; i++) {
      if (emails[i][0]) {
        var emailOriginal = emails[i][0].toString();
        var emailLimpio = emailOriginal.trim().toLowerCase();
        
        if (emailOriginal !== emailLimpio) {
          configSheet.getRange(i + 2, 2).setValue(emailLimpio);
          reparados++;
          Logger.log('Reparado: [' + emailOriginal + '] → [' + emailLimpio + ']');
        }
      }
    }
    
    var mensaje = '✅ REPARACIÓN COMPLETADA\n\n';
    mensaje += 'Emails reparados: ' + reparados + '\n\n';
    
    if (reparados > 0) {
      mensaje += '🔄 Recarga el archivo (F5) para aplicar cambios.';
    } else {
      mensaje += 'No se encontraron emails con problemas.';
    }
    
    ui.alert('Completado', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Prueba directa de obtenerRolUsuario con el email actual
 */
function probarObtenerRolDirecto() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var email = obtenerEmailUsuarioRobusto();
    
    if (!email) {
      ui.alert('Error', 'No se pudo obtener email', ui.ButtonSet.OK);
      return;
    }
    
    Logger.log('=== PRUEBA DIRECTA ===');
    Logger.log('Email: ' + email);
    
    var rol = obtenerRolUsuario(email);
    
    Logger.log('Rol obtenido: ' + rol);
    
    var mensaje = '🧪 PRUEBA DE obtenerRolUsuario()\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += '📧 Email: ' + email + '\n\n';
    mensaje += '👔 Rol devuelto: ' + rol + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (rol === 'NO_ENCONTRADO') {
      mensaje += '❌ NO SE ENCONTRÓ EL ROL\n\n';
      mensaje += 'Ejecuta: diagnosticarConfigPerfilesCompleto\n';
      mensaje += 'para más detalles.';
    } else {
      mensaje += '✅ ROL ENCONTRADO CORRECTAMENTE\n\n';
      mensaje += 'Si no ves el menú correcto:\n';
      mensaje += '1. Verifica que onOpen() use obtenerEmailUsuarioRobusto()\n';
      mensaje += '2. Recarga el archivo (F5)';
    }
    
    ui.alert('Prueba de Rol', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}