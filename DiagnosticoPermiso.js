/**
 * ========================================
 * DIAGNÓSTICO DE PERMISOS Y USUARIOS
 * ========================================
 * 
 * Detecta problemas con Session.getActiveUser().getEmail()
 * para usuarios que no son propietarios
 */

/**
 * Diagnóstico completo del usuario actual
 */
function diagnosticarUsuarioActual() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Intentar diferentes métodos para obtener el email
    var metodos = [];
    
    // Método 1: Session.getActiveUser()
    try {
      var email1 = Session.getActiveUser().getEmail();
      metodos.push({
        metodo: 'Session.getActiveUser().getEmail()',
        resultado: email1 || '(VACÍO)',
        funciona: email1 && email1.length > 0
      });
    } catch (e) {
      metodos.push({
        metodo: 'Session.getActiveUser().getEmail()',
        resultado: 'ERROR: ' + e.message,
        funciona: false
      });
    }
    
    // Método 2: Session.getEffectiveUser()
    try {
      var email2 = Session.getEffectiveUser().getEmail();
      metodos.push({
        metodo: 'Session.getEffectiveUser().getEmail()',
        resultado: email2 || '(VACÍO)',
        funciona: email2 && email2.length > 0
      });
    } catch (e) {
      metodos.push({
        metodo: 'Session.getEffectiveUser().getEmail()',
        resultado: 'ERROR: ' + e.message,
        funciona: false
      });
    }
    
    // Método 3: Propietario del archivo
    try {
      var owner = ss.getOwner();
      metodos.push({
        metodo: 'Spreadsheet.getOwner().getEmail()',
        resultado: owner ? owner.getEmail() : '(NO DISPONIBLE)',
        funciona: false // No es el usuario actual
      });
    } catch (e) {
      metodos.push({
        metodo: 'Spreadsheet.getOwner().getEmail()',
        resultado: 'ERROR: ' + e.message,
        funciona: false
      });
    }
    
    // Método 4: Editores del archivo
    try {
      var editors = ss.getEditors();
      var emailsEditores = [];
      for (var i = 0; i < Math.min(editors.length, 5); i++) {
        emailsEditores.push(editors[i].getEmail());
      }
      metodos.push({
        metodo: 'Spreadsheet.getEditors() - Total: ' + editors.length,
        resultado: emailsEditores.join(', ') + (editors.length > 5 ? '...' : ''),
        funciona: false // Son todos los editores, no el actual
      });
    } catch (e) {
      metodos.push({
        metodo: 'Spreadsheet.getEditors()',
        resultado: 'ERROR: ' + e.message,
        funciona: false
      });
    }
    
    // Construir mensaje
    var mensaje = '🔍 DIAGNÓSTICO DE USUARIO\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    for (var j = 0; j < metodos.length; j++) {
      var m = metodos[j];
      mensaje += (m.funciona ? '✅' : '❌') + ' ' + m.metodo + '\n';
      mensaje += '   → ' + m.resultado + '\n\n';
    }
    
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    // Determinar el problema
    var emailActual = metodos[0].resultado;
    var emailEffective = metodos[1].resultado;
    
    if (!metodos[0].funciona && !metodos[1].funciona) {
      mensaje += '❌ PROBLEMA CRÍTICO:\n\n';
      mensaje += 'No se puede obtener el email del usuario.\n\n';
      mensaje += 'SOLUCIONES:\n';
      mensaje += '1. El archivo debe compartirse desde un dominio de Google Workspace\n';
      mensaje += '2. O implementar autenticación manual\n';
    } else if (!metodos[0].funciona && metodos[1].funciona) {
      mensaje += '⚠️ PROBLEMA PARCIAL:\n\n';
      mensaje += 'Session.getActiveUser() no funciona\n';
      mensaje += 'pero Session.getEffectiveUser() SÍ funciona.\n\n';
      mensaje += '✅ SOLUCIÓN: Usar getEffectiveUser() en el código\n';
    } else if (metodos[0].funciona) {
      mensaje += '✅ TODO CORRECTO:\n\n';
      mensaje += 'El sistema puede identificar al usuario.\n';
      mensaje += 'Email: ' + emailActual;
    }
    
    // Logging
    Logger.log('=== DIAGNÓSTICO DE USUARIO ===');
    for (var k = 0; k < metodos.length; k++) {
      Logger.log(metodos[k].metodo + ': ' + metodos[k].resultado);
    }
    
    ui.alert('Diagnóstico de Usuario', mensaje, ui.ButtonSet.OK);
    
    return metodos;
    
  } catch (error) {
    Logger.log('Error en diagnóstico: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * Obtiene el email del usuario de forma robusta
 * Prueba múltiples métodos
 */
function obtenerEmailUsuarioRobusto() {
  try {
    // Método 1: getActiveUser (más confiable si funciona)
    var email = Session.getActiveUser().getEmail();
    if (email && email.length > 0 && email.indexOf('@') !== -1) {
      Logger.log('Email obtenido con getActiveUser: ' + email);
      return email;
    }
    
    // Método 2: getEffectiveUser (fallback)
    email = Session.getEffectiveUser().getEmail();
    if (email && email.length > 0 && email.indexOf('@') !== -1) {
      Logger.log('Email obtenido con getEffectiveUser: ' + email);
      return email;
    }
    
    // Si ninguno funciona
    Logger.log('⚠️ No se pudo obtener email del usuario');
    return null;
    
  } catch (error) {
    Logger.log('Error obteniendo email: ' + error.toString());
    return null;
  }
}

/**
 * Prueba el sistema de perfiles con el usuario actual
 */
function probarSistemaPerfilesConDiagnostico() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    Logger.log('=== PRUEBA DE SISTEMA DE PERFILES ===');
    
    // 1. Diagnóstico de email
    var email = obtenerEmailUsuarioRobusto();
    
    var mensaje = '🔍 PRUEBA DE PERFILES\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (!email) {
      mensaje += '❌ NO SE PUDO OBTENER EMAIL\n\n';
      mensaje += 'El sistema no puede identificar al usuario.\n\n';
      mensaje += 'Ejecuta: diagnosticarUsuarioActual\n';
      mensaje += 'para ver detalles del problema.';
      
      ui.alert('Error', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    mensaje += '✅ Email detectado:\n' + email + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    // 2. Buscar en CONFIG_PERFILES
    var rol = obtenerRolUsuario(email);
    var hoja = obtenerHojaAsignada(email);
    
    mensaje += 'ROL: ' + rol + '\n';
    mensaje += 'HOJA: ' + (hoja || 'No asignada') + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (rol === 'NO_ENCONTRADO') {
      mensaje += '⚠️ USUARIO NO REGISTRADO\n\n';
      mensaje += 'Tu email no está en CONFIG_PERFILES.\n\n';
      mensaje += 'Pide al supervisor que te agregue.';
    } else {
      mensaje += '✅ USUARIO REGISTRADO\n\n';
      mensaje += 'El sistema te reconoce correctamente.';
    }
    
    Logger.log('Email: ' + email);
    Logger.log('Rol: ' + rol);
    Logger.log('Hoja: ' + (hoja || 'Ninguna'));
    
    ui.alert('Prueba de Perfiles', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Verifica permisos del archivo actual
 */
function verificarPermisosArchivo() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var mensaje = '📋 PERMISOS DEL ARCHIVO\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    // Propietario
    try {
      var owner = ss.getOwner();
      mensaje += '👑 Propietario:\n' + (owner ? owner.getEmail() : 'Desconocido') + '\n\n';
    } catch (e) {
      mensaje += '👑 Propietario: ERROR\n\n';
    }
    
    // Editores
    try {
      var editors = ss.getEditors();
      mensaje += '✏️ Editores (' + editors.length + '):\n';
      for (var i = 0; i < Math.min(editors.length, 10); i++) {
        mensaje += '  • ' + editors[i].getEmail() + '\n';
      }
      if (editors.length > 10) {
        mensaje += '  ... y ' + (editors.length - 10) + ' más\n';
      }
      mensaje += '\n';
    } catch (e) {
      mensaje += '✏️ Editores: ERROR\n\n';
    }
    
    // Viewers
    try {
      var viewers = ss.getViewers();
      mensaje += '👁️ Lectores (' + viewers.length + '):\n';
      if (viewers.length === 0) {
        mensaje += '  (ninguno)\n';
      } else {
        for (var j = 0; j < Math.min(viewers.length, 5); j++) {
          mensaje += '  • ' + viewers[j].getEmail() + '\n';
        }
        if (viewers.length > 5) {
          mensaje += '  ... y ' + (viewers.length - 5) + ' más\n';
        }
      }
    } catch (e) {
      mensaje += '👁️ Lectores: ERROR\n';
    }
    
    ui.alert('Permisos del Archivo', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}