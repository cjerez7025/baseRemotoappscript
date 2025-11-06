/**
 * ========================================
 * MÓDULO: DIAGNÓSTICO DE PERFILES
 * ========================================
 * Herramientas para identificar problemas con el sistema de perfilamiento
 */

/**
 * DIAGNÓSTICO COMPLETO DEL SISTEMA DE PERFILES
 * Muestra toda la información relevante en un diálogo
 */
function diagnosticarSistemaPerfiles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    Logger.log('=== DIAGNÓSTICO DE PERFILES ===');
    
    // 1. Obtener email del usuario actual
    var emailUsuario = Session.getActiveUser().getEmail();
    Logger.log('Email detectado: "' + emailUsuario + '"');
    
    // 2. Verificar si CONFIG_PERFILES existe
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    var mensaje = '🔍 DIAGNÓSTICO DEL SISTEMA DE PERFILES\n\n';
    mensaje += '═══════════════════════════════\n\n';
    
    // Información del usuario actual
    mensaje += '👤 USUARIO ACTUAL\n';
    mensaje += '📧 Email detectado: ' + (emailUsuario || '(VACÍO)') + '\n';
    mensaje += '🔒 Email válido: ' + (emailUsuario && emailUsuario !== '' ? 'SÍ' : 'NO') + '\n\n';
    
    if (!configSheet) {
      mensaje += '❌ PROBLEMA CRÍTICO\n';
      mensaje += 'La hoja CONFIG_PERFILES NO EXISTE\n\n';
      mensaje += '🔧 SOLUCIÓN:\n';
      mensaje += '1. Ve a menú "🎯 Gestión Supervisores"\n';
      mensaje += '2. Click en "🔄 Actualizar CONFIG_PERFILES"\n';
      
      ui.alert('❌ Error Crítico', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    mensaje += '✅ CONFIG_PERFILES existe\n\n';
    
    // 3. Leer todos los perfiles
    var ultimaFila = configSheet.getLastRow();
    Logger.log('Última fila en CONFIG_PERFILES: ' + ultimaFila);
    
    if (ultimaFila < 2) {
      mensaje += '❌ PROBLEMA\n';
      mensaje += 'CONFIG_PERFILES está vacía (sin usuarios)\n\n';
      mensaje += '🔧 SOLUCIÓN:\n';
      mensaje += '1. Ejecuta el proceso de distribución\n';
      mensaje += '2. O usa "🔄 Actualizar CONFIG_PERFILES"\n';
      
      ui.alert('⚠️ Configuración Vacía', mensaje, ui.ButtonSet.OK);
      return;
    }
    
    mensaje += '📋 USUARIOS REGISTRADOS (' + (ultimaFila - 1) + '):\n';
    mensaje += '─────────────────────────────\n';
    
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    var encontrado = false;
    var rolEncontrado = '';
    var hojaEncontrada = '';
    
    for (var i = 0; i < datos.length; i++) {
      var nombre = datos[i][0];
      var email = datos[i][1];
      var rol = datos[i][2];
      var hoja = datos[i][3];
      
      if (!email) continue;
      
      // Mostrar en el log
      Logger.log((i+1) + '. ' + nombre + ' | ' + email + ' | ' + rol);
      
      // Comparar con el usuario actual (ignorando mayúsculas)
      var emailLimpio = email.toString().trim().toLowerCase();
      var emailUsuarioLimpio = emailUsuario ? emailUsuario.trim().toLowerCase() : '';
      
      var esUsuarioActual = (emailLimpio === emailUsuarioLimpio);
      
      if (esUsuarioActual) {
        encontrado = true;
        rolEncontrado = rol;
        hojaEncontrada = hoja;
        mensaje += '👉 ';
      } else {
        mensaje += '   ';
      }
      
      mensaje += (i+1) + '. ' + nombre + '\n';
      mensaje += '   📧 ' + email + '\n';
      mensaje += '   👔 Rol: ' + (rol || 'SIN ROL') + '\n';
      
      if (hoja) {
        mensaje += '   📊 Hoja: ' + hoja + '\n';
      }
      
      mensaje += '\n';
    }
    
    mensaje += '─────────────────────────────\n\n';
    
    // 4. Resultado del diagnóstico
    if (!emailUsuario || emailUsuario === '') {
      mensaje += '❌ PROBLEMA DETECTADO\n';
      mensaje += 'Google NO puede identificar tu email\n\n';
      mensaje += '🔧 POSIBLES CAUSAS:\n';
      mensaje += '• Archivo compartido de forma pública\n';
      mensaje += '• Sesión anónima\n\n';
      mensaje += '💡 SOLUCIÓN:\n';
      mensaje += '1. Cierra este archivo\n';
      mensaje += '2. El propietario debe:\n';
      mensaje += '   - Ir a "Compartir"\n';
      mensaje += '   - AGREGAR tu email específico\n';
      mensaje += '   - Dar permiso "Editor"\n';
      mensaje += '3. Vuelve a abrir el archivo\n';
      
    } else if (encontrado) {
      mensaje += '✅ PERFIL ENCONTRADO\n';
      mensaje += '📧 Email: ' + emailUsuario + '\n';
      mensaje += '👔 Rol asignado: ' + (rolEncontrado || 'NO DEFINIDO') + '\n';
      
      if (hojaEncontrada) {
        mensaje += '📊 Hoja asignada: ' + hojaEncontrada + '\n';
      }
      
      mensaje += '\n🎯 ESTADO: CONFIGURACIÓN CORRECTA\n';
      
      // Verificar si el rol está correcto
      if (!rolEncontrado || rolEncontrado === '') {
        mensaje += '\n⚠️ ADVERTENCIA:\n';
        mensaje += 'Tu perfil no tiene ROL asignado\n';
        mensaje += 'Actualiza CONFIG_PERFILES desde el menú\n';
      }
      
    } else {
      mensaje += '❌ PERFIL NO ENCONTRADO\n';
      mensaje += 'Tu email NO está en CONFIG_PERFILES\n\n';
      mensaje += '📧 Email buscado:\n';
      mensaje += emailUsuario + '\n\n';
      mensaje += '🔧 SOLUCIÓN:\n';
      mensaje += '1. El supervisor debe:\n';
      mensaje += '   - Ir a "👥 Ver CONFIG_PERFILES"\n';
      mensaje += '   - Agregar tu email manualmente:\n';
      mensaje += '     • NOMBRE: Tu nombre\n';
      mensaje += '     • EMAIL: ' + emailUsuario + '\n';
      mensaje += '     • ROL: EJECUTIVO o SUPERVISOR\n';
      mensaje += '     • HOJA_ASIGNADA: Nombre de tu hoja\n\n';
      mensaje += '2. O ejecutar distribución de nuevo\n';
    }
    
    mensaje += '\n═══════════════════════════════\n';
    
    // Mostrar el diálogo
    ui.alert('🔍 Diagnóstico de Perfiles', mensaje, ui.ButtonSet.OK);
    
    Logger.log('=== FIN DIAGNÓSTICO ===');
    
  } catch (error) {
    Logger.log('ERROR en diagnóstico: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'Error ejecutando diagnóstico:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Compara en detalle dos emails para debug
 */
function compararEmails(email1, email2) {
  Logger.log('=== COMPARACIÓN DE EMAILS ===');
  Logger.log('Email 1: "' + email1 + '"');
  Logger.log('Email 2: "' + email2 + '"');
  Logger.log('Longitud 1: ' + (email1 ? email1.length : 0));
  Logger.log('Longitud 2: ' + (email2 ? email2.length : 0));
  Logger.log('Iguales (exacto): ' + (email1 === email2));
  Logger.log('Iguales (lowercase): ' + (email1.toLowerCase() === email2.toLowerCase()));
  Logger.log('Iguales (trim): ' + (email1.trim() === email2.trim()));
}

/**
 * CORRECCIÓN RÁPIDA: Agregar usuario manualmente
 */
function agregarUsuarioManual() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Verificar que CONFIG_PERFILES existe
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    if (!configSheet) {
      ui.alert('❌ Error', 'CONFIG_PERFILES no existe. Primero créala desde el menú.', ui.ButtonSet.OK);
      return;
    }
    
    // Solicitar datos
    var respNombre = ui.prompt(
      '👤 Agregar Usuario - Paso 1/4',
      'Ingresa el NOMBRE COMPLETO del usuario:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (respNombre.getSelectedButton() !== ui.Button.OK) return;
    var nombre = respNombre.getResponseText().trim();
    
    var respEmail = ui.prompt(
      '📧 Agregar Usuario - Paso 2/4',
      'Ingresa el EMAIL del usuario:\n\n' +
      'Ejemplo: usuario@gmail.com',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (respEmail.getSelectedButton() !== ui.Button.OK) return;
    var email = respEmail.getResponseText().trim().toLowerCase();
    
    var respRol = ui.prompt(
      '👔 Agregar Usuario - Paso 3/4',
      'Ingresa el ROL del usuario:\n\n' +
      '1 = SUPERVISOR\n' +
      '2 = EJECUTIVO\n\n' +
      'Ingresa 1 o 2:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (respRol.getSelectedButton() !== ui.Button.OK) return;
    var rolNum = respRol.getResponseText().trim();
    var rol = (rolNum === '1') ? 'SUPERVISOR' : 'EJECUTIVO';
    
    var hoja = '';
    if (rol === 'EJECUTIVO') {
      var respHoja = ui.prompt(
        '📊 Agregar Usuario - Paso 4/4',
        'Ingresa el NOMBRE DE LA HOJA asignada:\n\n' +
        'Ejemplo: JUAN_PEREZ',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (respHoja.getSelectedButton() === ui.Button.OK) {
        hoja = respHoja.getResponseText().trim();
      }
    }
    
    // Validaciones
    if (!nombre || !email) {
      ui.alert('❌ Error', 'Nombre y email son obligatorios', ui.ButtonSet.OK);
      return;
    }
    
    if (email.indexOf('@') === -1) {
      ui.alert('❌ Error', 'Email inválido (debe contener @)', ui.ButtonSet.OK);
      return;
    }
    
    // Verificar si ya existe
    var ultimaFila = configSheet.getLastRow();
    if (ultimaFila > 1) {
      var datosExistentes = configSheet.getRange(2, 2, ultimaFila - 1, 1).getValues();
      for (var i = 0; i < datosExistentes.length; i++) {
        if (datosExistentes[i][0].toString().toLowerCase() === email) {
          ui.alert('⚠️ Advertencia', 'Este email ya está registrado en la fila ' + (i + 2), ui.ButtonSet.OK);
          return;
        }
      }
    }
    
    // Agregar el nuevo usuario
    var nuevaFila = ultimaFila + 1;
    var ahora = new Date();
    
    configSheet.getRange(nuevaFila, 1).setValue(nombre);
    configSheet.getRange(nuevaFila, 2).setValue(email);
    configSheet.getRange(nuevaFila, 3).setValue(rol);
    configSheet.getRange(nuevaFila, 4).setValue(hoja);
    configSheet.getRange(nuevaFila, 5).setValue(ahora);
    configSheet.getRange(nuevaFila, 6).setValue(ahora);
    
    // Aplicar formato
    var color = (nuevaFila % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
    configSheet.getRange(nuevaFila, 1, 1, 6).setBackground(color);
    
    Logger.log('✓ Usuario agregado: ' + nombre + ' (' + email + ') - ' + rol);
    
    var mensaje = '✅ USUARIO AGREGADO EXITOSAMENTE\n\n';
    mensaje += '👤 Nombre: ' + nombre + '\n';
    mensaje += '📧 Email: ' + email + '\n';
    mensaje += '👔 Rol: ' + rol + '\n';
    if (hoja) {
      mensaje += '📊 Hoja: ' + hoja + '\n';
    }
    mensaje += '\n✨ El usuario ya puede usar el sistema';
    
    ui.alert('✅ Completado', mensaje, ui.ButtonSet.OK);
    
    // Mostrar la hoja
    if (configSheet.isSheetHidden()) {
      configSheet.showSheet();
    }
    ss.setActiveSheet(configSheet);
    
  } catch (error) {
    Logger.log('ERROR agregando usuario: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * CORRECCIÓN: Sincronizar emails de usuarios con acceso
 * Lee los permisos del archivo y actualiza CONFIG_PERFILES
 */
function sincronizarUsuariosConAcceso() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var respuesta = ui.alert(
      '🔄 Sincronizar Usuarios',
      '¿Deseas sincronizar los usuarios que tienen acceso al archivo con CONFIG_PERFILES?\n\n' +
      'Esto detectará automáticamente a todos los usuarios con permisos de edición.',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) return;
    
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    if (!configSheet) {
      ui.alert('❌ Error', 'Primero crea CONFIG_PERFILES desde el menú', ui.ButtonSet.OK);
      return;
    }
    
    // Obtener usuarios con acceso usando DriveApp
    var file = DriveApp.getFileById(ss.getId());
    var editors = file.getEditors();
    var viewers = file.getViewers();
    
    Logger.log('Editores detectados: ' + editors.length);
    Logger.log('Visualizadores: ' + viewers.length);
    
    var emailsNuevos = [];
    var nombresNuevos = [];
    
    // Procesar editores
    for (var i = 0; i < editors.length; i++) {
      var email = editors[i].getEmail();
      var nombre = editors[i].getName() || email.split('@')[0];
      
      Logger.log('Editor ' + (i+1) + ': ' + nombre + ' (' + email + ')');
      
      // Verificar si ya existe en CONFIG_PERFILES
      var existe = false;
      var ultimaFila = configSheet.getLastRow();
      
      if (ultimaFila > 1) {
        var datosExistentes = configSheet.getRange(2, 2, ultimaFila - 1, 1).getValues();
        for (var j = 0; j < datosExistentes.length; j++) {
          if (datosExistentes[j][0].toString().toLowerCase() === email.toLowerCase()) {
            existe = true;
            break;
          }
        }
      }
      
      if (!existe) {
        emailsNuevos.push(email);
        nombresNuevos.push(nombre);
      }
    }
    
    if (emailsNuevos.length === 0) {
      ui.alert(
        'ℹ️ Sin Cambios',
        'Todos los usuarios con acceso ya están registrados en CONFIG_PERFILES',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Agregar usuarios nuevos
    var mensaje = '🆕 SE DETECTARON ' + emailsNuevos.length + ' USUARIO(S) NUEVO(S):\n\n';
    
    for (var k = 0; k < emailsNuevos.length; k++) {
      mensaje += (k+1) + '. ' + nombresNuevos[k] + '\n';
      mensaje += '   📧 ' + emailsNuevos[k] + '\n\n';
    }
    
    mensaje += '¿Deseas agregarlos como EJECUTIVOS?';
    
    var confirmar = ui.alert('🆕 Usuarios Nuevos', mensaje, ui.ButtonSet.YES_NO);
    
    if (confirmar !== ui.Button.YES) return;
    
    // Agregar a CONFIG_PERFILES
    var ultimaFila = configSheet.getLastRow();
    var ahora = new Date();
    
    for (var m = 0; m < emailsNuevos.length; m++) {
      var fila = ultimaFila + m + 1;
      
      configSheet.getRange(fila, 1).setValue(nombresNuevos[m]);
      configSheet.getRange(fila, 2).setValue(emailsNuevos[m]);
      configSheet.getRange(fila, 3).setValue('EJECUTIVO');
      configSheet.getRange(fila, 4).setValue(''); // Sin hoja asignada por ahora
      configSheet.getRange(fila, 5).setValue(ahora);
      configSheet.getRange(fila, 6).setValue(ahora);
      
      var color = (fila % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
      configSheet.getRange(fila, 1, 1, 6).setBackground(color);
    }
    
    ui.alert(
      '✅ Sincronización Completada',
      emailsNuevos.length + ' usuario(s) agregado(s) exitosamente.\n\n' +
      'Ahora necesitas asignarles las hojas correspondientes.',
      ui.ButtonSet.OK
    );
    
    // Mostrar CONFIG_PERFILES
    if (configSheet.isSheetHidden()) {
      configSheet.showSheet();
    }
    ss.setActiveSheet(configSheet);
    
  } catch (error) {
    Logger.log('ERROR sincronizando usuarios: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}