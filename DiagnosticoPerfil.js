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
 * ========================================
 * SINCRONIZAR USUARIOS - VERSIÓN MEJORADA
 * ========================================
 * Detecta:
 * 1. Usuarios con acceso al archivo (editores)
 * 2. Hojas de ejecutivos creadas en el sistema
 * 
 * Los agrega automáticamente a CONFIG_PERFILES
 */

function sincronizarUsuariosConAcceso() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var respuesta = ui.alert(
      '🔄 Sincronizar Usuarios',
      '¿Deseas sincronizar usuarios?\n\n' +
      'Esto detectará:\n' +
      '• Usuarios con acceso al archivo\n' +
      '• Hojas de ejecutivos creadas\n\n' +
      'Y los agregará a CONFIG_PERFILES',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) return;
    
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    if (!configSheet) {
      ui.alert('❌ Error', 'Primero crea CONFIG_PERFILES desde el menú', ui.ButtonSet.OK);
      return;
    }
    
    Logger.log('=== SINCRONIZACIÓN DE USUARIOS ===');
    
    var usuariosNuevos = [];
    var ahora = new Date();
    
    // ═══════════════════════════════════════
    // PARTE 1: DETECTAR USUARIOS CON ACCESO
    // ═══════════════════════════════════════
    
    Logger.log('');
    Logger.log('PARTE 1: Detectando usuarios con acceso al archivo...');
    
    var file = DriveApp.getFileById(ss.getId());
    var editores = file.getEditors();
    
    Logger.log('Total editores con acceso: ' + editores.length);
    
    for (var i = 0; i < editores.length; i++) {
      var editor = editores[i];
      var email = editor.getEmail();
      var nombre = editor.getName() || email.split('@')[0];
      
      Logger.log((i + 1) + '. Editor: ' + nombre + ' (' + email + ')');
      
      // Verificar si ya existe en CONFIG_PERFILES
      if (!existeEnConfigPerfiles(configSheet, email)) {
        usuariosNuevos.push({
          nombre: nombre,
          email: email,
          rol: 'EJECUTIVO',
          hoja: '',
          origen: 'ACCESO AL ARCHIVO'
        });
        Logger.log('   → Será agregado como EJECUTIVO');
      } else {
        Logger.log('   → Ya existe en CONFIG_PERFILES');
      }
    }
    
    // ═══════════════════════════════════════
    // PARTE 2: DETECTAR HOJAS DE EJECUTIVOS
    // ═══════════════════════════════════════
    
    Logger.log('');
    Logger.log('PARTE 2: Detectando hojas de ejecutivos...');
    
    var todasLasHojas = ss.getSheets();
    var hojasExcluidas = [
      'BBDD_REPORTE', 'RESUMEN', 'PRODUCTIVIDAD', 'LLAMADAS',
      'CONFIG_PERFILES', 'Sheet1', 'Hoja 1', 'Hoja1',
      'CONFIGURACION', 'DASHBOARD', 'TOTALES', 'GRAFICOS'
    ];
    
    var hojasEjecutivos = [];
    
    for (var j = 0; j < todasLasHojas.length; j++) {
      var hoja = todasLasHojas[j];
      var nombreHoja = hoja.getName();
      
      // Saltar hojas especiales
      if (hojasExcluidas.indexOf(nombreHoja) !== -1) continue;
      if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) continue;
      
      // Si la hoja tiene datos, probablemente es de un ejecutivo
      if (hoja.getLastRow() > 1) {
        hojasEjecutivos.push(nombreHoja);
        Logger.log((hojasEjecutivos.length) + '. Hoja ejecutivo: ' + nombreHoja);
      }
    }
    
    Logger.log('Total hojas de ejecutivos detectadas: ' + hojasEjecutivos.length);
    
    // Procesar cada hoja de ejecutivo
    Logger.log('');
    Logger.log('Procesando hojas de ejecutivos...');
    
    for (var k = 0; k < hojasEjecutivos.length; k++) {
      var nombreHoja = hojasEjecutivos[k];
      
      // Verificar si ya hay un usuario con esta hoja asignada
      var tieneUsuarioAsignado = verificarHojaAsignada(configSheet, nombreHoja);
      
      if (!tieneUsuarioAsignado) {
        // Crear nombre basado en la hoja
        var nombreEjecutivo = formatearNombreDesdeHoja(nombreHoja);
        
        // Verificar si ya existe un usuario con este nombre
        var yaExisteNombre = existeNombreEnConfigPerfiles(configSheet, nombreEjecutivo);
        
        if (!yaExisteNombre) {
          usuariosNuevos.push({
            nombre: nombreEjecutivo,
            email: '', // Sin email porque solo detectamos por hoja
            rol: 'EJECUTIVO',
            hoja: nombreHoja,
            origen: 'HOJA CREADA'
          });
          Logger.log('   → ' + nombreEjecutivo + ' será agregado (hoja: ' + nombreHoja + ')');
        } else {
          // Usuario existe pero sin hoja asignada, actualizar
          Logger.log('   → ' + nombreEjecutivo + ' existe, actualizando hoja asignada');
          actualizarHojaAsignada(configSheet, nombreEjecutivo, nombreHoja);
        }
      } else {
        Logger.log('   → ' + nombreHoja + ' ya tiene usuario asignado');
      }
    }
    
    // ═══════════════════════════════════════
    // PARTE 3: MOSTRAR RESUMEN Y CONFIRMAR
    // ═══════════════════════════════════════
    
    Logger.log('');
    Logger.log('=== RESUMEN DE SINCRONIZACIÓN ===');
    Logger.log('Total usuarios nuevos a agregar: ' + usuariosNuevos.length);
    
    if (usuariosNuevos.length === 0) {
      ui.alert(
        'ℹ️ Sin Cambios',
        'No se detectaron usuarios nuevos.\n\n' +
        '• Usuarios con acceso: Ya están en CONFIG_PERFILES\n' +
        '• Hojas de ejecutivos: Ya tienen usuarios asignados',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Mostrar resumen
    var mensaje = '🆕 SE DETECTARON ' + usuariosNuevos.length + ' USUARIO(S) NUEVO(S):\n\n';
    
    var desdeAcceso = usuariosNuevos.filter(function(u) { return u.origen === 'ACCESO AL ARCHIVO'; }).length;
    var desdeHojas = usuariosNuevos.filter(function(u) { return u.origen === 'HOJA CREADA'; }).length;
    
    mensaje += '📊 ORIGEN:\n';
    mensaje += '• Desde acceso al archivo: ' + desdeAcceso + '\n';
    mensaje += '• Desde hojas creadas: ' + desdeHojas + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    // Mostrar primeros 10 usuarios
    var limite = Math.min(10, usuariosNuevos.length);
    for (var m = 0; m < limite; m++) {
      var usuario = usuariosNuevos[m];
      mensaje += (m + 1) + '. ' + usuario.nombre + '\n';
      if (usuario.email) {
        mensaje += '   📧 ' + usuario.email + '\n';
      }
      if (usuario.hoja) {
        mensaje += '   📊 Hoja: ' + usuario.hoja + '\n';
      }
      mensaje += '   🏷️ Origen: ' + usuario.origen + '\n\n';
    }
    
    if (usuariosNuevos.length > 10) {
      mensaje += '... y ' + (usuariosNuevos.length - 10) + ' más.\n\n';
    }
    
    mensaje += '¿Deseas agregarlos como EJECUTIVOS?';
    
    var confirmar = ui.alert('🆕 Usuarios Nuevos', mensaje, ui.ButtonSet.YES_NO);
    
    if (confirmar !== ui.Button.YES) return;
    
    // ═══════════════════════════════════════
    // PARTE 4: AGREGAR USUARIOS A CONFIG_PERFILES
    // ═══════════════════════════════════════
    
    Logger.log('');
    Logger.log('Agregando usuarios a CONFIG_PERFILES...');
    
    var ultimaFila = configSheet.getLastRow();
    
    for (var n = 0; n < usuariosNuevos.length; n++) {
      var usuario = usuariosNuevos[n];
      var fila = ultimaFila + n + 1;
      
      configSheet.getRange(fila, 1).setValue(usuario.nombre);
      configSheet.getRange(fila, 2).setValue(usuario.email);
      configSheet.getRange(fila, 3).setValue(usuario.rol);
      configSheet.getRange(fila, 4).setValue(usuario.hoja);
      configSheet.getRange(fila, 5).setValue(ahora);
      configSheet.getRange(fila, 6).setValue(ahora);
      
      var color = (fila % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
      configSheet.getRange(fila, 1, 1, 6).setBackground(color);
      
      Logger.log('✓ Usuario agregado: ' + usuario.nombre);
    }
    
    // ═══════════════════════════════════════
    // PARTE 5: MOSTRAR RESULTADO
    // ═══════════════════════════════════════
    
    Logger.log('');
    Logger.log('✓ Sincronización completada exitosamente');
    
    ui.alert(
      '✅ Sincronización Completada',
      usuariosNuevos.length + ' usuario(s) agregado(s) exitosamente.\n\n' +
      '📊 RESUMEN:\n' +
      '• Desde acceso: ' + desdeAcceso + '\n' +
      '• Desde hojas: ' + desdeHojas + '\n\n' +
      '💡 SIGUIENTE PASO:\n' +
      '• Revisa CONFIG_PERFILES\n' +
      '• Asigna emails a usuarios sin email\n' +
      '• Asigna hojas a usuarios sin hoja',
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

/**
 * ═══════════════════════════════════════
 * FUNCIONES AUXILIARES
 * ═══════════════════════════════════════
 */

/**
 * Verifica si un email ya existe en CONFIG_PERFILES
 */
function existeEnConfigPerfiles(configSheet, email) {
  if (!email) return false;
  
  var ultimaFila = configSheet.getLastRow();
  if (ultimaFila < 2) return false;
  
  var emails = configSheet.getRange(2, 2, ultimaFila - 1, 1).getValues();
  
  for (var i = 0; i < emails.length; i++) {
    if (emails[i][0] && emails[i][0].toString().toLowerCase() === email.toLowerCase()) {
      return true;
    }
  }
  
  return false;
}

/**
 * Verifica si un nombre ya existe en CONFIG_PERFILES
 */
function existeNombreEnConfigPerfiles(configSheet, nombre) {
  if (!nombre) return false;
  
  var ultimaFila = configSheet.getLastRow();
  if (ultimaFila < 2) return false;
  
  var nombres = configSheet.getRange(2, 1, ultimaFila - 1, 1).getValues();
  
  for (var i = 0; i < nombres.length; i++) {
    if (nombres[i][0] && nombres[i][0].toString().toUpperCase() === nombre.toUpperCase()) {
      return true;
    }
  }
  
  return false;
}

/**
 * Verifica si una hoja ya tiene un usuario asignado
 */
function verificarHojaAsignada(configSheet, nombreHoja) {
  var ultimaFila = configSheet.getLastRow();
  if (ultimaFila < 2) return false;
  
  var hojas = configSheet.getRange(2, 4, ultimaFila - 1, 1).getValues();
  
  for (var i = 0; i < hojas.length; i++) {
    if (hojas[i][0] && hojas[i][0].toString() === nombreHoja) {
      return true;
    }
  }
  
  return false;
}

/**
 * Actualiza la hoja asignada de un usuario existente
 */
function actualizarHojaAsignada(configSheet, nombreUsuario, nombreHoja) {
  var ultimaFila = configSheet.getLastRow();
  if (ultimaFila < 2) return;
  
  var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
  
  for (var i = 0; i < datos.length; i++) {
    var nombre = datos[i][0];
    
    if (nombre && nombre.toString().toUpperCase() === nombreUsuario.toUpperCase()) {
      var fila = i + 2;
      configSheet.getRange(fila, 4).setValue(nombreHoja);
      configSheet.getRange(fila, 6).setValue(new Date()); // Actualizar fecha
      Logger.log('✓ Hoja actualizada para ' + nombreUsuario + ': ' + nombreHoja);
      return;
    }
  }
}

/**
 * Formatea el nombre de la hoja para crear un nombre de usuario legible
 * Ejemplo: "ANA_VILLANUEVA" → "Ana Villanueva"
 */
function formatearNombreDesdeHoja(nombreHoja) {
  // Reemplazar guiones bajos por espacios
  var nombre = nombreHoja.replace(/_/g, ' ');
  
  // Capitalizar cada palabra
  nombre = nombre.toLowerCase().replace(/\b\w/g, function(letra) {
    return letra.toUpperCase();
  });
  
  return nombre;
}