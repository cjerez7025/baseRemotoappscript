/**
 * ========================================
 * MÓDULO: SISTEMA DE PERFILAMIENTO - MEJORADO
 * ========================================
 * 
 * ✅ CORRECCIÓN: Obtención robusta de email de usuario
 * Funciona tanto para propietarios como para editores
 */

/**
 * ✅ NUEVO: Obtiene el email del usuario de forma robusta
 * Prueba múltiples métodos para usuarios que no son propietarios
 */
function obtenerEmailUsuarioRobusto() {
  try {
    // Método 1: getActiveUser (más confiable si funciona)
    var email = Session.getActiveUser().getEmail();
    if (email && email.length > 0 && email.indexOf('@') !== -1) {
      Logger.log('Email obtenido con getActiveUser: ' + email);
      return email;
    }
    
    // Método 2: getEffectiveUser (fallback para editores)
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
 * ✅ ACTUALIZADO: Obtiene el rol de un usuario desde CONFIG_PERFILES
 * Ahora usa obtenerEmailUsuarioRobusto()
 */
function obtenerRolUsuario(email) {
  try {
    // Si no se proporciona email, obtenerlo del usuario actual
    if (!email) {
      email = obtenerEmailUsuarioRobusto();
      if (!email) {
        Logger.log('No se pudo obtener email del usuario');
        return 'NO_ENCONTRADO';
      }
    }
    
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
      var emailRegistrado = datos[i][1]; // Columna EMAIL
      if (emailRegistrado && emailRegistrado.toString().toLowerCase() === email.toLowerCase()) {
        var rol = datos[i][2] || 'EJECUTIVO'; // Columna ROL
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
 * ✅ ACTUALIZADO: Obtiene la hoja asignada a un ejecutivo
 */
function obtenerHojaAsignada(email) {
  try {
    // Si no se proporciona email, obtenerlo del usuario actual
    if (!email) {
      email = obtenerEmailUsuarioRobusto();
      if (!email) {
        Logger.log('No se pudo obtener email del usuario');
        return null;
      }
    }
    
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
      var emailRegistrado = datos[i][1]; // Columna EMAIL
      if (emailRegistrado && emailRegistrado.toString().toLowerCase() === email.toLowerCase()) {
        return datos[i][3] || null; // Columna HOJA_ASIGNADA
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('Error obteniendo hoja asignada: ' + error.toString());
    return null;
  }
}

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
      
      Logger.log('✓ Hoja CONFIG_PERFILES creada exitosamente');
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
 */
function generarEmailDesdeNombre(nombreCompleto) {
  try {
    var nombreLimpio = nombreCompleto.toString().trim();
    nombreLimpio = nombreLimpio.replace(/_/g, ' ');
    nombreLimpio = nombreLimpio.toLowerCase();
    
    // Remover acentos
    nombreLimpio = nombreLimpio
      .replace(/á/g, 'a')
      .replace(/é/g, 'e')
      .replace(/í/g, 'i')
      .replace(/ó/g, 'o')
      .replace(/ú/g, 'u')
      .replace(/ñ/g, 'n');
    
    var palabras = nombreLimpio.split(/\s+/);
    
    if (palabras.length >= 2) {
      return palabras[0] + '.' + palabras[1] + '@empresa.com';
    } else if (palabras.length === 1) {
      return palabras[0] + '@empresa.com';
    }
    
    return nombreLimpio.replace(/\s+/g, '.') + '@empresa.com';
    
  } catch (error) {
    Logger.log('Error generando email: ' + error.toString());
    return nombreCompleto.toLowerCase().replace(/\s+/g, '.') + '@empresa.com';
  }
}