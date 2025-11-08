/**
 * ========================================
 * FUNCIÓN PARA CONSOLA: CREAR CONFIG_PERFILES
 * ========================================
 * Ejecuta esta función desde la consola de Apps Script
 * para crear/actualizar la hoja CONFIG_PERFILES automáticamente
 * 
 * USO:
 * 1. Abre Apps Script (Extensiones > Apps Script)
 * 2. Pega este código en un nuevo archivo
 * 3. En la consola, ejecuta: crearConfigPerfilesDesdeConsola()
 */

/**
 * ✅ FUNCIÓN PRINCIPAL - Ejecutar desde consola
 * Crea CONFIG_PERFILES con todos los usuarios que tienen acceso al archivo
 */
function crearConfigPerfilesDesdeConsola() {
  try {
    Logger.log('╔═══════════════════════════════════════════╗');
    Logger.log('║   CREANDO CONFIG_PERFILES DESDE CONSOLA  ║');
    Logger.log('╚═══════════════════════════════════════════╝');
    Logger.log('');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // PASO 1: Crear o limpiar la hoja CONFIG_PERFILES
    Logger.log('PASO 1: Verificando hoja CONFIG_PERFILES...');
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (configSheet) {
      Logger.log('⚠️  CONFIG_PERFILES ya existe. Limpiando contenido...');
      configSheet.clear();
    } else {
      Logger.log('✓ Creando nueva hoja CONFIG_PERFILES...');
      configSheet = ss.insertSheet('CONFIG_PERFILES');
    }
    
    // PASO 2: Crear encabezados
    Logger.log('');
    Logger.log('PASO 2: Creando estructura...');
    
    var encabezados = ['NOMBRE', 'EMAIL', 'ROL', 'HOJA_ASIGNADA', 'FECHA_CREACION', 'ULTIMA_MODIFICACION'];
    configSheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    
    // Formatear encabezados
    var rangoEncabezado = configSheet.getRange(1, 1, 1, encabezados.length);
    rangoEncabezado.setBackground('#4CAF50');
    rangoEncabezado.setFontColor('white');
    rangoEncabezado.setFontWeight('bold');
    rangoEncabezado.setHorizontalAlignment('center');
    
    Logger.log('✓ Encabezados creados');
    
    // PASO 3: Detectar usuarios con acceso al archivo
    Logger.log('');
    Logger.log('PASO 3: Detectando usuarios con acceso...');
    
    var file = DriveApp.getFileById(ss.getId());
    var propietario = file.getOwner();
    var editores = file.getEditors();
    
    Logger.log('📁 Archivo: ' + file.getName());
    Logger.log('👑 Propietario: ' + propietario.getName() + ' (' + propietario.getEmail() + ')');
    Logger.log('✏️  Editores detectados: ' + editores.length);
    Logger.log('');
    
    // PASO 4: Agregar usuarios a CONFIG_PERFILES
    Logger.log('PASO 4: Agregando usuarios...');
    Logger.log('─────────────────────────────────────────');
    
    var usuarios = [];
    var ahora = new Date();
    
    // Agregar propietario como SUPERVISOR
    usuarios.push({
      nombre: propietario.getName() || propietario.getEmail().split('@')[0],
      email: propietario.getEmail(),
      rol: 'SUPERVISOR',
      hoja: '',
      fechaCreacion: ahora,
      ultimaMod: ahora
    });
    
    Logger.log('1. ' + propietario.getName() + ' (' + propietario.getEmail() + ') → SUPERVISOR (propietario)');
    
    // Agregar editores como EJECUTIVO
    for (var i = 0; i < editores.length; i++) {
      var editor = editores[i];
      var email = editor.getEmail();
      
      // No duplicar al propietario
      if (email.toLowerCase() === propietario.getEmail().toLowerCase()) {
        Logger.log((i + 2) + '. ' + editor.getName() + ' (' + email + ') → OMITIDO (ya es propietario)');
        continue;
      }
      
      usuarios.push({
        nombre: editor.getName() || email.split('@')[0],
        email: email,
        rol: 'EJECUTIVO',
        hoja: '',
        fechaCreacion: ahora,
        ultimaMod: ahora
      });
      
      Logger.log((i + 2) + '. ' + editor.getName() + ' (' + email + ') → EJECUTIVO');
    }
    
    Logger.log('─────────────────────────────────────────');
    Logger.log('Total usuarios a agregar: ' + usuarios.length);
    Logger.log('');
    
    // PASO 5: Escribir datos en la hoja
    Logger.log('PASO 5: Escribiendo datos en CONFIG_PERFILES...');
    
    if (usuarios.length > 0) {
      var datos = usuarios.map(function(u) {
        return [u.nombre, u.email, u.rol, u.hoja, u.fechaCreacion, u.ultimaMod];
      });
      
      configSheet.getRange(2, 1, datos.length, 6).setValues(datos);
      
      // Aplicar formato alternado
      for (var j = 0; j < datos.length; j++) {
        var fila = j + 2;
        var color = (fila % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
        configSheet.getRange(fila, 1, 1, 6).setBackground(color);
      }
      
      Logger.log('✓ ' + datos.length + ' usuarios escritos correctamente');
    }
    
    // PASO 6: Ajustar columnas y formato
    Logger.log('');
    Logger.log('PASO 6: Aplicando formato...');
    
    configSheet.setColumnWidth(1, 200); // NOMBRE
    configSheet.setColumnWidth(2, 250); // EMAIL
    configSheet.setColumnWidth(3, 120); // ROL
    configSheet.setColumnWidth(4, 200); // HOJA_ASIGNADA
    configSheet.setColumnWidth(5, 150); // FECHA_CREACION
    configSheet.setColumnWidth(6, 150); // ULTIMA_MODIFICACION
    
    // Centrar columnas ROL y fechas
    configSheet.getRange(2, 3, usuarios.length, 1).setHorizontalAlignment('center'); // ROL
    configSheet.getRange(2, 5, usuarios.length, 2).setHorizontalAlignment('center'); // Fechas
    
    // Aplicar bordes
    configSheet.getRange(1, 1, usuarios.length + 1, 6).setBorder(true, true, true, true, true, true);
    
    Logger.log('✓ Formato aplicado');
    
    // PASO 7: Proteger la hoja (opcional - descomenta si quieres protección)
    /*
    Logger.log('');
    Logger.log('PASO 7: Protegiendo hoja...');
    var protection = configSheet.protect().setDescription('CONFIG_PERFILES - Solo supervisores');
    protection.setWarningOnly(true);
    Logger.log('✓ Hoja protegida');
    */
    
    // PASO 8: Resumen final
    Logger.log('');
    Logger.log('╔═══════════════════════════════════════════╗');
    Logger.log('║            ✅ PROCESO COMPLETADO          ║');
    Logger.log('╚═══════════════════════════════════════════╝');
    Logger.log('');
    Logger.log('📊 RESUMEN:');
    Logger.log('   • Hoja creada: CONFIG_PERFILES');
    Logger.log('   • Total usuarios: ' + usuarios.length);
    Logger.log('   • Supervisores: ' + usuarios.filter(function(u) { return u.rol === 'SUPERVISOR'; }).length);
    Logger.log('   • Ejecutivos: ' + usuarios.filter(function(u) { return u.rol === 'EJECUTIVO'; }).length);
    Logger.log('');
    Logger.log('🔍 SIGUIENTE PASO:');
    Logger.log('   1. Ve a la hoja CONFIG_PERFILES');
    Logger.log('   2. Asigna las hojas correspondientes en la columna HOJA_ASIGNADA');
    Logger.log('   3. Cambia ROL a SUPERVISOR si es necesario');
    Logger.log('');
    
    // Mostrar la hoja creada
    ss.setActiveSheet(configSheet);
    
    return '✅ CONFIG_PERFILES creada exitosamente con ' + usuarios.length + ' usuarios';
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌ ERROR: ' + error.toString());
    Logger.log('');
    Logger.log('Stack trace:');
    Logger.log(error.stack);
    throw error;
  }
}

/**
 * ✅ VERSIÓN ALTERNATIVA - Con roles personalizados
 * Permite especificar qué emails serán SUPERVISORES
 * 
 * USO:
 * crearConfigPerfilesConSupervisores([
 *   'supervisor1@gmail.com',
 *   'supervisor2@gmail.com'
 * ]);
 */
function crearConfigPerfilesConSupervisores(emailsSupervisores) {
  try {
    Logger.log('╔═══════════════════════════════════════════╗');
    Logger.log('║   CREANDO CONFIG_PERFILES CON ROLES       ║');
    Logger.log('╚═══════════════════════════════════════════╝');
    Logger.log('');
    
    // Validar parámetro
    if (!emailsSupervisores || !Array.isArray(emailsSupervisores)) {
      emailsSupervisores = [];
    }
    
    // Convertir a minúsculas para comparación
    var supervisoresLowercase = emailsSupervisores.map(function(e) {
      return e.toLowerCase().trim();
    });
    
    Logger.log('📋 Emails que serán SUPERVISORES:');
    supervisoresLowercase.forEach(function(e, i) {
      Logger.log('   ' + (i + 1) + '. ' + e);
    });
    Logger.log('');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Crear o limpiar hoja
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (configSheet) {
      Logger.log('⚠️  CONFIG_PERFILES ya existe. Limpiando...');
      configSheet.clear();
    } else {
      Logger.log('✓ Creando nueva hoja CONFIG_PERFILES...');
      configSheet = ss.insertSheet('CONFIG_PERFILES');
    }
    
    // Crear encabezados
    var encabezados = ['NOMBRE', 'EMAIL', 'ROL', 'HOJA_ASIGNADA', 'FECHA_CREACION', 'ULTIMA_MODIFICACION'];
    configSheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    
    var rangoEncabezado = configSheet.getRange(1, 1, 1, encabezados.length);
    rangoEncabezado.setBackground('#4CAF50');
    rangoEncabezado.setFontColor('white');
    rangoEncabezado.setFontWeight('bold');
    rangoEncabezado.setHorizontalAlignment('center');
    
    // Detectar usuarios
    Logger.log('Detectando usuarios con acceso...');
    var file = DriveApp.getFileById(ss.getId());
    var propietario = file.getOwner();
    var editores = file.getEditors();
    
    var usuarios = [];
    var ahora = new Date();
    var emailsProcesados = [];
    
    // Función auxiliar para determinar rol
    function determinarRol(email) {
      return supervisoresLowercase.indexOf(email.toLowerCase()) !== -1 ? 'SUPERVISOR' : 'EJECUTIVO';
    }
    
    // Agregar propietario
    var emailPropietario = propietario.getEmail().toLowerCase();
    usuarios.push({
      nombre: propietario.getName() || propietario.getEmail().split('@')[0],
      email: propietario.getEmail(),
      rol: determinarRol(emailPropietario) || 'SUPERVISOR', // Propietario siempre supervisor
      hoja: '',
      fechaCreacion: ahora,
      ultimaMod: ahora
    });
    emailsProcesados.push(emailPropietario);
    
    Logger.log('1. ' + propietario.getName() + ' (' + propietario.getEmail() + ') → ' + usuarios[0].rol + ' (propietario)');
    
    // Agregar editores
    for (var i = 0; i < editores.length; i++) {
      var editor = editores[i];
      var email = editor.getEmail().toLowerCase();
      
      // No duplicar
      if (emailsProcesados.indexOf(email) !== -1) {
        Logger.log((i + 2) + '. ' + editor.getName() + ' (' + email + ') → OMITIDO (duplicado)');
        continue;
      }
      
      var rol = determinarRol(email);
      
      usuarios.push({
        nombre: editor.getName() || editor.getEmail().split('@')[0],
        email: editor.getEmail(),
        rol: rol,
        hoja: '',
        fechaCreacion: ahora,
        ultimaMod: ahora
      });
      
      emailsProcesados.push(email);
      
      Logger.log((i + 2) + '. ' + editor.getName() + ' (' + editor.getEmail() + ') → ' + rol);
    }
    
    Logger.log('');
    Logger.log('Total usuarios: ' + usuarios.length);
    
    // Escribir datos
    if (usuarios.length > 0) {
      var datos = usuarios.map(function(u) {
        return [u.nombre, u.email, u.rol, u.hoja, u.fechaCreacion, u.ultimaMod];
      });
      
      configSheet.getRange(2, 1, datos.length, 6).setValues(datos);
      
      // Formato alternado
      for (var j = 0; j < datos.length; j++) {
        var fila = j + 2;
        var color = (fila % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
        configSheet.getRange(fila, 1, 1, 6).setBackground(color);
      }
    }
    
    // Ajustar columnas
    configSheet.setColumnWidth(1, 200);
    configSheet.setColumnWidth(2, 250);
    configSheet.setColumnWidth(3, 120);
    configSheet.setColumnWidth(4, 200);
    configSheet.setColumnWidth(5, 150);
    configSheet.setColumnWidth(6, 150);
    
    configSheet.getRange(2, 3, usuarios.length, 1).setHorizontalAlignment('center');
    configSheet.getRange(2, 5, usuarios.length, 2).setHorizontalAlignment('center');
    configSheet.getRange(1, 1, usuarios.length + 1, 6).setBorder(true, true, true, true, true, true);
    
    // Resumen
    Logger.log('');
    Logger.log('╔═══════════════════════════════════════════╗');
    Logger.log('║            ✅ PROCESO COMPLETADO          ║');
    Logger.log('╚═══════════════════════════════════════════╝');
    Logger.log('');
    Logger.log('📊 RESUMEN:');
    Logger.log('   • Total usuarios: ' + usuarios.length);
    Logger.log('   • Supervisores: ' + usuarios.filter(function(u) { return u.rol === 'SUPERVISOR'; }).length);
    Logger.log('   • Ejecutivos: ' + usuarios.filter(function(u) { return u.rol === 'EJECUTIVO'; }).length);
    Logger.log('');
    
    ss.setActiveSheet(configSheet);
    
    return '✅ CONFIG_PERFILES creada con ' + usuarios.length + ' usuarios';
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌ ERROR: ' + error.toString());
    throw error;
  }
}

/**
 * ✅ FUNCIÓN RÁPIDA - Solo listar usuarios sin crear hoja
 * Útil para ver qué usuarios se detectarían antes de crear CONFIG_PERFILES
 */
function listarUsuariosDelArchivo() {
  try {
    Logger.log('╔═══════════════════════════════════════════╗');
    Logger.log('║        USUARIOS CON ACCESO AL ARCHIVO     ║');
    Logger.log('╚═══════════════════════════════════════════╝');
    Logger.log('');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    
    var propietario = file.getOwner();
    var editores = file.getEditors();
    var visualizadores = file.getViewers();
    
    Logger.log('📁 Archivo: ' + file.getName());
    Logger.log('');
    
    Logger.log('👑 PROPIETARIO:');
    Logger.log('   ' + propietario.getName() + ' (' + propietario.getEmail() + ')');
    Logger.log('');
    
    if (editores.length > 0) {
      Logger.log('✏️  EDITORES (' + editores.length + '):');
      for (var i = 0; i < editores.length; i++) {
        Logger.log('   ' + (i + 1) + '. ' + editores[i].getName() + ' (' + editores[i].getEmail() + ')');
      }
      Logger.log('');
    }
    
    if (visualizadores.length > 0) {
      Logger.log('👁️  VISUALIZADORES (' + visualizadores.length + '):');
      for (var j = 0; j < visualizadores.length; j++) {
        Logger.log('   ' + (j + 1) + '. ' + visualizadores[j].getName() + ' (' + visualizadores[j].getEmail() + ')');
      }
      Logger.log('');
    }
    
    Logger.log('═══════════════════════════════════════════');
    Logger.log('Total con acceso de edición: ' + (1 + editores.length));
    Logger.log('');
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    throw error;
  }
}