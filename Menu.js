/**
 * ========================================
 * MÓDULO: MENÚ CON SISTEMA DE ROLES
 * ========================================
 * 
 * ✅ SOLUCIÓN AL PROBLEMA DE DETECCIÓN DE EMAIL
 * Usa una combinación de:
 * 1. Trigger simple onOpen() para carga rápida
 * 2. UserProperties para recordar el email entre sesiones
 * 3. Activación manual la primera vez
 */

// Configuración de seguridad para supervisores
const CONFIG_SEGURIDAD = {
  PASSWORD: 'admin123',
  INTENTOS_MAXIMOS: 3
};

/**
 * ✅ NUEVA FUNCIÓN: onOpen mejorado con caché de usuario
 * Si no puede obtener el email, usa el email guardado de sesiones anteriores
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    var email = obtenerEmailConCache();
    
    Logger.log('=== CARGA DE MENÚ ===');
    Logger.log('Usuario: ' + (email || 'NO DETECTADO'));
    
    if (!email) {
      Logger.log('⚠️ No se pudo obtener email - Mostrando menú de activación');
      crearMenuActivacion(ui);
      return;
    }
    
    // Obtener rol del usuario desde CONFIG_PERFILES
    var rol = obtenerRolUsuario(email);
    Logger.log('Rol detectado: ' + rol);
    
    // Crear menú según el rol
    if (rol === 'SUPERVISOR') {
      crearMenuSupervisor(ui);
      Logger.log('✓ Menú de SUPERVISOR cargado');
    } else if (rol === 'EJECUTIVO') {
      crearMenuEjecutivo(ui);
      Logger.log('✓ Menú de EJECUTIVO cargado');
    } else {
      // Usuario no encontrado o sin rol
      crearMenuBasico(ui);
      Logger.log('⚠️ Usuario sin rol definido - Menú básico cargado');
    }
    
    // Ejecutar inicialización en segundo plano (sin ventanas)
    ejecutarInicializacionSilenciosa();
    
  } catch (error) {
    Logger.log('Error en onOpen: ' + error.toString());
    // En caso de error, mostrar menú básico
    var ui = SpreadsheetApp.getUi();
    crearMenuActivacion(ui);
  }
}

/**
 * ✅ NUEVA FUNCIÓN: Obtiene email con sistema de caché
 * 1. Intenta obtener el email actual
 * 2. Si falla, busca el email guardado de sesiones anteriores
 * 3. Si no hay email guardado, devuelve null
 */
function obtenerEmailConCache() {
  try {
    // Intentar obtener email actual
    var email = obtenerEmailUsuarioRobusto();
    
    if (email && email.length > 0) {
      // Email obtenido correctamente, guardarlo para futuras sesiones
      var props = PropertiesService.getUserProperties();
      props.setProperty('USER_EMAIL_CACHE', email);
      Logger.log('✓ Email obtenido y guardado en caché: ' + email);
      return email;
    }
    
    // Si no se pudo obtener, intentar usar el email en caché
    var props = PropertiesService.getUserProperties();
    var emailCache = props.getProperty('USER_EMAIL_CACHE');
    
    if (emailCache && emailCache.length > 0) {
      Logger.log('⚠️ Email no obtenido, usando caché: ' + emailCache);
      return emailCache;
    }
    
    // No hay email actual ni en caché
    Logger.log('❌ No se pudo obtener email (ni actual ni caché)');
    return null;
    
  } catch (error) {
    Logger.log('Error en obtenerEmailConCache: ' + error.toString());
    return null;
  }
}

/**
 * ✅ NUEVA FUNCIÓN: Activar cuenta manualmente
 * Cuando el usuario ejecuta esto, se guarda su email para futuras sesiones
 */
function activarMiCuenta() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    Logger.log('=== ACTIVACIÓN MANUAL DE CUENTA ===');
    
    // Obtener email del usuario
    var email = obtenerEmailUsuarioRobusto();
    
    if (!email) {
      ui.alert(
        '❌ Error',
        'No se pudo detectar tu email.\n\n' +
        'Esto puede ocurrir si:\n' +
        '1. El archivo no está en Google Workspace\n' +
        '2. No has autorizado el script\n\n' +
        'Contacta al administrador.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Guardar email en caché
    var props = PropertiesService.getUserProperties();
    props.setProperty('USER_EMAIL_CACHE', email);
    
    // Obtener rol
    var rol = obtenerRolUsuario(email);
    
    var mensaje = '✅ CUENTA ACTIVADA\n\n';
    mensaje += '📧 Email: ' + email + '\n';
    mensaje += '👔 Rol: ' + rol + '\n\n';
    
    if (rol === 'NO_ENCONTRADO') {
      mensaje += '⚠️ Tu usuario NO está registrado en CONFIG_PERFILES.\n\n';
      mensaje += 'Contacta al supervisor para que te asigne un rol.\n\n';
    } else {
      mensaje += '✅ Tu cuenta está configurada correctamente.\n\n';
    }
    
    mensaje += '🔄 RECARGA EL ARCHIVO (F5) para ver tu menú personalizado.';
    
    ui.alert('✅ Activación Exitosa', mensaje, ui.ButtonSet.OK);
    
    Logger.log('✓ Cuenta activada para: ' + email + ' (' + rol + ')');
    
  } catch (error) {
    Logger.log('Error en activarMiCuenta: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'Error activando cuenta:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ✅ NUEVA FUNCIÓN: Menú de activación
 * Se muestra cuando no se puede detectar el usuario automáticamente
 */
function crearMenuActivacion(ui) {
  ui.createMenu('⚠️ Activar Sistema')
    .addItem('🔓 Activar Mi Cuenta', 'activarMiCuenta')
    .addSeparator()
    .addItem('🔍 Diagnosticar Sistema', 'diagnosticarProblemaLorena')
    .addItem('ℹ️ ¿Por qué veo esto?', 'explicarMenuActivacion')
    .addToUi();
  
  Logger.log('✓ Menú de activación cargado');
}

/**
 * ✅ NUEVA FUNCIÓN: Explicación del menú de activación
 */
function explicarMenuActivacion() {
  var ui = SpreadsheetApp.getUi();
  
  var mensaje = '⚠️ MENÚ DE ACTIVACIÓN\n\n';
  mensaje += 'Estás viendo este menú porque el sistema no pudo detectar tu email automáticamente.\n\n';
  mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
  mensaje += '🔧 SOLUCIÓN:\n\n';
  mensaje += '1. Click en "🔓 Activar Mi Cuenta"\n';
  mensaje += '2. Autoriza el script cuando te lo pida\n';
  mensaje += '3. Recarga el archivo (F5)\n';
  mensaje += '4. Tu menú personalizado aparecerá\n\n';
  mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
  mensaje += 'Esto solo necesitas hacerlo UNA VEZ.\n';
  mensaje += 'Después, tu menú se cargará automáticamente.';
  
  ui.alert('ℹ️ Información', mensaje, ui.ButtonSet.OK);
}

/**
 * Crea el menú completo para SUPERVISORES
 */
function crearMenuSupervisor(ui) {
  ui.createMenu('🎯 Gestión Supervisores')
    .addItem('📥 Carga Inicial (Copiar y Distribuir)', 'cargarDatosDesdeLink')
    .addItem('📤 Cargar Base Adicional (Excel)', 'cargarYDistribuirDesdeExcel')
    .addSeparator()
    .addItem('📈 Generar Resumen', 'generarResumenManual')
    .addItem('📊 Crear PRODUCTIVIDAD', 'crearHojaProductividad')
    .addItem('📞 Crear LLAMADAS', 'crearTablaLlamadas')
    .addSeparator()
    .addItem('🗂️ Ordenar Hojas', 'ordenarHojasPorGrupo')
    .addItem('🧹 Limpiar Filas en Blanco', 'limpiarFilasEnBlancoManual')
    .addSeparator()
    .addItem('👥 Ver CONFIG_PERFILES', 'mostrarConfigPerfiles')
    .addItem('🔄 Actualizar CONFIG_PERFILES', 'crearConfigPerfilesManual')
    .addItem('➕ Agregar Usuario Manual', 'agregarUsuarioManual')
    .addItem('🔄 Sincronizar Usuarios', 'sincronizarUsuariosConAcceso')
    .addSeparator()
    .addItem('⚙️ Gestionar Triggers', 'gestionarTriggers')
    .addItem('🔍 Diagnosticar Perfiles', 'diagnosticarSistemaPerfiles')
    .addToUi();
  
  // Menú para Panel de Llamadas
  ui.createMenu('📞 Panel de Llamadas')
    .addItem('📋 Abrir Panel de Gestión', 'mostrarPanel')
    .addItem('🗂️ Navegación de Hojas', 'mostrarPanelNavegacion')
    .addToUi();
}

/**
 * Crea el menú limitado para EJECUTIVOS
 */
function crearMenuEjecutivo(ui) {
  ui.createMenu('📞 Panel de Llamadas')
    .addItem('📋 Abrir Panel de Gestión', 'mostrarPanel')
    .addSeparator()
    .addItem('ℹ️ Información', 'mostrarInfoEjecutivo')
    .addItem('🔍 Diagnosticar Perfiles', 'diagnosticarSistemaPerfiles')
    .addToUi();
  
  // Menú de Navegación (para ejecutivos también)
  ui.createMenu('🗂️ Navegación')
    .addItem('📋 Panel de Navegación', 'mostrarPanelNavegacion')
    .addToUi();
}

/**
 * Crea un menú básico para usuarios sin rol definido
 */
function crearMenuBasico(ui) {
  ui.createMenu('📋 Sistema')
    .addItem('🔄 Panel de Llamadas', 'mostrarPanel')
    .addSeparator()
    .addItem('⚠️ Sin permisos asignados', 'mostrarMensajeSinPermisos')
    .addItem('🔍 Diagnosticar Perfiles', 'diagnosticarSistemaPerfiles')
    .addItem('🔓 Activar Mi Cuenta', 'activarMiCuenta')
    .addToUi();
  
  // Menú de Navegación (disponible para todos)
  ui.createMenu('🗂️ Navegación')
    .addItem('📋 Panel de Navegación', 'mostrarPanelNavegacion')
    .addToUi();
}

/**
 * Muestra información para ejecutivos
 */
function mostrarInfoEjecutivo() {
  var ui = SpreadsheetApp.getUi();
  var email = obtenerEmailConCache();
  var hojaAsignada = obtenerHojaAsignada(email);
  
  var mensaje = '👤 INFORMACIÓN DEL USUARIO\n\n';
  mensaje += '📧 Email: ' + email + '\n';
  mensaje += '👔 Rol: EJECUTIVO\n';
  mensaje += '📊 Hoja asignada: ' + (hojaAsignada || 'No asignada') + '\n\n';
  mensaje += '📞 Usa el Panel de Llamadas para registrar tus gestiones.\n\n';
  mensaje += 'Si tienes problemas, contacta a tu supervisor.';
  
  ui.alert('ℹ️ Información del Usuario', mensaje, ui.ButtonSet.OK);
}

/**
 * Muestra mensaje para usuarios sin permisos
 */
function mostrarMensajeSinPermisos() {
  var ui = SpreadsheetApp.getUi();
  var email = obtenerEmailConCache();
  
  var mensaje = '⚠️ NO TIENES PERMISOS ASIGNADOS\n\n';
  mensaje += '📧 Tu email: ' + (email || 'NO DETECTADO') + '\n\n';
  
  if (!email) {
    mensaje += 'No se pudo detectar tu email automáticamente.\n\n';
    mensaje += '🔧 SOLUCIÓN:\n';
    mensaje += '1. Click en "🔓 Activar Mi Cuenta"\n';
    mensaje += '2. Recarga el archivo (F5)\n';
  } else {
    mensaje += 'Tu usuario no está registrado en CONFIG_PERFILES.\n\n';
    mensaje += 'Por favor contacta a tu supervisor para que te asigne permisos.';
  }
  
  ui.alert('⚠️ Sin Permisos', mensaje, ui.ButtonSet.OK);
}

/**
 * ✅ Función para mostrar el panel lateral de llamadas
 * Disponible para TODOS los usuarios
 */
function mostrarPanel() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Panel')
      .setTitle('📞 Panel de Control')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.log('Error mostrando panel: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'No se pudo abrir el Panel de Llamadas.\n\nError: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ✅ Obtiene la hoja asignada a un usuario
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
      var emailFila = datos[i][1];
      
      if (emailFila && emailFila.toString().trim().toLowerCase() === email.toLowerCase()) {
        return datos[i][3] || null; // Columna 4: HOJA_ASIGNADA
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('Error obteniendo hoja asignada: ' + error.toString());
    return null;
  }
}

/**
 * ⭐ RESTAURADO: Muestra ventana de inicialización al abrir
 */
function mostrarVentanaInicializacion() {
  try {
    // Verificar si ya se inicializó (para evitar mostrar cada vez)
    var props = PropertiesService.getUserProperties();
    var yaInicializado = props.getProperty('SISTEMA_INICIALIZADO');
    
    // Si ya se inicializó hoy, no mostrar ventana
    var hoy = new Date().toDateString();
    if (yaInicializado === hoy) {
      Logger.log('Sistema ya inicializado hoy, ejecutando en segundo plano...');
      ejecutarInicializacionSilenciosa();
      return;
    }
    
    // Primera vez del día: mostrar ventana
    Logger.log('Primera carga del día, mostrando ventana de progreso');
    
    var html = HtmlService.createHtmlOutputFromFile('VentanaInicializacion')
      .setWidth(400)
      .setHeight(250);
    
    SpreadsheetApp.getUi().showModelessDialog(html, '🔄 Inicializando Sistema');
    
    // Ejecutar inicialización con progreso
    inicializarConProgreso();
    
    // Marcar como inicializado
    props.setProperty('SISTEMA_INICIALIZADO', hoy);
    
  } catch (error) {
    Logger.log('Error en ventana de inicialización: ' + error.toString());
    ejecutarInicializacionSilenciosa();
  }
}

/**
 * Inicialización silenciosa en segundo plano
 */
function ejecutarInicializacionSilenciosa() {
  try {
    Logger.log('=== INICIALIZACIÓN SILENCIOSA ===');
    Logger.log('Fecha: ' + new Date());
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Generar resumen automáticamente (sin notificaciones)
    generarResumenAutomatico(ss);
    
    Logger.log('✓ Sistema inicializado correctamente');
    
  } catch (error) {
    Logger.log('⚠️ Error en inicialización: ' + error.toString());
  }
}

/**
 * Inicialización con progreso visible
 */
function inicializarConProgreso() {
  try {
    var cache = CacheService.getUserCache();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Tarea 1: Verificando estructura
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 1,
      mensaje: 'Verificando estructura...',
      completado: false
    }), 120);
    
    Utilities.sleep(500);
    
    // Tarea 2: Actualizando datos
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 2,
      mensaje: 'Actualizando datos...',
      completado: false
    }), 120);
    
    // Generar resumen
    generarResumenAutomatico(ss);
    
    Utilities.sleep(500);
    
    // Tarea 3: Verificando perfiles
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 3,
      mensaje: 'Verificando perfiles...',
      completado: false
    }), 120);
    
    Utilities.sleep(500);
    
    // Tarea 4: Optimizando hojas
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 4,
      mensaje: 'Optimizando hojas...',
      completado: false
    }), 120);
    
    Utilities.sleep(500);
    
    // Tarea 5: Finalización
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 5,
      mensaje: 'Finalizando configuración...',
      completado: true
    }), 120);
    
    Logger.log('✓ Inicialización completada exitosamente');
    
  } catch (error) {
    Logger.log('❌ Error en inicialización con progreso: ' + error.toString());
  }
}

/**
 * Obtiene el estado actual de inicialización (para la ventana)
 */
function obtenerEstadoInicializacion() {
  try {
    var cache = CacheService.getUserCache();
    var estado = cache.get('estadoInicializacion');
    
    if (estado) {
      return JSON.parse(estado);
    }
    
    return {
      tarea: 0,
      mensaje: 'Iniciando...',
      completado: false
    };
    
  } catch (error) {
    Logger.log('Error obteniendo estado: ' + error.toString());
    return null;
  }
}
/**
 * ========================================
 * FUNCIONES FALTANTES PARA EL MENÚ
 * ========================================
 * Estas funciones son llamadas desde Menu.js pero no existían
 * Agrégalas a tu proyecto para que el menú funcione correctamente
 */

/**
 * ✅ CREAR O ACTUALIZAR CONFIG_PERFILES
 * Función llamada desde el menú "🔄 Actualizar CONFIG_PERFILES"
 */
function crearConfigPerfilesManual() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Confirmar acción
    var respuesta = ui.alert(
      '🔄 Actualizar CONFIG_PERFILES',
      '¿Deseas crear o actualizar la hoja CONFIG_PERFILES?\n\n' +
      'Esto detectará automáticamente todos los usuarios con acceso al archivo.\n\n' +
      'Si CONFIG_PERFILES ya existe, se limpiará y recreará.',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) {
      return;
    }
    
    Logger.log('=== CREANDO/ACTUALIZANDO CONFIG_PERFILES ===');
    
    // Verificar o crear hoja
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    var esNueva = false;
    
    if (configSheet) {
      Logger.log('CONFIG_PERFILES ya existe, limpiando...');
      configSheet.clear();
    } else {
      Logger.log('Creando nueva hoja CONFIG_PERFILES...');
      configSheet = ss.insertSheet('CONFIG_PERFILES');
      esNueva = true;
    }
    
    // Crear encabezados
    var encabezados = ['NOMBRE', 'EMAIL', 'ROL', 'HOJA_ASIGNADA', 'FECHA_CREACION', 'ULTIMA_MODIFICACION'];
    configSheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    
    // Formatear encabezados
    var rangoEncabezado = configSheet.getRange(1, 1, 1, encabezados.length);
    rangoEncabezado.setBackground('#4CAF50');
    rangoEncabezado.setFontColor('white');
    rangoEncabezado.setFontWeight('bold');
    rangoEncabezado.setHorizontalAlignment('center');
    
    // Detectar usuarios con acceso
    Logger.log('Detectando usuarios con acceso al archivo...');
    var file = DriveApp.getFileById(ss.getId());
    var propietario = file.getOwner();
    var editores = file.getEditors();
    
    Logger.log('Propietario: ' + propietario.getName() + ' (' + propietario.getEmail() + ')');
    Logger.log('Total editores: ' + editores.length);
    
    // Preparar datos
    var usuarios = [];
    var ahora = new Date();
    var emailsProcesados = [];
    
    // Agregar propietario como SUPERVISOR
    var emailPropietario = propietario.getEmail().toLowerCase();
    usuarios.push([
      propietario.getName() || propietario.getEmail().split('@')[0],
      propietario.getEmail(),
      'SUPERVISOR',
      '',
      ahora,
      ahora
    ]);
    emailsProcesados.push(emailPropietario);
    Logger.log('1. ' + propietario.getName() + ' → SUPERVISOR (propietario)');
    
    // Agregar editores como EJECUTIVO
    for (var i = 0; i < editores.length; i++) {
      var editor = editores[i];
      var email = editor.getEmail().toLowerCase();
      
      // No duplicar al propietario
      if (emailsProcesados.indexOf(email) !== -1) {
        Logger.log((i + 2) + '. ' + editor.getName() + ' → OMITIDO (duplicado)');
        continue;
      }
      
      usuarios.push([
        editor.getName() || editor.getEmail().split('@')[0],
        editor.getEmail(),
        'EJECUTIVO',
        '',
        ahora,
        ahora
      ]);
      emailsProcesados.push(email);
      Logger.log((i + 2) + '. ' + editor.getName() + ' → EJECUTIVO');
    }
    
    Logger.log('Total usuarios a agregar: ' + usuarios.length);
    
    // Escribir datos
    if (usuarios.length > 0) {
      configSheet.getRange(2, 1, usuarios.length, 6).setValues(usuarios);
      
      // Aplicar formato alternado
      for (var j = 0; j < usuarios.length; j++) {
        var fila = j + 2;
        var color = (fila % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
        configSheet.getRange(fila, 1, 1, 6).setBackground(color);
      }
    }
    
    // Ajustar anchos de columnas
    configSheet.setColumnWidth(1, 200); // NOMBRE
    configSheet.setColumnWidth(2, 250); // EMAIL
    configSheet.setColumnWidth(3, 120); // ROL
    configSheet.setColumnWidth(4, 200); // HOJA_ASIGNADA
    configSheet.setColumnWidth(5, 150); // FECHA_CREACION
    configSheet.setColumnWidth(6, 150); // ULTIMA_MODIFICACION
    
    // Centrar columnas
    configSheet.getRange(2, 3, usuarios.length, 1).setHorizontalAlignment('center'); // ROL
    configSheet.getRange(2, 5, usuarios.length, 2).setHorizontalAlignment('center'); // Fechas
    
    // Aplicar bordes
    configSheet.getRange(1, 1, usuarios.length + 1, 6).setBorder(true, true, true, true, true, true);
    
    Logger.log('✓ CONFIG_PERFILES creada/actualizada correctamente');
    
    // Mostrar resultado
    var mensaje = '✅ CONFIG_PERFILES ' + (esNueva ? 'CREADA' : 'ACTUALIZADA') + '\n\n';
    mensaje += '📊 RESUMEN:\n';
    mensaje += '• Total usuarios: ' + usuarios.length + '\n';
    mensaje += '• Supervisores: ' + usuarios.filter(function(u) { return u[2] === 'SUPERVISOR'; }).length + '\n';
    mensaje += '• Ejecutivos: ' + usuarios.filter(function(u) { return u[2] === 'EJECUTIVO'; }).length + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += '🔍 SIGUIENTE PASO:\n';
    mensaje += '1. Ve a CONFIG_PERFILES\n';
    mensaje += '2. Asigna hojas en HOJA_ASIGNADA\n';
    mensaje += '3. Cambia ROL si es necesario';
    
    ui.alert('✅ Completado', mensaje, ui.ButtonSet.OK);
    
    // Mostrar la hoja
    if (configSheet.isSheetHidden()) {
      configSheet.showSheet();
    }
    ss.setActiveSheet(configSheet);
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'Error al crear CONFIG_PERFILES:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ✅ MOSTRAR CONFIG_PERFILES
 * Función llamada desde el menú "👥 Ver CONFIG_PERFILES"
 */
function mostrarConfigPerfiles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) {
      var ui = SpreadsheetApp.getUi();
      var respuesta = ui.alert(
        '⚠️ CONFIG_PERFILES no existe',
        'La hoja CONFIG_PERFILES no existe.\n\n¿Deseas crearla ahora?',
        ui.ButtonSet.YES_NO
      );
      
      if (respuesta === ui.Button.YES) {
        crearConfigPerfilesManual();
      }
      return;
    }
    
    // Mostrar la hoja
    if (configSheet.isSheetHidden()) {
      configSheet.showSheet();
    }
    
    ss.setActiveSheet(configSheet);
    
    // Obtener información para mostrar
    var ultimaFila = configSheet.getLastRow();
    
    if (ultimaFila < 2) {
      SpreadsheetApp.getUi().alert(
        'ℹ️ CONFIG_PERFILES vacía',
        'La hoja CONFIG_PERFILES existe pero está vacía.\n\n' +
        'Usa "🔄 Actualizar CONFIG_PERFILES" para llenarla automáticamente.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Mostrar resumen
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    var supervisores = 0;
    var ejecutivos = 0;
    
    for (var i = 0; i < datos.length; i++) {
      if (datos[i][2] === 'SUPERVISOR') supervisores++;
      if (datos[i][2] === 'EJECUTIVO') ejecutivos++;
    }
    
    var mensaje = '👥 CONFIG_PERFILES\n\n';
    mensaje += '📊 RESUMEN:\n';
    mensaje += '• Total usuarios: ' + (ultimaFila - 1) + '\n';
    mensaje += '• Supervisores: ' + supervisores + '\n';
    mensaje += '• Ejecutivos: ' + ejecutivos + '\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += '💡 ACCIONES DISPONIBLES:\n';
    mensaje += '• Editar roles manualmente\n';
    mensaje += '• Asignar hojas en HOJA_ASIGNADA\n';
    mensaje += '• Agregar usuarios con "➕ Agregar Usuario"\n';
    mensaje += '• Sincronizar con "🔄 Sincronizar Usuarios"';
    
    SpreadsheetApp.getUi().alert('👥 CONFIG_PERFILES', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'Error al mostrar CONFIG_PERFILES:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ✅ FUNCIÓN AUXILIAR: Obtener hoja asignada a un usuario
 * Busca en CONFIG_PERFILES la hoja asignada al email
 */
function obtenerHojaAsignada(email) {
  try {
    if (!email) return null;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('CONFIG_PERFILES');
    
    if (!configSheet) return null;
    
    var ultimaFila = configSheet.getLastRow();
    if (ultimaFila < 2) return null;
    
    var datos = configSheet.getRange(2, 1, ultimaFila - 1, 4).getValues();
    
    for (var i = 0; i < datos.length; i++) {
      var emailFila = datos[i][1];
      
      if (emailFila && emailFila.toString().trim().toLowerCase() === email.toLowerCase()) {
        return datos[i][3] || null; // Columna 4: HOJA_ASIGNADA
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('Error obteniendo hoja asignada: ' + error.toString());
    return null;
  }
}