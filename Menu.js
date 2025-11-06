/**
 * ========================================
 * MÓDULO: MENÚ CON SISTEMA DE ROLES
 * ========================================
 * 
 * Gestiona menús diferenciados según el rol del usuario:
 * - EJECUTIVO: Solo Panel de Llamadas
 * - SUPERVISOR: Menú completo + Panel de Llamadas
 * 
 * Se configura automáticamente al abrir Google Sheets
 * Muestra ventana de progreso inicial en la primera carga
 */

// Configuración de seguridad para supervisores
const CONFIG_SEGURIDAD = {
  PASSWORD: 'admin123',
  INTENTOS_MAXIMOS: 3
};

/**
 * FUNCIÓN PRINCIPAL: Se ejecuta al abrir Google Sheets
 * Detecta el rol del usuario y muestra el menú apropiado
 * 
 * NOTA: La ventana de progreso inicial NO se muestra aquí.
 * Se muestra solo si tienes el trigger activado desde "Gestionar Triggers"
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    var email = obtenerEmailUsuarioRobusto();
    
    Logger.log('=== CARGA DE MENÚ ===');
    Logger.log('Usuario: ' + email);
    
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
    SpreadsheetApp.getUi().createMenu('📋 Sistema')
      .addItem('🔄 Panel de Llamadas', 'mostrarPanel')
      .addToUi();
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
    Logger.log('Primera carga del día, mostrando ventana de inicialización...');
    
    var html = HtmlService.createHtmlOutputFromFile('VentanaCargaInicio')
      .setWidth(450)
      .setHeight(550);
    
    SpreadsheetApp.getUi().showModalDialog(html, '🚀 Inicializando Sistema CRM');
    
    // Ejecutar inicialización en segundo plano
    ejecutarInicializacionConProgreso();
    
    // Marcar como inicializado
    props.setProperty('SISTEMA_INICIALIZADO', hoy);
    
  } catch (error) {
    Logger.log('Error mostrando ventana de inicialización: ' + error.toString());
    // Si falla, ejecutar silenciosamente
    ejecutarInicializacionSilenciosa();
  }
}

/**
 * Ejecuta inicialización CON ventana de progreso
 */
function ejecutarInicializacionConProgreso() {
  try {
    Logger.log('=== INICIALIZACIÓN CON PROGRESO ===');
    
    var cache = CacheService.getUserCache();
    
    // Tarea 1: Generar Resumen
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 1,
      mensaje: 'Generando resumen...',
      completado: false
    }), 120);
    
    generarResumenSeguro();
    Utilities.sleep(1000);
    
    // Tarea 2: Actualizar Llamadas
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 2,
      mensaje: 'Actualizando llamadas...',
      completado: false
    }), 120);
    
    Utilities.sleep(800);
    
    // Tarea 3: Ordenar Hojas
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 3,
      mensaje: 'Ordenando hojas...',
      completado: false
    }), 120);
    
    try {
      ordenarHojasPorGrupo();
    } catch (e) {
      Logger.log('Error ordenando hojas: ' + e.toString());
    }
    
    Utilities.sleep(800);
    
    // Tarea 4: Actualizar Productividad
    cache.put('estadoInicializacion', JSON.stringify({
      tarea: 4,
      mensaje: 'Actualizando productividad...',
      completado: false
    }), 120);
    
    Utilities.sleep(800);
    
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
 * Solo tienen acceso al Panel de Llamadas
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
  var email = Session.getActiveUser().getEmail();
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
  var email = Session.getActiveUser().getEmail();
  
  var mensaje = '⚠️ NO TIENES PERMISOS ASIGNADOS\n\n';
  mensaje += '📧 Tu email: ' + email + '\n\n';
  mensaje += 'Tu usuario no está registrado en el sistema.\n\n';
  mensaje += 'Por favor contacta a tu supervisor para que te asigne permisos.';
  
  ui.alert('⚠️ Sin Permisos', mensaje, ui.ButtonSet.OK);
}

/**
 * Función para mostrar el panel lateral de llamadas
 * Disponible para TODOS los usuarios
 */
function mostrarPanel() {
  var html = HtmlService.createHtmlOutputFromFile('Panel')
    .setTitle('Panel de Control')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Ejecuta inicialización en segundo plano sin ventanas
 * Se ejecuta automáticamente desde onOpen() si ya fue inicializado hoy
 */
function ejecutarInicializacionSilenciosa() {
  try {
    Logger.log('=== INICIALIZACIÓN SILENCIOSA ===');
    Logger.log('Fecha: ' + new Date());
    
    generarResumenSeguro();
    
    Logger.log('✓ Sistema inicializado correctamente');
    
  } catch (error) {
    Logger.log('❌ Error en inicialización: ' + error.toString());
  }
}

/**
 * Genera resumen de forma segura (sin mostrar notificaciones)
 * Se usa en inicialización automática
 */
function generarResumenSeguro() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      Logger.log('⚠️ BBDD_REPORTE no existe. Se omite generación de resumen.');
      return;
    }
    
    if (bddSheet.getLastRow() < 2) {
      Logger.log('⚠️ BBDD_REPORTE está vacía. Se omite generación de resumen.');
      return;
    }
    
    generarResumenAutomatico(spreadsheet);
    Logger.log('✓ Resumen generado correctamente');
    
  } catch (error) {
    Logger.log('❌ Error generando resumen: ' + error.toString());
  }
}

/**
 * Genera resumen manualmente (con confirmación)
 * Solo para SUPERVISORES
 */
function generarResumenManual() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var respuesta = ui.alert(
      '📈 Generar Resumen',
      '¿Deseas generar/actualizar la hoja RESUMEN?',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) {
      return;
    }
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      ui.alert('❌ Error', 'No se encontró la hoja BBDD_REPORTE', ui.ButtonSet.OK);
      return;
    }
    
    generarResumenAutomatico(spreadsheet);
    ui.alert('✅ Completado', 'Resumen generado exitosamente', ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Error al generar resumen:\n\n' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ========================================
 * FUNCIONES DE VALIDACIÓN DE PERMISOS
 * ========================================
 */

/**
 * Verifica si el usuario actual es supervisor
 * @return {boolean} true si es supervisor
 */
function esUsuarioSupervisor() {
  try {
    var email = Session.getActiveUser().getEmail();
    var rol = obtenerRolUsuario(email);
    return rol === 'SUPERVISOR';
  } catch (error) {
    Logger.log('Error verificando supervisor: ' + error.toString());
    return false;
  }
}

/**
 * Verifica si el usuario actual es ejecutivo
 * @return {boolean} true si es ejecutivo
 */
function esUsuarioEjecutivo() {
  try {
    var email = Session.getActiveUser().getEmail();
    var rol = obtenerRolUsuario(email);
    return rol === 'EJECUTIVO';
  } catch (error) {
    Logger.log('Error verificando ejecutivo: ' + error.toString());
    return false;
  }
}

/**
 * Bloquea el acceso si el usuario no es supervisor
 * Muestra mensaje y retorna false
 */
function validarAccesoSupervisor() {
  if (!esUsuarioSupervisor()) {
    SpreadsheetApp.getUi().alert(
      '🚫 Acceso Denegado',
      'Esta función solo está disponible para supervisores.\n\n' +
      'Si necesitas acceso, contacta a tu supervisor.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
  return true;
}

/**
 * ========================================
 * FUNCIONES DE DIAGNÓSTICO
 * ========================================
 */

/**
 * Muestra información del usuario actual
 * Útil para debugging
 */
function mostrarInfoUsuarioActual() {
  try {
    var ui = SpreadsheetApp.getUi();
    var email = Session.getActiveUser().getEmail();
    var rol = obtenerRolUsuario(email);
    var hoja = obtenerHojaAsignada(email);
    
    var mensaje = '🔍 INFORMACIÓN DEL USUARIO ACTUAL\n\n';
    mensaje += '📧 Email: ' + email + '\n';
    mensaje += '👔 Rol: ' + rol + '\n';
    mensaje += '📊 Hoja asignada: ' + (hoja || 'Ninguna') + '\n';
    mensaje += '✅ Es Supervisor: ' + (esUsuarioSupervisor() ? 'Sí' : 'No') + '\n';
    mensaje += '👤 Es Ejecutivo: ' + (esUsuarioEjecutivo() ? 'Sí' : 'No');
    
    ui.alert('🔍 Diagnóstico de Usuario', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Error al obtener información:\n\n' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Función de prueba para verificar el sistema de roles
 * Solo para desarrollo
 */
function probarSistemaRoles() {
  try {
    Logger.log('=== PRUEBA DE SISTEMA DE ROLES ===');
    
    var email = Session.getActiveUser().getEmail();
    Logger.log('Email actual: ' + email);
    
    var rol = obtenerRolUsuario(email);
    Logger.log('Rol obtenido: ' + rol);
    
    var hoja = obtenerHojaAsignada(email);
    Logger.log('Hoja asignada: ' + (hoja || 'Ninguna'));
    
    Logger.log('Es Supervisor: ' + esUsuarioSupervisor());
    Logger.log('Es Ejecutivo: ' + esUsuarioEjecutivo());
    
    Logger.log('=== PRUEBA COMPLETADA ===');
    
  } catch (error) {
    Logger.log('ERROR EN PRUEBA: ' + error.toString());
  }
}