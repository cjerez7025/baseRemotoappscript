/**
 * ARCHIVO: Menu.gs
 * MENÚ PRINCIPAL DEL SISTEMA CON PROTECCIÓN POR CONTRASEÑA
 */

// CONFIGURACIÓN DE SEGURIDAD
const CONFIG_SEGURIDAD = {
  PASSWORD: 'Admin2025',
  INTENTOS_MAXIMOS: 3,
  MENSAJE_ACCESO_DENEGADO: '🔒 Acceso denegado. Contraseña incorrecta.'
};

/**
 * Función que se ejecuta al abrir la hoja
 * NOTA: onOpen() tiene restricciones de seguridad, no puede mostrar diálogos
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Menú de Gestión de Supervisores (con contraseña)
  ui.createMenu('Gestión de Supervisores')
    .addItem('🔐 Acceder al Panel de Supervisores', 'solicitarAccesoGestion')
    .addToUi();
  
  // Menú de Panel de Llamadas (sin contraseña)
  ui.createMenu('📞 Panel de Llamadas')
    .addItem('Abrir Panel', 'mostrarPanel')
    .addToUi();
  
  // Menú de Navegación (sin contraseña - para todos)
  ui.createMenu('🗂️ Navegación')
    .addItem('📋 Abrir Panel de Navegación', 'mostrarPanelNavegacion')
    .addSeparator()
    .addItem('🔍 Diagnóstico de Hojas', 'diagnosticarHojas')
    .addItem('📊 Ordenar Hojas', 'ordenarHojasAutomaticamente')
    .addToUi();
  
  // NUEVO: Menú para inicialización manual
  ui.createMenu('⚙️ Sistema')
    .addItem('🚀 Inicializar Sistema (con ventana)', 'inicializarSistemaConVentana')
    .addSeparator()
    .addItem('🔧 Instalar Trigger Automático', 'instalarTriggerOnOpen')
    .addItem('🗑️ Desinstalar Trigger Automático', 'desinstalarTriggerOnOpen')
    .addToUi();
  
  // DESHABILITADO: No ejecutar inicialización automática en onOpen
  // Causa conflictos con hojas que se están creando/eliminando
  // Los usuarios deben usar el trigger instalable o inicializar manualmente
  
  Logger.log('✓ Menús cargados. Sistema listo.');
}

/**
 * Ejecuta inicialización en segundo plano sin ventanas
 * Esta se ejecuta automáticamente desde onOpen()
 */
function ejecutarInicializacionSilenciosa() {
  try {
    Logger.log('=== INICIALIZACIÓN SILENCIOSA ===');
    Logger.log('Fecha: ' + new Date());
    
    generarResumenSeguro();
    crearTablaLlamadas();
    ordenarHojasPorGrupo();
    crearHojaProductividad();
    
    Logger.log('✓ Sistema inicializado correctamente');
    
  } catch (error) {
    Logger.log('❌ Error en inicialización: ' + error.toString());
  }
}

/**
 * NUEVA FUNCIÓN: Inicializa el sistema CON ventana de progreso
 * Esta función SÍ puede mostrar ventanas porque es activada por el usuario
 */
function inicializarSistemaConVentana() {
  try {
    // Resetear estado
    guardarEstadoInicializacion({ tarea: 0, mensaje: 'Iniciando...', completado: false });
    
    // Mostrar ventana de carga
    const html = HtmlService.createHtmlOutputFromFile('VentanaCargaInicio')
      .setWidth(450)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModelessDialog(html, 'Inicializando Sistema');
    
    // Pequeño delay para que se muestre la ventana
    SpreadsheetApp.flush();
    Utilities.sleep(300);
    
    // TAREA 1: Generar Resumen
    guardarEstadoInicializacion({ tarea: 1, mensaje: 'Generando resumen...', completado: false });
    generarResumenSeguro();
    Utilities.sleep(500);
    
    // TAREA 2: Crear Tabla de Llamadas
    guardarEstadoInicializacion({ tarea: 2, mensaje: 'Creando tabla de llamadas...', completado: false });
    crearTablaLlamadas();
    Utilities.sleep(500);
    
    // TAREA 3: Ordenar Hojas
    guardarEstadoInicializacion({ tarea: 3, mensaje: 'Ordenando hojas...', completado: false });
    ordenarHojasPorGrupo();
    Utilities.sleep(500);
    
    // TAREA 4: Crear Hoja Productividad
    guardarEstadoInicializacion({ tarea: 4, mensaje: 'Creando hoja de productividad...', completado: false });
    crearHojaProductividad();
    Utilities.sleep(500);
    
    // TAREA 5: Finalizar
    guardarEstadoInicializacion({ tarea: 5, mensaje: 'Finalizando configuración...', completado: false });
    Utilities.sleep(500);
    
    // COMPLETADO
    guardarEstadoInicializacion({ tarea: 5, mensaje: '✅ Sistema listo', completado: true });
    
    Logger.log('✓ Sistema inicializado con ventana');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    guardarEstadoInicializacion({ tarea: 0, mensaje: 'Error: ' + error.message, completado: true });
    SpreadsheetApp.getUi().alert('Error', 'Hubo un problema en la inicialización: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Guarda el estado actual de la inicialización en Properties
 */
function guardarEstadoInicializacion(estado) {
  try {
    const props = PropertiesService.getUserProperties();
    props.setProperty('estadoInicializacion', JSON.stringify(estado));
  } catch (error) {
    Logger.log('Error guardando estado: ' + error.toString());
  }
}

/**
 * Obtiene el estado actual de la inicialización
 * Esta función es llamada desde el HTML para actualizar la UI
 */
function obtenerEstadoInicializacion() {
  try {
    const props = PropertiesService.getUserProperties();
    const estadoStr = props.getProperty('estadoInicializacion');
    
    if (estadoStr) {
      return JSON.parse(estadoStr);
    }
    
    return { tarea: 0, mensaje: 'Iniciando...', completado: false };
    
  } catch (error) {
    Logger.log('Error obteniendo estado: ' + error.toString());
    return { tarea: 0, mensaje: 'Iniciando...', completado: false };
  }
}

/**
 * INSTALAR TRIGGER: Esta función instala un trigger que se ejecuta al abrir
 * y SÍ puede mostrar ventanas
 */
function instalarTriggerOnOpen() {
  try {
    // Primero eliminar triggers existentes para evitar duplicados
    desinstalarTriggerOnOpen();
    
    // Crear nuevo trigger
    ScriptApp.newTrigger('inicializarSistemaConVentana')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
    
    SpreadsheetApp.getUi().alert(
      '✅ Trigger Instalado',
      'Ahora el sistema se inicializará automáticamente con ventana de progreso cada vez que abras la hoja.\n\n' +
      'Para desactivarlo, usa: Sistema → Desinstalar Trigger Automático',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log('✓ Trigger instalado correctamente');
    
  } catch (error) {
    Logger.log('❌ Error instalando trigger: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'No se pudo instalar el trigger: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * DESINSTALAR TRIGGER: Elimina el trigger automático
 */
function desinstalarTriggerOnOpen() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let eliminados = 0;
    
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'inicializarSistemaConVentana') {
        ScriptApp.deleteTrigger(triggers[i]);
        eliminados++;
      }
    }
    
    if (eliminados > 0) {
      SpreadsheetApp.getUi().alert(
        '✅ Trigger Desinstalado',
        'Se eliminaron ' + eliminados + ' trigger(s). Ya no se mostrará la ventana automáticamente al abrir.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      Logger.log('✓ ' + eliminados + ' trigger(s) eliminados');
    } else {
      SpreadsheetApp.getUi().alert(
        'ℹ️ Sin Cambios',
        'No había triggers instalados.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log('❌ Error desinstalando trigger: ' + error.toString());
  }
}

/**
 * Ejecuta las funciones de inicialización en segundo plano
 * CORREGIDO: Más tiempo de espera y validación de hojas
 */
function ejecutarInicializacionSilenciosa() {
  try {
    Logger.log('=== INICIALIZACIÓN SILENCIOSA ===');
    Logger.log('Fecha: ' + new Date());
    
    // Esperar más tiempo para que el spreadsheet esté completamente cargado
    Utilities.sleep(2000);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // PASO 1: Generar Resumen (si existe BBDD_REPORTE)
    try {
      Logger.log('1. Generando resumen...');
      generarResumenSeguro();
      SpreadsheetApp.flush(); // Forzar actualización
      Utilities.sleep(1000);
      Logger.log('✓ Resumen completado');
    } catch (e) {
      Logger.log('❌ Error en resumen: ' + e.toString());
    }
    
    // PASO 2: Crear Tabla Llamadas
    try {
      Logger.log('2. Creando tabla llamadas...');
      crearTablaLlamadasSegura();
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      Logger.log('✓ Llamadas completada');
    } catch (e) {
      Logger.log('❌ Error en llamadas: ' + e.toString());
    }
    
    // PASO 3: Crear Hoja Productividad
    try {
      Logger.log('3. Creando productividad...');
      crearHojaProductividadSegura();
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      Logger.log('✓ Productividad completada');
    } catch (e) {
      Logger.log('❌ Error en productividad: ' + e.toString());
    }
    
    // PASO 4: Ordenar Hojas (al final)
    try {
      Logger.log('4. Ordenando hojas...');
      ordenarHojasPorGrupo();
      SpreadsheetApp.flush();
      Logger.log('✓ Orden completado');
    } catch (e) {
      Logger.log('❌ Error ordenando: ' + e.toString());
    }
    
    Logger.log('✅ Inicialización completada');
    
  } catch (error) {
    Logger.log('❌ Error crítico: ' + error.toString());
  }
}

/**
 * Versión segura de crearTablaLlamadas
 */
function crearTablaLlamadasSegura() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = ss.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet || bddSheet.getLastRow() < 2) {
      Logger.log('⚠️ BBDD_REPORTE no disponible o vacía');
      return;
    }
    
    crearTablaLlamadas();
    
  } catch (error) {
    Logger.log('Error en crearTablaLlamadasSegura: ' + error.toString());
  }
}

/**
 * Versión segura de crearHojaProductividad
 */
function crearHojaProductividadSegura() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bddSheet = ss.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet || bddSheet.getLastRow() < 2) {
      Logger.log('⚠️ BBDD_REPORTE no disponible o vacía');
      return;
    }
    
    crearHojaProductividad();
    
  } catch (error) {
    Logger.log('Error en crearHojaProductividadSegura: ' + error.toString());
  }
}

/**
 * Genera el resumen de forma segura (sin showModelessDialog)
 * Esta versión NO muestra ventanas emergentes durante onOpen
 */
function generarResumenSeguro() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const bddSheet = spreadsheet.getSheetByName('BBDD_REPORTE');
    
    if (!bddSheet) {
      Logger.log('⚠️ BBDD_REPORTE no existe aún. Se omite generación de resumen.');
      return;
    }
    
    // VALIDAR que la hoja tenga datos antes de procesar
    if (bddSheet.getLastRow() < 2) {
      Logger.log('⚠️ BBDD_REPORTE está vacía. Se omite generación de resumen.');
      return;
    }
    
    // Llamar a la función de resumen pero sin mostrar notificaciones visuales
    generarResumenAutomatico(spreadsheet);
    Logger.log('✓ Resumen generado correctamente');
    
  } catch (error) {
    Logger.log('❌ Error generando resumen: ' + error.toString());
    // No lanzar el error para no interrumpir la inicialización
  }
}

/**
 * Solicita contraseña antes de mostrar el menú de gestión
 */
function solicitarAccesoGestion() {
  const ui = SpreadsheetApp.getUi();
  let intentos = 0;
  
  while (intentos < CONFIG_SEGURIDAD.INTENTOS_MAXIMOS) {
    const response = ui.prompt(
      '🔐 Acceso Restringido',
      'Ingresa la contraseña para acceder a Gestión de Supervisores:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('❌ Acceso cancelado');
      return;
    }
    
    const passwordIngresado = response.getResponseText();
    
    if (passwordIngresado === CONFIG_SEGURIDAD.PASSWORD) {
      ui.alert('✅ Acceso concedido', 'Bienvenido al panel de Gestión de Supervisores', ui.ButtonSet.OK);
      mostrarMenuGestion();
      return;
    }
    
    intentos++;
    const intentosRestantes = CONFIG_SEGURIDAD.INTENTOS_MAXIMOS - intentos;
    
    if (intentosRestantes > 0) {
      ui.alert(
        '❌ Contraseña incorrecta',
        'Te quedan ' + intentosRestantes + ' intento(s)',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '🚫 Acceso Bloqueado',
        'Has excedido el número máximo de intentos. Contacta al administrador.',
        ui.ButtonSet.OK
      );
      registrarIntentoFallido();
    }
  }
}

/**
 * Muestra el menú completo de gestión después de autenticación exitosa
 */
function mostrarMenuGestion() {
  const ui = SpreadsheetApp.getUi();
  
  const resultado = ui.alert(
    '🚀 Panel de Gestión Supervisores',
    '¿Qué deseas hacer?\n\n' +
    '1️⃣ Carga Inicial (Copiar y Distribuir)\n' +
    '2️⃣ Generar Resumen\n' +
    '3️⃣ Funciones Individuales\n' +
    '4️⃣ Limpiar Hojas de Ejecutivos\n' +
    '5️⃣ Cargar Base Adicional (Excel)\n\n' +
    'Selecciona una opción:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resultado === ui.Button.OK) {
    mostrarOpcionesGestion();
  }
}

/**
 * Muestra las opciones del menú de gestión
 */
function mostrarOpcionesGestion() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '📋 Selecciona una opción',
    'Ingresa el número de la opción:\n\n' +
    '1 - Carga Inicial (Copiar y Distribuir)\n' +
    '2 - Generar Resumen\n' +
    '3 - Funciones Individuales\n' +
    '4 - Limpiar Hojas de Ejecutivos\n' +
    '5 - Cargar Base Adicional (Excel)\n' +
    '0 - Salir',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const opcion = response.getResponseText().trim();
  
  switch(opcion) {
    case '1':
      cargarDatosDesdeLink();
      break;
      
    case '2':
      generateSummary(); // Aquí sí se puede usar porque es acción del usuario
      ui.alert('✅ Resumen generado', 'El resumen ha sido actualizado', ui.ButtonSet.OK);
      break;
      
    case '3':
      mostrarFuncionesIndividuales();
      break;
      
    case '4':
      const confirmar = ui.alert(
        '⚠️ Confirmar acción',
        '¿Estás seguro de que deseas limpiar las hojas de ejecutivos?\nEsta acción no se puede deshacer.',
        ui.ButtonSet.YES_NO
      );
      if (confirmar === ui.Button.YES) {
        limpiarHojasEjecutivos();
        ui.alert('✅ Hojas limpiadas', 'Las hojas de ejecutivos han sido limpiadas', ui.ButtonSet.OK);
      }
      break;
      
    case '5':
      cargarYDistribuirDesdeExcel();
      break;
      
    case '0':
      return;
      
    default:
      ui.alert('❌ Opción inválida', 'Por favor selecciona un número válido', ui.ButtonSet.OK);
      mostrarOpcionesGestion();
  }
}

/**
 * Muestra el submenú de funciones individuales
 */
function mostrarFuncionesIndividuales() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '🔧 Funciones Individuales',
    'Ingresa el número de la función:\n\n' +
    '1 - Actualizar Validaciones\n' +
    '2 - Aplicar Fórmulas ESTADO_COMPROMISO\n' +
    '3 - Verificar Hojas de Ejecutivos\n' +
    '4 - Crear Hoja BBDD_REPORTE\n' +
    '5 - Actualizar Reporte\n' +
    '6 - Aplicar Protección a TODAS las Hojas\n' +
    '7 - Verificar Protección (Hoja Actual)\n' +
    '8 - Aplicar Protección (Solo Hoja Actual)\n' +
    '9 - Eliminar Protecciones (Hoja Actual)\n' +
    '10 - Ordenar Hojas\n' +
    '11 - Regenerar Hoja PRODUCTIVIDAD\n' +
    '12 - Regenerar Hoja LLAMADAS\n' +
    '0 - Volver al menú anterior',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const opcion = response.getResponseText().trim();
  
  switch(opcion) {
    case '1':
      actualizarValidaciones();
      ui.alert('✅ Validaciones actualizadas', 'Las validaciones han sido aplicadas', ui.ButtonSet.OK);
      break;
      
    case '2':
      aplicarFormulasEstadoCompromiso();
      ui.alert('✅ Fórmulas aplicadas', 'Las fórmulas ESTADO_COMPROMISO han sido aplicadas', ui.ButtonSet.OK);
      break;
      
    case '3':
      verificarHojasEjecutivos();
      ui.alert('✅ Verificación completa', 'Revisa el registro de ejecución (Logs)', ui.ButtonSet.OK);
      break;
      
    case '4':
      crearHojaReporte();
      ui.alert('✅ Hoja creada', 'BBDD_REPORTE ha sido creada', ui.ButtonSet.OK);
      break;
      
    case '5':
      actualizarReporte();
      ui.alert('✅ Reporte actualizado', 'BBDD_REPORTE ha sido actualizado', ui.ButtonSet.OK);
      break;
      
    case '6':
      aplicarProteccionTodasHojas();
      ui.alert('✅ Protección aplicada', 'Todas las hojas han sido protegidas', ui.ButtonSet.OK);
      break;
      
    case '7':
      verificarProteccion();
      break;
      
    case '8':
      aplicarProteccionHojaActual();
      ui.alert('✅ Protección aplicada', 'La hoja actual ha sido protegida', ui.ButtonSet.OK);
      break;
      
    case '9':
      eliminarProteccionesHojaActual();
      ui.alert('✅ Protecciones eliminadas', 'Las protecciones de la hoja actual han sido eliminadas', ui.ButtonSet.OK);
      break;
      
    case '10':
      ordenarHojasAutomaticamente2024();
      ui.alert('✅ Hojas ordenadas', 'Las hojas han sido ordenadas correctamente', ui.ButtonSet.OK);
      break;
      
    case '11':
      crearHojaProductividad();
      ui.alert('✅ PRODUCTIVIDAD regenerada', 'La hoja PRODUCTIVIDAD ha sido regenerada', ui.ButtonSet.OK);
      break;
      
    case '12':
      crearTablaLlamadas();
      ui.alert('✅ LLAMADAS regenerada', 'La hoja LLAMADAS ha sido regenerada', ui.ButtonSet.OK);
      break;
      
    case '0':
      mostrarOpcionesGestion();
      return;
      
    default:
      ui.alert('❌ Opción inválida', 'Por favor selecciona un número válido', ui.ButtonSet.OK);
      mostrarFuncionesIndividuales();
  }
}

/**
 * Registra intento fallido de acceso
 */
function registrarIntentoFallido() {
  try {
    const email = Session.getActiveUser().getEmail();
    const fecha = new Date();
    Logger.log('Intento fallido de acceso - Usuario: ' + email + ' - Fecha: ' + fecha);
  } catch (error) {
    Logger.log('Error registrando intento fallido: ' + error.toString());
  }
}

/**
 * Cambiar contraseña (requiere contraseña actual)
 */
function cambiarContrasena() {
  const ui = SpreadsheetApp.getUi();
  
  const responseActual = ui.prompt(
    '🔐 Contraseña Actual',
    'Ingresa la contraseña actual:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (responseActual.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  if (responseActual.getResponseText() !== CONFIG_SEGURIDAD.PASSWORD) {
    ui.alert('❌ Error', 'Contraseña actual incorrecta', ui.ButtonSet.OK);
    return;
  }
  
  const responseNueva = ui.prompt(
    '🔐 Nueva Contraseña',
    'Ingresa la nueva contraseña:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (responseNueva.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const nuevaPassword = responseNueva.getResponseText();
  
  if (nuevaPassword.length < 6) {
    ui.alert('❌ Error', 'La contraseña debe tener al menos 6 caracteres', ui.ButtonSet.OK);
    return;
  }
  
  const responseConfirmar = ui.prompt(
    '🔐 Confirmar Contraseña',
    'Confirma la nueva contraseña:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (responseConfirmar.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  if (responseConfirmar.getResponseText() !== nuevaPassword) {
    ui.alert('❌ Error', 'Las contraseñas no coinciden', ui.ButtonSet.OK);
    return;
  }
  
  ui.alert(
    '⚠️ Atención',
    'Para cambiar la contraseña permanentemente, debes modificar la constante CONFIG_SEGURIDAD.PASSWORD en el código.\n\n' +
    'Nueva contraseña sugerida: ' + nuevaPassword + '\n\n' +
    'Ve a Extensiones > Apps Script > Menu.gs',
    ui.ButtonSet.OK
  );
}

/**
 * Función para mostrar el panel lateral de llamadas (SIN PROTECCIÓN)
 */
function mostrarPanel() {
  var html = HtmlService.createHtmlOutputFromFile('Panel')
    .setTitle('Panel de Control')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}