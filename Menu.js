/**
 * ARCHIVO: Menu.gs
 * MENÚ PRINCIPAL DEL SISTEMA CON PROTECCIÓN POR CONTRASEÑA
 */

// CONFIGURACIÓN DE SEGURIDAD
const CONFIG_SEGURIDAD = {
  PASSWORD: 'Admin2025', // Cambia esta contraseña por la que desees
  INTENTOS_MAXIMOS: 3,
  MENSAJE_ACCESO_DENEGADO: '🔒 Acceso denegado. Contraseña incorrecta.'
};

/**
 * Función que se ejecuta al abrir la hoja
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Menú principal: Gestión de Ejecutivos (PROTEGIDO)
  ui.createMenu('Gestión de Supervisores')
    .addItem('🔐 Acceder al Panel de Supervisores', 'solicitarAccesoGestion')
    .addToUi();
  
  // Menú del Panel de Llamadas (SIN PROTECCIÓN)
  ui.createMenu('📞 Panel de Llamadas')
    .addItem('Abrir Panel', 'mostrarPanel')
    .addToUi();
  
  // Generar resumen automáticamente
  generateSummary();
  crearTablaLlamadas();
  ordenarHojasPorGrupo();
  crearHojaProductividad();
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
    
    // Si el usuario cancela
    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('❌ Acceso cancelado');
      return;
    }
    
    const passwordIngresado = response.getResponseText();
    
    // Verificar contraseña
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
        `Te quedan ${intentosRestantes} intento(s)`,
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
    '1️⃣ Distribución inicial\n' +
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
    '1 - Distribución inicial\n' +
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
      procesarEjecutivos();
      ui.alert('✅ Procesamiento completado', 'Los ejecutivos han sido procesados exitosamente', ui.ButtonSet.OK);
      break;
      
    case '2':
      generateSummary();
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
      ui.alert('✅ Fórmulas aplicadas', 'Las fórmulas de ESTADO_COMPROMISO han sido aplicadas', ui.ButtonSet.OK);
      break;
      
    case '3':
      verificarHojasEjecutivos();
      ui.alert('✅ Verificación completada', 'Las hojas de ejecutivos han sido verificadas', ui.ButtonSet.OK);
      break;
      
    case '4':
      crearHojaReporte();
      ui.alert('✅ Hoja creada', 'La hoja BBDD_REPORTE ha sido creada', ui.ButtonSet.OK);
      break;
      
    case '5':
      actualizarReporte();
      ui.alert('✅ Reporte actualizado', 'El reporte ha sido actualizado', ui.ButtonSet.OK);
      break;
      
    case '6':
      var resultado = aplicarProteccionTodasLasHojas(SpreadsheetApp.getActiveSpreadsheet());
      ui.alert('✅ Protección aplicada', 
               'Hojas protegidas: ' + resultado.protegidas + '\n' +
               'Hojas saltadas: ' + resultado.saltadas + '\n' +
               'Errores: ' + resultado.errores, 
               ui.ButtonSet.OK);
      break;
      
    case '7':
      verificarProteccion();
      break;
      
    case '8':
      ejecutarProteccionHojaActual();
      break;
      
    case '9':
      eliminarTodasLasProtecciones();
      break;
      
    case '10':
      ordenarHojasPorGrupo();
      ui.alert('✅ Hojas ordenadas', 'Las hojas han sido ordenadas correctamente', ui.ButtonSet.OK);
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
 * Registra intentos fallidos de acceso (opcional)
 */
function registrarIntentoFallido() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usuario = Session.getActiveUser().getEmail();
    const fecha = new Date();
    
    // Intenta registrar en una hoja de logs si existe
    let logSheet = ss.getSheetByName('LOGS_ACCESO');
    if (!logSheet) {
      logSheet = ss.insertSheet('LOGS_ACCESO');
      logSheet.getRange('A1:C1').setValues([['Fecha', 'Usuario', 'Evento']]);
      logSheet.getRange('A1:C1').setFontWeight('bold');
    }
    
    logSheet.appendRow([fecha, usuario, 'Intento fallido de acceso a Gestión de Supervisores']);
    
  } catch (e) {
    console.log('Error al registrar intento fallido: ' + e.toString());
  }
}

/**
 * Función para cambiar la contraseña (solo para administradores)
 */
function cambiarPassword() {
  const ui = SpreadsheetApp.getUi();
  
  // Solicitar contraseña actual
  const responseActual = ui.prompt(
    '🔐 Cambiar Contraseña',
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
  
  // Solicitar nueva contraseña
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
  
  // Confirmar nueva contraseña
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