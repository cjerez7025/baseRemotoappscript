/**
 * ========================================
 * MÓDULO: GESTIÓN DE TRIGGERS
 * ========================================
 * 
 * Permite activar/desactivar triggers automáticos desde el menú
 * 
 * TRIGGERS DISPONIBLES:
 * - Ventana Inicial: Muestra progreso al abrir el archivo
 * - onEdit: Actualiza ESTADO_COMPROMISO al editar FECHA_COMPROMISO
 */

/**
 * Muestra el panel de gestión de triggers
 */
function gestionarTriggers() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Verificar triggers actuales
    var triggers = ScriptApp.getUserTriggers(ss);
    var triggersActivos = [];
    var tieneVentanaInicial = false;
    var tieneOnEdit = false;
    
    for (var i = 0; i < triggers.length; i++) {
      var tipo = triggers[i].getEventType().toString();
      var funcion = triggers[i].getHandlerFunction();
      triggersActivos.push(funcion + ' (' + tipo + ')');
      
      if (funcion === 'mostrarVentanaInicialAutomatica') {
        tieneVentanaInicial = true;
      }
      if (funcion === 'onEdit') {
        tieneOnEdit = true;
      }
    }
    
    var mensaje = '⚙️ GESTIÓN DE TRIGGERS AUTOMÁTICOS\n\n';
    
    if (triggersActivos.length > 0) {
      mensaje += '✅ TRIGGERS ACTIVOS:\n\n';
      for (var j = 0; j < triggersActivos.length; j++) {
        mensaje += '• ' + triggersActivos[j] + '\n';
      }
      mensaje += '\n━━━━━━━━━━━━━━━━━━━━━━\n\n';
    } else {
      mensaje += '⚠️ NO HAY TRIGGERS ACTIVOS\n\n';
      mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    }
    
    mensaje += '¿Qué deseas hacer?\n\n';
    mensaje += '1️⃣ = ' + (tieneVentanaInicial ? '❌ Desactivar' : '✅ Activar') + ' ventana inicial\n';
    mensaje += '2️⃣ = ' + (tieneOnEdit ? '❌ Desactivar' : '✅ Activar') + ' onEdit (compromisos)\n';
    mensaje += '3️⃣ = Instalar todos los triggers\n';
    mensaje += '4️⃣ = Eliminar todos los triggers\n';
    mensaje += '5️⃣ = Ver información de triggers\n';
    mensaje += '0️⃣ = Cancelar';
    
    var respuesta = ui.prompt(
      '⚙️ Triggers',
      mensaje,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (respuesta.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    var opcion = respuesta.getResponseText().trim();
    
    switch(opcion) {
      case '1':
        if (tieneVentanaInicial) {
          desactivarVentanaInicial();
        } else {
          activarVentanaInicial();
        }
        break;
      case '2':
        if (tieneOnEdit) {
          desactivarOnEdit();
        } else {
          activarOnEdit();
        }
        break;
      case '3':
        instalarTriggers();
        break;
      case '4':
        eliminarTodosLosTriggers();
        break;
      case '5':
        mostrarInformacionTriggers();
        break;
      case '0':
        return;
      default:
        ui.alert('❌ Opción inválida', 'Por favor selecciona una opción válida', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Error en gestionarTriggers: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Activa el trigger de ventana inicial
 */
function activarVentanaInicial() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Eliminar trigger existente si existe
    var triggers = ScriptApp.getUserTriggers(ss);
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'mostrarVentanaInicialAutomatica') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Crear nuevo trigger
    ScriptApp.newTrigger('mostrarVentanaInicialAutomatica')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
    
    Logger.log('✓ Trigger ventana inicial activado');
    
    ui.alert(
      '✅ Ventana Inicial Activada',
      'La ventana de progreso inicial se mostrará automáticamente cada vez que abras el archivo.\n\n' +
      '📋 Tareas que ejecuta:\n' +
      '• Generar Resumen\n' +
      '• Ordenar Hojas\n' +
      '• Actualizar Sistema\n\n' +
      '💡 Puedes desactivarla desde este mismo menú.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error activando ventana inicial: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Desactiva el trigger de ventana inicial
 */
function desactivarVentanaInicial() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var triggers = ScriptApp.getUserTriggers(ss);
    var eliminado = false;
    
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'mostrarVentanaInicialAutomatica') {
        ScriptApp.deleteTrigger(triggers[i]);
        eliminado = true;
        Logger.log('✓ Trigger ventana inicial desactivado');
      }
    }
    
    if (eliminado) {
      ui.alert(
        '✅ Ventana Inicial Desactivada',
        'La ventana de progreso inicial ya NO se mostrará al abrir el archivo.\n\n' +
        'El sistema seguirá ejecutando actualizaciones en segundo plano de forma silenciosa.\n\n' +
        '💡 Puedes reactivarla cuando quieras desde este menú.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'ℹ️ Información',
        'La ventana inicial ya estaba desactivada.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log('Error desactivando ventana inicial: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Activa el trigger onEdit
 */
function activarOnEdit() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Eliminar trigger existente si existe
    var triggers = ScriptApp.getUserTriggers(ss);
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Crear nuevo trigger
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    
    Logger.log('✓ Trigger onEdit activado');
    
    ui.alert(
      '✅ Trigger onEdit Activado',
      'El trigger onEdit está activo.\n\n' +
      '📝 Función:\n' +
      'Cuando edites FECHA_COMPROMISO, automáticamente actualizará ESTADO_COMPROMISO.\n\n' +
      'Estados posibles:\n' +
      '• SIN_COMPROMISO\n' +
      '• LLAMAR_HOY\n' +
      '• COMPROMISO_VENCIDO\n' +
      '• COMPROMISO_FUTURO',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error activando onEdit: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Desactiva el trigger onEdit
 */
function desactivarOnEdit() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var triggers = ScriptApp.getUserTriggers(ss);
    var eliminado = false;
    
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(triggers[i]);
        eliminado = true;
        Logger.log('✓ Trigger onEdit desactivado');
      }
    }
    
    if (eliminado) {
      ui.alert(
        '✅ Trigger onEdit Desactivado',
        'El trigger onEdit está desactivado.\n\n' +
        'ESTADO_COMPROMISO ya NO se actualizará automáticamente al editar FECHA_COMPROMISO.\n\n' +
        'Tendrás que actualizar los estados manualmente.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'ℹ️ Información',
        'El trigger onEdit ya estaba desactivado.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log('Error desactivando onEdit: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Instala todos los triggers necesarios
 */
function instalarTriggers() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Primero eliminar triggers existentes para evitar duplicados
    var triggersExistentes = ScriptApp.getUserTriggers(ss);
    for (var i = 0; i < triggersExistentes.length; i++) {
      ScriptApp.deleteTrigger(triggersExistentes[i]);
    }
    
    Logger.log('Triggers anteriores eliminados');
    
    var instalados = [];
    
    // 1. Trigger ventana inicial (onOpen)
    try {
      ScriptApp.newTrigger('mostrarVentanaInicialAutomatica')
        .forSpreadsheet(ss)
        .onOpen()
        .create();
      instalados.push('✅ Ventana Inicial - Se muestra al abrir');
      Logger.log('✓ Trigger ventana inicial instalado');
    } catch (e) {
      instalados.push('❌ Ventana Inicial - Error: ' + e.message);
      Logger.log('Error instalando ventana inicial: ' + e.toString());
    }
    
    // 2. Trigger onEdit
    try {
      ScriptApp.newTrigger('onEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      instalados.push('✅ onEdit - Actualiza estado de compromisos');
      Logger.log('✓ Trigger onEdit instalado');
    } catch (e) {
      instalados.push('❌ onEdit - Error: ' + e.message);
      Logger.log('Error instalando onEdit: ' + e.toString());
    }
    
    var mensaje = '✅ INSTALACIÓN DE TRIGGERS COMPLETADA\n\n';
    mensaje += 'Resultados:\n\n';
    for (var j = 0; j < instalados.length; j++) {
      mensaje += instalados[j] + '\n';
    }
    mensaje += '\n━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += '🔄 Los triggers están activos ahora.\n\n';
    mensaje += '📋 Funciones:\n';
    mensaje += '• Ventana inicial al abrir el archivo\n';
    mensaje += '• Actualización automática de compromisos';
    
    ui.alert('✅ Triggers Instalados', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error instalando triggers: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ⭐ FUNCIÓN QUE SE EJECUTA AUTOMÁTICAMENTE CON EL TRIGGER
 * Muestra la ventana de inicialización al abrir el archivo
 */
function mostrarVentanaInicialAutomatica() {
  try {
    Logger.log('=== VENTANA INICIAL AUTOMÁTICA (desde trigger) ===');
    
    // Llamar a la función que ya existe en Menu.js
    mostrarVentanaInicializacion();
    
  } catch (error) {
    Logger.log('Error mostrando ventana inicial: ' + error.toString());
  }
}

/**
 * Elimina todos los triggers del proyecto
 */
function eliminarTodosLosTriggers() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var confirmar = ui.alert(
      '⚠️ Confirmar Eliminación',
      '¿Estás seguro de eliminar TODOS los triggers?\n\n' +
      'Esto desactivará las actualizaciones automáticas.',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmar !== ui.Button.YES) {
      return;
    }
    
    var triggers = ScriptApp.getUserTriggers(ss);
    var eliminados = 0;
    
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
      eliminados++;
      Logger.log('Trigger eliminado: ' + triggers[i].getHandlerFunction());
    }
    
    var mensaje = '✅ TRIGGERS ELIMINADOS\n\n';
    mensaje += 'Total eliminados: ' + eliminados + '\n\n';
    mensaje += 'Las funciones automáticas están desactivadas.\n';
    mensaje += 'Puedes reinstalarlos cuando quieras.';
    
    ui.alert('✅ Completado', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error eliminando triggers: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Muestra información detallada sobre los triggers
 */
function mostrarInformacionTriggers() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var triggers = ScriptApp.getUserTriggers(ss);
    
    var mensaje = '📋 INFORMACIÓN DE TRIGGERS\n\n';
    mensaje += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    if (triggers.length === 0) {
      mensaje += '⚠️ NO HAY TRIGGERS INSTALADOS\n\n';
      mensaje += 'Los triggers automáticos están desactivados.\n\n';
      mensaje += '💡 Para instalarlos:\n';
      mensaje += 'Usa la opción "3️⃣ Instalar todos los triggers"';
    } else {
      mensaje += '✅ TRIGGERS ACTIVOS: ' + triggers.length + '\n\n';
      
      for (var i = 0; i < triggers.length; i++) {
        var trigger = triggers[i];
        var funcion = trigger.getHandlerFunction();
        var tipo = trigger.getEventType().toString();
        
        mensaje += '─────────────────────\n';
        mensaje += '📌 Trigger ' + (i + 1) + ':\n\n';
        mensaje += '  Función: ' + funcion + '\n';
        mensaje += '  Tipo: ' + tipo + '\n';
        
        // Descripción según el tipo
        if (funcion === 'mostrarVentanaInicialAutomatica') {
          mensaje += '\n  📝 Descripción:\n';
          mensaje += '  Muestra ventana de progreso\n';
          mensaje += '  al abrir el archivo\n';
        } else if (funcion === 'onEdit') {
          mensaje += '\n  📝 Descripción:\n';
          mensaje += '  Actualiza automáticamente el\n';
          mensaje += '  ESTADO_COMPROMISO cuando se\n';
          mensaje += '  modifica FECHA_COMPROMISO\n';
        }
        
        mensaje += '\n';
      }
      
      mensaje += '─────────────────────\n\n';
      mensaje += '🔄 Estado: FUNCIONANDO\n';
      mensaje += '✅ Los triggers se ejecutan automáticamente';
    }
    
    mensaje += '\n\n━━━━━━━━━━━━━━━━━━━━━━\n\n';
    mensaje += '💡 FUNCIONES DE LOS TRIGGERS:\n\n';
    mensaje += '• Ventana Inicial:\n';
    mensaje += '  Muestra progreso al abrir\n';
    mensaje += '  Ejecuta tareas de inicialización\n\n';
    mensaje += '• onEdit:\n';
    mensaje += '  Cuando editas FECHA_COMPROMISO,\n';
    mensaje += '  actualiza automáticamente\n';
    mensaje += '  ESTADO_COMPROMISO';
    
    ui.alert('📋 Información de Triggers', mensaje, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error mostrando información: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Trigger onEdit: Actualiza ESTADO_COMPROMISO automáticamente
 * Se ejecuta cuando se edita cualquier celda
 */
function onEdit(e) {
  try {
    var hoja = e.source.getActiveSheet();
    var nombreHoja = hoja.getName();
    
    // Saltar hojas del sistema
    if (/^BBDD_.*_REMOTO/i.test(nombreHoja)) return;
    if (nombreHoja === 'BBDD_REPORTE' || nombreHoja === 'RESUMEN' || 
        nombreHoja === 'LLAMADAS' || nombreHoja === 'PRODUCTIVIDAD' ||
        nombreHoja === 'CONFIG_PERFILES') return;
    
    // Obtener encabezados
    var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var idxFechaCompromiso = enc.indexOf('FECHA_COMPROMISO');
    var idxEstadoCompromiso = enc.indexOf('ESTADO_COMPROMISO');
    
    // Si no existen las columnas, salir
    if (idxFechaCompromiso === -1 || idxEstadoCompromiso === -1) return;
    
    // Si se editó la columna FECHA_COMPROMISO
    if (e.range.getColumn() === idxFechaCompromiso + 1) {
      var fila = e.range.getRow();
      var col = columnNumberToLetter(idxFechaCompromiso + 1);
      
      var celdaEstado = hoja.getRange(fila, idxEstadoCompromiso + 1);
      
      celdaEstado.clearContent();
      celdaEstado.clearFormat();
      
      celdaEstado.setNumberFormat('@STRING@');
      SpreadsheetApp.flush();
      celdaEstado.setNumberFormat('General');
      
      // FORMULA HIBRIDA: IF (inglés) + ; (separador español)
      var f = '=IF(ISBLANK(' + col + fila + ');"SIN_COMPROMISO";IF(' + col + fila + '=TODAY();"LLAMAR_HOY";IF(' + col + fila + '<TODAY();"COMPROMISO_VENCIDO";"COMPROMISO_FUTURO")))';
      
      celdaEstado.setFormula(f);
      SpreadsheetApp.flush();
      
      Logger.log('Estado de compromiso actualizado en fila ' + fila + ' de ' + nombreHoja);
    }
    
  } catch (error) {
    Logger.log('Error en onEdit trigger: ' + error.toString());
  }
}

/**
 * Función auxiliar para convertir número de columna a letra
 */
function columnNumberToLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}