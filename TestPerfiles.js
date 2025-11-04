/**
 * MÓDULO DE PRUEBA: SISTEMA DE PERFILAMIENTO
 * Valida que todas las funciones de perfilamiento funcionen correctamente
 */

/**
 * Ejecuta todas las pruebas del sistema de perfilamiento
 */
function ejecutarPruebasPerfilamiento() {
  Logger.log('╔════════════════════════════════════════════════╗');
  Logger.log('║   PRUEBAS DEL SISTEMA DE PERFILAMIENTO         ║');
  Logger.log('╚════════════════════════════════════════════════╝');
  Logger.log('');
  
  var resultados = {
    total: 0,
    exitosas: 0,
    fallidas: 0,
    pruebas: []
  };
  
  // Prueba 1: Crear CONFIG_PERFILES
  ejecutarPrueba(resultados, 'Crear CONFIG_PERFILES', function() {
    var hoja = crearHojaConfigPerfiles();
    if (!hoja) throw new Error('No se pudo crear la hoja');
    return '✓ Hoja creada correctamente';
  });
  
  // Prueba 2: Validar estructura
  ejecutarPrueba(resultados, 'Validar estructura de CONFIG_PERFILES', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(NOMBRE_HOJA_PERFILES);
    validarEstructuraPerfiles(hoja);
    return '✓ Estructura válida';
  });
  
  // Prueba 3: Agregar datos de prueba
  ejecutarPrueba(resultados, 'Agregar datos de prueba', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(NOMBRE_HOJA_PERFILES);
    
    var datosPrueba = [
      ['supervisor.test@empresa.cl', 'Supervisor Test', 'Supervisor', true],
      ['ejecutivo.test@empresa.cl', 'Ejecutivo Test', 'Ejecutivo', true],
      ['inactivo.test@empresa.cl', 'Usuario Inactivo', 'Ejecutivo', false]
    ];
    
    hoja.getRange(2, 1, datosPrueba.length, 4).setValues(datosPrueba);
    return '✓ ' + datosPrueba.length + ' usuarios de prueba agregados';
  });
  
  // Prueba 4: Obtener perfiles
  ejecutarPrueba(resultados, 'Obtener perfiles configurados', function() {
    var perfiles = obtenerPerfilesConfigurados();
    var keys = Object.keys(perfiles);
    
    // Contar únicos por email
    var unicos = {};
    for (var i = 0; i < keys.length; i++) {
      if (perfiles[keys[i]].email) {
        unicos[perfiles[keys[i]].email] = true;
      }
    }
    
    var cantidad = Object.keys(unicos).length;
    if (cantidad < 2) throw new Error('Se esperaban al menos 2 usuarios activos');
    
    return '✓ ' + cantidad + ' perfiles activos encontrados';
  });
  
  // Prueba 5: Buscar perfil por email
  ejecutarPrueba(resultados, 'Buscar perfil por email', function() {
    var perfil = buscarPerfilEjecutivo('supervisor.test@empresa.cl');
    if (perfil !== 'Supervisor') throw new Error('Perfil incorrecto: ' + perfil);
    return '✓ Supervisor encontrado correctamente';
  });
  
  // Prueba 6: Buscar perfil por nombre
  ejecutarPrueba(resultados, 'Buscar perfil por nombre', function() {
    var perfil = buscarPerfilEjecutivo('Ejecutivo Test');
    if (perfil !== 'Ejecutivo') throw new Error('Perfil incorrecto: ' + perfil);
    return '✓ Ejecutivo encontrado correctamente';
  });
  
  // Prueba 7: Normalización de nombres
  ejecutarPrueba(resultados, 'Normalización de nombres', function() {
    var casos = [
      { input: 'Juan Pérez', esperado: 'juan_perez' },
      { input: 'María José González', esperado: 'maria_jose_gonzalez' },
      { input: 'PEDRO SILVA', esperado: 'pedro_silva' },
      { input: '  Carlos   Ramos  ', esperado: 'carlos_ramos' }
    ];
    
    for (var i = 0; i < casos.length; i++) {
      var resultado = normalizarNombreEjecutivo(casos[i].input);
      if (resultado !== casos[i].esperado) {
        throw new Error('Error en normalización: "' + casos[i].input + 
                       '" -> "' + resultado + '" (esperado: "' + casos[i].esperado + '")');
      }
    }
    
    return '✓ ' + casos.length + ' casos de normalización correctos';
  });
  
  // Prueba 8: Usuario inactivo
  ejecutarPrueba(resultados, 'Usuario inactivo no debe aparecer', function() {
    var perfil = buscarPerfilEjecutivo('inactivo.test@empresa.cl');
    if (perfil !== 'Sin Perfil') throw new Error('Usuario inactivo debería retornar "Sin Perfil"');
    return '✓ Usuario inactivo manejado correctamente';
  });
  
  // Prueba 9: Usuario no existente
  ejecutarPrueba(resultados, 'Usuario no existente', function() {
    var perfil = buscarPerfilEjecutivo('noexiste@empresa.cl');
    if (perfil !== 'Sin Perfil') throw new Error('Usuario inexistente debería retornar "Sin Perfil"');
    return '✓ Usuario no existente manejado correctamente';
  });
  
  // Prueba 10: Crear hoja de prueba con perfil
  ejecutarPrueba(resultados, 'Crear hoja de prueba con perfil', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var nombreHoja = 'TEST_EJECUTIVO_PRUEBA';
    
    // Eliminar hoja si existe
    var hojaExistente = ss.getSheetByName(nombreHoja);
    if (hojaExistente) ss.deleteSheet(hojaExistente);
    
    // Crear hoja de prueba
    var hoja = ss.insertSheet(nombreHoja);
    
    // Agregar columna PERFIL
    var resultado = agregarColumnaPerfilAHoja(hoja, 'Ejecutivo Test');
    
    if (resultado === -1) throw new Error('No se pudo agregar columna PERFIL');
    
    // Verificar que tenga encabezado PERFIL
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var tienePerfil = false;
    
    for (var i = 0; i < encabezados.length; i++) {
      if (encabezados[i].toString().toUpperCase() === 'PERFIL') {
        tienePerfil = true;
        break;
      }
    }
    
    if (!tienePerfil) throw new Error('No se encontró columna PERFIL');
    
    return '✓ Hoja de prueba creada con columna PERFIL';
  });
  
  // Resumen final
  Logger.log('');
  Logger.log('╔════════════════════════════════════════════════╗');
  Logger.log('║              RESUMEN DE PRUEBAS                ║');
  Logger.log('╠════════════════════════════════════════════════╣');
  Logger.log('║  Total:     ' + pad(resultados.total, 3) + '                            ║');
  Logger.log('║  Exitosas:  ' + pad(resultados.exitosas, 3) + ' ✓                         ║');
  Logger.log('║  Fallidas:  ' + pad(resultados.fallidas, 3) + ' ✗                         ║');
  Logger.log('╚════════════════════════════════════════════════╝');
  
  if (resultados.fallidas === 0) {
    Logger.log('');
    Logger.log('🎉 ¡TODAS LAS PRUEBAS PASARON EXITOSAMENTE!');
    Logger.log('');
    Logger.log('El sistema de perfilamiento está listo para usar.');
  } else {
    Logger.log('');
    Logger.log('⚠️  ALGUNAS PRUEBAS FALLARON');
    Logger.log('');
    Logger.log('Revisa los errores arriba para corregir los problemas.');
  }
  
  // Limpiar datos de prueba
  Logger.log('');
  Logger.log('Limpiando datos de prueba...');
  limpiarDatosPrueba();
  Logger.log('✓ Datos de prueba eliminados');
  
  return resultados;
}

/**
 * Ejecuta una prueba individual
 */
function ejecutarPrueba(resultados, nombre, funcion) {
  resultados.total++;
  Logger.log('');
  Logger.log('───────────────────────────────────────────────');
  Logger.log('Prueba ' + resultados.total + ': ' + nombre);
  Logger.log('───────────────────────────────────────────────');
  
  try {
    var mensaje = funcion();
    resultados.exitosas++;
    resultados.pruebas.push({ nombre: nombre, exito: true, mensaje: mensaje });
    Logger.log(mensaje);
    Logger.log('✓ ÉXITO');
  } catch (error) {
    resultados.fallidas++;
    resultados.pruebas.push({ nombre: nombre, exito: false, error: error.message });
    Logger.log('✗ ERROR: ' + error.message);
    Logger.log('✗ FALLO');
  }
}

/**
 * Limpia los datos de prueba creados
 */
function limpiarDatosPrueba() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Eliminar hoja de prueba si existe
    var hojaPrueba = ss.getSheetByName('TEST_EJECUTIVO_PRUEBA');
    if (hojaPrueba) {
      ss.deleteSheet(hojaPrueba);
    }
    
    // Eliminar datos de prueba de CONFIG_PERFILES
    var hojaConfig = ss.getSheetByName(NOMBRE_HOJA_PERFILES);
    if (hojaConfig) {
      var ultimaFila = hojaConfig.getLastRow();
      if (ultimaFila > 1) {
        var datos = hojaConfig.getRange(2, 1, ultimaFila - 1, 1).getValues();
        var filasAEliminar = [];
        
        for (var i = 0; i < datos.length; i++) {
          var email = datos[i][0] ? datos[i][0].toString() : '';
          if (email.indexOf('test@empresa.cl') !== -1) {
            filasAEliminar.push(i + 2); // +2 porque empezamos en fila 2
          }
        }
        
        // Eliminar de abajo hacia arriba para no afectar índices
        for (var j = filasAEliminar.length - 1; j >= 0; j--) {
          hojaConfig.deleteRow(filasAEliminar[j]);
        }
        
        if (filasAEliminar.length > 0) {
          Logger.log('  - ' + filasAEliminar.length + ' usuarios de prueba eliminados');
        }
      }
    }
    
  } catch (error) {
    Logger.log('⚠️  Error limpiando datos de prueba: ' + error.message);
  }
}

/**
 * Función auxiliar para padding de números
 */
function pad(numero, longitud) {
  var str = numero.toString();
  while (str.length < longitud) {
    str = ' ' + str;
  }
  return str;
}

/**
 * Prueba rápida de configuración básica
 * Ejecuta solo las pruebas esenciales para validar instalación
 */
function pruebaRapida() {
  Logger.log('=== PRUEBA RÁPIDA DE PERFILAMIENTO ===');
  Logger.log('');
  
  try {
    // 1. Verificar/crear CONFIG_PERFILES
    Logger.log('1. Verificando CONFIG_PERFILES...');
    var hoja = crearHojaConfigPerfiles();
    if (!hoja) throw new Error('No se pudo crear CONFIG_PERFILES');
    Logger.log('   ✓ CONFIG_PERFILES OK');
    
    // 2. Validar estructura
    Logger.log('2. Validando estructura...');
    validarEstructuraPerfiles(hoja);
    Logger.log('   ✓ Estructura OK');
    
    // 3. Verificar funciones principales
    Logger.log('3. Verificando funciones...');
    var perfiles = obtenerPerfilesConfigurados();
    Logger.log('   ✓ obtenerPerfilesConfigurados() OK');
    
    var perfil = buscarPerfilEjecutivo('test');
    Logger.log('   ✓ buscarPerfilEjecutivo() OK');
    
    var normalizado = normalizarNombreEjecutivo('Test Usuario');
    if (normalizado !== 'test_usuario') throw new Error('Normalización incorrecta');
    Logger.log('   ✓ normalizarNombreEjecutivo() OK');
    
    Logger.log('');
    Logger.log('✅ PRUEBA RÁPIDA COMPLETADA EXITOSAMENTE');
    Logger.log('');
    Logger.log('El sistema de perfilamiento está instalado correctamente.');
    Logger.log('Ahora puedes:');
    Logger.log('  1. Agregar usuarios a CONFIG_PERFILES');
    Logger.log('  2. Ejecutar distribución de datos');
    Logger.log('  3. Verificar que las hojas tengan columna PERFIL');
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌ PRUEBA RÁPIDA FALLIDA');
    Logger.log('Error: ' + error.message);
    throw error;
  }
}

/**
 * Ejecuta desde el menú
 */
function menuEjecutarPruebas() {
  var ui = SpreadsheetApp.getUi();
  
  var respuesta = ui.alert(
    '🧪 Pruebas del Sistema',
    '¿Deseas ejecutar las pruebas del sistema de perfilamiento?\n\n' +
    'Esto creará datos de prueba temporales que serán eliminados al finalizar.\n\n' +
    'Los resultados aparecerán en los registros de ejecución.',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta === ui.Button.YES) {
    try {
      var resultados = ejecutarPruebasPerfilamiento();
      
      if (resultados.fallidas === 0) {
        ui.alert(
          '✅ Pruebas Completadas',
          'Todas las pruebas pasaron exitosamente.\n\n' +
          'Total: ' + resultados.total + '\n' +
          'Exitosas: ' + resultados.exitosas + '\n\n' +
          'El sistema de perfilamiento está funcionando correctamente.',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert(
          '⚠️ Pruebas con Errores',
          'Algunas pruebas fallaron.\n\n' +
          'Total: ' + resultados.total + '\n' +
          'Exitosas: ' + resultados.exitosas + '\n' +
          'Fallidas: ' + resultados.fallidas + '\n\n' +
          'Revisa los registros (Extensiones > Apps Script > Ejecuciones)',
          ui.ButtonSet.OK
        );
      }
    } catch (error) {
      ui.alert('❌ Error', 'Error ejecutando pruebas: ' + error.message, ui.ButtonSet.OK);
    }
  }
}