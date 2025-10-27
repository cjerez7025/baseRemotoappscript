function convertirFecha(fechaInput) {
  // Intentar convertir varios formatos de fecha
  if (!fechaInput) return null;
  
  // Si ya es un objeto Date válido
  if (fechaInput instanceof Date) {
    if (!isNaN(fechaInput.getTime())) {
      return fechaInput;
    }
  }
  
  // Si es un número (timestamp de Excel: días desde 1900)
  if (typeof fechaInput === 'number') {
    // Excel timestamp: días desde 1899-12-30
    var excelEpoch = new Date(1899, 11, 30);
    var fecha = new Date(excelEpoch.getTime() + (fechaInput * 24 * 60 * 60 * 1000));
    if (!isNaN(fecha.getTime())) {
      return fecha;
    }
  }
  
  // Si es string, intentar parsear
  if (typeof fechaInput === 'string') {
    fechaInput = fechaInput.trim();
    
    // Formato DD-MM-YYYY
    var match1 = fechaInput.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (match1) {
      var dia = parseInt(match1[1]);
      var mes = parseInt(match1[2]) - 1;
      var año = parseInt(match1[3]);
      var fecha = new Date(año, mes, dia);
      if (!isNaN(fecha.getTime())) {
        return fecha;
      }
    }
    
    // Formato YYYY-MM-DD
    var match2 = fechaInput.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match2) {
      var año = parseInt(match2[1]);
      var mes = parseInt(match2[2]) - 1;
      var dia = parseInt(match2[3]);
      var fecha = new Date(año, mes, dia);
      if (!isNaN(fecha.getTime())) {
        return fecha;
      }
    }
    
    // Intentar parseo estándar
    var fecha = new Date(fechaInput);
    if (!isNaN(fecha.getTime())) {
      return fecha;
    }
  }
  
  return null;
}// ========================================
// CÓDIGO APPS SCRIPT - PANEL DE LLAMADAS
// ========================================

function obtenerDatosEjecutivo(fechaSeleccionada) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBBDD = ss.getSheetByName('BBDD_REPORTE');
  
  if (!hojaBBDD) {
    Logger.log('ERROR: No se encontró la hoja BBDD_REPORTE');
    return {
      error: true,
      mensaje: 'No se encontró la hoja BBDD_REPORTE',
      tipoError: 'not_found'
    };
  }
  
  try {
    var ultimaFila = hojaBBDD.getLastRow();
    Logger.log('===== VALIDACIÓN BBDD_REPORTE =====');
    Logger.log('Última fila con datos: ' + ultimaFila);
    
    if (ultimaFila < 3) {
      Logger.log('ERROR: Menos de 3 filas. Solo hay ' + ultimaFila + ' filas');
      return {
        error: true,
        mensaje: 'La hoja BBDD_REPORTE no tiene suficientes datos. Por favor contacta al supervisor.',
        tipoError: 'data_error'
      };
    }
    
    Logger.log('VALIDACIÓN OK: La hoja tiene ' + ultimaFila + ' filas');
    
  } catch (e) {
    Logger.log('ERROR al verificar BBDD_REPORTE: ' + e.toString());
    return {
      error: true,
      mensaje: 'Error al verificar la base de datos: ' + e.toString(),
      tipoError: 'check_error'
    };
  }
  
  var hojaActiva = ss.getActiveSheet().getName();
  Logger.log('Hoja activa (ejecutivo): ' + hojaActiva);
  
  var nombreEjecutivo = hojaActiva.replace(/_/g, ' ').toUpperCase();
  
  var datos = hojaBBDD.getDataRange().getValues();
  
  var colEjecutivo = 2;
  var colFechaLlamada = 13;
  var colEstado = 15;
  
  // Determinar la fecha a consultar
  // El cliente envía en formato YYYY-MM-DD
  var fechaConsulta = new Date();
  if (fechaSeleccionada && fechaSeleccionada !== null && fechaSeleccionada !== '') {
    try {
      // Parsear la fecha YYYY-MM-DD como fecha local, no UTC
      var partes = fechaSeleccionada.split('-');
      if (partes.length === 3) {
        var año = parseInt(partes[0]);
        var mes = parseInt(partes[1]) - 1; // Los meses en JS van de 0-11
        var dia = parseInt(partes[2]);
        fechaConsulta = new Date(año, mes, dia);
      } else {
        fechaConsulta = new Date(fechaSeleccionada);
      }
    } catch (e) {
      Logger.log('Error al parsear fecha: ' + e.toString());
      fechaConsulta = new Date();
    }
  }
  fechaConsulta.setHours(0, 0, 0, 0);
  
  var llamadasDia = 0;
  var estadosCerrado = 0;
  var estadosInteresado = 0;
  var estadosEnGestion = 0;
  
  Logger.log('Consultando para: ' + nombreEjecutivo + ' - Fecha: ' + fechaConsulta.toDateString());
  
  for (var i = 1; i < datos.length; i++) {
    var ejecutivo = datos[i][colEjecutivo];
    var fechaLlamada = datos[i][colFechaLlamada];
    var estado = datos[i][colEstado];
    
    if (ejecutivo && ejecutivo.toString().toUpperCase().indexOf(nombreEjecutivo) !== -1) {
      
      if (fechaLlamada) {
        var fecha = convertirFecha(fechaLlamada);
        
        if (fecha) {
          fecha.setHours(0, 0, 0, 0);
          
          if (fecha.getTime() === fechaConsulta.getTime()) {
            llamadasDia++;
            
            if (estado) {
              var estadoStr = estado.toString().trim();
              if (estadoStr === 'Cerrado') {
                estadosCerrado++;
              } else if (estadoStr === 'Interesado') {
                estadosInteresado++;
              } else if (estadoStr === 'En Gestión') {
                estadosEnGestion++;
              }
            }
          }
        }
      }
    }
  }
  
  var metaDiaria = 70;
  var pendientesDia = metaDiaria - llamadasDia;
  var porcentajeDia = Math.round((llamadasDia / metaDiaria) * 100);
  
  var porcCerrado = llamadasDia > 0 ? Math.round((estadosCerrado / llamadasDia) * 100) : 0;
  var porcInteresado = llamadasDia > 0 ? Math.round((estadosInteresado / llamadasDia) * 100) : 0;
  var porcEnGestion = llamadasDia > 0 ? Math.round((estadosEnGestion / llamadasDia) * 100) : 0;
  
  var rotacionAguja = -90 + (porcentajeDia * 1.8);
  
  var badgeClass = 'badge-danger';
  var badgeText = porcentajeDia + '% Completado';
  if (porcentajeDia >= 70) {
    badgeClass = 'badge-success';
  } else if (porcentajeDia >= 40) {
    badgeClass = 'badge-warning';
  }
  
  var opciones = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  var fechaFormateada = fechaConsulta.toLocaleDateString('es-ES', opciones);
  fechaFormateada = fechaFormateada.charAt(0).toUpperCase() + fechaFormateada.slice(1);
  
  var hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  var horasRestantes = 0;
  
  if (fechaConsulta.getTime() === hoy.getTime()) {
    var ahora = new Date();
    var horaFin = new Date();
    horaFin.setHours(18, 0, 0, 0);
    horasRestantes = Math.max(0, Math.round((horaFin - ahora) / (1000 * 60 * 60)));
  }
  
  Logger.log('===== RESULTADOS =====');
  Logger.log('Fecha enviada (cliente): ' + fechaSeleccionada);
  Logger.log('Fecha parseada (servidor): ' + fechaConsulta.toDateString());
  Logger.log('Hoy (servidor): ' + hoy.toDateString());
  Logger.log('¿Es hoy?: ' + (fechaConsulta.getTime() === hoy.getTime()));
  Logger.log('Ejecutivo buscado: ' + nombreEjecutivo);
  Logger.log('Llamadas encontradas: ' + llamadasDia);
  Logger.log('Cerrado: ' + estadosCerrado + ', Interesado: ' + estadosInteresado + ', En Gestión: ' + estadosEnGestion);
  Logger.log('Porcentaje: ' + porcentajeDia + '%');
  Logger.log('======================');
  
  return {
    error: false,
    nombreEjecutivo: nombreEjecutivo,
    fechaFormateada: fechaFormateada,
    fechaISO: fechaConsulta.toISOString().split('T')[0],
    llamadasDia: llamadasDia,
    metaDiaria: metaDiaria,
    pendientesDia: Math.max(0, pendientesDia),
    porcentajeDia: porcentajeDia,
    rotacionAguja: rotacionAguja,
    badgeClass: badgeClass,
    badgeText: badgeText,
    horasRestantes: horasRestantes,
    estadosCerrado: estadosCerrado,
    estadosInteresado: estadosInteresado,
    estadosEnGestion: estadosEnGestion,
    porcCerrado: porcCerrado,
    porcInteresado: porcInteresado,
    porcEnGestion: porcEnGestion,
    esHoy: fechaConsulta.getTime() === hoy.getTime()
  };
}