/**
 * MÓDULO 2: UTILIDADES
 * Funciones auxiliares reutilizables
 */

/**
 * Convierte número de columna a letra (A, B, C, ..., Z, AA, AB, etc.)
 */
function columnNumberToLetter(num) {
  var letter = '';
  while (num > 0) {
    var remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

/**
 * Formatea el nombre del ejecutivo (Primera_Letra_Mayuscula)
 */
function formatearNombreEjecutivo(nombre) {
  var palabras = nombre.trim().split(' ');
  var resultado = '';
  
  for (var i = 0; i < palabras.length; i++) {
    if (palabras[i].length > 0) {
      resultado += palabras[i].charAt(0).toUpperCase() + palabras[i].slice(1).toLowerCase();
      if (i < palabras.length - 1) resultado += '_';
    }
  }
  
  return resultado.replace(/[^a-zA-Z0-9_]/g, '');
}

/**
 * Obtiene la hoja BBDD_*_REMOTO* (origen de datos)
 */
function obtenerHojaOrigen(ss) {
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    if (/^BBDD_.*_REMOTO/i.test(hojas[i].getName())) {
      return hojas[i];
    }
  }
  return null;
}

/**
 * Obtiene el nombre del ejecutivo de una fila de datos
 */
function obtenerNombreEjecutivo(fila, encabezados) {
  var posibles = ['EJECUTIVO', 'NOMBRE_EJECUTIVO', 'VENDEDOR', 'AGENTE'];
  
  // Buscar en columnas con nombres típicos
  for (var i = 0; i < posibles.length; i++) {
    for (var j = 0; j < encabezados.length; j++) {
      if (encabezados[j].toString().toUpperCase().indexOf(posibles[i]) !== -1) {
        if (fila[j]) return fila[j].toString();
      }
    }
  }
  
  // Buscar en últimas columnas (heurística)
  for (var k = fila.length - 1; k >= 0; k--) {
    var valor = fila[k];
    if (valor && typeof valor === 'string' && valor.indexOf(' ') !== -1) {
      if (/^[a-zA-ZÀ-ÿ\s]+$/.test(valor)) return valor;
    }
  }
  
  return null;
}

/**
 * Agrupa datos por ejecutivo
 */
function agruparPorEjecutivo(filasDatos, encabezados) {
  var ejecutivosPorNombre = {};
  
  for (var i = 0; i < filasDatos.length; i++) {
    var nombreEjecutivo = obtenerNombreEjecutivo(filasDatos[i], encabezados);
    
    if (nombreEjecutivo && nombreEjecutivo.trim() !== '') {
      var nombreFormateado = formatearNombreEjecutivo(nombreEjecutivo);
      
      if (!ejecutivosPorNombre[nombreFormateado]) {
        ejecutivosPorNombre[nombreFormateado] = [];
      }
      
      ejecutivosPorNombre[nombreFormateado].push(filasDatos[i]);
    }
  }
  
  return ejecutivosPorNombre;
}

/**
 * Valida ejecutivos en la base contra las hojas existentes
 */
function validarEjecutivosEnBase(ejecutivosPorNombre, hojas) {
  var alertas = { hojasHuerfanas: [], ejecutivosNuevos: [] };
  var hojasEjecutivas = [];
  
  try {
    for (var i = 0; i < hojas.length; i++) {
      var nombre = hojas[i].getName();
      var esOriginal = /^BBDD_.*_REMOTO/i.test(nombre);
      
      if (!esOriginal && HOJAS_EXCLUIDAS.indexOf(nombre) === -1 && hojas[i].getLastRow() > 1) {
        hojasEjecutivas.push(nombre);
      }
    }
    
    var ejecutivosEnBase = Object.keys(ejecutivosPorNombre);
    
    for (var j = 0; j < hojasEjecutivas.length; j++) {
      if (ejecutivosEnBase.indexOf(hojasEjecutivas[j]) === -1) {
        alertas.hojasHuerfanas.push(hojasEjecutivas[j]);
      }
    }
    
    for (var k = 0; k < ejecutivosEnBase.length; k++) {
      if (hojasEjecutivas.indexOf(ejecutivosEnBase[k]) === -1) {
        alertas.ejecutivosNuevos.push(ejecutivosEnBase[k]);
      }
    }
  } catch (e) {
    Logger.log('Error en validación: ' + e.toString());
  }
  
  return alertas;
}