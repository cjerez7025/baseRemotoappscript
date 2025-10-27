/**
 * MÓDULO 1: CONFIGURACIÓN GLOBAL
 * Variables y constantes del sistema
 */

var ESTADOS_GESTION = [
  'Sin Gestión', 'Cerrado', 'En Gestión', 
  'No Contactado', 'Interesado', 'Sin Interés'
];

var SUB_ESTADOS_GESTION = [
  'Sin Gestión', 'Volver a llamar', 'Sin motivo',
  'Problema económico', 'No cumple requisitos', 'Mala Experiencia',
  'Cuenta con seguro', 'Prefiere competencia', 'No contesta',
  'Teléfono erróneo', 'Apagado', 'Ya habia sido llamado'
];

var COLUMNAS_NUEVAS = [
  'Propensión', 'FECHA_LLAMADA', 'FECHA_COMPROMISO', 
  'ESTADO', 'SUB_ESTADO', 'NOTA_EJECUTIVO', 
  'ORIGEN_VENTA', 'ESTADO_COMPROMISO'
];

var HOJAS_EXCLUIDAS = [
  'BBDD_REPORTE', 'RESUMEN', 'Sheet1', 'Hoja1', 'Hoja 1', 
  'LLAMADAS', 'PRODUCTIVIDAD', 'Datos desplegables', 
  'DASHBOARD', 'CONFIGURACION', 'TOTALES', 'GRAFICOS'
];

/**
 * Configuración de colores
 */
var COLORES = {
  HEADER_EJECUTIVO: '#FFD700',
  HEADER_REPORTE: '#4CAF50',
  HEADER_RESUMEN: '#00B0F0',
  HEADER_PRODUCTIVIDAD: '#4472C4'
};