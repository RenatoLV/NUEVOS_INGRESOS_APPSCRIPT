// =================================================================================
// --- CONSTANTES GLOBALES ---
// =================================================================================
const ID_HOJA_DOTACION = '1YO9yrSC7gtLnmUrcfcM39tJ_5BN_BCsyiWloKAzFlc8';
const REGISTRO_MAESTRO_SHEET = 'REGISTRO_MAESTRO';
const DICCIONARIO_SHEET = 'DICCIONARIO';
const ID_CARPETA_RAIZ = '1mrhLGqTF1I_MWUigLdiolTmeNigsUimx';
const CHECKBOX_COLUMN_NAME = 'GENERAR DOCUMENTOS';

// Constantes para la Lógica de Validación de Estado
const COL_ESTADO_INGRESO = "ESTADO DE INFORMACION"; // Columna que se actualizará
const COL_CALIDAD_CONTRACTUAL = "CALIDAD CONTRACTUAL"; // Columna que determina la lógica a usar
const COL_RUT = "RUT"; // Columna principal para determinar si una fila está en uso

const TEMPLATES_UNIVERSALES = {
  'FORMULARIO DE INGRESO': '18Yl2VKT1wA-hyPiyRIc2xKU5LaCz72enA9w-w9Yxyh4',
  'DECLARACION SIMPLE DE PARENTESCO': '19rTkZ0xg1h9D0hCB5PHhwJV6iQPldAhmOPwn66Af6mk',
  'DECLARACIÓN JURADA SOBRE ESTUDIOS': '1Z0BOFNDVJbgWdRuiv4T_bTll2XWi6_PBGNZx3b2B148',
  'CERTIFICADO DE SALUD': '12JHqMHhS1X_Z_B8ZRJKFZOWkHCRDrhL_5yn4jZLiQSE',
  'DECLARACIONES': '15cjcgiKj4qqDMz1vKZa_zqZMLSBnAVTHyPIJ5uvw5E4',
};

// =================================================================================
// --- FUNCIONES DEL MENÚ Y DISPARADORES (TRIGGERS) ---
// =================================================================================

function doGet(e) {
  // Crea el template HTML desde el archivo 'Formulario.html'
  const template = HtmlService.createTemplateFromFile('Formulario');
  
  // Pasa el modo de la aplicación al HTML (en este caso, 'WEBAPP')
  template.mode = 'WEBAPP'; 

  // Evalúa el template y devuelve el resultado como una página HTML
  const htmlOutput = template.evaluate()
    .setTitle('Gestión de Nuevos Ingresos') // Establece el título de la pestaña del navegador
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Asegura que se vea bien en móviles

  return htmlOutput;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Formulario de Gestión')
    .addItem('Abrir Formulario', 'mostrarFormulario')
    .addItem('⚡ Actualizar Todos los Estados (Rápido)', 'coreUpdateAllStatusOptimized') 
    .addSeparator()
    .addItem('Verificar Mis Permisos', 'verificarPermisosCriticos')
    .addToUi();
}

/**
 * Disparador que se activa al enviar un formulario (si está configurado)
 * o al editar la hoja.
 */
function manejarEnvioDeFormulario(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const FilaIngresada = range.getRow();

  if (sheet.getName() === REGISTRO_MAESTRO_SHEET) {
    Logger.log(`Nuevo envío de formulario detectado en la fila: ${FilaIngresada}. Actualizando estado.`);
    Utilities.sleep(500); 
    // Llama a la función global que contiene la lógica de validación
    coreUpdateAllStatusOptimized(); // <-- ¡Esta es la llamada correcta a la función del Validacion.gs!
  }
}