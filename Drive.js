// =================================================================================
// --- FUNCIONES DE AYUDA PARA DRIVE ---
// =================================================================================

/**
 * Obtiene o crea la carpeta de un empleado en Drive.
 */
function obtenerOCrearCarpetaPersona(calidadContractual, nombreCompleto) {
  const carpetaRaiz = DriveApp.getFolderById(ID_CARPETA_RAIZ);
  let iteratorCalidad = carpetaRaiz.getFoldersByName(calidadContractual.toUpperCase());
  const carpetaCalidad = iteratorCalidad.hasNext() ? iteratorCalidad.next() : carpetaRaiz.createFolder(calidadContractual.toUpperCase());
  let iteratorPersona = carpetaCalidad.getFoldersByName(nombreCompleto);
  return iteratorPersona.hasNext() ? iteratorPersona.next() : carpetaCalidad.createFolder(nombreCompleto);
}

/**
 * Verifica que el usuario tenga permisos para acceder a carpetas y plantillas críticas.
 */
function verificarPermisosCriticos() {
  const ui = SpreadsheetApp.getUi();
  const errores = [];
  try {
    const carpeta = DriveApp.getFolderById(ID_CARPETA_RAIZ);
    carpeta.setDescription("Verificación de permisos realizada el " + new Date().toLocaleString());
  } catch (e) {
    errores.push(`- Carpeta Raíz: ACCESO DENEGADO. Necesitas permisos de 'Editor'.`);
  }
  for (const nombreTemplate in TEMPLATES_UNIVERSALES) {
    const idTemplate = TEMPLATES_UNIVERSALES[nombreTemplate];
    try {
      DriveApp.getFileById(idTemplate).getName();
    } catch (e) {
      errores.push(`- Plantilla '${nombreTemplate}': ACCESO DENEGADO.`);
    }
  }
  if (errores.length > 0) {
    const mensajeError = "Se encontraron problemas de permisos:\n\n" + errores.join('\n') + "\n\nContacta al administrador.";
    ui.alert('❌ Verificación Fallida', mensajeError, ui.ButtonSet.OK);
  } else {
    ui.alert('✅ ¡Éxito!', 'Tienes todos los permisos necesarios.', ui.ButtonSet.OK);
  }
}