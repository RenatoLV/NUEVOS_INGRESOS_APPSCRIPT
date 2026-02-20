// =================================================================================
// --- FUNCIONES DEL BACKEND DEL FORMULARIO (OPTIMIZADO Y ROBUSTO) ---
// =================================================================================

function obtenerOCrearCarpetaPersona(calidad, nombre) {
  const lock = LockService.getScriptLock();
  try {
    // Esperamos hasta 30s si hay otro proceso creando carpetas
    lock.waitLock(30000); 

    // 1. Limpieza de entradas para evitar carpetas con espacios fantasma
    const nombreCalidad = (calidad || 'SIN CALIDAD DEFINIDA').trim().toUpperCase();
    const nombreFuncionario = (nombre || 'SIN NOMBRE').trim().toUpperCase();

    // 2. Buscar/Crear carpeta de CALIDAD (Nivel 1)
    let carpetaCalidad = null;
    const iterCalidad = DriveApp.getFoldersByName(nombreCalidad);
    
    while (iterCalidad.hasNext()) {
      const f = iterCalidad.next();
      if (!f.isTrashed()) { 
        carpetaCalidad = f;
        break;
      }
    }
    
    if (!carpetaCalidad) {
      carpetaCalidad = DriveApp.createFolder(nombreCalidad);
    }

    // 3. Buscar/Crear carpeta del FUNCIONARIO (Nivel 2) dentro de la Calidad
    let carpetaFuncionario = null;
    const iterFuncionario = carpetaCalidad.getFoldersByName(nombreFuncionario);
    
    while (iterFuncionario.hasNext()) {
      const f = iterFuncionario.next();
      if (!f.isTrashed()) {
        carpetaFuncionario = f;
        break;
      }
    }
    
    if (!carpetaFuncionario) {
      carpetaFuncionario = carpetaCalidad.createFolder(nombreFuncionario);
    }

    return carpetaFuncionario;

  } catch (e) {
    Logger.log("Error crítico en obtenerOCrearCarpetaPersona: " + e.message);
    throw new Error("Error gestionando carpetas en Drive: " + e.message);
  } finally {
    lock.releaseLock(); // Siempre liberamos el bloqueo
  }
}

/**
 * Obtiene los datos de la hoja DICCIONARIO para poblar los selects.
 */
function obtenerDatosDiccionario() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const diccionarioSheet = ss.getSheetByName(DICCIONARIO_SHEET);
    if (!diccionarioSheet) throw new Error(`Hoja ${DICCIONARIO_SHEET} no encontrada.`);
    
    const dataRange = diccionarioSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift().map(h => h.trim());

    const resultado = {};
    const jerarquiaData = [];
    const gradosEscalafones = [];
    const perfilesConDescripcion = [];
    const escalaDeSueldos = [];

    const COLS = {
      direccion: headers.indexOf('DIRECCION'),
      departamento: headers.indexOf('DEPARTAMENTO'),
      oficina: headers.indexOf('SECCION'),
      asignado: headers.indexOf('ASIGNADO A'),
      grado: headers.indexOf('GRADO_NUM'),
      escalafon: headers.indexOf('ESCALAFON_ASOCIADO'),
      perfil: headers.indexOf('PERFIL_HONORARIOS'),
      descPerfil: headers.indexOf('DESCRIPCION_PERFIL'),
      sueldoGrado: headers.indexOf('SUELDO_GRADO'),
      sueldoEscalafon: headers.indexOf('SUELDO_ESCALAFON'),
      montoBruto: headers.indexOf('MONTO_BRUTO')
    };

    values.forEach(row => {
      // Grados
      if (COLS.grado > -1 && row[COLS.grado]) {
        gradosEscalafones.push({
          grado: row[COLS.grado].toString().trim(),
          escalafon: COLS.escalafon > -1 && row[COLS.escalafon] ? row[COLS.escalafon].toString().trim() : ''
        });
      }
      // Perfiles (IMPORTANTE: Mantiene el texto original de la descripción)
      if (COLS.perfil > -1 && row[COLS.perfil]) {
        perfilesConDescripcion.push({
          perfil: row[COLS.perfil].toString().trim(), 
          descripcion: COLS.descPerfil > -1 && row[COLS.descPerfil] ? row[COLS.descPerfil].toString().trim() : ''
        });
      }
      // Sueldos
      if (COLS.sueldoGrado > -1 && row[COLS.sueldoGrado]) {
        escalaDeSueldos.push({
          grado: row[COLS.sueldoGrado].toString().trim(),
          escalafon: COLS.sueldoEscalafon > -1 && row[COLS.sueldoEscalafon] ? row[COLS.sueldoEscalafon].toString().trim() : '',
          monto: COLS.montoBruto > -1 ? row[COLS.montoBruto] : 0
        });
      }
      // Jerarquía
      if (COLS.direccion > -1 && row[COLS.direccion]) {
        jerarquiaData.push({
          direccion: row[COLS.direccion].toString().trim(),
          departamento: COLS.departamento > -1 && row[COLS.departamento] ? row[COLS.departamento].toString().trim() : '',
          oficina: COLS.oficina > -1 && row[COLS.oficina] ? row[COLS.oficina].toString().trim() : '',
          asignado: COLS.asignado > -1 && row[COLS.asignado] ? row[COLS.asignado].toString().trim().toUpperCase() : ''
        });
      }
    });

    resultado.gradosEscalafones = gradosEscalafones;
    resultado.perfilesConDescripcion = perfilesConDescripcion;
    resultado.escalaDeSueldos = escalaDeSueldos;
    resultado.jerarquia = jerarquiaData;

    // --- LECTURA DE JEFATURAS (NOMBRE + CORREO) ---
    const colEmail = headers.indexOf('DIRECTORIO_JEFATURAS');
    const colNombre = headers.indexOf('NOMBRESC_CORREOS_JEFATURAS');
    
    const jefaturas = [];
    if (colEmail > -1) {
        values.forEach(row => {
            const email = row[colEmail] ? String(row[colEmail]).trim() : '';
            const nombre = (colNombre > -1 && row[colNombre]) ? String(row[colNombre]).trim() : '';
            
            if (email && email.includes('@')) {
                // Formato para Select2: id (valor real), text (lo que se ve y busca)
                jefaturas.push({
                    id: email,
                    text: nombre ? `${nombre} <${email}>` : email, // "Juan Perez <jperez@...>"
                    nombreSolo: nombre || email.split('@')[0]
                });
            }
        });
    }
    resultado['DIRECTORIO_JEFATURAS_DATA'] = jefaturas;

    // Campos restantes dinámicos
    const excludedHeaders = Object.values(COLS).map(index => headers[index]);
    headers.forEach((header, index) => {
      if (header && !excludedHeaders.includes(header)) {
        const valoresUnicos = [...new Set(values.map(row => row[index]).filter(Boolean).map(String))];
        resultado[header] = valoresUnicos;
      }
    });

    resultado.templates = Object.keys(TEMPLATES_UNIVERSALES);
    return resultado;
  } catch (e) {
    Logger.log(e);
    throw new Error('Error cargando diccionario: ' + e.message);
  }
}

/**
 * Busca registros por RUT en la hoja REGISTRO_MAESTRO.
 */
function buscarPorRut(rutBuscado) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    if (!hoja) throw new Error(`Hoja '${REGISTRO_MAESTRO_SHEET}' no encontrada.`);

    const datos = hoja.getDataRange().getValues();
    const headers = datos.shift();
    const indiceRut = headers.indexOf("RUT");
    if (indiceRut === -1) throw new Error("No se encontró columna 'RUT'.");

    const normalizarRut = (rut) => {
      if (!rut || typeof rut !== 'string') return '';
      return rut.replace(/[\.\-]/g, '').toLowerCase();
    };
    
    const rutBuscadoNormalizado = normalizarRut(rutBuscado);
    const resultados = [];

    datos.forEach((fila, idx) => {
      const rutEnFila = fila[indiceRut];
      if (normalizarRut(rutEnFila) === rutBuscadoNormalizado) {
        // ID real o fallback al índice
        const registro = { ID: fila[headers.indexOf('ID')] || (idx + 1) };
        headers.forEach((header, i) => {
          if (fila[i] instanceof Date) {
            registro[header] = fila[i].toISOString();
          } else {
            registro[header] = fila[i];
          }
        });
        resultados.push(registro);
      }
    });

    return resultados;
  } catch (error) {
    console.error("Error en buscarPorRut:", error);
    return { error: true, message: error.message };
  }
}

/**
 * Guarda o actualiza un registro.
 * CORRECCIÓN IMPORTANTE: Filtra URLs basura y preserva formato de descripciones.
 */
function guardarDatos(registro, fileData) {
  try {
    Logger.log("--- INICIO guardado --- ID Recibido: " + registro.ID);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    if (!sheet) throw new Error(`Hoja ${REGISTRO_MAESTRO_SHEET} no existe.`);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let newId = null;
    
    // --- 1. LIMPIEZA DE DATOS (PRESERVANDO FORMATO 1-1) ---
    // Lista de campos que NO deben forzarse a mayúsculas para coincidir exactamente con Diccionario
    const CAMPOS_PRESERVAR_FORMATO = [
      'DESCRIPCION PERFIL', 
      'FUNCION', 
      'OBSERVACIONES', 
      'URL ORDEN DE SERVICIO', 
      'DESCRIPCION DISCAPACIDAD',
      'FUNCION PRIMER SEMESTRE',
      'FUNCION SEGUNDO SEMESTRE',
      'PERFIL', // A veces los perfiles tienen casing mixto en el diccionario
      'PROGRAMA',
      'NOMBRE DE CONVENIO'
    ];

    for (const key in registro) {
      if (typeof registro[key] === 'string') {
        if (CAMPOS_PRESERVAR_FORMATO.includes(key)) {
           // Solo quitamos espacios de los extremos, mantenemos Mayúsculas/Minúsculas/Comillas
           registro[key] = registro[key].trim();
        } else {
           // Resto de campos (Nombres, Direcciones, etc.) van en Mayúsculas estándar
           registro[key] = registro[key].trim().toUpperCase();
        }
      }
    }

    // Fix común: Horario 24/7 con comilla simple para evitar que Excel lo tome como fórmula
    if (registro['HORARIO'] && registro['HORARIO'].includes('24/7')) {
      registro['HORARIO'] = registro['HORARIO'].replace(/24\/7/g, "'24/7");
    }
    
    // Calculo automático de MES y AÑO
    if (registro['FECHA DE REGISTRO']) {
      const fecha = new Date(registro['FECHA DE REGISTRO']);
      const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
      if(!isNaN(fecha.getTime())){
          registro['MES'] = meses[fecha.getUTCMonth()];
          registro['AÑO'] = fecha.getUTCFullYear();
      }
    }
    
    // --- 2. GESTIÓN DE ID ---
    let esNuevoRegistro = false;
    let rowIndex = -1;

    if (registro.ID && registro.ID !== '') {
      const data = sheet.getDataRange().getValues();
      const idIndex = headers.indexOf('ID');
      // Buscamos coincidencia exacta de ID
      rowIndex = data.findIndex(row => String(row[idIndex]) === String(registro.ID));
      
      if (rowIndex <= 0) { // No encontrado (rowIndex 0 son headers)
         Logger.log("ID no encontrado, se creará nuevo registro.");
         registro.ID = ''; // Forzar creación
         rowIndex = -1;
      }
    }
    
    if (!registro.ID || rowIndex === -1) {
      esNuevoRegistro = true;
      const idIndex = headers.indexOf('ID');
      let lastId = 0;
      if (sheet.getLastRow() > 1) {
          const ids = sheet.getRange(2, idIndex + 1, sheet.getLastRow() - 1, 1).getValues().flat();
          const maxId = ids.reduce((max, val) => {
              const num = parseInt(val);
              return !isNaN(num) && num > max ? num : max;
          }, 0);
          lastId = maxId;
      }
      newId = lastId + 1;
      registro.ID = newId;
      
      // Valor por defecto
      if (!registro['ESTADO DE TRASPASO A NOMINA']) {
        registro['ESTADO DE TRASPASO A NOMINA'] = 'EN ESPERA DE MOVER';
      }
    }

    const idActual = newId || registro.ID;

    // --- 3. LÓGICA ROBUSTA DE ARCHIVOS ---
    let fileUrl = registro['URL ORDEN DE SERVICIO'] || '';
    
    // A. Limpieza Preventiva: Si la URL existente parece un link al Sheet o Script, se borra.
    if (fileUrl) {
        const urlLower = fileUrl.toLowerCase();
        if (urlLower.includes('spreadsheets') || 
            urlLower.includes('script.google.com') || 
            urlLower.includes('/edit') ||
            !urlLower.startsWith('http')) {
            Logger.log(`URL INVÁLIDA detectada y eliminada: ${fileUrl}`);
            fileUrl = ''; 
        }
    }

    // B. Procesamiento de Nuevo Archivo (Prioridad Alta)
    if (fileData && fileData.data && fileData.fileName) {
      try {
          Logger.log("Iniciando subida de archivo nuevo...");
          
          // Usamos la función helper robusta (con bloqueo y verificación de papelera)
          const carpetaPersona = obtenerOCrearCarpetaPersona(registro['CALIDAD CONTRACTUAL'], registro['NOMBRE COMPLETO']);
          
          const decodedData = Utilities.base64Decode(fileData.data);
          const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
          const archivoSubido = carpetaPersona.createFile(blob);
          
          // Pequeña espera para propagación en Drive (ayuda a generar la URL correcta)
          Utilities.sleep(1000); 
          fileUrl = archivoSubido.getUrl();
          Logger.log("Archivo subido con éxito: " + fileUrl);

          // Validación final post-subida para asegurar que es un link válido
          if (!fileUrl.includes('drive.google.com') && !fileUrl.includes('docs.google.com')) {
             throw new Error("La URL generada por Drive no parece válida.");
          }

      } catch (uploadError) {
          Logger.log("Error al subir archivo: " + uploadError.message);
          return { success: false, message: "Error al guardar el archivo adjunto: " + uploadError.message };
      }
    }
    
    // Asignar la URL final (ya sea la nueva o la existente limpia)
    registro['URL ORDEN DE SERVICIO'] = fileUrl;

    // --- 4. ESCRITURA EN HOJA ---
    const rowData = headers.map(header => registro[header] !== undefined ? registro[header] : '');
    
    if (esNuevoRegistro) {
        sheet.appendRow(rowData);
    } else {
        // rowIndex viene de data (con headers), por lo que rowIndex 1 es la fila 2 de la hoja.
        // getRange usa base 1. Entonces getRange(rowIndex + 1...)
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([rowData]);
    }

    // Validación de estado post-guardado (si existe la función de validación)
    try { if(typeof coreUpdateAllStatusOptimized === 'function') coreUpdateAllStatusOptimized(); } catch(e){}

    const message = esNuevoRegistro ? `Registro creado exitosamente (ID: ${idActual})` : `Registro actualizado exitosamente (ID: ${idActual})`;
    return { success: true, message: message, newId: (esNuevoRegistro ? idActual : null) };
    
  } catch (e) {
    Logger.log(`ERROR FATAL: ${e.message}`);
    return { success: false, message: 'Error crítico en el servidor: ' + e.message };
  }
}

/**
 * Verifica si un RUT ya existe.
 */
function verificarRutExistente(rutAVerificar) {
  try {
    if (!rutAVerificar) return false;
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    const indiceRut = headers.indexOf('RUT');
    if (indiceRut === -1) return false;

    const rutsDeLaHoja = hoja.getRange(2, indiceRut + 1, hoja.getLastRow() - 1, 1).getValues();
    const normalizar = rut => String(rut).replace(/[.\-]/g, '').toLowerCase();
    const rutNormalizado = normalizar(rutAVerificar);

    return rutsDeLaHoja.some(fila => fila[0] && normalizar(fila[0]) === rutNormalizado);
  } catch (error) {
    return false;
  }
}

/**
 * Busca un registro por su ID único.
 */
function buscarPorId(idBuscado) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const datos = hoja.getDataRange().getValues();
    const headers = datos.shift();
    const idColumnIndex = headers.indexOf("ID");
    if (idColumnIndex === -1) throw new Error("Falta columna ID");

    const filaEncontrada = datos.find(fila => String(fila[idColumnIndex]) === String(idBuscado));

    if (filaEncontrada) {
      const registro = {};
      headers.forEach((header, i) => {
        if (filaEncontrada[i] instanceof Date) {
          registro[header] = filaEncontrada[i].toISOString();
        } else {
          registro[header] = filaEncontrada[i];
        }
      });
      return registro;
    }
    return null;
  } catch (error) {
    return { error: true, message: error.message };
  }
}

/**
 * Función independiente para subir archivo (por si se usa fuera del guardado normal).
 */
function subirArchivoORNS(fileData, idEmpleado) {
  try {
    if (!fileData || !idEmpleado) throw new Error("Faltan datos.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColumnIndex = headers.indexOf('ID');

    const rowIndexInData = data.findIndex(row => String(row[idColumnIndex]) === String(idEmpleado));
    if (rowIndexInData === -1) throw new Error(`ID ${idEmpleado} no encontrado.`);
    
    const filaReal = rowIndexInData + 2;
    const registro = {};
    headers.forEach((header, i) => { registro[header] = data[rowIndexInData][i]; });
    
    // Usamos el helper robusto
    const carpetaPersona = obtenerOCrearCarpetaPersona(registro['CALIDAD CONTRACTUAL'], registro['NOMBRE COMPLETO']);

    const decodedData = Utilities.base64Decode(fileData.data, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    const archivoSubido = carpetaPersona.createFile(blob);
    
    Utilities.sleep(1000); 
    const fileUrl = archivoSubido.getUrl();

    if (fileUrl.includes('spreadsheets')) throw new Error("Error generando URL de archivo.");

    const urlColumnIndex = headers.indexOf('URL ORDEN DE SERVICIO');
    if (urlColumnIndex !== -1) {
      sheet.getRange(filaReal, urlColumnIndex + 1).setValue(fileUrl);
    }
    
    try { if(typeof coreUpdateAllStatusOptimized === 'function') coreUpdateAllStatusOptimized(); } catch(e){}
    
    return { success: true, fileUrl: fileUrl, fileName: fileData.fileName };

  } catch (e) {
    Logger.log(`Error en subirArchivoORNS: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Genera documentos PDF a partir de plantillas.
 */
function generarDocumentosDesdeWeb(idEmpleado, listaDocumentos) {
  try {
    if (!idEmpleado || !listaDocumentos || listaDocumentos.length === 0) throw new Error("Faltan datos.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REGISTRO_MAESTRO_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColumnIndex = headers.indexOf('ID');
    const rowIndexInData = data.findIndex(row => String(row[idColumnIndex]) === String(idEmpleado));
    
    if (rowIndexInData === -1) throw new Error("Registro no encontrado");
    
    const registro = {};
    headers.forEach((header, index) => {
      let value = data[rowIndexInData][index];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      }
      registro[header] = value;
    });

    const nombreCompleto = registro['NOMBRE COMPLETO'].trim();
    const calidadContractual = registro['CALIDAD CONTRACTUAL'].trim();
    
    // Helper robusto para encontrar carpeta
    const carpetaPersona = obtenerOCrearCarpetaPersona(calidadContractual, nombreCompleto);

    // Lógica para campo compuesto
    let funcionOPerfilCompleto = '';
    const calidad = registro['CALIDAD CONTRACTUAL'];
    if (calidad === 'HONORARIOS SUMA ALZADA') {
      funcionOPerfilCompleto = `${registro['PERFIL'] || ''}\n${registro['DESCRIPCION PERFIL'] || ''}`.trim();
    } else if (calidad === 'ADMINISTRACION DE FONDOS') {
      funcionOPerfilCompleto = `${registro['NOMBRE DE CONVENIO'] || ''}\n${registro['CARGO'] || ''}\n${registro['FUNCION'] || ''}`.trim();
    } else {
      let funciones = [];
      if (registro['FUNCION'] && registro['FUNCION'] !== 'NO APLICA') funciones.push(registro['FUNCION']);
      if (registro['FUNCION PRIMER SEMESTRE'] && registro['FUNCION PRIMER SEMESTRE'] !== 'NO APLICA') funciones.push(`1er Semestre: ${registro['FUNCION PRIMER SEMESTRE']}`);
      if (registro['FUNCION SEGUNDO SEMESTRE'] && registro['FUNCION SEGUNDO SEMESTRE'] !== 'NO APLICA') funciones.push(`2do Semestre: ${registro['FUNCION SEGUNDO SEMESTRE']}`);
      funcionOPerfilCompleto = funciones.join('\n');
    }

    listaDocumentos.forEach(nombreTemplate => {
      const templateId = TEMPLATES_UNIVERSALES[nombreTemplate];
      if (templateId) {
        const plantilla = DriveApp.getFileById(templateId);
        const nuevoDoc = plantilla.makeCopy(`${nombreTemplate} - ${nombreCompleto}`, carpetaPersona);
        const doc = DocumentApp.openById(nuevoDoc.getId());
        const body = doc.getBody();

        body.replaceText('{{FUNCION_O_PERFIL_COMPLETO}}', funcionOPerfilCompleto);
        headers.forEach(header => {
           body.replaceText(`{{${header}}}`, registro[header] || '');
        });
        
        doc.saveAndClose();
      }
    });

    return { success: true, folderUrl: carpetaPersona.getUrl(), nombre: nombreCompleto };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Helper para el frontend: Obtiene la URL de la carpeta.
 */
function obtenerUrlCarpeta(idEmpleado) {
  try {
    const registro = buscarPorId(idEmpleado);
    if (registro) {
      const carpetaPersona = obtenerOCrearCarpetaPersona(registro['CALIDAD CONTRACTUAL'], registro['NOMBRE COMPLETO']);
      return carpetaPersona.getUrl();
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Obtiene los datos completos y calcula faltantes.
 */
function obtenerDatosCompletosPorId(id) {
  try {
    const registro = buscarPorId(id);
    if (!registro || registro.error) return registro;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Helpers internos para validación
    const isBlank = (val) => !val || String(val).trim() === '';
    const isMissing = (val) => {
        if (!val || String(val).trim() === '' || String(val).trim() === '-' || String(val).trim() === '0') return true;
        return false;
    };

    const calidad = registro['CALIDAD CONTRACTUAL'];
    let faltantesGeneral = [];
    let faltantesContratos = [];

    const camposBaseContratos = [
      'RUT', 'NOMBRES', 'APELLIDOS', 'CALIDAD CONTRACTUAL', 'ASIGNADO A', 'DOMICILIO', 
      'DIRECCION', 'MONTO BRUTO', 'FECHA INGRESO', 'FECHA DE TERMINO', 'JORNADA', 'HORARIO'
    ];
    camposBaseContratos.forEach(f => {
      if (isMissing(registro[f])) faltantesContratos.push(f);
    });

    const noBlancos_General = [
      'CODIGO', 'DIRECCION', 'NOMBRES', 'APELLIDOS', 'RUT', 'FECHA DE NACIMIENTO', 
      'ESTADO CIVIL', 'SEXO', 'INSCRIPCION', 'NACIONALIDAD', 'PUEBLOS ORIGINARIOS', 
      'DISCAPACIDAD', 'DOMICILIO', 'NUMERO CONTACTO', 'CORREO ELECTRONICO', 'FECHA DE REGISTRO'
    ];
    noBlancos_General.forEach(f => { if (isBlank(registro[f])) faltantesGeneral.push(f); });
    
    const calidadPrincipal = String(calidad).startsWith("PRESTADORES DE SERVICIO") ? "PRESTADORES DE SERVICIO" : calidad;
    
    if (["CONTRATA", "PLANTA", "PLANTA SUPLENCIA"].includes(calidadPrincipal)) {
        ['DECRETO', 'GRADO', 'ESCALAFON'].forEach(f => { if (isMissing(registro[f])) faltantesGeneral.push(f); });
    }

    registro.faltantesGeneral = [...new Set(faltantesGeneral)];
    registro.faltantesContratos = [...new Set(faltantesContratos)];

    return registro;

  } catch (e) {
    return { error: true, message: e.message };
  } 
}

function verificarIntegridadDeDatos() {
  try {
    const sheetName = 'REGISTRO_MAESTRO'; // Asegúrate de que este nombre sea correcto
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(sheetName);
    
    // Si no existe la hoja, devolvemos 0 para forzar una recarga o manejo de error
    if (!ws) return { ultimoId: 0, totalFilas: 0 };

    const lastRow = ws.getLastRow();
    
    // Si solo hay encabezados o está vacía
    if (lastRow < 2) return { ultimoId: 0, totalFilas: 0 };

    // Obtenemos la columna de IDs (Columna A = 1) para encontrar el ID más alto real.
    // Esto es más seguro que usar solo el número de filas si borraste registros intermedios.
    const rangoIds = ws.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let maxId = 0;
    // Recorremos rápido el array en memoria
    for (let i = 0; i < rangoIds.length; i++) {
      const val = Number(rangoIds[i][0]);
      if (!isNaN(val) && val > maxId) {
        maxId = val;
      }
    }

    return {
      ultimoId: maxId,
      totalFilas: lastRow
    };
    
  } catch (e) {
    Logger.log("Error en verificarIntegridadDeDatos: " + e.toString());
    // En caso de error, devolvemos ceros para evitar bloqueos, 
    // el cliente intentará cargar normal.
    return { ultimoId: 0, totalFilas: 0, error: e.toString() };
  }
}