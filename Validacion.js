// =================================================================================
// --- FUNCIONES DE VALIDACIÓN DE ESTADO ---
// =================================================================================

/**
 * Función principal que recorre toda la hoja y actualiza las columnas de estado.
 */
function coreUpdateAllStatusOptimized() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`No se encontró la hoja llamada "${REGISTRO_MAESTRO_SHEET}".`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  const headers = allData.shift(); 
  const headerMap = {};
  headers.forEach((header, i) => { headerMap[header.trim()] = i; });

  const indices = {
    estadoGeneral: headerMap["ESTADO DE INFORMACION"],
    faltanteGeneral: headerMap["INFORMACION FALTANTE"],
    estadoContratos: headerMap["ESTADO PARA CONTRATOS"],
    faltanteContratos: headerMap["FALTANTE_CONTRATOS"],
    calidad: headerMap["CALIDAD CONTRACTUAL"],
    rut: headerMap["RUT"]
  };
  
  const NOMBRE_COLUMNA_ESTADO_REAL = "ESTADO INGRESO"; 

  const results = allData.map(rowData => {
    if (isBlank(rowData[headerMap['RUT']])) {
      return ["", "", "PENDIENTE", ""];
    }
    
    const estadoRealDelIngreso = rowData[headerMap[NOMBRE_COLUMNA_ESTADO_REAL]];
    if (estadoRealDelIngreso && String(estadoRealDelIngreso).trim().toUpperCase() === "INGRESO CANCELADO") {
      return ["INFORMACION COMPLETA", "", "CANCELADO", ""];
    }

    const calidad = rowData[indices.calidad];
    let faltantesGeneral = [];
    let faltantesContratos = [];

    // --- LISTA 1: CAMPOS OBLIGATORIOS PARA CONTRATOS ---
    const camposBaseContratos = [
        'RUT', 'NOMBRES', 'APELLIDOS', 'CALIDAD CONTRACTUAL', 'ASIGNADO A', 'DOMICILIO', 
        'DIRECCION', 'MONTO BRUTO', 'FECHA INGRESO', 
        'FECHA DE TERMINO', 'JORNADA', 'HORARIO'
    ];
    camposBaseContratos.forEach(f => {
        if (isMissing(rowData[headerMap[f]], f)) faltantesContratos.push(f);
    });

    // --- LISTA 2: CAMPOS OBLIGATORIOS GENERALES ---
    const noBlancos_General = [
      'CODIGO', 'DIRECCION', 'NOMBRES', 'APELLIDOS', 'RUT', 'FECHA DE NACIMIENTO', 
      'ESTADO CIVIL', 'SEXO', 'INSCRIPCION', 'NACIONALIDAD', 
      'PUEBLOS ORIGINARIOS', 'DISCAPACIDAD', 'DOMICILIO', 'NUMERO CONTACTO', 
      'CORREO ELECTRONICO', 'FECHA DE REGISTRO'
    ];
    const noFaltantes_General = [
      'AREA DE GESTION', 'PROFESION', 'INSTITUCION DE ESTUDIOS',
      'NUMERO ORDEN DE SERVICIO', 'ASIGNADO A', 'FECHA ORNS/NI', 
      'FECHA VENCIMIENTO CI', 'FECHA C ANTECEDENTES', 'FECHA C SALUD', 
      'CERTIFICADO AFILIACION', 'SITUACION MILITAR', 'ENTREVISTA', 
      'ENROLAJE MARCAJE', 'PENSION DE ALIMENTOS', 'FOLIO PENSION', 
      'FECHA INGRESO', 'FECHA INGRESO MUNICIPALIDAD', 'FECHA DE TERMINO', 
      'MONTO BRUTO', 'JORNADA', 'HORARIO', 'URL ORDEN DE SERVICIO'
    ];

    noBlancos_General.forEach(field => {
        if (isBlank(rowData[headerMap[field]])) {
            if (field === 'DIRECCION') {
                faltantesGeneral.push('DIRECCION (Y DEPTO/OFICINA SI CORRESPONDE)');
            } else {
                faltantesGeneral.push(field);
            }
        }
    });
    const pensionAlimentosIndex = headerMap["PENSION DE ALIMENTOS"];
    noFaltantes_General.forEach(field => {
      if (field === 'AREA DE GESTION' && calidad === 'ADMINISTRACION DE FONDOS') {
        return; // Skip
      }
      if (field === 'FOLIO PENSION') {
        const pensionValue = (pensionAlimentosIndex !== -1 && rowData[pensionAlimentosIndex]) ? String(rowData[pensionAlimentosIndex]).trim().toUpperCase() : '';
        if (pensionValue === 'SI' && isMissing(rowData[headerMap[field]], field)) {
          faltantesGeneral.push(field);
        }
      }
      else if (isMissing(rowData[headerMap[field]], field)) {
        faltantesGeneral.push(field);
      }
    });

    // --- VALIDACIÓN ESPECÍFICA POR CALIDAD ---
    const calidadPrincipal = String(calidad).startsWith("PRESTADORES DE SERVICIO") ? "PRESTADORES DE SERVICIO" : calidad;
    switch (calidadPrincipal) {
      case "CONTRATA":
      case "PLANTA":
      case "PLANTA SUPLENCIA":
        ['DECRETO', 'GRADO', 'ESCALAFON', 'CENTRO DE COSTO', 'AREA DE GESTION'].forEach(f => {
            if (isMissing(rowData[headerMap[f]], f)) {
                faltantesGeneral.push(f);
                faltantesContratos.push(f);
            }
        });
        if (isMissing(rowData[headerMap['FUNCION']], 'FUNCION') && isMissing(rowData[headerMap['FUNCION PRIMER SEMESTRE']], 'FUNCION PRIMER SEMESTRE') && isMissing(rowData[headerMap['FUNCION SEGUNDO SEMESTRE']], 'FUNCION SEGUNDO SEMESTRE')) {
            faltantesGeneral.push("FUNCION (Anual o Semestral)");
            faltantesContratos.push("FUNCION (Anual o Semestral)");
        }
        break;
        
      case "HONORARIOS SUMA ALZADA":
        if (isMissing(rowData[headerMap['PERFIL']], 'PERFIL')) {
            faltantesGeneral.push("PERFIL");
            faltantesContratos.push("PERFIL");
        }
        ['CENTRO DE COSTO', 'AREA DE GESTION'].forEach(f => {
            if (isMissing(rowData[headerMap[f]], f)) {
                faltantesGeneral.push(f);
                faltantesContratos.push(f);
            }
        });
        break;
        case "CODIGO DEL TRABAJO":
        // 1. Validar campos obligatorios de ubicación (Igual que el resto)
        ['CENTRO DE COSTO', 'AREA DE GESTION'].forEach(f => {
            if (isMissing(rowData[headerMap[f]], f)) {
                faltantesGeneral.push(f);
                faltantesContratos.push(f);
            }
        });

        // 2. LÓGICA CRÍTICA: Validar Función (Anual o Semestral)
        // Verificamos si existe la función Anual
        const hayFuncionAnual = !isMissing(rowData[headerMap['FUNCION']], 'FUNCION');
        
        // Verificamos si existen AMBOS semestres
        const haySemestre1 = !isMissing(rowData[headerMap['FUNCION PRIMER SEMESTRE']], 'FUNCION PRIMER SEMESTRE');
        const haySemestre2 = !isMissing(rowData[headerMap['FUNCION SEGUNDO SEMESTRE']], 'FUNCION SEGUNDO SEMESTRE');

        // Si NO hay anual Y tampoco está completo el ciclo semestral (faltan ambos o falta uno)
        if (!hayFuncionAnual && (!haySemestre1 || !haySemestre2)) {
             const mensajeError = "FUNCION (Debe indicar Anual o ambos Semestres)";
             faltantesGeneral.push(mensajeError);
             faltantesContratos.push(mensajeError);
        }
        break;
      case "ADMINISTRACION DE FONDOS":
        ['NOMBRE DE CONVENIO', 'CARGO', 'FUNCION', 'N CUENTA'].forEach(f => {
            if (isMissing(rowData[headerMap[f]], f)) {
                faltantesGeneral.push(f);
                if (f !== 'N CUENTA') {
                    faltantesContratos.push(f);
                }
            }
        });
        break;

      case "PRESTADORES DE SERVICIO":
      case "PLANTA CEMENTERIO":
        const progAnual = rowData[headerMap['PROGRAMA']], funcAnual = rowData[headerMap['FUNCION']];
        const sem1Completo = !isMissing(rowData[headerMap['PROGRAMA PRIMER SEMESTRE']]) && !isMissing(rowData[headerMap['FUNCION PRIMER SEMESTRE']]);
        const sem2Completo = !isMissing(rowData[headerMap['PROGRAMA SEGUNDO SEMESTRE']]) && !isMissing(rowData[headerMap['FUNCION SEGUNDO SEMESTRE']]);
        if (!(!isMissing(progAnual) && !isMissing(funcAnual)) && !sem1Completo && !sem2Completo) {
            faltantesGeneral.push("Par PROGRAMA/FUNCION (Anual o Semestral)");
            faltantesContratos.push("Par PROGRAMA/FUNCION (Anual o Semestral)");
        }
        ['CENTRO DE COSTO', 'AREA DE GESTION'].forEach(f => {
            if (isMissing(rowData[headerMap[f]], f)) {
                faltantesGeneral.push(f);
                faltantesContratos.push(f);
            }
        });
        break;
        
      default:
        ['CENTRO DE COSTO', 'AREA DE GESTION'].forEach(f => {
            if (isMissing(rowData[headerMap[f]], f)) {
                faltantesGeneral.push(f);
                faltantesContratos.push(f);
            }
        });
        break;
    }
    
    const estadoContratos = faltantesContratos.length === 0 ? "LISTO" : "PENDIENTE";
    const textoFaltanteContratos = faltantesContratos.length > 0 ? "Falta: " + [...new Set(faltantesContratos)].join(', ') : "";
    const estadoGeneral = faltantesGeneral.length === 0 ? "INFORMACION COMPLETA" : "INFORMACION INCOMPLETA";
    const textoFaltanteGeneral = faltantesGeneral.length > 0 ? "Falta: " + [...new Set(faltantesGeneral)].join(', ') : "";

    return [estadoGeneral, textoFaltanteGeneral, estadoContratos, textoFaltanteContratos];
  });

  if (results.length > 0) {
    if (indices.estadoGeneral !== -1) sheet.getRange(2, indices.estadoGeneral + 1, results.length, 1).setValues(results.map(r => [r[0]]));
    if (indices.faltanteGeneral !== -1) sheet.getRange(2, indices.faltanteGeneral + 1, results.length, 1).setValues(results.map(r => [r[1]]));
    if (indices.estadoContratos !== -1) sheet.getRange(2, indices.estadoContratos + 1, results.length, 1).setValues(results.map(r => [r[2]]));
    if (indices.faltanteContratos !== -1) sheet.getRange(2, indices.faltanteContratos + 1, results.length, 1).setValues(results.map(r => [r[3]]));
  }
}

// =================================================================================
// --- FUNCIONES AUXILIARES DE VALIDACIÓN ---
// =================================================================================

/**
 * Verifica si un valor es nulo, indefinido o una cadena vacía.
 */
function isBlank(value) {
  return value === null || value === undefined || value === "";
}

/**
 * Verifica si un valor se considera "faltante" (vacío, '-', o 'NO APLICA').
 * Tiene una excepción para "SITUACION MILITAR".
 */
function isMissing(value, fieldName) {
  if (isBlank(value)) return true;
  
  const strValue = String(value).trim().toUpperCase();

  if (fieldName === 'SITUACION MILITAR' && strValue === 'NO APLICA') {
    return false;
  }

  return strValue === '-' || strValue === 'NO APLICA';
}

/**
 * Verifica si un valor se considera "con datos" (no vacío, no '-', no 'NO APLICA').
 */
function tieneDatos(valor) {
  if (valor === null || valor === undefined || valor === '') return false;
  const valorStr = String(valor).trim().toUpperCase();
  return valorStr !== '-' && valorStr !== 'NO APLICA';
}

// --- (Las funciones 'checkBaseFields', 'checkContrataPlantaFields', etc., 
//      fueron eliminadas de tu código original, ya que 'coreUpdateAllStatusOptimized' 
//      parece haberlas reemplazado. Si todavía las usas, también irían aquí.) ---