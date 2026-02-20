// =================================================================================
// --- FUNCIONES PARA EL MÓDULO DE TRASPASO DE PERSONAL ---
// =================================================================================

/**
 * Convierte una fila de datos (array) en un objeto con claves según los encabezados.
 */
function filaAObjeto(fila, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[String(header).trim().toUpperCase()] = fila[index];
  });
  return obj;
}

/**
 * Obtiene los registros listos para ser movidos a la nómina (Dotación).
 */
function obtenerRegistrosParaNomina() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).trim().toUpperCase());

    const registros = data.map(row => filaAObjeto(row, headers));

    const resultados = registros.filter(reg => {
      const estadoTraspaso = String(reg['ESTADO DE TRASPASO A NOMINA'] || '').trim().toUpperCase();
      const estadoInfo = String(reg['ESTADO DE INFORMACION'] || '').trim().toUpperCase();
      const estadoIngreso = String(reg['ESTADO INGRESO'] || '').trim().toUpperCase();

      const pendienteDeNomina = (estadoTraspaso === 'EN ESPERA DE MOVER' || estadoTraspaso === 'MOVIDO A CONTRATOS');

      return estadoInfo === 'INFORMACION COMPLETA'
              && pendienteDeNomina
              && estadoIngreso !== 'INGRESO CANCELADO'; 

    }).map(reg => ({
      id: reg['ID'],
      nombre: reg['NOMBRE COMPLETO'],
      rut: reg['RUT'],
      calidad: reg['CALIDAD CONTRACTUAL'],
      asignadoA: reg['ASIGNADO A']
    }));

    return resultados;
  } catch (e) {
    Logger.log(`Error en obtenerRegistrosParaNomina: ${e.message}`);
    return { error: true, message: e.message };
  }
}

/**
 * Obtiene los registros listos para ser movidos a Contratos.
 * Excluye los que son "CAMBIO DE MODALIDAD".
 */
function obtenerRegistrosParaContratos() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).trim().toUpperCase());

    const registros = data.map(row => filaAObjeto(row, headers));

    const resultados = registros.filter(reg => {
      const estadoTraspaso = String(reg['ESTADO DE TRASPASO A NOMINA'] || '').trim().toUpperCase();
      const estadoContratos = String(reg['ESTADO PARA CONTRATOS'] || '').trim().toUpperCase();
      const estadoIngreso = String(reg['ESTADO INGRESO'] || '').trim().toUpperCase();

      const pendienteDeContrato = (estadoTraspaso === 'EN ESPERA DE MOVER' || estadoTraspaso === 'MOVIDO A DOTACION');

      return estadoContratos === 'LISTO'
              && pendienteDeContrato
              && estadoIngreso !== 'INGRESO CANCELADO'
              && estadoIngreso !== 'CAMBIO DE MODALIDAD'; // Exclusión solicitada

    }).map(reg => ({
      id: reg['ID'],
      nombre: reg['NOMBRE COMPLETO'],
      rut: reg['RUT'],
      calidad: reg['CALIDAD CONTRACTUAL'],
      asignadoA: reg['ASIGNADO A']
    }));

    return resultados;
  } catch (e) {
    Logger.log(`Error en obtenerRegistrosParaContratos: ${e.message}`);
    return { error: true, message: e.message };
  }
}


function obtenerEstadoAvanceGlobal() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).trim().toUpperCase());

    const registros = data.map(row => filaAObjeto(row, headers));

    const resultados = [];

    registros.forEach(reg => {
      const estadoIngreso = String(reg['ESTADO INGRESO'] || '').trim().toUpperCase();
      const estadoTraspaso = String(reg['ESTADO DE TRASPASO A NOMINA'] || '').trim().toUpperCase();
      
      // 1. Descartamos los que ya terminaron o están cancelados
      if (estadoTraspaso === 'PROCESO COMPLETO' || estadoIngreso === 'INGRESO CANCELADO') {
        return;
      }

      // 2. Determinamos qué le falta
      let estadoEtiqueta = '';
      let colorEtiqueta = '';
      let accionSugerida = '';

      if (estadoTraspaso === 'EN ESPERA DE MOVER' || estadoTraspaso === '') {
          // Caso especial: Cambio de modalidad solo requiere Dotación
          if (estadoIngreso === 'CAMBIO DE MODALIDAD' || estadoIngreso === 'NUEVO INGRESO - CAMBIO DE MODALIDAD') {
             estadoEtiqueta = 'FALTA DOTACIÓN';
             colorEtiqueta = 'badge-warning';
             accionSugerida = 'Mover a Nómina';
          } else {
             estadoEtiqueta = 'PENDIENTE TOTAL';
             colorEtiqueta = 'badge-danger';
             accionSugerida = 'Mover a Contratos y Nómina';
          }
      } else if (estadoTraspaso === 'MOVIDO A DOTACION') {
          estadoEtiqueta = 'FALTA CONTRATO';
          colorEtiqueta = 'badge-info';
          accionSugerida = 'Mover a Contratos';
      } else if (estadoTraspaso === 'MOVIDO A CONTRATOS') {
          estadoEtiqueta = 'FALTA DOTACIÓN';
          colorEtiqueta = 'badge-primary';
          accionSugerida = 'Mover a Nómina';
      } else {
          // Estado desconocido o intermedio
          estadoEtiqueta = estadoTraspaso;
          colorEtiqueta = 'badge-secondary';
      }

      // 3. Verificamos si tiene datos completos (para saber si está bloqueado por datos)
      const estadoInfo = String(reg['ESTADO DE INFORMACION'] || '').trim().toUpperCase();
      const datosListos = (estadoInfo === 'INFORMACION COMPLETA');

      resultados.push({
        id: reg['ID'],
        nombre: reg['NOMBRE COMPLETO'],
        rut: reg['RUT'],
        calidad: reg['CALIDAD CONTRACTUAL'],
        estadoEtiqueta: estadoEtiqueta,
        colorEtiqueta: colorEtiqueta,
        accion: accionSugerida,
        datosListos: datosListos // True si puede moverse, False si le faltan datos
      });
    });

    return resultados;

  } catch (e) {
    Logger.log("Error en obtenerEstadoAvanceGlobal: " + e.message);
    return { error: true, message: e.message };
  }
}


/**
 * Mueve un empleado del REGISTRO_MAESTRO a la hoja de DOTACIÓN.
 */
function ejecutarTraspaso(idEmpleado) {
  try {
    const ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
    const registroMaestroSheet = ssOrigen.getSheetByName(REGISTRO_MAESTRO_SHEET);
    
    // Obtenemos datos y headers para mapear por nombre
    const data = registroMaestroSheet.getDataRange().getValues();
    const rawHeaders = data.shift(); // Quitamos headers del array de datos
    const headers = rawHeaders.map(h => String(h).trim().toUpperCase());
    
    // 1. Encontrar índice de columna ID y la fila correspondiente
    const idColIndex = headers.indexOf('ID');
    if (idColIndex === -1) throw new Error("Columna ID no encontrada.");

    const dataRowIndex = data.findIndex(row => String(row[idColIndex]) === String(idEmpleado));
    if (dataRowIndex === -1) throw new Error(`Empleado ID ${idEmpleado} no encontrado.`);

    // 2. Convertir a objeto para trabajar cómodamente
    const registroOrigen = filaAObjeto(data[dataRowIndex], headers);
    const calidadContractual = String(registroOrigen['CALIDAD CONTRACTUAL'] || '').trim().toUpperCase();

    if (!calidadContractual) throw new Error("Sin Calidad Contractual.");

    // --- LÓGICA DE DESTINO ---
    let nombreHojaBuscada = calidadContractual;
    if (calidadContractual.includes('BANDA LOS MENAS')) nombreHojaBuscada = 'BANDA LOS MENAS';
    else if (calidadContractual.includes('CAMPEONES PARA COQUIMBO')) nombreHojaBuscada = 'CAMPEONES PARA COQUIMBO';
    else if (calidadContractual.includes('SIN MARCAJE')) nombreHojaBuscada = 'SIN MARCAJE';
    else if (calidadContractual.startsWith('PRESTADORES DE SERVICIO')) nombreHojaBuscada = 'PRESTADORES DE SERVICIO';

    const ssDestino = SpreadsheetApp.openById(ID_HOJA_DOTACION);
    const hojaDestino = ssDestino.getSheetByName(nombreHojaBuscada);
    if (!hojaDestino) throw new Error(`Hoja destino '${nombreHojaBuscada}' no encontrada.`);

    // --- CONSTRUCCIÓN DE LA NUEVA FILA ---
    const destinoHeaders = hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).getValues()[0];
    const CAMPOS_EXCLUIDOS = ['EDAD', 'CENTRO DE COSTO']; // Campos que NO se deben copiar

    const nuevaFila = destinoHeaders.map(headerDestino => {
        const headerLimpio = String(headerDestino).trim().toUpperCase();
        
        // Exclusión explícita
        if (CAMPOS_EXCLUIDOS.includes(headerLimpio)) return '';

        let valor = registroOrigen[headerLimpio];

        // Fallback para nombre compuesto si no existe la columna
        if (valor === undefined && headerLimpio === 'NOMBRE COMPLETO') {
             const n = registroOrigen['NOMBRES'] || '';
             const a = registroOrigen['APELLIDOS'] || '';
             valor = `${a} ${n}`.trim();
        }

        // Lógica de Estado Ingreso
        if (headerLimpio === 'ESTADO INGRESO' && String(valor).trim() === 'CAMBIO DE MODALIDAD') {
            return 'NUEVO INGRESO - CAMBIO DE MODALIDAD';
        }
        
        // Formateo de RUT
        if (headerLimpio === 'RUT') {
            return formatearRut(valor);
        }

        return valor !== undefined ? valor : '';
    });

    hojaDestino.appendRow(nuevaFila);

    // --- ACTUALIZACIÓN DE ESTADO EN ORIGEN ---
    const estadoActual = String(registroOrigen['ESTADO DE TRASPASO A NOMINA'] || '').trim().toUpperCase();
    const estadoIngresoVal = String(registroOrigen['ESTADO INGRESO'] || '').trim().toUpperCase();
    
    let nuevoEstado = 'MOVIDO A DOTACION'; 

    if (estadoActual === 'MOVIDO A CONTRATOS') {
        nuevoEstado = 'PROCESO COMPLETO';
    } else if (estadoIngresoVal === 'CAMBIO DE MODALIDAD') {
        // Los cambios de modalidad terminan al pasar a dotación (no pasan por contratos)
        nuevoEstado = 'PROCESO COMPLETO';
    } else if (estadoActual === 'PROCESO COMPLETO') {
        nuevoEstado = 'PROCESO COMPLETO';
    }

    const colEstadoIndex = headers.indexOf('ESTADO DE TRASPASO A NOMINA');
    if (colEstadoIndex > -1) {
        // +2: 1 por header + 1 por base 1 de sheets
        registroMaestroSheet.getRange(dataRowIndex + 2, colEstadoIndex + 1).setValue(nuevoEstado);
    }
    
    Logger.log(`Empleado ID ${idEmpleado} movido a '${nombreHojaBuscada}'. Estado: ${nuevoEstado}`);

    // Notificación
    try {
        if (typeof EditordeDotacinUltimo !== 'undefined' && typeof EditordeDotacinUltimo.dispararNotificacionExterna === 'function') {
            EditordeDotacinUltimo.dispararNotificacionExterna();
        }
    } catch (e) { Logger.log("Error notificación dotación: " + e); }

    return { success: true, message: `Empleado traspasado a '${nombreHojaBuscada}'. Estado: ${nuevoEstado}` };

  } catch (e) {
    Logger.log(`Error ejecutarTraspaso: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Mueve un empleado del REGISTRO_MAESTRO a la hoja de Contratos Pendientes.
 */
function ejecutarTraspasoAContratos(idEmpleado) {
  const ID_HOJA_CONTRATOS = "1K7HRYwcf4h69so0YNyuSmydp-zNqWw9Shp9vFdp7ySA";
  const NOMBRE_HOJA_CONTRATOS = "Contratos_Pendientes";

  try {
    const ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
    const registroMaestroSheet = ssOrigen.getSheetByName(REGISTRO_MAESTRO_SHEET);
    
    const data = registroMaestroSheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).trim().toUpperCase());

    const idColIndex = headers.indexOf('ID');
    if (idColIndex === -1) throw new Error("Columna ID no encontrada.");

    const dataRowIndex = data.findIndex(row => String(row[idColIndex]) === String(idEmpleado));
    if (dataRowIndex === -1) throw new Error(`Empleado ID ${idEmpleado} no encontrado.`);

    // Convertimos a Objeto
    const registroOrigen = filaAObjeto(data[dataRowIndex], headers);

    // --- LÓGICA DE PRIORIZACIÓN DE PROGRAMA Y FUNCIÓN (MODIFICADO) ---
    let programaActivo = "";
    let funcionActiva = "";

    // 1. Determinar Semestre Actual
    const hoy = new Date();
    const esPrimerSemestre = hoy.getMonth() < 6; // Enero(0) a Junio(5)

    // 2. Helper interno para validar datos reales (excluye vacíos, guiones y 'NO APLICA')
    const tieneInfo = (val) => {
      if (!val) return false;
      const s = String(val).trim().toUpperCase();
      return s !== '' && s !== '-' && s !== 'NO APLICA';
    };

    // 3. Obtener valores
    const prog1 = registroOrigen['PROGRAMA PRIMER SEMESTRE'];
    const func1 = registroOrigen['FUNCION PRIMER SEMESTRE'];
    const prog2 = registroOrigen['PROGRAMA SEGUNDO SEMESTRE'];
    const func2 = registroOrigen['FUNCION SEGUNDO SEMESTRE'];
    const progAnual = registroOrigen['PROGRAMA'];
    const funcAnual = registroOrigen['FUNCION'];

    // 4. Aplicar prioridad según semestre actual
    if (esPrimerSemestre) {
        // Prioridad: Semestre 1 -> Semestre 2 -> Anual
        programaActivo = tieneInfo(prog1) ? prog1 : (tieneInfo(prog2) ? prog2 : progAnual);
        funcionActiva = tieneInfo(func1) ? func1 : (tieneInfo(func2) ? func2 : funcAnual);
    } else {
        // Prioridad: Semestre 2 -> Semestre 1 -> Anual
        programaActivo = tieneInfo(prog2) ? prog2 : (tieneInfo(prog1) ? prog1 : progAnual);
        funcionActiva = tieneInfo(func2) ? func2 : (tieneInfo(func1) ? func1 : funcAnual);
    }
    // ---------------------------------------------------------------

    const ssDestino = SpreadsheetApp.openById(ID_HOJA_CONTRATOS);
    const hojaDestino = ssDestino.getSheetByName(NOMBRE_HOJA_CONTRATOS);
    if (!hojaDestino) throw new Error(`Hoja destino no encontrada.`);

    // Lógica para ID Solicitud
    const destinoHeaders = hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).getValues()[0];
    const idSolIndex = destinoHeaders.indexOf('ID_SOLICITUD');
    let ultimoNumero = 0;
    if (hojaDestino.getLastRow() > 1 && idSolIndex > -1) {
       const val = hojaDestino.getRange(hojaDestino.getLastRow(), idSolIndex + 1).getValue();
       if (String(val).includes('-')) ultimoNumero = parseInt(String(val).split('-')[1]) || 0;
    }
    const nuevoIdCorrelativo = "SOL-" + (ultimoNumero + 1);

    const nuevaFila = destinoHeaders.map(headerDestino => {
        const headerLimpio = String(headerDestino).trim().toUpperCase();
        const valorOriginal = registroOrigen[headerLimpio];

        switch (headerLimpio) {
            case 'ID_SOLICITUD': return nuevoIdCorrelativo;
            case 'TIMESTAMP': return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            
            case 'TIPOCONTRATO': 
               if (String(registroOrigen['ESTADO INGRESO']).trim() === 'CAMBIO DE MODALIDAD') return 'NUEVO INGRESO - CAMBIO DE MODALIDAD';
               return 'NUEVO INGRESO';
            
            case 'RUT': return formatearRut(valorOriginal);
            case 'ESTADO': return 'Pendiente';
            
            case 'PROGRAMA':
               if (registroOrigen['CALIDAD CONTRACTUAL'] === 'HONORARIOS SUMA ALZADA') return registroOrigen['PERFIL'] || '';
               if (registroOrigen['CALIDAD CONTRACTUAL'] === 'ADMINISTRACION DE FONDOS') return registroOrigen['NOMBRE DE CONVENIO'] || '';
               return programaActivo || '';
            
            case 'FUNCION':
               if (registroOrigen['CALIDAD CONTRACTUAL'] === 'HONORARIOS SUMA ALZADA') return registroOrigen['DESCRIPCION PERFIL'] || '';
               // Lógica especial para CÓDIGO DEL TRABAJO y ADMINISTRACION DE FONDOS (Usan el campo FUNCION directo)
               if (registroOrigen['CALIDAD CONTRACTUAL'] === 'CODIGO DEL TRABAJO') return registroOrigen['FUNCION'] || '';
               if (registroOrigen['CALIDAD CONTRACTUAL'] === 'ADMINISTRACION DE FONDOS') return registroOrigen['FUNCION'] || '';
               
               // Para prestadores, usamos la lógica semestral calculada arriba
               return funcionActiva || '';
            
            case 'CENTRO DE COSTO':
               if (valorOriginal && String(valorOriginal).length === 6) return String(valorOriginal).substring(2, 4);
               return valorOriginal;
               
            case 'DEPARTAMENTO':
            case 'OFICINA':
            case 'DIRECCION':
               return valorOriginal ? String(valorOriginal).replace(/^\(\d+\)\s*/, '') : ""; 
               
            case 'AREA DE GESTION': 
               if (valorOriginal) {
                   const match = String(valorOriginal).match(/\((\d+)\)/);
                   return match ? match[1] : valorOriginal;
               }
               return "";
            
            case 'NOMBRE_AREA_GESTION':
               return registroOrigen['AREA DE GESTION'] ? String(registroOrigen['AREA DE GESTION']).replace(/^\(\d+\)\s*/, '') : "";

            case 'FECHA INGRESO':
            case 'FECHA DE TERMINO':
            case 'FECHA ORNS/NI': 
               if (valorOriginal instanceof Date) return Utilities.formatDate(valorOriginal, Session.getScriptTimeZone(), "d 'de' MMMM 'de' yyyy");
               return valorOriginal;

            case 'MONTO EN PALABRAS': return numeroAJson(registroOrigen['MONTO PROPORCIONAL']);
            case 'MONTO_BRUTO_EN_PALABRAS': return numeroAJson(registroOrigen['MONTO BRUTO']);
            
            // Limpieza de campos intermedios
            case 'PROGRAMA PRIMER SEMESTRE':
            case 'FUNCION PRIMER SEMESTRE':
            case 'PROGRAMA SEGUNDO SEMESTRE':
            case 'FUNCION SEGUNDO SEMESTRE': return "";
            
            default:
               if (headerLimpio === 'ASIGNADO A' && typeof valorOriginal === 'string') return valorOriginal.toLowerCase();
               return valorOriginal !== undefined ? valorOriginal : "";
        }
    });

    hojaDestino.appendRow(nuevaFila);

    // --- ACTUALIZACIÓN ESTADO ---
    const estadoActual = String(registroOrigen['ESTADO DE TRASPASO A NOMINA'] || '').trim().toUpperCase();
    let nuevoEstado = 'MOVIDO A CONTRATOS';

    if (estadoActual === 'MOVIDO A DOTACION') {
         nuevoEstado = 'PROCESO COMPLETO';
    } else if (estadoActual === 'PROCESO COMPLETO') {
         nuevoEstado = 'PROCESO COMPLETO';
    }

    const colEstadoIndex = headers.indexOf('ESTADO DE TRASPASO A NOMINA');
    if (colEstadoIndex > -1) {
        registroMaestroSheet.getRange(dataRowIndex + 2, colEstadoIndex + 1).setValue(nuevoEstado);
    }

    // Notificación
    let notificacionMsg = "";
    try {
        const asignadoA = registroOrigen['ASIGNADO A'] ? String(registroOrigen['ASIGNADO A']).toLowerCase() : '';
        if (asignadoA && typeof SistemadeGestindeContratos !== 'undefined' && typeof SistemadeGestindeContratos.notificarNuevoIngreso === 'function') {
            SistemadeGestindeContratos.notificarNuevoIngreso(asignadoA);
            notificacionMsg = " (Notificación enviada)";
        } else {
            notificacionMsg = " (Sin notificación)";
        }
    } catch (e) { notificacionMsg = " (Error notif)"; }

    return { success: true, message: `Empleado traspasado a Contratos. Nuevo estado: ${nuevoEstado}.${notificacionMsg}` };

  } catch (e) {
    Logger.log(`Error en ejecutarTraspasoAContratos: ${e.message}`);
    return { success: false, message: e.message };
  }
}

// --- HELPER: FORMATEO DE RUT SIN PUNTOS (SOLO GUION) ---
function formatearRut(rut) {
  if (!rut) return '';
  const valorLimpio = String(rut).replace(/[^0-9kK]/g, '').toUpperCase();
  if (valorLimpio.length < 2) return rut;
  
  const dv = valorLimpio.slice(-1);
  const cuerpo = valorLimpio.slice(0, -1);
  
  // SIN PUNTOS, SOLO GUION
  return cuerpo + "-" + dv;
}

// --- HELPER: MONTO A PALABRAS (Requerido por la lógica de contratos) ---
function numeroAJson(numero) {
  if (!numero || isNaN(numero)) return "";
  return numero; 
}