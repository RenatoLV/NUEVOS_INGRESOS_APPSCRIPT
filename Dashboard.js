// =================================================================================
// --- FUNCIONES PARA EL DASHBOARD Y VISTAS DE LISTA ---
// =================================================================================

/**
 * Obtiene las estadísticas principales para el Dashboard.
 */
function obtenerDatosDashboard() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const datos = hoja.getDataRange().getValues();
    const headers = datos.shift();
    
    const indices = {
      rut: headers.indexOf("RUT"),
      estadoInfo: headers.indexOf("ESTADO DE INFORMACION"),
      estadoTraspaso: headers.indexOf("ESTADO DE TRASPASO A NOMINA"),
      estadoIngreso: headers.indexOf("ESTADO INGRESO"),
      id: headers.indexOf('ID'),
      nombreCompleto: headers.indexOf('NOMBRE COMPLETO'),
      fechaIngreso: headers.indexOf('FECHA INGRESO'),
      calidad: headers.indexOf('CALIDAD CONTRACTUAL')
    };

    if (Object.values(indices).some(i => i === -1)) {
        throw new Error("Faltan columnas requeridas en la hoja.");
    }
    
    let totalIngresos = 0, incompletos = 0, pendientesContrato = 0, pendientesNomina = 0, cancelados = 0;
    const conteoPorCalidad = {};
    const conteoPorMesDetallado = {};
    
    const primeraFecha = datos.length > 0 && datos[0][indices.fechaIngreso] ? datos[0][indices.fechaIngreso] : null;

    datos.forEach(fila => {
      if (fila[indices.rut]) totalIngresos++;
      
      const esCancelado = String(fila[indices.estadoIngreso] || '').toUpperCase() === 'INGRESO CANCELADO';
      const calidad = fila[indices.calidad];
      let fechaIngreso = fila[indices.fechaIngreso]; 

      // --- CORRECCIÓN CRÍTICA DE FECHAS ---
      // Si no es un objeto Date pero es un String con formato de fecha, intentamos convertirlo.
      if (!(fechaIngreso instanceof Date) && typeof fechaIngreso === 'string' && fechaIngreso.length >= 8) {
          // Intentamos parsear strings tipo "2026-01-06" o "06/01/2026"
          const parsed = Date.parse(fechaIngreso.replace(/-/g, '/')); 
          if (!isNaN(parsed)) {
              fechaIngreso = new Date(parsed);
          } else {
              // Intento secundario para formato ISO puro si falló el anterior
              const isoParsed = Date.parse(fechaIngreso);
              if(!isNaN(isoParsed)) {
                  fechaIngreso = new Date(isoParsed);
              }
          }
      }

      if (esCancelado) {
        cancelados++;
      } else {
        if (String(fila[indices.estadoInfo]).toUpperCase() === 'INFORMACION INCOMPLETA') incompletos++;
        const estadoTraspaso = String(fila[indices.estadoTraspaso]).toUpperCase();
        if (estadoTraspaso === 'EN ESPERA DE MOVER') pendientesContrato++;
        else if (estadoTraspaso === 'MOVIDO A CONTRATOS') pendientesNomina++;

        if (calidad) {
          conteoPorCalidad[calidad] = (conteoPorCalidad[calidad] || 0) + 1;
        }

        // Ahora validamos con la fecha ya procesada/corregida
        if (fechaIngreso instanceof Date && !isNaN(fechaIngreso.getTime())) {
          // Usamos el TimeZone del Script para asegurar que caiga en el mes correcto
          const mesAnio = Utilities.formatDate(fechaIngreso, Session.getScriptTimeZone(), "yyyy-MM");
          if (!conteoPorMesDetallado[mesAnio]) {
            conteoPorMesDetallado[mesAnio] = {};
          }
          conteoPorMesDetallado[mesAnio][calidad] = (conteoPorMesDetallado[mesAnio][calidad] || 0) + 1;
        }
      }
    });

    const ultimosIngresos = datos.slice(-5).reverse().map(fila => {
      let fecha = fila[indices.fechaIngreso];
      // Aplicamos la misma corrección para la lista visual
      if (!(fecha instanceof Date) && typeof fecha === 'string') {
          const p = Date.parse(fecha);
          if(!isNaN(p)) fecha = new Date(p);
      }

      if (fecha instanceof Date) {
        fecha = fecha.toISOString();
      }
      return {
        ID: fila[indices.id],
        NOMBRE_COMPLETO: fila[indices.nombreCompleto],
        FECHA_INGRESO: fecha,
        RUT: fila[indices.rut],
        CALIDAD_CONTRACTUAL: fila[indices.calidad],
        ESTADO_INGRESO: fila[indices.estadoIngreso] 
      };
    });

    const objetoDeRetorno = {
      stats: {
        totalIngresos, incompletos, pendientesContrato, pendientesNomina, cancelados
      },
      ultimosIngresos: ultimosIngresos,
      conteoPorCalidad: conteoPorCalidad,
      conteoPorMesDetallado: conteoPorMesDetallado,
      primeraFecha: primeraFecha instanceof Date ? primeraFecha.toISOString() : null
    };

    return JSON.stringify(objetoDeRetorno);
    
  } catch (e) {
    Logger.log("Error en obtenerDatosDashboard: " + e.message);
    return JSON.stringify({ error: true, message: e.message });
  }
}

/**
 * Obtiene todos los registros con "INFORMACION INCOMPLETA".
 */
function obtenerRegistrosIncompletos() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    if (!hoja) {
      throw new Error(`La hoja de cálculo '${REGISTRO_MAESTRO_SHEET}' no fue encontrada.`);
    }

    const datos = hoja.getDataRange().getValues();
    const headers = datos.shift(); 

    const indices = {
        id: headers.indexOf("ID"),
        nombreCompleto: headers.indexOf("NOMBRE COMPLETO"),
        rut: headers.indexOf("RUT"),
        calidadContractual: headers.indexOf("CALIDAD CONTRACTUAL"),
        estadoGeneral: headers.indexOf("ESTADO DE INFORMACION"),
        faltanteGeneral: headers.indexOf("INFORMACION FALTANTE"),
        estadoContratos: headers.indexOf("ESTADO PARA CONTRATOS"),
        faltanteContratos: headers.indexOf("FALTANTE_CONTRATOS")
    };

    for (const key in indices) {
        if (indices[key] === -1) {
            throw new Error(`No se encontró la columna requerida: '${key.toUpperCase()}' en la hoja.`);
        }
    }

    const resultados = [];
    for (let i = datos.length - 1; i >= 0; i--) {
      const fila = datos[i];
      const estadoGeneralActual = fila[indices.estadoGeneral];

      if (typeof estadoGeneralActual === 'string' && estadoGeneralActual.trim().toUpperCase() === 'INFORMACION INCOMPLETA') {
        const registro = {
            ID: fila[indices.id],
            'NOMBRE COMPLETO': fila[indices.nombreCompleto],
            RUT: fila[indices.rut],
            'CALIDAD CONTRACTUAL': fila[indices.calidadContractual],
            estadoContratos: fila[indices.estadoContratos] || 'PENDIENTE',
            faltanteContratos: fila[indices.faltanteContratos] || '',
            estadoGeneral: estadoGeneralActual,
            faltanteGeneral: fila[indices.faltanteGeneral] || ''
        };
        resultados.push(registro);
      }
    }

    Logger.log(`obtenerRegistrosIncompletos: Encontrados ${resultados.length} registros incompletos.`);
    return resultados;

  } catch (error) {
    console.error("Error en obtenerRegistrosIncompletos:", error.message, error.stack);
    Logger.log(`Error en obtenerRegistrosIncompletos: ${error.message}`);
    return { error: true, message: error.message };
  }
}

/**
 * Obtiene todos los registros de la hoja REGISTRO_MAESTRO.
 */
function obtenerTodosLosIngresos() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const datos = hoja.getDataRange().getValues();
    const headers = datos.shift(); 

    const resultados = datos.map(fila => {
      const registro = {};
      headers.forEach((header, index) => {
        let valor = fila[index];
        if (valor instanceof Date) {
          registro[header] = valor.toISOString();
        } else {
          registro[header] = valor;
        }
      });
      return registro;
    });

    return resultados;

  } catch (error) {
    Logger.log("Error en obtenerTodosLosIngresos:", error.message, error.stack);
    return { error: true, message: error.message };
  }
}

function verificarIntegridadDeDatos() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRO_MAESTRO_SHEET);
    const lastRow = sheet.getLastRow();
    // Obtenemos el ID de la última fila para mayor seguridad
    const lastId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0; 
    
    return {
      totalFilas: lastRow,
      ultimoId: lastId
    };
  } catch (e) {
    return { error: true };
  }
}