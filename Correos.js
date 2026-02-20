// =================================================================================
// --- FUNCIONES PARA EL MÓDULO DE CORREOS ---
// =================================================================================

function obtenerListaCC() {
  return [
    { nombre: "Juan Carlos Alegría Barraza", email: "jalegriab@municoquimbo.cl" },
    { nombre: "Mario Barrios Almendares", email: "mbarriosa@municoquimbo.cl" },
    { nombre: "Jocelyn Godoy Arancibia", email: "jocelyn.godoy@municoquimbo.cl" },
    { nombre: "Bastian Figueroa Baez", email: "bastian.figueroa@municoquimbo.cl" },
    { nombre: "María José Mora Styl", email: "mariajose.mora@municoquimbo.cl" },
    { nombre: "Nicole Espinoza", email: "nespinozav@municoquimbo.cl" },
    { nombre: "Yanina Mellado Rojas", email: "ymellador@municoquimbo.cl" },
    { nombre: "Aaron Cortés Ramos", email: "aaroncortes@municoquimbo.cl" },
    { nombre: "Isabel Ramos Ramos", email: "isabelramos@municoquimbo.cl" },
    { nombre: "Alejandra Romero Bugueño", email: "aromerob@municoquimbo.cl" },
    { nombre: "Claudia Encalada Muñoz", email: "claudiaencalada@municoquimbo.cl" },
    { nombre: "Fiscalizacion RRHH", email: "fiscalizacion_rrhh@municoquimbo.cl" },
    { nombre: "Angela Barraza Garcia", email: "abarraza@municoquimbo.cl" }
  ];
}

function obtenerEmpleadosParaBienvenida() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTRO_MAESTRO');
    if (!hoja) throw new Error("Hoja REGISTRO_MAESTRO no encontrada.");
    const datos = hoja.getDataRange().getValues();
    const headers = datos.shift();
    
    const indices = {
      id: headers.indexOf("ID"),
      nombreCompleto: headers.indexOf("NOMBRE COMPLETO"),
      nombres: headers.indexOf("NOMBRES"),
      apellidos: headers.indexOf("APELLIDOS"),
      rut: headers.indexOf("RUT"),
      calidad: headers.indexOf("CALIDAD CONTRACTUAL"),
      estadoTraspaso: headers.indexOf("ESTADO DE TRASPASO A NOMINA"),
      correoEnviado: headers.indexOf("CORREO BIENVENIDA ENVIADO"),
      email: headers.indexOf("CORREO ELECTRONICO"),
      fechaIngreso: headers.indexOf("FECHA INGRESO"),
      urlOrden: headers.indexOf("URL ORDEN DE SERVICIO"),
      // --- CAMPOS REQUERIDOS ADICIONALES ---
      numOrden: headers.indexOf("NUMERO ORDEN DE SERVICIO"),
      direccion: headers.indexOf("DIRECCION"),
      departamento: headers.indexOf("DEPARTAMENTO"),
      oficina: headers.indexOf("OFICINA"),
      areaGestion: headers.indexOf("AREA DE GESTION"),
      codigo: headers.indexOf("CODIGO")
    };

    if (indices.id === -1) throw new Error("Faltan columnas críticas.");

    // Helper: Verifica que haya dato (no vacío). "NO APLICA" o "-" se consideran datos válidos (listos).
    const tieneDato = (val) => val && String(val).trim() !== '';
    // Helper: URL válida
    const tieneUrl = (val) => val && String(val).trim().length > 10 && String(val).toLowerCase().startsWith('http');

    return datos.filter(fila => {
      const estado = String(fila[indices.estadoTraspaso] || '').toUpperCase().trim();
      const movido = (estado === 'MOVIDO A NOMINA' || estado === 'MOVIDO A DOTACION' || estado === 'PROCESO COMPLETO');
      const noEnviado = !fila[indices.correoEnviado];

      if (!movido || !noEnviado) return false;

      // Validación estricta de campos solicitada
      const camposListos = 
           tieneDato(fila[indices.numOrden]) &&
           tieneUrl(fila[indices.urlOrden]) &&
           tieneDato(fila[indices.rut]) &&
           tieneDato(fila[indices.calidad]) &&
           tieneDato(fila[indices.email]) &&
           tieneDato(fila[indices.areaGestion]) &&
           tieneDato(fila[indices.codigo]) &&
           tieneDato(fila[indices.fechaIngreso]) &&
           // Ubicación (Deben tener algo, aunque sea NO APLICA o -)
           tieneDato(fila[indices.direccion]) &&
           tieneDato(fila[indices.departamento]) &&
           tieneDato(fila[indices.oficina]) &&
           // Nombre
           (tieneDato(fila[indices.nombreCompleto]) || (tieneDato(fila[indices.nombres]) && tieneDato(fila[indices.apellidos])));

      return camposListos;

    }).map(fila => {
      let fecha = fila[indices.fechaIngreso];
      if (fecha instanceof Date) fecha = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      const rawUrl = String(fila[indices.urlOrden] || '').trim();
      const tieneOrden = rawUrl.length > 20 && rawUrl.toLowerCase().startsWith('http');
      
      let nombreDisplay = fila[indices.nombreCompleto];
      const n = String(fila[indices.nombres] || '').trim();
      const a = String(fila[indices.apellidos] || '').trim();
      if (n && a) nombreDisplay = `${n} ${a}`;

      return {
        id: fila[indices.id],
        nombre: toTitleCase(nombreDisplay),
        rut: fila[indices.rut],
        calidad: fila[indices.calidad],
        email: String(fila[indices.email] || '').trim(),
        fechaIngreso: fecha,
        urlOrden: tieneOrden ? rawUrl : null
      };
    });
  } catch (e) {
    return { error: true, message: e.message };
  }
}

function toTitleCase(str) {
  if (!str) return "";
  return String(str).toLowerCase().replace(/(^|\s)\S/g, l => l.toUpperCase());
}

function generarContenidoCorreo(idEmpleado, tipoFormato) {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTRO_MAESTRO');
    const todosLosDatos = hoja.getDataRange().getValues();
    const headers = todosLosDatos.shift();
    const idCol = headers.indexOf('ID');
    const fila = todosLosDatos.find(row => String(row[idCol]) === String(idEmpleado));
    if (!fila) return null;

    const registro = {};
    headers.forEach((h, i) => registro[h] = fila[i]);

    // --- CONSTRUCCIÓN DE DATOS ---

    // 1. Ubicación
    let ubicacion = registro['DIRECCION'] || '';
    if (registro['DEPARTAMENTO'] && registro['DEPARTAMENTO'] !== '-' && registro['DEPARTAMENTO'] !== 'NO APLICA') {
        ubicacion += `, dependiente del ${registro['DEPARTAMENTO']}`;
        if (registro['OFICINA'] && registro['OFICINA'] !== '-' && registro['OFICINA'] !== 'NO APLICA') {
             ubicacion += `, oficina ${registro['OFICINA']}`;
        }
    }

    // 2. Nombre
    let nombreC = registro['NOMBRE COMPLETO'];
    if (registro['NOMBRES'] && registro['APELLIDOS']) nombreC = `${registro['NOMBRES']} ${registro['APELLIDOS']}`;
    const nombreSaludo = toTitleCase(registro['NOMBRES'] || nombreC);

    // 3. Función Inteligente
    const funcAnual = String(registro['FUNCION'] || '').trim();
    const funcSem1 = String(registro['FUNCION PRIMER SEMESTRE'] || '').trim();
    const funcSem2 = String(registro['FUNCION SEGUNDO SEMESTRE'] || '').trim();
    
    let funcionParaMostrar = "No especificada";
    const esValida = (s) => s.length > 3 && s !== '-' && s !== 'NO APLICA' && s.toUpperCase() !== 'UNDEFINED';
    
    if (esValida(funcAnual)) funcionParaMostrar = funcAnual;
    else if (esValida(funcSem1)) {
        funcionParaMostrar = funcSem1;
        if (esValida(funcSem2) && funcSem1 !== funcSem2) funcionParaMostrar += ` / ${funcSem2}`;
    } else if (esValida(funcSem2)) funcionParaMostrar = funcSem2;

    // --- CONSTRUCCIÓN DEL TEXTO (TEXTOS DEFINITIVOS) ---
    
    const intro = `Junto con saludarle cordialmente, en nombre del Alcalde de Coquimbo Sr. Alí Manouchehri Moghadam Kashan Lobos, por el presente correo le damos a usted la más cordial bienvenida a la I. Municipalidad de Coquimbo, nos enorgullece poder contar con personas tan talentosas como usted y le deseamos lo mejor en su incorporación, cabe mencionar que estamos a su disposición para ayudarle en lo que necesite.`;

    // Usamos toTitleCase para que la ubicación no salga en mayúsculas agresivas
    const detalle = `En virtud de su incorporación le enviamos a usted copia digital de la orden de servicio N° <strong>{{NUMERO ORDEN DE SERVICIO}}</strong> que da cuenta de su ingreso cuya contratación es bajo la modalidad de <strong>{{CALIDAD CONTRACTUAL}}</strong> y que cumplirá funciones en <strong>${toTitleCase(ubicacion)}</strong> (ID {{CODIGO}} en sistema CAS Chile) a contar del <strong>{{FECHA INGRESO}}</strong>.`;

    const cierre = `Se agradece además al Depto. requirente, que informe el ingreso oportuno del Sr./Srta. <strong>${toTitleCase(nombreC)}</strong>, con el objeto de corroborar el ingreso en la fecha y funciones designadas en la Oficina de Movimiento de Personal como en las diversas Unidades pertenecientes a la Dirección de Recursos Humanos.<br><br>
    En virtud de lo anterior, solicitamos acusar recibo y asimismo se adjunta la orden de servicio N° {{NUMERO ORDEN DE SERVICIO}}, para su conocimiento y cumplimiento, sin otro particular le saluda atentamente a usted.`;

    // Bloque Previred (SOLO FORMATO A)
    let previredBlock = "";
    if (tipoFormato === 'A') {
        previredBlock = `
        <div style="background-color: #fefce8; border-left: 4px solid #facc15; padding: 15px; margin: 20px 0; border-radius: 4px;">
            <p style="margin: 0; color: #854d0e; font-weight: bold; font-size: 14px;">Información Importante (Trabajador Independiente):</p>
            <p style="margin: 5px 0 0 0; color: #713f12; font-size: 13px;">Se solicita ingresar al Link de PREVIRED para conocer derechos y obligaciones:</p>
            <p style="margin-top: 8px;"><a href="https://www.youtube.com/watch?v=GGwUvjP_Gvc" style="color: #ca8a04; font-weight: bold; text-decoration: underline;">Ver Video Explicativo</a></p>
        </div>`;
    }

    // Link Orden Adjunta (SI EXISTE URL)
    let linkOrdenBlock = "";
    const urlOrden = registro['URL ORDEN DE SERVICIO'];
    if (urlOrden && urlOrden.toString().length > 10) {
        linkOrdenBlock = `<p style="margin-top: 15px;"><strong>Documento Adjunto:</strong> <a href="${urlOrden}" style="color: #2b6cb0; text-decoration: none; font-weight: bold;">Ver Orden de Servicio 📎</a></p>`;
    }

    // Unimos las partes del cuerpo
    let cuerpoFinal = `${intro}<br><br>${detalle}${previredBlock}<br>${cierre}`;

    // Reemplazos de variables del Sheet (Fechas, IDs, etc)
    for (const k in registro) {
        let val = registro[k];
        if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
        // Reemplazo global
        cuerpoFinal = cuerpoFinal.split(`{{${k}}}`).join(val || '');
    }

    // --- HTML ESTRUCTURA FINAL ---
    const htmlTemplate = `
      <div style="font-family: 'Segoe UI', Helvetica, Arial, sans-serif; color: #374151; max-width: 680px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
        
        <!-- Header Azul -->
        <div style="background: linear-gradient(90deg, #1e3a8a 0%, #2563eb 100%); padding: 30px 40px;">
           <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 700;">Bienvenido/a a la I. Municipalidad de Coquimbo</h1>
           <p style="color: #bfdbfe; margin: 5px 0 0 0; font-size: 14px;">Departamento De Personal</p>
        </div>

        <!-- Body -->
        <div style="padding: 40px;">
           <p style="font-size: 18px; color: #1e3a8a; font-weight: 600; margin-top: 0;">Estimado/a ${nombreSaludo},</p>
           
           <div style="font-size: 15px; line-height: 1.6; color: #4b5563;">
              ${cuerpoFinal}
           </div>

           <!-- Tarjeta Resumen -->
           <div style="margin-top: 30px; background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px;">
             <h3 style="margin: 0 0 10px 0; font-size: 14px; color: #1e40af; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px;">Detalles del Ingreso</h3>
             <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
               <tr>
                  <td style="padding: 5px 0; width: 120px; font-weight: bold; color: #64748b;">Calidad:</td>
                  <td style="padding: 5px 0; color: #334155;">${toTitleCase(registro['CALIDAD CONTRACTUAL'] || '')}</td>
               </tr>
               <tr>
                  <td style="padding: 5px 0; font-weight: bold; color: #64748b;">Fecha:</td>
                  <td style="padding: 5px 0; color: #334155;">${registro['FECHA INGRESO'] instanceof Date ? Utilities.formatDate(registro['FECHA INGRESO'], Session.getScriptTimeZone(), "dd/MM/yyyy") : registro['FECHA INGRESO']}</td>
               </tr>
               <tr>
                  <td style="padding: 5px 0; font-weight: bold; color: #64748b; vertical-align: top;">Función:</td>
                  <td style="padding: 5px 0; color: #334155;">${toTitleCase(funcionParaMostrar)}</td>
               </tr>
             </table>
             ${linkOrdenBlock}
           </div>

           <!-- Firma -->
           <div style="margin-top: 40px; padding-top: 25px; border-top: 1px solid #f3f4f6;">
              <table style="width: 100%;">
                 <tr>
                    <td style="vertical-align: top;">
                        <p style="margin: 0; font-weight: bold; color: #111827; font-size: 15px;">Dirección de Recursos Humanos</p>
                        <p style="margin: 2px 0 0 0; color: #4b5563; font-size: 14px;">I. Municipalidad de Coquimbo</p>
                        <p style="margin: 2px 0 0 0; color: #6b7280; font-size: 13px;">Edificio Consistorial (Piso 7) - Av. Varela 1112</p>
                    </td>
                 </tr>
                 <tr>
                    <td style="padding-top: 20px;">
                       <img src="https://ci3.googleusercontent.com/mail-sig/AIorK4wEkBCmQ9hUXuoKZntbxbO_8LN-c7WwNUwhVYGQWs6grlb9lLqv62x51b5qeawh8mhPiLUMw4Z7N7yh" alt="Firma Institucional" style="max-width: 100%; height: auto; border-radius: 6px; border: 1px solid #e5e7eb;">
                    </td>
                 </tr>
              </table>
           </div>
        </div>
        
        <!-- Footer -->
        <div style="background-color: #f9fafb; padding: 15px; text-align: center; border-top: 1px solid #e5e7eb;">
           <p style="margin: 0; font-size: 11px; color: #9ca3af;">Este es un mensaje automático. Por favor no responder a esta dirección.</p>
        </div>
      </div>
    `;

    return {
        asunto: `ORDEN DE SERVICIO ${registro['NUMERO ORDEN DE SERVICIO']}/${nombreC} - INGRESO NUEVO`,
        cuerpoHtml: htmlTemplate,
        emailDestino: registro['CORREO ELECTRONICO']
    };
}

function previsualizarCorreo(id, tipo) {
    try {
        return { success: true, data: generarContenidoCorreo(id, tipo) };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * Función Principal de Envío (Personalizada)
 * @param {Array} listaEnvios - Array de objetos: [{ id: "123", extraCc: "jefe@municoquimbo.cl" }, ...]
 * @param {Array} globalCcs - Array de strings con los CC fijos
 * @param {String} tipo - 'A' o 'B'
 */
function enviarCorreosDeBienvenida(listaEnvios, globalCcs, tipo) {
    try {
        const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTRO_MAESTRO');
        const data = hoja.getDataRange().getValues();
        const headers = data.shift();
        const idIdx = headers.indexOf('ID');
        const sentIdx = headers.indexOf('CORREO BIENVENIDA ENVIADO');
        
        let count = 0;
        let errores = [];

        const ccsBase = (globalCcs || []).filter(email => email && email.includes('@'));

        listaEnvios.forEach(item => {
            try {
                const idEmpleado = item.id;
                const ccEspecifico = item.extraCc;

                const contenido = generarContenidoCorreo(idEmpleado, tipo);
                
                if (contenido && contenido.emailDestino) {
                    
                    let listaCCFinal = [...ccsBase];
                    
                    if (ccEspecifico && ccEspecifico.includes('@')) {
                        const extras = ccEspecifico.split(',').map(e => e.trim());
                        listaCCFinal = listaCCFinal.concat(extras);
                    }

                    const ccFinalString = [...new Set(listaCCFinal)].join(',');

                    MailApp.sendEmail({
                        to: contenido.emailDestino,
                        cc: ccFinalString,
                        subject: contenido.asunto,
                        htmlBody: contenido.cuerpoHtml
                    });

                    const row = data.findIndex(r => String(r[idIdx]) === String(idEmpleado));
                    if (row !== -1) {
                        hoja.getRange(row + 2, sentIdx + 1).setValue(new Date());
                    }
                    count++;
                }
            } catch (errLocal) {
                console.error(`Error enviando a ID ${item.id}: ${errLocal.message}`);
                errores.push(`ID ${item.id}: ${errLocal.message}`);
            }
        });

        if (errores.length > 0) {
            return { success: true, message: `${count} enviados. Hubo errores en: ${errores.join(', ')}` };
        }

        return { success: true, message: `Se enviaron ${count} correos exitosamente.` };

    } catch (e) {
        return { success: false, message: e.message };
    }
}