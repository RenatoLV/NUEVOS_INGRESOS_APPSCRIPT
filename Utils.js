// =================================================================================
// --- FUNCIONES DE UTILIDAD Y UI DEL SERVIDOR ---
// =================================================================================

/**
 * Permite incluir archivos HTML dentro de otros archivos HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Muestra el diálogo modal del formulario principal.
 */
function mostrarFormulario() {
  const html = HtmlService.createTemplateFromFile('Formulario')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Formulario de Ingreso y Edición de Personal');
}

/**
 * Muestra un pop-up personalizado con un enlace clickeable.
 */
function mostrarPopupConEnlace(url, nombre) {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 10px;">
      <p>✅ <b>¡Éxito!</b></p>
      <p>Los documentos para <i>${nombre}</i> se han generado correctamente.</p>
      <br>
      <a href="${url}" target="_blank" onclick="google.script.host.close()" style="background-color: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
        Abrir Carpeta de Documentos
      </a>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(400)
      .setHeight(180);
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generación Completa');
}

/**
 * Convierte un número a su representación en palabras (pesos chilenos).
 */
function numeroAJson(numero) {
  const unidades = ['', 'Un', 'Dos', 'Tres', 'Cuatro', 'Cinco', 'Seis', 'Siete', 'Ocho', 'Nueve'];
  const decenas = ['', 'Diez', 'Veinte', 'Treinta', 'Cuarenta', 'Cincuenta', 'Sesenta', 'Setenta', 'Ochenta', 'Noventa'];
  const especiales = ['Diez', 'Once', 'Doce', 'Trece', 'Catorce', 'Quince', 'Dieciséis', 'Diecisiete', 'Dieciocho', 'Diecinueve'];
  const centenas = ['', 'Ciento', 'Doscientos', 'Trescientos', 'Cuatrocientos', 'Quinientos', 'Seiscientos', 'Setecientos', 'Ochocientos', 'Novecientos'];

  function convertirGrupo(n) {
    if (n > 999) return 'Máximo 999';
    if (n === 100) return 'Cien';
    let output = '';
    const c = Math.floor(n / 100);
    const d = Math.floor((n % 100) / 10);
    const u = n % 10;
    output += centenas[c];
    const du = n % 100;
    if (du > 0) {
      if (c > 0) output += ' ';
      if (du < 20 && du > 9) {
        output += especiales[du - 10];
      } else {
        output += decenas[d];
        if (u > 0) {
          if (d > 2) output += ' y ';
          output += unidades[u];
        }
      }
    }
    return output;
  }

  if (numero === null || isNaN(numero)) return '';
  if (numero === 0) return 'Cero Pesos';

  const entero = Math.floor(numero);
  let texto = '';
  const millones = Math.floor(entero / 1000000);
  let resto = entero % 1000000;
  if (millones > 0) {
    texto += (millones === 1 ? 'Un Millón' : convertirGrupo(millones) + ' Millones');
    if (resto > 0) texto += ' ';
  }
  const miles = Math.floor(resto / 1000);
  resto %= 1000;
  if (miles > 0) {
    if (miles === 1) texto += 'Mil';
    else texto += convertirGrupo(miles) + ' Mil';
    if (resto > 0) texto += ' ';
  }
  if (resto > 0) {
    texto += convertirGrupo(resto);
  }
  return (texto + ' Pesos').trim();
}

/**
 * Función de depuración para diagnosticar problemas con nombres de pestañas.
 */
function diagnosticarNombresDePestaña() {
  const idArchivoNominas = ID_HOJA_NOMINAS; 
  try {
    const ssDestino = SpreadsheetApp.openById(idArchivoNominas);
    const todasLasPestañas = ssDestino.getSheets();
    console.log('--- INICIO DEL DIAGNÓSTICO ---');
    const nombreProblematico = 'PRESTADORES DE SERVICIO';
    console.log(`Buscando la pestaña: "${nombreProblematico}" (Longitud: ${nombreProblematico.length})`);
    console.log('\nPestañas encontradas en el archivo de destino:');
    let encontrada = false;
    todasLasPestañas.forEach(pestaña => {
      const nombrePestañaActual = pestaña.getName();
      console.log(`- Pestaña: "${nombrePestañaActual}" (Longitud: ${nombrePestañaActual.length})`);
      if (nombrePestañaActual === nombreProblematico) {
        console.log('  ✅ ¡COINCIDENCIA EXACTA ENCONTRADA!');
        encontrada = true;
      }
    });
    if (!encontrada) {
        console.log('\n❌ NO SE ENCONTRÓ COINCIDENCIA EXACTA. Esto confirma el problema.');
    }
    console.log('--- FIN DEL DIAGNÓSTICO ---');
  } catch (e) {
    console.error('Error durante el diagnóstico: ' + e.message);
  }
}