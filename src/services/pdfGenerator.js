// ============================================================================
// GENERADOR DE PDF — S&S IA
// ============================================================================
// Convierte un elemento ya renderizado en pantalla (por ejemplo, la
// plantilla de cotización) en un archivo PDF real de tamaño carta — para
// poder subirlo a OneDrive como archivo, no solo enviar filas a una tabla
// de Excel. Antes de esto, la app solo podía "imprimir" con el diálogo del
// navegador, que no genera ningún archivo que la app pueda manipular.
//
// IMPORTANTE — margen y llenado de página: antes esto encogía la imagen
// para que cupiera completa en una sola página (ancho Y alto a la vez), así
// que una cotización corta (pocos ítems) quedaba como una tarjeta chiquita
// pegada arriba, con medio pliego en blanco debajo — se veía "desconfigurada".
// Ahora se ajusta SIEMPRE al ancho útil de la página (con margen real en los
// 4 lados, como cualquier documento impreso), y si el contenido no cabe de
// alto en una sola página (cotizaciones con muchos ítems), se reparte en
// varias páginas automáticamente — igual que lo haría Word o Excel al
// imprimir un documento largo.
// ============================================================================

import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";

const MARGEN_IN = 0.4; // margen real en los 4 lados, en pulgadas

/**
 * @param {HTMLElement} elemento - el nodo del DOM a convertir (debe estar
 *   visible/renderizado en el momento de llamar esta función).
 * @returns {Promise<Blob>} el PDF ya armado, tamaño carta, listo para
 *   descargar o subir a OneDrive.
 */
export async function generarPdfDesdeElemento(elemento) {
  // ANCHO_VENTANA_VIRTUAL — en celular, el navegador angosta el contenedor
  // de la plantilla (max-w-3xl) al ancho real de la pantalla del teléfono,
  // así que la captura salía chiquita, apretada y con columnas mal armadas
  // — se veía "desconfigurada" solo ahí, nunca en escritorio (donde la
  // pantalla ya es más ancha que la plantilla, así que nunca la achicaba).
  // windowWidth le dice a html2canvas que calcule el layout como si el
  // navegador tuviera este ancho, sin importar el ancho real del
  // dispositivo — el mismo resultado en celular que en escritorio, y no
  // cambia nada en escritorio porque ahí ya se renderizaba a este ancho o
  // más angosto de forma natural.
  const ANCHO_VENTANA_VIRTUAL = 1000;
  const canvas = await html2canvas(elemento, {
    scale: 2, // más nitidez que una captura a resolución normal
    useCORS: true,
    backgroundColor: "#ffffff",
    windowWidth: ANCHO_VENTANA_VIRTUAL,
  });

  const pdf = new jsPDF({ unit: "in", format: "letter" });
  const anchoPagina = pdf.internal.pageSize.getWidth();
  const altoPagina = pdf.internal.pageSize.getHeight();
  const anchoUtil = anchoPagina - MARGEN_IN * 2;
  const altoUtil = altoPagina - MARGEN_IN * 2;

  // Se ajusta SIEMPRE por ancho (llena el ancho útil de la página) — el
  // alto sale proporcional al contenido real, sin estirar ni encoger la
  // imagen para "rellenar" espacio que no existe.
  const escala = anchoUtil / canvas.width;
  const altoTotalEscalado = canvas.height * escala;

  if (altoTotalEscalado <= altoUtil) {
    // Cabe completa en una sola página.
    const imagenDatos = canvas.toDataURL("image/png");
    pdf.addImage(imagenDatos, "PNG", MARGEN_IN, MARGEN_IN, anchoUtil, altoTotalEscalado);
  } else {
    // No cabe de alto — se reparte en varias páginas, cortando el canvas
    // original (a resolución real, no la imagen ya escalada) en franjas de
    // la altura que sí cabe por página.
    const altoUtilEnPx = Math.floor(altoUtil / escala); // alto de cada franja, en px del canvas original
    let yaCopiadoPx = 0;
    let primeraPagina = true;

    while (yaCopiadoPx < canvas.height) {
      const altoFranjaPx = Math.min(altoUtilEnPx, canvas.height - yaCopiadoPx);

      const franja = document.createElement("canvas");
      franja.width = canvas.width;
      franja.height = altoFranjaPx;
      franja.getContext("2d").drawImage(
        canvas,
        0, yaCopiadoPx, canvas.width, altoFranjaPx, // origen: recorte del canvas completo
        0, 0, canvas.width, altoFranjaPx // destino: canvas de la franja
      );

      if (!primeraPagina) pdf.addPage();
      pdf.addImage(franja.toDataURL("image/png"), "PNG", MARGEN_IN, MARGEN_IN, anchoUtil, altoFranjaPx * escala);

      yaCopiadoPx += altoFranjaPx;
      primeraPagina = false;
    }
  }

  return pdf.output("blob");
}
