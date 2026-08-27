// ============================================================================
// SERVICIO DE EXCEL EN ONEDRIVE VÍA MICROSOFT GRAPH — S&S IA
// ============================================================================
// Lee y escribe filas en libros de Excel alojados en OneDrive, usando la
// API de Excel de Microsoft Graph (workbook / tables). Requiere que exista
// una sesión de Microsoft activa (ver authConfig.js y App.jsx) — cada
// función recibe la instancia de MSAL y la cuenta activa para poder pedir
// el token de acceso.
//
// REQUISITO IMPORTANTE EN EL LADO DE EXCEL: para usar las funciones de
// filas (leerFilasTabla / agregarFilaTabla / actualizarFilaTabla) el rango
// de datos en el libro debe estar convertido en una "Tabla" real de Excel
// (seleccionar el rango → Insertar → Tabla, o Ctrl+T) y esa tabla debe
// tener un nombre (Diseño de tabla → Nombre de la tabla). Si solo tienes
// datos en celdas sueltas sin tabla, usa leerRango / escribirRango en su
// lugar, indicando el rango tipo "A1:F50".
// ============================================================================

import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { graphScopes, graphBaseUrl } from "../authConfig";

// -------- Rutas y nombres de tabla — AJUSTA ESTO a tus archivos reales --------
// Las rutas son relativas a la raíz de OneDrive de la cuenta conectada
// (subgerencia@suaservice.com). Cambia estos valores para que apunten a
// tus libros reales antes de usar las funciones de conveniencia de más
// abajo (sincronizarCotizacion, sincronizarRegistro, sincronizarNomina).
export const ARCHIVOS_EXCEL = {
  cotizaciones: { ruta: "/S&SIA/Cotizaciones.xlsx", tabla: "TablaCotizaciones" },
  registros: { ruta: "/S&SIA/Registro_Operaciones.xlsx", tabla: "TablaRegistros" },
  nomina: { ruta: "/S&SIA/Nomina.xlsx", tabla: "TablaNomina" },
};

/**
 * Consigue un token de acceso para Graph — intenta en silencio primero
 * (sin interrumpir al usuario); si la sesión de Microsoft expiró o hace
 * falta volver a autorizar, abre el popup de login solo en ese caso.
 * @param {import("@azure/msal-browser").PublicClientApplication} instance
 * @param {import("@azure/msal-browser").AccountInfo} account
 */
export async function obtenerTokenGraph(instance, account) {
  if (!account) throw new Error("No hay una cuenta de Microsoft conectada. Conecta Microsoft 365 primero.");
  const solicitud = { scopes: graphScopes, account };
  try {
    const resultado = await instance.acquireTokenSilent(solicitud);
    return resultado.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const resultado = await instance.acquireTokenPopup(solicitud);
      return resultado.accessToken;
    }
    throw error;
  }
}

/** Llamada genérica a Graph, con el token ya incluido en el encabezado. */
async function llamarGraph(token, metodo, rutaRelativa, cuerpo) {
  const respuesta = await fetch(`${graphBaseUrl}${rutaRelativa}`, {
    method: metodo,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: cuerpo ? JSON.stringify(cuerpo) : undefined,
  });
  if (!respuesta.ok) {
    const detalle = await respuesta.text().catch(() => "");
    throw new Error(`Graph API respondió ${respuesta.status}: ${detalle || respuesta.statusText}`);
  }
  if (respuesta.status === 204) return null; // sin contenido (ej. tras un DELETE)
  return respuesta.json();
}

// ============================================================================
// LECTURA / ESCRITURA POR TABLA DE EXCEL (recomendado)
// ============================================================================

/**
 * Lee todas las filas de una tabla de Excel.
 * @returns {Promise<Array<{index: number, values: any[]}>>} cada fila con su
 *   índice dentro de la tabla y sus valores en el mismo orden de columnas.
 */
export async function leerFilasTabla(token, rutaArchivo, nombreTabla) {
  const datos = await llamarGraph(
    token,
    "GET",
    `/me/drive/root:${rutaArchivo}:/workbook/tables/${nombreTabla}/rows`
  );
  return (datos?.value || []).map((fila, i) => ({ index: fila.index ?? i, values: fila.values?.[0] || [] }));
}

/**
 * Agrega una fila nueva al final de la tabla.
 * @param {any[]} valoresFila - un valor por columna, en el mismo orden que
 *   las columnas de la tabla en Excel.
 */
export async function agregarFilaTabla(token, rutaArchivo, nombreTabla, valoresFila) {
  return llamarGraph(
    token,
    "POST",
    `/me/drive/root:${rutaArchivo}:/workbook/tables/${nombreTabla}/rows/add`,
    { values: [valoresFila] }
  );
}

/**
 * Actualiza una fila existente de la tabla, identificada por su índice
 * (el mismo "index" que devuelve leerFilasTabla).
 */
export async function actualizarFilaTabla(token, rutaArchivo, nombreTabla, indiceFila, valoresFila) {
  return llamarGraph(
    token,
    "PATCH",
    `/me/drive/root:${rutaArchivo}:/workbook/tables/${nombreTabla}/rows/itemAt(index=${indiceFila})`,
    { values: [valoresFila] }
  );
}

/** Elimina una fila de la tabla por su índice. */
export async function eliminarFilaTabla(token, rutaArchivo, nombreTabla, indiceFila) {
  return llamarGraph(
    token,
    "POST",
    `/me/drive/root:${rutaArchivo}:/workbook/tables/${nombreTabla}/rows/itemAt(index=${indiceFila})/delete`
  );
}

// ============================================================================
// LECTURA / ESCRITURA POR RANGO (para hojas sin tabla con nombre)
// ============================================================================

/** Lee un rango de celdas, ej. leerRango(token, ruta, "Hoja1", "A1:F20"). */
export async function leerRango(token, rutaArchivo, nombreHoja, rango) {
  const datos = await llamarGraph(
    token,
    "GET",
    `/me/drive/root:${rutaArchivo}:/workbook/worksheets/${nombreHoja}/range(address='${rango}')`
  );
  return datos?.values || [];
}

/** Escribe valores en un rango de celdas (sobreescribe lo que haya ahí). */
export async function escribirRango(token, rutaArchivo, nombreHoja, rango, valores) {
  return llamarGraph(
    token,
    "PATCH",
    `/me/drive/root:${rutaArchivo}:/workbook/worksheets/${nombreHoja}/range(address='${rango}')`,
    { values: valores }
  );
}

// ============================================================================
// FUNCIONES DE CONVENIENCIA — cotizaciones, registros y nómina
// ============================================================================
// Ajusta el orden de columnas dentro de cada "valoresFila" para que
// coincida exactamente con el orden real de columnas de tu tabla en Excel.

/** Envía una cotización (ya calculada) como fila nueva a la tabla de Excel. */
export async function sincronizarCotizacion(instance, account, cotizacion) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, tabla } = ARCHIVOS_EXCEL.cotizaciones;
  const fila = [
    `${cotizacion.numero}-${cotizacion.revision}`,
    cotizacion.fecha,
    cotizacion.clienteNombre || "",
    cotizacion.referencia,
    cotizacion.estado,
    cotizacion.subtotalBruto ?? cotizacion.subtotal,
    cotizacion.iva,
    cotizacion.total,
  ];
  return agregarFilaTabla(token, ruta, tabla, fila);
}

/** Envía un registro de operación (venta/compra/gasto/cobro/pago) como fila nueva. */
export async function sincronizarRegistro(instance, account, operacion) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, tabla } = ARCHIVOS_EXCEL.registros;
  const fila = [
    operacion.id,
    operacion.tipo,
    operacion.fecha,
    operacion.concepto,
    operacion.valor,
    operacion.iva || 0,
  ];
  return agregarFilaTabla(token, ruta, tabla, fila);
}

/** Envía un pago de nómina como fila nueva. */
export async function sincronizarNomina(instance, account, nomina) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, tabla } = ARCHIVOS_EXCEL.nomina;
  const fila = [
    nomina.colaboradorNombre || nomina.colaboradorId,
    nomina.periodo || nomina.fecha,
    nomina.valor,
  ];
  return agregarFilaTabla(token, ruta, tabla, fila);
}
