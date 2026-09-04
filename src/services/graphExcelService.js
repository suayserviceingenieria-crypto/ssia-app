// ============================================================================
// SERVICIO DE EXCEL EN ONEDRIVE VÍA MICROSOFT GRAPH — S&S IA
// ============================================================================
// Lee el libro completo de OneDrive, lo edita en memoria con SheetJS (xlsx)
// y lo vuelve a subir — en vez de usar la API de "Tablas" de Excel de
// Graph. Esto evita que la sincronización dependa de que cada rango esté
// convertido en una Tabla con nombre exacto en Excel: basta con que la
// hoja exista y tenga el nombre correcto (ver ARCHIVOS_EXCEL más abajo).
// Requiere que exista una sesión de Microsoft activa (ver authConfig.js y
// App.jsx) — cada función recibe la instancia de MSAL y la cuenta activa
// para poder pedir el token de acceso.
// ============================================================================

import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { graphScopes, graphBaseUrl } from "../authConfig";
import * as XLSX from "xlsx";

// -------- DRIVE FIJO — todo se escribe siempre aquí, sin importar quién esté conectado --------
// En vez de escribir en "el OneDrive de quien esté conectado" (/me/drive),
// la app escribe siempre en ESTE Drive específico (el de
// subgerencia@suaservice.com) — así comercial@ y gerencia@ pueden
// conectarse cada uno con SU PROPIO correo (sin conocer la clave de
// subgerencia) y aun así todos los archivos caen en el mismo lugar,
// siempre que esa carpeta se haya compartido con ellos con permiso de
// edición (ver instrucciones más abajo).
//
// CÓMO OBTENER ESTE ID (una sola vez, como subgerencia):
//   1. Entra a https://developer.microsoft.com/graph/graph-explorer
//   2. Inicia sesión ahí con subgerencia@suaservice.com
//   3. Pega esta consulta y dale "Run query": GET https://graph.microsoft.com/v1.0/me/drive
//   4. En la respuesta, copia el valor del campo "id" (una cadena larga
//      con letras y números) y pégalo aquí abajo, reemplazando el texto
//      de ejemplo. SOLO el valor de "id" entre comillas — nunca pegues la
//      respuesta completa de Graph Explorer aquí.
//
// CÓMO COMPARTIR LA CARPETA (una sola vez, como subgerencia):
//   1. Entra a onedrive.com con subgerencia@suaservice.com
//   2. Busca la carpeta "SSIA" → clic derecho → "Compartir"
//   3. Agrega a comercial@suaservice.com y gerencia@suaservice.com
//   4. Asegúrate de que el permiso diga "Puede editar" (no "Puede ver")
const DRIVE_ID = "b!SgDOPot9M0i8PRasYq_rXvTyfY4Iy6NDqfFH_oLITU2T4rkW2B2JQpoR-u-pYFfA";

// Correo cuyo calendario usa el módulo de tareas/agenda (crear/actualizar/
// eliminar eventos) — igual que con el Drive fijo de arriba, todos los
// eventos caen siempre en ESTE calendario sin importar quién esté conectado.
const CORREO_CALENDARIO = "comercial@suaservice.com";

/** Falla temprano y con un mensaje claro si alguien dejó DRIVE_ID mal pegado. */
function validarDriveId() {
  if (DRIVE_ID.startsWith("PEGA_AQUI")) {
    throw new Error(
      'DRIVE_ID no está configurado correctamente en graphExcelService.js — debe ser un texto simple, el valor del campo "id" de GET /me/drive (no la respuesta completa, y no el id de GET /me).'
    );
  }
}

// Prefijo de ruta que reemplaza a "/me/drive" en todas las llamadas —
// apunta siempre al Drive fijo de arriba, no al de quien esté conectado.
const rutaDrive = `/drives/${DRIVE_ID}`;

// -------- Rutas, tabla y hoja de cada libro — AJUSTA ESTO a tus archivos reales --------
// "hoja" es el nombre real de la pestaña dentro del libro — es lo que se
// usa para leer/escribir (ya no se depende de que el rango esté convertido
// en una Tabla con nombre). "tabla" se deja documentado por si algún día
// se vuelve a necesitar, pero las funciones de este archivo ya no lo usan.
export const ARCHIVOS_EXCEL = {
  cotizaciones: { ruta: "/SSIA/Cotizaciones.xlsx", tabla: "TablaCotizaciones", hoja: "Datos" },
  cotizacionItems: { ruta: "/SSIA/Cotizaciones.xlsx", tabla: "TablaCotizacionItems", hoja: "Items" },
  catalogoItems: { ruta: "/SSIA/Cotizaciones.xlsx", tabla: "TablaCatalogoItems", hoja: "CatalogoItems" },
  registros: { ruta: "/SSIA/Registro_Operaciones.xlsx", tabla: "TablaRegistros", hoja: "Datos" },
  nomina: { ruta: "/SSIA/Nomina.xlsx", tabla: "TablaNomina", hoja: "Datos" },
  nominaMensual: { ruta: "/SSIA/Nomina.xlsx", tabla: "TablaNominaMensual", hoja: "NominaMensual" },
  cesantias: { ruta: "/SSIA/Nomina.xlsx", tabla: "TablaCesantias", hoja: "Cesantias" },
  prima: { ruta: "/SSIA/Nomina.xlsx", tabla: "TablaPrima", hoja: "Prima" },
  vacaciones: { ruta: "/SSIA/Nomina.xlsx", tabla: "TablaVacaciones", hoja: "Vacaciones" },
  liquidacionContrato: { ruta: "/SSIA/Nomina.xlsx", tabla: "TablaLiquidacionContrato", hoja: "LiquidacionContrato" },
  terceroClientes: { ruta: "/SSIA/Terceros.xlsx", tabla: "TablaClientes", hoja: "Clientes" },
  terceroProveedores: { ruta: "/SSIA/Terceros.xlsx", tabla: "TablaProveedores", hoja: "Proveedores" },
  proyectos: { ruta: "/SSIA/Proyectos.xlsx", tabla: "TablaProyectos", hoja: "Proyectos" },
  cartera: { ruta: "/SSIA/Cartera_CxC.xlsx", tabla: "TablaCartera", hoja: "Cartera" },
  cxpComprasGastos: { ruta: "/SSIA/CuentasPorPagar.xlsx", tabla: "TablaComprasGastos", hoja: "ComprasGastos" },
  cxpPrestaciones: { ruta: "/SSIA/CuentasPorPagar.xlsx", tabla: "TablaPrestacionesPendientes", hoja: "PrestacionesPendientes" },
  cxpProvision: { ruta: "/SSIA/CuentasPorPagar.xlsx", tabla: "TablaProvision", hoja: "Provision" },
};

// Carpetas donde se guardan los PDF (y el Excel mensual de nómina)
// generados — dentro de la misma carpeta SSIA de OneDrive, en subcarpetas
// aparte para no mezclarlos con los libros de Excel.
export const CARPETA_PDFS = "/SSIA/PDFs";
export const CARPETA_PDFS_NOMINA = "/SSIA/PDFs/Nomina";
export const CARPETA_PDFS_PRESTACIONES = "/SSIA/PDFs/Prestaciones";
export const CARPETA_PDFS_REGISTRO = "/SSIA/PDFs/Registro";
export const CARPETA_PDFS_INFORMES = "/SSIA/PDFs/Informes";

export async function subirPdfCotizacion(instance, account, cotizacion, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const numeroCompleto = `${cotizacion.numero}-${cotizacion.revision}`;
  const nombreArchivo = `Cotizacion_${numeroCompleto}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS}/${nombreArchivo}`, blob, "application/pdf");
}

/** Comprobante de VENTA/COMPRA/GASTO — módulo "Registrar hecho económico". */
export async function subirPdfComprobante(instance, account, operacion, nombreTercero, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const tipoLabel = { VENTA: "Venta", COMPRA: "Compra", GASTO: "Gasto" }[operacion.tipo] || operacion.tipo;
  const nombreArchivo = `${tipoLabel}_${operacion.id}_${nombreTercero || "sin_tercero"}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_REGISTRO}/${nombreArchivo}`, blob, "application/pdf");
}

/** Comprobante de COBRO/PAGO — módulo Cuentas por cobrar / por pagar. */
export async function subirPdfComprobanteMovimiento(instance, account, operacion, nombreTercero, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const tipoLabel = operacion.tipo === "COBRO" ? "Cobro" : "Pago";
  const nombreArchivo = `${tipoLabel}_${operacion.id}_${nombreTercero || "sin_tercero"}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_REGISTRO}/${nombreArchivo}`, blob, "application/pdf");
}

/** Comprobante individual de un pago de nómina (una quincena, un colaborador). */
export async function subirPdfNomina(instance, account, nomina, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const nombreArchivo = `Nomina_${nomina.colaboradorNombre || nomina.colaboradorId || "colaborador"}_${nomina.periodo || nomina.fecha || ""}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_NOMINA}/${nombreArchivo}`, blob, "application/pdf");
}

/**
 * Sube el Excel de nómina mensual compilado (con provisiones) — un archivo
 * APARTE por mes (NominaMensual_2026-09.xlsx), distinto de Nomina.xlsx (que
 * es el historial fila por fila de cada pago individual). Cada mes que se
 * vuelve a generar el compilado, este archivo se reemplaza — es una "foto"
 * del mes, no un historial que crece.
 */
export async function subirExcelNominaMensual(instance, account, mesNomina, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const nombreArchivo = `NominaMensual_${mesNomina}.xlsx`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_NOMINA}/${nombreArchivo}`, blob, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

/** Sube el PDF consolidado de nómina mensual (lista de pagos del mes) — igual que el Excel de arriba, un archivo aparte por mes. */
export async function subirPdfNominaMensual(instance, account, mesNomina, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const nombreArchivo = `NominaMensual_${mesNomina}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_NOMINA}/${nombreArchivo}`, blob, "application/pdf");
}

/** Comprobante de una liquidación de cesantías, prima, vacaciones o de contrato. */
export async function subirPdfPrestacion(instance, account, tipoPrestacion, liquidacion, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const tipoLabel = { CESANTIAS: "Cesantias", PRIMA: "Prima", VACACIONES: "Vacaciones", LIQUIDACION_FINAL: "LiquidacionFinal" }[tipoPrestacion] || tipoPrestacion;
  const periodoLabel = liquidacion.anioLiquidado || liquidacion.fechaTerminacion || liquidacion.fechaCorte || (liquidacion.semestrePrima ? `${liquidacion.semestrePrima}S-${liquidacion.anioPrima}` : "");
  const nombreArchivo = `${tipoLabel}_${liquidacion.colaboradorNombre || "colaborador"}_${periodoLabel}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_PRESTACIONES}/${nombreArchivo}`, blob, "application/pdf");
}

/** Informe genérico (proyectos, presupuesto, general, indicadores) — un PDF con la fecha del día. */
export async function subirPdfInforme(instance, account, nombre, blob) {
  const token = await obtenerTokenGraph(instance, account);
  const fechaHoy = new Date().toISOString().slice(0, 10);
  const nombreArchivo = `${nombre}_${fechaHoy}.pdf`.replace(/[\\/:*?"<>|]/g, "-");
  return subirArchivoOneDrive(token, `${CARPETA_PDFS_INFORMES}/${nombreArchivo}`, blob, "application/pdf");
}

/**
 * Consigue un token de acceso para Graph — intenta en silencio primero
 * (sin interrumpir al usuario); si la sesión de Microsoft expiró o hace
 * falta volver a autorizar, abre el popup de login solo en ese caso.
 * @param {import("@azure/msal-browser").PublicClientApplication} instance
 * @param {import("@azure/msal-browser").AccountInfo} account
 */
export async function obtenerTokenGraph(instance, account) {
  validarDriveId();
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

/**
 * Sube un archivo completo (no filas de una hoja) a OneDrive — para PDFs
 * de cotizaciones, comprobantes, el Excel mensual de nómina, etc. Si ya
 * existe un archivo con ese nombre exacto en esa ruta, lo reemplaza
 * (comportamiento normal de OneDrive al subir con el mismo nombre).
 * @param {Blob} blob - el archivo, ya generado (ej. un PDF).
 * @param {string} tipoContenido - ej. "application/pdf".
 */
export async function subirArchivoOneDrive(token, rutaArchivo, blob, tipoContenido) {
  const respuesta = await fetch(`${graphBaseUrl}${rutaDrive}/root:${rutaArchivo}:/content`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": tipoContenido,
    },
    body: blob,
  });
  if (!respuesta.ok) {
    const detalle = await respuesta.text().catch(() => "");
    throw new Error(`Graph API respondió ${respuesta.status} al subir archivo: ${detalle || respuesta.statusText}`);
  }
  return respuesta.json();
}

/** Descarga el contenido crudo (bytes) de un archivo de OneDrive. */
async function descargarArchivoOneDrive(token, rutaArchivo) {
  const respuesta = await fetch(`${graphBaseUrl}${rutaDrive}/root:${rutaArchivo}:/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!respuesta.ok) {
    const detalle = await respuesta.text().catch(() => "");
    throw new Error(`Graph API respondió ${respuesta.status} al descargar archivo: ${detalle || respuesta.statusText}`);
  }
  return respuesta.arrayBuffer();
}

function esperar(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Cola de promesas por archivo — serializa todas las operaciones de
// descargar→editar→resubir sobre un mismo libro, para que dos escrituras
// casi simultáneas (dos pestañas, dos usuarios) no se pisen entre sí
// dentro de esta misma sesión de la app.
const colaPorArchivo = new Map();
function encolarPorArchivo(rutaArchivo, tarea) {
  const siguiente = (colaPorArchivo.get(rutaArchivo) || Promise.resolve()).catch(() => {}).then(tarea);
  colaPorArchivo.set(rutaArchivo, siguiente);
  return siguiente;
}

/**
 * Agrega filas nuevas al final de una hoja (debajo de lo que ya haya),
 * descargando el libro, editándolo con SheetJS y resubiéndolo. Reintenta
 * hasta 4 veces si el archivo está abierto por otra persona (error 423 de
 * OneDrive/Excel), esperando 4s entre intento e intento.
 */
export async function agregarFilasSinSesionInterno(token, rutaArchivo, nombreHoja, filasNuevas) {
  return encolarPorArchivo(rutaArchivo, () => intentarAgregarFilas(token, rutaArchivo, nombreHoja, filasNuevas));
}
async function intentarAgregarFilas(token, rutaArchivo, nombreHoja, filasNuevas) {
  for (let intento = 1; intento <= 4; intento++) {
    try {
      const buffer = await descargarArchivoOneDrive(token, rutaArchivo);
      const libro = XLSX.read(buffer, { type: "array" });
      const hoja = libro.Sheets[nombreHoja];
      if (!hoja) throw new Error(`La hoja "${nombreHoja}" no existe en ${rutaArchivo} — revisa el nombre exacto de la pestaña.`);
      const filasActuales = [...XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" }), ...filasNuevas];
      libro.Sheets[nombreHoja] = XLSX.utils.aoa_to_sheet(filasActuales);
      const bufferNuevo = XLSX.write(libro, { type: "array", bookType: "xlsx" });
      const blob = new Blob([bufferNuevo], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      return await subirArchivoOneDrive(token, rutaArchivo, blob, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } catch (error) {
      if (/423|locked/i.test(String(error?.message || error)) && intento < 4) {
        console.warn(`[OneDrive] Archivo bloqueado, reintentando (${intento}/3)…`);
        await esperar(4000);
        continue;
      }
      throw error;
    }
  }
}

/**
 * Recorre filas existentes de una o varias hojas y las transforma con un
 * callback — usado para actualizar (o borrar, si el callback devuelve
 * null) filas puntuales sin tocar el resto del libro. Mismo reintento por
 * archivo bloqueado que agregarFilasSinSesionInterno.
 * @param {Array<{hoja: string, transformarFila: (fila: any[]) => (any[]|null)}>} especificaciones
 */
async function transformarFilasSinSesionInterno(token, rutaArchivo, especificaciones) {
  return encolarPorArchivo(rutaArchivo, () => intentarTransformarFilas(token, rutaArchivo, especificaciones));
}
async function intentarTransformarFilas(token, rutaArchivo, especificaciones) {
  for (let intento = 1; intento <= 4; intento++) {
    try {
      const buffer = await descargarArchivoOneDrive(token, rutaArchivo);
      const libro = XLSX.read(buffer, { type: "array" });
      for (const { hoja: nombreHoja, transformarFila } of especificaciones) {
        const hoja = libro.Sheets[nombreHoja];
        if (!hoja) continue;
        const [encabezado, ...filas] = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" });
        const filasTransformadas = filas.map(transformarFila).filter((fila) => fila !== null);
        const filasFinal = encabezado ? [encabezado, ...filasTransformadas] : filasTransformadas;
        libro.Sheets[nombreHoja] = XLSX.utils.aoa_to_sheet(filasFinal);
      }
      const bufferNuevo = XLSX.write(libro, { type: "array", bookType: "xlsx" });
      const blob = new Blob([bufferNuevo], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      return await subirArchivoOneDrive(token, rutaArchivo, blob, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } catch (error) {
      if (/423|locked/i.test(String(error?.message || error)) && intento < 4) {
        console.warn(`[OneDrive] Archivo bloqueado, reintentando (${intento}/3)…`);
        await esperar(4000);
        continue;
      }
      throw error;
    }
  }
}

/**
 * Reemplaza TODO el contenido (menos el encabezado) de una o varias hojas
 * de un mismo libro por un conjunto nuevo de filas — para datos que son
 * una "foto" del estado actual (cartera, cuentas por pagar, proyectos,
 * terceros, nómina mensual), no un historial que solo crece. Mismo
 * reintento por archivo bloqueado que las funciones de arriba.
 * @param {Array<{hoja: string, filasNuevas: any[][]}>} especificaciones
 */
async function reemplazarFilasVariasHojas(token, rutaArchivo, especificaciones) {
  return encolarPorArchivo(rutaArchivo, () => intentarReemplazarFilas(token, rutaArchivo, especificaciones));
}
async function intentarReemplazarFilas(token, rutaArchivo, especificaciones) {
  for (let intento = 1; intento <= 4; intento++) {
    try {
      const buffer = await descargarArchivoOneDrive(token, rutaArchivo);
      const libro = XLSX.read(buffer, { type: "array" });
      for (const { hoja: nombreHoja, filasNuevas } of especificaciones) {
        const hoja = libro.Sheets[nombreHoja];
        if (!hoja) continue;
        const [encabezado] = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" });
        const filasFinal = encabezado ? [encabezado, ...filasNuevas] : filasNuevas;
        libro.Sheets[nombreHoja] = XLSX.utils.aoa_to_sheet(filasFinal);
      }
      const bufferNuevo = XLSX.write(libro, { type: "array", bookType: "xlsx" });
      const blob = new Blob([bufferNuevo], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      return await subirArchivoOneDrive(token, rutaArchivo, blob, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } catch (error) {
      if (/423|locked/i.test(String(error?.message || error)) && intento < 4) {
        console.warn(`[OneDrive] Archivo bloqueado, reintentando (${intento}/3)…`);
        await esperar(4000);
        continue;
      }
      throw error;
    }
  }
}

// ============================================================================
// FUNCIONES DE CONVENIENCIA — cotizaciones, registros, nómina, terceros…
// ============================================================================

/** Envía una cotización (ya calculada) como fila nueva a la hoja "Datos", y cada ítem a la hoja "Items". */
export async function sincronizarCotizacion(instance, account, cotizacion) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.cotizaciones;
  const numeroCompleto = `${cotizacion.numero}-${cotizacion.revision}`;
  const motivoVersion = cotizacion.motivoVersion || "";
  const fila = [
    numeroCompleto,
    cotizacion.fecha,
    cotizacion.clienteNombre || "",
    cotizacion.referencia,
    cotizacion.estado,
    cotizacion.subtotalBruto ?? cotizacion.subtotal,
    cotizacion.iva,
    cotizacion.total,
    motivoVersion,
    "",
  ];
  await agregarFilasSinSesionInterno(token, ruta, hoja, [fila]);

  // Cada ítem de la cotización también se envía como su propia fila, en la
  // hoja "Items" — sin esto, el Excel solo mostraba el total, sin el
  // detalle de qué se cotizó exactamente.
  if (Array.isArray(cotizacion.items) && cotizacion.items.length) {
    const { ruta: rutaItems, hoja: hojaItems } = ARCHIVOS_EXCEL.cotizacionItems;
    const filasItems = cotizacion.items.map((it) => {
      const cantidad = Number(it.cantidad) || 0;
      const precio = Number(it.precio) || 0;
      return [
        numeroCompleto, cotizacion.fecha, cotizacion.clienteNombre || "",
        it.descripcion || "", it.unidad || "", cantidad, precio, cantidad * precio,
        motivoVersion, cotizacion.estado, "",
      ];
    });
    await agregarFilasSinSesionInterno(token, rutaItems, hojaItems, filasItems);
  }
}

/**
 * Actualiza el estado (y, si aplica, el motivo de pérdida) de una
 * cotización ya sincronizada — busca la fila por su número completo tanto
 * en "Datos" como en "Items" y la corrige in-place, sin tocar el resto.
 */
export async function actualizarEstadoCotizacionEnGraph(instance, account, numeroCompleto, cambios = {}) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.cotizaciones;
  const { hoja: hojaItems } = ARCHIVOS_EXCEL.cotizacionItems;
  // Índices de columna (0-based) del estado y del motivo de pérdida en
  // cada hoja — ver el orden de columnas armado en sincronizarCotizacion.
  const COL_ESTADO_DATOS = 4;
  const COL_MOTIVO_DATOS = 9;
  const COL_ESTADO_ITEMS = 9;
  const COL_MOTIVO_ITEMS = 10;
  const aplicarCambios = (fila, colEstado, colMotivo) => {
    const filaNueva = [...fila];
    if (cambios.estado) filaNueva[colEstado] = cambios.estado;
    if ("motivoPerdida" in cambios) filaNueva[colMotivo] = cambios.motivoPerdida || "";
    return filaNueva;
  };
  await transformarFilasSinSesionInterno(token, ruta, [
    { hoja, transformarFila: (fila) => (fila[0] !== numeroCompleto ? fila : aplicarCambios(fila, COL_ESTADO_DATOS, COL_MOTIVO_DATOS)) },
    { hoja: hojaItems, transformarFila: (fila) => (fila[0] !== numeroCompleto ? fila : aplicarCambios(fila, COL_ESTADO_ITEMS, COL_MOTIVO_ITEMS)) },
  ]);
}

/** Envía un registro de operación (venta/compra/gasto/cobro/pago) como fila nueva. */
export async function sincronizarRegistro(instance, account, operacion) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.registros;
  const fila = [operacion.id, operacion.tipo, operacion.fecha, operacion.concepto, operacion.valor, operacion.iva || 0];
  return agregarFilasSinSesionInterno(token, ruta, hoja, [fila]);
}

/** Envía un pago de nómina como fila nueva. */
export async function sincronizarNomina(instance, account, nomina) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.nomina;
  const fila = [nomina.colaboradorNombre || nomina.colaboradorId, nomina.periodo || nomina.fecha, nomina.valor];
  return agregarFilasSinSesionInterno(token, ruta, hoja, [fila]);
}

/** Envía un ítem nuevo del catálogo de cotizaciones como fila nueva. */
export async function sincronizarItemCatalogo(instance, account, item) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.catalogoItems;
  return agregarFilasSinSesionInterno(token, ruta, hoja, [[item.descripcion, item.unidad, item.precio]]);
}

/** Envía una liquidación de cesantías, prima, vacaciones o de contrato — cada una a su propia hoja. */
export async function sincronizarPrestacion(instance, account, tipoPrestacion, valores) {
  const token = await obtenerTokenGraph(instance, account);
  const clave = { CESANTIAS: "cesantias", PRIMA: "prima", VACACIONES: "vacaciones", LIQUIDACION_FINAL: "liquidacionContrato" }[tipoPrestacion];
  if (!clave) return;
  const { ruta, hoja } = ARCHIVOS_EXCEL[clave];
  return agregarFilasSinSesionInterno(token, ruta, hoja, [valores]);
}

/** Sincroniza la nómina mensual consolidada — una fila por colaborador del mes (foto actual). */
export async function sincronizarNominaMensual(instance, account, filasDelMes) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.nominaMensual;
  return reemplazarFilasVariasHojas(token, ruta, [{ hoja, filasNuevas: filasDelMes }]);
}

/** Sincroniza la lista completa de clientes y proveedores (foto actual) — mismo libro, dos hojas. */
export async function sincronizarTercerosCompleto(instance, account, clientes, proveedores) {
  const token = await obtenerTokenGraph(instance, account);
  return reemplazarFilasVariasHojas(token, ARCHIVOS_EXCEL.terceroClientes.ruta, [
    { hoja: ARCHIVOS_EXCEL.terceroClientes.hoja, filasNuevas: clientes.map((c) => [c.nombre, c.nit || "", c.direccion || "", c.telefono || ""]) },
    { hoja: ARCHIVOS_EXCEL.terceroProveedores.hoja, filasNuevas: proveedores.map((p) => [p.nombre, p.nit || "", p.direccion || "", p.telefono || ""]) },
  ]);
}

/** Sincroniza el comparativo de proyectos (foto actual). */
export async function sincronizarProyectosCompleto(instance, account, filas) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.proyectos;
  return reemplazarFilasVariasHojas(token, ruta, [{ hoja, filasNuevas: filas }]);
}

/** Sincroniza la cartera de cuentas por cobrar (foto actual). */
export async function sincronizarCarteraCompleto(instance, account, filas) {
  const token = await obtenerTokenGraph(instance, account);
  const { ruta, hoja } = ARCHIVOS_EXCEL.cartera;
  return reemplazarFilasVariasHojas(token, ruta, [{ hoja, filasNuevas: filas }]);
}

/** Sincroniza cuentas por pagar — las 3 hojas del mismo libro (foto actual). */
export async function sincronizarCxpCompleto(instance, account, { comprasGastos, prestaciones, provision }) {
  const token = await obtenerTokenGraph(instance, account);
  return reemplazarFilasVariasHojas(token, ARCHIVOS_EXCEL.cxpComprasGastos.ruta, [
    { hoja: ARCHIVOS_EXCEL.cxpComprasGastos.hoja, filasNuevas: comprasGastos },
    { hoja: ARCHIVOS_EXCEL.cxpPrestaciones.hoja, filasNuevas: prestaciones },
    { hoja: ARCHIVOS_EXCEL.cxpProvision.hoja, filasNuevas: provision },
  ]);
}

// ============================================================================
// CALENDARIO (Outlook) — módulo de tareas/agenda
// ============================================================================

function sumarMinutos(hora, minutos) {
  const [h, m] = hora.split(":").map(Number);
  const total = Math.min(23 * 60 + 59, h * 60 + m + minutos);
  const hh = String(Math.floor(total / 60)).padStart(2, "0");
  const mm = String(total % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

function diaSiguiente(fechaISO) {
  const fecha = new Date(`${fechaISO}T00:00:00`);
  fecha.setDate(fecha.getDate() + 1);
  return fecha.toISOString().slice(0, 10);
}

function horaAMinutos(hora) {
  const [h, m] = hora.split(":").map(Number);
  return h * 60 + m;
}

function construirEvento(tarea) {
  const asunto = `${tarea.tipo}${tarea.tituloExtra ? " — " + tarea.tituloExtra : ""}`;
  const zonaHoraria = "America/Bogota";
  if (tarea.horaInicio) {
    const fin = tarea.horaFin && horaAMinutos(tarea.horaFin) > horaAMinutos(tarea.horaInicio) ? tarea.horaFin : sumarMinutos(tarea.horaInicio, 30);
    return {
      subject: asunto,
      body: { contentType: "text", content: tarea.nota || "" },
      start: { dateTime: `${tarea.fecha}T${tarea.horaInicio}:00`, timeZone: zonaHoraria },
      end: { dateTime: `${tarea.fecha}T${fin}:00`, timeZone: zonaHoraria },
    };
  }
  return {
    subject: asunto,
    body: { contentType: "text", content: tarea.nota || "" },
    isAllDay: true,
    start: { dateTime: `${tarea.fecha}T00:00:00`, timeZone: zonaHoraria },
    end: { dateTime: `${diaSiguiente(tarea.fecha)}T00:00:00`, timeZone: zonaHoraria },
  };
}

export async function crearEventoCalendario(instance, account, tarea) {
  const token = await obtenerTokenGraph(instance, account);
  const evento = construirEvento(tarea);
  return llamarGraph(token, "POST", `/users/${CORREO_CALENDARIO}/events`, evento);
}

export async function actualizarEventoCalendario(instance, account, tarea) {
  if (!tarea.eventoCalendarioId) return null;
  const token = await obtenerTokenGraph(instance, account);
  const evento = construirEvento(tarea);
  return llamarGraph(token, "PATCH", `/users/${CORREO_CALENDARIO}/events/${tarea.eventoCalendarioId}`, evento);
}

export async function eliminarEventoCalendario(instance, account, eventoCalendarioId) {
  if (!eventoCalendarioId) return null;
  const token = await obtenerTokenGraph(instance, account);
  try {
    return await llamarGraph(token, "DELETE", `/users/${CORREO_CALENDARIO}/events/${eventoCalendarioId}`);
  } catch (error) {
    if (String(error.message).includes("404")) return null;
    throw error;
  }
}
