import ExcelJS from "exceljs";
import JSZip from "jszip";
import { LOGO_NOMINA_BASE64 } from "./logoNominaBase64.js";

// ExcelJS (probado con la versión 4.4.0) tiene un bug conocido: al insertar
// una imagen, siempre escribe el tamaño interno de la forma
// (xdr:spPr/a:xfrm/a:ext) como cx="0" cy="0", sin importar el tamaño real
// con el que se ancló la imagen (xdr:ext, que sí queda correcto). Un
// tamaño interno 0x0 es contenido inválido para Excel — al abrir el
// archivo, Excel lo detecta, muestra el diálogo de "encontramos un
// problema con el contenido" y lo repara solo (quita la fórmula que sea
// vecina y reconstruye el dibujo), lo cual es alarmante para quien lo
// abre aunque el resultado final se vea bien. Como el bug está en el XML
// que arma la librería (no en la API pública), se corrige después de
// generar el archivo: se abre el .xlsx como zip, se reemplaza ese tamaño
// interno inválido por el tamaño real del logo, y se vuelve a comprimir.
const LOGO_ANCHO_PX = 243;
const LOGO_ALTO_PX = 99;
const EMU_POR_PIXEL = 9525; // constante fija de OOXML (96 DPI)
async function corregirTamanoImagenIncrustada(bufferXlsx) {
  const cx = LOGO_ANCHO_PX * EMU_POR_PIXEL;
  const cy = LOGO_ALTO_PX * EMU_POR_PIXEL;
  const zip = await JSZip.loadAsync(bufferXlsx);
  const rutasDibujos = Object.keys(zip.files).filter((ruta) => /^xl\/drawings\/drawing\d+\.xml$/.test(ruta));
  for (const ruta of rutasDibujos) {
    const xmlOriginal = await zip.file(ruta).async("string");
    const xmlCorregido = xmlOriginal.replace(/<a:ext cx="0" cy="0"\/>/g, `<a:ext cx="${cx}" cy="${cy}"/>`);
    zip.file(ruta, xmlCorregido);
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// ============================================================================
// Generador del libro de nómina mensual, replicando exactamente la plantilla
// oficial de la empresa (Plantilla_app.xlsm): logo, cajas de encabezado,
// tabla de colaboradores, fila de totales, "Neto pagado" en letras, caja de
// provisiones patronales y líneas de firma.
//
// Se genera como .xlsx normal (sin macros) usando ExcelJS, con fórmulas de
// Excel reales donde la plantilla las usa (totales, fondo de solidaridad
// pensional, neto, provisiones patronales), así el archivo queda auditable
// y editable en Excel igual que el original.
// ============================================================================

// ---- Conversión de número a letras (Colombia, pesos) ----
// Portado de la macro VBA LETRAS/NUMERORECURSIVO de la plantilla original,
// verificado carácter por carácter contra el ejemplo real de la plantilla
// (27.738.318 -> "Veintisiete Millones Setecientos Treinta y Ocho Mil
// Trescientos Dieciocho Pesos"). ExcelJS no puede evaluar macros VBA, así
// que este texto se calcula aquí mismo en JS y se escribe como valor fijo.
const UNIDADES = ["", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veintiuno", "Veintidós", "Veintitrés", "Veinticuatro", "Veinticinco", "Veintiséis", "Veintisiete", "Veintiocho", "Veintinueve"];
const DECENAS = ["", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa", "Cien"];
const CENTENAS = ["", "Ciento", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", "Seiscientos", "Setecientos", "Ochocientos", "Novecientos"];

function numeroRecursivo(n) {
  n = Math.floor(n);
  if (n === 0) return "Cero";
  if (n <= 29) return UNIDADES[n];
  if (n <= 100) {
    const d = Math.floor(n / 10), r = n % 10;
    return DECENAS[d] + (r !== 0 ? " y " + numeroRecursivo(r) : "");
  }
  if (n <= 999) {
    const c = Math.floor(n / 100), r = n % 100;
    return CENTENAS[c] + (r !== 0 ? " " + numeroRecursivo(r) : "");
  }
  if (n <= 1999) {
    const r = n % 1000;
    return "Mil" + (r !== 0 ? " " + numeroRecursivo(r) : "");
  }
  if (n <= 999999) {
    const r = n % 1000;
    return numeroRecursivo(Math.floor(n / 1000)) + " Mil" + (r !== 0 ? " " + numeroRecursivo(r) : "");
  }
  if (n <= 1999999) {
    const r = n % 1000000;
    return "Un Millón" + (r !== 0 ? " " + numeroRecursivo(r) : "");
  }
  const r = n % 1000000;
  return numeroRecursivo(Math.floor(n / 1000000)) + " Millones" + (r !== 0 ? " " + numeroRecursivo(r) : "");
}

export function numeroALetras(valor) {
  const numero = Math.abs(Math.round(valor));
  if (numero > 1999999999) return "ERROR: El número excede los límites.";
  const letra = numeroRecursivo(numero);
  return letra + " " + (numero === 1 ? "Peso" : "Pesos");
}

// ---- Colores exactos extraídos de la plantilla real (tema Office estándar) ----
const COLOR_HEADER_AZUL_OSCURO = "FF2F5597"; // Azul, Énfasis 1, Oscuro 25%
const COLOR_CAJA_GRIS_CLARO = "FFF2F2F2"; // Blanco, Fondo 1, Oscuro 5%
const COLOR_DEVENGADO_AZUL = "FF9DC3E6"; // Azul, Énfasis 5, Claro 40%
const COLOR_DEDUCCIONES_AZUL = "FFBDD7EE"; // Azul, Énfasis 5, Claro 60%
const COLOR_NETO_AZUL = "FFB4C7E7"; // Azul, Énfasis 1, Claro 60%
const COLOR_TOTALES_AZUL_CLARO = "FFDAE3F3"; // Azul, Énfasis 1, Claro 80%
const COLOR_PROVISION_NARANJA = "FFC55A11"; // Naranja, Énfasis 2, Oscuro 25%
const COLOR_BLANCO = "FFFFFFFF";
const COLOR_NEGRO = "FF000000";

const FMT_MONEDA = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-';

function borde(estilo = "thin") {
  return { style: estilo, color: { argb: COLOR_NEGRO } };
}
function bordeCompleto(estilo = "thin") {
  return { top: borde(estilo), left: borde(estilo), bottom: borde(estilo), right: borde(estilo) };
}

/**
 * Genera el libro de Excel de nómina mensual replicando exactamente la
 * plantilla oficial de la empresa (logo, colores, fórmulas, caja de
 * provisiones patronales, "Neto pagado" en letras, líneas de firma).
 *
 * @param {object} datos
 * @param {Array}  datos.filas - una fila por colaborador, con los mismos
 *   campos que ya calcula exportarNominaMensualExcel en App.jsx.
 * @param {string} datos.mesLabel - texto para el período, ej. "1 MARZO 2026 - MARZO 31 2026"
 * @param {Date|string} datos.fechaLiquidacion
 * @param {number} datos.smmlv
 * @param {number} datos.auxilioTransporteLey
 * @param {boolean} datos.exoneradoAportes
 * @param {Array}  [datos.liquidacionesDelMes] - filas informativas de
 *   prestaciones sociales realmente causadas/pagadas ese mes (hoja aparte).
 * @returns {Promise<Uint8Array>} los bytes finales del .xlsx, ya listos
 *   para descargarse o subirse — no el objeto Workbook de ExcelJS, porque
 *   esta función también corrige el bug de la imagen (ver
 *   corregirTamanoImagenIncrustada) después de serializar el libro.
 */
export async function generarLibroNominaMensual({ filas, mesLabel, fechaLiquidacion, smmlv, auxilioTransporteLey, exoneradoAportes, liquidacionesDelMes }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Nómina", { views: [{ showGridLines: false }] });

  // ---- Anchos de columna (tomados de la plantilla real) ----
  const anchos = { A: 2.16, B: 22.66, C: 27.5, D: 14.33, E: 11.16, F: 12.16, G: 10.16, H: 17.66, I: 11.16, J: 11.33, K: 18.16, L: 10.33, M: 10.5, N: 14.83, O: 10.5, P: 10.5, Q: 13.33, R: 13.66, S: 14.16, T: 12, U: 16, V: 14.66, W: 22.16, X: 15.16 };
  Object.entries(anchos).forEach(([col, width]) => { ws.getColumn(col).width = width; });

  const logoId = wb.addImage({ base64: `data:image/jpeg;base64,${LOGO_NOMINA_BASE64}`, extension: "jpeg" });
  ws.addImage(logoId, { tl: { col: 1, row: 0 }, ext: { width: 243, height: 99 } });
  ws.mergeCells("B1:I4");

  // ---- Caja "Diligencie los siguientes valores de referencia" ----
  ws.mergeCells("J1:N1");
  ws.getCell("J1").value = "Diligencie los siguientes valores de referencia:";
  ws.mergeCells("J2:L2"); ws.getCell("J2").value = "Salario mínimo";
  ws.getCell("N2").value = smmlv;
  ws.mergeCells("J3:L3"); ws.getCell("J3").value = "Auxilio de transporte";
  ws.getCell("N3").value = auxilioTransporteLey;
  ws.mergeCells("J4:L4"); ws.getCell("J4").value = "¿Está exonerado del pago de aportes?";
  ws.getCell("N4").value = exoneradoAportes ? "Si" : "No";
  ["N2", "N3"].forEach((c) => { ws.getCell(c).numFmt = FMT_MONEDA; });
  ws.mergeCells("O1:R4"); // espaciador

  // ---- Caja "Periodo de pago" ----
  ws.mergeCells("S1:W1"); ws.getCell("S1").value = "Periodo de pago";
  ws.mergeCells("S2:W2"); ws.getCell("S2").value = mesLabel;
  ws.mergeCells("S3:W3"); ws.getCell("S3").value = "Fecha de liquidación";
  ws.mergeCells("S4:W4");
  ws.getCell("S4").value = fechaLiquidacion instanceof Date ? fechaLiquidacion : new Date(fechaLiquidacion);
  ws.getCell("S4").numFmt = "d/mm/yyyy";

  // Estilos de las dos cajas de encabezado
  [["J1", true], ["S1", true]].forEach(([c]) => {
    const cell = ws.getCell(c);
    cell.font = { bold: true, size: c === "J1" ? 13 : 18, color: { argb: COLOR_BLANCO } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_AZUL_OSCURO } };
    cell.alignment = { horizontal: "center", vertical: "center" };
  });
  ["J2", "J3", "J4"].forEach((c) => { ws.getCell(c).font = { bold: true, size: 12 }; ws.getCell(c).alignment = { vertical: "center" }; });
  ["N2", "N3", "N4"].forEach((c) => { ws.getCell(c).alignment = { horizontal: "center", vertical: "center" }; });
  ["S2", "S3", "S4"].forEach((c) => {
    const cell = ws.getCell(c);
    cell.font = { bold: c === "S3", size: c === "S2" ? 14 : 11 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_CAJA_GRIS_CLARO } };
    cell.alignment = { horizontal: "center", vertical: "center" };
  });
  ["J1", "J2", "J3", "J4", "N2", "N3", "N4", "S1", "S2", "S3", "S4"].forEach((c) => { ws.getCell(c).border = bordeCompleto(); });

  ws.getRow(1).height = 25; ws.getRow(2).height = 25; ws.getRow(3).height = 25; ws.getRow(4).height = 25; ws.getRow(5).height = 15;

  // ---- Encabezados de la tabla (filas 6-7) ----
  ws.mergeCells("B6:B7"); ws.getCell("B6").value = "Nombre del\nempleado";
  ws.mergeCells("D5:N5"); ws.mergeCells("O5:S5"); // filas espaciadoras
  ws.mergeCells("D6:N6"); ws.getCell("D6").value = "Devengado";
  ws.mergeCells("O6:S6"); ws.getCell("O6").value = "Deducciones";
  ws.mergeCells("V6:V7"); ws.getCell("V6").value = "Neto\n pagado";
  ws.mergeCells("W6:W7"); ws.getCell("W6").value = "Firma del \n empleado";

  const encabezadosFila7 = {
    C7: "Cedula", D7: "Salario \nbásico", E7: "Días \nliquidados", F7: "Salario \ndevengado",
    G7: "Incapacidad general", H7: "Incapaciadad laboral", I7: "Horas \nextras", J7: "vacaciones",
    K7: "Trabajo dominical \ny festivo", L7: "Auxilio de\ntransporte", M7: "Comisiones", N7: "Total\ndevengado",
    O7: "Salud", P7: "Pensión", Q7: "Fondo de solidaridad\npensional", R7: "EMPRESA", S7: "Otras\n deducciones",
    T7: "Intereses cesantias", U7: "Auxilio no salarial / reconocimiento extra legal ", X7: "Verificacion pago",
  };
  Object.entries(encabezadosFila7).forEach(([c, v]) => { ws.getCell(c).value = v; });

  ["B6", "C7"].forEach((c) => { ws.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_AZUL_OSCURO } }; ws.getCell(c).font = { bold: true, color: { argb: COLOR_BLANCO } }; });
  ["D6"].forEach((c) => { ws.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_DEVENGADO_AZUL } }; ws.getCell(c).font = { bold: true }; });
  ["O6"].forEach((c) => { ws.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_DEDUCCIONES_AZUL } }; ws.getCell(c).font = { bold: true }; });
  ["V6", "W6"].forEach((c) => { ws.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_NETO_AZUL } }; ws.getCell(c).font = { bold: true }; });
  Object.keys(encabezadosFila7).forEach((c) => {
    const cell = ws.getCell(c);
    if (!cell.font?.bold) cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
  });
  ["B6", "D6", "O6", "V6", "W6"].forEach((c) => { ws.getCell(c).alignment = { horizontal: "center", vertical: "center", wrapText: true }; });
  for (const col of ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X"]) {
    ws.getCell(`${col}6`).border = bordeCompleto();
    ws.getCell(`${col}7`).border = bordeCompleto();
  }
  ws.getRow(7).height = 55;

  // ---- Filas de colaboradores (dinámico) ----
  const filaInicio = 8;
  filas.forEach((f, i) => {
    const r = filaInicio + i;
    ws.getCell(`B${r}`).value = f.Nombre_del_empleado;
    ws.getCell(`C${r}`).value = f.Cedula;
    ws.getCell(`D${r}`).value = f.Salario_basico;
    ws.getCell(`E${r}`).value = f.Dias_liquidados;
    ws.getCell(`F${r}`).value = f.Salario_devengado;
    ws.getCell(`G${r}`).value = f.Incapacidad_general;
    ws.getCell(`H${r}`).value = f.Incapacidad_laboral;
    ws.getCell(`I${r}`).value = f.Horas_extras;
    ws.getCell(`J${r}`).value = f.Vacaciones;
    ws.getCell(`K${r}`).value = f.Trabajo_dominical_y_festivo;
    ws.getCell(`L${r}`).value = f.Auxilio_de_transporte;
    ws.getCell(`M${r}`).value = f.Comisiones;
    ws.getCell(`N${r}`).value = { formula: `SUM(F${r}:M${r})` };
    ws.getCell(`O${r}`).value = f.Salud;
    ws.getCell(`P${r}`).value = f.Pension;
    // Fondo de solidaridad pensional: fórmula legal por tramos de IBC/SMMLV (igual a la plantilla)
    ws.getCell(`Q${r}`).value = {
      formula: `(N${r}-L${r})*(IF(AND(((N${r}-L${r})/$N$2)>=4,((N${r}-L${r})/$N$2)<16),1,IF(AND(((N${r}-L${r})/$N$2)>=16,((N${r}-L${r})/$N$2)<=17),1.2,IF(AND(((N${r}-L${r})/$N$2)>17,((N${r}-L${r})/$N$2)<=18),1.4,IF(AND(((N${r}-L${r})/$N$2)>18,((N${r}-L${r})/$N$2)<=19),1.6,IF(AND(((N${r}-L${r})/$N$2)>19,((N${r}-L${r})/$N$2)<=20),1.8,IF(((N${r}-L${r})/$N$2)>20,2,0)))))))/100`,
    };
    ws.getCell(`S${r}`).value = f.Otras_deducciones;
    ws.getCell(`T${r}`).value = f.Intereses_cesantias;
    ws.getCell(`U${r}`).value = f.Auxilio_no_salarial;
    ws.getCell(`V${r}`).value = { formula: `N${r}-O${r}-P${r}-Q${r}-R${r}-S${r}+U${r}+T${r}` };
    ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V"].forEach((col) => { ws.getCell(`${col}${r}`).numFmt = FMT_MONEDA; });
    ["E"].forEach((col) => { ws.getCell(`${col}${r}`).numFmt = "0"; });
    for (const col of ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X"]) {
      ws.getCell(`${col}${r}`).border = bordeCompleto();
    }
  });

  const ultimaFilaColaborador = filaInicio + filas.length - 1;
  const filaTotales = ultimaFilaColaborador + 1;

  // ---- Fila de Totales ----
  ws.mergeCells(`B${filaTotales}:E${filaTotales}`);
  ws.getCell(`B${filaTotales}`).value = "Totales";
  ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V"].forEach((col) => {
    ws.getCell(`${col}${filaTotales}`).value = { formula: `SUM(${col}${filaInicio}:${col}${ultimaFilaColaborador})` };
    ws.getCell(`${col}${filaTotales}`).numFmt = FMT_MONEDA;
  });
  for (const col of ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X"]) {
    const cell = ws.getCell(`${col}${filaTotales}`);
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_TOTALES_AZUL_CLARO } };
    cell.border = bordeCompleto();
  }
  ws.getCell(`B${filaTotales}`).alignment = { horizontal: "center" };

  // ---- Caja "Neto pagado: $X" + en letras ----
  const filaNeto = filaTotales + 2;
  ws.mergeCells(`B${filaNeto}:B${filaNeto + 1}`);
  ws.getCell(`B${filaNeto}`).value = { formula: `ROUND(V${filaTotales},0)` };
  ws.getCell(`B${filaNeto}`).numFmt = '"Neto pagado:" "$"0#,##0;\\-#,##0';
  ws.mergeCells(`E${filaNeto}:Q${filaNeto + 1}`);
  // ExcelJS no puede evaluar la macro LETRAS al vuelo — el texto en letras
  // se calcula aquí mismo en JS y se escribe como valor fijo, así el
  // archivo no depende de macros ni de "habilitar contenido" para verse bien.
  const totalNeto = filas.reduce((s, f) => s + (f.Neto_pagado || 0), 0);
  ws.getCell(`E${filaNeto}`).value = numeroALetras(totalNeto);
  [`B${filaNeto}`, `E${filaNeto}`].forEach((c) => {
    const cell = ws.getCell(c);
    cell.font = { bold: true, size: 16 };
    cell.alignment = { horizontal: c.startsWith("B") ? "center" : "right", vertical: "center" };
    cell.border = bordeCompleto("medium");
  });

  // ---- Caja de provisiones patronales (columna derecha) ----
  const filaProvisiones0 = filaNeto + 3;
  ws.mergeCells(`N${filaProvisiones0}:Q${filaProvisiones0 + 1}`);
  ws.getCell(`N${filaProvisiones0}`).value = "Provisiones de nómina a \ncargo del empleador.";
  ws.getCell(`N${filaProvisiones0}`).font = { bold: true, size: 13, color: { argb: COLOR_BLANCO } };
  ws.getCell(`N${filaProvisiones0}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_AZUL_OSCURO } };
  ws.getCell(`N${filaProvisiones0}`).alignment = { horizontal: "center", vertical: "center", wrapText: true };

  const nRef = `N${filaTotales}`, lRef = `L${filaTotales}`, iRef = `I${filaTotales}`, kRef = `K${filaTotales}`;
  const provisiones = [
    ["Aportes a pensión", `(${nRef}-${lRef})*0.12`],
    ["Aportes a salud", `IF(N4="Si",(SUMIF(N${filaInicio}:N${ultimaFilaColaborador},">"&N2*10,N${filaInicio}:N${ultimaFilaColaborador}))*0.085,(${nRef}-${lRef})*0.085)`],
    ["Aportes a riesgos laborales", `(${nRef}-${lRef})*0.0696`],
    ["Sena", `IF(N4="Si",(SUMIF(N${filaInicio}:N${ultimaFilaColaborador},">"&N2*10,N${filaInicio}:N${ultimaFilaColaborador}))*0.02,(${nRef}-${lRef})*0.02)`],
    ["Icbf", `IF(N4="Si",(SUMIF(N${filaInicio}:N${ultimaFilaColaborador},">"&N2*10,N${filaInicio}:N${ultimaFilaColaborador}))*0.03,(${nRef}-${lRef})*0.03)`],
    ["Cajas de compensación", `(${nRef}-${lRef})*0.04`],
    ["Prima de servicios", `${nRef}*0.0833`],
    ["Cesantía", `${nRef}*0.0833`],
    ["Intereses sobre cesantías", null], // se calcula abajo referenciando la fila de Cesantía
    ["Provisón de vacaciones", `(${nRef}-${iRef}-${kRef}-${lRef})*0.0417`],
  ];

  let filaCursor = filaProvisiones0 + 2;
  provisiones.forEach(([label, formula], idx) => {
    const r = filaProvisiones0 + 2 + idx;
    ws.mergeCells(`N${r}:P${r}`);
    ws.getCell(`N${r}`).value = label;
    ws.getCell(`N${r}`).border = bordeCompleto();
    if (label === "Intereses sobre cesantías") {
      const filaCesantiaReal = filaProvisiones0 + 2 + provisiones.findIndex(([l]) => l === "Cesantía");
      ws.getCell(`Q${r}`).value = { formula: `Q${filaCesantiaReal}*0.12` };
    } else {
      ws.getCell(`Q${r}`).value = { formula };
    }
    ws.getCell(`Q${r}`).numFmt = FMT_MONEDA;
    ws.getCell(`Q${r}`).border = bordeCompleto();
    filaCursor = r;
  });
  const filaTotalProv = filaCursor + 1;
  ws.mergeCells(`N${filaTotalProv}:P${filaTotalProv}`);
  ws.getCell(`N${filaTotalProv}`).value = "Total provisiones";
  const primeraFilaProv = filaProvisiones0 + 2;
  ws.getCell(`Q${filaTotalProv}`).value = { formula: `SUM(Q${primeraFilaProv}:Q${filaCursor})` };
  ws.getCell(`Q${filaTotalProv}`).numFmt = FMT_MONEDA;
  [`N${filaTotalProv}`, `Q${filaTotalProv}`].forEach((c) => {
    ws.getCell(c).font = { bold: true, color: { argb: COLOR_BLANCO } };
    ws.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_PROVISION_NARANJA } };
    ws.getCell(c).border = bordeCompleto();
  });

  // ---- Comprobante Nº / Fecha / Observaciones (columna izquierda) ----
  const filaComprobante = filaNeto + 3;
  ws.getCell(`B${filaComprobante}`).value = "Comprobante de Pago Nº:";
  ws.mergeCells(`D${filaComprobante}:K${filaComprobante}`);
  const filaFecha = filaComprobante + 2;
  ws.getCell(`B${filaFecha}`).value = "Fecha de";
  ws.mergeCells(`D${filaFecha}:K${filaFecha}`);
  ws.getCell(`D${filaFecha}`).value = fechaLiquidacion instanceof Date ? fechaLiquidacion : new Date(fechaLiquidacion);
  ws.getCell(`D${filaFecha}`).numFmt = "d/mm/yyyy";
  const filaObs = filaFecha + 2;
  ws.getCell(`B${filaObs}`).value = "Observaciones";
  const filaObsFin = filaObs + 3;
  ws.mergeCells(`D${filaObs}:K${filaObsFin}`);
  [`B${filaComprobante}`, `B${filaFecha}`, `B${filaObs}`].forEach((c) => { ws.getCell(c).font = { bold: true }; });
  [`B${filaComprobante}`, `D${filaComprobante}`, `B${filaFecha}`, `D${filaFecha}`, `B${filaObs}`, `D${filaObs}`].forEach((c) => { ws.getCell(c).border = bordeCompleto("medium"); });

  // ---- Firmas ----
  // Se colocan debajo de lo que termine más abajo entre la columna
  // izquierda (Comprobante/Fecha/Observaciones) y la caja de provisiones a
  // la derecha, para que nunca se encimen sin importar cuántas filas tenga
  // cada una.
  const filaFirmas = Math.max(filaObsFin, filaTotalProv) + 2;
  ws.getCell(`B${filaFirmas}`).value = "Elaborador Por:";
  ws.mergeCells(`D${filaFirmas}:I${filaFirmas}`);
  ws.getCell(`D${filaFirmas}`).value = "Revisado Por:";
  ws.mergeCells(`J${filaFirmas}:K${filaFirmas}`);
  ws.getCell(`J${filaFirmas}`).value = "Aprobado Por:";
  [`B${filaFirmas}`, `D${filaFirmas}`, `J${filaFirmas}`].forEach((c) => {
    ws.getCell(c).font = { bold: true };
    ws.getCell(c).alignment = { horizontal: "center" };
    ws.getCell(c).border = bordeCompleto("medium");
  });
  for (let i = 1; i <= 3; i++) {
    const r = filaFirmas + i;
    ws.mergeCells(`D${r}:I${r}`);
    ws.mergeCells(`J${r}:K${r}`);
    [`B${r}`, `D${r}`, `J${r}`].forEach((c) => { ws.getCell(c).border = bordeCompleto(); });
  }

  // ---- Hoja informativa: Liquidaciones del mes (prestaciones realmente
  // causadas/pagadas ese mes, vía módulo Prestaciones sociales) ----
  if (liquidacionesDelMes && liquidacionesDelMes.length > 0) {
    const ws2 = wb.addWorksheet("Liquidaciones del mes");
    const columnas = [
      { header: "Colaborador", key: "Colaborador", width: 28 },
      { header: "Tipo", key: "Tipo", width: 20 },
      { header: "Fecha corte", key: "Fecha_corte", width: 14 },
      { header: "Cesantías al fondo", key: "Cesantias_al_fondo", width: 18 },
      { header: "Intereses cesantías", key: "Intereses_cesantias", width: 18 },
      { header: "Prima", key: "Prima", width: 14 },
      { header: "Días vacaciones causados", key: "Dias_vacaciones_causados", width: 20 },
      { header: "Valor vacaciones", key: "Valor_vacaciones", width: 16 },
      { header: "Valor pagado al trabajador", key: "Valor_pagado_al_trabajador", width: 20 },
      { header: "Estado", key: "Estado", width: 14 },
    ];
    ws2.columns = columnas;
    ws2.getRow(1).font = { bold: true, color: { argb: COLOR_BLANCO } };
    ws2.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_AZUL_OSCURO } };
    liquidacionesDelMes.forEach((l) => ws2.addRow(l));
    ["Cesantias_al_fondo", "Intereses_cesantias", "Prima", "Valor_vacaciones", "Valor_pagado_al_trabajador"].forEach((key) => {
      const col = ws2.getColumn(columnas.findIndex((c) => c.key === key) + 1);
      col.numFmt = FMT_MONEDA;
    });
  }

  const bufferBruto = await wb.xlsx.writeBuffer();
  return corregirTamanoImagenIncrustada(bufferBruto);
}
