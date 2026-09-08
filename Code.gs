/**
 * TEASCHILE — Registro de Actividades (versión compartida en línea)
 * Backend en Google Apps Script. Guarda cada registro en una Google Sheet
 * que se crea automáticamente la primera vez que alguien usa la app,
 * y que sirve como base de datos compartida para todo el equipo.
 */

const SHEET_NAME = "ACTIVIDADES PERIODO";
const HEADERS = ["FECHA","MES","ASESOR","EQUIPO","ACTIVIDAD","DESCRIPCIÓN","OPERATIVO","N° SERIE","STATUS ACTIVIDAD","COMENTARIOS","RECOMENDACIONES","REGISTRADO POR","FECHA REGISTRO"];

/**
 * Sirve la app (si no viene ?action=...) o responde como API JSON
 * (si viene ?action=listar|exportar|sheetUrl), para que un HTML puro
 * hospedado en cualquier lado pueda usar fetch() en vez del puente
 * interno de Google (google.script.run), que falla en algunos
 * navegadores integrados (ej. el de WhatsApp).
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Registro de Actividades — TEASCHILE')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  try {
    let data;
    if (action === 'listar') {
      data = listarRegistros(e.parameter.fecha || "");
    } else if (action === 'exportar') {
      data = exportarRango(e.parameter.desde || "", e.parameter.hasta || "");
    } else if (action === 'sheetUrl') {
      data = getSheetUrl();
    } else {
      return jsonOutput_({ ok:false, error: "Acción GET desconocida: " + action });
    }
    return jsonOutput_({ ok:true, data: data });
  } catch (err) {
    return jsonOutput_({ ok:false, error: String(err) });
  }
}

/** Recibe guardar/eliminar desde un HTML puro vía fetch (POST, body texto plano con JSON). */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || {};
    let data;
    if (action === 'guardar') {
      data = guardarRegistro(payload);
    } else if (action === 'eliminar') {
      data = eliminarRegistro(Number(payload.rowIndex));
    } else {
      return jsonOutput_({ ok:false, error: "Acción POST desconocida: " + action });
    }
    return jsonOutput_({ ok:true, data: data });
  } catch (err) {
    return jsonOutput_({ ok:false, error: String(err) });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Obtiene (o crea la primera vez) la planilla compartida y su hoja. */
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  let ss;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); }
    catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('TEASCHILE - Registro de Actividades');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }
  if (sheet.getRange(1,1).getValue() !== HEADERS[0]) {
    sheet.clear();
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange("A:A").setNumberFormat("@"); // FECHA como texto plano (evita que Sheets la reformatee)
    sheet.getRange("M:M").setNumberFormat("@"); // FECHA REGISTRO como texto plano
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return sheet;
}

/** Guarda un nuevo registro. entry es un objeto con los campos del formulario. */
function guardarRegistro(entry) {
  const sheet = getSheet_();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
  sheet.appendRow([
    entry.fecha || "", entry.mes || "", entry.asesor || "", entry.equipo || "",
    entry.actividad || "", entry.descripcion || "", entry.operativo || "",
    entry.serie || "", entry.status || "", entry.comentarios || "",
    entry.recomendaciones || "", entry.registradoPor || "", now
  ]);
  return { ok: true };
}

/** Lista los registros de una fecha (formato yyyy-MM-dd). Si fecha es "", trae todo. */
function listarRegistros(fecha) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const out = [];
  values.forEach((r, i) => {
    if (!fecha || String(r[0]) === fecha) {
      out.push({
        rowIndex: i + 2,
        fecha: r[0], mes: r[1], asesor: r[2], equipo: r[3], actividad: r[4],
        descripcion: r[5], operativo: r[6], serie: r[7], status: r[8],
        comentarios: r[9], recomendaciones: r[10], registradoPor: r[11], registradoEn: r[12]
      });
    }
  });
  out.reverse();
  return out;
}

/** Elimina un registro por su número de fila real en la hoja. */
function eliminarRegistro(rowIndex) {
  const sheet = getSheet_();
  sheet.deleteRow(rowIndex);
  return { ok: true };
}

/** Devuelve filas (arreglos) entre dos fechas yyyy-MM-dd, para exportar. */
function exportarRango(desde, hasta) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return values
    .filter(r => (!desde || String(r[0]) >= desde) && (!hasta || String(r[0]) <= hasta))
    .map(r => r.slice(0, 11)); // sin las columnas de auditoría
}

/** URL de la planilla en Google Sheets, para verla/editarla completa. */
function getSheetUrl() {
  return getSheet_().getParent().getUrl();
}
