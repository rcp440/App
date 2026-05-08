// ================================================================
// APPS SCRIPT — pegar en Google Sheets > Extensiones > Apps Script
// Desplegar como: Aplicación web
//   Ejecutar como: Yo
//   Quién tiene acceso: Cualquier usuario (incluso anónimo)
// Después de cambiar el código: Desplegar > Nueva versión
// ================================================================

var SHEET_NAME = 'Asistencia';

function doGet(e) {
  const cb = (e.parameter && e.parameter.callback) || '';

  // Traer personas por DNI
  if (e.parameter && e.parameter.action === 'getPersonas') {
    try {
      const result = getPersonasByDni(e.parameter.dni || '');
      return jsonpOutput(result, cb);
    } catch(err) {
      return jsonpOutput({ ok: false, error: err.toString() }, cb);
    }
  }

  // Borrar filas de una fecha+id (previo a re-guardar)
  if (e.parameter && e.parameter.action === 'clear') {
    try {
      borrarFilas(e.parameter.fecha || '', e.parameter.id || '');
      return jsonpOutput({ ok: true }, cb);
    } catch(err) {
      return jsonpOutput({ ok: false, error: err.toString() }, cb);
    }
  }

  // Agregar una fila individual
  if (e.parameter && e.parameter.action === 'add') {
    try {
      agregarFila(e.parameter);
      return jsonpOutput({ ok: true }, cb);
    } catch(err) {
      return jsonpOutput({ ok: false, error: err.toString() }, cb);
    }
  }

  return jsonpOutput({ ok: true, msg: 'Apps Script activo' }, cb);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    guardarDatos(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function jsonpOutput(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Fecha', 'Hora', 'ID', 'Maestro/a', 'Nombre', 'Celular', 'Estado', 'Observaciones']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#4A4090').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function borrarFilas(fecha, id) {
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const vals = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const toStr = d => (d instanceof Date) ? d.toISOString().slice(0, 10) : String(d).slice(0, 10);
  for (let i = vals.length - 1; i >= 0; i--) {
    if (toStr(vals[i][0]) === fecha && String(vals[i][2]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 2);
    }
  }
}

function agregarFila(p) {
  const sheet = ensureSheet();
  sheet.appendRow([p.fecha, p.hora, p.id || '', p.maestro, p.nombre, p.celular, p.estado === 'P' ? 'Presente' : 'Ausente', p.obs || '']);
  const row = sheet.getLastRow();
  sheet.getRange(row, 1, 1, 8).setBackground(p.estado === 'P' ? '#EBF8F3' : '#FCECEA');
  sheet.autoResizeColumns(1, 8);
}

// Devuelve personas de la última reunión del DNI dado
function getPersonasByDni(dni) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return { ok: true, personas: [], maestro: '' };

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  const filas = data.filter(r => String(r[2]).trim() === String(dni).trim());
  if (!filas.length) return { ok: true, personas: [], maestro: '' };

  const toStr = d => (d instanceof Date) ? d.toISOString().slice(0, 10) : String(d).slice(0, 10);
  const fechas = filas.map(r => toStr(r[0])).sort();
  const ultimaFecha = fechas[fechas.length - 1];

  const maestro = (filas.find(r => toStr(r[0]) === ultimaFecha) || [])[3] || '';

  const seen = new Set();
  const personas = filas
    .filter(r => toStr(r[0]) === ultimaFecha)
    .filter(r => { const n = String(r[4]); if (seen.has(n)) return false; seen.add(n); return true; })
    .map(r => ({ name: String(r[4]), phone: String(r[5]), obs: String(r[7] || '') }));

  return { ok: true, personas, maestro, fecha: ultimaFecha };
}

function guardarDatos(data) {
  const fecha   = data.fecha   || '';
  const id      = String(data.id || '').trim();
  borrarFilas(fecha, id);
  (data.rows || []).forEach(r => {
    agregarFila({ fecha, hora: data.hora || '', id, maestro: data.maestro || '', ...r, estado: r.estado === 'Presente' ? 'P' : 'A' });
  });
}
