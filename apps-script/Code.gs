const SHEET_ID = '1Evr1lpNSwLYXWgcRf5D5NSPUD5nR-YS_EczVK4PVAHI';
const SHEET_GIDS = {
  'Mantenimiento': 0,
  'Lavado': 819388144,
  'Engrasada': 2024356449,
  'Cambio de aceite': 464443967,
  'Control horometro': null
};

function doGet() {
  return jsonResponse({ ok: true, message: 'Punto PAS API activa' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetName = data.tipoServicio;
    const sheet = sheetName === 'Control horometro' ? getOrCreateHourMeterSheet() : getSheetByGid(SHEET_GIDS[sheetName]);

    if (!sheet) {
      return jsonResponse({ ok: false, error: 'No existe la hoja: ' + sheetName });
    }

    const item = Math.max(sheet.getLastRow(), 1);
    if (sheetName === 'Control horometro') {
      sheet.appendRow([
        item,
        data.fecha || '',
        data.lugar || '',
        data.chofer || '',
        Number(data.horometroAnterior || 0),
        Number(data.horometroActual || 0),
        Number(data.horasTrabajadas || 0),
        String(data.placa || '').toUpperCase(),
        String(data.cambioAceite || 'NO').toUpperCase(),
        data.descripcion || ''
      ]);
      return jsonResponse({ ok: true });
    }
    sheet.appendRow([
      item,
      data.fecha || '',
      data.lugar || '',
      data.maestro || '',
      data.taller || '',
      data.chofer || '',
      String(data.placa || '').toUpperCase(),
      Number(data.cantidadPago || 0),
      data.descripcion || ''
    ]);

    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function getOrCreateHourMeterSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName('CONTROL HOROMETRO');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('CONTROL HOROMETRO');
    sheet.appendRow(['ITEM', 'FECHA', 'LUGAR', 'CHOFER', 'HOROMETRO ANTERIOR', 'HOROMETRO ACTUAL', 'HORAS TRABAJADAS', 'PLACA', 'CAMBIO DE ACEITE?', 'DESCRIPCION']);
  }
  return sheet;
}

function getSheetByGid(gid) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) {
      return sheets[i];
    }
  }
  return null;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
