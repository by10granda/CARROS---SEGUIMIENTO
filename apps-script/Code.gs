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
      processHourMeterAlert(String(data.placa || '').toUpperCase());
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

function processHourMeterAlert(placa) {
  if (!placa) return;

  const sheet = getOrCreateHourMeterSheet();
  const values = sheet.getDataRange().getValues();
  let accumulatedHours = 0;
  let lastDate = '';

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowPlate = String(row[7] || '').toUpperCase();
    if (rowPlate !== placa) continue;

    lastDate = row[1] || lastDate;
    const changedOil = String(row[8] || '').toUpperCase() === 'SI';
    if (changedOil) {
      accumulatedHours = 0;
    } else {
      accumulatedHours += Number(row[6] || 0);
    }
  }

  if (accumulatedHours < 250 || wasHourMeterAlertSent(placa, accumulatedHours)) return;

  sendWhatsAppHourMeterAlert(placa, accumulatedHours, lastDate);
  registerHourMeterAlert(placa, accumulatedHours, lastDate);
}

function sendWhatsAppHourMeterAlert(placa, hours, lastDate) {
  const properties = PropertiesService.getScriptProperties();
  const phoneNumberId = properties.getProperty('WHATSAPP_PHONE_NUMBER_ID');
  const token = properties.getProperty('WHATSAPP_ACCESS_TOKEN');
  const templateName = properties.getProperty('WHATSAPP_TEMPLATE_NAME') || 'alerta_horometro';
  const languageCode = properties.getProperty('WHATSAPP_LANGUAGE_CODE') || 'es';
  const recipients = ['593939069555', '593997882191'];

  if (!phoneNumberId || !token) {
    registerHourMeterAlert(placa, hours, lastDate, 'PENDIENTE_CONFIGURACION_WHATSAPP');
    return;
  }

  const url = 'https://graph.facebook.com/v20.0/' + phoneNumberId + '/messages';

  recipients.forEach(function (recipient) {
    const payload = {
      messaging_product: 'whatsapp',
      to: recipient,
      type: 'template',
      template: {
        name: templateName,
        language: { code: languageCode },
        components: [{
          type: 'body',
          parameters: [
            { type: 'text', text: placa },
            { type: 'text', text: String(hours) },
            { type: 'text', text: String(lastDate || '') }
          ]
        }]
      }
    };

    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  });
}

function wasHourMeterAlertSent(placa, hours) {
  const sheet = getOrCreateAlertsSheet();
  const values = sheet.getDataRange().getValues();
  const bucket = Math.floor(Number(hours || 0) / 250);

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase() === placa && Number(values[i][3]) === bucket) {
      return true;
    }
  }
  return false;
}

function registerHourMeterAlert(placa, hours, lastDate, status) {
  const sheet = getOrCreateAlertsSheet();
  const bucket = Math.floor(Number(hours || 0) / 250);
  sheet.appendRow([placa, hours, lastDate, bucket, status || 'ENVIADA', new Date()]);
}

function getOrCreateAlertsSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName('ALERTAS HOROMETRO');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('ALERTAS HOROMETRO');
    sheet.appendRow(['PLACA', 'HORAS', 'ULTIMA FECHA', 'BLOQUE 250H', 'ESTADO', 'FECHA ALERTA']);
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
