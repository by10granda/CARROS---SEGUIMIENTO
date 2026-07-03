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

    if (data.action === 'configureWhatsApp') {
      configureWhatsApp(data);
      return jsonResponse({ ok: true });
    }

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
        String(data.placa || '').toUpperCase(),
        Number(data.horometroAnterior || 0),
        Number(data.horometroActual || 0),
        data.descripcion || '',
        String(data.cambioAceite || 'NO').toUpperCase()
      ]);
      processHourMeterAlert(String(data.placa || '').toUpperCase());
      return jsonResponse({ ok: true });
    }
    if (sheetName === 'Cambio de aceite') {
      const kmInicial = Number(data.kilometrajeInicial || 0);
      const kmFinal = Number(data.kilometrajeFinal || 0);
      const sumaKm = calculateOilMileageSum(sheet, String(data.placa || '').toUpperCase(), kmInicial, kmFinal);
      sheet.appendRow([
        item,
        data.fecha || '',
        data.lugar || '',
        data.maestro || '',
        data.taller || '',
        data.chofer || '',
        String(data.placa || '').toUpperCase(),
        Number(data.cantidadPago || 0),
        data.descripcion || '',
        kmInicial,
        kmFinal,
        sumaKm
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

function calculateOilMileageSum(sheet, placa, kmInicial, kmFinal) {
  const values = sheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < values.length; i++) {
    const rowPlate = String(values[i][6] || '').toUpperCase();
    if (rowPlate !== placa) continue;
    const storedSum = Number(values[i][11] || 0);
    if (storedSum > total) total = storedSum;
    if (!storedSum) total += Math.max(0, Number(values[i][10] || 0) - Number(values[i][9] || 0));
  }
  return total + Math.max(0, Number(kmFinal || 0) - Number(kmInicial || 0));
}

function configureWhatsApp(data) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('WHATSAPP_PHONE_NUMBER_ID', data.phoneNumberId || '');
  properties.setProperty('WHATSAPP_ACCESS_TOKEN', data.accessToken || '');
  properties.setProperty('WHATSAPP_TEMPLATE_NAME', data.templateName || 'jaspers_market_order_confirmation_v1');
  properties.setProperty('WHATSAPP_LANGUAGE_CODE', data.languageCode || 'en_US');
}

function getOrCreateHourMeterSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName('CONTROL HOROMETRO');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('CONTROL HOROMETRO');
    sheet.appendRow(['ITEM', 'FECHA', 'LUGAR', 'CHOFER', 'PLACA', 'HOROMETRO INICIAL', 'HOROMETRO FINAL', 'DESCRIPCION', 'CAMBIO DE ACEITE']);
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
    const rowPlate = String(row[4] || '').toUpperCase();
    if (rowPlate !== placa) continue;

    lastDate = row[1] || lastDate;
    const changedOil = String(row[8] || '').toUpperCase() === 'SI';
    if (changedOil) {
      accumulatedHours = 0;
    } else {
      accumulatedHours += Math.max(0, Number(row[6] || 0) - Number(row[5] || 0));
    }
  }

  if (accumulatedHours < 250 || wasHourMeterAlertSent(placa, accumulatedHours)) return;

  const results = sendWhatsAppHourMeterAlert(placa, accumulatedHours, lastDate);
  results.forEach(function (result) {
    registerHourMeterAlert(placa, accumulatedHours, lastDate, result.status, result.response);
  });
}

function sendWhatsAppHourMeterAlert(placa, hours, lastDate) {
  const properties = PropertiesService.getScriptProperties();
  const phoneNumberId = properties.getProperty('WHATSAPP_PHONE_NUMBER_ID');
  const token = properties.getProperty('WHATSAPP_ACCESS_TOKEN');
  const templateName = properties.getProperty('WHATSAPP_TEMPLATE_NAME') || 'alerta_horometro';
  const languageCode = properties.getProperty('WHATSAPP_LANGUAGE_CODE') || 'es';
  const recipients = ['593939069555'];

  if (!phoneNumberId || !token) {
    return [{ status: 'PENDIENTE_CONFIGURACION_WHATSAPP', response: 'Faltan WHATSAPP_PHONE_NUMBER_ID o WHATSAPP_ACCESS_TOKEN' }];
  }

  const url = 'https://graph.facebook.com/v25.0/' + phoneNumberId + '/messages';
  const results = [];

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

    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      const body = response.getContentText();
      results.push({
        status: code >= 200 && code < 300 ? 'ENVIADA' : 'ERROR_META_' + code,
        response: body
      });
    } catch (error) {
      results.push({ status: 'ERROR_APPS_SCRIPT', response: error.message });
    }
  });

  return results;
}

function wasHourMeterAlertSent(placa, hours) {
  const sheet = getOrCreateAlertsSheet();
  const values = sheet.getDataRange().getValues();
  const bucket = Math.floor(Number(hours || 0) / 250);

  for (let i = 1; i < values.length; i++) {
    const response = String(values[i][6] || '');
    if (String(values[i][0]).toUpperCase() === placa && Number(values[i][3]) === bucket && String(values[i][4]).toUpperCase() === 'ENVIADA' && response.indexOf('messages') !== -1) {
      return true;
    }
  }
  return false;
}

function registerHourMeterAlert(placa, hours, lastDate, status, response) {
  const sheet = getOrCreateAlertsSheet();
  const bucket = Math.floor(Number(hours || 0) / 250);
  sheet.appendRow([placa, hours, lastDate, bucket, status || 'ENVIADA', new Date(), response || '']);
}

function getOrCreateAlertsSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName('ALERTAS HOROMETRO');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('ALERTAS HOROMETRO');
    sheet.appendRow(['PLACA', 'HORAS', 'ULTIMA FECHA', 'BLOQUE 250H', 'ESTADO', 'FECHA ALERTA', 'RESPUESTA META']);
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
