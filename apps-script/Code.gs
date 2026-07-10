const SHEET_ID = '1Evr1lpNSwLYXWgcRf5D5NSPUD5nR-YS_EczVK4PVAHI';
const SHEET_GIDS = {
  'Mantenimiento': 0,
  'Lavado': 819388144,
  'Engrasada': 2024356449,
  'Cambio de aceite': 464443967
};

function doGet() {
  return jsonResponse({ ok: true, message: 'Punto PAS API activa' });
}

function authorizeExternalRequest() {
  const response = UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis');
  return response.getResponseCode();
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'configureWhatsApp') {
      configureWhatsApp(data);
      return jsonResponse({ ok: true });
    }

    const sheetName = data.tipoServicio;
    const sheet = getSheetByGid(SHEET_GIDS[sheetName]);

    if (!sheet) {
      return jsonResponse({ ok: false, error: 'No existe la hoja: ' + sheetName });
    }

    const item = Math.max(sheet.getLastRow(), 1);
    if (sheetName === 'Cambio de aceite') {
      const placa = String(data.placa || '').toUpperCase();
      const kmInicial = Number(data.kilometrajeInicial || 0);
      const kmFinal = Number(data.kilometrajeFinal || 0);
      const horometroInicial = Number(data.horometroAnterior || 0);
      const horometroFinal = Number(data.horometroActual || 0);
      const horasTrabajadas = Math.max(0, horometroFinal - horometroInicial);
      const sumaKm = calculateOilMileageSum(sheet, placa, kmInicial, kmFinal, String(data.cambioAceite || 'NO').toUpperCase() === 'SI');
      sheet.appendRow([
        item,
        data.fecha || '',
        data.lugar || '',
        data.maestro || '',
        data.taller || '',
        data.chofer || '',
        placa,
        Number(data.cantidadPago || 0),
        data.descripcion || '',
        kmInicial,
        kmFinal,
        sumaKm,
        normalizeTransport(data.transporte),
        horometroInicial,
        horometroFinal,
        horasTrabajadas,
        String(data.cambioAceite || 'NO').toUpperCase() === 'SI' ? 'SI' : 'NO'
      ]);
      processOilAlert(placa);
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

function normalizeTransport(value) {
  return String(value || '').toUpperCase().indexOf('PEQUE') !== -1 || String(value || '').toUpperCase() === 'PEQUENO' ? 'PEQUENO' : 'GRANDE';
}

function calculateOilMileageSum(sheet, placa, kmInicial, kmFinal, changedOil) {
  const values = sheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < values.length; i++) {
    const rowPlate = String(values[i][6] || '').toUpperCase();
    if (rowPlate !== placa) continue;
    if (String(values[i][16] || '').toUpperCase() === 'SI') {
      total = 0;
      continue;
    }
    const storedSum = Number(values[i][11] || 0);
    if (storedSum > total) total = storedSum;
    if (!storedSum) total += Math.max(0, Number(values[i][10] || 0) - Number(values[i][9] || 0));
  }
  return changedOil ? 0 : total + Math.max(0, Number(kmFinal || 0) - Number(kmInicial || 0));
}

function configureWhatsApp(data) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('WHATSAPP_PHONE_NUMBER_ID', data.phoneNumberId || '');
  properties.setProperty('WHATSAPP_ACCESS_TOKEN', data.accessToken || '');
  properties.setProperty('WHATSAPP_TEMPLATE_NAME', data.templateName || 'alerta_horometro');
  properties.setProperty('WHATSAPP_LANGUAGE_CODE', data.languageCode || 'es_EC');
}

function processOilAlert(placa) {
  if (!placa) return;

  const sheet = getSheetByGid(SHEET_GIDS['Cambio de aceite']);
  const values = sheet.getDataRange().getValues();
  let accumulatedHours = 0;
  let accumulatedKm = 0;
  let transport = 'GRANDE';
  let lastDate = '';

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowPlate = String(row[6] || '').toUpperCase();
    if (rowPlate !== placa) continue;

    lastDate = row[1] || lastDate;
    transport = normalizeTransport(row[12] || transport);
    const changedOil = String(row[16] || '').toUpperCase() === 'SI';
    if (changedOil) {
      accumulatedHours = 0;
      accumulatedKm = 0;
    } else {
      accumulatedKm += Math.max(0, Number(row[10] || 0) - Number(row[9] || 0));
      accumulatedHours += Number(row[15] || 0) || Math.max(0, Number(row[14] || 0) - Number(row[13] || 0));
    }
  }

  const kmLimit = transport === 'PEQUENO' ? 5000 : 10000;
  const alerts = [];
  if (accumulatedKm >= kmLimit && !wasOilAlertSent(placa, 'KM', accumulatedKm, kmLimit)) alerts.push({ type: 'KM', value: accumulatedKm, limit: kmLimit, label: accumulatedKm + ' km' });
  if (accumulatedHours >= 250 && !wasOilAlertSent(placa, 'HORAS', accumulatedHours, 250)) alerts.push({ type: 'HORAS', value: accumulatedHours, limit: 250, label: accumulatedHours + ' horas' });

  alerts.forEach(function (alert) {
    const results = sendWhatsAppOilAlert(placa, alert.label, lastDate);
    results.forEach(function (result) {
      registerOilAlert(placa, alert.type, alert.value, alert.limit, lastDate, result.status, result.response);
    });
  });
}

function sendWhatsAppOilAlert(placa, value, lastDate) {
  const properties = PropertiesService.getScriptProperties();
  const phoneNumberId = properties.getProperty('WHATSAPP_PHONE_NUMBER_ID');
  const token = properties.getProperty('WHATSAPP_ACCESS_TOKEN');
  const templateName = properties.getProperty('WHATSAPP_TEMPLATE_NAME') || 'alerta_horometro';
  const languageCode = properties.getProperty('WHATSAPP_LANGUAGE_CODE') || 'es_EC';
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
            { type: 'text', text: String(value) },
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

function wasOilAlertSent(placa, type, value, limit) {
  const sheet = getOrCreateAlertsSheet();
  const values = sheet.getDataRange().getValues();
  const bucket = Math.floor(Number(value || 0) / Number(limit || 1));

  for (let i = 1; i < values.length; i++) {
    const response = String(values[i][7] || '');
    if (String(values[i][0]).toUpperCase() === placa && String(values[i][1]).toUpperCase() === type && Number(values[i][4]) === bucket && String(values[i][5]).toUpperCase() === 'ENVIADA' && response.indexOf('messages') !== -1) {
      return true;
    }
  }
  return false;
}

function registerOilAlert(placa, type, value, limit, lastDate, status, response) {
  const sheet = getOrCreateAlertsSheet();
  const bucket = Math.floor(Number(value || 0) / Number(limit || 1));
  sheet.appendRow([placa, type, value, limit, bucket, status || 'ENVIADA', new Date(), response || '', lastDate]);
}

function getOrCreateAlertsSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName('ALERTAS ACEITE');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('ALERTAS ACEITE');
    sheet.appendRow(['PLACA', 'TIPO', 'VALOR', 'LIMITE', 'BLOQUE', 'ESTADO', 'FECHA ALERTA', 'RESPUESTA META', 'ULTIMA FECHA SERVICIO']);
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
