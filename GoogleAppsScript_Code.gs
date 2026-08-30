const SHEET_NAME = 'Orders';
const ADMIN_PASSWORD = 'AAN@2026';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Order ID','Date','Status','Product','Size','Quantity','Total',
      'Customer Name','Phone','Email','Delivery Address'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonp_(callback, data) {
  const json = JSON.stringify(data);
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback || '')) {
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function rowsFromSheet_() {
  const values = getSheet_().getDataRange().getValues();
  return values.slice(1).map(r => ({
    orderId: String(r[0] || ''),
    date: String(r[1] || ''),
    status: String(r[2] || 'Order Placed'),
    productName: String(r[3] || ''),
    size: String(r[4] || ''),
    quantity: Number(r[5] || 1),
    total: Number(r[6] || 0),
    name: String(r[7] || ''),
    phone: String(r[8] || ''),
    email: String(r[9] || ''),
    address: String(r[10] || '')
  }));
}

function doPost(e) {
  const p = e.parameter || {};
  if (p.action !== 'createOrder') {
    return ContentService.createTextOutput('Invalid action');
  }

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
    if (ids.includes(String(p.orderId || ''))) {
      return ContentService.createTextOutput('OK');
    }
  }

  sheet.appendRow([
    p.orderId || '',
    p.date || new Date(),
    p.status || 'Order Placed',
    p.product || '',
    p.size || '',
    Number(p.quantity || 1),
    Number(p.total || 0),
    p.name || '',
    p.phone || '',
    p.email || '',
    p.address || ''
  ]);

  return ContentService.createTextOutput('OK');
}

function doGet(e) {
  const p = e.parameter || {};
  const callback = p.callback || '';
  const action = p.action || '';

  if (action === 'customerHistory') {
    const phone = String(p.phone || '').replace(/\D/g, '');
    const email = String(p.email || '').trim().toLowerCase();

    const rows = rowsFromSheet_()
      .filter(o => {
        const rowPhone = String(o.phone || '').replace(/\D/g, '');
        const rowEmail = String(o.email || '').trim().toLowerCase();
        return phone && email && rowPhone === phone && rowEmail === email;
      })
      .reverse()
      .slice(0, 50);

    return jsonp_(callback, rows);
  }

  if (action === 'adminHistory') {
    if (String(p.password || '') !== ADMIN_PASSWORD) {
      return jsonp_(callback, { error: 'Unauthorized' });
    }

    const rows = rowsFromSheet_().reverse().slice(0, 500);
    return jsonp_(callback, rows);
  }

  return ContentService.createTextOutput('AAN Orders API is running.');
}
