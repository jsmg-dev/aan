/**
 * AAN PERFUME - Google Sheets Order Backend
 *
 * 1. Create a Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Paste this entire file into Code.gs.
 * 4. Deploy -> New deployment -> Web app.
 * 5. Execute as: Me
 * 6. Who has access: Anyone
 * 7. Copy the /exec URL and put it into AAN_ORDERS_ENDPOINT in index.html.
 *
 * The first request creates an "Orders" sheet with headers.
 */

const SHEET_NAME = 'Orders';

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

function doPost(e) {
  const p = e.parameter || {};
  if (p.action !== 'createOrder') {
    return ContentService.createTextOutput('Invalid action');
  }

  const sheet = getSheet_();

  // Prevent duplicate order IDs.
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
    if (ids.includes(p.orderId)) {
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
  if (p.action !== 'history') {
    return ContentService.createTextOutput('AAN Orders API is running.');
  }

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  const rows = values.slice(1).reverse().slice(0,50).map(r => ({
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

  const json = JSON.stringify(rows);
  const callback = p.callback || '';

  // JSONP is used for GitHub Pages -> Google Apps Script reads without CORS setup.
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
