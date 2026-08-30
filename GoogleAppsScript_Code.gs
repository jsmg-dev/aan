/**
 * AAN PERFUME - Secure-ish Google Sheets Order Backend
 *
 * Google Sheet columns:
 * Order ID, Date, Status, Product, Size, Quantity, Total,
 * Customer Name, Phone, Email, Delivery Address
 *
 * SET THESE BEFORE DEPLOYING:
 * 1) ADMIN_PASSWORD: strong admin password for admin.html
 * 2) Deploy as Web App -> Execute as Me -> Who has access: Anyone
 */

const SHEET_NAME = 'Orders';
const ADMIN_PASSWORD = 'CHANGE_THIS_TO_A_STRONG_ADMIN_PASSWORD';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Order ID','Date','Status','Product','Size','Quantity','Total','Customer Name','Phone','Email','Delivery Address']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizePhone_(v) { return String(v || '').replace(/\D/g,'').slice(-10); }
function normalizeEmail_(v) { return String(v || '').trim().toLowerCase(); }

function rowToOrder_(r) {
  return {orderId:String(r[0]||''),date:String(r[1]||''),status:String(r[2]||'Order Placed'),productName:String(r[3]||''),size:String(r[4]||''),quantity:Number(r[5]||1),total:Number(r[6]||0),name:String(r[7]||''),phone:String(r[8]||''),email:String(r[9]||''),address:String(r[10]||'')};
}

function doPost(e) {
  const p=e.parameter||{};
  if(p.action!=='createOrder') return ContentService.createTextOutput('Invalid action');
  const sheet=getSheet_();
  const last=sheet.getLastRow();
  if(last>1){
    const ids=sheet.getRange(2,1,last-1,1).getValues().flat().map(String);
    if(ids.includes(String(p.orderId||''))) return ContentService.createTextOutput('OK');
  }
  sheet.appendRow([p.orderId||'',p.date||new Date(),p.status||'Order Placed',p.product||'',p.size||'',Number(p.quantity||1),Number(p.total||0),p.name||'',p.phone||'',normalizeEmail_(p.email),p.address||'']);
  return ContentService.createTextOutput('OK');
}

function jsonp_(data, callback) {
  const json=JSON.stringify(data);
  if(!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback||'')) return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doGet(e) {
  const p=e.parameter||{};
  const action=p.action||'';
  const sheet=getSheet_();
  const values=sheet.getDataRange().getValues();
  const rows=values.slice(1).map(rowToOrder_);
  const callback=p.callback||'';

  if(action==='customerHistory'){
    const phone=normalizePhone_(p.phone);
    const email=normalizeEmail_(p.email);
    if(!phone || !email) return jsonp_([],callback);
    const result=rows.filter(o=>normalizePhone_(o.phone)===phone && normalizeEmail_(o.email)===email).reverse().slice(0,50);
    return jsonp_(result,callback);
  }

  if(action==='adminHistory'){
    if(String(p.password||'')!==ADMIN_PASSWORD) return jsonp_({ok:false,error:'Unauthorized'},callback);
    return jsonp_({ok:true,orders:rows.reverse()},callback);
  }

  return ContentService.createTextOutput('AAN Orders API is running.');
}
