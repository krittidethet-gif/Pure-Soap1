// Google Apps Script (ผูกกับ Google Sheet) — Deploy เป็น Web app
const ADMIN_KEY = 'soap1234'; // เปลี่ยนเป็นรหัสของคุณ ใช้เปิดหน้า orders.html

function setup() { // รันครั้งเดียวเพื่อสร้างชีตและข้อมูลตัวอย่าง
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defs = {
    Products: [['sku','name','price','stock','image'],
      ['SOAP-LAV','สบู่ลาเวนเดอร์',89,100,'images/lavender.svg'],
      ['SOAP-LEM','สบู่เลมอน',79,100,'images/lemon.svg']],
    Orders: [['order_id','timestamp','sku','product','qty','unit_price','line_total','status']],
    Shipping: [['order_id','timestamp','name','phone','address','subdistrict','district','province','postcode','note']]
  };
  Object.keys(defs).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    sh.clear(); sh.getRange(1,1,defs[n].length,defs[n][0].length).setValues(defs[n]);
    sh.setFrozenRows(1);
  });
}

const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const rows = name => {
  const v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getDataRange().getValues();
  const h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k,i) => [k, r[i]])));
};

function doGet(e) {
  const a = (e.parameter.action || 'products');
  if (a === 'products') return json({ products: rows('Products') });
  if (a === 'orders') {
    if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
    return json({ orders: rows('Orders'), shipping: rows('Shipping') });
  }
  return json({ error: 'bad action' });
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prods = rows('Products');
    const id = 'SO' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmmss');
    const ts = new Date();
    const lines = [];
    (d.items || []).forEach(it => {
      const p = prods.find(x => x.sku === it.sku), q = parseInt(it.qty, 10);
      if (p && q > 0) lines.push([id, ts, p.sku, p.name, q, p.price, q * p.price, 'รอแพ็ค']);
    });
    if (!lines.length) return json({ ok: false, error: 'no items' });
    const o = ss.getSheetByName('Orders');
    o.getRange(o.getLastRow() + 1, 1, lines.length, 8).setValues(lines);
    ss.getSheetByName('Shipping').appendRow([id, ts, d.name, "'" + d.phone, d.address,
      d.subdistrict, d.district, d.province, "'" + d.postcode, d.note || '']);
    return json({ ok: true, order_id: id });
  } finally { lock.releaseLock(); }
}
