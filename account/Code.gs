/**
 * YANG Haus — จดบัญชีผ่าน LINE (หลังบ้าน)
 * ผูกกับ Google Sheet "YANG Haus บัญชีซื้อ-ขาย 2026"
 *
 * Deploy: Apps Script → Deploy → New deployment → Web app
 *   Execute as: Me / Who has access: Anyone
 * Script Properties (Project Settings → Script properties):
 *   GEMINI_API_KEY = คีย์จาก https://aistudio.google.com/apikey  (ไม่ใส่ = กรอกเองทุกช่อง)
 */
const SHEET_ID  = '17gY8l9pkJjIXqZWRfStYwga-xasJjef97HaHfI7KIB4';
const FOLDER_ID = '1Pfy5Bo6vfkxWDyglsAJwDhg9rh5pajV-';   // โฟลเดอร์ ส่งบัญชี
const GEMINI_MODEL = 'gemini-2.5-flash';

const TAB_EXP = 'ค่าใช้จ่าย';
const TAB_BUY = 'ซื้อ';
const TAB_CFG = 'ตั้งค่า';
const SUB_VAT = 'ภาษีซื้อ';
const SUB_OTHER = 'ใบเสร็จอื่นๆ';

function doGet() {
  return json_({ ok: true, service: 'yanghaus-account', time: new Date().toISOString() });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    switch (req.action) {
      case 'config': return json_(getConfig_());
      case 'read':   return json_(readReceipt_(req.image, req.mime));
      case 'save':   return json_(saveRecord_(req));
      default:       return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- config: รายการ dropdown จากแท็บ ตั้งค่า ---------- */
function getConfig_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_CFG);
  const col = c => sh.getRange(5, c, 40, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
  return { ok: true, business: col(1), suppliers: col(5), categories: col(7), hasAI: !!geminiKey_() };
}

/* ---------- read: ให้ Gemini อ่านใบเสร็จ ---------- */
function geminiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function readReceipt_(b64, mime) {
  const key = geminiKey_();
  if (!key) return { ok: true, noai: true };
  if (!b64) return { ok: false, error: 'no image' };

  const cats = getConfig_().categories;
  const prompt =
    'You are reading a Thai receipt / tax invoice photo for bookkeeping. Return ONLY JSON with these keys:\n' +
    'date (YYYY-MM-DD; Thai year พ.ศ. minus 543; empty if unreadable), ' +
    'vendor (seller name, Thai as printed, no address), ' +
    'tax_id (seller 13-digit tax ID, digits only, empty if none), ' +
    'invoice_no (document number), ' +
    'items (short Thai summary of what was bought, max 60 chars), ' +
    'total (grand total paid, number), ' +
    'vat (VAT amount if shown, number or 0), ' +
    'is_tax_invoice (true only if the document is a ใบกำกับภาษี: has the words ใบกำกับภาษี, seller tax ID and a VAT line; ABB/ใบกำกับภาษีอย่างย่อ also true), ' +
    'is_goods (true if items are wheels/tires/ล้อ/ยาง bought for resale, else false), ' +
    'category (best match from this list: ' + cats.join(' | ') + '), ' +
    'confidence (0-1).';

  const body = {
    contents: [{ parts: [
      { text: prompt },
      { inline_data: { mime_type: mime || 'image/jpeg', data: b64 } }
    ]}],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + key;
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return { ok: true, noai: true, error: 'Gemini ' + res.getResponseCode() };
  const out = JSON.parse(res.getContentText());
  const text = (((out.candidates || [])[0] || {}).content || {}).parts?.[0]?.text || '{}';
  let data = {};
  try { data = JSON.parse(text.replace(/^```json|```$/g, '').trim()); } catch (e) { return { ok: true, noai: true, error: 'parse' }; }
  return { ok: true, data: data };
}

/* ---------- save: เก็บรูปลง Drive + เขียนแถวลงชีต ---------- */
function saveRecord_(r) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const isBuy = r.type === 'buy';
  const sh = ss.getSheetByName(isBuy ? TAB_BUY : TAB_EXP);
  const d = r.date ? new Date(r.date + 'T00:00:00+07:00') : new Date();
  const total = Number(r.total) || 0;
  const hasVat = !!r.hasVat;
  const invNo = String(r.invoiceNo || '').trim();

  // กันบันทึกซ้ำ (เลขที่ใบกำกับเดิมในแท็บเดียวกัน)
  if (invNo && !r.force) {
    const invCol = isBuy ? 10 : 12;   // ซื้อ!J / ค่าใช้จ่าย!L
    const last = sh.getLastRow();
    if (last >= 2) {
      const vals = sh.getRange(2, invCol, last - 1, 1).getValues().map(x => String(x[0]).trim());
      if (vals.indexOf(invNo) >= 0) return { ok: false, dup: true, invoiceNo: invNo };
    }
  }

  // 1) เก็บรูป
  let photoUrl = '';
  if (r.image) {
    const ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
    const ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
    const folder = getOrCreate_(getOrCreate_(DriveApp.getFolderById(FOLDER_ID), ym), hasVat ? SUB_VAT : SUB_OTHER);
    const safe = s => String(s || '').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 40);
    const name = ymd + '_' + (safe(r.vendor) || 'receipt') + '_' + total.toFixed(2) + (invNo ? '_' + safe(invNo) : '') + '.jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(r.image), r.mime || 'image/jpeg', name);
    const file = folder.createFile(blob);
    photoUrl = file.getUrl();
  }

  // 2) เขียนแถว (ข้ามคอลัมน์สูตร)
  const row = nextRow_(sh);
  const note = [r.note, r.by ? 'จดผ่านไลน์: ' + r.by : ''].filter(Boolean).join(' | ');
  if (isBuy) {
    // A วันที่ B ธุรกิจ C ซัพพลายเออร์ D สินค้า E จำนวน F ยอดรวม G มีใบกำกับ | H,I สูตร | J เลขใบกำกับ K ขนส่งเข้า L อ้างอิงออเดอร์ M หมายเหตุ N รูป O เลขผู้เสียภาษี
    sh.getRange(row, 1, 1, 7).setValues([[d, r.business || 'YANG Haus', r.vendor || '', r.items || '', Number(r.qty) || '', total, hasVat]]);
    sh.getRange(row, 10, 1, 6).setValues([[invNo, Number(r.shipIn) || '', r.orderRef || '', note, photoUrl, r.taxId || '']]);
  } else {
    // A วันที่ B ธุรกิจ C ประเภท D รายละเอียด E ยอด F มีใบกำกับ | G,H สูตร | I ผู้ขาย J อ้างอิงออเดอร์ K หมายเหตุ L เลขใบกำกับ M เลขผู้เสียภาษี N รูป
    sh.getRange(row, 1, 1, 6).setValues([[d, r.business || 'YANG Haus', r.category || 'อื่นๆ', r.items || '', total, hasVat]]);
    sh.getRange(row, 9, 1, 6).setValues([[r.vendor || '', r.orderRef || '', note, invNo, r.taxId || '', photoUrl]]);
  }
  sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
  if (invNo) sh.getRange(row, isBuy ? 10 : 12).setNumberFormat('@');
  if (r.taxId) sh.getRange(row, isBuy ? 15 : 13).setNumberFormat('@');

  return { ok: true, row: row, tab: sh.getName(), photoUrl: photoUrl, vat: hasVat ? +(total - total / 1.07).toFixed(2) : 0 };
}

function nextRow_(sh) {
  const a = sh.getRange('A:A').getValues();
  let last = 0;
  for (let i = 0; i < a.length; i++) if (a[i][0] !== '' && a[i][0] !== null) last = i + 1;
  return last + 1;
}

function getOrCreate_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ทดสอบใน editor: เลือก testConfig แล้วกด Run (ครั้งแรกจะขอสิทธิ์) */
function testConfig() { Logger.log(JSON.stringify(getConfig_())); }
