/**
 * YANG Haus — จดบัญชีผ่าน LINE (หลังบ้าน)
 * ผูกกับ Google Sheet "YANG Haus บัญชีซื้อ-ขาย 2026"
 *
 * Deploy: Apps Script → Deploy → New deployment → Web app
 *   Execute as: Me / Who has access: Anyone
 * Script Properties (Project Settings → Script properties):
 *   GEMINI_API_KEY   = คีย์จาก https://aistudio.google.com/apikey  (ไม่ใส่ = กรอกเองทุกช่อง)
 *   LINE_CHANNEL_ID  = Channel ID ของ LINE Login channel "YH ขนส่ง Login" (หน้า Basic settings)
 *   COD_ADMINS       = LINE userId ที่ติ๊กเงิน COD เข้าได้ คั่นด้วย , (ดู userId ได้ท้ายหน้าฟอร์ม)
 */
const SHEET_ID  = '17gY8l9pkJjIXqZWRfStYwga-xasJjef97HaHfI7KIB4';
const FOLDER_ID = '1Pfy5Bo6vfkxWDyglsAJwDhg9rh5pajV-';   // โฟลเดอร์ ส่งบัญชี
// รุ่นที่อยากใช้ตามลำดับ (prefix พอ) — ระบบดึงรายชื่อรุ่นที่มีจริงจาก Google แล้วจับคู่ ถ้ารุ่นแรกล้น (503/429) จะไปรุ่นถัดไป
const GEMINI_PREFER = ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3.5-flash', 'gemini-3', 'gemini-2.5-flash'];

const TAB_EXP = 'ค่าใช้จ่าย';
const TAB_BUY = 'ซื้อ';
const TAB_COD = 'COD';
const TAB_SALE = 'ขาย';
const TAB_CFG = 'ตั้งค่า';
const CAT_SHIP = 'Lalamove/ค่าส่งสินค้า';
// เบอร์ของเราเอง (บจก.ตามโกล์บ / YANG Haus) — ห้ามถือเป็นเบอร์ลูกค้า
const OUR_PHONES = ['0811363332', '0811197855', '0811239924'];
const OUR_NAMES = ['yanghaus', 'yang haus', 'ยางเฮ้าส์', 'ตามโกล์บ', 'tamglobe', 'ทิพย์วรรณ', 'thipwon'];
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
      case 'read':   return json_(readReceipt_(req.image, req.mime, req.type));
      case 'save':   return json_(saveRecord_(req));
      case 'whoami':     return json_(whoami_(req.idToken));
      case 'buyPending': return json_(buyPending_(req.vendor));
      case 'buyMatch':   return json_(buyMatch_(req));
      case 'codPending': { const u = requireAdmin_(req.idToken); return json_(u.ok ? codPending_() : u); }
      case 'codPaid':    { const u = requireAdmin_(req.idToken); return json_(u.ok ? codPaid_(req.items, u.name) : u); }
      case 'codList':    { const u = requireAdmin_(req.idToken); return json_(u.ok ? codList_(req.month) : u); }
      case 'dash':       { const u = requireAdmin_(req.idToken); return json_(u.ok ? dash_(req.month) : u); }
      default:       return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}


/* วันที่จากฟอร์ม (yyyy-MM-dd) → Date: ปี พ.ศ. (≥2400) ลบ 543 · ปีเพี้ยน (<2020 หรือ >ปีนี้+1) → ใช้วันนี้ · กันเคส 2483-11-02 */
function parseDate_(s) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s || '').trim());
  if (!m) return new Date();
  let y = Number(m[1]); const mo = Number(m[2]), da = Number(m[3]);
  if (y >= 2400) y -= 543;
  const thisY = new Date().getFullYear();
  if (y < 2020 || y > thisY + 1 || mo < 1 || mo > 12 || da < 1 || da > 31) return new Date();
  return new Date(y + '-' + String(mo).padStart(2, '0') + '-' + String(da).padStart(2, '0') + 'T00:00:00+07:00');
}

/* ---------- config: รายการ dropdown จากแท็บ ตั้งค่า ---------- */
function getConfig_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_CFG);
  const col = c => sh.getRange(5, c, 40, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
  return { ok: true, business: col(1), suppliers: col(5), categories: col(7), carriers: col(11), sources: col(13), hasAI: !!geminiKey_() };
}

/* ---------- read: ให้ Gemini อ่านใบเสร็จ ---------- */
function geminiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function readReceipt_(b64, mime, type) {
  const key = geminiKey_();
  if (!key) return { ok: true, noai: true };
  if (!b64) return { ok: false, error: 'no image' };

  const cfg = getConfig_();
  const cats = cfg.categories;
  const prompt = type === 'sale' ?
    'You are reading, for bookkeeping, evidence of a sale by our shop YANG Haus (ยางเฮ้าส์ / lorcenter): a bank transfer slip received from a customer, a chat screenshot (Facebook/LINE) with an order, or a handwritten/typed bill. Return ONLY JSON with keys:\n' +
    'date (YYYY-MM-DD; Thai year พ.ศ. minus 543; empty if unreadable), ' +
    'IMPORTANT: our own shop is YANGHAUS / ยางเฮ้าส์ / บริษัท ตามโกล์บ จำกัด (phones 0811363332, 0811197855, 0811239924) — never return us as the customer. The customer is the BUYER/payer/recipient of goods. ' +
    'customer (buyer/payer name), phone (buyer phone digits only, empty if none or if it is one of our phones), province (buyer province in Thai, empty if none), ' +
    'items (what was sold, short Thai, e.g. ล้อแม็ก KAYS 17 นิ้ว x4), qty (number of pieces if clear, else 0), ' +
    'amount (sale amount paid/agreed, number), ' +
    'ref (transfer reference / slip number / order no if any), ' +
    'channel (best guess from: Facebook | LINE | ขายส่ง | อื่นๆ), confidence (0-1).'
    : type === 'cod' ?
    'You are reading a Thai parcel/COD shipping receipt or waybill (Flash Express, NNP, J&T, Kerry...) for bookkeeping. Return ONLY JSON with keys:\n' +
    'date (YYYY-MM-DD; Thai year พ.ศ. minus 543; empty if unreadable), ' +
    'carrier (best match from: ' + cfg.carriers.join(' | ') + '), ' +
    'tracking (parcel/tracking/waybill number, e.g. TH2301996BHA0G or OT6924638), ' +
    'IMPORTANT: the SENDER/shipper is always our own shop YANGHAUS / ยางเฮ้าส์ / บริษัท ตามโกล์บ จำกัด (phones 0811363332, 0811197855, 0811239924, address 756/199 ถนนมอเตอร์เวย์ ประเวศ กรุงเทพฯ). Never return the sender as customer and never return our phones as phone. The customer is the RECEIVER/consignee (ผู้รับ / To). ' +
    'customer (receiver/consignee name), phone (receiver phone, digits only, empty if none or if it is one of our phones), province (receiver province in Thai, e.g. พัทลุง, กรุงเทพมหานคร; empty if none), ' +
    'items (goods description, short Thai), ' +
    'cod_amount (amount to collect from receiver; on NNP waybill it is the number in *** after the item e.g. ***3,090; 0 if not COD), ' +
    'freight (shipping fee number), cod_fee (COD service fee number, 0 if none), vat (VAT amount shown, 0 if none), ' +
    'total_paid (total the sender paid now; 0 if paid by deduction later), ' +
    'receipt_no (receipt/document number if any), seller_tax_id (carrier 13-digit tax id, digits only, empty if none), ' +
    'confidence (0-1).'
    :
    'You are reading, for bookkeeping, one of: a Thai receipt / tax invoice photo, a screenshot of an online order page (Shopee, TikTok Shop, Lazada, Facebook, LINE, a website), or a bank transfer slip (สลิปโอนเงิน from a Thai banking app: shows payer, recipient name/account, amount, date-time, transaction ref). Return ONLY JSON with these keys:\n' +
    'doc_type (one of: tax_invoice_original (ต้นฉบับใบกำกับภาษี) | tax_invoice_copy (สำเนาใบกำกับภาษี) | delivery_note (ใบส่งสินค้า/ใบแจ้งหนี้/ใบส่งของ) | receipt | order_screenshot | transfer_slip), ' +
    'source (where it was bought, best match from: ' + cfg.sources.join(' | ') + ' — detect from app UI/logo/colors: orange=Shopee, black=TikTok, blue/Lazada; a paper receipt from a shop = หน้าร้าน), ' +
    'is_screenshot (true if this is a phone/app screenshot of an order rather than a paper receipt), ' +
    'date (YYYY-MM-DD; Thai year พ.ศ. minus 543; empty if unreadable), ' +
    'vendor (seller name as printed, no address; for a transfer slip use the recipient/payee name), ' +
    'tax_id (seller 13-digit tax ID, digits only, empty if none), ' +
    'invoice_no (document number; online order → order number; transfer slip → transaction reference/เลขที่รายการ), ' +
    'items (short Thai summary of what was bought, max 60 chars; transfer slip → use the memo/บันทึกช่วยจำ if any, else empty), ' +
    'total (grand total paid, number), ' +
    'vat (VAT amount if shown, number or 0), ' +
    'is_tax_invoice (true if the document is or references a ใบกำกับภาษี with seller tax ID and VAT line — original or copy both true; a delivery note that shows a VAT line and tax-invoice number is also true; ABB/ใบกำกับภาษีอย่างย่อ true; an order screenshot is false), ' +
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
  const res = geminiFetch_(key, body);
  const code = res.getResponseCode(), raw = res.getContentText();
  Logger.log('Gemini ' + code + ' ' + raw.slice(0, 500));
  if (code !== 200) {
    let m = ''; try { m = JSON.parse(raw).error.message; } catch (e) {}
    return { ok: true, noai: true, error: 'Gemini ' + code + (m ? ': ' + m.slice(0, 120) : '') };
  }
  const out = JSON.parse(raw);
  const cand = (out.candidates || [])[0] || {};
  const text = ((cand.content || {}).parts || []).map(p => p.text || '').join('') || '';
  if (!text) return { ok: true, noai: true, error: 'ไม่มีคำตอบ (' + (cand.finishReason || out.promptFeedback?.blockReason || '?') + ')' };
  let data = {};
  try { data = JSON.parse(text.replace(/^```json|```$/g, '').trim()); } catch (e) { return { ok: true, noai: true, error: 'parse: ' + text.slice(0, 80) }; }
  return { ok: true, data: data };
}

/* ---------- save: เก็บรูปลง Drive + เขียนแถวลงชีต ---------- */
function saveRecord_(r) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  if (r.type === 'cod') return saveCod_(ss, r);
  if (r.type === 'sale') { const u = requireAdmin_(r.idToken); if (!u.ok) return u; return saveSale_(ss, r); }
  const isBuy = r.type === 'buy';
  const sh = ss.getSheetByName(isBuy ? TAB_BUY : TAB_EXP);
  const d = parseDate_(r.date);
  const total = Number(r.total) || 0;
  const hasVat = !!r.hasVat;
  const invNo = String(r.invoiceNo || '').trim();

  // ซื้อ: ใบกำกับตัวจริงมาทีหลัง → ถ้าเลขที่เดิมมีอยู่แล้ว ไม่ลงซ้ำ แค่ติ๊ก Q "ใบกำกับตัวจริงรับแล้ว" + เก็บรูปตัวจริง
  if (isBuy && r.docType === 'original' && invNo) {
    const last = sh.getLastRow();
    if (last >= 2) {
      const inv = sh.getRange(2, 10, last - 1, 1).getValues().map(x => String(x[0]).trim().toUpperCase());
      const i = inv.indexOf(invNo.toUpperCase());
      if (i >= 0) {
        const row = i + 2;
        let url = '';
        if (r.image) {
          const ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM'), ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
          const folder = getOrCreate_(getOrCreate_(DriveApp.getFolderById(FOLDER_ID), ym), SUB_VAT);
          url = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.image), r.mime || 'image/jpeg', ymd + '_ต้นฉบับ_' + invNo.replace(/[\\/:*?"<>|]/g, '') + '.jpg')).getUrl();
        }
        sh.getRange(row, 17).setValue(true);
        const note = sh.getRange(row, 13).getValue();
        sh.getRange(row, 13).setValue([note, 'ต้นฉบับใบกำกับรับแล้ว ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd/M') + (url ? ' ' + url : '')].filter(Boolean).join(' | '));
        return { ok: true, markedOriginal: true, row: row, tab: sh.getName(), photoUrl: url };
      }
    }
  }

  // กันบันทึกซ้ำ (ปฏิเสธ ไม่ถาม): 1) เลขที่ใบกำกับซ้ำในแท็บไหนก็ได้  2) วันที่+ผู้ขาย+ยอด เท่ากัน
  const dup = findDuplicate_(ss, invNo, d, r.vendor, total);
  if (dup) return { ok: false, dup: true, where: dup };

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
    const gotOriginal = r.docType === 'original' || !hasVat;
    sh.getRange(row, 10, 1, 8).setValues([[invNo, Number(r.shipIn) || '', r.orderRef || '', [note, (hasVat && !gotOriginal) ? 'รอใบกำกับตัวจริง' : ''].filter(Boolean).join(' | '), photoUrl, r.taxId || '', r.source || '', gotOriginal]]);
  } else {
    // A วันที่ B ธุรกิจ C ประเภท D รายละเอียด E ยอด F มีใบกำกับ | G,H สูตร | I ผู้ขาย J อ้างอิงออเดอร์ K หมายเหตุ L เลขใบกำกับ M เลขผู้เสียภาษี N รูป
    sh.getRange(row, 1, 1, 6).setValues([[d, r.business || 'YANG Haus', r.category || 'อื่นๆ', r.items || '', total, hasVat]]);
    sh.getRange(row, 9, 1, 7).setValues([[r.vendor || '', r.orderRef || '', note, invNo, r.taxId || '', photoUrl, r.source || '']]);
  }
  sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
  if (invNo) sh.getRange(row, isBuy ? 10 : 12).setNumberFormat('@');
  if (r.taxId) sh.getRange(row, isBuy ? 15 : 13).setNumberFormat('@');

  return { ok: true, row: row, tab: sh.getName(), photoUrl: photoUrl, vat: hasVat ? +(total - total / 1.07).toFixed(2) : 0 };
}

/* ---------- ซื้อ: ใบส่งของที่ยังรอใบกำกับตัวจริง (Q = FALSE, G = TRUE) ---------- */
function buyPending_(vendor) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_BUY);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, items: [] };
  const v = sh.getRange(2, 1, last - 1, 17).getValues();
  const norm = x => String(x || '').replace(/\s+/g, '').toLowerCase();
  const items = [];
  v.forEach((row, i) => {
    if (!row[0] || row[6] !== true || row[16] === true) return;
    const sameVendor = !vendor || norm(row[2]).indexOf(norm(vendor).slice(0, 6)) >= 0 || norm(vendor).indexOf(norm(row[2]).slice(0, 6)) >= 0;
    items.push({ row: i + 2, date: row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(row[0]),
      vendor: row[2], items: row[3], total: Number(row[5]) || 0, docNo: String(row[9] || ''), sameVendor: sameVendor });
  });
  items.sort((a, b) => (b.sameVendor - a.sameVendor) || (a.date < b.date ? -1 : 1));
  return { ok: true, items: items };
}

/* จับคู่ใบกำกับตัวจริง 1 ใบ กับใบส่งของหลายแถว: ใส่เลขใบกำกับ + ติ๊ก Q + เก็บรูป */
function buyMatch_(r) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_BUY);
  const invNo = String(r.invoiceNo || '').trim();
  if (!invNo) return { ok: false, error: 'ต้องมีเลขที่ใบกำกับ' };
  const rows = (r.rows || []).map(Number).filter(Boolean);
  if (!rows.length) return { ok: false, error: 'ยังไม่ได้เลือกใบส่งของ' };
  let url = '';
  if (r.image) {
    const d = parseDate_(r.date);
    const ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM'), ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
    const folder = getOrCreate_(getOrCreate_(DriveApp.getFolderById(FOLDER_ID), ym), SUB_VAT);
    url = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.image), r.mime || 'image/jpeg', ymd + '_ต้นฉบับ_' + invNo.replace(/[\\/:*?"<>|]/g, '') + '.jpg')).getUrl();
  }
  let sum = 0;
  rows.forEach(row => {
    const cur = sh.getRange(row, 1, 1, 17).getValues()[0];
    if (!cur[0] || cur[16] === true) return;
    sum += Number(cur[5]) || 0;
    const oldNo = String(cur[9] || '');
    sh.getRange(row, 10).setValue(invNo).setNumberFormat('@');
    if (r.taxId) sh.getRange(row, 15).setValue(String(r.taxId)).setNumberFormat('@');
    if (r.invoiceDate) { sh.getRange(row, 1).setValue(new Date(r.invoiceDate + 'T00:00:00+07:00')).setNumberFormat('yyyy-mm-dd'); }
    sh.getRange(row, 17).setValue(true);
    sh.getRange(row, 13).setValue([String(cur[12] || ''), (oldNo && oldNo !== invNo ? 'ใบส่งของ ' + oldNo : ''), 'ต้นฉบับใบกำกับรับแล้ว ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd/M') + (url ? ' ' + url : '')].filter(Boolean).join(' | '));
  });
  return { ok: true, count: rows.length, sum: sum, invoiceTotal: Number(r.total) || 0, diff: +(sum - (Number(r.total) || 0)).toFixed(2), photoUrl: url };
}

/* ---------- สิทธิ์: ตรวจ ID token กับ LINE แล้วเช็ค userId ใน COD_ADMINS ---------- */
function verifyLine_(idToken) {
  const cid = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID');
  if (!idToken || !cid) return null;
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', payload: { id_token: idToken, client_id: cid }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const j = JSON.parse(res.getContentText());
  return { userId: j.sub, name: j.name || '' };
}
function isAdmin_(userId) {
  const list = (PropertiesService.getScriptProperties().getProperty('COD_ADMINS') || '').split(/[,\s]+/).filter(Boolean);
  return list.indexOf(userId) >= 0;
}
function whoami_(idToken) {
  const u = verifyLine_(idToken);
  if (!u) return { ok: true, userId: '', admin: false, note: !PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID') ? 'ยังไม่ได้ตั้ง LINE_CHANNEL_ID' : 'ตรวจ token ไม่ผ่าน' };
  return { ok: true, userId: u.userId, name: u.name, admin: isAdmin_(u.userId) };
}
function requireAdmin_(idToken) {
  const u = verifyLine_(idToken);
  if (!u) return { ok: false, error: 'ยืนยันตัวตนกับ LINE ไม่ผ่าน (ตั้ง LINE_CHANNEL_ID หรือเปิดหน้าใหม่)' };
  if (!isAdmin_(u.userId)) return { ok: false, error: 'ไม่มีสิทธิ์ติ๊กเงิน COD (userId ' + u.userId + ')' };
  return { ok: true, userId: u.userId, name: u.name };
}

/* รายการ COD ที่ยังไม่ได้เงิน (L = FALSE) */
function codPending_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_COD);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, items: [] };
  const v = sh.getRange(2, 1, last - 1, 12).getValues();
  const items = [];
  v.forEach((row, i) => {
    if (!row[2] || row[11] === true) return;
    items.push({ row: i + 2, date: row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(row[0]),
      carrier: row[1], tracking: String(row[2]), customer: row[3], items: row[4], cod: Number(row[5]) || 0 });
  });
  items.sort((a, b) => a.date < b.date ? -1 : 1);
  return { ok: true, items: items, total: items.reduce((t, x) => t + x.cod, 0) };
}

/* รายการ COD ทั้งหมด (เลือกเดือน yyyy-MM หรือ '' = ทั้งหมด) + รายชื่อเดือนที่มี */
function codList_(month) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_COD);
  const last = sh.getLastRow();
  const fmt = (d, p) => d instanceof Date ? Utilities.formatDate(d, 'Asia/Bangkok', p) : '';
  const items = [], months = {};
  if (last >= 2) {
    const v = sh.getRange(2, 1, last - 1, 14).getValues();
    v.forEach((row, i) => {
      if (!row[2]) return;
      const ym = fmt(row[0], 'yyyy-MM'); months[ym] = true;
      if (month && ym !== month) return;
      items.push({ row: i + 2, date: fmt(row[0], 'd/M'), carrier: row[1], tracking: String(row[2]), customer: row[3], cod: Number(row[5]) || 0,
        paid: row[11] === true, paidDate: fmt(row[12], 'd/M'), paidAmt: Number(row[13]) || 0, sortKey: fmt(row[0], 'yyyy-MM-dd') });
    });
  }
  items.sort((a, b) => a.sortKey < b.sortKey ? 1 : -1);
  const total = items.reduce((t, x) => t + x.cod, 0), paid = items.filter(x => x.paid);
  return { ok: true, items: items, months: Object.keys(months).sort().reverse(), total: total, paidCount: paid.length, paidSum: paid.reduce((t, x) => t + x.paidAmt, 0), pendingSum: items.filter(x => !x.paid).reduce((t, x) => t + x.cod, 0) };
}

/* ---------- Dashboard รายจ่ายรายเดือน (เฉพาะ COD_ADMINS) — อ่านอย่างเดียว ไม่แก้ชีท ---------- */
function dash_(month) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const fmt = (d, p) => d instanceof Date ? Utilities.formatDate(d, 'Asia/Bangkok', p) : '';
  const num = x => Number(x) || 0;
  const months = {};
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
  month = month || now;

  // ซื้อ: A วันที่ C ซัพพลายเออร์ D สินค้า E จำนวน F ยอดรวม G มีใบกำกับ I VAT J เลขใบกำกับ K ขนส่งเข้า M หมายเหตุ N รูป Q ตัวจริงรับแล้ว
  const buy = { rows: [], total: 0, vat: 0, qty: 0, shipIn: 0, byVendor: {}, pendingOriginal: 0 };
  const isYH = b => /yang\s*haus|ยางเฮ้าส์/i.test(String(b || ''));   // dashboard นี้ของ YANG Haus เท่านั้น (ชีทมีธุรกิจอื่นปนอยู่)
  const sb = ss.getSheetByName(TAB_BUY);
  if (sb.getLastRow() >= 2) {
    sb.getRange(2, 1, sb.getLastRow() - 1, 17).getValues().forEach((r, i) => {
      if (!r[0] || !num(r[5]) || !isYH(r[1])) return;
      const ym = fmt(r[0], 'yyyy-MM'); months[ym] = true;
      if (ym !== month) return;
      const waiting = r[6] === true && r[16] !== true;
      const v = String(r[2] || '(ไม่ระบุ)');
      buy.rows.push({ row: i + 2, date: fmt(r[0], 'd/M'), vendor: v, items: r[3], qty: num(r[4]), total: num(r[5]), vat: num(r[8]), hasVat: r[6] === true,
        invNo: String(r[9] || ''), shipIn: num(r[10]), waiting: waiting, photo: String(r[13] || '') });
      buy.total += num(r[5]); buy.vat += num(r[8]); buy.qty += num(r[4]); buy.shipIn += num(r[10]);
      buy.byVendor[v] = (buy.byVendor[v] || 0) + num(r[5]);
      if (waiting) buy.pendingOriginal++;
    });
  }

  // ค่าใช้จ่าย: A วันที่ C ประเภท D รายละเอียด E ยอด F มีใบกำกับ H VAT I ผู้ขาย N รูป
  const exp = { rows: [], total: 0, vat: 0, byCat: {} };
  const se = ss.getSheetByName(TAB_EXP);
  if (se.getLastRow() >= 2) {
    se.getRange(2, 1, se.getLastRow() - 1, 14).getValues().forEach((r, i) => {
      if (!r[0] || !num(r[4]) || !isYH(r[1])) return;
      const ym = fmt(r[0], 'yyyy-MM'); months[ym] = true;
      if (ym !== month) return;
      if (/^COD\b/.test(String(r[10] || ''))) return;   // แถวค่าส่ง COD ที่ระบบลงเอง → ไปรวมในส่วน "ค่าส่ง+ค่าธรรมเนียม COD" (อ่านจากแท็บ COD) กันนับซ้ำ
      const c = String(r[2] || 'อื่นๆ');
      exp.rows.push({ row: i + 2, date: fmt(r[0], 'd/M'), cat: c, detail: r[3], total: num(r[4]), vat: num(r[7]), hasVat: r[5] === true, vendor: r[8], photo: String(r[13] || '') });
      exp.total += num(r[4]); exp.vat += num(r[7]);
      exp.byCat[c] = (exp.byCat[c] || 0) + num(r[4]);
    });
  }

  // รายรับที่จดผ่านไลน์ (แท็บ ขาย): A วันที่ C ช่องทาง D เลขออเดอร์ E ลูกค้า F สินค้า G จำนวน H ยอด M เงินเข้าจริง S รับเงินแล้ว T หมายเหตุ
  // แยก 3 กลุ่ม: cod (หมายเหตุขึ้นต้น COD) / ws (ช่องทาง ขายส่ง) / other (Facebook/LINE โอน) — ไม่นับแพลตฟอร์ม (Shopee/TikTok/Lazada ลงวันที่ 7 เป็นระบบแยก)
  const PLATFORM = /shopee|tiktok|lazada/i;
  const inc = { rows: [], total: 0, paid: 0, qty: 0, groups: { cod: { total: 0, paid: 0, count: 0 }, ws: { total: 0, paid: 0, count: 0 }, other: { total: 0, paid: 0, count: 0 } }, byCustomer: {} };
  const si = ss.getSheetByName(TAB_SALE);
  if (si && si.getLastRow() >= 2) {
    si.getRange(2, 1, si.getLastRow() - 1, 20).getValues().forEach((r, i) => {
      if (!r[0] || !num(r[7]) || !isYH(r[1])) return;
      const ch = String(r[2] || '');
      if (PLATFORM.test(ch)) return;
      const ym = fmt(r[0], 'yyyy-MM'); months[ym] = true;
      if (ym !== month) return;
      const note = String(r[19] || '');
      const g = /^COD\b/i.test(note) ? 'cod' : /ขายส่ง/.test(ch) ? 'ws' : 'other';
      const paid = r[18] === true, amt = num(r[7]), got = num(r[12]);
      const cust = String(r[4] || '(ไม่ระบุ)');
      inc.rows.push({ row: i + 2, date: fmt(r[0], 'd/M'), group: g, channel: ch, ref: String(r[3] || ''), customer: cust, items: r[5], qty: num(r[6]), amount: amt, paid: paid, got: got });
      inc.total += amt; inc.qty += num(r[6]); if (paid) inc.paid += (got || amt);
      inc.groups[g].total += amt; inc.groups[g].count++; if (paid) inc.groups[g].paid += (got || amt);
      inc.byCustomer[cust] = (inc.byCustomer[cust] || 0) + amt;
    });
  }

  // COD: ใช้ codList_ (เดือนที่ส่ง) + ค้างรับทั้งหมดทุกเดือน
  const cod = codList_(month);
  cod.months.forEach(m => months[m] = true);
  const pendAll = codPending_();

  // ค่าส่ง + ค่าธรรมเนียม COD (รายจ่าย) จากแท็บ COD: A วันที่ B ขนส่ง C เลขพัสดุ D ลูกค้า G ค่าส่ง H COD fee I VAT Q หมายเหตุ (จ่ายต้นทาง / ขนส่งหัก)
  const codCost = { rows: [], freight: 0, fee: 0, vat: 0, total: 0, paidNow: 0, deducted: 0, byCarrier: {} };
  const sc = ss.getSheetByName(TAB_COD);
  if (sc.getLastRow() >= 2) {
    sc.getRange(2, 1, sc.getLastRow() - 1, 17).getValues().forEach((r, i) => {
      if (!r[0] || !r[2]) return;
      if (fmt(r[0], 'yyyy-MM') !== month) return;
      const f = num(r[6]), fee = num(r[7]), v = num(r[8]), t = f + fee + v;
      if (!t) return;
      const upfront = /จ่ายค่าส่งต้นทาง/.test(String(r[16] || ''));
      const car = String(r[1] || '(ไม่ระบุ)');
      codCost.rows.push({ row: i + 2, date: fmt(r[0], 'd/M'), carrier: car, tracking: String(r[2]), customer: r[3], freight: f, fee: fee, vat: v, total: t, upfront: upfront });
      codCost.freight += f; codCost.fee += fee; codCost.vat += v; codCost.total += t;
      if (upfront) codCost.paidNow += t; else codCost.deducted += t;
      codCost.byCarrier[car] = (codCost.byCarrier[car] || 0) + t;
    });
  }

  const sortRows = a => a.sort((x, y) => x.row < y.row ? 1 : -1);
  sortRows(buy.rows); sortRows(exp.rows); sortRows(inc.rows); sortRows(codCost.rows);
  const top = o => Object.keys(o).map(k => ({ name: k, total: o[k] })).sort((a, b) => b.total - a.total);
  return { ok: true, month: month, months: Object.keys(months).sort().reverse(),
    inc: { rows: inc.rows, total: inc.total, paid: inc.paid, qty: inc.qty, count: inc.rows.length, groups: inc.groups, byCustomer: top(inc.byCustomer).slice(0, 8) },
    buy: { rows: buy.rows, total: buy.total, vat: buy.vat, qty: buy.qty, shipIn: buy.shipIn, count: buy.rows.length, byVendor: top(buy.byVendor), pendingOriginal: buy.pendingOriginal },
    exp: { rows: exp.rows, total: exp.total, vat: exp.vat, count: exp.rows.length, byCat: top(exp.byCat) },
    codCost: { rows: codCost.rows, freight: codCost.freight, fee: codCost.fee, vat: codCost.vat, total: codCost.total, paidNow: codCost.paidNow, deducted: codCost.deducted, count: codCost.rows.length, byCarrier: top(codCost.byCarrier) },
    cod: { items: cod.items, total: cod.total, paidCount: cod.paidCount, paidSum: cod.paidSum, pendingSum: cod.pendingSum, count: cod.items.length,
      pendingAllSum: pendAll.total || 0, pendingAllCount: (pendAll.items || []).length } };
}

/* ติ๊กเงินเข้า: items = [{row, tracking, amount, date}] */
function codPaid_(items, by) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_COD);
  let n = 0, sum = 0; const skipped = [];
  (items || []).forEach(it => {
    const row = Number(it.row);
    const cur = sh.getRange(row, 3).getValue();
    if (String(cur).trim() !== String(it.tracking).trim()) { skipped.push(it.tracking); return; }   // แถวเลื่อน → ไม่แตะ
    const d = it.date ? new Date(it.date + 'T00:00:00+07:00') : new Date();
    sh.getRange(row, 12, 1, 3).setValues([[true, d, Number(it.amount) || 0]]);
    sh.getRange(row, 13).setNumberFormat('yyyy-mm-dd');
    if (by) { const q = sh.getRange(row, 17); q.setValue([q.getValue(), 'เงินเข้า: ' + by].filter(Boolean).join(' | ')); }
    n++; sum += Number(it.amount) || 0;
  });
  return { ok: true, count: n, sum: sum, skipped: skipped };
}

/* COD: 1 ใบ → แถว COD + แถวขาย (ผูกด้วยเลขพัสดุ) + แถวค่าส่งในค่าใช้จ่าย (ถ้าจ่ายต้นทาง) */
function scrubCustomer_(r) {
  const ph = String(r.phone || '').replace(/\D/g, '');
  if (OUR_PHONES.indexOf(ph) >= 0) r.phone = '';
  const nm = String(r.customer || '').toLowerCase();
  if (nm && OUR_NAMES.some(x => nm.indexOf(x) >= 0)) r.customer = '';
  return r;
}

function saveCod_(ss, r) {
  scrubCustomer_(r);
  const tracking = String(r.tracking || '').trim();
  if (!tracking) return { ok: false, error: 'ต้องมีเลขพัสดุ' };
  const d = parseDate_(r.date);
  const cod = Number(r.codAmount) || 0, freight = Number(r.freight) || 0, codFee = Number(r.codFee) || 0, vat = Number(r.vat) || 0;
  const paidNow = !!r.paidNow;
  const hasVat = !!r.hasVat;
  const shCod = ss.getSheetByName(TAB_COD);

  // กันซ้ำด้วยเลขพัสดุ
  const lastC = shCod.getLastRow();
  if (lastC >= 2) {
    const have = shCod.getRange(2, 3, lastC - 1, 1).getValues().map(x => String(x[0]).trim().toUpperCase());
    const i = have.indexOf(tracking.toUpperCase());
    if (i >= 0) return { ok: false, dup: true, where: TAB_COD + ' แถว ' + (i + 2) + ' (เลขพัสดุ ' + tracking + ')' };
  }

  // รูป (1 = ใบปะหน้า, 2 = ใบเสร็จขนส่ง)
  let photoUrl = '';
  if (r.image || r.image2) {
    const ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM'), ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
    const folder = getOrCreate_(getOrCreate_(DriveApp.getFolderById(FOLDER_ID), ym), hasVat ? SUB_VAT : SUB_OTHER);
    const base = ymd + '_COD_' + (r.carrier || '') + '_' + tracking;
    const urls = [];
    if (r.image) urls.push(folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.image), r.mime || 'image/jpeg', base + '_ใบปะหน้า.jpg')).getUrl());
    if (r.image2) urls.push(folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.image2), r.mime || 'image/jpeg', base + '_ใบเสร็จขนส่ง.jpg')).getUrl());
    photoUrl = urls.join(' | ');
  }
  const by = r.by ? 'จดผ่านไลน์: ' + r.by : '';

  // 1) แถว COD: A วันที่ B ขนส่ง C เลขพัสดุ D ลูกค้า E สินค้า F ยอดCOD G ค่าส่ง H COD fee I VAT | J สูตร | K อ้างอิง L เงินเข้า M วันที่เงินเข้า N ยอดโอน | O สูตร | P รูป Q หมายเหตุ R ผู้บันทึก
  const rowC = nextRow_(shCod);
  shCod.getRange(rowC, 1, 1, 9).setValues([[d, r.carrier || '', tracking, r.customer || '', r.items || '', cod, freight, codFee, vat]]);
  shCod.getRange(rowC, 11, 1, 4).setValues([[r.orderRef || '', false, '', '']]);
  shCod.getRange(rowC, 16, 1, 5).setValues([[photoUrl, [paidNow ? 'จ่ายค่าส่งต้นทาง' : 'ขนส่งหักค่าส่งจากยอด COD', r.note].filter(Boolean).join(' | '), r.by || '', String(r.phone || '').replace(/\D/g, ''), r.province || '']]);
  if (r.phone) shCod.getRange(rowC, 19).setNumberFormat('@');
  shCod.getRange(rowC, 1).setNumberFormat('yyyy-mm-dd');
  shCod.getRange(rowC, 3).setNumberFormat('@');

  // 2) แถวขาย (เฉพาะที่มียอด COD): D = เลขพัสดุ, M/S ดึงจากแท็บ COD อัตโนมัติ
  let rowS = 0;
  if (cod > 0) {
    const shS = ss.getSheetByName(TAB_SALE);
    rowS = nextRow_(shS);
    const m = "IFERROR(MATCH($D" + rowS + ",COD!$C:$C,0),0)";
    shS.getRange(rowS, 1, 1, 9).setValues([[d, r.business || 'YANG Haus', r.channel || 'Facebook', tracking, r.customer || '', r.items || '', Number(r.qty) || '', cod, 'ไม่ออก']]);
    shS.getRange(rowS, 13).setFormula('=IF(' + m + '>0,IF(INDEX(COD!$L:$L,' + m + ')=TRUE,INDEX(COD!$N:$N,' + m + '),""),"")');
    shS.getRange(rowS, 15).setValue(0);
    shS.getRange(rowS, 19).setFormula('=IF(' + m + '>0,INDEX(COD!$L:$L,' + m + ')=TRUE,FALSE)');
    shS.getRange(rowS, 20, 1, 3).setValues([[['COD ' + (r.carrier || '') + ' ' + tracking, by].filter(Boolean).join(' | '), String(r.phone || '').replace(/\D/g, ''), r.province || '']]);
    shS.getRange(rowS, 1).setNumberFormat('yyyy-mm-dd');
    shS.getRange(rowS, 4).setNumberFormat('@'); shS.getRange(rowS, 21).setNumberFormat('@');
  }

  // 3) แถวค่าส่งในค่าใช้จ่าย (เฉพาะจ่ายเงินสดต้นทาง) ผูกเลขพัสดุ → ไปโผล่ในกำไรต่อออเดอร์
  let rowE = 0;
  const paid = paidNow ? (Number(r.totalPaid) || (freight + codFee + vat)) : 0;
  if (paid > 0) {
    const shE = ss.getSheetByName(TAB_EXP);
    rowE = nextRow_(shE);
    shE.getRange(rowE, 1, 1, 6).setValues([[d, r.business || 'YANG Haus', CAT_SHIP, 'ค่าส่ง COD ' + (r.carrier || '') + ' ' + tracking + (r.customer ? ' → ' + r.customer : ''), paid, hasVat]]);
    shE.getRange(rowE, 9, 1, 6).setValues([[r.carrier || '', tracking, ['COD', by].filter(Boolean).join(' | '), r.receiptNo || '', r.taxId || '', photoUrl]]);
    shE.getRange(rowE, 1).setNumberFormat('yyyy-mm-dd');
  }
  return { ok: true, tab: TAB_COD, row: rowC, saleRow: rowS, expRow: rowE, paid: paid, photoUrl: photoUrl };
}

/* ขาย (โอน/ขายส่ง/Facebook ที่ไม่ใช่ COD): 1 ใบ → แถวขาย (+ รูปสลิป/แชทใน Drive/ภาษีขาย) */
function saveSale_(ss, r) {
  scrubCustomer_(r);
  const d = parseDate_(r.date);
  const amount = Number(r.amount) || 0;
  if (!amount) return { ok: false, error: 'ต้องมียอดขาย' };
  const shS = ss.getSheetByName(TAB_SALE);
  const ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
  // เลขอ้างอิงขาย: ใช้ ref ที่ให้มา ไม่งั้นสร้าง FB-YYYYMMDD-NNN
  let ref = String(r.ref || '').trim();
  const lastS = shS.getLastRow();
  const refs = lastS >= 2 ? shS.getRange(2, 4, lastS - 1, 1).getValues().map(x => String(x[0]).trim()) : [];
  if (!ref) {
    const base = (r.channel === 'ขายส่ง' ? 'WS-' : r.channel === 'LINE' ? 'LN-' : 'FB-') + ymd.replace(/-/g, '');
    let n = 1; while (refs.indexOf(base + '-' + String(n).padStart(3, '0')) >= 0) n++;
    ref = base + '-' + String(n).padStart(3, '0');
  } else if (refs.indexOf(ref) >= 0) {
    return { ok: false, dup: true, where: TAB_SALE + ' แถว ' + (refs.indexOf(ref) + 2) + ' (เลขอ้างอิง ' + ref + ')' };
  }
  // กันซ้ำ: วันที่ + ลูกค้า + ยอด
  if (lastS >= 2) {
    const v = shS.getRange(2, 1, lastS - 1, 8).getValues();
    for (let i = 0; i < v.length; i++) {
      const rd = v[i][0] instanceof Date ? Utilities.formatDate(v[i][0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(v[i][0]);
      if (rd === ymd && String(v[i][4]).trim() === String(r.customer || '').trim() && Math.abs(Number(v[i][7]) - amount) < 0.005)
        return { ok: false, dup: true, where: TAB_SALE + ' แถว ' + (i + 2) + ' (วันที่/ลูกค้า/ยอด เท่ากัน)' };
    }
  }
  let photoUrl = '';
  if (r.image) {
    const ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
    const folder = getOrCreate_(getOrCreate_(DriveApp.getFolderById(FOLDER_ID), ym), 'ภาษีขาย');
    const safe = x => String(x || '').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 40);
    photoUrl = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.image), r.mime || 'image/jpeg', ymd + '_ขาย_' + (safe(r.customer) || 'ลูกค้า') + '_' + amount.toFixed(2) + '_' + safe(ref) + '.jpg')).getUrl();
  }
  const by = r.by ? 'จดผ่านไลน์: ' + r.by : '';
  const row = nextRow_(shS);
  const paid = !!r.paid;
  // A วันที่ B ธุรกิจ C ช่องทาง D เลขออเดอร์ E ลูกค้า F สินค้า G จำนวน H ยอด I ออกใบกำกับ | J,K สูตร | L เลขใบกำกับ M เงินเข้าจริง | N สูตร | O ขนส่งออก | P,Q,R สูตร | S รับเงินแล้ว T หมายเหตุ U เบอร์ V จังหวัด
  shS.getRange(row, 1, 1, 9).setValues([[d, r.business || 'YANG Haus', r.channel || 'Facebook', ref, r.customer || '', r.items || '', Number(r.qty) || '', amount, r.invoice || 'ไม่ออก']]);
  shS.getRange(row, 12, 1, 2).setValues([[r.invoiceNo || '', paid ? amount : '']]);
  shS.getRange(row, 15).setValue(Number(r.shipOut) || 0);
  shS.getRange(row, 19, 1, 4).setValues([[paid, [r.note, photoUrl ? 'รูป: ' + photoUrl : '', by].filter(Boolean).join(' | '), String(r.phone || '').replace(/\D/g, ''), r.province || '']]);
  shS.getRange(row, 1).setNumberFormat('yyyy-mm-dd'); shS.getRange(row, 4).setNumberFormat('@'); shS.getRange(row, 21).setNumberFormat('@');
  return { ok: true, tab: TAB_SALE, row: row, ref: ref, photoUrl: photoUrl };
}

function findDuplicate_(ss, invNo, d, vendor, total) {
  const ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
  const norm = v => String(v || '').replace(/\s+/g, '').toLowerCase();
  const tabs = [
    { name: TAB_EXP, inv: 12, vendor: 9, total: 5 },   // ค่าใช้จ่าย: L, I, E
    { name: TAB_BUY, inv: 10, vendor: 3, total: 6 }    // ซื้อ: J, C, F
  ];
  for (const t of tabs) {
    const sh = ss.getSheetByName(t.name);
    const last = sh.getLastRow();
    if (last < 2) continue;
    const vals = sh.getRange(2, 1, last - 1, 12).getValues();
    for (let i = 0; i < vals.length; i++) {
      const row = vals[i];
      if (row[0] === '' || row[0] === null) continue;
      const rowInv = String(row[t.inv - 1] || '').trim();
      if (invNo && rowInv && norm(rowInv) === norm(invNo))
        return t.name + ' แถว ' + (i + 2) + ' (เลขที่ ' + rowInv + ')';
      const rowDate = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(row[0]);
      if (rowDate === ymd && Math.abs(Number(row[t.total - 1]) - total) < 0.005 && norm(row[t.vendor - 1]) === norm(vendor))
        return t.name + ' แถว ' + (i + 2) + ' (วันที่/ผู้ขาย/ยอด เท่ากัน)';
    }
  }
  return '';
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

/* รายชื่อรุ่นที่คีย์นี้ใช้ได้จริง (ถาม Google วันละครั้ง เก็บใน cache) เรียงตาม GEMINI_PREFER */
function geminiModels_(key) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('gemini_models');
  if (hit) return JSON.parse(hit);
  let names = [];
  try {
    const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=' + key, { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      names = (JSON.parse(r.getContentText()).models || [])
        .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
        .map(m => m.name.replace('models/', ''))
        .filter(n => !/tts|audio|image|embedding|live|preview|exp|thinking/i.test(n));
    }
  } catch (e) {}
  const ordered = [];
  GEMINI_PREFER.forEach(p => names.forEach(n => { if (n.indexOf(p) === 0 && ordered.indexOf(n) < 0) ordered.push(n); }));
  const list = ordered.length ? ordered.slice(0, 4) : GEMINI_PREFER;
  cache.put('gemini_models', JSON.stringify(list), 21600);
  return list;
}

/* เรียก Gemini: รุ่นที่เคยใช้ได้ก่อน → 404 ข้าม / 503,429 ลองซ้ำ 1 ครั้งแล้วไปรุ่นถัดไป */
function geminiFetch_(key, body) {
  const props = PropertiesService.getScriptProperties();
  const models = geminiModels_(key);
  const remembered = props.getProperty('GEMINI_MODEL_OK');
  const order = remembered && models.indexOf(remembered) >= 0 ? [remembered].concat(models.filter(m => m !== remembered)) : models;
  const opt = { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true };
  let res = null, best = null;
  for (const m of order) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent?key=' + key;
    for (let attempt = 0; attempt < 2; attempt++) {
      res = UrlFetchApp.fetch(url, opt);
      const c = res.getResponseCode();
      if (c === 404) break;
      if (c === 503 || c === 429) { best = res; Utilities.sleep(1000); continue; }
      if (c === 200 && m !== remembered) props.setProperty('GEMINI_MODEL_OK', m);
      return res;
    }
  }
  return best || res;
}

/* ทดสอบใน editor: เลือก testConfig แล้วกด Run (ครั้งแรกจะขอสิทธิ์) */
function testConfig() { Logger.log(JSON.stringify(getConfig_())); }

/* ทดสอบคีย์ + รุ่นโมเดล Gemini โดยไม่ต้องใช้รูป: เลือก testGemini แล้วกด Run ดู Execution log */
function testGemini() {
  const key = geminiKey_();
  if (!key) { Logger.log('ไม่พบ GEMINI_API_KEY ใน Script properties'); return; }
  Logger.log('key ขึ้นต้น ' + key.slice(0, 6) + '… ยาว ' + key.length);
  const res = geminiFetch_(key, { contents: [{ parts: [{ text: 'ตอบคำว่า OK' }] }] });
  Logger.log('รุ่นที่ใช้ได้: ' + JSON.stringify(geminiModels_(key)));
  Logger.log('รุ่นที่ใช้ล่าสุด: ' + PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL_OK'));
  Logger.log('HTTP ' + res.getResponseCode() + ' → ' + res.getContentText().slice(0, 400));
}
