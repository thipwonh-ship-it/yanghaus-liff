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
// รุ่นที่อยากใช้ตามลำดับ (prefix พอ) — ระบบดึงรายชื่อรุ่นที่มีจริงจาก Google แล้วจับคู่ ถ้ารุ่นแรกล้น (503/429) จะไปรุ่นถัดไป
const GEMINI_PREFER = ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3.5-flash', 'gemini-3', 'gemini-2.5-flash'];

const TAB_EXP = 'ค่าใช้จ่าย';
const TAB_BUY = 'ซื้อ';
const TAB_COD = 'COD';
const TAB_SALE = 'ขาย';
const TAB_CFG = 'ตั้งค่า';
const CAT_SHIP = 'Lalamove/ค่าส่งสินค้า';
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
  return { ok: true, business: col(1), suppliers: col(5), categories: col(7), carriers: col(11), hasAI: !!geminiKey_() };
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
  const prompt = type === 'cod' ?
    'You are reading a Thai parcel/COD shipping receipt or waybill (Flash Express, NNP, J&T, Kerry...) for bookkeeping. Return ONLY JSON with keys:\n' +
    'date (YYYY-MM-DD; Thai year พ.ศ. minus 543; empty if unreadable), ' +
    'carrier (best match from: ' + cfg.carriers.join(' | ') + '), ' +
    'tracking (parcel/tracking/waybill number, e.g. TH2301996BHA0G or OT6924638), ' +
    'customer (consignee/receiver name), ' +
    'items (goods description, short Thai), ' +
    'cod_amount (amount to collect from receiver; on NNP waybill it is the number in *** after the item e.g. ***3,090; 0 if not COD), ' +
    'freight (shipping fee number), cod_fee (COD service fee number, 0 if none), vat (VAT amount shown, 0 if none), ' +
    'total_paid (total the sender paid now; 0 if paid by deduction later), ' +
    'receipt_no (receipt/document number if any), seller_tax_id (carrier 13-digit tax id, digits only, empty if none), ' +
    'confidence (0-1).'
    :
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
  const isBuy = r.type === 'buy';
  const sh = ss.getSheetByName(isBuy ? TAB_BUY : TAB_EXP);
  const d = r.date ? new Date(r.date + 'T00:00:00+07:00') : new Date();
  const total = Number(r.total) || 0;
  const hasVat = !!r.hasVat;
  const invNo = String(r.invoiceNo || '').trim();

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

/* COD: 1 ใบ → แถว COD + แถวขาย (ผูกด้วยเลขพัสดุ) + แถวค่าส่งในค่าใช้จ่าย (ถ้าจ่ายต้นทาง) */
function saveCod_(ss, r) {
  const tracking = String(r.tracking || '').trim();
  if (!tracking) return { ok: false, error: 'ต้องมีเลขพัสดุ' };
  const d = r.date ? new Date(r.date + 'T00:00:00+07:00') : new Date();
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

  // รูป
  let photoUrl = '';
  if (r.image) {
    const ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM'), ymd = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
    const folder = getOrCreate_(getOrCreate_(DriveApp.getFolderById(FOLDER_ID), ym), hasVat ? SUB_VAT : SUB_OTHER);
    const name = ymd + '_COD_' + (r.carrier || '') + '_' + tracking + '.jpg';
    photoUrl = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.image), r.mime || 'image/jpeg', name)).getUrl();
  }
  const by = r.by ? 'จดผ่านไลน์: ' + r.by : '';

  // 1) แถว COD: A วันที่ B ขนส่ง C เลขพัสดุ D ลูกค้า E สินค้า F ยอดCOD G ค่าส่ง H COD fee I VAT | J สูตร | K อ้างอิง L เงินเข้า M วันที่เงินเข้า N ยอดโอน | O สูตร | P รูป Q หมายเหตุ R ผู้บันทึก
  const rowC = nextRow_(shCod);
  shCod.getRange(rowC, 1, 1, 9).setValues([[d, r.carrier || '', tracking, r.customer || '', r.items || '', cod, freight, codFee, vat]]);
  shCod.getRange(rowC, 11, 1, 4).setValues([[r.orderRef || '', false, '', '']]);
  shCod.getRange(rowC, 16, 1, 3).setValues([[photoUrl, [paidNow ? 'จ่ายค่าส่งต้นทาง' : 'ขนส่งหักค่าส่งจากยอด COD', r.note].filter(Boolean).join(' | '), r.by || '']]);
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
    shS.getRange(rowS, 20).setValue(['COD ' + (r.carrier || '') + ' ' + tracking, by].filter(Boolean).join(' | '));
    shS.getRange(rowS, 1).setNumberFormat('yyyy-mm-dd');
    shS.getRange(rowS, 4).setNumberFormat('@');
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
