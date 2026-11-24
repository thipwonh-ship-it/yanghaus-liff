/**
 * YANG Haus — เช็คสต๊อกซัพพลายเออร์ (เจสตาร์: ล้อแม็ก KAYS / ยาง) ผ่าน LINE (หลังบ้าน)
 * ผูกกับ Google Sheet "YANG Haus สต๊อกเจสตาร์" (container-bound: เปิดชีท → ส่วนขยาย → Apps Script)
 *
 * Deploy: Apps Script → Deploy → New deployment → Web app
 *   Execute as: Me / Who has access: Anyone
 * Script Properties (Project Settings → Script properties):
 *   LINE_CHANNEL_ID = Channel ID ของ LINE Login channel "YH ขนส่ง Login" (หน้า Basic settings)
 *   STOCK_FOLDER_ID = (ระบบตั้งเอง) โฟลเดอร์เก็บไฟล์ Excel ต้นฉบับ
 *
 * สิทธิ์: แท็บ "ผู้ใช้" คอลัมน์ C = อัปโหลด / ดู / รออนุมัติ / ปิด
 *   คนที่เปิดหน้าครั้งแรกจะถูกเพิ่มเป็น "รออนุมัติ" อัตโนมัติ เจ้าของแก้เป็น "อัปโหลด" หรือ "ดู" ในชีท
 */
// ชุดข้อมูล: 1 ชุด = 1 แท็บ · เพิ่มชุดใหม่ = เพิ่มบรรทัดที่นี่ + สร้างแท็บ + เพิ่มตัวอ่านไฟล์ใน index.html (KINDS)
const KIND_TAB = { wheel: 'สต๊อกล้อ', tire: 'สต๊อกยาง' };
const TAB_PHOTO = 'รูปสินค้า';
const TAB_USER  = 'ผู้ใช้';
const TAB_LOG   = 'ประวัติอัปโหลด';
const ROLE_UPLOAD = 'อัปโหลด', ROLE_VIEW = 'ดู', ROLE_PENDING = 'รออนุมัติ';
const MAX_PENDING = 30;      // กันคนนอกที่ได้ลิงก์มากดจนแท็บผู้ใช้บวม
const MIN_ITEMS = 50;        // ไฟล์ที่อ่านได้น้อยกว่านี้ถือว่าผิดไฟล์
const STOCK_HEAD = ['ขอบ', 'รหัส', 'ชื่อ'];
const TH_MONTH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function doGet() {
  return json_({ ok: true, service: 'yanghaus-stock', time: new Date().toISOString() });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    switch (req.action) {
      case 'whoami': return json_(whoami_(req.idToken));
      case 'data':   { const u = requireRole_(req.idToken, false); return json_(u.ok ? data_(u, req.kind, req.have) : u); }
      case 'upload': { const u = requireRole_(req.idToken, true);  return json_(u.ok ? upload_(req, u) : u); }
      default:       return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- สิทธิ์ ---------- */
/* ตรวจ ID token กับ LINE · จำผลไว้ใน cache (ไม่เกิน 10 นาที และไม่เกินอายุ token) เพื่อไม่ต้องยิง LINE ซ้ำทุกคำขอ */
function verifyLine_(idToken) {
  const cid = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID');
  if (!idToken || !cid) return null;
  const cache = CacheService.getScriptCache();
  const key = 'v:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cid + '|' + idToken));
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', payload: { id_token: idToken, client_id: cid }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const j = JSON.parse(res.getContentText());
  if (!j.sub) return null;
  const u = { userId: j.sub, name: j.name || '' };
  const left = j.exp ? Math.floor(Number(j.exp) - Date.now() / 1000) - 5 : 300;
  if (left > 10) { try { cache.put(key, JSON.stringify(u), Math.min(600, left)); } catch (e) {} }
  return u;
}

/* อ่านแท็บผู้ใช้ → {userId: {row, name, role}} */
function users_(sh) {
  const last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r, i) {
    const id = String(r[0] || '').trim();
    if (id) out[id] = { row: i + 2, name: String(r[1] || ''), role: String(r[2] || '').trim() };
  });
  return out;
}

/* คืน role ของ userId · ถ้ายังไม่มีในแท็บ → เพิ่มแถว "รออนุมัติ" (ไม่เกิน MAX_PENDING) */
function roleOf_(u) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_USER);
  const all = users_(sh);
  if (all[u.userId]) return all[u.userId].role;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return ROLE_PENDING;
  try {
    const again = users_(sh);
    if (again[u.userId]) return again[u.userId].role;
    const pending = Object.keys(again).filter(function (k) { return again[k].role === ROLE_PENDING; }).length;
    if (pending < MAX_PENDING) sh.appendRow([u.userId, u.name, ROLE_PENDING, new Date(), '']);
  } finally { lock.releaseLock(); }
  return ROLE_PENDING;
}

function whoami_(idToken) {
  const u = verifyLine_(idToken);
  if (!u) return authFail_();
  const role = roleOf_(u);
  return { ok: true, userId: u.userId, name: u.name, canView: role === ROLE_VIEW || role === ROLE_UPLOAD, canUpload: role === ROLE_UPLOAD, role: role };
}

function authFail_() {
  const noCid = !PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID');
  return { ok: false, code: 'auth', error: noCid ? 'หลังบ้านยังไม่ได้ตั้ง LINE_CHANNEL_ID' : 'ยืนยันตัวตนกับ LINE ไม่ผ่าน (หมดอายุ) เปิดหน้านี้ใหม่อีกครั้ง' };
}

function requireRole_(idToken, needUpload) {
  const u = verifyLine_(idToken);
  if (!u) return authFail_();
  const role = roleOf_(u);
  const ok = needUpload ? role === ROLE_UPLOAD : (role === ROLE_UPLOAD || role === ROLE_VIEW);
  if (!ok) return { ok: false, code: 'role', error: needUpload ? 'บัญชีนี้ไม่มีสิทธิ์อัปโหลด' : 'บัญชีนี้ยังไม่ได้รับสิทธิ์ดูสต๊อก', userId: u.userId, name: u.name, role: role };
  return { ok: true, userId: u.userId, name: u.name, role: role };
}

/* ---------- อ่านข้อมูล ---------- */
/* แถวล่าสุดของชุดนั้นในแท็บประวัติอัปโหลด: A เวลา B ชุด C ไฟล์ D ผู้อัปโหลด E userId F รายการ G มีของ H ไฟล์ต้นฉบับ */
function meta_(ss, kind) {
  const sh = ss.getSheetByName(TAB_LOG), last = sh.getLastRow();
  if (last < 2) return null;
  const v = sh.getRange(2, 1, last - 1, 7).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    const r = v[i];
    if (String(r[1]).trim() !== kind) continue;
    const t = r[0] instanceof Date ? r[0] : new Date(r[0]);
    return { file: String(r[2] || ''), at: thaiTime_(t), atIso: Utilities.formatDate(t, 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss"), by: String(r[3] || ''), count: Number(r[5]) || 0, inStock: Number(r[6]) || 0 };
  }
  return null;
}

function thaiTime_(d) {
  const p = Utilities.formatDate(d, 'Asia/Bangkok', 'd|M|yyyy|HH:mm').split('|');
  return p[0] + ' ' + TH_MONTH[+p[1] - 1] + ' ' + String((+p[2] + 543) % 100) + ' ' + p[3];
}

function data_(u, kind, have) {
  kind = kind || 'wheel';
  if (!KIND_TAB[kind]) return { ok: false, error: 'ไม่รู้จักชุดข้อมูล ' + kind };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KIND_TAB[kind]);
  if (!sh) return { ok: false, error: 'ไม่พบแท็บ ' + KIND_TAB[kind] };
  // อ่านใต้ lock เดียวกับตอนอัปโหลด กันอ่านเจอแท็บครึ่ง ๆ กลาง ๆ ระหว่างที่อีกคนกำลังเขียนทับ
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'ระบบกำลังอัปเดตไฟล์อยู่ ลองใหม่อีกครั้ง' };
  let v = null, meta = null;
  try {
    const lastR = sh.getLastRow(), lastC = sh.getLastColumn();
    if (lastR >= 2 && lastC > STOCK_HEAD.length) v = sh.getRange(1, 1, lastR, lastC).getValues();
    meta = meta_(ss, kind);
  } finally { lock.releaseLock(); }
  let branches = [], items = [];
  if (v) {
    let head = v[0].map(function (h) { return String(h || '').trim(); });
    let end = head.length;
    while (end > 0 && (!head[end - 1] || head[end - 1] === 'รวม')) end--;       // ตัดคอลัมน์ "รวม" และคอลัมน์ว่างท้ายตาราง
    branches = head.slice(STOCK_HEAD.length, end);
    for (let i = 1; i < v.length; i++) {
      const r = v[i];
      if (!r[1]) continue;
      const q = [];
      for (let c = STOCK_HEAD.length; c < end; c++) q.push(Number(r[c]) || 0);
      items.push([String(r[0]), String(r[1]), String(r[2]), q]);
    }
  }
  const photos = kind === 'wheel' ? photos_(ss) : { urls: [], map: {} };
  const me = { userId: u.userId, name: u.name, role: u.role, canView: true, canUpload: u.role === ROLE_UPLOAD };
  // ver = ลายนิ้วมือของข้อมูลทั้งชุด · ถ้าตรงกับที่เครื่องผู้ใช้จำไว้ (have) ตอบสั้น ๆ ว่า same ไม่ต้องส่งข้อมูลซ้ำ
  const ver = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([meta, branches, items, photos]), Utilities.Charset.UTF_8));
  if (have && have === ver) return { ok: true, kind: kind, same: true, ver: ver, meta: meta, me: me };
  return { ok: true, kind: kind, ver: ver, meta: meta, branches: branches, items: items, photos: photos, me: me };
}

/* แท็บรูปสินค้า → {urls:[...], map:{รหัส:[[ชื่อรูป, index ของ url, มีใน Shopee 0/1], ...]}} (url ซ้ำกันเยอะ เก็บครั้งเดียว) */
function photos_(ss) {
  const sh = ss.getSheetByName(TAB_PHOTO), last = sh ? sh.getLastRow() : 0;
  const urls = [], idx = {}, map = {};
  if (last < 2) return { urls: urls, map: map };
  sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
    const code = String(r[0] || '').trim(), url = String(r[2] || '').trim();
    if (!code || !/^https:\/\//.test(url)) return;
    if (!(url in idx)) { idx[url] = urls.length; urls.push(url); }
    const sp = r[3] === true || /^(true|y|yes|ใช่|1)$/i.test(String(r[3]).trim()) ? 1 : 0;
    (map[code] = map[code] || []).push([String(r[1] || ''), idx[url], sp]);
  });
  return { urls: urls, map: map };
}

/* ---------- อัปโหลด ---------- */
function upload_(req, u) {
  const kind = req.kind || 'wheel';
  if (!KIND_TAB[kind]) return { ok: false, error: 'ไม่รู้จักชุดข้อมูล ' + kind };
  const file = String(req.file || '').trim().slice(0, 120);
  const branches = (req.branches || []).map(function (b) { return String(b).trim(); }).filter(Boolean);
  const src = req.items || [];
  if (!file) return { ok: false, error: 'ไม่มีชื่อไฟล์' };
  if (branches.length < 1 || branches.length > 40) return { ok: false, error: 'อ่านชื่อสาขาจากไฟล์ไม่ได้' };
  if (src.length < MIN_ITEMS) return { ok: false, error: 'ไฟล์มีแค่ ' + src.length + ' รายการ น้อยผิดปกติ ไม่บันทึก' };
  if (src.length > 20000) return { ok: false, error: 'ไฟล์ใหญ่เกินไป (' + src.length + ' รายการ)' };

  const rows = [];
  let inStock = 0;
  for (let i = 0; i < src.length; i++) {
    const it = src[i], q = it[3] || [];
    const code = String(it[1] || '').trim(), name = String(it[2] || '').trim();
    if (!code || !name || q.length !== branches.length) return { ok: false, error: 'ข้อมูลแถวที่ ' + (i + 1) + ' ไม่ครบ ไม่บันทึก' };
    let total = 0;
    const nums = q.map(function (x) { const n = Math.round(Number(x)) || 0; total += n; return n; });
    if (total > 0) inStock++;
    rows.push([String(it[0] || ''), code, name].concat(nums, [total]));
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'มีคนกำลังอัปโหลดอยู่ รอสักครู่แล้วลองใหม่' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(KIND_TAB[kind]);
    if (!sh) return { ok: false, error: 'ไม่พบแท็บ ' + KIND_TAB[kind] };
    const head = STOCK_HEAD.concat(branches, ['รวม']);
    const needR = rows.length + 1, needC = head.length;
    if (sh.getMaxRows() < needR) sh.insertRowsAfter(sh.getMaxRows(), needR - sh.getMaxRows());
    if (sh.getMaxColumns() < needC) sh.insertColumnsAfter(sh.getMaxColumns(), needC - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearContent();
    sh.getRange(1, 1, sh.getMaxRows(), STOCK_HEAD.length).setNumberFormat('@');     // ขอบ/รหัส/ชื่อ เป็นข้อความ กันรหัสถูกแปลงเป็นตัวเลข
    sh.getRange(1, 1, 1, needC).setValues([head]);
    sh.getRange(2, 1, rows.length, needC).setValues(rows);
    SpreadsheetApp.flush();

    let link = '';
    try { link = saveOriginal_(ss, req.fileB64, kind + '_' + file); } catch (e) { link = 'เก็บไฟล์ต้นฉบับไม่ได้: ' + e; }
    ss.getSheetByName(TAB_LOG).appendRow([new Date(), kind, file, u.name, u.userId, rows.length, inStock, link]);
    SpreadsheetApp.flush();
    return { ok: true, kind: kind, meta: meta_(ss, kind) };
  } finally { lock.releaseLock(); }
}

function saveOriginal_(ss, b64, name) {
  if (!b64) return '';
  const props = PropertiesService.getScriptProperties();
  let folder = null;
  const fid = props.getProperty('STOCK_FOLDER_ID');
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; } }
  if (!folder) {
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const it = parent.getFoldersByName('สต๊อกเจสตาร์ ไฟล์ต้นฉบับ');
    folder = it.hasNext() ? it.next() : parent.createFolder('สต๊อกเจสตาร์ ไฟล์ต้นฉบับ');
    props.setProperty('STOCK_FOLDER_ID', folder.getId());
  }
  const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmm');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', stamp + '_' + name);
  return folder.createFile(blob).getUrl();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ทดสอบใน editor ---------- */
function testConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(JSON.stringify({
    LINE_CHANNEL_ID: PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID') ? 'มี' : 'ไม่มี',
    tabs: [KIND_TAB.wheel, KIND_TAB.tire, TAB_PHOTO, TAB_USER, TAB_LOG].map(function (t) { return t + ':' + (ss.getSheetByName(t) ? 'มี' : 'ไม่มี'); }),
    meta: { wheel: meta_(ss, 'wheel'), tire: meta_(ss, 'tire') },
  }));
}
