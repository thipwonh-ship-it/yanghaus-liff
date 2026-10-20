// ===== ตั้งค่า =====
const SHEET_ID  = '1t9phG6yVcseuX042WWoExS-nNbSnnOaUb3jJSDllzLc'; // YANG Haus ส่งของ
const FOLDER_ID = '1skXBDHIJXVrswcniLbYAYQlxaN56HJy_';             // รูปส่งของ
// ===================

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const now = new Date();
    const tz = 'Asia/Bangkok';
    const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    const stamp = timeStr.replace(/:/g, '');

    // โฟลเดอร์รายวัน
    const root = DriveApp.getFolderById(FOLDER_ID);
    const it = root.getFoldersByName(dateStr);
    let day;
    if (it.hasNext()) day = it.next();
    else { day = root.createFolder(dateStr); day.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } // ตั้งสิทธิ์ครั้งเดียวต่อวัน ไฟล์ข้างในสืบทอด

    const tracks = String(d.tracking || '').split(/[\n,]+/).map(s => s.trim()).filter(Boolean);

    // กันเลขพัสดุซ้ำ: ซ้ำในฟอร์มเดียวกัน หรือเคยบันทึกในชีตแล้ว
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const dupInForm = tracks.filter((t, i) => tracks.indexOf(t) !== i);
    if (dupInForm.length) throw new Error('เลขพัสดุซ้ำในฟอร์ม: ' + [...new Set(dupInForm)].join(', '));
    if (sheet.getLastRow() > 1) {
      const existing = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim());
      const dup = tracks.filter(t => existing.includes(t));
      if (dup.length) {
        const where = dup.map(t => { const r = existing.indexOf(t) + 2; return `${t} (แถว ${r}, ${sheet.getRange(r,1).getDisplayValue()} ${sheet.getRange(r,2).getDisplayValue()})`; });
        throw new Error('เลขพัสดุนี้เคยบันทึกแล้ว: ' + where.join(' / '));
      }
    }

    // บันทึกรูป แยกกลุ่ม สินค้า / ใบรับฝาก
    const itemLinks = [], receiptLinks = [];
    (d.photos || []).forEach(p => {
      const isItem = p.group !== 'receipt';
      const idx = isItem ? itemLinks.length : receiptLinks.length;
      const label = isItem ? (tracks[idx] || `item${idx+1}`) : `receipt${idx+1}`;
      const blob = Utilities.newBlob(Utilities.base64Decode(p.data), 'image/jpeg',
        `${dateStr}_${stamp}_${d.carrier}_${label}.jpg`);
      const f = day.createFile(blob); // สิทธิ์แชร์สืบทอดจากโฟลเดอร์ รูปส่งของ
      (isItem ? itemLinks : receiptLinks).push(f.getUrl());
    });

    const mapLink = (d.lat && d.lng) ? `https://maps.google.com/?q=${d.lat},${d.lng}` : '';
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['วันที่','เวลา','ผู้ส่ง','ขนส่ง','สาขา','เลขพัสดุ','จำนวนชิ้นรวม','GPS','รูปสินค้า','รูปใบรับฝาก/เคาน์เตอร์','ยิงรับแล้ว','พนักงานรับฝาก','ตรวจสอบแล้ว','หมายเหตุ','LINE userId']);
    }

    const scanned = d.scanned === true;
    // 1 แถว / 1 เลขพัสดุ — รูปสินค้าชิ้นที่ n คู่กับเลขที่ n (ถ้าจำนวนไม่ตรง ใส่ทุกรูปทุกแถว)
    const paired = itemLinks.length === tracks.length;
    const rows = (tracks.length ? tracks : ['']).map((t, i) =>
      [dateStr, timeStr, d.name, d.carrier, d.branch, t, tracks.length, mapLink,
       paired ? itemLinks[i] : itemLinks.join('\n'), receiptLinks.join('\n'),
       scanned ? 'ใช่' : 'ยัง', d.staff || '', scanned ? 'ไม่ต้อง' : 'รอตรวจสอบ', d.note || '', d.userId]);
    const startRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
    range.setValues(rows);
    // แถวที่ยังไม่ยิงรับ → สีเหลือง ให้ตามเช็ค
    range.setBackground(scanned ? null : '#fff3cd');

    return ContentService.createTextOutput(JSON.stringify({ok:true, time:timeStr, count:tracks.length}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() { return ContentService.createTextOutput('ok'); }
