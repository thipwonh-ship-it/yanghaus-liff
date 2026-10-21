# YANG Haus — ระบบบันทึกส่งของผ่าน LINE

พนักงานนำพัสดุไปวางที่สาขาขนส่ง → กดปุ่มในไลน์ YANG Haus → สแกนบาร์โค้ด + ถ่ายรูป → ระบบบันทึก เวลา / GPS / รูป / สถานะยิงรับ ลง Google Sheet อัตโนมัติ

## ส่วนประกอบ (3 ชิ้น)

| ชิ้น | อยู่ที่ไหน | หน้าที่ |
|---|---|---|
| **index.html** (หน้าฟอร์ม) | GitHub repo `thipwonh-ship-it/yanghaus-liff` → GitHub Pages `https://thipwonh-ship-it.github.io/yanghaus-liff/` | หน้าจอที่เปิดในไลน์ สแกนบาร์โค้ด ถ่ายรูป ขอ GPS แล้วส่งข้อมูลไป Apps Script |
| **Code.gs** (หลังบ้าน) | Google Sheet "YANG Haus ส่งของ" → ส่วนขยาย → Apps Script → Deploy เป็น Web app | รับข้อมูล บันทึกรูปลง Drive เขียนแถวลงชีต |
| **LINE** | developers.line.biz → Provider **YangHaus** | OA "YANG Haus" (Messaging API) ใช้ทำ Rich menu / Channel "YH ขนส่ง Login" มี LIFF ID `2012109669-jTtL2UEr` ชี้ Endpoint URL ไปที่ GitHub Pages |

ข้อมูลอยู่ที่ Google Drive โฟลเดอร์ **YANG Haus ขนส่ง**
- ชีต **YANG Haus ส่งของ** (spreadsheetId `1t9phG6yVcseuX042WWoExS-nNbSnnOaUb3jJSDllzLc`)
- โฟลเดอร์ **รูปส่งของ** (id `1skXBDHIJXVrswcniLbYAYQlxaN56HJy_`) แยกโฟลเดอร์รายวัน ชื่อไฟล์ = วันที่_เวลา_ขนส่ง_เลขพัสดุ

เปิดฟอร์มตรง: `https://liff.line.me/2012109669-jTtL2UEr`

## ขั้นตอนใช้งานของพนักงาน
1. กดปุ่ม "📦 ส่งของที่สาขา" ในไลน์ YANG Haus
2. เลือกขนส่ง + พิมพ์สาขา
3. **ถ่ายรูปสินค้าทีละชิ้น** ให้เห็นใบปะหน้า → ระบบอ่านเลขพัสดุจากรูปเอง (ถ้าอ่านไม่ติดมีปุ่มสแกนด้วยกล้อง หรือพิมพ์)
   - รูปที่อ่านบาร์โค้ดไม่ได้จะไม่รับ ต้องถ่ายใหม่
   - เลขพัสดุซ้ำ (ในฟอร์ม หรือเคยบันทึกในชีต) จะถูกปฏิเสธ
4. ตอบ "ขนส่งยิงรับแล้วหรือยัง"
   - ยิงแล้ว → ถ่ายสลิป/หน้าจอที่ยิงรับ 1 รูป
   - ยังไม่ยิง → ถ่ายรูปของทั้งหมด 1 รูป + ติ๊กยืนยันเห็นครบ N ชิ้น + ระบุพนักงานที่รับฝาก → แถวในชีตเป็นสีเหลือง "รอตรวจสอบ"
5. บันทึก → ข้อความสรุปเด้งในแชท

## การไหลของข้อมูล
```
ไลน์ (Rich menu) → LIFF เปิด index.html (GitHub Pages)
  → liff.getProfile() ได้ชื่อ+userId, navigator.geolocation ได้ GPS
  → POST JSON {carrier, branch, tracking, scanned, staff, note, lat, lng, photos[base64]} ไป GAS_URL
Code.gs doPost → เซฟรูปลง Drive/รูปส่งของ/YYYY-MM-DD/ → append แถวลงชีต (1 แถว / 1 เลขพัสดุ)
  → ตอบ {ok, time, count} → หน้าฟอร์มส่งข้อความสรุปเข้าแชท แล้วปิดตัวเอง
```

## คอลัมน์ในชีต
A วันที่ | B เวลา | C ผู้ส่ง | D ขนส่ง | E สาขา | F เลขพัสดุ | G จำนวนชิ้นรวม | H GPS | I รูปสินค้า | J รูปใบรับฝาก/เคาน์เตอร์ | K ยิงรับแล้ว (ใช่/ยัง) | L พนักงานรับฝาก | M ตรวจสอบแล้ว | N หมายเหตุ | O LINE userId

แถว **สีเหลือง** = ยังไม่ยิงรับ → ตอนเย็นเช็คเลขพัสดุใน Seller Center แล้วแก้ M เป็น "ยิงแล้ว" / ถ้ายังไม่ยิงให้แจ้งขนส่งพร้อมหลักฐานจากแถวนั้น

## วิธีแก้ไข / อัปเดต
**แก้หน้าฟอร์ม** (เพิ่มช่อง เปลี่ยนข้อความ เพิ่มขนส่ง)
1. แก้ index.html
2. GitHub → repo yanghaus-liff → Add file → Upload files → ลากทับ → Commit
3. รอ 1–2 นาที ไม่ต้องแก้อะไรใน LINE

**แก้หลังบ้าน** (เพิ่มคอลัมน์ เปลี่ยนที่เก็บ)
1. แก้ Code.gs ใน Apps Script → บันทึก
2. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (ถ้าไม่ทำ โค้ดเก่าจะยังทำงาน)
3. URL เดิม ไม่ต้องแก้ index.html
4. ถ้าเพิ่มคอลัมน์ ให้แก้หัวตารางในชีตแถว 1 ให้ตรงด้วย

**ย้ายชีต/โฟลเดอร์** → แก้ `SHEET_ID` / `FOLDER_ID` บรรทัดบนสุดของ Code.gs แล้ว Deploy ใหม่

**เพิ่มพนักงาน** → ให้แอดไลน์ OA "YANG Haus" แล้วกด Rich menu ได้เลย ไม่ต้องตั้งค่าอะไร

**ปุ่มในไลน์** → manager.line.biz → YANG Haus → ริชเมนู → ลิงก์ = `https://liff.line.me/2012109669-jTtL2UEr`

## ค่าที่ตั้งไว้ในโค้ด
- index.html บรรทัดบนของ `<script>`: `LIFF_ID`, `GAS_URL`
- Code.gs บรรทัดบนสุด: `SHEET_ID`, `FOLDER_ID`
- ขนส่งใน dropdown: SPX / J&T / Flash / BEST / อื่นๆ (แก้ที่ `<select id="carrier">`)
- สาขาประจำต่อขนส่ง: แก้ที่ `const BRANCHES` ใน index.html (SPX=วัดกระทุ่ม, J&T=74/2, Flash=หลังโรงไม้, BEST=รามสอง) เพิ่มได้หลายสาขาต่อขนส่ง
- บาร์โค้ดที่อ่าน: CODE_128, CODE_39, QR (ไลบรารี html5-qrcode จาก cdnjs)

## ข้อจำกัดที่ควรรู้
- GPS ต้องกด "อนุญาต" ครั้งแรก ถ้าไม่อนุญาตจะบันทึกโดยไม่มีพิกัด
- รูปถูกย่อเหลือกว้าง 1280px ก่อนส่ง (ประหยัด Drive)
- Apps Script ฟรีรับได้ ~ 50MB/ครั้ง ถ้าส่งรูปเยอะมาก (>20 รูป) อาจช้า
- ใครมีลิงก์รูปใน Drive เปิดดูได้ (ตั้ง ANYONE_WITH_LINK เพื่อให้แนบเคลมขนส่งได้)

## บัญชีที่ใช้
Google: thipwon.h@gmail.com (เจ้าของชีต/Drive/Apps Script) · GitHub: thipwonh-ship-it · LINE Developers: บัญชีไลน์ของ Thipwon

## account/ — จดบัญชีผ่าน LINE
ระบบที่ 2 ในไลน์เดียวกัน: ถ่ายรูปใบเสร็จ → ลงชีทบัญชี + เก็บรูปใน Drive ดู [account/README.md](account/README.md)
