# YANG Haus — จดบัญชีผ่าน LINE (ถ่ายรูปใบเสร็จ → ลงชีท + เก็บรูปใน Drive)

กดปุ่มในไลน์ YANG Haus → ถ่ายรูปใบเสร็จ/ใบกำกับภาษี → AI (Gemini) อ่าน วันที่ ผู้ขาย เลขผู้เสียภาษี เลขที่ ยอด → ตรวจ/แก้ → บันทึก
→ แถวใหม่ในชีท **YANG Haus บัญชีซื้อ-ขาย 2026** (แท็บ `ค่าใช้จ่าย` หรือ `ซื้อ`) + รูปถูกเก็บใน Drive `ส่งบัญชี/YYYY-MM/ภาษีซื้อ` (ถ้ามีใบกำกับ) หรือ `ใบเสร็จอื่นๆ`

## ส่วนประกอบ

| ชิ้น | อยู่ที่ไหน | หน้าที่ |
|---|---|---|
| `index.html` | โฟลเดอร์ `account/` ใน repo yanghaus-liff → GitHub Pages `https://thipwonh-ship-it.github.io/yanghaus-liff/account/` | หน้าฟอร์มในไลน์: ถ่ายรูป ย่อรูป ส่งให้ Apps Script |
| `Code.gs` | ชีท YANG Haus บัญชีซื้อ-ขาย 2026 → ส่วนขยาย → Apps Script → Deploy เป็น Web app | เรียก Gemini อ่านรูป / เซฟรูปลง Drive / เขียนแถวลงชีท |
| LINE | developers.line.biz → Provider **YangHaus** → channel "YH ขนส่ง Login" → LIFF ใหม่ | ปุ่ม Rich menu เปิดหน้าฟอร์ม |

- ชีท: `17gY8l9pkJjIXqZWRfStYwga-xasJjef97HaHfI7KIB4`
- โฟลเดอร์ Drive ส่งบัญชี: `1Pfy5Bo6vfkxWDyglsAJwDhg9rh5pajV-` (สร้างโฟลเดอร์เดือน/โฟลเดอร์ย่อยเองถ้ายังไม่มี)
- ชื่อไฟล์รูป: `วันที่_ผู้ขาย_ยอด_เลขที่.jpg`

## ติดตั้งครั้งแรก (ทำตามลำดับ)

### 1. Gemini API key (ฟรี ไม่ต้องผูกบัตร)
1. เปิด https://aistudio.google.com/apikey ด้วยบัญชี thipwon.h@gmail.com
2. **Create API key** → copy ไว้ (ขึ้นต้น `AIza...`)
3. ห้ามวางคีย์ลงใน repo — ใส่ใน Apps Script ข้อ 2 เท่านั้น

### 2. Apps Script
1. เปิดชีท YANG Haus บัญชีซื้อ-ขาย 2026 → เมนู **ส่วนขยาย → Apps Script**
2. ลบโค้ดเดิมในหน้า แล้ววาง `Code.gs` ทั้งไฟล์ → บันทึก (ตั้งชื่อโปรเจกต์ เช่น "จดบัญชี LINE")
3. ⚙️ **Project Settings → Script properties → Add script property**
   - Property: `GEMINI_API_KEY` Value: คีย์จากข้อ 1 → Save
4. ด้านบนเลือกฟังก์ชัน `testConfig` → **Run** → อนุญาตสิทธิ์ (Advanced → Go to … (unsafe) → Allow) → ดู Execution log ต้องเห็นรายการ ธุรกิจ/ประเภท
5. **Deploy → New deployment** → ⚙️ เลือก **Web app**
   - Execute as: **Me** / Who has access: **Anyone** → Deploy
6. copy **Web app URL** (ลงท้าย `/exec`)

### 3. GitHub Pages
1. ไฟล์อยู่ในโฟลเดอร์ `account/` ของ repo **yanghaus-liff** (ใช้ Pages ของ repo เดิม ไม่ต้องตั้งค่าเพิ่ม)
2. (ข้าม — Claude push ให้แล้ว)
3. แก้ `account/index.html` 2 บรรทัดบนของ `<script>`: `GAS_URL` = URL จากข้อ 2.6 (LIFF_ID ใส่หลังข้อ 4)
4. Pages เปิดอยู่แล้วจาก repo นี้ → รอ 1–2 นาทีหลัง commit

### 4. LINE LIFF
1. https://developers.line.biz → Provider YangHaus → channel **YH ขนส่ง Login** → แท็บ **LIFF → Add**
   - LIFF app name: `จดบัญชี` / Size: **Full** / Endpoint URL: `https://thipwonh-ship-it.github.io/yanghaus-liff/account/`
   - Scopes: ✅ profile ✅ openid ✅ chat_message.write / Add
2. copy **LIFF ID** → ใส่ใน `index.html` บรรทัด `LIFF_ID` → upload ทับใน GitHub
3. manager.line.biz → YANG Haus → **ริชเมนู** → เพิ่มปุ่ม "🧾 จดบัญชี" → ประเภท ลิงก์ = `https://liff.line.me/<LIFF ID>`

## วิธีใช้
1. กดปุ่ม 🧾 จดบัญชี ในไลน์
2. เลือก **ค่าใช้จ่าย** หรือ **ซื้อยาง/ล้อ** → ถ่ายรูปใบเสร็จ
3. รอ AI อ่าน 3–5 วิ → ตรวจยอด/วันที่/ติ๊ก "มีใบกำกับภาษี" → บันทึก
4. ได้ข้อความสรุปในแชท + ลิงก์รูป / กด "จดใบต่อไป" ได้เลย

กติกาที่ฝังไว้
- ค่าน้ำมัน → ติ๊ก "มีใบกำกับ" ออกอัตโนมัติ (เคลม VAT ไม่ได้)
- เลขที่ใบกำกับซ้ำในแท็บเดิม → ถามก่อนบันทึก
- ไม่มี GEMINI_API_KEY / เกินโควตา → ยังใช้ได้ แค่กรอกเอง
- คอลัมน์สูตร (ฐานภาษี/VAT) ไม่ถูกเขียนทับ — ชีทคำนวณเอง → ขึ้นแท็บ **รายงานภาษีซื้อ** อัตโนมัติ

## คอลัมน์ที่เขียน
- **ค่าใช้จ่าย**: A วันที่ B ธุรกิจ C ประเภท D รายละเอียด E ยอด F มีใบกำกับ · I ผู้ขาย J อ้างอิงออเดอร์ K หมายเหตุ L เลขที่ใบกำกับ M เลขผู้เสียภาษี N รูปใบเสร็จ
- **ซื้อ**: A วันที่ B ธุรกิจ C ซัพพลายเออร์ D สินค้า E จำนวน F ยอดรวม G มีใบกำกับ · J เลขที่ใบกำกับ K ขนส่งเข้า L อ้างอิงออเดอร์ M หมายเหตุ N รูปใบเสร็จ O เลขผู้เสียภาษี

## แก้ไข / อัปเดต
- **แก้หน้าฟอร์ม** → แก้ `account/index.html` → upload ทับใน GitHub → รอ 1–2 นาที
- **แก้หลังบ้าน** → แก้ `Code.gs` ใน Apps Script → บันทึก → **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (ไม่ทำ = โค้ดเก่ายังทำงาน) URL เดิม
- **เปลี่ยนชีท/โฟลเดอร์** → แก้ `SHEET_ID` / `FOLDER_ID` บนสุดของ Code.gs → Deploy ใหม่
- **รุ่น AI** → `GEMINI_PREFER` บนสุดของ Code.gs คือลำดับที่อยากใช้ ระบบดึงรายชื่อรุ่นที่คีย์ใช้ได้จริงจาก Google (cache 6 ชม.) แล้วจับคู่ ถ้ารุ่นแรกล้น (503/429) ไปรุ่นถัดไปเอง รุ่นที่ใช้ได้ล่าสุดอยู่ใน Script property `GEMINI_MODEL_OK`
- รายการ dropdown (ธุรกิจ/ประเภท/ซัพพลายเออร์) ดึงจากแท็บ `ตั้งค่า` ของชีท แก้ที่ชีทได้เลย

## บัญชีที่ใช้
Google: thipwon.h@gmail.com · GitHub: thipwonh-ship-it · LINE Developers: บัญชีไลน์ของ Thipwon
