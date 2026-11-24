# Rich menu — OA "YH In House" (2 ชุดตามบทบาท ตั้งแต่ 2026-11-24)

| ชุด | ใครเห็น | รูป | ตั้งที่ไหน |
|---|---|---|---|
| **พนักงาน** (default) | ทุกคนที่แอด OA และไม่ใช่ผู้ดูแล | `yanghaus-richmenu-staff.jpg` 2500×843 (compact 3 คอลัมน์) | OA Manager → Rich menus (เมนู "InhouseJobs") |
| **ผู้ดูแล** | Thipwon + คุณอ้อย (ผูกรายคน) | `yanghaus-richmenu-admin.jpg` 2500×1686 (large 6 ช่อง) | Messaging API — รัน `richMenuAdminSetup` ใน Apps Script ของระบบจดบัญชี (`account/Code.gs`) |

ชุดรายคนถูกแสดงแทน default เสมอ (LINE: per-user rich menu มี priority สูงกว่า) · ริชเมนูที่สร้างทาง API ไม่ขึ้นใน OA Manager และแก้ที่นั่นไม่ได้

## ปุ่มและลิงก์
ชุดพนักงาน (ซ้าย→ขวา):

| ปุ่ม | ลิงก์ |
|---|---|
| ส่งของที่สาขา | `https://liff.line.me/2012109669-jTtL2UEr` |
| ส่ง COD | `https://liff.line.me/2012109669-kGPX1UtI?mode=cod` |
| ตั้งเบิก | `https://liff.line.me/2012109669-N03WhvLZ` |

ชุดผู้ดูแล (แถวบน → แถวล่าง):

| ปุ่ม | ลิงก์ |
|---|---|
| จดบัญชี | `https://liff.line.me/2012109669-kGPX1UtI` |
| เช็คสต๊อก | `https://liff.line.me/2012109669-qQAVcyBR` |
| อนุมัติเบิก | `https://liff.line.me/2012109669-N03WhvLZ?mode=approve` |
| ส่งของที่สาขา | `https://liff.line.me/2012109669-jTtL2UEr` |
| ส่ง COD | `https://liff.line.me/2012109669-kGPX1UtI?mode=cod` |
| รายจ่ายเดือนนี้ | `https://liff.line.me/2012109669-kGPX1UtI?mode=dash` |

## ชุดผู้ดูแล — ตั้งครั้งแรก
1. LINE Developers Console → Provider YangHaus → channel **YH In House** (Messaging API) → แท็บ Messaging API → Channel access token (long-lived) → Issue → คัดลอก
2. Apps Script "YANG Haus จดบัญชี Line" → Project Settings → Script properties → เพิ่ม `LINE_OA_TOKEN` = token นั้น (**ห้ามใส่ใน repo / แชท**)
3. ใน editor เลือกฟังก์ชัน `richMenuAdminSetup` → Run → Log ขึ้น `linked: N`
4. ตรวจด้วย `richMenuStatus` (ตอบแค่ มี/ไม่มี + จำนวน)

หลังจากนั้น: เพิ่ม/ถอดผู้ดูแล (แท็บ ผู้ใช้ หรือ `COD_ADMINS`) → ระบบผูก/ถอดเมนูให้เองตอนคนนั้นเปิดหน้าจดบัญชีหรือตั้งเบิกครั้งถัดไป · เลิกใช้ชุดผู้ดูแล → รัน `richMenuAdminRemove`

## แก้รูป / ปุ่ม
- `menu.html` = ต้นฉบับ (รายการปุ่มอยู่ใน `B` และ `SETS`) · `python3 build.py` สร้าง jpg ทั้ง 2 ชุด (ต้องมี Google Chrome + playwright) · ฟอนต์ Prompt (OFL) อยู่ใน `fonts/`
- ชุดพนักงาน: อัปโหลดรูปใหม่ใน OA Manager → Edit → Image → Change
- ชุดผู้ดูแล: push รูปใหม่ขึ้น main แล้วรัน `richMenuAdminSetup` อีกครั้ง (สร้างเมนูใหม่ ลบของเดิม ผูกใหม่ทุกคน) · ถ้าเปลี่ยนลิงก์/จำนวนปุ่ม แก้ในฟังก์ชันเดียวกัน
- `yanghaus-richmenu-3btn.jpg` = ชุดเดิม (ส่งของ / จดบัญชี / เช็คสต๊อก) ที่ใช้ช่วง 24 พ.ย. 2026 ก่อนแยกบทบาท — เก็บไว้เผื่อย้อนกลับ (`menu.html?set=all3`)

หมายเหตุ: LINE บนคอม (Windows/macOS) ไม่แสดงริชเมนู — บนคอมให้เปิดลิงก์ LIFF ในเบราว์เซอร์แทน
