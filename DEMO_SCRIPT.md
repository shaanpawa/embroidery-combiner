# Micro Automation — Embroidery Stacker Demo Script
# สคริปต์สาธิต Micro Automation — Embroidery Stacker

**Live URL:** https://embroidery-combiner.vercel.app
**Local URL:** http://localhost:3000
**Login Password:** micro2026
**Sample Data:** Built-in (click "Load sample data" / "โหลดข้อมูลตัวอย่าง")

---

## Pre-Demo Checklist / เตรียมตัวก่อนสาธิต

- [ ] Browser open to the app URL / เปิดเบราว์เซอร์ที่ URL ของแอป
- [ ] Have sample Excel + DST zip ready (or use built-in sample data)
  - มีไฟล์ Excel ตัวอย่าง + DST zip พร้อม (หรือใช้ข้อมูลตัวอย่างในระบบ)
- [ ] If using localhost: backend running on port 8000, frontend on port 3000

---

## PART 1: Introduction / แนะนำ (1 min)

### English:
> "Welcome to Micro Automation — Embroidery Stacker. This tool automates the most time-consuming part of embroidery production: **turning name orders into production-ready combo files.**
>
> Right now, this process is fully manual — combining DST files one by one, assigning MA and COM numbers by hand, checking for errors. Our tool does all of this automatically in minutes."

### ภาษาไทย:
> "ยินดีต้อนรับสู่ Micro Automation — Embroidery Stacker เครื่องมือนี้ทำให้ขั้นตอนที่ใช้เวลานานที่สุดในการผลิตงานปักเป็นอัตโนมัติ: **เปลี่ยนออเดอร์ชื่อให้เป็นไฟล์คอมโบที่พร้อมผลิต**
>
> ตอนนี้ขั้นตอนนี้เป็นงานมือทั้งหมด — รวมไฟล์ DST ทีละตัว กำหนดเลข MA และ COM ด้วยมือ ตรวจสอบข้อผิดพลาด เครื่องมือของเราทำทั้งหมดนี้อัตโนมัติภายในไม่กี่นาที"

---

## PART 2: Login / เข้าสู่ระบบ (30 sec)

**Action / ทำ:** Enter password `micro2026` → Click "Sign In" / ใส่รหัส `micro2026` → กด "Sign In"

> EN: "The system is password-protected so only authorized staff can access it."
>
> TH: "ระบบมีรหัสผ่านป้องกัน เฉพาะพนักงานที่ได้รับอนุญาตเท่านั้นที่เข้าถึงได้"

---

## PART 3: Start Session / เริ่มเซสชัน (30 sec)

**Action / ทำ:** Type session name (e.g. today's date) → Click "Start Session"

> EN: "Create a new session — name it by date or order batch. Previous sessions are listed below with a timer showing when they expire."
>
> TH: "สร้างเซสชันใหม่ — ตั้งชื่อตามวันที่หรือล็อตออเดอร์ เซสชันก่อนหน้าแสดงด้านล่างพร้อมนาฬิกาบอกเวลาหมดอายุ"

**Point out:** Session expiry timer on existing sessions (e.g. "18h left")

---

## PART 4: MA & COM Assignment / กำหนดเลข MA & COM (3 min)

**What the audience sees:** Two cards side by side — "Generate MA & COM" (recommended) and "I already have MA & COM"

### Step 4a: Upload Excel

**Action / ทำ:** Click "Generate MA & COM" drop zone → Upload the order Excel

> EN: "You'll see two options. Most of the time you'll use the left card — upload your order Excel and we'll automatically assign MA and COM numbers. If you already have them, use the right card to skip straight to the stacker."
>
> TH: "จะเห็นสองตัวเลือก ส่วนใหญ่จะใช้การ์ดซ้าย — อัปโหลด Excel ออเดอร์แล้วระบบจะกำหนดเลข MA และ COM อัตโนมัติ ถ้ามีอยู่แล้ว ใช้การ์ดขวาเพื่อข้ามไปที่ stacker เลย"

### Step 4b: Column Detection

**Action / ทำ:** System auto-detects columns → Review the 4 fields (Size, Fabric Colour, Frame Colour, Embroidery Colour) → Click "Generate MA & COM"

> EN: "The system detected which columns contain the size and colour information. If any are wrong, just click the field and then click the correct column. When everything looks right, hit Generate."
>
> TH: "ระบบตรวจจับว่าคอลัมน์ไหนมีข้อมูลไซส์และสี ถ้าอันไหนผิดก็แค่คลิกที่ฟิลด์แล้วคลิกคอลัมน์ที่ถูกต้อง เมื่อทุกอย่างถูกต้องแล้วกดสร้าง"

### Step 4c: Review Results

**What the audience sees:** Summary stats (MA Groups, COM Groups, Total Rows) + grouped MA cards with COM tables inside

> EN: "Here you see the results. Each MA group represents a different size — MA1, MA2, and so on. Inside each MA, COM numbers group rows that share the same fabric, frame, and embroidery colour combination. Notice the colour-coded columns — blue for fabric, purple for frame, green for embroidery."
>
> TH: "ตรงนี้เห็นผลลัพธ์ แต่ละกลุ่ม MA แทนไซส์ต่างกัน — MA1, MA2 เป็นต้น ภายในแต่ละ MA เลข COM จัดกลุ่มแถวที่มีสีผ้า สีเฟรม และสีปักเหมือนกัน สังเกตคอลัมน์สี — น้ำเงินคือผ้า ม่วงคือเฟรม เขียวคือปัก"

### Step 4d: Download Updated Excel

**Action / ทำ:** Click the outlined "Download Updated Excel" button

> EN: "You can download the Excel with MA and COM columns added at the end — your original data is untouched. This is useful for your records."
>
> TH: "ดาวน์โหลด Excel ที่เพิ่มคอลัมน์ MA และ COM ต่อท้ายได้ — ข้อมูลเดิมไม่เปลี่ยน เป็นประโยชน์สำหรับเก็บบันทึก"

### Step 4e: Proceed

**Action / ทำ:** Click "Proceed to Embroidery Stacker →"

---

## PART 5: Upload Order / อัปโหลดออเดอร์ (1 min)

**What the audience sees:** Excel + DST upload zones, column mapping step

> EN: "Now we map the production columns — Program Number, Name, Quantity, Combo Number, and Machine Program. The system auto-detects these too. Just confirm and proceed."
>
> TH: "ตอนนี้เราแมปคอลัมน์การผลิต — เลขโปรแกรม ชื่อ จำนวน เลขคอมโบ และเครื่องจักร ระบบตรวจจับอัตโนมัติเช่นกัน แค่ยืนยันแล้วดำเนินการต่อ"

**Action / ทำ:** Review auto-detected columns → Confirm mapping

---

## PART 6: Upload Programs / อัปโหลดโปรแกรม (1 min)

**Action / ทำ:** Upload DST zip file → See matching results

> EN: "Upload your DST program files as a zip. The system matches each program number from the Excel to its DST file. If any are missing, you'll see a clear warning with the exact file numbers."
>
> TH: "อัปโหลดไฟล์โปรแกรม DST เป็น zip ระบบจับคู่เลขโปรแกรมจาก Excel กับไฟล์ DST ถ้าไฟล์ไหนหาย จะเห็นคำเตือนชัดเจนพร้อมเลขไฟล์"

**Point out:** "All 300 programs matched" / "300 โปรแกรมจับคู่ครบ"

---

## PART 7: Export / ส่งออก (1 min)

**Action / ทำ:** Review combo list → Click "Export X files" → Download zip

> EN: "Review your combos — each one shows which names go in which slots, left and right columns matching your machine layout. Hit Export, download the zip, and load it straight into the machine."
>
> TH: "ตรวจสอบคอมโบ — แต่ละตัวแสดงว่าชื่อไหนอยู่ช่องไหน คอลัมน์ซ้ายขวาตรงกับเลย์เอาต์เครื่อง กดส่งออก ดาวน์โหลด zip แล้วโหลดเข้าเครื่องได้เลย"

**Point out:** "300 names → 31 combos in seconds" / "300 ชื่อ → 31 คอมโบในไม่กี่วินาที"

---

## PART 8: Key Benefits / ข้อดีสำคัญ (1 min)

### English:
> "To summarize what Embroidery Stacker delivers:
> 1. **Speed** — Hours of manual work → minutes
> 2. **Accuracy** — Auto-matching eliminates human errors
> 3. **MA & COM Assignment** — Automatic grouping by size and colour
> 4. **Validation** — Missing files caught before production, not during
> 5. **Bilingual** — Full Thai and English support
> 6. **No training needed** — Uses your existing Excel workflow"

### ภาษาไทย:
> "สรุปสิ่งที่ Embroidery Stacker มอบให้:
> 1. **ความเร็ว** — งานมือหลายชั่วโมง → ไม่กี่นาที
> 2. **ความแม่นยำ** — จับคู่อัตโนมัติกำจัดข้อผิดพลาดจากคน
> 3. **กำหนด MA & COM** — จัดกลุ่มอัตโนมัติตามไซส์และสี
> 4. **ตรวจสอบ** — จับไฟล์ที่หายก่อนเข้าผลิต ไม่ใช่ระหว่างผลิต
> 5. **สองภาษา** — รองรับภาษาไทยและอังกฤษเต็มรูปแบบ
> 6. **ไม่ต้องเทรนนิ่ง** — ใช้ Excel เดิมที่คุ้นเคย"

---

## Additional Features / ฟีเจอร์เพิ่มเติม

| Feature | EN | TH |
|---------|----|----|
| Language Toggle | Click "TH"/"EN" in top nav | กดปุ่ม "TH"/"EN" ที่แถบบน |
| Dark/Light Mode | Click sun/moon icon | กดไอคอนดวงอาทิตย์/พระจันทร์ |
| Session History | Sessions auto-save for 24 hours | เซสชันบันทึกอัตโนมัติ 24 ชั่วโมง |
| Session Timer | Each session shows time remaining | แต่ละเซสชันแสดงเวลาคงเหลือ |

---

## Q&A / ถาม-ตอบ

| Question | EN | TH |
|----------|----|----|
| File formats? | Excel (.xlsx) in, DST out | รับ Excel (.xlsx) ส่งออก DST |
| Max names per combo? | 20 slots (10L + 10R) | 20 ช่อง (10 ซ้าย + 10 ขวา) |
| Excel format changes? | Auto-detects columns | ตรวจจับคอลัมน์อัตโนมัติ |
| Internet required? | Yes for cloud, no for local desktop app | ใช่สำหรับคลาวด์ ไม่ใช้สำหรับแอปบนเครื่อง |
| How long are sessions saved? | 24 hours with visible countdown | 24 ชั่วโมงพร้อมนับถอยหลัง |
| What if colours are inconsistent? | Auto-normalized (White = white = WHITE) | ปรับอัตโนมัติ (White = white = WHITE) |

---

## Demo Timing / เวลาสาธิต

| Part | Time |
|------|------|
| 1. Introduction / แนะนำ | 1 min |
| 2. Login / เข้าสู่ระบบ | 30 sec |
| 3. Start Session / เริ่มเซสชัน | 30 sec |
| 4. MA & COM Assignment / กำหนด MA & COM | 3 min |
| 5. Upload Order / อัปโหลดออเดอร์ | 1 min |
| 6. Upload Programs / อัปโหลดโปรแกรม | 1 min |
| 7. Export / ส่งออก | 1 min |
| 8. Key Benefits / ข้อดี | 1 min |
| **Total / รวม** | **~10 min** |
