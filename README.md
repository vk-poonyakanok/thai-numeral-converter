# IT๙ Converter (Word Add-in)

เครื่องมือแปลงเลขไทยสำหรับ Microsoft Word ที่มาพร้อมกับระบบ "Smart Ignore" เพื่อความปลอดภัยในการใช้งานร่วมกับเอกสารที่มีภาษาอังกฤษและ URL

## ฟีเจอร์หลัก (Features)
- **Smart Ignore (ข้ามคำอังกฤษอัตโนมัติ)**: ระบบจะข้ามตัวเลขที่อยู่ในคำภาษาอังกฤษ (เช่น spin9, iPhone 15) และ URL/Email โดยอัตโนมัติ
- **Deep Search & Flatten**: ปุ่มพิเศษสำหรับแปลงและ "แช่แข็ง" องค์ประกอบที่เข้าถึงยาก:
  - ส่วนหัวและท้ายกระดาษ (Headers & Footers)
  - เลขหน้าและคำบรรยายภาพ (Fields & Page Numbers)
  - ลำดับรายการอัตโนมัติ (Auto-lists 1.1 -> ๑.๑)
- **Formatting Preservation**: รักษาฟอนต์ สี และขนาดดั้งเดิมไว้ครบถ้วน

## 📥 ดาวน์โหลด (One-Click Download)
[**คลิกที่นี่เพื่อดาวน์โหลดไฟล์ manifest.xml**](https://vk-poonyakanok.github.io/thai-numeral-converter/manifest.xml)

## เอกสารทางกฎหมาย (Legal Documents)
สำหรับการตรวจสอบโดย Microsoft Partner Center:
- **Privacy Policy**: [Link](https://vk-poonyakanok.github.io/thai-numeral-converter/privacy.html)
- **End User License Agreement (EULA)**: [Link](https://vk-poonyakanok.github.io/thai-numeral-converter/eula.html)

## วิธีการใช้งานสำหรับ MacOS (How to Use)

เพื่อให้ปุ่ม **IT๙** ปรากฏในโปรแกรม Word บน Mac ให้ทำตามขั้นตอนดังนี้:

1. **Download & Copy**: [ดาวน์โหลดไฟล์ `manifest.xml`](https://vk-poonyakanok.github.io/thai-numeral-converter/manifest.xml) แล้วคัดลอกไปไว้ที่:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml`
   *(หากไม่มีโฟลเดอร์ `wef` ให้สร้างขึ้นมาใหม่)*
2. **Restart Word**: ปิดโปรแกรม Word (Cmd + Q) แล้วเปิดใหม่
3. **เปิดใช้งาน Add-in**:
   - ไปที่แถบ **หน้าแรก (Home)**
   - มองหาปุ่ม **Add-ins** ทางด้านขวา (หรือไปที่ **Insert** > **Add-ins** > **My Add-ins**)
   - เลือกหัวข้อ **DEVELOPER ADD-INS**
   - จะพบ **IT๙ Converter** ปรากฏอยู่ ให้กดเพิ่ม (Add)

## การตั้งค่า
- **แปลงเนื้อหาหลัก**: แปลงเฉพาะตัวเลขในส่วนเนื้อหาหลักของเอกสาร
- **แปลงรวมเลขหน้าและเลขหัวข้อแบบ Flatten**: แปลงและ flatten เลขในหัวกระดาษ, เลขหน้า และลำดับรายการอัตโนมัติ (จะกลายเป็นข้อความถาวร)

---
**พัฒนาโดย:** Vitchakorn Poonyakanok (IT๙)
**เวอร์ชัน:** 1.22.0
