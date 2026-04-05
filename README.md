# IT๙ Converter (Word Add-in)

เครื่องมือแปลงเลขไทยระดับมืออาชีพสำหรับ Microsoft Word ที่มาพร้อมกับระบบ "Smart Ignore" เพื่อความปลอดภัยในการใช้งานร่วมกับเอกสารที่มีภาษาอังกฤษและ URL

## ฟีเจอร์หลัก (Features)
- **Smart Ignore (ข้ามคำอังกฤษอัตโนมัติ)**: ระบบจะข้ามตัวเลขที่อยู่ในคำภาษาอังกฤษ (เช่น spin9, iPhone 15) และ URL/Email โดยอัตโนมัติ
- **Deep Search**: รองรับการแปลงตัวเลขในทุกส่วนของเอกสาร:
  - เนื้อหาหลัก (Body)
  - กล่องข้อความและรูปร่าง (Shapes & Textboxes)
  - ส่วนหัวและท้ายกระดาษ (Headers & Footers)
  - เลขหน้าและคำบรรยายภาพ (Fields & Page Numbers)
- **List Flattening (แช่แข็งลำดับรายการ)**: แปลงเลขลำดับอัตโนมัติ (1.1, 1.2) ให้เป็นข้อความเลขไทย (๑.๑, ๑.๒) แบบถาวร
- **Auto-Date Formatter**: แปลงวันที่ภาษาอังกฤษ (ค.ศ.) เป็นวันที่ไทย (พ.ศ.) พร้อมเลขไทยโดยอัตโนมัติ (เช่น 5 May 2024 -> ๕ พฤษภาคม ๒๕๖๗)
- **Formatting Preservation**: รักษาฟอนต์ สี และขนาดดั้งเดิมไว้ครบถ้วน

## วิธีการใช้งานสำหรับ MacOS (How to Use)

เพื่อให้ปุ่ม **IT๙** ปรากฏในโปรแกรม Word บน Mac ให้ทำตามขั้นตอนดังนี้:

1. **Copy ไฟล์ Manifest**: คัดลอกไฟล์ `manifest.xml` จากโฟลเดอร์โปรเจกต์ไปไว้ที่:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml`
   *(หากไม่มีโฟลเดอร์ `wef` ให้สร้างขึ้นมาใหม่)*
2. **Restart Word**: ปิดโปรแกรม Word แบบสนิท (Cmd + Q) แล้วเปิดใหม่
3. **เปิดใช้งาน Add-in**:
   - ไปที่แถบ **หน้าแรก (Home)**
   - มองหาปุ่ม **Add-ins** ทางด้านขวา (หรือไปที่ **Insert** > **Add-ins** > **My Add-ins**)
   - เลือกหัวข้อ **DEVELOPER ADD-INS**
   - จะพบ **IT๙ Converter** ปรากฏอยู่ ให้กดเพิ่ม (Add)

## วิธีการใช้งานสำหรับ Word Online (word.new)
1. ไปที่แถบ **หน้าแรก (Home)** > **Add-ins** > **More Add-ins**
2. เลือก **My Add-ins** > **Upload My Add-in**
3. อัปโหลดไฟล์ `manifest.xml` จากเครื่องของคุณ

---
**พัฒนาโดย:** Vitchakorn Poonyakanok (IT๙)
**เวอร์ชัน:** 1.10.0
