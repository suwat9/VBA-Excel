#ที่จัดเก็บโปรแกรม VBA for Microsoft Excel
1. Arabic2Thai ใช้แปลงเลขอาระบิคเป็นเลขไทย โดยสร้างสไตล์ชื่อ ArabicToThai พร้อมแสดงตัวอย่างการใช้ในช่อง A1 หากต้องการสร้างสไตล์อย่างเดียวให้ไปลบ code ในส่วนของการแสดงตัวอย่างออก 
  * หลังการ Run สามารถเรียกใช้สไตล์ ArabicToThai ได้ตลอดเวลา ไปกำหนด Cell Style ให้กับเซลที่จะแสดงเลขไทยได้เลย
1. ClearStyle(name) ใช้สำหรับลบ Style ที่ระบุผ่านตัวแปร name
  * ClearStyle("Demo1") คือลบ Style ชื่อ Demo1 ออกจาก MS Excel
1. ClearExtendStyles() ใช้ลบ Style ที่สร้างเพิ่มจาก Style เดิมของ MS Excel
  * ClearExtendStyles()
1. MakeAStyle() ใช้สร้าง Style
  * MakeAStyle()
1. createSheet(name) ใช้สร้างชีทชื่อ name
  * createSheet("Demo") คือสร้างชีทชื่อ Demo
1. deleteSheet(name) ใช้ลบชีทตาม name ที่ระบุ
  * DeleteSheet("Demo") คือ ลบชีทชื่อ Demo
1. ListStyles() ใช้แสดงรายชื่อ Style ทั้งหมดออกมาที่ชีทชื่อ Config-Styles
  * ListStyles() 
1. Multi_FindReplace() ใช้สำหรับค้นหาและแทนที่คำทั้งหมด ครั้งละหลาย ๆ คำ โดยเปลี่ยนค่าใน fndList และ rplcList ซึ่งก็คือ รายการคำค้น(Find List) และ รายการคำแทน(Replace List)
