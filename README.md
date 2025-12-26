**Excel Extractor Application**
โปรแกรมสำหรับ ดึงข้อมูลจากไฟล์ Excel หลายไฟล์พร้อมกัน
โดยโครงสร้างไฟล์เหมือนกัน (Pattern เดียวกัน)
ผู้ใช้เพียงเลือกไฟล์ Excel → โปรแกรมจะรวมข้อมูลออกมาเป็นไฟล์เดียว

**Features**

รองรับการเลือก Excel หลายไฟล์พร้อมกัน

ดึงข้อมูลจาก Sheet เดียวกันทุกไฟล์

ดึงข้อมูลช่วงคอลัมน์ A–AH

อ่านข้อมูลตั้งแต่แถวข้อมูลแรกลงไปจนกว่าจะเจอแถวว่าง

รวมข้อมูลทั้งหมดเป็นไฟล์ Excel ใหม่ 1 ไฟล์

มี GUI ใช้งานง่าย (ไม่ต้องพิมพ์คำสั่ง)

รองรับการ pack เป็นไฟล์ .exe

**โครงสร้างไฟล์**
excel_extractor_app/
│
├─ gui.py               # หน้าจอโปรแกรม (GUI)
├─ extractor.py         # Logic ดึงข้อมูลจาก Excel
├─ documentation.ico    # ไอคอนโปรแกรม (optional)
└─ README.md

**ข้อกำหนดของไฟล์ Excel**

ชื่อ Sheet ต้องเป็น
System Daily Report Entry

หัวตารางอยู่ที่แถวบน (เช่น แถว 1–2)

ข้อมูลเริ่มที่ A3 – AH3

อ่านลงไปเรื่อย ๆ จนกว่า คอลัมน์ A จะว่าง

**Requirements**
Python
แนะนำ: Python 3.10+ (64-bit)
ตรวจสอบเวอร์ชัน: python --version

Python Libraries
ติดตั้งไลบรารีที่จำเป็น: pip install pandas openpyxl PySide6
ถ้าใช้ Python หลายเวอร์ชัน: "C:/Program Files/Python314/python.exe" -m pip install pandas openpyxl PySide6

วิธีรันโปรแกรม (แบบไม่ pack)
โค้ดรัน : python gui.py

Pack เป็นไฟล์ .exe
ติดตั้ง PyInstaller : pip install pyinstaller

คำสั่ง Pack (พร้อม icon)
  -m PyInstaller `
  --onefile `
  --windowed `
  --name ExcelExtractor `
  --icon documentation.ico `
  --hidden-import PySide6.QtWidgets `
  --hidden-import openpyxl `
  gui.py

 ** เขียนบรรทัดเดียว (กันพลาด)**
 & "C:/Program Files/Python314/python.exe" -m PyInstaller --onefile --windowed --name ExcelExtractor --icon documentation.ico --hidden-import PySide6.QtWidgets --hidden-import openpyxl gui.py


ไฟล์ .exe จะอยู่ที่: dist/ExcelExtractor.exe


