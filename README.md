# Printer Dashboard

Dashboard สรุปสัญญาเครื่องพิมพ์ จากชีต "รายละเอียดตามคำอธิบาย" ในไฟล์ `สรุปสัญญา_Printer.xlsx`

- เปิดดู: `index.html` (หรือผ่าน GitHub Pages) — ใส่ `?q=คำค้น` ต่อท้ายลิงก์เพื่อเปิดพร้อมผลค้นหา
- ค้นหาได้ทั้งตัวย่อและชื่อเต็ม เช่น `กดส.` หรือ `กองเทคโนโลยีดิจิทัล`

## ไฟล์

| ไฟล์ | หน้าที่ |
|---|---|
| `abbreviations.csv` | ตาราง ตัวย่อ → ชื่อเต็ม (แก้/เพิ่มเองได้ใน Excel) |
| `build_abbr.py` | สร้าง `abbreviations.csv` จาก `Computer Nameสค69.xlsx` (ดึงเฉพาะชื่อหน่วยงาน ไม่ดึง IP) |
| `build_data.py` | สร้าง `data.js` จาก `สรุปสัญญา_Printer.xlsx` + `abbreviations.csv` |

## อัปเดตข้อมูล

```
python build_abbr.py "path/to/Computer Nameสค69.xlsx"   # เมื่อไฟล์ชื่อหน่วยงานเปลี่ยน
python build_data.py "path/to/สรุปสัญญา_Printer.xlsx"
git add -A && git commit -m "update data" && git push
```

## เพิ่มชื่อเต็มเอง

เพิ่มแถวใน `abbreviations.csv` โดยใส่คอลัมน์ `ที่มา` = `เพิ่มเอง` (แถวนี้จะไม่ถูกลบเมื่อรัน `build_abbr.py` ใหม่)

- `สังกัด` ว่าง = ใช้ได้ทุกที่ (กอง/ฝ่าย/กฟส./กฟจ.)
- `สังกัด` = `กดส.` = แผนกในกองนั้น · `กฟส` / `กฟจ` = แผนกในการไฟฟ้าสาขา/จังหวัด

ตัวอย่าง: `กบห.,,ชื่อเต็มของ กบห.,เพิ่มเอง`
