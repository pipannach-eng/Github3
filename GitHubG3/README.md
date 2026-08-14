# G3 Dashboard Deploy

ไฟล์สำหรับ deploy dashboard แบบ static web ผ่าน GitHub Pages

ไฟล์หลัก:
- `index.html`
- `styles.css`
- `app.js`
- `data.js`

วิธีอัปเดตข้อมูล:
1. รัน `C:\Users\User\Documents\G3\dashboard\build_dashboard_data.ps1`
2. คัดลอก `dashboard\data.js` มาแทนไฟล์ `data.js` ในโฟลเดอร์นี้
3. commit และ push ขึ้น GitHub

วิธีเปิดใช้ GitHub Pages:
1. push repository นี้ขึ้น GitHub
2. ไปที่ `Settings > Pages`
3. เลือก `Deploy from a branch`
4. เลือก branch `main`
5. เลือก folder `/ (root)`

หมายเหตุ:
- dashboard ใช้ชื่อแฝงโรงเรียนในส่วนกราฟรายโรงเรียน
- ระดับคุณภาพ A/B/C/D ใช้สีเดียวกันทั้ง dashboard และ D เป็นสีแดง
