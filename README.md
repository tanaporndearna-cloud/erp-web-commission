# TRC Commission Calculator

Streamlit app คิดค่าคอม ERP → เขียนลง Google Sheet อัตโนมัติ

## วิธี Deploy บน Streamlit Community Cloud

### ขั้นตอนที่ 1: สร้าง GitHub Repository
1. ไปที่ github.com → New repository ชื่อ `trc-commission-app`
2. Upload ไฟล์ทั้งหมดในโฟลเดอร์นี้ขึ้น GitHub

### ขั้นตอนที่ 2: สร้าง Google Service Account
1. ไปที่ console.cloud.google.com
2. สร้าง Project ใหม่ (หรือใช้ที่มีอยู่)
3. เปิดใช้งาน **Google Sheets API** และ **Google Drive API**
4. ไปที่ IAM & Admin → Service Accounts → Create Service Account
5. ตั้งชื่อ เช่น `trc-commission-bot`
6. คลิก Keys → Add Key → Create new key → JSON → Download
7. **เปิด Google Sheet "test commission SL/2026"** แล้ว Share ให้ email ของ Service Account (ตรง `client_email` ในไฟล์ JSON) เป็น Editor

### ขั้นตอนที่ 3: Deploy บน Streamlit Community Cloud
1. ไปที่ share.streamlit.io → Sign in ด้วย GitHub
2. New app → เลือก repo `trc-commission-app` → Main file: `app.py`
3. ก่อน Deploy: คลิก **Advanced settings → Secrets**
4. วาง secrets ดังนี้:

```toml
[gcp_service_account]
type = "service_account"
project_id = "xxx"
private_key_id = "xxx"
private_key = "-----BEGIN RSA PRIVATE KEY-----\nxxx\n-----END RSA PRIVATE KEY-----\n"
client_email = "trc-commission-bot@xxx.iam.gserviceaccount.com"
client_id = "xxx"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/trc-commission-bot%40xxx.iam.gserviceaccount.com"
```

> **หมายเหตุ:** ค่าทั้งหมดอยู่ในไฟล์ JSON ที่ดาวน์โหลดมา

5. คลิก Deploy!

### วิธีใช้งาน
1. เปิด URL ของ Streamlit app
2. อัปโหลดไฟล์ ERP (.xlsx)
3. อัปโหลดไฟล์ PPS (ถ้ามี)
4. เลือกเดือน/ปี
5. กด "คิดค่าคอม"
6. ข้อมูลจะเขียนลง Google Sheet อัตโนมัติ + ดาวน์โหลดไฟล์ได้

## โครงสร้างไฟล์
```
trc-commission-app/
├── app.py                  # หน้า Streamlit หลัก
├── pipeline_runner.py      # รัน ERP pipeline
├── google_writer.py        # เขียนลง Google Sheet
├── requirements.txt        # dependencies
├── .streamlit/
│   └── secrets.toml        # credentials (อย่า commit ขึ้น GitHub!)
└── scripts/
    ├── decrypt_erp.py
    ├── preprocess_erp.py
    ├── build_pps.py
    ├── build_pps_from_erp.py
    ├── build_t_sheets.py
    ├── build_cal05.py
    └── build_sum.py
```

## หมายเหตุ
- ไฟล์ `.streamlit/secrets.toml` ต้องไม่ commit ขึ้น GitHub — ใส่ credentials ผ่าน Streamlit Cloud UI แทน
- เพิ่ม `.streamlit/secrets.toml` ใน `.gitignore`
