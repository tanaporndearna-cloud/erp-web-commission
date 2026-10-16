#!/usr/bin/env python3
"""
generate_payslips.py — สร้างสลีปเงินเดือนรายเดือน TRC Motorsport
สร้างไฟล์ Google Sheets ใหม่ใน folder ที่กำหนด และ generate ชีตพนักงานทุกคน

Usage (standalone):
    python3 generate_payslips.py

Usage (from Streamlit):
    from generate_payslips import run_generate_payslips
    file_id, file_name = run_generate_payslips(gc, drive_svc, sheets_svc, progress_cb)
"""
import re
import time
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ── CONFIG ────────────────────────────────────────────────────────────────────
CREDENTIALS_FILE  = "credentials.json"          # ← ใช้เฉพาะตอนรันแบบ standalone
PAYROLL_SHEET_ID  = "1oWaO8U1nNItgozvdSAcfN6QC-Ih78d0H9n5mJRfK6W4"
TEMPLATE_SHEET_ID = "1H44tl2BHMBIO-pYpxTngVKrullpfeWpuYeUZHKv8TBw"
OUTPUT_FOLDER_ID  = "16IFsEN8w7idA9T7hfcHEXVva6FUSUuYH"
TEMPLATE_TAB      = "ชีต1"
SALARY_SHEET      = "Salary"
OWNER_EMAIL       = "ai02.trc@gmail.com"        # ← เจ้าของไฟล์จริง (รับ ownership)

# Block structure (เหมือน payroll_page_updated.py)
BLOCK_SIZE      = 13
BLOCK_START_COL = 27   # คอลัมน์แรกของ block 1 (1-indexed)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── HELPERS ───────────────────────────────────────────────────────────────────
def col_letter(n: int) -> str:
    """แปลงเลขคอลัมน์ (1-indexed) เป็นตัวอักษร A, B, ..., Z, AA, ..."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result

def block_col_letter(block_n: int, col_1indexed: int) -> str:
    """ตัวอักษรคอลัมน์ของ block_n, col_1indexed (นับจาก 1 ภายใน block)"""
    abs_col = BLOCK_START_COL + (block_n - 1) * BLOCK_SIZE + (col_1indexed - 1)
    return col_letter(abs_col)

def importrange(payroll_id: str, sheet_name: str, col_str: str, row: int) -> str:
    """สร้าง IMPORTRANGE formula"""
    safe = sheet_name.replace("'", "\\'")
    return f"=IMPORTRANGE(\"{payroll_id}\", \"'{safe}'!{col_str}{row}\")"

def imp_block(payroll_id: str, sheet_name: str, block_n: int,
              col_1: int, row: int) -> str:
    """IMPORTRANGE จาก block col (1-indexed) และ row ที่กำหนด"""
    return importrange(payroll_id, sheet_name, block_col_letter(block_n, col_1), row)

def detect_latest_block(salary_ws) -> int:
    """หา block ล่าสุดจากชีต Salary"""
    latest = 1
    for bn in range(1, 20):
        col = block_col_letter(bn, 8)   # col 8 ของแต่ละ block
        try:
            val = salary_ws.acell(f"{col}2").value
        except Exception:
            break
        if val and str(val).strip():
            latest = bn
        else:
            break
    return latest

def get_month_year_from_salary(salary_ws):
    """ดึงชื่อเดือนและปีจาก B1 ของ Salary sheet"""
    b1 = salary_ws.acell("B1").value or ""
    match = re.search(r"สรุปเงินเดือน\s+(\S+)\s+(\d+)", b1)
    if match:
        return match.group(1), match.group(2)
    return "ไม่ทราบ", "2569"

# ── BUILD FORMULAS ────────────────────────────────────────────────────────────
def build_formulas(emp_name: str, block_n: int) -> dict:
    """สร้าง dict ของ {cell_address: formula} สำหรับพนักงาน 1 คน"""
    pid = PAYROLL_SHEET_ID

    def ib(col_1, row):
        return imp_block(pid, emp_name, block_n, col_1, row)

    # ── Grade/Amount refs สำหรับ IF ────────────────────────────────────────
    grade_col = block_col_letter(block_n, 5)
    amt_col   = block_col_letter(block_n, 10)
    grade_ir  = f"IMPORTRANGE(\"{pid}\", \"'{emp_name}'!{grade_col}63\")"
    amt_ir    = f"IMPORTRANGE(\"{pid}\", \"'{emp_name}'!{amt_col}63\")"

    # เงื่อนไขเกรด
    cond_abc = f"OR({grade_ir}=\"A\",{grade_ir}=\"B\",{grade_ir}=\"C\")"
    cond_de  = f"OR({grade_ir}=\"D\",{grade_ir}=\"E\")"

    return {
        # ── Header info ─────────────────────────────────────────────────────
        "C3":  ib(6, 2),      # รหัสพนักงาน
        "C4":  ib(12, 3),     # ตำแหน่ง
        "J3":  ib(7, 2),      # ชื่อสกุล
        # L2: VLOOKUP เลขบัญชีจาก Salary sheet (match รหัสใน col A)
        "L2":  (f"=IFERROR(VLOOKUP(C3,"
                f"IMPORTRANGE(\"{pid}\",\"Salary!A:Q\"),17,0),\"\")"),

        # ── รายรับ ──────────────────────────────────────────────────────────
        "H8":  ib(10, 47),    # อัตราเงินเดือน
        "H9":  ib(10, 51),    # ค่าตำแหน่ง
        "H10": "0.00",         # ค่าดูแลงาน/ประสบการณ์ — เว้นว่าง
        "H11": ib(10, 49),    # ค่าเดินทาง/ค่าสึกหรอ
        "H12": ib(10, 48),    # เบี้ยขยัน
        "H13": "0.00",         # ค่าเช่าบ้าน/ค่าอื่นๆ — เว้นว่าง
        "H14": ib(10, 64),    # ค่าเป้าสาขา

        "P8":  "=H8/30*M8",   # เงินเดือน (บาท) — M8 กรอกมือ

        "M9":  ib(10, 42),    # ค่าล่วงเวลา (นาที)
        "P9":  "=M9",          # ค่าล่วงเวลา (บาท)

        # คอมมิชชั่นจากการสอบ — แยกเพิ่ม/ลด ตาม grade
        "M10": f"=IF({cond_abc},{grade_ir},\"\")",   # เกรด A/B/C → แสดงใน income
        "P10": f"=IF({cond_abc},{amt_ir},0)",

        "P11": ib(10, 58),    # ค่าคอมฯหน้าร้าน
        "P12": ib(10, 59),    # ค่าคอมฯตจว./ติดตั้ง
        "P13": "0.00",         # ค่าคอมพิเศษ/ค่า KPI — เว้นว่าง
        "P14": "=SUM(P8:P13)+SUM(H9:H14)",  # รวมเงินได้

        # ── รายหัก ──────────────────────────────────────────────────────────
        "E17": ib(3, 40),     # หักลากิจ (นาที)
        "H17": "=H8/30/8/60*E17",

        "E18": ib(6, 40),     # ลาป่วย (นาที)
        "H18": "=H8/30/8/60*E18",

        "E19": "0.00",         # หักขาดงาน — เว้นว่าง
        "H19": "=H8/30/8/60*E19",

        "E20": "0.00",         # หักลาอื่นๆ — เว้นว่าง
        "H20": "=H8/30/8/60*E20",

        "M17": ib(3, 41),     # หักมาสาย (นาที)
        "P17": "=H8/30/8/60*M17",

        "M18": ib(3, 39),     # มาสาย (ครั้ง)
        "P18": "=IF(M18<4,0,(M18-3)*50)",

        # คอมมิชชั่นลดจากการสอบ — grade D/E → แสดงใน deduction
        "M19": f"=IF({cond_de},{grade_ir},\"\")",
        "P19": f"=IF({cond_de},{amt_ir},0)",

        "P20": "0.00",         # หักสินค้าเสียหาย — เว้นว่าง
        "P21": ib(10, 44),    # ประกันสังคม
        "P22": "0.00",         # หักภาษี — เว้นว่าง

        "H23": "=SUM(P17:P22)+SUM(H17:H22)",  # รวมหัก
        "P23": "=P14-H23-P26",                  # เงินรับสุทธิ (P26 ว่าง = กรอกมือถ้ามี)
    }

# ── CORE FUNCTION (ใช้ได้ทั้ง standalone และ Streamlit) ──────────────────────
def run_generate_payslips(gc, drive_svc, sheets_svc, progress_cb=None) -> tuple:
    """
    สร้างไฟล์ Slip เงินเดือนใหม่ใน Google Drive

    Parameters:
        gc         — gspread client (authorized)
        drive_svc  — Google Drive API service
        sheets_svc — Google Sheets API service
        progress_cb — callback(msg: str, pct: int | None) สำหรับแสดง progress

    Returns:
        (file_id: str, file_name: str)
    """
    def log(msg, pct=None):
        print(msg)
        if progress_cb:
            progress_cb(msg, pct)

    # ── เปิด Payroll Sheet ────────────────────────────────────────────────────
    log("กำลังอ่านข้อมูลจาก Payroll Sheet...", 5)
    payroll_wb  = gc.open_by_key(PAYROLL_SHEET_ID)
    salary_ws   = payroll_wb.worksheet(SALARY_SHEET)

    latest_block          = detect_latest_block(salary_ws)
    month_name, year      = get_month_year_from_salary(salary_ws)
    slip_filename         = f"Slip_{month_name}_{year}"

    log(f"Block ล่าสุด: {latest_block} | เดือน: {month_name} {year}", 10)

    # ── รายชื่อชีตพนักงาน (pattern: รหัส_ชื่อ) ────────────────────────────
    all_sheets = payroll_wb.worksheets()
    emp_sheets = [ws.title for ws in all_sheets
                  if re.match(r"^\d+_.+", ws.title)]
    log(f"พบพนักงาน {len(emp_sheets)} คน", 15)

    # ── สร้างไฟล์ว่างก่อน (ไม่ copy เพราะ Service Account มี quota = 0) ──────
    log("กำลังสร้างไฟล์ Slip ใหม่ใน Google Drive...", 20)
    new_file = drive_svc.files().create(
        body={
            "name": slip_filename,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [OUTPUT_FOLDER_ID],
        },
        fields="id",
    ).execute()
    new_file_id = new_file["id"]
    log(f"สร้างไฟล์: {slip_filename}", 23)

    # ── Transfer ownership ทันที ก่อนเพิ่ม content ใดๆ ───────────────────────
    log("กำลัง transfer ownership...", 25)
    try:
        drive_svc.permissions().create(
            fileId=new_file_id,
            body={
                "type": "user",
                "role": "owner",
                "emailAddress": OWNER_EMAIL,
            },
            transferOwnership=True,
        ).execute()
        log("Transfer ownership สำเร็จ ✅", 27)
    except Exception as e:
        log(f"⚠️ transfer ownership ไม่สำเร็จ: {e}", 27)

    # ── Copy template tab จาก Template file เข้าไฟล์ใหม่ ───────────────────
    log("กำลัง copy template sheet...", 28)
    template_wb      = gc.open_by_key(TEMPLATE_SHEET_ID)
    template_tab_src = template_wb.worksheet(TEMPLATE_TAB)
    template_gid_src = template_tab_src.id

    dup = sheets_svc.spreadsheets().sheets().copyTo(
        spreadsheetId=TEMPLATE_SHEET_ID,
        sheetId=template_gid_src,
        body={"destinationSpreadsheetId": new_file_id},
    ).execute()
    copied_gid = dup["sheetId"]

    # Rename copied tab → TEMPLATE_TAB + ลบ Sheet1 default ที่ติดมา
    slip_wb        = gc.open_by_key(new_file_id)
    default_sheets = slip_wb.worksheets()
    requests_batch = [
        {
            "updateSheetProperties": {
                "properties": {"sheetId": copied_gid, "title": TEMPLATE_TAB},
                "fields": "title",
            }
        }
    ]
    for ws in default_sheets:
        if ws.id != copied_gid:
            requests_batch.append({"deleteSheet": {"sheetId": ws.id}})
    slip_wb.batch_update({"requests": requests_batch})

    template_tab = slip_wb.worksheet(TEMPLATE_TAB)
    template_gid = template_tab.id

    # ── Generate ชีตพนักงาน ─────────────────────────────────────────────────
    total = len(emp_sheets)
    for i, emp_name in enumerate(emp_sheets):
        pct = 25 + int((i / total) * 65)
        log(f"[{i+1}/{total}] สร้างชีต {emp_name}...", pct)

        # Copy template tab
        dup = sheets_svc.spreadsheets().sheets().copyTo(
            spreadsheetId=new_file_id,
            sheetId=template_gid,
            body={"destinationSpreadsheetId": new_file_id},
        ).execute()
        new_gid = dup["sheetId"]

        # Rename
        slip_wb.batch_update({"requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": new_gid, "title": emp_name},
                "fields": "title",
            }
        }]})

        # Fill formulas
        emp_ws   = slip_wb.worksheet(emp_name)
        formulas = build_formulas(emp_name, latest_block)
        data     = [{"range": cell, "values": [[val]]}
                    for cell, val in formulas.items()]
        emp_ws.batch_update(data, value_input_option="USER_ENTERED")

        time.sleep(1.2)   # avoid Google API rate limit

    # ── ลบ Template Tab ออก ─────────────────────────────────────────────────
    log("กำลังลบ template tab...", 92)
    slip_wb.batch_update({"requests": [{
        "deleteSheet": {"sheetId": template_gid}
    }]})

    log(f"เสร็จสิ้น! ไฟล์: {slip_filename}", 100)
    return new_file_id, slip_filename


# ── MAIN (standalone) ─────────────────────────────────────────────────────────
def main():
    print("── เชื่อม Google APIs ──")
    creds      = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc         = gspread.authorize(creds)
    drive_svc  = build("drive",  "v3", credentials=creds)
    sheets_svc = build("sheets", "v4", credentials=creds)

    file_id, file_name = run_generate_payslips(gc, drive_svc, sheets_svc)

    print(f"\n✅ เสร็จสิ้น!")
    print(f"   ไฟล์: https://docs.google.com/spreadsheets/d/{file_id}")
    print(f"   ⚠️  เปิดไฟล์แล้วกด 'Allow access' สำหรับ IMPORTRANGE ครั้งแรก")
    print(f"   📝  M8 (เงินเดือน วัน) ต้องกรอกมือทุกชีต")


if __name__ == "__main__":
    main()
