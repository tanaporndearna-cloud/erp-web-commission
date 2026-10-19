#!/usr/bin/env python3
"""
generate_payslips.py — สร้างสลีปเงินเดือนรายเดือน TRC Motorsport
เปิดไฟล์ Google Sheets ที่ผู้ใช้สร้างไว้แล้ว และ generate ชีตพนักงานทุกคน

Usage (standalone):
    python3 generate_payslips.py <target_file_id>

Usage (from Streamlit):
    from generate_payslips import run_generate_payslips
    file_name = run_generate_payslips(gc, sheets_svc, target_file_id, progress_cb)
"""
import re
import sys
import time
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ── RETRY HELPER ──────────────────────────────────────────────────────────────
def api_call_with_retry(fn, max_retries=5):
    """เรียก API พร้อม retry อัตโนมัติเมื่อเจอ 429 Rate Limit"""
    for attempt in range(max_retries):
        try:
            return fn()
        except (HttpError, gspread.exceptions.APIError) as e:
            err_str = str(e)
            if "429" in err_str or "RATE_LIMIT" in err_str or "Quota exceeded" in err_str:
                wait = 15 * (attempt + 1)  # 15s, 30s, 45s, 60s, 75s
                print(f"  ⏳ Rate limit — รอ {wait}s แล้วลองใหม่ (ครั้งที่ {attempt+1}/{max_retries})")
                time.sleep(wait)
            else:
                raise
    raise RuntimeError("เกิน retry limit สำหรับ API call")

# ── CONFIG ────────────────────────────────────────────────────────────────────
CREDENTIALS_FILE  = "credentials.json"          # ← ใช้เฉพาะตอนรันแบบ standalone
PAYROLL_SHEET_ID  = "1oWaO8U1nNItgozvdSAcfN6QC-Ih78d0H9n5mJRfK6W4"
TEMPLATE_SHEET_ID = "1nr0100FNOa9UA-TpXfEvgA1uaoscD41C8sbl9HbHz_l"
TEMPLATE_TAB      = "Template"
SALARY_SHEET      = "Salary"

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

_THAI_MONTHS = {
    "มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3,
    "เมษายน": 4, "พฤษภาคม": 5, "มิถุนายน": 6,
    "กรกฎาคม": 7, "สิงหาคม": 8, "กันยายน": 9,
    "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12,
}

def _parse_month_from_cell(val: str) -> int | None:
    """แปลงค่าเดือน/ปีใน cell (เช่น 'ตุลาคม 2569' หรือ '10/2569') เป็นเลขเดือน"""
    if not val:
        return None
    val = str(val).strip()
    # รูปแบบ เลข/ปี เช่น "10/2569" หรือ "10-2569"
    m = re.match(r"^(\d{1,2})[/\-]", val)
    if m:
        return int(m.group(1))
    # รูปแบบชื่อเดือนไทย เช่น "ตุลาคม 2569"
    for name, num in _THAI_MONTHS.items():
        if name in val:
            return num
    return None

def detect_latest_block(salary_ws) -> int:
    """หา block ปัจจุบัน = block สุดท้ายที่ col 12 แถว 1 มีค่า (Bookmark เดือน/ปี)

    ตรรกะ: เดือนปัจจุบันอยู่ใน block ขวาสุดที่มีข้อมูลเสมอ
    """
    latest = 1
    empty_streak = 0   # นับ block ว่างต่อเนื่อง
    for bn in range(1, 20):
        col = block_col_letter(bn, 12)
        try:
            val = salary_ws.acell(f"{col}1").value
        except Exception:
            break
        if val and str(val).strip():
            latest = bn
            empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= 2:   # ว่าง 2 block ติดกัน = ถึงจุดสิ้นสุดแน่นอน
                break
    return latest

def find_block_by_month(salary_ws, target_month: int) -> int:
    """หา block ที่ตรงกับเดือนที่เลือก โดยอ่านเดือน/ปีจาก col 12 แถว 1"""
    for bn in range(1, 20):
        col = block_col_letter(bn, 12)
        try:
            val = salary_ws.acell(f"{col}1").value
        except Exception:
            break
        parsed = _parse_month_from_cell(val)
        if parsed == target_month:
            return bn
        if not val or not str(val).strip():
            break
    # ถ้าหาไม่เจอ fallback เป็น detect_latest_block
    return detect_latest_block(salary_ws)

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

    def ib_num(col_1, row):
        """IMPORTRANGE ที่คืน 0 ถ้าค่าไม่ใช่ตัวเลข (เช่น "-" หรือว่าง)"""
        raw = imp_block(pid, emp_name, block_n, col_1, row)[1:]  # ตัด "=" ออก
        return f"=IFERROR(VALUE({raw}),0)"

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
        "L3":  ib(7, 2),      # ชื่อสกุล
        # L2: VLOOKUP เลขบัญชีจาก Salary sheet (match รหัสใน col A)
        "L2":  (f"=IFERROR(VLOOKUP(C3,"
                f"IMPORTRANGE(\"{pid}\",\"Salary!A:Q\"),17,0),\"\")"),

        # ── รายรับ ──────────────────────────────────────────────────────────
        "H8":  ib_num(10, 47),    # อัตราเงินเดือน
        "H9":  ib_num(10, 51),    # ค่าตำแหน่ง
        "H10": "0.00",             # ค่าดูแลงาน/ประสบการณ์ — เว้นว่าง
        "H11": ib_num(10, 49),    # ค่าเดินทาง/ค่าสึกหรอ
        "H12": ib_num(10, 48),    # เบี้ยขยัน
        "H13": "0.00",             # ค่าเช่าบ้าน/ค่าอื่นๆ — เว้นว่าง
        "H14": ib_num(10, 64),    # ค่าเป้าสาขา

        "P8":  "=H8/30*M8",   # เงินเดือน (บาท) — M8 กรอกมือ

        "M9":  ib_num(10, 42),    # ค่าล่วงเวลา (นาที)
        "P9":  "=M9",          # ค่าล่วงเวลา (บาท)

        # คอมมิชชั่นจากการสอบ — แยกเพิ่ม/ลด ตาม grade
        "M10": f"=IF({cond_abc},{grade_ir},\"\")",   # เกรด A/B/C → แสดงใน income
        "P10": f"=IF({cond_abc},{amt_ir},0)",

        "P11": ib_num(10, 58),    # ค่าคอมฯหน้าร้าน
        "P12": ib_num(10, 59),    # ค่าคอมฯตจว./ติดตั้ง
        "P13": "0.00",         # ค่าคอมพิเศษ/ค่า KPI — เว้นว่าง
        "P14": "=SUM(P8:P13)+SUM(H9:H14)",  # รวมเงินได้

        # ── รายหัก ──────────────────────────────────────────────────────────
        "E17": ib_num(3, 40),     # หักลากิจ (นาที)
        "H17": "=H8/30/8/60*E17",

        "E18": ib_num(6, 40),     # ลาป่วย (นาที)
        "H18": "=H8/30/8/60*E18",

        "E19": "0.00",         # หักขาดงาน — เว้นว่าง
        "H19": "=H8/30/8/60*E19",

        "E20": "0.00",         # หักลาอื่นๆ — เว้นว่าง
        "H20": "=H8/30/8/60*E20",

        "M17": ib_num(3, 41),     # หักมาสาย (นาที)
        "P17": "=H8/30/8/60*M17",

        "M18": ib_num(3, 39),     # มาสาย (ครั้ง)
        "P18": "=IF(M18<4,0,(M18-3)*50)",

        # คอมมิชชั่นลดจากการสอบ — grade D/E → แสดงใน deduction
        "M19": f"=IF({cond_de},{grade_ir},\"\")",
        "P19": f"=IF({cond_de},{amt_ir},0)",

        "P20": "0.00",         # หักสินค้าเสียหาย — เว้นว่าง
        "P21": "=" + ib_num(10, 44)[1:] + "*-1",    # ประกันสังคม (×-1 ให้เป็นค่าหัก)
        "P22": "0.00",         # หักภาษี — เว้นว่าง

        "H23": "=SUM(P17:P22)+SUM(H17:H22)",  # รวมหัก
        "P23": "=P14-H23-P26",                  # เงินรับสุทธิ (P26 ว่าง = กรอกมือถ้ามี)
    }

# ── CORE FUNCTION (ใช้ได้ทั้ง standalone และ Streamlit) ──────────────────────
def run_generate_payslips(gc, sheets_svc, target_file_id,
                          template_sheet_id=None, month=None, progress_cb=None) -> tuple:
    """
    เติมชีตพนักงานลงในไฟล์ Google Sheets ที่ผู้ใช้สร้างไว้แล้ว

    Parameters:
        gc                 — gspread client (authorized)
        sheets_svc         — Google Sheets API service
        target_file_id     — ID ของไฟล์ Slip ที่ผู้ใช้สร้างเอง
        template_sheet_id  — ID ของไฟล์ Template (กรอกในหน้าเว็บ); ถ้าไม่ส่งใช้ค่า TEMPLATE_SHEET_ID
        month              — เดือนที่เลือกใน sidebar (int 1-12); ถ้า None จะ auto-detect
        progress_cb        — callback(msg: str, pct: int | None) สำหรับแสดง progress

    Returns:
        (file_id: str, file_name: str)
    """
    def log(msg, pct=None):
        print(msg)
        if progress_cb:
            progress_cb(msg, pct)

    # ถ้าไม่ได้ส่ง template_sheet_id มา ใช้ target_file_id แทน (ไฟล์เดียวกัน)
    # fallback สุดท้ายถึงจะใช้ TEMPLATE_SHEET_ID (hardcode)
    tmpl_id = (template_sheet_id or "").strip() or target_file_id or TEMPLATE_SHEET_ID

    # ── เปิด Payroll Sheet ────────────────────────────────────────────────────
    log("กำลังอ่านข้อมูลจาก Payroll Sheet...", 5)
    payroll_wb  = gc.open_by_key(PAYROLL_SHEET_ID)
    salary_ws   = payroll_wb.worksheet(SALARY_SHEET)

    # ── รายชื่อชีตพนักงาน (pattern: รหัส_ชื่อ) ────────────────────────────
    all_sheets = payroll_wb.worksheets()
    emp_sheets = [ws.title for ws in all_sheets
                  if re.match(r"^\d+_.+", ws.title)]
    log(f"พบพนักงาน {len(emp_sheets)} คน", 15)

    # ── หา block ปัจจุบันจากชีตพนักงานคนแรก (Salary มีแค่ 26 col ไม่มี block) ──
    if not emp_sheets:
        raise ValueError("ไม่พบชีตพนักงาน (รูปแบบ รหัส_ชื่อ) ใน Payroll Sheet")
    first_emp_ws = payroll_wb.worksheet(emp_sheets[0])
    latest_block = detect_latest_block(first_emp_ws)
    log(f"Block ปัจจุบัน (block สุดท้ายที่มีข้อมูล): {latest_block}", 8)

    # ดึงชื่อเดือน/ปีจาก col 12 แถว 1 ของ block ที่เลือก
    bm_col = block_col_letter(latest_block, 12)
    bm_val = first_emp_ws.acell(f"{bm_col}1").value or ""
    if not bm_val.strip():
        month_name, year = get_month_year_from_salary(salary_ws)
    else:
        m = re.match(r"^(\d{1,2})[/\-](\d+)", bm_val.strip())
        if m:
            month_num  = int(m.group(1))
            year       = m.group(2)
            month_name = next((k for k, v in _THAI_MONTHS.items() if v == month_num), str(month_num))
        else:
            parts      = bm_val.strip().split()
            month_name = parts[0] if parts else "ไม่ทราบ"
            year       = parts[1] if len(parts) > 1 else "2569"

    slip_filename = f"Slip_{month_name}_{year}"
    log(f"Block: {latest_block} | {month_name} {year}", 10)

    # ── เปิดไฟล์ที่ผู้ใช้สร้างไว้แล้ว ────────────────────────────────────────
    log("กำลังเปิดไฟล์ Slip ที่สร้างไว้...", 20)
    try:
        slip_wb = gc.open_by_key(target_file_id)
    except Exception as e:
        raise ValueError(
            f"ไม่สามารถเปิดไฟล์ได้ (ID: {target_file_id})\n"
            f"โปรดตรวจสอบ: 1) File ID ถูกต้อง  2) Share ไฟล์ให้ Service Account แล้ว\n"
            f"Error: {e}"
        )
    log(f"เปิดไฟล์ '{slip_wb.title}' สำเร็จ ✅", 23)

    # ── เตรียม template tab ────────────────────────────────────────────────────
    # ดึง template tab จาก Template file แล้ว copy เข้าไฟล์ target
    log(f"กำลัง copy template sheet (tmpl_id={tmpl_id[:20]}...)...", 25)

    # ถ้า template == target ใช้ slip_wb ที่เปิดไว้แล้ว (ป้องกัน 404 redundant)
    if tmpl_id.strip() == target_file_id.strip():
        template_wb = slip_wb
        log("Template อยู่ไฟล์เดียวกับ Slip — ใช้ไฟล์ที่เปิดอยู่", 26)
    else:
        try:
            template_wb = gc.open_by_key(tmpl_id)
        except Exception as e:
            raise ValueError(
                f"ไม่สามารถเปิดไฟล์ Template ได้ (ID: {tmpl_id})\n"
                f"โปรดตรวจสอบ: 1) File ID ถูกต้อง  2) Share ไฟล์ให้ erp-streamlit@erp-commission.iam.gserviceaccount.com\n"
                f"Error: {e}"
            )

    template_tab_src = template_wb.worksheet(TEMPLATE_TAB)
    template_gid_src = template_tab_src.id

    dup = sheets_svc.spreadsheets().sheets().copyTo(
        spreadsheetId=tmpl_id,
        sheetId=template_gid_src,
        body={"destinationSpreadsheetId": target_file_id},
    ).execute()
    copied_gid = dup["sheetId"]

    # Rename copied tab → TEMPLATE_TAB + ลบชีตอื่นที่ไม่ใช่ template
    slip_wb        = gc.open_by_key(target_file_id)   # refresh
    existing_sheets = slip_wb.worksheets()

    rename_req = [{
        "updateSheetProperties": {
            "properties": {"sheetId": copied_gid, "title": TEMPLATE_TAB},
            "fields": "title",
        }
    }]
    # ลบทุกชีตที่ไม่ใช่ template ที่เพิ่ง copy มา (ชีตเดิมที่ผู้ใช้สร้างมาด้วย)
    delete_reqs = [
        {"deleteSheet": {"sheetId": ws.id}}
        for ws in existing_sheets
        if ws.id != copied_gid
    ]
    # ลบชีตเก่าก่อน แล้วค่อย rename — ป้องกัน "sheet name already exists"
    slip_wb.batch_update({"requests": delete_reqs + rename_req})
    log("เตรียม template สำเร็จ ✅", 28)

    # refresh อีกครั้งหลัง delete
    slip_wb     = gc.open_by_key(target_file_id)
    template_tab = slip_wb.worksheet(TEMPLATE_TAB)
    template_gid = template_tab.id

    # ── Generate ชีตพนักงาน ─────────────────────────────────────────────────
    total = len(emp_sheets)
    for i, emp_name in enumerate(emp_sheets):
        pct = 30 + int((i / total) * 60)
        log(f"[{i+1}/{total}] สร้างชีต {emp_name}...", pct)

        # Copy template tab (retry on rate limit)
        dup = api_call_with_retry(lambda: sheets_svc.spreadsheets().sheets().copyTo(
            spreadsheetId=target_file_id,
            sheetId=template_gid,
            body={"destinationSpreadsheetId": target_file_id},
        ).execute())
        new_gid = dup["sheetId"]

        # Rename (retry on rate limit)
        api_call_with_retry(lambda: slip_wb.batch_update({"requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": new_gid, "title": emp_name},
                "fields": "title",
            }
        }]}))

        # Fill formulas (retry on rate limit)
        emp_ws   = slip_wb.worksheet(emp_name)
        formulas = build_formulas(emp_name, latest_block)
        data     = [{"range": cell, "values": [[val]]}
                    for cell, val in formulas.items()]
        api_call_with_retry(lambda: emp_ws.batch_update(data, value_input_option="USER_ENTERED"))

        time.sleep(2.5)   # เพิ่ม delay ป้องกัน rate limit

    # ── ลบ Template Tab ออก ─────────────────────────────────────────────────
    log("กำลังลบ template tab...", 92)
    slip_wb.batch_update({"requests": [{
        "deleteSheet": {"sheetId": template_gid}
    }]})

    log(f"เสร็จสิ้น! ไฟล์: {slip_wb.title}", 100)
    return target_file_id, slip_wb.title


# ── MAIN (standalone) ─────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("Usage: python3 generate_payslips.py <target_file_id>")
        print("  target_file_id: ID ของไฟล์ Google Sheets ที่ copy จาก template ไว้แล้ว")
        sys.exit(1)

    target_file_id = sys.argv[1]

    print("── เชื่อม Google APIs ──")
    creds      = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc         = gspread.authorize(creds)
    sheets_svc = build("sheets", "v4", credentials=creds)

    file_id, file_name = run_generate_payslips(gc, sheets_svc, target_file_id)

    print(f"\n✅ เสร็จสิ้น!")
    print(f"   ไฟล์: https://docs.google.com/spreadsheets/d/{file_id}")
    print(f"   ⚠️  เปิดไฟล์แล้วกด 'Allow access' สำหรับ IMPORTRANGE ครั้งแรก")
    print(f"   📝  M8 (เงินเดือน วัน) ต้องกรอกมือทุกชีต")


if __name__ == "__main__":
    main()
