"""
google_writer.py — เขียนข้อมูลค่าคอมลง Google Sheet
ใช้ gspread + Service Account ที่เก็บใน Streamlit secrets
"""
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SHEET_ID = "1CkK7KZS0J_GzSAB46Xxh4dP_BVKRdx1GtbUHM9ZwikQ"
SHEET_COM_ERP = "Com ERP"
SHEET_SUMMARIZE = "สรุปCom"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_client():
    """สร้าง gspread client จาก Service Account ใน secrets"""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def write_com_erp(rows: list, dates: list, progress_cb=None) -> str:
    """
    เขียนข้อมูลลงชีท Com ERP
    rows: list of dict {branch, name, daily:[32], com_pp, com_tot, sales, pct_com}
    dates: list of 32 strings เช่น ["20/07/69", ...]
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_COM_ERP)

    all_values = ws.get_all_values()

    # หา header row (col A = "สาขา")
    header_row_idx = None
    for i, row in enumerate(all_values):
        if row and row[0] == "สาขา":
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("หา header row (สาขา) ใน Com ERP ไม่เจอ")

    header_row = header_row_idx + 1  # 1-indexed
    data_start_row = header_row + 1
    date_start_col = 3   # col C
    summary_start = 36   # col AJ

    # เขียนวันที่ใน header
    date_updates = []
    for i, d in enumerate(dates):
        date_updates.append({
            "range": gspread.utils.rowcol_to_a1(header_row, date_start_col + i),
            "values": [[d]]
        })
    ws.batch_update(date_updates)

    # เขียนข้อมูลทีละแถว (batch)
    batch_data = []
    for ri, d in enumerate(rows):
        row = data_start_row + ri
        row_vals = [d["branch"], d["name"]]
        # daily 32 วัน
        for di in range(32):
            val = d["daily"][di] if d["daily"] and di < len(d["daily"]) else 0
            row_vals.append(round(float(val or 0), 2))
        # col AJ (idx 34) = empty, then summary at 36
        while len(row_vals) < 34:
            row_vals.append("")
        row_vals.append("")  # col AI = empty
        row_vals.append(round(float(d.get("com_pp") or 0), 2))  # AJ
        row_vals.append(round(float(d.get("com_tot") or 0), 2))  # AK
        row_vals.append(round(float(d.get("sales") or 0), 2))    # AL
        # % Com — ส่งเป็น fraction แล้วตั้ง format ทีหลัง
        row_vals.append(round(float(d.get("pct_com") or 0), 6))  # AM

        batch_data.append({
            "range": f"A{row}:{gspread.utils.rowcol_to_a1(row, len(row_vals))}",
            "values": [row_vals]
        })

        if progress_cb:
            progress_cb(ri + 1, len(rows))

    ws.batch_update(batch_data)

    # ตั้ง number format % Com
    pct_range = f"{gspread.utils.rowcol_to_a1(data_start_row, summary_start + 3)}:{gspread.utils.rowcol_to_a1(data_start_row + len(rows) - 1, summary_start + 3)}"
    ws.format(pct_range, {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}})

    return f"✅ เขียน Com ERP สำเร็จ {len(rows)} แถว"


def _serial_to_date_str(val: str) -> str:
    """
    แปลง Google Sheets / Excel serial number (เช่น '46031' หรือ '46,031.00')
    ให้เป็น string 'dd/mm/yy' — ถ้าแปลงไม่ได้คืนค่าเดิม
    """
    from datetime import datetime as _dt, timedelta as _td
    try:
        clean = str(val).replace(",", "").strip()
        n = float(clean)
        if 40000 < n < 60000:          # ช่วงปี ~2009-2064
            return (_dt(1899, 12, 30) + _td(days=int(n))).strftime("%d/%m/%y")
    except (ValueError, TypeError):
        pass
    return str(val)


def _be_date_to_ddmmyy(date_str: str) -> str:
    """
    แปลงวันที่ปีพุทธศักราช (dd/mm/YYYY เช่น '03/10/2569')
    ให้เป็น 'dd/mm/yy' ปีคริสต์ (เช่น '03/10/26')
    ถ้าเป็นปีคริสต์อยู่แล้ว (year < 2300) คืนค่าเดิม
    """
    try:
        parts = str(date_str).strip().split("/")
        if len(parts) == 3:
            day, month, year = parts
            year_int = int(year)
            if year_int > 2300:          # ปีพุทธศักราช
                year_int -= 543
            return f"{day}/{month}/{str(year_int)[2:]}"
    except (ValueError, IndexError):
        pass
    return str(date_str)


def parse_absent_days(xlsx_path: str) -> dict:
    """
    อ่านไฟล์ InOutDailyReport.xlsx
    คืน dict: {(emp_id_str, "dd/mm/yy"): True}
    สำหรับพนักงานที่ไม่มีเวลาเข้างาน AND ไม่มีเวลาออกงาน (= ขาด/ลา/หยุด)

    โครงสร้างไฟล์:
    - Col B (idx 2): รหัสพนักงาน
    - Col C (idx 3): ชื่อ-นามสกุล
    - Col D (idx 4): แผนก  (เช่น TRC2, TRC4)
    - Col F (idx 6): วันที่  dd/mm/YYYY ปีพุทธศักราช
    - Col G (idx 7): เวลาเข้างาน  — None ถ้าไม่มา
    - Col H (idx 8): เวลาออกงาน  — None ถ้าไม่มา
    """
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    absent: dict = {}
    for r in range(2, ws.max_row + 1):
        emp_id   = ws.cell(r, 2).value   # B
        date_val = ws.cell(r, 6).value   # F
        time_in  = ws.cell(r, 7).value   # G
        time_out = ws.cell(r, 8).value   # H

        if not emp_id or not date_val:
            continue

        emp_id_str = str(emp_id).strip()
        date_ddmmyy = _be_date_to_ddmmyy(str(date_val).strip())

        def _empty(v) -> bool:
            return v is None or str(v).strip() in ("", "None")

        if _empty(time_in) and _empty(time_out):
            absent[(emp_id_str, date_ddmmyy)] = True

    return absent


def _is_sunday(date_str: str) -> bool:
    """คืน True ถ้าวันที่นั้นเป็นวันอาทิตย์
    รองรับ dd/mm/yy, dd/mm/yyyy และ Google Sheets serial number"""
    from datetime import datetime
    if not date_str:
        return False
    normalized = _serial_to_date_str(date_str)
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(normalized.strip(), fmt).weekday() == 6
        except ValueError:
            continue
    return False


def write_summarize_com(branch_totals: dict, dates: list = None, date_start_day: int = 20) -> str:
    """
    เขียนข้อมูลลงชีท สรุปCom แบบ positional — เริ่มจาก col E เสมอ
    branch_totals: {"T2": [v1..v32], "T3": [...], ...}
    vals[0] = วันที่ 20 (col E), vals[1] = วันที่ 21 (col F), ไปเรื่อยๆ
    ถ้า dates ส่งมาด้วย จะเขียนวันที่ลง date header row (row 5) ทันที
    หมายเหตุ: ไม่ใช้ day-matching เพราะวันที่ 20 ปรากฏสองครั้งใน header
    (ต้นรอบ col E + ปลายรอบ col AJ) ทำให้ mapping ผิดพลาด
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_SUMMARIZE)

    all_values = ws.get_all_values()

    # Date header อยู่ที่ row 5 เสมอ (0-indexed = 4) — เหมือน mark_absent_h
    DATE_HEADER_IDX = 4  # 0-indexed
    date_header_row_idx = DATE_HEADER_IDX

    date_col_start = 5  # col E (1-indexed) = วันที่ 20 เสมอ

    # นับจากวันที่จริงที่ส่งมา ไม่ hardcode 32
    # กรอง empty ออกก่อน แล้วใช้ความยาวจริง (28/29/30/31 วัน)
    if dates:
        clean_dates = [str(d) for d in dates if str(d).strip()]
    else:
        clean_dates = []
    num_day_cols = len(clean_dates) if clean_dates else 32

    sheet_last_date_col = date_col_start + num_day_cols - 1

    # ── เขียนวันที่ลง date header row ──
    if clean_dates:
        date_header_1idx = date_header_row_idx + 1  # 1-indexed
        start_cell = gspread.utils.rowcol_to_a1(date_header_1idx, date_col_start)
        end_cell   = gspread.utils.rowcol_to_a1(date_header_1idx, date_col_start + num_day_cols - 1)
        ws.update(f"{start_cell}:{end_cell}", [clean_dates], value_input_option="USER_ENTERED")
    dates = clean_dates  # ใช้ตัวที่ clean แล้วตลอด

    # หา Commission TRC rows
    trc_row_map = {}
    for i in range(date_header_row_idx + 1, len(all_values)):
        row = all_values[i]
        if len(row) >= 3 and row[1] == "Commission" and str(row[2]).startswith("TRC "):
            num = str(row[2]).replace("TRC ", "").strip()
            trc_row_map[f"T{num}"] = i + 1

    if not trc_row_map:
        raise ValueError("หา Commission TRC rows ใน สรุปCom ไม่เจอ")

    batch_data = []
    written = 0

    for branch, vals in branch_totals.items():
        if branch not in trc_row_map:
            continue
        row_idx = trc_row_map[branch]

        # เขียนแบบ positional: vals[0] → col E, vals[1] → col F, ...
        # แถว Commission TRC — ใส่ค่าจริงเสมอ (ไม่ใส่ H แม้วันอาทิตย์)
        # H จะถูกใส่เฉพาะแถวรหัสพนักงานโดย mark_absent_h() เท่านั้น
        row_vals = []
        for i, v in enumerate(vals[:num_day_cols]):
            row_vals.append(round(float(v or 0), 2))
        start_cell = gspread.utils.rowcol_to_a1(row_idx, date_col_start)
        end_cell = gspread.utils.rowcol_to_a1(row_idx, date_col_start + len(row_vals) - 1)
        batch_data.append({"range": f"{start_cell}:{end_cell}", "values": [row_vals]})

        written += 1

    ws.batch_update(batch_data)

    # ตั้ง number format #,##0.00 เฉพาะช่องที่ไม่ใช่วันอาทิตย์
    sunday_cols = set()
    if dates:
        for i, d in enumerate(dates[:num_day_cols]):
            if _is_sunday(d):
                sunday_cols.add(date_col_start + i)  # 1-indexed col

    fmt_requests = []
    for branch in branch_totals:
        if branch not in trc_row_map:
            continue
        row_idx = trc_row_map[branch]
        # format เป็น range ย่อยๆ ข้ามคอลัมน์ที่เป็น Sunday
        range_start = None
        for ci in range(date_col_start, date_col_start + num_day_cols):
            if ci in sunday_cols:
                # ปิด range ที่เปิดอยู่
                if range_start is not None:
                    fmt_requests.append({
                        "range": f"{gspread.utils.rowcol_to_a1(row_idx, range_start)}:{gspread.utils.rowcol_to_a1(row_idx, ci - 1)}",
                        "format": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}
                    })
                    range_start = None
            else:
                if range_start is None:
                    range_start = ci
        # ปิด range สุดท้าย
        if range_start is not None:
            fmt_requests.append({
                "range": f"{gspread.utils.rowcol_to_a1(row_idx, range_start)}:{gspread.utils.rowcol_to_a1(row_idx, date_col_start + num_day_cols - 1)}",
                "format": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}
            })

    if fmt_requests:
        ws.batch_format(fmt_requests)

    return f"✅ เขียน สรุปCom สำเร็จ {written}/{len(trc_row_map)} สาขา"


SHEET_SUM_O2O = "sum (ตัดO2O)"

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

# ชีทที่ต้องอัปเดตหัวตามเดือน (A1)
MONTHLY_TITLE_SHEETS = [
    "สรุปCom",
    "สรุปComช่าง",
    "ยืมระหว่างสาขาโซนใกล้",
    "บิลกทม",
    "เปลี่ยนสินค้า",
]


def update_monthly_titles(month: int, year_be: int) -> str:
    """
    อัปเดตหัวชีท A1 ของ 5 ชีท โดยอ่านค่าจาก A1 ของชีท Com ERP
    ค้นหาชีทแบบ case-insensitive
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)

    all_worksheets = sh.worksheets()
    all_sheet_names = {ws.title.lower(): ws for ws in all_worksheets}

    # อ่าน A1 จากชีท Com ERP
    ws_com_erp = all_sheet_names.get(SHEET_COM_ERP.lower())
    if ws_com_erp is None:
        return f"⚠️ หาชีท '{SHEET_COM_ERP}' ไม่เจอ"
    title = ws_com_erp.acell("A1").value or ""

    updated = []

    for sheet_name in MONTHLY_TITLE_SHEETS:
        ws = all_sheet_names.get(sheet_name.lower())
        if ws is None:
            continue
        ws.update("A1", [[title]], value_input_option="USER_ENTERED")
        updated.append(ws.title)

    return f"✅ อัปเดตหัวชีท {len(updated)} ชีท: {', '.join(updated)}"


def _make_com_title(month: int, year_be: int) -> str:
    """สร้างชื่อหัวชีท เช่น 'คอมมิชชั่นหน้าร้าน เดือนกันยายน 2569 (21 สิงหาคม 69 - 20 กันยายน 69)'"""
    month_name = THAI_MONTHS[month] if 1 <= month <= 12 else f"เดือน{month}"
    prev_month = 12 if month == 1 else month - 1
    prev_month_name = THAI_MONTHS[prev_month]
    year_short = year_be % 100  # เช่น 2569 → 69
    return (
        f"คอมมิชชั่นหน้าร้าน เดือน{month_name} {year_be} "
        f"(21 {prev_month_name} {year_short} - 20 {month_name} {year_short})"
    )


def copy_sum_o2o_to_com_erp(xlsx_path: str, progress_cb=None, month: int = None, year_be: int = None) -> str:
    """
    อ่านข้อมูลทั้งหมดจากชีท sum(ตัดO2O) ในไฟล์ xlsx
    แล้วเขียนลง Google Sheet ชีท Com ERP เป็นค่าธรรมดา (ไม่มีสูตร)
    หน้าตาเหมือน sum(ตัดO2O) เป๊ะๆ
    """
    import openpyxl

    if progress_cb:
        progress_cb("อ่านข้อมูลจาก sum(ตัดO2O) ในไฟล์ xlsx...")

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # หาชีท sum(ตัดO2O)
    ws_src = None
    for name in wb.sheetnames:
        if "sum" in name.lower() and "o2o" in name.lower():
            ws_src = wb[name]
            break

    if ws_src is None:
        raise ValueError(f"ไม่พบชีท sum(ตัดO2O) ในไฟล์ xlsx\nชีทที่มี: {wb.sheetnames}")

    # อ่านทุก row เป็น plain values (ตรวจจับ date cell ด้วย)
    from datetime import datetime as _dt, timedelta as _td

    def _excel_serial_to_str(val):
        """แปลง Excel date serial → YYYY-MM-DD (ISO) ให้ Google Sheets อ่านถูก locale"""
        try:
            return (_dt(1899, 12, 30) + _td(days=int(val))).strftime("%Y-%m-%d")
        except Exception:
            return val

    all_rows = []
    cell_formats = []  # เก็บ (row_1idx, col_1idx, bg_hex, is_bold) สำหรับ cell ที่มีสี/ตัวหนา

    for ri, row in enumerate(ws_src.iter_rows(values_only=False)):
        processed = []
        for ci, cell in enumerate(row):
            val = cell.value
            if val is None:
                processed.append("")
            elif hasattr(val, 'strftime'):
                processed.append(val.strftime("%Y-%m-%d"))  # ISO format — Google Sheets อ่านถูกทุก locale
            elif isinstance(val, (int, float)) and (cell.is_date or 40000 < val < 60000):
                processed.append(_excel_serial_to_str(val))
            else:
                processed.append(val)

            # อ่านสีพื้นหลัง
            bg_hex = None
            try:
                fill = cell.fill
                if fill and fill.patternType not in (None, 'none'):
                    fg = fill.fgColor
                    if fg and fg.type == 'rgb':
                        argb = fg.rgb  # AARRGGBB
                        if argb and argb not in ('00000000', 'FFFFFFFF', 'FF000000'):
                            bg_hex = argb[2:]  # เอาแค่ RRGGBB
            except Exception:
                pass

            # อ่าน bold
            is_bold = False
            try:
                if cell.font and cell.font.bold:
                    is_bold = True
            except Exception:
                pass

            if bg_hex or is_bold:
                cell_formats.append((ri + 1, ci + 1, bg_hex, is_bold))

        all_rows.append(processed)

    wb.close()

    if not all_rows:
        raise ValueError("ชีท sum(ตัดO2O) ว่างเปล่า")

    if progress_cb:
        progress_cb("เชื่อมต่อ Google Sheet...")

    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws_dst = sh.worksheet(SHEET_COM_ERP)

    if progress_cb:
        progress_cb("เขียนข้อมูลลง Com ERP...")

    max_cols = max(len(r) for r in all_rows)
    normalized = [list(r) + [""] * (max_cols - len(r)) for r in all_rows]
    total_rows = len(normalized)

    write_range = f"A1:{gspread.utils.rowcol_to_a1(total_rows, max_cols)}"
    ws_dst.update(write_range, normalized, value_input_option="USER_ENTERED")

    full_range = f"A1:{gspread.utils.rowcol_to_a1(total_rows, max_cols)}"

    # ===== Apply สีพื้นหลัง + ตัวหนา ก่อน =====
    if cell_formats:
        if progress_cb:
            progress_cb(f"ใส่สีและตัวหนา ({len(cell_formats)} cells)...")

        fmt_requests = []
        for (r, c, bg_hex, is_bold) in cell_formats:
            cell_a1 = gspread.utils.rowcol_to_a1(r, c)
            fmt = {}
            if bg_hex:
                rr = int(bg_hex[0:2], 16) / 255
                gg = int(bg_hex[2:4], 16) / 255
                bb = int(bg_hex[4:6], 16) / 255
                fmt["backgroundColor"] = {"red": rr, "green": gg, "blue": bb}
            if is_bold:
                fmt["textFormat"] = {"bold": True}
            if fmt:
                fmt_requests.append({"range": cell_a1, "format": fmt})

        if fmt_requests:
            ws_dst.batch_format(fmt_requests)

    # ===== รีเซ็ต number format ทีหลัง (ทับสีได้ แต่ไม่ทับสี) =====
    if progress_cb:
        progress_cb("รีเซ็ตรูปแบบตัวเลข...")

    ws_dst.format(
        full_range,
        {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}
    )

    # ===== แถวที่ 3 = วันที่ → ใส่ DATE format ทับ (เพราะ USER_ENTERED แปลง "13/09/26" เป็น serial แล้ว #,##0.00 ทำให้โชว์เป็นตัวเลข) =====
    if progress_cb:
        progress_cb("ใส่ format วันที่แถว 3...")
    date_row_range = f"A3:{gspread.utils.rowcol_to_a1(3, max_cols)}"
    ws_dst.format(date_row_range, {"numberFormat": {"type": "DATE", "pattern": "d/m/yy"}})

    # ล้าง B1 ที่อาจมีชื่อเดือนเก่าจาก template (เช่น ตุลาคม ทั้งๆ ที่เดือนจริงคือ กันยายน)
    if progress_cb:
        progress_cb("ล้าง B1 ที่มีชื่อเดือนเก่า...")
    ws_dst.update("B1", [[""]], value_input_option="USER_ENTERED")

    # ===== Column AL = % Com → ใส่ PERCENT format =====
    if progress_cb:
        progress_cb("ใส่ format % Col AL...")
    al_col_range = f"AL1:AL{total_rows}"
    ws_dst.format(al_col_range, {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}})

    return f"✅ คัดลอก {total_rows} แถวจาก sum(ตัดO2O) → Com ERP สำเร็จ"


def clear_and_copy_from_sum_o2o(progress_cb=None) -> str:
    """
    ล้างข้อมูลทั้งหมดใน Com ERP แล้วคัดลอกข้อมูลทั้งหมดจากชีท sum(ตัดO2O)
    มาวางเป็นค่าธรรมดา (ไม่มีสูตร) พร้อม reset number format
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)

    # ---- อ่านข้อมูลทั้งหมดจากชีท sum(ตัดO2O) ----
    if progress_cb:
        progress_cb("อ่านข้อมูลจากชีท sum(ตัดO2O)...")
    all_titles = [ws.title for ws in sh.worksheets()]
    ws_src = None
    for t in all_titles:
        if "sum" in t.lower() and "o2o" in t.lower():
            ws_src = sh.worksheet(t)
            break
    if ws_src is None:
        raise ValueError(
            f"หาชีท sum(ตัดO2O) ไม่เจอ ชีทที่มีทั้งหมด: {all_titles}"
        )
    src_values = ws_src.get_all_values()

    if not src_values:
        raise ValueError("ชีท sum(ตัดO2O) ว่างเปล่า")

    # ---- ล้างทุกอย่างใน Com ERP ----
    if progress_cb:
        progress_cb("ล้างข้อมูลทั้งหมดใน Com ERP...")
    ws_dst = sh.worksheet(SHEET_COM_ERP)
    ws_dst.clear()

    # ---- normalize ทุก row ให้ยาวเท่ากัน ----
    max_cols = max(len(r) for r in src_values)
    normalized = [r + [""] * (max_cols - len(r)) for r in src_values]

    # ---- เขียนข้อมูลทั้งหมดลง Com ERP ----
    if progress_cb:
        progress_cb("คัดลอกข้อมูลจาก sum(ตัดO2O) → Com ERP...")
    total_rows = len(normalized)
    write_range = f"A1:{gspread.utils.rowcol_to_a1(total_rows, max_cols)}"
    ws_dst.update(write_range, normalized, value_input_option="USER_ENTERED")

    # ---- reset number format ทุก cell ให้เป็นตัวเลขธรรมดา ----
    # ป้องกัน cell ที่ถูก format % ไว้ก่อน แสดงผิด เช่น 85 → 8500%
    if progress_cb:
        progress_cb("รีเซ็ตรูปแบบตัวเลข...")
    ws_dst.format(
        f"A1:{gspread.utils.rowcol_to_a1(total_rows, max_cols)}",
        {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}
    )

    return f"✅ คัดลอก {total_rows} แถวจาก sum(ตัดO2O) → Com ERP สำเร็จ"


def copy_com_erp_to_summarize(progress_cb=None) -> str:
    """
    อ่านข้อมูลจากชีท Com ERP ใน Google Sheet
    แล้วเขียนยอดรวมแต่ละสาขาลงชีท สรุปCom

    โครงสร้าง Com ERP:
    - header row: col A = "สาขา", col B = ชื่อ, col C-AH = วันที่ (32 วัน)
    - แถว "รวม": col A = branch (เช่น "T2"), col B = "รวม", col C-AH = ยอดรวมรายวัน
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)

    # ---- อ่าน Com ERP ----
    if progress_cb:
        progress_cb("อ่านข้อมูลจากชีท Com ERP...")

    ws_com = sh.worksheet(SHEET_COM_ERP)
    all_values = ws_com.get_all_values()

    if not all_values:
        raise ValueError("ชีท Com ERP ว่างเปล่า")

    # หา header row (col A = "สาขา")
    header_idx = None
    for i, row in enumerate(all_values):
        if row and row[0] == "สาขา":
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("หา header row (สาขา) ใน Com ERP ไม่เจอ")

    # อ่านวันที่จาก header (col C-AH = index 2-33)
    # แปลง serial number (เช่น "46,031.00") → "dd/mm/yy" string
    header = all_values[header_idx]
    dates = []
    for j in range(2, 34):
        v = header[j] if j < len(header) else ""
        dates.append(_serial_to_date_str(str(v)) if v else "")

    # หาแถว "รวม" ของแต่ละสาขา
    branch_totals = {}
    for i in range(header_idx + 1, len(all_values)):
        row = all_values[i]
        if len(row) < 2:
            continue
        branch = str(row[0]).strip()
        name = str(row[1]).strip()
        if name == "รวม" and branch:
            daily = []
            for j in range(2, 34):
                v = row[j] if j < len(row) else ""
                try:
                    daily.append(float(str(v).replace(",", "")) if v != "" else 0.0)
                except (ValueError, TypeError):
                    daily.append(0.0)
            branch_totals[branch] = daily

    if not branch_totals:
        raise ValueError("หาแถว 'รวม' ในชีท Com ERP ไม่เจอ (ต้องมี col B = 'รวม')")

    if progress_cb:
        progress_cb(f"พบข้อมูล {len(branch_totals)} สาขา — กำลังเขียนลง สรุปCom...")

    # เขียนลง สรุปCom
    result = write_summarize_com(branch_totals, dates)

    return f"✅ คัดลอก Com ERP → สรุปCom สำเร็จ ({len(branch_totals)} สาขา)\n{result}"


def mark_absent_h(absent_days: dict, progress_cb=None) -> str:
    """
    เขียน 'H' ลงชีท สรุปCom สำหรับพนักงานที่ขาด/ลา หรือวันอาทิตย์

    absent_days: dict จาก parse_absent_days()  → {(emp_id_str, "dd/mm/yy"): True}

    โครงสร้าง สรุปCom:
    - Row 5 (index 4): Date header — col E (idx 4) เป็นต้นไป เช่น "20/08/26"
    - แถวพนักงาน: col A มีรหัสพนักงาน (ตัวเลข), col E-AJ = ค่าคอมรายวัน
    - ถ้า col A ว่าง = header row / รวม row → ข้าม

    กฎการเขียน H:
    - เขียน H เฉพาะ cell ที่ "ว่าง" อยู่
    - เงื่อนไข: (emp_id, date) อยู่ใน absent_days  OR  date นั้นเป็นวันอาทิตย์
    - ไม่แตะ cell ที่มีค่าอยู่แล้ว
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_SUMMARIZE)

    if progress_cb:
        progress_cb("อ่านข้อมูลจากชีท สรุปCom...")

    all_values = ws.get_all_values()

    # ---- หา date header row (row 5 = index 4) ----
    DATE_HEADER_IDX = 4          # 0-indexed (Row 5)
    DATE_COL_START  = 4          # col E = index 4 (0-indexed)

    date_row = all_values[DATE_HEADER_IDX] if len(all_values) > DATE_HEADER_IDX else []

    # อ่านวันที่จริงจาก Row 5 — หยุดที่ cell ว่างแรก
    # รองรับ 28/29/30/31 วัน โดยไม่ hardcode
    dates = []
    for i in range(50):           # เผื่อสูงสุด 50 col ไว้ก่อน
        ci = DATE_COL_START + i
        v = str(date_row[ci]).strip() if ci < len(date_row) else ""
        if not v:
            break                 # cell ว่าง = หมดวันที่แล้ว
        dates.append(v)

    NUM_DAY_COLS = len(dates)     # จำนวนจริงของเดือนนั้น

    if progress_cb:
        valid_dates = [d for d in dates if d]
        progress_cb(f"พบวันที่ {len(valid_dates)} วัน ({valid_dates[0] if valid_dates else '?'} - {valid_dates[-1] if valid_dates else '?'})")

    # ---- สแกนแถวพนักงาน ----
    updates: list[dict] = []   # [{range, values}]
    sunday_addrs: list[str] = []   # H สีแดง (วันอาทิตย์)
    absent_addrs: list[str] = []   # H สีน้ำเงิน (ขาด/ลา)
    marked_count = 0

    for row_idx, row in enumerate(all_values):
        if row_idx <= DATE_HEADER_IDX:
            continue  # ข้ามแถว header

        emp_id_raw = row[0] if row else ""
        emp_id = str(emp_id_raw).strip()

        # ข้ามแถวที่ col A ไม่ใช่ตัวเลข (header, รวม, ว่าง)
        if not emp_id or not emp_id.isdigit():
            continue

        # วนทุก day column
        for i in range(NUM_DAY_COLS):
            ci = DATE_COL_START + i         # 0-indexed column
            date_str = dates[i] if i < len(dates) else ""

            # ข้ามถ้าไม่มีวันที่ใน header
            if not date_str:
                continue

            # ค่าปัจจุบันใน cell
            cur = row[ci] if ci < len(row) else ""
            cur = str(cur).strip()

            # เขียน H เฉพาะ cell ว่างเท่านั้น
            if cur != "":
                continue

            is_sun = _is_sunday(date_str)
            is_absent = absent_days.get((emp_id, date_str), False)

            if is_sun or is_absent:
                cell_addr = gspread.utils.rowcol_to_a1(row_idx + 1, ci + 1)
                updates.append({"range": cell_addr, "values": [["H"]]})
                marked_count += 1
                if is_sun:
                    sunday_addrs.append(cell_addr)   # สีแดง
                else:
                    absent_addrs.append(cell_addr)   # สีน้ำเงิน

    if not updates:
        return "✅ ไม่มี cell ที่ต้องเติม H (อาจเติมไปแล้วหรือไม่มีวันขาด)"

    if progress_cb:
        progress_cb(f"กำลังเขียน H จำนวน {marked_count} cell...")

    # batch update ครั้งละ 500 cells
    BATCH = 500
    for i in range(0, len(updates), BATCH):
        ws.batch_update(updates[i:i + BATCH], value_input_option="USER_ENTERED")

    # format: center + สีตามประเภท
    RED  = {"red": 1.0, "green": 0.0, "blue": 0.0}   # วันอาทิตย์
    BLUE = {"red": 0.0, "green": 0.0, "blue": 1.0}   # ขาด/ลา

    BASE_FMT = {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}

    fmt_updates = []
    for addr in sunday_addrs:
        fmt_updates.append({"range": addr, "format": {**BASE_FMT, "textFormat": {"foregroundColor": RED}}})
    for addr in absent_addrs:
        fmt_updates.append({"range": addr, "format": {**BASE_FMT, "textFormat": {"foregroundColor": BLUE}}})

    for i in range(0, len(fmt_updates), BATCH):
        ws.batch_format(fmt_updates[i:i + BATCH])

    return f"✅ เติม H สำเร็จ {marked_count} cell (🔴 อาทิตย์ {len(sunday_addrs)} | 🔵 ขาด/ลา {len(absent_addrs)})"


def read_sum_sheet(xlsx_path: str) -> tuple:
    """
    อ่านข้อมูลจากชีท sum(ตัดO2O) ในไฟล์ xlsx
    คืนค่า (rows, dates, branch_totals)
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    # หาชีท sum(ตัดO2O) แบบ case-insensitive
    ws = None
    for name in wb.sheetnames:
        if "sum" in name.lower() and "o2o" in name.lower():
            ws = wb[name]
            break
    if ws is None:
        raise ValueError(f"ไม่พบชีท sum(ตัดO2O) ในไฟล์\nชีทที่มี: {wb.sheetnames}")

    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        raise ValueError("ชีท sum(ตัดO2O) ว่างเปล่า")

    # หา header row (col A = "สาขา")
    header_idx = None
    for i, row in enumerate(all_rows):
        if row and row[0] == "สาขา":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("หา header ใน sum(ตัดO2O) ไม่เจอ")

    # อ่านวันที่จาก header (col 3-34)
    import datetime as _dt
    header = all_rows[header_idx]
    dates = []
    for j in range(2, 34):  # index 2-33 = col C-AH
        v = header[j] if j < len(header) else None
        if v is not None:
            if hasattr(v, 'strftime'):
                # datetime/date object จาก openpyxl
                dates.append(v.strftime("%d/%m/%y"))
            elif isinstance(v, (int, float)) and 40000 < v < 60000:
                # Excel serial number → แปลงเป็นวันที่
                # Excel epoch: Dec 30, 1899 (รวม leap-year bug ของ Excel)
                dt = _dt.date(1899, 12, 30) + _dt.timedelta(days=int(v))
                dates.append(dt.strftime("%d/%m/%y"))
            else:
                dates.append(str(v))
        else:
            dates.append("")

    # อ่านข้อมูลแถวหลัง header
    rows = []
    branch_totals = {}  # {"T2": [32 vals]}

    for i in range(header_idx + 1, len(all_rows)):
        row = all_rows[i]
        if not row or row[0] is None:
            continue

        branch = str(row[0]) if row[0] else ""
        name = str(row[1]) if row[1] else ""

        # daily values (col 3-34, index 2-33)
        daily = []
        for j in range(2, 34):
            v = row[j] if j < len(row) else None
            try:
                daily.append(float(v) if v is not None else 0.0)
            except (ValueError, TypeError):
                daily.append(0.0)  # กรณี #NAME?, #REF!, #N/A ฯลฯ

        # summary cols (index 35=คอมรายคน, 36=คอมรวม, 37=ยอดขาย, 38=%Com)
        def _f(idx):
            try:
                return float(row[idx]) if idx < len(row) and row[idx] is not None else 0.0
            except (ValueError, TypeError):
                return 0.0  # กรณี #NAME?, #REF!, #N/A ฯลฯ

        com_pp = _f(35)
        com_tot = _f(36)
        sales = _f(37)
        pct_com = _f(38)

        is_total = (name == "รวม")

        rows.append({
            "branch": branch,
            "name": name,
            "daily": daily,
            "com_pp": com_pp,
            "com_tot": com_tot,
            "sales": sales,
            "pct_com": pct_com,
            "is_total": is_total,
        })

        # เก็บ branch totals (แถว "รวม" ของแต่ละสาขา)
        if is_total and branch:
            branch_totals[branch] = daily

    return rows, dates, branch_totals
