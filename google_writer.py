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


def write_summarize_com(branch_totals: dict, dates: list = None, date_start_day: int = 20) -> str:
    """
    เขียนข้อมูลลงชีท สรุปCom โดย match วันที่จาก xlsx กับ header ของ Google Sheet
    branch_totals: {"T2": [v1..v32], "T3": [...], ...}
    dates: list of date strings เช่น ["21/08/69", "22/08/69", ...] จาก read_sum_sheet
           ถ้าไม่ส่งมา จะวางตามตำแหน่ง (พฤติกรรมเดิม)
    """
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_SUMMARIZE)

    all_values = ws.get_all_values()

    # หา date header row (col E = date_start_day หรือ 1)
    date_header_row_idx = None
    for i, row in enumerate(all_values):
        if len(row) >= 5:
            try:
                v = int(row[4])
                if v == date_start_day or v == 1:
                    date_header_row_idx = i
                    break
            except (ValueError, TypeError):
                pass

    if date_header_row_idx is None:
        raise ValueError(f"หา date header row ใน สรุปCom ไม่เจอ (col E = {date_start_day})")

    date_header_row = date_header_row_idx + 1  # 1-indexed
    date_col_start = 5  # col E

    # สร้าง mapping: day_number → col_index (1-indexed) จาก header ของ Google Sheet
    sheet_date_row = all_values[date_header_row_idx]
    day_to_col = {}  # {21: 5, 22: 6, ..., 31: 15, 1: 16, ...}
    sheet_last_date_col = date_col_start
    for ci in range(date_col_start - 1, len(sheet_date_row)):
        try:
            day_num = int(sheet_date_row[ci])
            day_to_col[day_num] = ci + 1  # 1-indexed
            sheet_last_date_col = ci + 1
        except (ValueError, TypeError):
            pass

    def _date_str_to_day(date_str: str):
        """แปลง '21/08/69' → 21"""
        try:
            return int(str(date_str).split('/')[0])
        except Exception:
            return None

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

        if dates and day_to_col:
            # ===== Match by day number =====
            # สร้าง dict: day_number → value จาก xlsx
            day_val_map = {}
            for vi, date_str in enumerate(dates[:32]):
                if vi >= len(vals):
                    break
                day_num = _date_str_to_day(date_str)
                if day_num is not None:
                    day_val_map[day_num] = round(float(vals[vi] or 0), 2)

            # เขียนทีละ cell ตาม col ที่ตรงกับ day_number
            for day_num, col in day_to_col.items():
                val = day_val_map.get(day_num, 0)
                cell_a1 = gspread.utils.rowcol_to_a1(row_idx, col)
                batch_data.append({"range": cell_a1, "values": [[val]]})
        else:
            # ===== Fallback: วางตามตำแหน่ง (พฤติกรรมเดิม) =====
            row_vals = [round(float(v or 0), 2) for v in vals[:32]]
            start_cell = gspread.utils.rowcol_to_a1(row_idx, date_col_start)
            end_cell = gspread.utils.rowcol_to_a1(row_idx, date_col_start + len(row_vals) - 1)
            batch_data.append({"range": f"{start_cell}:{end_cell}", "values": [row_vals]})

        written += 1

    ws.batch_update(batch_data)

    # ตั้ง number format #,##0.00 ให้ช่วงวันที่ทั้งหมด
    for branch in branch_totals:
        if branch not in trc_row_map:
            continue
        row_idx = trc_row_map[branch]
        start_cell = gspread.utils.rowcol_to_a1(row_idx, date_col_start)
        end_cell = gspread.utils.rowcol_to_a1(row_idx, sheet_last_date_col)
        ws.format(f"{start_cell}:{end_cell}", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}})

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
    อัปเดตหัวชีท A1 ของ 4 ชีทตามเดือน/ปีที่คิดค่าคอม
    รูปแบบ: 'สรุปคอมมิชชั่น เดือน กันยายน 2569'
    ค้นหาชีทแบบ case-insensitive
    """
    month_name = THAI_MONTHS[month] if 1 <= month <= 12 else f"เดือน{month}"
    title = f"สรุปคอมมิชชั่น เดือน {month_name} {year_be}"

    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)

    all_worksheets = sh.worksheets()
    all_sheet_names = {ws.title.lower(): ws for ws in all_worksheets}
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
        """แปลง Excel date serial → dd/mm/yy"""
        try:
            return (_dt(1899, 12, 30) + _td(days=int(val))).strftime("%d/%m/%y")
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
                processed.append(val.strftime("%d/%m/%y"))
            elif isinstance(val, (int, float)) and cell.is_date:
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

    # ===== อัปเดต title ตามเดือน/ปีที่เลือก =====
    if month and year_be:
        if progress_cb:
            progress_cb("อัปเดตชื่อเดือนใน Com ERP...")
        title = _make_com_title(month, year_be)
        ws_dst.update("B1", [[title]], value_input_option="USER_ENTERED")

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
            daily.append(float(v) if v is not None else 0.0)

        # summary cols (index 35=คอมรายคน, 36=คอมรวม, 37=ยอดขาย, 38=%Com)
        def _f(idx):
            return float(row[idx]) if idx < len(row) and row[idx] is not None else 0.0

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
