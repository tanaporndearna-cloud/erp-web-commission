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
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def write_com_erp(rows: list, dates: list, progress_cb=None) -> str:
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_COM_ERP)

    all_values = ws.get_all_values()

    header_row_idx = None
    for i, row in enumerate(all_values):
        if row and row[0] == "สาขา":
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("หา header row (สาขา) ใน Com ERP ไม่เจอ")

    header_row = header_row_idx + 1
    data_start_row = header_row + 1
    date_start_col = 3
    summary_start = 36

    date_updates = []
    for i, d in enumerate(dates):
        date_updates.append({
            "range": gspread.utils.rowcol_to_a1(header_row, date_start_col + i),
            "values": [[d]]
        })
    ws.batch_update(date_updates)

    batch_data = []
    for ri, d in enumerate(rows):
        row = data_start_row + ri
        row_vals = [d["branch"], d["name"]]
        for di in range(32):
            val = d["daily"][di] if d["daily"] and di < len(d["daily"]) else 0
            row_vals.append(round(float(val or 0), 2))
        while len(row_vals) < 34:
            row_vals.append("")
        row_vals.append("")
        row_vals.append(round(float(d.get("com_pp") or 0), 2))
        row_vals.append(round(float(d.get("com_tot") or 0), 2))
        row_vals.append(round(float(d.get("sales") or 0), 2))
        row_vals.append(round(float(d.get("pct_com") or 0), 6))

        batch_data.append({
            "range": f"A{row}:{gspread.utils.rowcol_to_a1(row, len(row_vals))}",
            "values": [row_vals]
        })

        if progress_cb:
            progress_cb(ri + 1, len(rows))

    ws.batch_update(batch_data)

    pct_range = f"{gspread.utils.rowcol_to_a1(data_start_row, summary_start + 3)}:{gspread.utils.rowcol_to_a1(data_start_row + len(rows) - 1, summary_start + 3)}"
    ws.format(pct_range, {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}})

    return f"✅ เขียน Com ERP สำเร็จ {len(rows)} แถว"


def write_summarize_com(branch_totals: dict, date_start_day: int = 20) -> str:
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(SHEET_SUMMARIZE)

    all_values = ws.get_all_values()

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

    date_header_row = date_header_row_idx + 1
    date_col_start = 5

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
        row = trc_row_map[branch]
        row_vals = [round(float(v or 0), 2) for v in vals[:32]]
        start_cell = gspread.utils.rowcol_to_a1(row, date_col_start)
        end_cell = gspread.utils.rowcol_to_a1(row, date_col_start + len(row_vals) - 1)
        batch_data.append({
            "range": f"{start_cell}:{end_cell}",
            "values": [row_vals]
        })
        written += 1

    ws.batch_update(batch_data)
    return f"✅ เขียน สรุปCom สำเร็จ {written}/{len(trc_row_map)} สาขา"


SHEET_SUM_O2O = "sum (ตัดO2O)"


def clear_and_copy_from_sum_o2o(progress_cb=None) -> str:
    gc = get_client()
    sh = gc.open_by_key(SHEET_ID)

    if progress_cb:
        progress_cb("อ่านข้อมูลจากชีท sum(ตัดO2O)...")
    ws_src = sh.worksheet(SHEET_SUM_O2O)
    src_values = ws_src.get_all_values()

    if not src_values:
        raise ValueError("ชีท sum(ตัดO2O) ว่างเปล่า")

    if progress_cb:
        progress_cb("ล้างข้อมูลใน Com ERP...")
    ws_dst = sh.worksheet(SHEET_COM_ERP)
    dst_values = ws_dst.get_all_values()

    header_row_idx = None
    for i, row in enumerate(dst_values):
        if row and row[0] == "สาขา":
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("หา header row (สาขา) ใน Com ERP ไม่เจอ")

    data_start_row = header_row_idx + 2
    total_rows = len(dst_values)

    if total_rows >= data_start_row:
        n_cols = max(len(r) for r in src_values) if src_values else 39
        clear_range = f"A{data_start_row}:{gspread.utils.rowcol_to_a1(total_rows, n_cols)}"
        ws_dst.batch_clear([clear_range])

    src_header_idx = None
    for i, row in enumerate(src_values):
        if row and row[0] == "สาขา":
            src_header_idx = i
            break

    if src_header_idx is None:
        raise ValueError("หา header row (สาขา) ใน sum(ตัดO2O) ไม่เจอ")

    if progress_cb:
        progress_cb("คัดลอกวันที่จาก sum(ตัดO2O)...")
    src_header = src_values[src_header_idx]
    dst_header_row = header_row_idx + 1

    date_col_start = 3
    date_vals = src_header[2:]
    if date_vals:
        date_range = (
            f"{gspread.utils.rowcol_to_a1(dst_header_row, date_col_start)}"
            f":{gspread.utils.rowcol_to_a1(dst_header_row, date_col_start + len(date_vals) - 1)}"
        )
        ws_dst.update(date_range, [date_vals], value_input_option="USER_ENTERED")

    if progress_cb:
        progress_cb("คัดลอกข้อมูลพนักงานจาก sum(ตัดO2O)...")

    data_rows = src_values[src_header_idx + 1:]
    data_rows = [r for r in data_rows if any(v for v in r)]

    if data_rows:
        max_cols = max(len(r) for r in data_rows)
        data_rows = [r + [""] * (max_cols - len(r)) for r in data_rows]

        write_range = (
            f"A{data_start_row}"
            f":{gspread.utils.rowcol_to_a1(data_start_row + len(data_rows) - 1, max_cols)}"
        )
        ws_dst.update(write_range, data_rows, value_input_option="USER_ENTERED")

    return f"✅ คัดลอก {len(data_rows)} แถวจาก sum(ตัดO2O) → Com ERP สำเร็จ"


def read_sum_sheet(xlsx_path: str) -> tuple:
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["sum (ตัดO2O)"]

    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        raise ValueError("ชีท sum(ตัดO2O) ว่างเปล่า")

    header_idx = None
    for i, row in enumerate(all_rows):
        if row and row[0] == "สาขา":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("หา header ใน sum(ตัดO2O) ไม่เจอ")

    header = all_rows[header_idx]
    dates = []
    for j in range(2, 34):
        v = header[j] if j < len(header) else None
        if v is not None:
            if hasattr(v, 'strftime'):
                dates.append(v.strftime("%d/%m/%y"))
            else:
                dates.append(str(v))
        else:
            dates.append("")

    rows = []
    branch_totals = {}

    for i in range(header_idx + 1, len(all_rows)):
        row = all_rows[i]
        if not row or row[0] is None:
            continue

        branch = str(row[0]) if row[0] else ""
        name = str(row[1]) if row[1] else ""

        daily = []
        for j in range(2, 34):
            v = row[j] if j < len(row) else None
            daily.append(float(v) if v is not None else 0.0)

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

        if is_total and branch:
            branch_totals[branch] = daily

    return rows, dates, branch_totals
