"""
payroll_page.py — TRC Motorsport: จัดการเงินเดือนรายเดือน
เพิ่มเข้า Streamlit app เดิมได้เลย

วิธีใช้ใน app.py หลัก:
    from payroll_page import render_payroll_page
    render_payroll_page(gc, sheet_id)
    # หรือถ้ามี history DB:
    render_payroll_page(gc, sheet_id, history_db_id="1UMB2LlO_8BKevg_dvIrNmFeBpOG0OadUrYcgejBzPA4")
"""

import streamlit as st
import pandas as pd
import gspread
from datetime import date, timedelta
import re
import time


def _sheets_retry(func, *args, max_retries: int = 6, **kwargs):
    """เรียก Google Sheets API พร้อม retry อัตโนมัติเมื่อเจอ 429 (rate limit)"""
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = getattr(e.response, "status_code", None)
            if status == 429 and attempt < max_retries - 1:
                wait = 10 * (2 ** attempt)   # 10 → 20 → 40 → 80 → 160 วินาที
                time.sleep(wait)
            else:
                raise

# ── Config ──────────────────────────────────────────────────────
# ค่า config ตรงกับ Code.gs ที่ deploy ไว้ใน Google Sheet
CFG = {
    "HEADER_ROW"    : 5,    # แถวหัวตาราง (1-indexed)
    "DATA_START"    : 6,    # แถวแรกที่มีข้อมูล
    "DATA_END"      : 40,   # แถวสุดท้าย
    "FIRST_ATT_COL" : 13,   # คอลัมน์ M (1-indexed) — เริ่มต้นช่องเวลา

    "HDR_DATE"     : "วันที่",
    "HDR_TIME_IN"  : "เวลาเข้า",
    "HDR_TIME_OUT" : "เวลาออก",
    "HDR_NOTE"     : "หมายเหตุ",
    "HDR_DAY_NAME" : "ชื่อวัน",

    # คอลัมน์ในไฟล์เวลาเข้า-ออก (0-indexed) — ปรับตามไฟล์จริงถ้าจำเป็น
    "ATT_DAY"     : 0,   # ชื่อวัน
    "ATT_DATE"    : 5,   # วันที่
    "ATT_TIME_IN" : 6,   # เวลาเข้างาน
    "ATT_TIME_OUT": 7,   # เวลาออกงาน
    "ATT_NOTE"    : 14,  # หมายเหตุ

    "HISTORY_SHEET": "📚 ประวัติเงินเดือน",
}

HISTORY_COLS = ["บันทึกเมื่อ", "Sheet", "เดือน/ปี", "วันที่",
                "ชื่อวัน", "เวลาเข้า", "เวลาออก", "หมายเหตุ"]

DAY_TH   = ["อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส."]
MONTH_TH = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
             "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


# ── Helpers ─────────────────────────────────────────────────────

def col_letter(n: int) -> str:
    """แปลงเลขคอลัมน์ (1-indexed) เป็นตัวอักษร เช่น 13 → M"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def norm_date(val) -> str | None:
    """แปลงค่าต่าง ๆ เป็น 'DD/MM/YYYY' (พ.ศ.)"""
    if isinstance(val, str):
        val = val.strip()
        if re.match(r"^\d{2}/\d{2}/\d{4}$", val):
            return val
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                d = pd.to_datetime(val, format=fmt)
                return d.strftime("%d/%m/") + str(d.year + 543)
            except Exception:
                pass
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d/%m/") + str(val.year + 543)
    return None


def find_col_indices(headers: list, *target_names) -> dict:
    """หาตำแหน่งคอลัมน์ทุกตัวที่ตรงกับ target_names → {name: [col1, col2, ...]}"""
    result = {name: [] for name in target_names}
    for i, h in enumerate(headers):
        h = str(h).strip()
        for name in target_names:
            if h == name:
                result[name].append(i + 1)   # 1-indexed
    return result


def detect_period_from_df(att_df: pd.DataFrame):
    """ดึงเดือน/ปี (พ.ศ.) จากข้อมูลในไฟล์ → (month, year_be)"""
    try:
        dates = att_df["วันที่"].dropna()
        for d_str in reversed(dates.tolist()):
            m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(d_str))
            if m:
                day, mon, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if day == 25:   # วันสิ้นสุดรอบ = 25
                    return mon, yr
        last = dates.iloc[-1]
        m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(last))
        if m:
            return int(m.group(2)), int(m.group(3))
    except Exception:
        pass
    return date.today().month, date.today().year + 543


def _get_or_create_history_sheet(ss: gspread.Spreadsheet) -> gspread.Worksheet:
    """เปิดหรือสร้างชีต '📚 ประวัติเงินเดือน'"""
    try:
        return ss.worksheet(CFG["HISTORY_SHEET"])
    except gspread.WorksheetNotFound:
        hist = ss.add_worksheet(CFG["HISTORY_SHEET"], rows=2000, cols=10)
        hist.append_row(HISTORY_COLS)
        hist.freeze(rows=1)
        return hist


# ── Main render function ─────────────────────────────────────────

def render_payroll_page(gc: gspread.Client,
                        sheet_id: str,
                        history_db_id: str | None = "1UMB2LlO_8BKevg_dvIrNmFeBpOG0OadUrYcgejBzPA4"):
    """
    gc            = gspread.Client ที่ authenticate แล้ว
    sheet_id      = Google Sheet ID ของไฟล์เทมเพลตเงินเดือน
    history_db_id = Google Sheet ID ของฐานเก็บประวัติเงินเดือน (None = ไม่บันทึก DB)
    """
    st.header("👷 จัดการเงินเดือนรายเดือน")

    try:
        ss = gc.open_by_key(sheet_id)
    except Exception as e:
        st.error(f"เปิด Spreadsheet ไม่ได้: {e}")
        return

    # เปิด history DB ถ้ามี
    ss_db = None
    if history_db_id:
        try:
            ss_db = gc.open_by_key(history_db_id)
        except Exception:
            ss_db = None

    visible_sheets = [ws.title for ws in ss.worksheets()]

    # ══════════════════════════════════════════════════════════════
    # คิดเงินเดือนพนักงาน — Import เวลาเข้า-ออกงาน
    # ══════════════════════════════════════════════════════════════
    with st.expander("🧾  คิดเงินเดือนพนักงาน (Import เวลาเข้า-ออกงาน)", expanded=False):
        st.caption("อัปโหลดไฟล์เวลา → ล้างเก่า → สร้างวันที่ → วางเวลา → บันทึกประวัติ (ประมวลผลทุก Sheet พนักงานอัตโนมัติ)")

        att_file = st.file_uploader("📂 ไฟล์เวลาเข้า-ออกงาน (.xlsx / .xls)",
                                    type=["xlsx", "xls"], key="att_file")

        if att_file:
            att_df = _read_attendance_file(att_file)
            if att_df is not None and not att_df.empty:
                auto_month, auto_year = detect_period_from_df(att_df)
                m_label = f"{MONTH_TH[auto_month]} {auto_year}"
                st.info(f"📅 พบ **{len(att_df)} วัน** — รอบเดือน {m_label}")
                st.dataframe(
                    att_df[["วัน", "วันที่", "เวลาเข้า", "เวลาออก", "หมายเหตุ"]].head(10),
                    use_container_width=True, hide_index=True
                )

                # Auto-detect employee sheets — ยกเว้น history sheet
                EXCLUDE_SHEETS = {CFG["HISTORY_SHEET"]}
                emp_sheets = [s for s in visible_sheets if s not in EXCLUDE_SHEETS]
                st.info(f"🗂 พบ **{len(emp_sheets)} Sheet** พนักงานที่จะประมวลผล")

                save_hist = st.checkbox(
                    "📚 บันทึกประวัติ (เทมเพลต" +
                    (" + ฐานข้อมูล" if ss_db else "") + ")",
                    value=True, key="chk_hist"
                )

                if st.button("▶ เริ่มคิดเงินเดือนทุก Sheet พนักงาน (ล้างเก่า → สร้างวันที่ → วางเวลา)",
                             type="primary", key="btn_payroll"):

                    success_count = 0
                    fail_count    = 0
                    errors        = []
                    progress_bar  = st.progress(0)
                    status_text   = st.empty()

                    for idx, sh_pay in enumerate(emp_sheets):
                        status_text.info(
                            f"⏳ กำลังประมวลผล Sheet **{sh_pay}** ({idx+1}/{len(emp_sheets)})..."
                        )
                        try:
                            # Step 1: ล้างข้อมูลเดือนเก่า
                            r1 = _clear_attendance(ss, sh_pay)
                            if not r1["ok"]:
                                errors.append(f"{sh_pay}: ล้างไม่ได้ — {r1['msg']}")
                                fail_count += 1
                                progress_bar.progress((idx + 1) / len(emp_sheets))
                                continue

                            # Step 2: สร้างวันที่
                            r2 = _generate_dates(ss, sh_pay, auto_month, auto_year)
                            if not r2["ok"]:
                                errors.append(f"{sh_pay}: สร้างวันที่ไม่ได้ — {r2['msg']}")
                                fail_count += 1
                                progress_bar.progress((idx + 1) / len(emp_sheets))
                                continue

                            # Step 3: วางข้อมูลเวลา
                            r3 = _import_attendance(ss, sh_pay, att_df.to_dict("records"))
                            if not r3["ok"]:
                                errors.append(f"{sh_pay}: วางเวลาไม่ได้ — {r3['msg']}")
                                fail_count += 1
                                progress_bar.progress((idx + 1) / len(emp_sheets))
                                continue

                            # Step 4: บันทึกประวัติ
                            if save_hist:
                                month_yr = f"{auto_month}/{auto_year}"
                                r4 = _export_history(ss, sh_pay, month_yr, ss_db)
                                if not r4["ok"]:
                                    errors.append(f"{sh_pay}: บันทึกประวัติไม่ได้ — {r4['msg']}")

                            success_count += 1

                        except Exception as e:
                            errors.append(f"{sh_pay}: ข้อผิดพลาด — {str(e)}")
                            fail_count += 1

                        progress_bar.progress((idx + 1) / len(emp_sheets))

                    # สรุปผล
                    status_text.empty()
                    if fail_count == 0:
                        st.success(f"✅ เสร็จสิ้น! ประมวลผลครบ {success_count} Sheet ค่ะ")
                    else:
                        st.warning(
                            f"⚠️ เสร็จ {success_count} Sheet / ล้มเหลว {fail_count} Sheet"
                        )
                    if errors:
                        with st.expander("❌ รายการที่มีปัญหา"):
                            for err in errors:
                                st.text(err)
            else:
                st.warning("⚠️ อ่านไฟล์ไม่ได้ หรือไม่พบข้อมูล")


# ── Backend: ERP Import ──────────────────────────────────────────

def _read_erp_file(file) -> pd.DataFrame | None:
    """อ่านไฟล์ ERP — ลองหลาย header row จนได้ข้อมูล"""
    try:
        file.seek(0)
        for hdr in [0, 1, 12, 13]:
            try:
                file.seek(0)
                df = pd.read_excel(file, header=hdr)
                df.columns = [str(c).strip() for c in df.columns]
                df = df.dropna(how="all")
                if len(df) > 0 and len(df.columns) > 3:
                    return df
            except Exception:
                pass
        return None
    except Exception as e:
        st.warning(f"อ่านไฟล์ ERP ไม่ได้: {e}")
        return None


def _import_erp(ss: gspread.Spreadsheet, sheet_name: str, df: pd.DataFrame) -> dict:
    """นำเข้าข้อมูล ERP ลง Google Sheet"""
    try:
        ws   = ss.worksheet(sheet_name)
        rows = [df.columns.tolist()] + df.astype(str).values.tolist()
        ws.clear()
        ws.update("A1", rows)
        return {"ok": True, "msg": f"นำเข้า ERP {len(df)} แถว ลงชีต \"{sheet_name}\" เรียบร้อยค่ะ"}
    except gspread.WorksheetNotFound:
        return {"ok": False, "msg": f"ไม่พบ sheet: {sheet_name}"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Backend: Copy sum(ตัดO2O) → Com ERP ─────────────────────────

def _copy_sum_to_com(ss: gspread.Spreadsheet, src_sheet: str, dst_sheet: str) -> dict:
    """คัดลอกยอดจากชีต sum(ตัดO2O) → คอลัมน์ Com ERP"""
    try:
        src = ss.worksheet(src_sheet)
        dst = ss.worksheet(dst_sheet)

        src_data = src.get_all_values()
        if len(src_data) < 2:
            return {"ok": False, "msg": f"ชีต \"{src_sheet}\" ไม่มีข้อมูล"}

        headers  = [str(h).strip() for h in src_data[0]]
        name_col = next((i for i, h in enumerate(headers)
                         if any(k in h for k in ["ชื่อ", "พนักงาน", "name"])), 0)
        com_col  = next((i for i, h in enumerate(headers)
                         if any(k in h for k in ["คอม", "commission", "ยอด", "รวม"])), 1)

        commission_map: dict[str, str] = {}
        for row in src_data[1:]:
            if len(row) > max(name_col, com_col) and str(row[name_col]).strip():
                commission_map[str(row[name_col]).strip()] = str(row[com_col]).strip()

        if not commission_map:
            return {"ok": False, "msg": "ไม่พบข้อมูลค่าคอมในชีตต้นทาง"}

        dst_data    = dst.get_all_values()
        dst_headers = [str(h).strip() for h in dst_data[0]] if dst_data else []

        com_erp_col = next(
            (i for i, h in enumerate(dst_headers)
             if "com erp" in h.lower() or "comErp" in h or "ค่าคอมerp" in h.lower()),
            None
        )
        if com_erp_col is None:
            return {
                "ok": False,
                "msg": (f"ไม่พบคอลัมน์ 'Com ERP' ใน \"{dst_sheet}\" "
                        f"(พบ: {', '.join(dst_headers[:10])})")
            }

        dst_name_col = next((i for i, h in enumerate(dst_headers)
                             if any(k in h for k in ["ชื่อ", "พนักงาน", "name"])), 0)
        updates = []
        matched = 0
        for r_i, row in enumerate(dst_data[1:], start=2):
            name = str(row[dst_name_col]).strip() if len(row) > dst_name_col else ""
            if name in commission_map:
                updates.append({"range": f"{col_letter(com_erp_col + 1)}{r_i}",
                                "values": [[commission_map[name]]]})
                matched += 1

        if updates:
            dst.batch_update(updates)
        return {"ok": True,
                "msg": f"คัดลอกค่าคอม {matched} คน จาก \"{src_sheet}\" → \"{dst_sheet}\" เรียบร้อยค่ะ"}

    except gspread.WorksheetNotFound as e:
        return {"ok": False, "msg": f"ไม่พบ sheet: {e}"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Backend: Attendance & Payroll ────────────────────────────────

def _clear_attendance(ss: gspread.Spreadsheet, sheet_name: str) -> dict:
    """ล้างข้อมูลเวลาที่กรอกมือ
    เฉพาะ col P-S (16-19) และ col X-Y (24-25) แถว 6-36 เท่านั้น
    (คงสูตรไว้ — ไม่แตะ cell ที่เป็น formula)
    """
    ROW_START  = 6
    ROW_END    = 36
    # กลุ่มคอลัมน์ที่ต้องล้าง: [(col_start, col_end), ...]  1-indexed
    COL_GROUPS = [(16, 19), (24, 25)]   # P-S, X-Y

    try:
        ws = ss.worksheet(sheet_name)

        to_clear = []
        for col_s, col_e in COL_GROUPS:
            rng_a1 = (f"{col_letter(col_s)}{ROW_START}:"
                      f"{col_letter(col_e)}{ROW_END}")
            formulas = ws.get(rng_a1, value_render_option="FORMULA")
            values   = ws.get(rng_a1, value_render_option="FORMATTED_VALUE")

            num_rows = ROW_END - ROW_START + 1
            num_cols = col_e - col_s + 1
            for r_i in range(num_rows):
                frow = formulas[r_i] if r_i < len(formulas) else []
                vrow = values[r_i]   if r_i < len(values)   else []
                for c_i in range(num_cols):
                    f = frow[c_i] if c_i < len(frow) else ""
                    v = vrow[c_i] if c_i < len(vrow) else ""
                    if not str(f).startswith("=") and str(v).strip() not in ("", "None"):
                        to_clear.append(
                            f"{col_letter(col_s + c_i)}{ROW_START + r_i}"
                        )

        if not to_clear:
            return {"ok": True, "msg": "ไม่มีข้อมูลที่ต้องล้าง (สะอาดอยู่แล้ว)"}
        ws.batch_clear(to_clear)
        return {"ok": True, "msg": f"ล้างแล้ว {len(to_clear)} เซลล์ (สูตรยังอยู่ครบ)"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def _generate_dates(ss: gspread.Spreadsheet, sheet_name: str,
                    month: int, year_be: int) -> dict:
    """สร้างวันที่รอบ 26 เดือนก่อน → 25 เดือนนี้"""
    try:
        ws      = ss.worksheet(sheet_name)
        year_ce = year_be - 543

        prev_month = 12 if month == 1 else month - 1
        prev_year  = year_ce - 1 if month == 1 else year_ce
        start = date(prev_year, prev_month, 26)
        end   = date(year_ce, month, 25)

        days, cur = [], start
        while cur <= end:
            days.append(cur)
            cur += timedelta(days=1)

        headers       = ws.row_values(CFG["HEADER_ROW"])
        cols          = find_col_indices(headers, CFG["HDR_DATE"], CFG["HDR_DAY_NAME"])
        date_cols     = cols[CFG["HDR_DATE"]]
        day_name_cols = cols[CFG["HDR_DAY_NAME"]]

        if not date_cols:
            return {"ok": False,
                    "msg": f'ไม่พบคอลัมน์ "{CFG["HDR_DATE"]}" ใน row {CFG["HEADER_ROW"]}'}

        updates  = []
        num_rows = min(len(days), CFG["DATA_END"] - CFG["DATA_START"] + 1)

        for bi, dc in enumerate(date_cols):
            dnc = day_name_cols[bi] if bi < len(day_name_cols) else None
            for i in range(num_rows):
                row      = CFG["DATA_START"] + i
                d        = days[i]
                date_str = d.strftime("%d/%m/") + str(d.year + 543)
                day_idx  = (d.weekday() + 1) % 7   # 0=อา. 1=จ. ...
                updates.append({"range": f"{col_letter(dc)}{row}",
                                "values": [[date_str]]})
                if dnc:
                    DAY_EN_LIST = ["Su","Mo","Tu","We","Th","Fr","Sa"]
                    updates.append({"range": f"{col_letter(dnc)}{row}",
                                    "values": [[DAY_EN_LIST[day_idx]]]})
            # ไม่ clear แถวที่เกิน เพื่อไม่ให้ทับส่วนสรุปด้านล่าง

        ws.batch_update(updates)
        return {"ok": True,
                "msg": f"สร้างวันที่ {num_rows} วัน ใน {len(date_cols)} block เรียบร้อยค่ะ"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def _read_attendance_file(file) -> pd.DataFrame | None:
    """อ่านไฟล์เวลาเข้า-ออก → DataFrame (columns: วัน, วันที่, เวลาเข้า, เวลาออก, หมายเหตุ)"""
    try:
        file.seek(0)
        df_raw  = pd.read_excel(file, header=None)
        hdr_idx = 0
        for i in range(min(5, len(df_raw))):
            row_str = "|".join(df_raw.iloc[i].astype(str).tolist())
            if "วันที่" in row_str or "เวลาเข้างาน" in row_str:
                hdr_idx = i
                break

        rows = []
        for i in range(hdr_idx + 1, len(df_raw)):
            row      = df_raw.iloc[i]
            # ตรวจสอบว่ามีข้อมูล index ครบก่อน access
            max_idx = max(CFG["ATT_DATE"], CFG["ATT_TIME_IN"],
                          CFG["ATT_TIME_OUT"], CFG["ATT_NOTE"], CFG["ATT_DAY"])
            if len(row) <= max_idx:
                continue
            date_val = row.iloc[CFG["ATT_DATE"]]
            date_str = norm_date(str(date_val).strip()) if pd.notna(date_val) else None
            if not date_str:
                continue

            def safe_str(idx):
                v = row.iloc[idx]
                return str(v).strip() if pd.notna(v) else ""

            rows.append({
                "วัน"     : safe_str(CFG["ATT_DAY"]),
                "วันที่"  : date_str,
                "เวลาเข้า": safe_str(CFG["ATT_TIME_IN"]),
                "เวลาออก" : safe_str(CFG["ATT_TIME_OUT"]),
                "หมายเหตุ": safe_str(CFG["ATT_NOTE"]),
            })
        return pd.DataFrame(rows) if rows else None
    except Exception as e:
        st.warning(f"อ่านไฟล์ไม่ได้: {e}")
        return None


def _import_attendance(ss: gspread.Spreadsheet, sheet_name: str,
                       rows: list) -> dict:
    """วางข้อมูลเวลาเข้า-ออกลง template (ข้ามเซลล์ที่มีสูตร)"""
    try:
        ws      = ss.worksheet(sheet_name)
        att_map = {r["วันที่"]: r for r in rows if r.get("วันที่")}

        headers   = ws.row_values(CFG["HEADER_ROW"])
        cols      = find_col_indices(headers,
                                     CFG["HDR_DATE"], CFG["HDR_TIME_IN"],
                                     CFG["HDR_TIME_OUT"], CFG["HDR_NOTE"],
                                     CFG["HDR_DAY_NAME"])
        date_cols = cols[CFG["HDR_DATE"]]
        ti_cols   = cols[CFG["HDR_TIME_IN"]]
        to_cols   = cols[CFG["HDR_TIME_OUT"]]
        note_cols = cols[CFG["HDR_NOTE"]]
        day_cols  = cols[CFG["HDR_DAY_NAME"]]

        if not date_cols:
            return {"ok": False,
                    "msg": f'ไม่พบคอลัมน์ "{CFG["HDR_DATE"]}" ในเทมเพลต'}

        # ดึง formula ทั้ง block เพื่อตรวจ
        start_c  = min(date_cols)
        end_c    = ws.col_count
        rng      = (f"{col_letter(start_c)}{CFG['DATA_START']}:"
                    f"{col_letter(end_c)}{CFG['DATA_END']}")
        formulas = ws.get(rng, value_render_option="FORMULA")

        # สร้าง map วันที่ → row number สำหรับแต่ละ block
        date_row_map: dict[tuple, int] = {}
        for bi, dc in enumerate(date_cols):
            col_data = ws.col_values(dc)
            for r_i, val in enumerate(col_data[CFG["DATA_START"] - 1: CFG["DATA_END"]]):
                date_key = norm_date(str(val).strip())
                if date_key:
                    date_row_map[(bi, date_key)] = CFG["DATA_START"] + r_i

        def is_formula(c, r_num):
            if c is None:
                return True
            ci        = c - start_c
            f_row_idx = r_num - CFG["DATA_START"]
            if f_row_idx >= len(formulas):
                return False
            row_f = formulas[f_row_idx]
            return ci < len(row_f) and str(row_f[ci]).startswith("=")

        DAY_EN = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]   # 0=Mon…6=Sun
        COL_DAY_EN = 14   # column N — ชื่อวัน (En) สำหรับสูตร IF(N6="Su",...)

        def day_en_from_datestr(ds: str) -> str:
            """แปลง 'dd/mm/YYYY_BE' เป็น 'Su','Mo',... (ปีพุทธ → คริสต์ -543)"""
            try:
                parts = ds.split("/")
                if len(parts) != 3:
                    return ""
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y > 2400:          # ปีพุทธศักราช
                    y -= 543
                from datetime import date as _date
                return DAY_EN[_date(y, m, d).weekday()]
            except Exception:
                return ""

        updates = []
        pasted  = 0

        for bi, dc in enumerate(date_cols):
            ti_col = ti_cols[bi]   if bi < len(ti_cols)   else None
            to_col = to_cols[bi]   if bi < len(to_cols)   else None
            nt_col = note_cols[bi] if bi < len(note_cols) else None
            dn_col = day_cols[bi]  if bi < len(day_cols)  else None

            for key, r_num in date_row_map.items():
                if key[0] != bi:
                    continue
                date_key = key[1]
                if date_key not in att_map:
                    continue
                att = att_map[date_key]

                if ti_col and not is_formula(ti_col, r_num):
                    updates.append({"range": f"{col_letter(ti_col)}{r_num}",
                                    "values": [[att["เวลาเข้า"]]]})
                if to_col and not is_formula(to_col, r_num):
                    updates.append({"range": f"{col_letter(to_col)}{r_num}",
                                    "values": [[att["เวลาออก"]]]})
                if nt_col and not is_formula(nt_col, r_num):
                    updates.append({"range": f"{col_letter(nt_col)}{r_num}",
                                    "values": [[att["หมายเหตุ"]]]})
                if dn_col and not is_formula(dn_col, r_num):
                    updates.append({"range": f"{col_letter(dn_col)}{r_num}",
                                    "values": [[att["วัน"]]]})

                # เขียนชื่อวัน (En) ลง column N เสมอ (Mo/Tu/We/Th/Fr/Sa/Su)
                day_abbr = day_en_from_datestr(date_key)
                if day_abbr and not is_formula(COL_DAY_EN, r_num):
                    updates.append({"range": f"N{r_num}",
                                    "values": [[day_abbr]]})

                pasted += 1

        if updates:
            # USER_ENTERED ให้ Sheets parse "08:47" เป็น serial time จริงๆ
            # (ไม่ใช่ text ธรรมดา) เพื่อให้สูตรเปรียบเทียบเวลาได้ถูกต้อง
            _sheets_retry(ws.batch_update, updates,
                          value_input_option="USER_ENTERED")
        return {"ok": True, "msg": f"วางข้อมูล {pasted} วัน เรียบร้อยค่ะ"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def _export_history(ss: gspread.Spreadsheet, sheet_name: str,
                    month_yr: str,
                    ss_db: gspread.Spreadsheet | None = None) -> dict:
    """
    บันทึกประวัติเงินเดือน — copy คอลัมน์ M–Y (col 13–25) จากชีตพนักงาน
    แล้ววางต่อในชีตเดิม โดยเว้น 1 คอลัมน์จากข้อมูลล่าสุด
    """
    COL_HIST_START = 13   # M (1-indexed)
    COL_HIST_END   = 25   # Y (1-indexed)

    try:
        NUM_ROWS = 100   # copy row 1–100 ทั้งหมด (รวม "มาสาย" และ section ด้านล่าง)
        width    = COL_HIST_END - COL_HIST_START + 1

        ws           = ss.worksheet(sheet_name)
        src_sheet_id = ws.id

        # ── หา last used column (retry กัน 429) ───────────────────────
        all_vals = _sheets_retry(ws.get_all_values)
        last_col = COL_HIST_END
        for row in all_vals:
            for ci in range(len(row) - 1, -1, -1):
                if row[ci].strip():
                    if ci + 1 > last_col:
                        last_col = ci + 1
                    break

        paste_start = last_col + 2          # เว้น 1 คอลัมน์
        paste_end   = paste_start + width   # exclusive

        # ── ขยาย Sheet ถ้าจำเป็น ──────────────────────────────────
        if paste_end > ws.col_count:
            _sheets_retry(ws.resize,
                          rows=max(ws.row_count, NUM_ROWS),
                          cols=paste_end + 10)

        src_range = {
            "sheetId"         : src_sheet_id,
            "startRowIndex"   : 0,
            "endRowIndex"     : NUM_ROWS,
            "startColumnIndex": COL_HIST_START - 1,   # 0-indexed
            "endColumnIndex"  : COL_HIST_END,          # 0-indexed exclusive
        }
        dst_range = {
            "sheetId"         : src_sheet_id,
            "startRowIndex"   : 0,
            "endRowIndex"     : NUM_ROWS,
            "startColumnIndex": paste_start - 1,
            "endColumnIndex"  : paste_end - 1,
        }

        # ── Step 1: อ่านค่า source ก่อน paste ────────────────────────
        src_a1 = (f"{col_letter(COL_HIST_START)}1:"
                  f"{col_letter(COL_HIST_END)}{NUM_ROWS}")
        src_values = _sheets_retry(ws.get, src_a1,
                                   value_render_option="FORMATTED_VALUE")
        time.sleep(1)   # หน่วงเล็กน้อยก่อน batch_update

        # ── Step 2: PASTE_NORMAL — copy ทุกอย่างรวมสี+เส้น+รูป ──────────
        body = {"requests": [{
            "copyPaste": {
                "source"          : src_range,
                "destination"     : dst_range,
                "pasteType"       : "PASTE_NORMAL",
                "pasteOrientation": "NORMAL"
            }
        }]}
        _sheets_retry(ss.batch_update, body)
        time.sleep(1)   # หน่วงก่อน write values

        # ── Step 3: write source values ทับ destination ───────────────
        dst_a1 = (f"{col_letter(paste_start)}1:"
                  f"{col_letter(paste_end - 1)}{NUM_ROWS}")
        if src_values:
            _sheets_retry(ws.update, dst_a1, src_values,
                          value_input_option="RAW")

        return {"ok": True,
                "msg": (f"บันทึกประวัติ {month_yr} ที่คอลัมน์ "
                        f"{col_letter(paste_start)} ในชีต \"{sheet_name}\" เรียบร้อยค่ะ")}
    except Exception as e:
        return {"ok": False, "msg": str(e)}
