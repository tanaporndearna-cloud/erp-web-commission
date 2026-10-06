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
                wait = 5 * (2 ** attempt)   # 5 → 10 → 20 → 40 → 80 วินาที
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
    "ATT_EMP_ID"  : 1,   # รหัสพนักงาน
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
MONTH_TH_FULL = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน",
                  "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม",
                  "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]


# ── Helpers ─────────────────────────────────────────────────────

def col_letter(n: int) -> str:
    """แปลงเลขคอลัมน์ (1-indexed) เป็นตัวอักษร เช่น 13 → M"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _freeze_formulas_to_values(ss, sheet_pairs, col_start: int, num_rows: int, chunk: int = 50):
    """
    แปลงสูตรใน history area → ค่านิ่ง โดยไม่กระทบรูปที่อยู่ใน cell

    วิธีการ (segment-based):
      - อ่าน FORMULA + FORMATTED_VALUE ของแต่ละ sheet
      - แต่ละคอลัมน์: รวม row ที่เป็นสูตรติดกัน → write เป็น 1 range
      - row ที่ไม่ใช่สูตร (ว่าง / มีรูป) → ข้ามทั้งหมด → รูปปลอดภัย

    sheet_pairs: list of (sheet_name, ws_title)
                 sheet_name = ชื่อสำหรับ lookup (อาจต่างจาก ws_title ถ้า escape)
                 ws_title   = ชื่อจริงของ Google Sheet tab
    col_start:  คอลัมน์แรกของ history (1-indexed)  เช่น 26 = Z
    num_rows:   จำนวนแถวสูงสุดที่อ่าน

    Returns: dict[sheet_name] -> คอลัมน์สุดท้ายที่มีข้อมูล (1-indexed)
    """
    def _esc(sh: str) -> str:
        return sh.replace("'", "''")

    COL_END_MAX = col_start + 574   # รองรับประวัติ ~575 คอลัมน์ (~41 เดือน × 14 col)

    fz_ranges = [
        f"'{_esc(wt)}'!{col_letter(col_start)}1:{col_letter(COL_END_MAX)}{num_rows}"
        for _, wt in sheet_pairs
    ]

    last_cols:    dict[str, int]        = {}
    sheet_writes: dict[str, list[dict]] = {}   # ws_title → write ops (แยกต่อ sheet)
    failed_reads: list[str]             = []

    for fi in range(0, len(fz_ranges), chunk):
        cr = fz_ranges[fi:fi + chunk]
        cp = sheet_pairs[fi:fi + chunk]
        try:
            rf = _sheets_retry(ss.values_batch_get, cr,
                               params={"valueRenderOption": "FORMULA"})
            rd = _sheets_retry(ss.values_batch_get, cr,
                               params={"valueRenderOption": "FORMATTED_VALUE"})
            for (sname, ws_title), vr_f, vr_d in zip(
                cp,
                rf.get("valueRanges", []),
                rd.get("valueRanges", []),
            ):
                rows_f = vr_f.get("values", [])
                rows_d = vr_d.get("values", [])
                if not rows_f:
                    continue
                max_w = max((len(r) for r in rows_f), default=0)
                # FIX Bug2: ใช้ rows_d (FORMATTED_VALUE) สำหรับ max_w_data แทน rows_f (FORMULA)
                # เพราะ FORMULA mode จะ return formula text เช่น '=""' ซึ่ง non-empty
                # ทำให้ len(row) บวมจาก trailing formula cells ที่ return ""
                # FORMATTED_VALUE mode trim trailing empty display values → len() ถูกต้อง
                data_rows_d = rows_d[5:]   # FORMATTED_VALUE rows สำหรับ data rows (แถว 6+)
                max_w_data = max((len(r) for r in data_rows_d), default=0)
                if max_w_data:
                    last_cols[sname] = col_start + max_w_data - 1
                elif max_w:
                    # มีแค่ rows 1-5 — หา rightmost formula cell เพื่อหลีกเลี่ยง static metadata
                    max_fml_w = 0
                    for ri, row_f in enumerate(rows_f):
                        for ci, cf in enumerate(row_f):
                            if isinstance(cf, str) and cf.startswith("="):
                                max_fml_w = max(max_fml_w, ci + 1)
                    if max_fml_w:
                        last_cols[sname] = col_start + max_fml_w - 1

                for ci in range(max_w):
                    # เก็บเฉพาะ row ที่เป็นสูตร (= หัวข้อ '=')
                    fml_rows: list[tuple[int, str]] = []
                    for ri, row_f in enumerate(rows_f):
                        cf = row_f[ci] if ci < len(row_f) else ""
                        if isinstance(cf, str) and cf.startswith("="):
                            dv = (rows_d[ri][ci]
                                  if ri < len(rows_d) and ci < len(rows_d[ri])
                                  else "")
                            fml_rows.append((ri, dv))

                    if not fml_rows:
                        continue

                    # สร้าง segment ของ row ต่อเนื่อง → 1 range ต่อ segment
                    segs: list[tuple[int, list[str]]] = []
                    s_ri   = fml_rows[0][0]
                    s_vals = [fml_rows[0][1]]
                    for (p_ri, _), (n_ri, n_dv) in zip(fml_rows, fml_rows[1:]):
                        if n_ri == p_ri + 1:
                            s_vals.append(n_dv)
                        else:
                            segs.append((s_ri, s_vals))
                            s_ri, s_vals = n_ri, [n_dv]
                    segs.append((s_ri, s_vals))

                    cl = col_letter(col_start + ci)
                    for seg_ri, seg_vs in segs:
                        sheet_writes.setdefault(ws_title, []).append({
                            "range" : (f"'{_esc(ws_title)}'!"
                                       f"{cl}{seg_ri + 1}:{cl}{seg_ri + len(seg_vs)}"),
                            "values": [[v] for v in seg_vs],
                        })
        except Exception:
            failed_reads.extend(sn for sn, _ in cp)

    # ── Write แยกต่อ sheet (sheet ใด fail ก็ไม่กระทบ sheet อื่น) ──
    failed_writes: list[str] = []
    for ws_title, sw in sheet_writes.items():
        for fi in range(0, len(sw), chunk):
            try:
                _sheets_retry(ss.values_batch_update, {
                    "valueInputOption": "RAW",
                    "data"            : sw[fi:fi + chunk],
                })
            except Exception as e:
                failed_writes.append(f"{ws_title}: {e}")

    # แนบ failed info ไว้ใน last_cols เพื่อให้ caller ดูได้ (ผ่าน attribute พิเศษ)
    last_cols["__failed_reads__"]  = failed_reads   # type: ignore[assignment]
    last_cols["__failed_writes__"] = failed_writes  # type: ignore[assignment]

    return last_cols


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


def _day_en_from_datestr(ds: str) -> str:
    """แปลง 'dd/mm/YYYY_BE' → 'Su'/'Mo'/... (ปีพุทธ -543)"""
    try:
        parts = ds.split("/")
        if len(parts) != 3:
            return ""
        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
        if y > 2400:
            y -= 543
        from datetime import date as _date
        DAY_EN = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]   # 0=Mon…6=Sun
        return DAY_EN[_date(y, m, d).weekday()]
    except Exception:
        return ""


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
        ss = _sheets_retry(gc.open_by_key, sheet_id)
    except Exception as e:
        st.error(f"เปิด Spreadsheet ไม่ได้: {e}")
        return

    # เปิด history DB ถ้ามี
    ss_db = None
    if history_db_id:
        try:
            ss_db = _sheets_retry(gc.open_by_key, history_db_id)
        except Exception:
            ss_db = None

    visible_sheets = [ws.title for ws in ss.worksheets()]

    # ══════════════════════════════════════════════════════════════
    # คิดเงินเดือนพนักงาน — Import เวลาเข้า-ออกงาน
    # ══════════════════════════════════════════════════════════════
    with st.expander("🧾  คิดเงินเดือนพนักงาน (Import เวลาเข้า-ออกงาน)", expanded=True):
        st.caption("อัปโหลดไฟล์เวลา → ล้างเก่า → สร้างวันที่ → วางเวลา → บันทึกประวัติ (ประมวลผลทุก Sheet พนักงานอัตโนมัติ)")

        att_file = st.file_uploader("📂 ไฟล์เวลาเข้า-ออกงาน (.xlsx / .xls)",
                                    type=["xlsx", "xls"], key="att_file")

        if att_file:
            att_df = _read_attendance_file(att_file)
            if att_df is not None and not att_df.empty:
                auto_month, auto_year = detect_period_from_df(att_df)

                # จับคู่ sheet กับรหัสพนักงาน
                # ชื่อ sheet รูปแบบ "99043_ณี" → prefix ก่อน "_" คือรหัสพนักงาน
                EXCLUDE_SHEETS = {CFG["HISTORY_SHEET"]}
                emp_ids_in_file = att_df["รหัสพนักงาน"].dropna().unique().tolist()
                emp_ids_in_file = [str(e) for e in emp_ids_in_file if str(e).strip()]

                # สร้าง map: รหัสพนักงาน → ชื่อ sheet
                id_to_sheet = {}
                for sh in visible_sheets:
                    if sh in EXCLUDE_SHEETS:
                        continue
                    prefix = sh.split("_")[0].strip()
                    if prefix in emp_ids_in_file:
                        id_to_sheet[prefix] = sh

                # แจ้งถ้ามีรหัสในไฟล์ที่ไม่มี sheet
                missing = [e for e in emp_ids_in_file if e not in id_to_sheet]
                if missing:
                    st.warning(f"⚠️ ไม่พบ Sheet สำหรับรหัส: {', '.join(missing)} — จะข้ามค่ะ")

                matched_pairs = list(id_to_sheet.items())   # [(emp_id, sheet_name), ...]

                def _run_sheets(pairs, att_df, do_history=False):
                    success_count = 0
                    fail_count    = 0
                    errors        = []
                    progress_bar  = st.progress(0)
                    status_text   = st.empty()

                    if do_history:
                        # ── FAST BATCH MODE สำหรับ history ──────────────────
                        month_yr    = f"{auto_month}/{auto_year}"
                        sheet_names = [sh_pay for _, sh_pay in pairs]
                        batch_results = _export_history_batch(
                            ss, sheet_names, month_yr, ss_db,
                            status_text=status_text,
                            progress_bar=progress_bar,
                        )
                        for idx2, r in enumerate(batch_results):
                            progress_bar.progress((idx2 + 1) / len(pairs))
                            if not r["ok"]:
                                errors.append(f"{r['sheet']}: บันทึกประวัติไม่ได้ — {r['msg']}")
                                fail_count += 1
                            else:
                                success_count += 1
                        status_text.empty()
                        if fail_count == 0:
                            st.success(f"✅ บันทึกประวัติครบ {success_count} Sheet ค่ะ")
                        else:
                            st.warning(f"⚠️ เสร็จ {success_count} / ล้มเหลว {fail_count} Sheet")
                        if errors:
                            with st.expander("❌ รายการที่มีปัญหา"):
                                for err in errors:
                                    st.text(err)
                        return

                    # ── Freeze ประวัติเก่า → ค่านิ่ง (segment-based, รูปปลอดภัย) ──
                    try:
                        _all_ws2 = _sheets_retry(ss.worksheets)
                        _ws2_map = {w.title: w for w in _all_ws2}
                        _fz2_pairs = [(sh, sh) for _, sh in pairs if sh in _ws2_map]
                        _freeze_formulas_to_values(ss, _fz2_pairs, col_start=26, num_rows=100)
                    except Exception:
                        pass   # freeze ไม่ได้ก็ข้าม ไม่หยุดการคิดเงินเดือน

                    for idx, (emp_id, sh_pay) in enumerate(pairs):
                        label = "ประมวลผล"
                        status_text.info(
                            f"⏳ กำลัง{label} Sheet **{sh_pay}** ({idx+1}/{len(pairs)})..."
                        )
                        try:
                            if False:   # placeholder — history ถูกจัดการข้างบนแล้ว
                                pass
                            else:
                                emp_rows = att_df[att_df["รหัสพนักงาน"] == emp_id].to_dict("records")

                                r1 = _clear_attendance(ss, sh_pay)
                                if not r1["ok"]:
                                    errors.append(f"{sh_pay}: ล้างไม่ได้ — {r1['msg']}")
                                    fail_count += 1
                                    progress_bar.progress((idx + 1) / len(pairs))
                                    continue

                                # เปิด ws + headers ครั้งเดียว แล้วส่งต่อให้ทั้ง 2 ฟังก์ชัน (ลด API call)
                                try:
                                    shared_ws      = _sheets_retry(ss.worksheet, sh_pay)
                                    shared_headers = _sheets_retry(shared_ws.row_values, CFG["HEADER_ROW"])
                                except Exception as e_ws:
                                    errors.append(f"{sh_pay}: เปิด worksheet ไม่ได้ — {e_ws}")
                                    fail_count += 1
                                    progress_bar.progress((idx + 1) / len(pairs))
                                    continue

                                r2 = _generate_dates(ss, sh_pay, auto_month, auto_year,
                                                     ws=shared_ws, headers=shared_headers)
                                if not r2["ok"]:
                                    errors.append(f"{sh_pay}: สร้างวันที่ไม่ได้ — {r2['msg']}")
                                    fail_count += 1
                                    progress_bar.progress((idx + 1) / len(pairs))
                                    continue

                                r3 = _import_attendance(ss, sh_pay, emp_rows,
                                                        ws=shared_ws, headers=shared_headers)
                                if not r3["ok"]:
                                    errors.append(f"{sh_pay}: วางเวลาไม่ได้ — {r3['msg']}")
                                    fail_count += 1
                                    progress_bar.progress((idx + 1) / len(pairs))
                                    continue

                                success_count += 1

                        except Exception as e:
                            errors.append(f"{sh_pay}: ข้อผิดพลาด — {str(e)}")
                            fail_count += 1

                        progress_bar.progress((idx + 1) / len(pairs))
                        if idx < len(pairs) - 1:
                            # cooldown ทุก 10 sheet เพื่อไม่ให้ 429
                            if (idx + 1) % 10 == 0:
                                status_text.info(
                                    f"⏸ พัก 30 วินาที (cooldown หลังครบ {idx+1} Sheet)..."
                                )
                                time.sleep(30)
                            else:
                                time.sleep(6)

                    status_text.empty()
                    if fail_count == 0:
                        st.success(f"✅ เสร็จสิ้น! ประมวลผลครบ {success_count} Sheet ค่ะ")
                    else:
                        st.warning(f"⚠️ เสร็จ {success_count} Sheet / ล้มเหลว {fail_count} Sheet")
                    if errors:
                        with st.expander("❌ รายการที่มีปัญหา"):
                            for err in errors:
                                st.text(err)

                use_fast = st.toggle(
                    "⚡ Fast Mode (batch API — ทุก Sheet พร้อมกัน ~2 นาที แทน ~45 นาที)",
                    value=True, key="toggle_fast"
                )

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("▶ เริ่มคิดเงินเดือน (ล้างเก่า → สร้างวันที่ → วางเวลา)",
                                 type="primary", key="btn_payroll", use_container_width=True):
                        if use_fast:
                            # โหลดวันหยุดตามประเพณีก่อนส่งให้ fast mode
                            _hmap = {}
                            try:
                                _hws  = ss.worksheet("วันหยุดประเพณี")
                                _hrows = _hws.get_all_records()
                                for _hr in _hrows:
                                    _d = int(_hr.get("วันที่", 0) or 0)
                                    _m = int(_hr.get("เดือน",  0) or 0)
                                    _y = int(_hr.get("ปี (พ.ศ.)", 0) or 0)
                                    if _d and _m and _y:
                                        _dk = f"{str(_d).zfill(2)}/{str(_m).zfill(2)}/{_y}"
                                        _hmap[_dk] = True
                            except Exception:
                                pass
                            _run_sheets_fast(matched_pairs, att_df,
                                             auto_month, auto_year, ss,
                                             holidays_map=_hmap)
                            # ── ล้างวันหยุดออกจาก sheet อัตโนมัติหลังรันสำเร็จ ──
                            try:
                                _hws2 = ss.worksheet("วันหยุดประเพณี")
                                _hws2.clear()
                                _hws2.append_row(["ปี (พ.ศ.)", "เดือน", "วันที่"])
                                _hws2.freeze(rows=1)
                                # ล้าง cache ด้วย
                                st.session_state.pop("_hol_data", None)
                                st.session_state.pop("_hol_ws",   None)
                            except Exception:
                                pass
                        else:
                            _run_sheets(matched_pairs, att_df, do_history=False)

                with col2:
                    hist_label = "📚 บันทึกประวัติ (เทมเพลต" + (" + DB" if ss_db else "") + ")"
                    if st.button(hist_label, key="btn_history", use_container_width=True):
                        _run_sheets(matched_pairs, att_df, do_history=True)

            else:
                st.warning("⚠️ อ่านไฟล์ไม่ได้ หรือไม่พบข้อมูล")

        # ── Freeze ประวัติ (แสดงเสมอ ไม่ต้องอัปโหลดไฟล์ก่อน) ──────────
        st.divider()
        st.caption("🔒 Freeze ประวัติ — แปลงสูตรในคอลัมน์ประวัติ (Z เป็นต้นไป) ให้เป็นค่านิ่ง กดได้ทุกเวลา ไม่ต้องรันเงินเดือนใหม่")
        if st.button("🔒 Freeze ประวัติ → ค่านิ่ง (กดได้ทุกเวลา)",
                     key="btn_freeze_only", use_container_width=True):
            # ดึง sheet list สดจาก Google Sheets ตอนกดปุ่ม (ไม่ใช้ visible_sheets ที่ cache ไว้ตอน page load)
            # เพื่อให้นับชีทถูกต้องแม้มีการลบ/เพิ่มชีทหลัง page load
            try:
                _fresh_sheets = [ws.title for ws in _sheets_retry(ss.worksheets)]
            except Exception:
                _fresh_sheets = visible_sheets   # fallback ถ้า API ล้มเหลว
            _fz_all_pairs = [
                (sh, sh) for sh in _fresh_sheets
                if re.match(r"^\d+_", sh)
            ]
            _freeze_history_standalone(_fz_all_pairs, ss)

    # ══════════════════════════════════════════════════════════════
    # วันหยุดตามประเพณี — บันทึกวันหยุดพิเศษประจำเดือน
    # ══════════════════════════════════════════════════════════════
    with st.expander("📅  วันหยุดตามประเพณี", expanded=False):
        st.caption("บันทึกวันหยุดตามประเพณีประจำเดือน เช่น วันเฉลิมพระชนมพรรษา ข้อมูลเก็บในชีต 'วันหยุดประเพณี'")

        THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน",
                       "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม",
                       "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

        import datetime as _dt
        _now = _dt.date.today()
        _def_month = _now.month
        _def_year  = _now.year + 543   # CE → BE

        # ดึงค่าเดือน/ปีจาก session_state (หรือ default) เพื่อใช้โหลดข้อมูลก่อน render widget
        h_month = int(st.session_state.get("h_month", _def_month))
        h_year  = int(st.session_state.get("h_year",  _def_year))

        # ── อ่านวันหยุดจาก Google Sheet ──
        HOLIDAY_SHEET   = "วันหยุดประเพณี"
        _HOL_WS_KEY     = "_hol_ws"       # cache worksheet object
        _HOL_DATA_KEY   = "_hol_data"     # cache records list

        def _get_holiday_ws(ss: gspread.Spreadsheet) -> gspread.Worksheet:
            """คืน worksheet (cache ไว้ใน session_state — ss.worksheet() 1 ครั้งต่อ session)"""
            if _HOL_WS_KEY not in st.session_state:
                try:
                    st.session_state[_HOL_WS_KEY] = ss.worksheet(HOLIDAY_SHEET)
                except gspread.exceptions.WorksheetNotFound:
                    ws = ss.add_worksheet(title=HOLIDAY_SHEET, rows=200, cols=4)
                    ws.append_row(["ปี (พ.ศ.)", "เดือน", "วันที่"])
                    ws.freeze(rows=1)
                    st.session_state[_HOL_WS_KEY] = ws
            return st.session_state[_HOL_WS_KEY]

        def _load_all_holidays(ss):
            """โหลดจาก sheet ครั้งเดียวต่อ session (cache ใน session_state)"""
            if _HOL_DATA_KEY not in st.session_state:
                try:
                    ws = _get_holiday_ws(ss)
                    st.session_state[_HOL_DATA_KEY] = ws.get_all_records()
                except Exception as e:
                    st.error(f"โหลดวันหยุดไม่ได้: {e}")
                    return []
            return st.session_state[_HOL_DATA_KEY]

        def _invalidate_hol_cache():
            st.session_state.pop(_HOL_DATA_KEY, None)

        def _add_holidays_batch(ss, year, month, days: list[int]):
            """เพิ่มหลายวันพร้อมกัน — ใช้ cache เช็ค dup, append ทีเดียว"""
            try:
                ws   = _get_holiday_ws(ss)
                rows = _load_all_holidays(ss)          # ใช้ cache
                existing = {
                    int(r.get("วันที่", 0))
                    for r in rows
                    if str(r.get("ปี (พ.ศ.)", "")) == str(year)
                    and str(r.get("เดือน", "")) == str(month)
                }
                to_add  = [d for d in days if d not in existing]
                skipped = [d for d in days if d in existing]
                if to_add:
                    ws.append_rows([[year, month, d] for d in sorted(to_add)],
                                   value_input_option="RAW")
                    _invalidate_hol_cache()            # ล้าง cache ให้ fetch ใหม่รอบหน้า
                return to_add, skipped
            except Exception as e:
                raise RuntimeError(str(e))

        def _rewrite_sheet(ws, keep_rows: list):
            """Clear + เขียน keep_rows กลับ — เร็วกว่า delete_rows ทีละแถว (3 API calls คงที่)"""
            header = ["ปี (พ.ศ.)", "เดือน", "วันที่"]
            ws.clear()
            ws.freeze(rows=0)                          # unfreeze ก่อน (freeze หลัง update)
            if keep_rows:
                ws.update([header] + keep_rows, value_input_option="RAW")
            else:
                ws.update([header], value_input_option="RAW")
            ws.freeze(rows=1)

        def _delete_holiday(ss, year, month, day):
            try:
                ws       = _get_holiday_ws(ss)
                all_vals = ws.get_all_values()         # row 0 = header
                before   = len(all_vals) - 1           # จำนวน data rows เดิม
                keep     = [
                    row for i, row in enumerate(all_vals)
                    if i > 0 and not (
                        str(row[0]) == str(year)
                        and str(row[1]) == str(month)
                        and str(row[2]) == str(day)
                    )
                ]
                if len(keep) == before:
                    return False                        # ไม่เจอ row ที่จะลบ
                _rewrite_sheet(ws, keep)
                _invalidate_hol_cache()
                return True
            except Exception as e:
                st.error(f"ลบไม่ได้: {e}")
                return False

        def _delete_month_holidays(ss, year, month):
            try:
                ws       = _get_holiday_ws(ss)
                all_vals = ws.get_all_values()
                before   = len(all_vals) - 1
                keep     = [
                    row for i, row in enumerate(all_vals)
                    if i > 0 and not (
                        str(row[0]) == str(year)
                        and str(row[1]) == str(month)
                    )
                ]
                deleted = before - len(keep)
                if deleted == 0:
                    return 0
                _rewrite_sheet(ws, keep)
                _invalidate_hol_cache()
                return deleted
            except Exception as e:
                st.error(f"ลบทั้งเดือนไม่ได้: {e}")
                return 0

        # โหลดรายการ **ทั้งหมด** (ไม่กรองเดือน)
        h_data_all = _load_all_holidays(ss)
        h_data_sorted = sorted(
            h_data_all,
            key=lambda r: (int(r.get("ปี (พ.ศ.)", 0)),
                           int(r.get("เดือน", 0)),
                           int(r.get("วันที่", 0)))
        )

        if h_data_sorted:
            top_l, top_r = st.columns([5, 2])
            top_l.markdown(f"**วันหยุดตามประเพณีทั้งหมด — {len(h_data_sorted)} วัน**")
            if top_r.button("🗑️ ล้างทั้งหมด", key="del_all_hol", use_container_width=True):
                try:
                    _ws_all = _get_holiday_ws(ss)
                    _ws_all.clear()
                    _ws_all.append_row(["ปี (พ.ศ.)", "เดือน", "วันที่"])
                    _ws_all.freeze(rows=1)
                    _invalidate_hol_cache()
                    st.success("ล้างวันหยุดทั้งหมดแล้วค่ะ")
                    st.rerun()
                except Exception as _e:
                    st.error(f"ล้างไม่ได้: {_e}")
            # จัดกลุ่มตามปี-เดือน
            from itertools import groupby
            for (yr, mo), grp in groupby(
                h_data_sorted,
                key=lambda r: (int(r.get("ปี (พ.ศ.)", 0)), int(r.get("เดือน", 0)))
            ):
                month_name = THAI_MONTHS[mo] if 1 <= mo <= 12 else str(mo)
                grp_list = list(grp)
                hdr_c, del_c = st.columns([6, 2])
                hdr_c.markdown(f"📅 **{month_name} {yr}** ({len(grp_list)} วัน)")
                if del_c.button(f"🗑️ ลบทั้งเดือน", key=f"del_month_{yr}_{mo}",
                                 use_container_width=True):
                    n = _delete_month_holidays(ss, yr, mo)
                    if n:
                        st.success(f"ลบ {n} วันของ{month_name} {yr} แล้วค่ะ")
                        st.rerun()
                for row in grp_list:
                    day_val   = row.get("วันที่", "")
                    row_yr    = int(row.get("ปี (พ.ศ.)", yr))
                    row_mo    = int(row.get("เดือน", mo))
                    month_pad = str(row_mo).zfill(2)
                    col_d, col_x = st.columns([5, 1])
                    col_d.markdown(
                        f"<span style='background:#dc3545;color:#fff;padding:3px 10px;"
                        f"border-radius:6px;font-weight:bold;font-size:0.9em'>"
                        f"{str(day_val).zfill(2)}/{month_pad}/{str(row_yr)[2:]}</span>",
                        unsafe_allow_html=True
                    )
                    if col_x.button("🗑️", key=f"del_hol_{row_yr}_{row_mo}_{day_val}",
                                      help="ลบวันหยุดนี้"):
                        ok = _delete_holiday(ss, row_yr, row_mo, int(day_val))
                        if ok:
                            st.success(f"ลบวันที่ {day_val}/{month_pad} แล้วค่ะ")
                            st.rerun()
        else:
            st.info("ยังไม่มีวันหยุดตามประเพณีค่ะ")

        st.divider()

        # ── ฟอร์มเพิ่มวันหยุด — วันที่ (หลายวัน) / เดือน / ปี / ปุ่ม บรรทัดเดียว ──
        st.markdown("**➕ เพิ่มวันหยุด**")
        f1, f2, f3, f4 = st.columns([2, 2, 1, 1])
        with f1:
            new_days_str = st.text_input("วันที่ (คั่นด้วยจุลภาค เช่น 5, 17, 18)",
                                         placeholder="5, 17, 18",
                                         key="new_hdays")
        with f2:
            h_month = st.selectbox("เดือน", list(range(1, 13)),
                                   index=h_month - 1,
                                   format_func=lambda m: THAI_MONTHS[m],
                                   key="h_month")
        with f3:
            h_year = st.number_input("ปี (พ.ศ.)", min_value=2560, max_value=2599,
                                     value=h_year, step=1, key="h_year")
        with f4:
            st.write("")
            st.write("")
            if st.button("เพิ่ม", key="btn_add_hol", use_container_width=True):
                raw_days = [d.strip() for d in new_days_str.replace("،", ",").split(",") if d.strip()]
                if not raw_days:
                    st.warning("กรุณากรอกวันที่ค่ะ")
                else:
                    invalid = [d for d in raw_days if not d.isdigit() or not (1 <= int(d) <= 31)]
                    valid   = [int(d) for d in raw_days if d.isdigit() and 1 <= int(d) <= 31]
                    if invalid:
                        st.error(f"❌ ค่าไม่ถูกต้อง: {', '.join(invalid)}")
                    if valid:
                        try:
                            added, skipped = _add_holidays_batch(ss, int(h_year), int(h_month), valid)
                            if added:
                                st.success(f"✅ เพิ่มวันที่ {', '.join(str(d) for d in added)} {THAI_MONTHS[int(h_month)]} {int(h_year)} สำเร็จค่ะ")
                            if skipped:
                                st.warning(f"⚠️ วันที่ {', '.join(str(d) for d in skipped)} มีอยู่แล้ว ข้ามไปค่ะ")
                            if added:
                                st.rerun()
                        except RuntimeError as e:
                            st.error(f"❌ {e}")


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
    เขียน "" ทับทุก cell ในคอลัมน์ N-S (14-19) + X-Y (24-25) แถว 6-36
    ใช้ values_batch_update แทน batch_clear เพื่อให้ work แม้มี merged cells
    """
    ROW_START = 6
    ROW_END   = 36

    try:
        ws = _sheets_retry(ss.worksheet, sheet_name)

        num_rows = ROW_END - ROW_START + 1   # 31 แถว
        # สร้างข้อมูลว่าง: N-S = 6 คอลัมน์, X-Y = 2 คอลัมน์
        empty_ns = [[""] * 6 for _ in range(num_rows)]
        empty_xy = [[""] * 2 for _ in range(num_rows)]

        # เขียน "" ทับทุก cell แทน batch_clear
        # (ทำงานได้แม้มี merged cells ซึ่ง batch_clear บางครั้งข้ามไป)
        updates = [
            {"range": f"N{ROW_START}:S{ROW_END}", "values": empty_ns},
            {"range": f"X{ROW_START}:Y{ROW_END}", "values": empty_xy},
        ]
        _sheets_retry(ws.batch_update, updates, value_input_option="RAW")
        return {"ok": True, "msg": f"ล้าง N-S และ X-Y แถว {ROW_START}-{ROW_END} เรียบร้อยค่ะ"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def _freeze_history_standalone(pairs, ss):
    """
    🔒 Freeze ประวัติ → ค่านิ่ง (standalone — กดได้ทุกเวลา ไม่ต้องรันเงินเดือนใหม่)
    ใช้ segment-based freeze → รูปที่อยู่ใน cell ปลอดภัย (ไม่ถูกเขียนทับ)
    """
    FZ_COL_START = 26    # col Z
    FZ_ROWS      = 100

    status_text  = st.empty()
    progress_bar = st.progress(0.0)

    # ── Step 1: เตรียม sheet list ────────────────────────────────
    status_text.info("🔍 [1/3] กำลังโหลด Sheet list...")
    try:
        all_ws = _sheets_retry(ss.worksheets)
        ws_map = {w.title: w for w in all_ws}
    except Exception as e:
        st.error(f"❌ เปิด Spreadsheet ไม่ได้: {e}")
        progress_bar.empty(); status_text.empty()
        return

    valid_pairs = [(sn, sh) for sn, sh in pairs if sh in ws_map]
    if not valid_pairs:
        st.warning("⚠️ ไม่พบ Sheet ที่ตรงกันเลยค่ะ")
        progress_bar.empty(); status_text.empty()
        return

    # valid_pairs ใช้ (emp_id, sheet_title) → สร้าง (sheet_title, sheet_title) สำหรับ helper
    sheet_pairs = [(sh, sh) for _, sh in valid_pairs]

    status_text.info(f"📖 [2/3] กำลังอ่านและ freeze ประวัติ {len(sheet_pairs)} Sheet...")
    progress_bar.progress(0.1)

    result = _freeze_formulas_to_values(ss, sheet_pairs, FZ_COL_START, FZ_ROWS)

    progress_bar.progress(1.0)
    status_text.empty()
    progress_bar.empty()

    # แยก error info ออกจาก last_cols ก่อนนับ
    failed_reads  = result.pop("__failed_reads__",  [])
    failed_writes = result.pop("__failed_writes__", [])

    if result:
        st.success(f"✅ Freeze เสร็จแล้วค่ะ — ครอบคลุม {len(result)} Sheet รูปใน cell ปลอดภัย")
    else:
        st.info("ℹ️ ไม่พบสูตรในประวัติ (อาจ freeze ไปแล้ว หรือยังไม่มีประวัติ)")

    if failed_reads:
        st.warning(f"⚠️ อ่านข้อมูลไม่ได้ {len(failed_reads)} Sheet: {', '.join(str(x) for x in failed_reads)}")
    if failed_writes:
        st.error(f"❌ เขียนค่านิ่งไม่สำเร็จ {len(failed_writes)} รายการ:\n" +
                 "\n".join(f"• {x}" for x in failed_writes))


def _run_sheets_fast(pairs, att_df, auto_month, auto_year, ss, holidays_map=None):
    """
    ⚡ Batch mode — ประมวลผลทุก Sheet ด้วย ~5 API calls ต่อ chunk
    แทนที่ ~1000 calls + sleep แบบเดิม (~2 นาที vs ~45 นาที)

    การป้องกันความเสี่ยง:
      - แบ่ง chunk ≤33 Sheet ต่อรอบ (33×3=99 ranges ต่อ call, limit=120)
      - escape apostrophe ในชื่อ Sheet (ป้องกัน A1 notation error)
      - validate คอลัมน์ก่อนเขียน (ข้าม Sheet ที่ผิดโครงสร้าง)
      - batch write fail → fallback slow mode เฉพาะ chunk นั้น
    """
    CHUNK_SIZE  = 33   # จำนวน Sheet สูงสุดต่อ batch round (33×3=99 ranges ต่อ call, limit=120)
    st_errors   = []
    success_cnt = 0
    fallback_pairs = []   # Sheet ที่ต้อง retry ด้วย slow mode

    # ── escape apostrophe ในชื่อ Sheet สำหรับ A1 notation ───────
    def _esc(sh: str) -> str:
        return sh.replace("'", "''")

    # ── ข้อมูลวันที่รอบบัญชี ─────────────────────────────────────
    year_ce = auto_year - 543
    prev_m  = 12 if auto_month == 1 else auto_month - 1
    prev_y  = year_ce - 1 if auto_month == 1 else year_ce
    start   = date(prev_y, prev_m, 26)
    end     = date(year_ce, auto_month, 25)
    days, cur = [], start
    while cur <= end:
        days.append(cur)
        cur += timedelta(days=1)
    num_rows    = min(len(days), CFG["DATA_END"] - CFG["DATA_START"] + 1)
    DAY_EN_LIST = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]   # (weekday+1)%7
    month_label = f"{MONTH_TH_FULL[auto_month]} {auto_year}"

    # ── Phase 1: Get all worksheets ──────────────────────────────
    progress_bar = st.progress(0.0)
    status_text  = st.empty()
    status_text.info("📂 [1/5] เปิด worksheet ทั้งหมด...")

    try:
        all_ws_list = _sheets_retry(ss.worksheets)
        ws_map = {w.title: w for w in all_ws_list}
    except Exception as e:
        st.error(f"เปิด worksheets ไม่ได้: {e}")
        return

    valid_pairs = [(eid, sh) for eid, sh in pairs if sh in ws_map]
    for eid, sh in pairs:
        if sh not in ws_map:
            st_errors.append(f"{sh}: ไม่พบ sheet นี้")

    if not valid_pairs:
        st.error("ไม่พบ Sheet ที่ตรงกันเลยค่ะ")
        return

    # ── แบ่ง chunks ≤ CHUNK_SIZE ─────────────────────────────────
    chunks = [valid_pairs[i:i + CHUNK_SIZE]
              for i in range(0, len(valid_pairs), CHUNK_SIZE)]
    total_chunks = len(chunks)

    # ── Phase 1.5: Freeze ประวัติเก่า → ค่านิ่ง (segment-based, รูปปลอดภัย) ────
    status_text.info("🔒 [1.5/5] กำลัง freeze ประวัติเก่า → ค่านิ่ง...")
    _freeze_formulas_to_values(
        ss,
        [(sh, sh) for _, sh in valid_pairs],
        col_start=26,   # col Z
        num_rows=100,
    )

    GRAY    = {"red": 211/255, "green": 211/255, "blue": 211/255}
    WHITE   = {"red": 1.0,     "green": 1.0,     "blue": 1.0}
    COL_N_0 = 13    # N (0-indexed)
    COL_Y_0 = 25    # Y+1 (exclusive)

    for chunk_idx, chunk in enumerate(chunks):
        chunk_label = f"chunk {chunk_idx+1}/{total_chunks}" if total_chunks > 1 else ""
        base_prog   = chunk_idx / total_chunks
        prog_span   = 1.0 / total_chunks

        progress_bar.progress(base_prog + prog_span * 0.10)

        # ── Phase 2: Batch read (FORMATTED_VALUE) ────────────────
        status_text.info(
            f"📖 [2/5] อ่านข้อมูล {len(chunk)} Sheet {chunk_label}..."
        )
        ranges_val = []
        for _, sh in chunk:
            e = _esc(sh)
            ranges_val.append(f"'{e}'!A1:Z10")
            ranges_val.append(f"'{e}'!A5:ZZ5")
            ranges_val.append(f"'{e}'!A6:ZZ40")

        try:
            res_val    = _sheets_retry(ss.values_batch_get, ranges_val,
                                       params={"valueRenderOption": "FORMATTED_VALUE"})
            val_ranges = res_val.get("valueRanges", [])
        except Exception as e:
            st.error(f"batch read (value) ล้มเหลว [{chunk_label}]: {e}")
            fallback_pairs.extend(chunk)
            continue

        progress_bar.progress(base_prog + prog_span * 0.25)

        # ── Phase 3: Batch read (FORMULA) ────────────────────────
        status_text.info(
            f"📐 [3/5] อ่านสูตร {len(chunk)} Sheet {chunk_label}..."
        )
        ranges_fml = [f"'{_esc(sh)}'!M5:ZZ40" for _, sh in chunk]

        try:
            res_fml    = _sheets_retry(ss.values_batch_get, ranges_fml,
                                       params={"valueRenderOption": "FORMULA"})
            fml_ranges = res_fml.get("valueRanges", [])
        except Exception as e:
            st.error(f"batch read (formula) ล้มเหลว [{chunk_label}]: {e}")
            fallback_pairs.extend(chunk)
            continue

        progress_bar.progress(base_prog + prog_span * 0.40)

        # ── Phase 4: Compute all updates in memory ───────────────
        status_text.info(f"⚙️ [4/5] คำนวณ {chunk_label}...")

        all_value_upd = []
        all_fmt_reqs  = []

        for i, (eid, sh) in enumerate(chunk):
            ws       = ws_map[sh]
            sheet_id = ws.id

            top_data = val_ranges[i * 3 + 0].get("values", [])
            hdr_data = val_ranges[i * 3 + 1].get("values", [])
            dat_data = val_ranges[i * 3 + 2].get("values", [])
            fml_data = fml_ranges[i].get("values", [])

            headers   = hdr_data[0] if hdr_data else []
            cols      = find_col_indices(headers,
                                         CFG["HDR_DATE"], CFG["HDR_DAY_NAME"],
                                         CFG["HDR_TIME_IN"], CFG["HDR_TIME_OUT"],
                                         CFG["HDR_NOTE"])
            # ── กรองเฉพาะ column ใน template (A-Y = col 1-25) ───────────
            # row 5 อ่านจาก A5:ZZ5 ดังนั้นได้ header ทั้ง template + history block
            # ต้องตัด col > 25 ออก เพื่อไม่ให้เขียนทับ history block เก่า
            _MAX_TMPL_COL = 25
            date_cols     = [c for c in cols[CFG["HDR_DATE"]]     if c <= _MAX_TMPL_COL]
            day_name_cols = [c for c in cols[CFG["HDR_DAY_NAME"]] if c <= _MAX_TMPL_COL]
            ti_cols       = [c for c in cols[CFG["HDR_TIME_IN"]]  if c <= _MAX_TMPL_COL]
            to_cols       = [c for c in cols[CFG["HDR_TIME_OUT"]] if c <= _MAX_TMPL_COL]
            note_cols     = [c for c in cols[CFG["HDR_NOTE"]]     if c <= _MAX_TMPL_COL]

            # ── validate โครงสร้าง ────────────────────────────
            if not date_cols:
                st_errors.append(
                    f"{sh}: ⚠️ ไม่พบหัวคอลัมน์ \"{CFG['HDR_DATE']}\" "
                    f"ในแถว {CFG['HEADER_ROW']} — ข้ามค่ะ "
                    f"(พบหัว: {headers[:8]})"
                )
                continue
            if not ti_cols and not to_cols:
                st_errors.append(
                    f"{sh}: ⚠️ ไม่พบหัวคอลัมน์เวลาเข้า/ออก — ข้ามค่ะ "
                    f"(พบหัว: {headers[:8]})"
                )
                continue

            FML_ROW_BASE = 5
            FML_COL_BASE = 13

            def is_formula(col_1idx, row_1idx,
                           _fd=fml_data, _rb=FML_ROW_BASE, _cb=FML_COL_BASE):
                fml_r = row_1idx - _rb
                fml_c = col_1idx - _cb
                if fml_r < 0 or fml_r >= len(_fd) or fml_c < 0:
                    return False
                row_f = _fd[fml_r]
                return fml_c < len(row_f) and str(row_f[fml_c]).startswith("=")

            DAT_ROW_BASE = CFG["DATA_START"]

            def dat_val(col_1idx, row_1idx,
                        _dd=dat_data, _rb=DAT_ROW_BASE):
                dr = row_1idx - _rb
                dc = col_1idx - 1
                if dr < 0 or dr >= len(_dd):
                    return ""
                row_d = _dd[dr]
                return row_d[dc] if dc < len(row_d) else ""

            esh = _esc(sh)   # escaped sheet name สำหรับ range

            # ── ล้างข้อมูลเก่า (N-S, X-Y) ─────────────────────
            # ใช้ range-based bulk clear แทน cell-by-cell
            # เพื่อล้างทุก cell รวมถึง cell ที่เป็น formula หรือ merged
            _row_s = CFG["DATA_START"]
            _row_e = 36   # ล้างเฉพาะแถวข้อมูลเวลา (6-36) — แถว 37+ มี label template เช่น "มาสาย","ขาดงาน"
            _num_r = _row_e - _row_s + 1
            all_value_upd.append({
                "range" : f"'{esh}'!N{_row_s}:S{_row_e}",
                "values": [[""] * 6 for _ in range(_num_r)]
            })
            all_value_upd.append({
                "range" : f"'{esh}'!X{_row_s}:Y{_row_e}",
                "values": [[""] * 2 for _ in range(_num_r)]
            })

            # ── ประจำเดือน ──────────────────────────────────────
            for r_i, row_vals in enumerate(top_data):
                found = False
                for c_i, cell_val in enumerate(row_vals):
                    if "ประจำเดือน" in str(cell_val):
                        target_col = max(c_i + 2, 16)
                        all_value_upd.append({
                            "range" : f"'{esh}'!{col_letter(target_col)}{r_i + 1}",
                            "values": [[month_label]]
                        })
                        found = True
                        break
                if found:
                    break

            # ── วันที่ + ชื่อวัน ────────────────────────────────
            COL_DAY_EN_FB = 14
            for bi, dc in enumerate(date_cols):
                dnc = (day_name_cols[bi] if bi < len(day_name_cols)
                       else (COL_DAY_EN_FB if bi == 0 else None))
                for j in range(num_rows):
                    row      = CFG["DATA_START"] + j
                    d        = days[j]
                    date_str = d.strftime("%d/%m/") + str(d.year + 543)
                    day_idx  = (d.weekday() + 1) % 7
                    all_value_upd.append({
                        "range" : f"'{esh}'!{col_letter(dc)}{row}",
                        "values": [[date_str]]
                    })
                    if dnc:
                        all_value_upd.append({
                            "range" : f"'{esh}'!{col_letter(dnc)}{row}",
                            "values": [[DAY_EN_LIST[day_idx]]]
                        })

            # ── map วันที่ → row ────────────────────────────────
            # ใช้ days[] ที่คำนวณแล้ว (ไม่ใช่ค่าเก่าใน sheet)
            # เพื่อให้ match กับข้อมูลเวลาของเดือนปัจจุบันได้ถูกต้อง
            date_row_map: dict[tuple, int] = {}
            for bi, dc in enumerate(date_cols):
                for j in range(num_rows):
                    row_1idx = CFG["DATA_START"] + j
                    d        = days[j]
                    date_key = d.strftime("%d/%m/") + str(d.year + 543)
                    date_row_map[(bi, date_key)] = row_1idx

            # ── เวลาเข้า-ออก ────────────────────────────────────
            att_rows = att_df[att_df["รหัสพนักงาน"] == eid].to_dict("records")
            att_map  = {r["วันที่"]: r for r in att_rows if r.get("วันที่")}
            COL_DAY_EN = 14

            for (bi, date_key), r_num in date_row_map.items():
                if date_key not in att_map:
                    continue
                att    = att_map[date_key]
                ti_col = ti_cols[bi]   if bi < len(ti_cols)   else None
                to_col = to_cols[bi]   if bi < len(to_cols)   else None
                nt_col = note_cols[bi] if bi < len(note_cols) else None

                if ti_col and not is_formula(ti_col, r_num):
                    all_value_upd.append({"range": f"'{esh}'!{col_letter(ti_col)}{r_num}",
                                           "values": [[att["เวลาเข้า"]]]})
                if to_col and not is_formula(to_col, r_num):
                    all_value_upd.append({"range": f"'{esh}'!{col_letter(to_col)}{r_num}",
                                           "values": [[att["เวลาออก"]]]})
                if nt_col and not is_formula(nt_col, r_num):
                    all_value_upd.append({"range": f"'{esh}'!{col_letter(nt_col)}{r_num}",
                                           "values": [[att["หมายเหตุ"]]]})
                day_abbr = _day_en_from_datestr(date_key)
                if day_abbr and not is_formula(COL_DAY_EN, r_num):
                    all_value_upd.append({"range": f"'{esh}'!N{r_num}",
                                           "values": [[day_abbr]]})

            # ── วันหยุดตามประเพณี ─────────────────────────────────
            # เขียน "วันหยุด" (+ ชื่อจาก att ถ้ามี) ลงหมายเหตุเสมอ
            # เพื่อให้สูตรขาดงานจับเจอคำว่า "วันหยุด" ได้
            if holidays_map:
                for (bi, date_key), r_num in date_row_map.items():
                    if date_key not in holidays_map:
                        continue
                    nt_col = (note_cols[bi] if bi < len(note_cols) else None)
                    if nt_col and not is_formula(nt_col, r_num):
                        # ดึงชื่อวันหยุดจากหมายเหตุในรายงานเวลาเข้าออก (ถ้ามี)
                        existing_note = ""
                        if date_key in att_map:
                            existing_note = str(att_map[date_key].get("หมายเหตุ", "") or "").strip()
                        note_val = f"วันหยุด {existing_note}".strip() if existing_note else "วันหยุด"
                        all_value_upd.append({
                            "range" : f"'{esh}'!{col_letter(nt_col)}{r_num}",
                            "values": [[note_val]]})

            # ── formatting ──────────────────────────────────────
            for j in range(num_rows):
                row = CFG["DATA_START"] + j
                bg  = GRAY if days[j].weekday() == 6 else WHITE
                all_fmt_reqs.append({"repeatCell": {
                    "range": {"sheetId": sheet_id,
                               "startRowIndex": row - 1, "endRowIndex": row,
                               "startColumnIndex": COL_N_0, "endColumnIndex": COL_Y_0},
                    "cell" : {"userEnteredFormat": {"backgroundColor": bg}},
                    "fields": "userEnteredFormat.backgroundColor"
                }})
            for (bi, date_key), r_num in date_row_map.items():
                if date_key not in att_map:
                    continue
                if str(att_map[date_key].get("หมายเหตุ", "")).startswith("วันหยุด"):
                    all_fmt_reqs.append({"repeatCell": {
                        "range": {"sheetId": sheet_id,
                                   "startRowIndex": r_num - 1, "endRowIndex": r_num,
                                   "startColumnIndex": COL_N_0, "endColumnIndex": COL_Y_0},
                        "cell" : {"userEnteredFormat": {"backgroundColor": GRAY}},
                        "fields": "userEnteredFormat.backgroundColor"
                    }})

            # ── ไฮไลต์วันหยุดตามประเพณี (สีเทาเดียวกับวันหยุด) ─────────
            if holidays_map:
                for (bi, date_key), r_num in date_row_map.items():
                    if date_key in holidays_map:
                        all_fmt_reqs.append({"repeatCell": {
                            "range": {"sheetId": sheet_id,
                                       "startRowIndex": r_num - 1, "endRowIndex": r_num,
                                       "startColumnIndex": COL_N_0, "endColumnIndex": COL_Y_0},
                            "cell" : {"userEnteredFormat": {"backgroundColor": GRAY}},
                            "fields": "userEnteredFormat.backgroundColor"
                        }})

            # ── format เวลาเข้า/ออก → h:mm "น." (ทุกครั้งที่รัน) ─────────
            # DATA_END - 2 เพื่อไม่ให้ format ทับแถว summary (row 39-40)
            for tc in (ti_cols + to_cols):
                # ใส่ format h:mm "น." เฉพาะ data rows
                all_fmt_reqs.append({"repeatCell": {
                    "range": {
                        "sheetId"         : sheet_id,
                        "startRowIndex"   : CFG["DATA_START"] - 1,
                        "endRowIndex"     : CFG["DATA_END"] - 2,
                        "startColumnIndex": tc - 1,
                        "endColumnIndex"  : tc,
                    },
                    "cell" : {"userEnteredFormat": {
                        "numberFormat": {"type": "TIME", "pattern": 'h:mm "น."'}
                    }},
                    "fields": "userEnteredFormat.numberFormat"
                }})
                # clear format แถว summary (row 39-40) ที่อาจค้างจาก run ก่อน
                all_fmt_reqs.append({"repeatCell": {
                    "range": {
                        "sheetId"         : sheet_id,
                        "startRowIndex"   : CFG["DATA_END"] - 2,
                        "endRowIndex"     : CFG["DATA_END"],
                        "startColumnIndex": tc - 1,
                        "endColumnIndex"  : tc,
                    },
                    "cell" : {"userEnteredFormat": {
                        "numberFormat": {"type": "TIME", "pattern": "h:mm"}
                    }},
                    "fields": "userEnteredFormat.numberFormat"
                }})

            success_cnt += 1

        progress_bar.progress(base_prog + prog_span * 0.65)

        # ── Phase 5a: Batch write values ─────────────────────────
        status_text.info(
            f"✍️ [5/5] เขียน {len(all_value_upd)} เซลล์ "
            f"+ จัดสี {len(all_fmt_reqs)} แถว {chunk_label}..."
        )
        if all_value_upd:
            try:
                _sheets_retry(ss.values_batch_update, {
                    "valueInputOption": "USER_ENTERED",
                    "data"            : all_value_upd
                })
            except Exception as e:
                # batch write ล้มเหลว → ส่ง chunk นี้ไป slow mode
                st.warning(
                    f"⚠️ batch write ล้มเหลว [{chunk_label}]: {e}\n"
                    f"จะลอง slow mode ให้อัตโนมัติค่ะ..."
                )
                success_cnt -= sum(1 for _ in chunk)   # ยกเลิก success ของ chunk นี้
                fallback_pairs.extend(chunk)
                progress_bar.progress(base_prog + prog_span)
                continue

        progress_bar.progress(base_prog + prog_span * 0.85)

        # ── Phase 5b: Batch format ────────────────────────────────
        if all_fmt_reqs:
            try:
                _sheets_retry(ss.batch_update, {"requests": all_fmt_reqs})
            except Exception as e:
                st.warning(f"⚠️ batch format ล้มเหลว [{chunk_label}]: {e} (ข้อมูลถูกเขียนแล้วค่ะ)")

        progress_bar.progress(base_prog + prog_span)

    # ── Fallback: slow mode สำหรับ chunk ที่ batch write ล้มเหลว ─
    if fallback_pairs:
        st.info(f"🔄 ลอง slow mode ให้ {len(fallback_pairs)} Sheet ที่ค้างค่ะ...")
        for idx, (eid, sh) in enumerate(fallback_pairs):
            try:
                emp_rows   = att_df[att_df["รหัสพนักงาน"] == eid].to_dict("records")
                shared_ws  = _sheets_retry(ss.worksheet, sh)
                shared_hdr = _sheets_retry(shared_ws.row_values, CFG["HEADER_ROW"])
                r1 = _clear_attendance(ss, sh)
                if not r1["ok"]:
                    st_errors.append(f"{sh} (fallback): ล้างไม่ได้ — {r1['msg']}")
                    continue
                r2 = _generate_dates(ss, sh, auto_month, auto_year,
                                     ws=shared_ws, headers=shared_hdr)
                if not r2["ok"]:
                    st_errors.append(f"{sh} (fallback): สร้างวันที่ไม่ได้ — {r2['msg']}")
                    continue
                r3 = _import_attendance(ss, sh, emp_rows,
                                        ws=shared_ws, headers=shared_hdr)
                if not r3["ok"]:
                    st_errors.append(f"{sh} (fallback): วางเวลาไม่ได้ — {r3['msg']}")
                    continue
                success_cnt += 1
            except Exception as e:
                st_errors.append(f"{sh} (fallback): {e}")
            if idx < len(fallback_pairs) - 1:
                time.sleep(6)

    progress_bar.progress(1.0)
    status_text.empty()

    if not st_errors:
        st.success(f"✅ เสร็จสิ้น! ประมวลผลครบ {success_cnt} Sheet ค่ะ")
    else:
        st.warning(f"⚠️ เสร็จ {success_cnt} Sheet / มีปัญหา {len(st_errors)} Sheet")
        with st.expander("❌ รายการที่มีปัญหา"):
            for err in st_errors:
                st.text(err)


def _generate_dates(ss: gspread.Spreadsheet, sheet_name: str,
                    month: int, year_be: int,
                    ws=None, headers=None) -> dict:
    """สร้างวันที่รอบ 26 เดือนก่อน → 25 เดือนนี้
    รับ ws/headers ที่เปิดไว้แล้วจากภายนอกได้ เพื่อลด API call"""
    try:
        if ws is None:
            ws = _sheets_retry(ss.worksheet, sheet_name)
        year_ce = year_be - 543

        prev_month = 12 if month == 1 else month - 1
        prev_year  = year_ce - 1 if month == 1 else year_ce
        start = date(prev_year, prev_month, 26)
        end   = date(year_ce, month, 25)

        days, cur = [], start
        while cur <= end:
            days.append(cur)
            cur += timedelta(days=1)

        if headers is None:
            headers = _sheets_retry(ws.row_values, CFG["HEADER_ROW"])
        cols          = find_col_indices(headers, CFG["HDR_DATE"], CFG["HDR_DAY_NAME"])
        # กรองเฉพาะ template area (A-Y = col 1-25) ไม่ให้เขียนทับ history block
        _MAX_TMPL_COL = 25
        date_cols     = [c for c in cols[CFG["HDR_DATE"]]     if c <= _MAX_TMPL_COL]
        day_name_cols = [c for c in cols[CFG["HDR_DAY_NAME"]] if c <= _MAX_TMPL_COL]

        if not date_cols:
            return {"ok": False,
                    "msg": f'ไม่พบคอลัมน์ "{CFG["HDR_DATE"]}" ใน row {CFG["HEADER_ROW"]}'}

        updates  = []
        num_rows = min(len(days), CFG["DATA_END"] - CFG["DATA_START"] + 1)

        DAY_EN_LIST   = ["Su","Mo","Tu","We","Th","Fr","Sa"]
        COL_DAY_EN_FB = 14   # column N — fallback ถ้าไม่มีหัว "ชื่อวัน"

        for bi, dc in enumerate(date_cols):
            # ถ้าหาหัว "ชื่อวัน" ไม่เจอ ให้ใช้ column N (14) สำหรับ block แรก
            # block ที่สองใช้คอลัมน์ถัดจาก date col ของ block นั้น (ถ้ามี)
            if bi < len(day_name_cols):
                dnc = day_name_cols[bi]
            elif bi == 0:
                dnc = COL_DAY_EN_FB   # fallback: column N
            else:
                dnc = None

            for i in range(num_rows):
                row      = CFG["DATA_START"] + i
                d        = days[i]
                date_str = d.strftime("%d/%m/") + str(d.year + 543)
                day_idx  = (d.weekday() + 1) % 7   # 0=อา. 1=จ. ...
                updates.append({"range": f"{col_letter(dc)}{row}",
                                "values": [[date_str]]})
                if dnc:
                    updates.append({"range": f"{col_letter(dnc)}{row}",
                                    "values": [[DAY_EN_LIST[day_idx]]]})
            # ล้างแถวที่เกิน num_rows (กรณีเดือนนี้สั้นกว่าเดือนที่แล้ว)
            # เช่น กรกฎาคม 31 วัน แต่มิถุนายน 30 วัน → ต้องล้าง row 36 ออก
            # จำกัดแค่ row 36 (DATA_START+30) ไม่แตะ row 37+ ซึ่งเป็น label/สรุป
            _MAX_DATA_ROW = CFG["DATA_START"] + 30  # สูงสุด 31 วัน → row 36
            for extra in range(num_rows, CFG["DATA_END"] - CFG["DATA_START"] + 1):
                extra_row = CFG["DATA_START"] + extra
                if extra_row > _MAX_DATA_ROW:       # ห้ามลบ row 37+ (label ไม่ใช่วัน)
                    break
                updates.append({"range": f"{col_letter(dc)}{extra_row}",
                                "values": [[""]]})
                if dnc:
                    updates.append({"range": f"{col_letter(dnc)}{extra_row}",
                                    "values": [[""]]})

        # ── อัปเดตเซลล์ "ประจำเดือน" (ค้นหาใน 10 แถวแรก) ──────────────
        month_label = f"{MONTH_TH_FULL[month]} {year_be}"
        try:
            top_rows = _sheets_retry(ws.get, "A1:Z10", value_render_option="FORMATTED_VALUE")
            for r_i, row_vals in enumerate(top_rows):
                for c_i, cell_val in enumerate(row_vals):
                    if "ประจำเดือน" in str(cell_val):
                        # เขียนค่าในเซลล์ถัดไป (คอลัมน์ P = 16 หรือคอลัมน์ c_i+2)
                        target_col = max(c_i + 2, 16)   # อย่างน้อยคอลัมน์ P
                        updates.append({
                            "range": f"{col_letter(target_col)}{r_i + 1}",
                            "values": [[month_label]]
                        })
                        break
        except Exception:
            pass   # ถ้าหาไม่เจอก็ข้ามไป ไม่ให้ error หลัก

        _sheets_retry(ws.batch_update, updates)

        # ── ไฮไลต์แถวอาทิตย์ N-Y สีเทา / ล้างวันอื่นกลับขาว ────────────
        GRAY  = {"red": 211/255, "green": 211/255, "blue": 211/255}
        WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
        COL_N_0 = 13   # N = col 14 (0-indexed)
        COL_Y_0 = 25   # Y = col 25 → endColumnIndex exclusive

        hl_requests = []
        for i in range(num_rows):
            row = CFG["DATA_START"] + i
            d   = days[i]
            bg  = GRAY if d.weekday() == 6 else WHITE   # 6 = Sunday
            hl_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": COL_N_0,
                        "endColumnIndex": COL_Y_0,
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": bg}},
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
        if hl_requests:
            _sheets_retry(ss.batch_update, {"requests": hl_requests})

        return {"ok": True,
                "msg": f"สร้างวันที่ {num_rows} วัน ใน {len(date_cols)} block เรียบร้อยค่ะ"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def _read_attendance_file(file) -> pd.DataFrame | None:
    """อ่านไฟล์เวลาเข้า-ออก → DataFrame
    columns: รหัสพนักงาน, วัน, วันที่, เวลาเข้า, เวลาออก, หมายเหตุ
    รองรับไฟล์ที่มีพนักงานหลายคนรวมในชีตเดียว"""
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
            row     = df_raw.iloc[i]
            max_idx = max(CFG["ATT_EMP_ID"], CFG["ATT_DATE"], CFG["ATT_TIME_IN"],
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

            def norm_time(val: str) -> str:
                """Normalize เวลา → "H:MM" (ไม่มี น.) เพื่อให้ USER_ENTERED parse เป็น time จริงใน Sheets
                รองรับ: "8:14 น.", "08:14:00", "8:54:00 AM", "0.347222..." (Excel fraction), "7:37"
                """
                if not val:
                    return ""
                # จัดการรูปแบบ AM/PM ก่อน เช่น "8:54:00 AM"
                am_pm = re.match(r"^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$", val.strip(), re.IGNORECASE)
                if am_pm:
                    h, m, period = int(am_pm.group(1)), int(am_pm.group(2)), am_pm.group(3).upper()
                    if period == "PM" and h != 12:
                        h += 12
                    elif period == "AM" and h == 12:
                        h = 0
                    return f"{h}:{m:02d}"
                # ตัด "น." และช่องว่างออก
                t = val.replace("น.", "").replace("น", "").strip()
                # Excel เก็บเวลาเป็น float fraction เช่น 0.347222 = 08:20
                try:
                    f = float(t)
                    if 0 < f < 1:
                        total_min = round(f * 1440)
                        h, m = divmod(total_min, 60)
                        return f"{h}:{m:02d}"
                except ValueError:
                    pass
                # รูปแบบ HH:MM:SS หรือ H:MM
                m2 = re.match(r"^(\d{1,2}):(\d{2})", t)
                if m2:
                    h, m = int(m2.group(1)), int(m2.group(2))
                    return f"{h}:{m:02d}"
                return val  # fallback ไม่แตะ

            emp_id_raw = row.iloc[CFG["ATT_EMP_ID"]]
            emp_id = str(int(emp_id_raw)) if pd.notna(emp_id_raw) else ""

            rows.append({
                "รหัสพนักงาน": emp_id,
                "วัน"         : safe_str(CFG["ATT_DAY"]),
                "วันที่"      : date_str,
                "เวลาเข้า"    : norm_time(safe_str(CFG["ATT_TIME_IN"])),
                "เวลาออก"     : norm_time(safe_str(CFG["ATT_TIME_OUT"])),
                "หมายเหตุ"    : safe_str(CFG["ATT_NOTE"]),
            })
        return pd.DataFrame(rows) if rows else None
    except Exception as e:
        st.warning(f"อ่านไฟล์ไม่ได้: {e}")
        return None


def _import_attendance(ss: gspread.Spreadsheet, sheet_name: str,
                       rows: list,
                       ws=None, headers=None) -> dict:
    """วางข้อมูลเวลาเข้า-ออกลง template (ข้ามเซลล์ที่มีสูตร)
    รับ ws/headers ที่เปิดไว้แล้วจากภายนอกได้ เพื่อลด API call"""
    try:
        if ws is None:
            ws = _sheets_retry(ss.worksheet, sheet_name)
        att_map = {r["วันที่"]: r for r in rows if r.get("วันที่")}

        if headers is None:
            headers = _sheets_retry(ws.row_values, CFG["HEADER_ROW"])
        cols      = find_col_indices(headers,
                                     CFG["HDR_DATE"], CFG["HDR_TIME_IN"],
                                     CFG["HDR_TIME_OUT"], CFG["HDR_NOTE"],
                                     CFG["HDR_DAY_NAME"])
        # กรองเฉพาะ template area (A-Y = col 1-25) ไม่ให้เขียนทับ history block
        _MAX_TMPL_COL = 25
        date_cols = [c for c in cols[CFG["HDR_DATE"]]     if c <= _MAX_TMPL_COL]
        ti_cols   = [c for c in cols[CFG["HDR_TIME_IN"]]  if c <= _MAX_TMPL_COL]
        to_cols   = [c for c in cols[CFG["HDR_TIME_OUT"]] if c <= _MAX_TMPL_COL]
        note_cols = [c for c in cols[CFG["HDR_NOTE"]]     if c <= _MAX_TMPL_COL]
        day_cols  = [c for c in cols[CFG["HDR_DAY_NAME"]] if c <= _MAX_TMPL_COL]

        if not date_cols:
            return {"ok": False,
                    "msg": f'ไม่พบคอลัมน์ "{CFG["HDR_DATE"]}" ในเทมเพลต'}

        # ดึง formula ทั้ง block เพื่อตรวจ
        start_c  = min(date_cols)
        end_c    = ws.col_count
        rng      = (f"{col_letter(start_c)}{CFG['DATA_START']}:"
                    f"{col_letter(end_c)}{CFG['DATA_END']}")
        formulas = _sheets_retry(ws.get, rng, value_render_option="FORMULA")

        # สร้าง map วันที่ → row number สำหรับแต่ละ block
        date_row_map: dict[tuple, int] = {}
        for bi, dc in enumerate(date_cols):
            col_data = _sheets_retry(ws.col_values, dc)
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

        # ── ไฮไลต์แถวที่หมายเหตุ**ขึ้นต้นด้วย** "วันหยุด" N-Y สีเทา ─────────
        GRAY    = {"red": 211/255, "green": 211/255, "blue": 211/255}
        COL_N_0 = 13   # N (0-indexed)
        COL_Y_0 = 25   # Y+1 (0-indexed, exclusive)

        holiday_requests = []
        for (bi, date_key), r_num in date_row_map.items():
            if date_key not in att_map:
                continue
            note = str(att_map[date_key].get("หมายเหตุ", ""))
            if note.startswith("วันหยุด"):
                holiday_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": r_num - 1,
                            "endRowIndex": r_num,
                            "startColumnIndex": COL_N_0,
                            "endColumnIndex": COL_Y_0,
                        },
                        "cell": {"userEnteredFormat": {"backgroundColor": GRAY}},
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
        if holiday_requests:
            _sheets_retry(ss.batch_update, {"requests": holiday_requests})

        return {"ok": True, "msg": f"วางข้อมูล {pasted} วัน เรียบร้อยค่ะ"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def _export_history(ss: gspread.Spreadsheet, sheet_name: str,
                    month_yr: str,
                    ss_db: gspread.Spreadsheet | None = None) -> dict:
    """
    บันทึกประวัติเงินเดือน — copy คอลัมน์ M–Y (col 13–25) จากชีตพนักงาน
    แล้ววางต่อในชีตเดิม โดยเว้น 1 คอลัมน์จากข้อมูลล่าสุด
    (single-sheet wrapper — ใช้ _export_history_batch สำหรับ bulk)
    """
    results = _export_history_batch(ss, [sheet_name], month_yr, ss_db)
    return results[0] if results else {"ok": False, "msg": "ไม่มีข้อมูล"}


def _export_history_batch(ss: gspread.Spreadsheet, sheet_names: list,
                           month_yr: str,
                           ss_db: gspread.Spreadsheet | None = None,
                           status_text=None, progress_bar=None) -> list:
    """
    บันทึกประวัติทุก sheet — copy ทั้ง format + สูตร (PASTE_NORMAL) แล้วเขียน bookmark:
      Phase 1: อ่านทุก sheet ด้วย batch_get (~3 API calls)
      Phase 2: copyPaste PASTE_NORMAL (เส้นตาราง/สี/header + สูตรครบ)
      Phase 3: write bookmark month_yr เท่านั้น (ไม่ต้อง write values ทับ)
    """
    COL_HIST_START = 13
    COL_HIST_END   = 25
    NUM_ROWS       = 100   # บันทึกถึงแถว 100
    width          = COL_HIST_END - COL_HIST_START + 1
    CHUNK          = 100   # max ranges per batch_get call

    results      = []
    copy_reqs     = []   # สำหรับ PASTE_FORMAT เท่านั้น
    write_tasks   = []
    bookmark_tasks = []  # ranges สำหรับเขียน month_yr bookmark ที่คอลัมน์สุดท้าย
    n_sheets     = len(sheet_names)

    def _upd(pct: float, msg: str):
        """อัปเดต progress bar + status text (ถ้ามี)"""
        if progress_bar  is not None: progress_bar.progress(min(pct, 1.0))
        if status_text   is not None: status_text.info(msg)

    # ── helpers: ขยับ column reference ในสูตร (สำหรับ copy formula ไป history) ──
    TW_START = 20   # col T (1-indexed)
    TW_END   = 23   # col W (1-indexed)

    # ── Config: cells ใน history block ที่ให้เขียนสูตรอ้างอิง col A-L คงที่ ──────
    # key: (template_row, template_col) — 1-indexed
    # value: สูตรที่จะเขียน (อ้างอิง col A-L เท่านั้น ไม่ขยับตาม offset)
    HISTORY_FORMULA_REFS = {
        (2,  19): "=F2",    # Row 2  Col S → =F2
        (2,  20): "=G2",    # Row 2  Col T → =G2
        (3,  16): "=C3",    # Row 3  Col P → =C3
        (4,  16): "=C4",    # Row 4  Col P → =C4
        (4,  19): "=F4",    # Row 4  Col S → =F4
        (39, 20): "=G39",   # Row 39 Col T → =G39
        (45, 20): "=G45",   # Row 45 Col T → =G45
        (49, 20): "=G49",   # Row 49 Col T → =G49
        (50, 20): "=G50",   # Row 50 Col T → =G50
    }

    def _col_str_to_num(s: str) -> int:
        n = 0
        for c in s:
            n = n * 26 + (ord(c) - 64)
        return n

    def _col_num_to_str(n: int) -> str:
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    def _shift_formula(formula: str, offset: int) -> str:
        """ขยับ relative column reference ทุกตัวในสูตรไป offset คอลัมน์"""
        if not isinstance(formula, str) or not formula.startswith("="):
            return formula
        def _repl(m):
            dol1, col_s, dol2, row_s = m.group(1), m.group(2), m.group(3), m.group(4)
            if not dol1:   # relative col → shift
                col_s = _col_num_to_str(_col_str_to_num(col_s) + offset)
            return f"{dol1}{col_s}{dol2}{row_s}"
        return re.sub(r'(\$?)([A-Z]+)(\$?)(\d+)', _repl, formula)

    # ── Phase 1a: โหลด worksheet objects ทุกชีตพร้อมกัน (1 API call) ───
    _upd(0.05, f"⏳ กำลังโหลดรายชื่อชีต ({n_sheets} Sheet)...")
    try:
        all_ws_list = _sheets_retry(ss.worksheets)
        ws_dict = {ws.title: ws for ws in all_ws_list}
    except Exception as e:
        if status_text: status_text.empty()
        return [{"ok": False, "sheet": n, "msg": f"โหลด worksheets ล้มเหลว: {e}"}
                for n in sheet_names]

    valid_names = [n for n in sheet_names if n in ws_dict]
    missing     = [n for n in sheet_names if n not in ws_dict]
    for n in missing:
        results.append({"ok": False, "sheet": n, "msg": "ไม่พบชีตในไฟล์"})

    # ── Phase 1b: batch-read แถว 1-5 ทุกชีต เพื่อหา last column (~1-2 calls) ─
    # อ่านแถว 1-5 แทนแถว 1 อย่างเดียว เพราะแถว 1 ของ history block มักว่างเปล่า
    # แต่แถว 2-4 มีหัว "ใบประเมิน / ประจำเดือน" ที่บอกว่ามีประวัติอยู่แล้ว
    _upd(0.15, f"⏳ กำลังตรวจสอบตำแหน่งประวัติ ({len(valid_names)} Sheet)...")
    header_data: dict[str, list[list]] = {}   # name → list of rows (max 5 rows)
    for i in range(0, len(valid_names), CHUNK):
        chunk_names  = valid_names[i:i + CHUNK]
        chunk_ranges = [f"'{n}'!1:5" for n in chunk_names]
        try:
            resp = _sheets_retry(ss.values_batch_get, chunk_ranges)
            for name, vr in zip(chunk_names,
                                 resp.get("valueRanges", [])):
                header_data[name] = vr.get("values", [])
        except Exception:
            for name in chunk_names:
                header_data[name] = []

    # ── Phase 1b.5: Freeze history เก่า → แปลง formula → static value ─────
    # ป้องกันไม่ให้ history เก่าที่เป็น formula เปลี่ยนตามข้อมูลเดือนใหม่
    # อ่านด้วย FORMATTED_VALUE (ได้ค่าที่คำนวณแล้ว) แล้วเขียนทับด้วย RAW (ตัดสูตรออก)
    #
    # FIX v2: เดิมใช้ header_data (แถว 1-5) หา last_col ก่อน ถ้าไม่เจอก็ข้าม freeze
    #          แต่ history เก่า (ที่มี formula) อาจมีข้อมูลแค่แถว 6-36 ไม่มีแถว 1-5
    #          ทำให้ last_col ไม่เพิ่ม → freeze ไม่ถูกเรียก → history ยังมี formula อยู่
    #          แก้: อ่าน cols COL_HIST_END+1 ถึง FREEZE_MAX_COL ทุกแถว (1-NUM_ROWS)
    #          ถ้ามีข้อมูลก็เขียนทับด้วย RAW — ถ้าว่างก็ไม่มีผล
    #          และเก็บ actual last_col ต่อ Phase 1c ด้วย เพื่อคำนวณ paste_start ที่ถูก
    FREEZE_MAX_COL = 150   # ครอบคลุม history ได้ประมาณ 9 เดือน (13 cols × 9 = 117)
    _upd(0.22, f"⏳ กำลัง freeze history เก่า ({len(valid_names)} Sheet)...")

    freeze_read_ranges: list[str]   = []
    freeze_sheet_info : list[tuple] = []   # (sheet_name, ws_title, start_col)
    frozen_last_col   : dict[str, int] = {}   # sheet_name → last non-empty col (for Phase 1c)

    for sheet_name in valid_names:
        ws = ws_dict[sheet_name]
        freeze_read_ranges.append(
            f"'{sheet_name}'!"
            f"{col_letter(COL_HIST_END + 1)}1"
            f":{col_letter(FREEZE_MAX_COL)}{NUM_ROWS}"
        )
        freeze_sheet_info.append((sheet_name, ws.title, COL_HIST_END + 1))

    # Freeze history เก่า → segment-based (รูปปลอดภัย) + track last_col สำหรับ Phase 1c
    _fz_pairs = [
        (sname, ws_dict[sname].title)
        for sname in valid_names
        if sname in ws_dict
    ]
    _lc = _freeze_formulas_to_values(
        ss,
        _fz_pairs,
        col_start=COL_HIST_END + 1,
        num_rows=NUM_ROWS,
    )
    _fz_failed_r = _lc.pop("__failed_reads__",  [])
    _fz_failed_w = _lc.pop("__failed_writes__", [])
    frozen_last_col.update(_lc)
    if _fz_failed_r:
        st.warning(f"⚠️ Freeze อ่านไม่ได้ {len(_fz_failed_r)} Sheet: {', '.join(str(x) for x in _fz_failed_r)}")
    if _fz_failed_w:
        st.warning(f"⚠️ Freeze เขียนค่านิ่งไม่สำเร็จ {len(_fz_failed_w)} รายการ:\n" +
                   "\n".join(f"• {x}" for x in _fz_failed_w))

    # ── Phase 1c: คำนวณ paste position + resize ถ้าจำเป็น ────────────
    paste_info: dict[str, tuple] = {}   # name → (paste_start, paste_end, ws)
    src_ranges_list: list[str]   = []

    resize_reqs = []   # รวม resize ทุกชีตไว้ก่อน — batch ครั้งเดียว
    _debug_paste: list[str] = []   # เก็บ debug info สำหรับแสดงผล

    for sheet_name in valid_names:
        ws      = ws_dict[sheet_name]
        rows    = header_data.get(sheet_name, [])

        # หา rightmost non-empty cell ใน ROW 1 เท่านั้น เฉพาะส่วน history (col >= COL_HIST_START)
        # ใช้แค่แถว 1 เพราะ bookmark วางที่แถว 1 ท้าย block เสมอ
        # ห้ามใช้แถว 2-5 เพราะ template metadata (ชื่อพนักงาน, วันที่ ฯลฯ) ในแถวนั้น
        # อาจยืดเกิน col Y ทำให้ inflate last_col_hdr จนเกิด gap เกิน 1 ช่อง
        last_col_hdr = COL_HIST_END
        hist_start_idx = COL_HIST_START - 1   # 0-based index ของ col M
        for row in rows[:1]:   # ← สแกนเฉพาะแถว 1 เท่านั้น (bookmark row)
            if len(row) <= hist_start_idx:
                continue   # แถวนี้สั้นกว่า col M — ข้ามได้เลย
            sub = row[hist_start_idx:]   # ตัดส่วน template (A-L) ออก
            for ci_sub in range(len(sub) - 1, -1, -1):
                if str(sub[ci_sub]).strip():
                    actual_col = hist_start_idx + ci_sub + 1   # 1-based
                    if actual_col > last_col_hdr:
                        last_col_hdr = actual_col
                    break

        # ── FIX v23: ใช้ bookmark แถว 1 เป็นหลักเสมอ ──────────────────────
        # bookmark ถูกเขียนที่ paste_end-1 (จุดสิ้นสุด block จริงๆ) ทุกครั้งที่บันทึก
        # → last_col_hdr ชี้ตำแหน่งสิ้นสุด block ที่แม่นยำ → gap = 1 col เสมอ
        #
        # (v20 เคยใช้ frozen_last_col override เพื่อป้องกัน bookmark เก่าที่ผิด
        #  แต่ frozen_last_col scan data แถว 6+ ซึ่งบางสูตรอาจ evaluate นอก block
        #  ทำให้ frozen_last_col สูงกว่าจริง 1-2 col → gap บวมเป็น 2-3 col แทน)
        fzlc = frozen_last_col.get(sheet_name)   # ยังเก็บไว้ใช้ใน debug log
        last_col = last_col_hdr
        _src = f"bookmark={last_col_hdr}"

        paste_start = last_col + 2
        paste_end   = paste_start + width

        _debug_paste.append(
            f"{sheet_name}: hdr={last_col_hdr}  fz={fzlc or '–'}  "
            f"→ {_src}  last_col={last_col}  paste_start={paste_start}"
        )

        if paste_end > ws.col_count:
            resize_reqs.append({
                "updateSheetProperties": {
                    "properties": {
                        "sheetId"    : ws.id,
                        "gridProperties": {
                            "rowCount"   : max(ws.row_count, NUM_ROWS),
                            "columnCount": paste_end + 10,
                        },
                    },
                    "fields": "gridProperties.rowCount,gridProperties.columnCount",
                }
            })

        paste_info[sheet_name] = (paste_start, paste_end, ws)
        src_ranges_list.append(
            f"'{sheet_name}'!{col_letter(COL_HIST_START)}1"
            f":{col_letter(COL_HIST_END)}{NUM_ROWS}"
        )

    # ── DEBUG: แสดง paste position ของทุกชีต (ใช้ diagnose gap issue) ──
    if _debug_paste:
        with st.expander("🔍 DEBUG — paste position (คลิกเพื่อดู)", expanded=False):
            st.code("\n".join(_debug_paste), language=None)

    # batch resize ทุกชีตที่ต้องขยายพร้อมกัน (1 API call แทน N calls)
    if resize_reqs:
        try:
            _sheets_retry(ss.batch_update, {"requests": resize_reqs})
        except Exception:
            pass

    # ── Phase 1d: batch-read source ranges ทุกชีตพร้อมกัน (~1-2 calls) ─
    _upd(0.35, f"⏳ กำลังอ่านข้อมูลเวลา ({len(valid_names)} Sheet)...")
    # ส่ง params เป็น dict เพื่อให้ได้ค่า formatted (ไม่ใช่สูตร)
    src_data: dict[str, list] = {}
    for i in range(0, len(valid_names), CHUNK):
        chunk_names  = valid_names[i:i + CHUNK]
        chunk_ranges = src_ranges_list[i:i + CHUNK]
        try:
            # ไม่ส่ง params เพิ่ม — gspread default คือ FORMATTED_VALUE อยู่แล้ว
            resp = _sheets_retry(ss.values_batch_get, chunk_ranges)
            for name, vr in zip(chunk_names,
                                  resp.get("valueRanges", [])):
                src_data[name] = vr.get("values", [])
        except Exception as e:
            for name in chunk_names:
                src_data[name] = []

    # ── Phase 1f: อ่าน FORMULA mode สำหรับ col T-W (batch — ~1 API call) ────
    # อ่านสูตรจาก template ก่อน แล้วจะเขียนลง history block ใน Phase 3
    # (PASTE_NORMAL อาจ copy สูตรได้อยู่แล้ว แต่ถ้า template มีค่านิ่งใน T-W
    #  Phase 3 จะ write สูตรที่อ่านมาได้ทับให้แน่ใจ)
    _upd(0.42, f"⏳ กำลังอ่านสูตร col T-W ({len(valid_names)} Sheet)...")
    formula_tw: dict[str, list] = {}   # name → 2D list (DATA_START:DATA_END × T:W)
    tw_rng_suffix = (f"{col_letter(TW_START)}{CFG['DATA_START']}"
                     f":{col_letter(TW_END)}{CFG['DATA_END']}")
    tw_ranges_list = [f"'{n}'!{tw_rng_suffix}" for n in valid_names]
    for i in range(0, len(valid_names), CHUNK):
        chunk_names  = valid_names[i:i + CHUNK]
        chunk_ranges = tw_ranges_list[i:i + CHUNK]
        try:
            resp = _sheets_retry(ss.values_batch_get, chunk_ranges,
                                 params={"valueRenderOption": "FORMULA"})
            for name, vr in zip(chunk_names, resp.get("valueRanges", [])):
                formula_tw[name] = vr.get("values", [])
        except Exception:
            # fallback: อ่านทีละชีต
            for sheet_name in chunk_names:
                ws_tmp = ws_dict[sheet_name]
                try:
                    formula_tw[sheet_name] = _sheets_retry(
                        ws_tmp.get, tw_rng_suffix,
                        value_render_option="FORMULA")
                except Exception:
                    formula_tw[sheet_name] = []

    # ── Phase 1e: สร้าง copy_reqs (format only) + write_tasks ──────────
    for sheet_name in valid_names:
        paste_start, paste_end, ws = paste_info[sheet_name]
        src_values = src_data.get(sheet_name, [])

        # unmerge destination ก่อน เพื่อป้องกัน "partially intersects a merge" error
        # แล้วค่อย PASTE_FORMAT: copy เส้นตาราง/สี/header style — ไม่มีสูตร ไม่มี value
        dest_range = {
            "sheetId"         : ws.id,
            "startRowIndex"   : 0,
            "endRowIndex"     : NUM_ROWS,
            "startColumnIndex": paste_start - 1,
            "endColumnIndex"  : paste_end - 1,
        }
        copy_reqs.append((sheet_name, {"unmergeCells": {"range": dest_range}}))
        copy_reqs.append((sheet_name, {
            "copyPaste": {
                "source": {
                    "sheetId"         : ws.id,
                    "startRowIndex"   : 0,
                    "endRowIndex"     : NUM_ROWS,
                    "startColumnIndex": COL_HIST_START - 1,
                    "endColumnIndex"  : COL_HIST_END,
                },
                "destination": dest_range,
                "pasteType"       : "PASTE_NORMAL",   # copy ทั้ง format + สูตร (relative refs ขยับอัตโนมัติ)
                "pasteOrientation": "NORMAL",
            }
        }))

        if src_values:
            # ไม่กรองวันที่ — รอบบัญชีครอบคลุม 2 เดือน (เช่น 26 มิ.ย.–25 ก.ค.)
            # การกรองเฉพาะเดือนปัจจุบันทำให้ข้อมูลปลายเดือนก่อนหายไป (เช่น 26-30 มิ.ย.)
            # ทำให้จำนวนแถวใน history น้อยกว่า template → ข้อมูลเหลื่อมตำแหน่ง
            # → เก็บทุกแถวตามที่อ่านมาจาก template เลยค่ะ
            filtered = list(src_values)

            dst_a1 = (f"{col_letter(paste_start)}1:"
                      f"{col_letter(paste_end - 1)}{NUM_ROWS}")
            # bookmark_short: ใช้ใน fallback (ws.update ไม่ต้องการชื่อชีต)
            bm_short = f"{col_letter(paste_end - 1)}1"
            write_tasks.append((ws, dst_a1, filtered, sheet_name, paste_start, bm_short))

        # ── bookmark: เขียน month_yr ที่คอลัมน์สุดท้ายของ block แถว 1 ──────
        # ทำให้ครั้งถัดไปสแกนแถว 1-5 แล้วเจอ bookmark จะรู้ว่า block จบที่ไหน
        # ป้องกัน history ซ้อนทับกัน แม้คอลัมน์ท้ายของ template จะว่างเปล่า
        bm_short_all = f"{col_letter(paste_end - 1)}1"
        bookmark_tasks.append((ws, bm_short_all))

        results.append({"ok": True, "sheet": sheet_name,
                        "msg": f"บันทึกประวัติ {month_yr} ที่คอลัมน์ "
                               f"{col_letter(paste_start)} เรียบร้อยค่ะ"})

    # ── Phase 2: batch copyPaste PASTE_FORMAT (เส้นตาราง/สีเท่านั้น) ──────
    CP_CHUNK = 20   # เล็กลงเพื่อไม่ timeout + ทำให้ progress ดูไม่ค้าง
    if copy_reqs:
        reqs_only   = [r for _, r in copy_reqs]
        names_only  = [n for n, _ in copy_reqs]
        total_reqs  = len(reqs_only)
        for i in range(0, total_reqs, CP_CHUNK):
            chunk       = reqs_only[i:i + CP_CHUNK]
            chunk_names = names_only[i:i + CP_CHUNK]
            done_sheets = i // 2          # แต่ละชีตมี 2 req (unmerge + copyPaste)
            pct = 0.55 + 0.20 * (i / total_reqs)
            _upd(pct, f"⏳ กำลัง copy ตาราง/สี ({done_sheets}/{len(valid_names)} Sheet)...")
            try:
                _sheets_retry(ss.batch_update, {"requests": chunk})
            except Exception:
                for req, sname in zip(chunk, chunk_names):
                    try:
                        _sheets_retry(ss.batch_update, {"requests": [req]})
                    except Exception as e2:
                        err_str = str(e2)
                        # ถ้า error เป็นเรื่อง merge ให้ treat เป็น warning (ค่ายังถูก write ใน Phase 3)
                        is_merge_err = "intersects a merge" in err_str or "partially intersects" in err_str
                        for r in results:
                            if r["sheet"] == sname:
                                if is_merge_err:
                                    # ok ยังเป็น True — Phase 3 จะ write ค่าให้อยู่ดี
                                    r["msg"] = "⚠️ copy format ข้ามเพราะมี merged cell (ค่าบันทึกแล้ว)"
                                else:
                                    r["ok"]  = False
                                    r["msg"] = f"copy format ล้มเหลว: {e2}"

    # ── Phase 3: bookmark + เขียนสูตร col T-W ลง history block ──────────────────
    # รวม writes ทุก sheet เป็น 2 batch calls (RAW + USER_ENTERED)
    # แทนที่จะ loop ทีละ sheet ทีละ range (~700 calls → ~2 calls)
    _upd(0.90, f"⏳ กำลังบันทึก bookmark + สูตร ({len(bookmark_tasks)} Sheet)...")

    raw_writes      : list[dict] = []   # bookmark (RAW)
    formula_writes  : list[dict] = []   # สูตร T-W + HISTORY_FORMULA_REFS (USER_ENTERED)

    for bm_ws, bm_short in bookmark_tasks:
        ws_title   = bm_ws.title
        sheet_name = ws_title

        # ── bookmark (RAW) ──
        raw_writes.append({
            "range" : f"'{ws_title}'!{bm_short}",
            "values": [[month_yr]],
        })

        # ── เขียนสูตร T-W (USER_ENTERED) ──
        fdata = formula_tw.get(sheet_name, [])
        if sheet_name not in paste_info:
            continue
        paste_start, _pe, _ = paste_info[sheet_name]
        col_offset = paste_start - COL_HIST_START
        hist_t     = paste_start + (TW_START - COL_HIST_START)

        if fdata:
            adjusted = []
            for row in fdata:
                adj_row = [_shift_formula(f, col_offset) for f in row]
                while len(adj_row) < (TW_END - TW_START + 1):
                    adj_row.append("")
                adjusted.append(adj_row)
            if adjusted:
                rng = (f"{col_letter(hist_t)}{CFG['DATA_START']}"
                       f":{col_letter(hist_t + TW_END - TW_START)}{CFG['DATA_END']}")
                formula_writes.append({
                    "range" : f"'{ws_title}'!{rng}",
                    "values": adjusted,
                })

        # ── HISTORY_FORMULA_REFS (USER_ENTERED) ──
        for (tmpl_row, tmpl_col), formula in HISTORY_FORMULA_REFS.items():
            h_col = paste_start + (tmpl_col - COL_HIST_START)
            rng_f = f"{col_letter(h_col)}{tmpl_row}"
            formula_writes.append({
                "range" : f"'{ws_title}'!{rng_f}",
                "values": [[formula]],
            })

    # ── batch write RAW (bookmarks) ──
    for i in range(0, len(raw_writes), CHUNK):
        try:
            _sheets_retry(ss.values_batch_update, {
                "valueInputOption": "RAW",
                "data"            : raw_writes[i:i + CHUNK],
            })
        except Exception:
            for entry in raw_writes[i:i + CHUNK]:
                sname = entry["range"].split("'!")[0].strip("'")
                ws_tmp = ws_dict.get(sname)
                if ws_tmp:
                    try:
                        short = entry["range"].split("'!")[-1]
                        _sheets_retry(ws_tmp.update, short, entry["values"],
                                      value_input_option="RAW")
                    except Exception:
                        pass

    # ── batch write USER_ENTERED (สูตร T-W + formula refs) ──
    for i in range(0, len(formula_writes), CHUNK):
        try:
            _sheets_retry(ss.values_batch_update, {
                "valueInputOption": "USER_ENTERED",
                "data"            : formula_writes[i:i + CHUNK],
            })
        except Exception:
            for entry in formula_writes[i:i + CHUNK]:
                sname = entry["range"].split("'!")[0].strip("'")
                ws_tmp = ws_dict.get(sname)
                if ws_tmp:
                    try:
                        short = entry["range"].split("'!")[-1]
                        _sheets_retry(ws_tmp.update, short, entry["values"],
                                      value_input_option="USER_ENTERED")
                    except Exception:
                        pass

    # v22: ไม่ auto-freeze block ใหม่ — ปล่อยให้เป็นสูตรไว้ก่อน
    # การ freeze จะเกิดขึ้นใน Phase 1b.5 ของการรัน ครั้งถัดไป (เดือนหน้า)
    # หรือกด "บันทึกค่านิ่ง" ด้วยตนเองเมื่อพร้อม

    _upd(0.99, "⏳ เกือบเสร็จแล้วค่ะ...")
    return results
