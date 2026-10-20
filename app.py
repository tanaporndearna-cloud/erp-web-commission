"""
TRC Commission App — Streamlit
อัปโหลดไฟล์ ERP → คิดค่าคอม → เขียนลง Google Sheet อัตโนมัติ
"""
import streamlit as st
import tempfile
import os
from datetime import date

# ===== Page Config =====
st.set_page_config(
    page_title="TRC Commission Calculator",
    page_icon="🏎️",
    layout="centered",
)

st.title("🏎️ TRC Commission Calculator")
st.caption("อัปโหลดไฟล์ ERP → คิดค่าคอม → เขียนลง Google Sheet อัตโนมัติ")

# ===== Sidebar Settings =====
with st.sidebar:
    st.header("⚙️ ตั้งค่า")
    erp_password = st.text_input(
        "รหัสผ่านไฟล์ ERP",
        value="weTeam15",
        type="password",
        help="รหัสผ่านสำหรับถอดรหัสไฟล์ ERP"
    )
    today = date.today()
    month = st.number_input("เดือน", min_value=1, max_value=12, value=today.month)
    year_be = st.number_input("ปี (พ.ศ.)", min_value=2560, max_value=2600, value=today.year + 543)

    st.divider()

# ===== File Upload =====
st.subheader("📁 อัปโหลดไฟล์")
col1, col2 = st.columns(2)
with col1:
    erp_file = st.file_uploader(
        "ไฟล์ ERP (บังคับ)",
        type=["xlsx"],
        help="Com_ERP*.xlsx ที่ได้จากระบบ eFlowsys"
    )
with col2:
    pps_file = st.file_uploader(
        "ไฟล์ PPS (ถ้ามี)",
        type=["xlsx"],
        help="PPS.xlsx ข้อมูลราคาสินค้า (ถ้าไม่มีจะใช้จากไฟล์ ERP)"
    )

# ===== Run Button =====
st.divider()
run_btn = st.button(
    "🚀 คิดค่าคอม",
    type="primary",
    disabled=(erp_file is None),
    use_container_width=True,
)

if erp_file is None:
    st.info("กรุณาอัปโหลดไฟล์ ERP ก่อนค่ะ")

# ===== Processing =====
if run_btn and erp_file is not None:
    # บันทึกไฟล์ upload ลง temp
    with tempfile.TemporaryDirectory() as tmp_upload:
        erp_path = os.path.join(tmp_upload, erp_file.name)
        with open(erp_path, "wb") as f:
            f.write(erp_file.getvalue())

        pps_path = None
        if pps_file:
            pps_path = os.path.join(tmp_upload, pps_file.name)
            with open(pps_path, "wb") as f:
                f.write(pps_file.getvalue())

        # --- Step 1: Run Pipeline ---
        st.subheader("⚙️ คิดค่าคอม...")
        status_box = st.empty()
        log_box = st.expander("แสดง Log", expanded=False)

        def update_status(msg):
            status_box.info(f"⏳ {msg}")

        try:
            from pipeline_runner import run_pipeline
            with st.spinner("กำลังประมวลผล..."):
                output_path, logs = run_pipeline(
                    erp_path=erp_path,
                    pps_path=pps_path,
                    password=erp_password,
                    month=int(month),
                    year=int(year_be),
                    progress_cb=update_status,
                )

            with log_box:
                for line in logs:
                    st.text(line)

            status_box.success("✅ คิดค่าคอมเสร็จแล้วค่ะ!")

            # --- Step 2: Read Result ---
            from google_writer import read_sum_sheet
            rows, dates, branch_totals = read_sum_sheet(output_path)
            st.success(f"อ่านข้อมูลได้ {len(rows)} แถว, {len(branch_totals)} สาขา")

            # แสดง Debug + Preview
            with st.expander("🔍 Debug: ค่าจริงใน xlsx (ก่อนเขียน Google Sheet)"):
                import pandas as pd

                # แสดงวันที่ที่ได้
                st.write("**วันที่ (5 วันแรก):**", dates[:5])

                # แสดงแถวทั้งหมดรวมถึงแถวรวม
                preview = []
                for r in rows:
                    # หาผลรวม daily
                    daily_sum = sum(v for v in r["daily"] if v)
                    preview.append({
                        "สาขา": r["branch"],
                        "ชื่อ": r["name"],
                        "daily_sum": daily_sum,
                        "com_pp": r["com_pp"],
                        "com_tot": r["com_tot"],
                        "sales": r["sales"],
                        "is_total": r["is_total"],
                    })
                df_preview = pd.DataFrame(preview)
                st.dataframe(df_preview, use_container_width=True)

                # สรุป: มีค่า > 0 ไหม
                non_zero = df_preview[df_preview["com_pp"] > 0]
                if len(non_zero) == 0:
                    st.error("⚠️ ทุกแถวมีค่า com_pp = 0 — ปัญหาอยู่ที่การคำนวณใน xlsx (LibreOffice / SUMIFS)")
                else:
                    st.success(f"✅ มีค่า > 0 จำนวน {len(non_zero)} แถว")

            # --- Step 3: Download + เก็บไฟล์ไว้ใน session ---
            st.subheader("⬇️ ดาวน์โหลดไฟล์ผลลัพธ์")
            with open(output_path, "rb") as f:
                output_bytes = f.read()

            # เก็บไว้ใน session_state เพื่อให้ปุ่มคัดลอกใช้ได้
            st.session_state["last_output_bytes"] = output_bytes
            st.session_state["last_output_name"] = os.path.basename(output_path)

            output_name = os.path.basename(output_path)
            st.download_button(
                label=f"⬇️ ดาวน์โหลด {output_name}",
                data=output_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            status_box.error(f"❌ เกิดข้อผิดพลาด: {e}")
            st.exception(e)

# ===== Commission SL =====
st.divider()
st.subheader("📊 Commission SL")

# แสดงสถานะไฟล์ด้านบน (ไม่ให้ทำให้คอลัมน์สูงไม่เท่ากัน)
if "last_output_name" in st.session_state:
    st.info(f"📁 ไฟล์พร้อม: **{st.session_state['last_output_name']}**")
else:
    st.warning("⚠️ ยังไม่มีไฟล์ — คิดค่าคอมก่อน หรืออัปโหลดในช่องซ้ายด้านล่าง")

_col_left, _col_right = st.columns(2)

# ── ซ้าย: คัดลอก sum(ตัดO2O) → Com ERP ──
with _col_left:
    st.markdown("**📋 คัดลอก sum(ตัดO2O) → Com ERP**")
    st.caption("ล้างชีทด้วย App Script ก่อน แล้วค่อยกดปุ่ม")

    upload_xlsx = st.file_uploader(
        "ไฟล์ xlsx ที่มีชีท sum(ตัดO2O)",
        type=["xlsx"],
        key="copy_xlsx"
    )
    if upload_xlsx:
        st.session_state["last_output_bytes"] = upload_xlsx.getvalue()
        st.session_state["last_output_name"] = upload_xlsx.name

    copy_btn = st.button(
        "📋 คัดลอก → Com ERP",
        type="secondary",
        use_container_width=True,
        disabled=("last_output_bytes" not in st.session_state),
    )

# ── ขวา: เติม H วันขาด/ลา ──
with _col_right:
    st.markdown("**🗓️ เติม H วันขาด/ลา → สรุปCom**")
    st.caption("ไม่อัปโหลดไฟล์ InOut = ใส่ H เฉพาะวันอาทิตย์")

    inout_file = st.file_uploader(
        "ไฟล์ InOutDailyReport.xlsx (ไม่บังคับ)",
        type=["xlsx"],
        key="inout_file",
        help="ถ้าไม่อัปโหลด ระบบจะใส่ H เฉพาะวันอาทิตย์เท่านั้น"
    )

    mark_h_btn = st.button(
        "🗓️ เติม H" + (" วันอาทิตย์ + ขาด/ลา" if inout_file else " วันอาทิตย์"),
        type="secondary",
        use_container_width=True,
    )

# ── Logic คัดลอก ──
if copy_btn and "last_output_bytes" in st.session_state:
    import tempfile as _tmp
    try:
        from google_writer import copy_sum_o2o_to_com_erp, write_summarize_com, read_sum_sheet, update_monthly_titles
        status_copy = st.empty()

        def copy_progress(msg):
            status_copy.info(f"⏳ {msg}")

        with _tmp.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tf.write(st.session_state["last_output_bytes"])
            tf_path = tf.name

        with st.spinner("กำลังเขียน Com ERP..."):
            result = copy_sum_o2o_to_com_erp(tf_path, progress_cb=copy_progress, month=int(month), year_be=int(year_be))
        status_copy.success(result)

        with st.spinner("กำลังเขียน สรุปCom..."):
            rows, dates, branch_totals = read_sum_sheet(tf_path)
            msg2 = write_summarize_com(branch_totals, dates=dates)
        st.success(msg2)

        with st.spinner("กำลังอัปเดตหัวชีท..."):
            msg3 = update_monthly_titles(int(month), int(year_be))
        st.success(msg3)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        st.exception(e)

# ── Logic เติม H ──
if mark_h_btn:
    import tempfile as _tmp2
    try:
        from google_writer import mark_absent_h, parse_absent_days
        status_h = st.empty()

        absent_days = {}
        if inout_file:
            with _tmp2.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf2:
                tf2.write(inout_file.getvalue())
                inout_path = tf2.name
            status_h.info("⏳ อ่านไฟล์ InOut...")
            absent_days = parse_absent_days(inout_path)
            status_h.info(f"⏳ พบ {len(absent_days)} รายการขาด/ลา — กำลังเติม H...")
        else:
            status_h.info("⏳ ไม่มีไฟล์ InOut — เติม H เฉพาะวันอาทิตย์...")

        def h_progress(msg):
            status_h.info(f"⏳ {msg}")

        with st.spinner("กำลังเติม H..."):
            result_h = mark_absent_h(absent_days, progress_cb=h_progress)
        status_h.success(result_h)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        st.exception(e)


# ===== Payroll Section =====
st.divider()

# เชื่อม gspread สำหรับ Payroll (lazy — เชื่อมเฉพาะตอนที่จะใช้)
PAYROLL_SHEET_ID  = "1x32_CtmkUFbhAHOJNVt2qBFbuIH0F8FQj5uEhLxqG6w"

@st.cache_resource
def get_gspread_client():
    """สร้าง gspread client จาก st.secrets["gcp_service_account"] (เหมือน google_writer.py)"""
    import gspread
    from google.oauth2.service_account import Credentials
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# เรนเดอร์ Payroll Page
try:
    from payroll_page import render_payroll_page
    gc = get_gspread_client()
    if gc is None:
        st.subheader("💰 จัดการเงินเดือน")
        st.error("❌ ไม่พบ credentials.json — กรุณาวางไฟล์ Service Account ในโฟลเดอร์เดียวกับ app.py")
        st.info("วิธีตั้งค่า: ไปที่ Google Cloud Console → Service Accounts → สร้าง Key (JSON) → วางไฟล์ชื่อ credentials.json")
    else:
        render_payroll_page(gc, PAYROLL_SHEET_ID)
except ImportError:
    st.subheader("💰 จัดการเงินเดือน")
    st.warning("⚠️ ไม่พบ payroll_page.py — กรุณาวางไฟล์ในโฟลเดอร์เดียวกับ app.py")


# ===== Payslip Section =====
st.divider()
st.subheader("🧾 สร้างสลีปเงินเดือน")
st.caption("เติมข้อมูลลงในไฟล์ Slip ที่สร้างไว้แล้ว พร้อมสูตร IMPORTRANGE อัตโนมัติ")

# ── คำแนะนำวิธีใช้ ──────────────────────────────────────────────────────────
with st.expander("📋 วิธีใช้ก่อนกด (คลิกเพื่อดู)", expanded=False):
    st.markdown("""
**ขั้นตอน (ทำครั้งแรกของแต่ละเดือน):**

1. เปิด Google Drive → หา **ไฟล์ Template Slip**
2. คลิกขวา → **Copy** (ทำสำเนา)
3. ตั้งชื่อไฟล์ใหม่ เช่น `Slip_ตุลาคม_2569`
4. Share ไฟล์นั้นให้ Service Account:
   - คลิกขวาที่ไฟล์ → **Share**
   - เพิ่ม email: `erp-streamlit@erp-commission.iam.gserviceaccount.com`
   - สิทธิ์: **Editor**
5. เปิดไฟล์ → Copy **File ID** จาก URL
   - URL ตัวอย่าง: `https://docs.google.com/spreadsheets/d/`**`1ABC...xyz`**`/edit`
   - เอาเฉพาะส่วนที่อยู่ระหว่าง `/d/` และ `/edit`
6. วาง ID ด้านล่างแล้วกดปุ่ม
""")

# ── Input File ID ────────────────────────────────────────────────────────────
# inject autocomplete=off เพื่อปิด browser suggestion
st.markdown("""
<style>
input[data-testid="stTextInput-RootElement"],
div[data-testid="stTextInput"] input {
    autocomplete: off !important;
}
</style>
<script>
window.addEventListener('load', function() {
    const inputs = window.parent.document.querySelectorAll('input[type="text"]');
    inputs.forEach(el => el.setAttribute('autocomplete', 'off'));
});
</script>
""", unsafe_allow_html=True)

target_file_id = st.text_input(
    "📋 วาง File ID ของไฟล์ Slip (09Payslip_template)",
    placeholder="เช่น 1nr0100FNOa9UA-TpXfEvgA1uaoscD41C8sbl9HbHz_l",
    help="ID ของไฟล์ Google Sheets ที่มีชีต 'Template' — copy จาก URL ระหว่าง /d/ กับ /edit",
    label_visibility="visible",
)

_col1, _col2 = st.columns(2)
with _col1:
    slip_btn = st.button(
        "🧾 เติมข้อมูล Slip",
        type="primary",
        use_container_width=True,
        disabled=not bool(target_file_id and target_file_id.strip()),
    )
with _col2:
    pdf_btn = st.button(
        "📄 Export PDF (ZIP)",
        type="primary",
        use_container_width=True,
        disabled=not bool(target_file_id and target_file_id.strip()),
    )

if slip_btn:
    file_id_clean = target_file_id.strip()
    try:
        from googleapiclient.discovery import build as gapi_build
        from google.oauth2.service_account import Credentials as SACredentials

        SLIP_SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict  = dict(st.secrets["gcp_service_account"])
        slip_creds  = SACredentials.from_service_account_info(creds_dict, scopes=SLIP_SCOPES)
        slip_sheets = gapi_build("sheets", "v4", credentials=slip_creds)
        slip_gc     = get_gspread_client()

        from generate_payslips import run_generate_payslips

        status_slip  = st.empty()
        progress_bar = st.progress(0)

        def slip_progress(msg, pct=None):
            status_slip.info(f"⏳ {msg}")
            if pct is not None:
                progress_bar.progress(min(pct, 100))

        with st.spinner("กำลังเติมข้อมูล Slip..."):
            file_id, file_name = run_generate_payslips(
                gc=slip_gc,
                sheets_svc=slip_sheets,
                target_file_id=file_id_clean,
                template_sheet_id=file_id_clean,
                progress_cb=slip_progress,
            )

        progress_bar.progress(100)
        status_slip.success(f"✅ เติมข้อมูล Slip เสร็จแล้วค่ะ! — {file_name}")
        st.markdown(
            f"[📄 เปิดไฟล์ {file_name}]"
            f"(https://docs.google.com/spreadsheets/d/{file_id})"
        )
        st.warning(
            "⚠️ เปิดไฟล์แล้วกด **Allow access** ครั้งแรก "
            "เพื่อให้ IMPORTRANGE ดึงข้อมูลได้ (กดครั้งเดียวครอบคลุมทุกชีต)\n\n"
            "📝 จากนั้นกรอก **M8 (เงินเดือน วัน)** ทีละชีตตามจริงค่ะ"
        )

    except ValueError as e:
        st.error(f"❌ {e}")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        st.exception(e)


# ===== PDF Export Section =====
if pdf_btn:
    pdf_file_id_clean = target_file_id.strip()
    try:
        from google.oauth2.service_account import Credentials as SACredentials
        from generate_payslips import export_payslip_pdfs

        SLIP_SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict  = dict(st.secrets["gcp_service_account"])
        pdf_creds   = SACredentials.from_service_account_info(creds_dict, scopes=SLIP_SCOPES)

        pdf_status   = st.empty()
        pdf_progress = st.progress(0)

        def pdf_progress_cb(msg, pct=None):
            pdf_status.info(f"⏳ {msg}")
            if pct is not None:
                pdf_progress.progress(min(int(pct), 100))

        with st.spinner("กำลัง Export PDF..."):
            zip_bytes, month_en, year_be, err_list = export_payslip_pdfs(
                creds=pdf_creds,
                target_file_id=pdf_file_id_clean,
                progress_cb=pdf_progress_cb,
            )

        pdf_progress.progress(100)
        zip_name = f"Slip_{month_en}_{year_be}.zip"
        if err_list:
            pdf_status.warning(f"⚠️ Export เสร็จ — มี {len(err_list)} ไฟล์ที่ error (ดูรายละเอียดด้านล่าง)")
            with st.expander(f"❌ รายการ error ({len(err_list)} รายการ)", expanded=True):
                for e in err_list:
                    st.text(e)
        else:
            pdf_status.success(f"✅ Export PDF เสร็จแล้วค่ะ! — {zip_name}")

        st.download_button(
            label=f"⬇️ ดาวน์โหลด {zip_name}",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
            use_container_width=True,
        )

    except ValueError as e:
        st.error(f"❌ {e}")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        st.exception(e)


# ===== Footer =====
st.divider()
st.caption("TRC Motorsport · Commission Calculator · Powered by Streamlit")
