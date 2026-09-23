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
    write_to_sheet = st.toggle("📤 เขียนลง Google Sheet", value=True)
    if write_to_sheet:
        st.info("จะเขียนข้อมูลลงชีท **Com ERP** และ **สรุปCom** ใน test commission SL/2026")

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

            # --- Step 3: Write to Google Sheet ---
            if write_to_sheet:
                st.subheader("📤 เขียนลง Google Sheet...")

                try:
                    from google_writer import write_com_erp, write_summarize_com

                    prog = st.progress(0, "กำลังเขียนชีท Com ERP...")

                    def erp_progress(current, total):
                        prog.progress(current / total, f"เขียน Com ERP: {current}/{total} แถว")

                    with st.spinner("เขียนชีท Com ERP..."):
                        msg1 = write_com_erp(rows, dates, erp_progress)
                    st.success(msg1)

                    with st.spinner("เขียนชีท สรุปCom..."):
                        msg2 = write_summarize_com(branch_totals)
                    st.success(msg2)

                    prog.empty()

                except Exception as e:
                    st.error(f"เขียน Google Sheet ไม่สำเร็จ: {e}")
                    st.warning("กรุณาตรวจสอบ Service Account credentials ใน secrets.toml")

            # --- Step 4: Download + เก็บไฟล์ไว้ใน session ---
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

# ===== Copy sum(ตัดO2O) → Com ERP =====
st.divider()
st.subheader("📋 คัดลอก sum(ตัดO2O) → Com ERP")
st.caption("อ่านข้อมูลจากไฟล์ xlsx แล้วเขียนลง Com ERP เป็นค่าธรรมดา (ไม่มีสูตร) — ล้างชีทก่อนด้วย App Script แล้วค่อยกดปุ่มนี้")

# แสดงสถานะไฟล์ที่พร้อมใช้
if "last_output_name" in st.session_state:
    st.info(f"📁 ไฟล์พร้อม: **{st.session_state['last_output_name']}**")
else:
    st.warning("⚠️ ยังไม่มีไฟล์ — กรุณาคิดค่าคอมก่อน หรืออัปโหลดไฟล์ด้านล่าง")

# อัปโหลดไฟล์เพิ่มเติม (กรณีเปิด app ใหม่)
upload_xlsx = st.file_uploader(
    "หรืออัปโหลดไฟล์ xlsx ที่มีชีท sum(ตัดO2O)",
    type=["xlsx"],
    key="copy_xlsx"
)
if upload_xlsx:
    st.session_state["last_output_bytes"] = upload_xlsx.getvalue()
    st.session_state["last_output_name"] = upload_xlsx.name

copy_btn = st.button(
    "📋 คัดลอก sum(ตัดO2O) → Com ERP",
    type="secondary",
    use_container_width=True,
    disabled=("last_output_bytes" not in st.session_state),
)

if copy_btn and "last_output_bytes" in st.session_state:
    import tempfile as _tmp
    try:
        from google_writer import copy_sum_o2o_to_com_erp, write_summarize_com, read_sum_sheet
        status_copy = st.empty()

        def copy_progress(msg):
            status_copy.info(f"⏳ {msg}")

        # เขียนไฟล์ bytes ลง temp file ชั่วคราว
        with _tmp.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tf.write(st.session_state["last_output_bytes"])
            tf_path = tf.name

        # Step 1: คัดลอก sum(ตัดO2O) → Com ERP
        with st.spinner("กำลังเขียน Com ERP..."):
            result = copy_sum_o2o_to_com_erp(tf_path, progress_cb=copy_progress)
        status_copy.success(result)

        # Step 2: เขียนยอดรายสาขา → สรุปCom
        with st.spinner("กำลังเขียน สรุปCom..."):
            rows, dates, branch_totals = read_sum_sheet(tf_path)
            msg2 = write_summarize_com(branch_totals)
        st.success(msg2)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        st.exception(e)

# ===== Footer =====
st.divider()
st.caption("TRC Motorsport · Commission Calculator · Powered by Streamlit")
