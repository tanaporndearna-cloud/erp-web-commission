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
    if progress_cb:
        progress_cb("รีเซ็ตรูปแบบตัวเลข...")
    ws_dst.format(
        f"A1:{gspread.utils.rowcol_to_a1(total_rows, max_cols)}",
        {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}
    )

    return f"✅ คัดลอก {total_rows} แถวจาก sum(ตัดO2O) → Com ERP สำเร็จ"
