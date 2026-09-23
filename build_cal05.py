"""
Build the "CAL" commission-calculation sheet from a workbook that already
has T2-T18 (output of build_t_sheets.py), and reorder sheets for final delivery.

อ่านคอลัมน์จาก T sheets ด้วย header name (robust ต่อการขยับ column จาก preprocess)

CAL columns (1-indexed):
  1 สาขา, 2 วันที่เอกสาร, 3 เลขที่เอกสาร, 4 เลขที่เอกสารอ้างอิง,
  5 รหัสลูกหนี้, 6 ชื่อลูกค้า, 7 รหัสสินค้า, 8 cetagory (VLOOKUP pps05),
  9 รายละเอียด, 10 จำนวนที่สั่ง, 11 หน่วย, 12 ราคา/หน่วย, 13 อื่นๆ,
  14 ส่วนลด, 15 ราคารวม, 16 จำนวนเงินมัดจำ, 17 พนักงานขาย, 18 สาขา(dup),
  19 ราคาปลีก, 20 ราคาส่ง, 21 ราคากลาง=(S+T)/2, 22 ราคาขาย=O/J,
  23 อัตราคำนวญ, 24 ปรับคอมชุด, 25-26 (empty), 27 คอมพิเศษ, 28 คอมพิเศษ,
  29 เงินสด, 30 other, 31 อัตราค่าคอม, 32 ค่าคอม

Branch-conditional pps05 columns for retail(19)/wholesale(20):
  branches 10,11,13,14,15,18 -> pps05 H(8)/I(9)
  all other branches         -> pps05 F(6)/G(7)

Usage:
    python3 build_cal05.py <base_with_T_sheets.xlsx> <final_output.xlsx>
"""
import sys
import datetime
import openpyxl
from openpyxl.styles import PatternFill

IN_PATH = sys.argv[1]
OUT_PATH = sys.argv[2]

wb = openpyxl.load_workbook(IN_PATH)

# หาชีต PPS (pps05 หรือชื่อที่ขึ้นต้นด้วย pps)
PPS_SHEET = 'pps05'
for _sn in wb.sheetnames:
    if _sn.lower().startswith('pps'):
        PPS_SHEET = _sn
        break
PPS_Q = f"'{PPS_SHEET}'" if ' ' in PPS_SHEET else PPS_SHEET

cal_header = [
    'สาขา', 'วันที่เอกสาร', 'เลขที่เอกสาร', 'เลขที่เอกสารอ้างอิง',
    'รหัสลูกหนี้', 'ชื่อลูกค้า', 'รหัสสินค้า', 'cetagory',
    'รายละเอียด', 'จำนวนที่สั่ง', 'หน่วย', 'ราคา/หน่วย',
    'อื่นๆ', 'ส่วนลด', 'ราคารวม', 'จำนวนเงินมัดจำ',
    'พนักงานขาย', 'สาขา', 'ราคาปลีก', 'ราคาส่ง',
    'ราคากลาง', 'ราคาขาย', 'อัตราคำนวญ', 'ปรับคอมชุด',
    'คอมพิเศษ', 'คอมพิเศษ', 'เงินสด', 'เงินสด',
    'other', 'other', 'อัตราค่าคอม', 'ค่าคอม',
]

for old_name in ('Cal05', 'CAL', 'cal'):
    if old_name in wb.sheetnames:
        del wb[old_name]
ws = wb.create_sheet('CAL')

PINK = PatternFill(start_color='FFFFC0CB', end_color='FFFFC0CB', fill_type='solid')
RED  = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')

for c, header in enumerate(cal_header, start=1):
    ws.cell(row=1, column=c, value=header)


def build_t_col_map(src_ws):
    """Return dict: header_name (str) -> 0-based index ใน row tuple"""
    result = {}
    for cell in src_ws[1]:
        if cell.value is not None:
            key = str(cell.value)
            # ถ้ามี header ซ้ำ (เช่น สาขา 2 ตัว) เก็บตัวแรกเท่านั้น
            if key not in result:
                result[key] = cell.column - 1
    return result


r = 2
branch_sheets = [f'T{n}' for n in range(2, 19) if f'T{n}' in wb.sheetnames]

for sheet_name in branch_sheets:
    n = int(sheet_name[1:])
    src = wb[sheet_name]

    # Build column map จาก header row ของ T sheet
    t_map = build_t_col_map(src)

    def tv(row_vals, name):
        idx = t_map.get(name)
        if idx is None:
            return None
        return row_vals[idx] if idx < len(row_vals) else None

    # หา data_last
    data_last = 1
    for rr in range(2, src.max_row + 1):
        if src.cell(row=rr, column=1).value is None:
            break
        data_last = rr

    # Branch-conditional PPS columns
    if n in (10, 11, 13, 14, 15, 18):
        retail_col, wholesale_col = 8, 9
    else:
        retail_col, wholesale_col = 6, 7

    for row in src.iter_rows(min_row=2, max_row=data_last, values_only=True):
        # ─── คอลัมน์ 1-18: ข้อมูลจาก T sheet ─────────────────────
        branch_val = tv(row, 'สาขา')
        date_val   = tv(row, 'วันที่เอกสาร')
        prod_val   = tv(row, 'รหัสสินค้า')
        desc_val   = tv(row, 'รายละเอียด')
        qty_val    = tv(row, 'จำนวนที่สั่ง')
        total_val  = tv(row, 'ราคารวม')

        ws.cell(row=r, column=1,  value=branch_val)
        ws.cell(row=r, column=2,  value=date_val)
        if isinstance(date_val, datetime.datetime):
            ws.cell(row=r, column=2).number_format = 'dd/mm/yyyy'
        ws.cell(row=r, column=3,  value=tv(row, 'เลขที่เอกสาร'))
        ws.cell(row=r, column=4,  value=tv(row, 'เลขที่เอกสารอ้างอิง'))
        ws.cell(row=r, column=5,  value=tv(row, 'รหัสลูกหนี้'))
        ws.cell(row=r, column=6,  value=tv(row, 'ชื่อลูกค้า'))
        ws.cell(row=r, column=7,  value=prod_val)
        # col 8 = cetagory (VLOOKUP จาก pps05)
        ws.cell(row=r, column=8,  value=f'=VLOOKUP(G{r},{PPS_Q}!A:C,3,FALSE)')
        ws.cell(row=r, column=9,  value=desc_val)
        ws.cell(row=r, column=10, value=qty_val)
        ws.cell(row=r, column=11, value=tv(row, 'หน่วย'))
        ws.cell(row=r, column=12, value=tv(row, 'ราคา/หน่วย'))
        ws.cell(row=r, column=13, value=tv(row, 'อื่นๆ'))
        ws.cell(row=r, column=14, value=tv(row, 'ส่วนลด'))
        ws.cell(row=r, column=15, value=total_val)
        ws.cell(row=r, column=16, value=tv(row, 'จำนวนเงินมัดจำ'))
        ws.cell(row=r, column=17, value=tv(row, 'พนักงานขาย'))
        ws.cell(row=r, column=18, value=branch_val)  # สาขา dup

        # ─── คอลัมน์ 19-32: สูตรคำนวณ ──────────────────────────
        ws.cell(row=r, column=19,
                value=f"=VLOOKUP(G{r},{PPS_Q}!A:I,{retail_col},FALSE)")
        ws.cell(row=r, column=20,
                value=f"=VLOOKUP($G{r},{PPS_Q}!$A:$I,{wholesale_col},0)")
        ws.cell(row=r, column=21, value=f"=(S{r}+T{r})/2")
        ws.cell(row=r, column=22, value=f"=O{r}/J{r}")
        ws.cell(row=r, column=23,
                value=(f"=IF(V{r}<2000,(IF(V{r}>=U{r},2%,IF(V{r}<T{r},0%,1%))),"
                       f"(IF(V{r}>=U{r},1%,IF(V{r}<T{r},0,0.5%))))"))
        ws.cell(row=r, column=24,
                value=(f"=IF(V{r}>=3700,W{r},IF(H{r}=\"Skir-FR\",0.5%,"
                       f"IF(H{r}=\"Skir-RE\",0.5%,IF(H{r}=\"Skir-SK\",0.5%,"
                       f"IF(H{r}=\"Skir-CO\",0.5%,IF(H{r}=\"Skir-TO\",0.5%,"
                       f"IF(H{r}=\"Skir-TB\",0.5%,IF(H{r}=\"ProD\",0.5%,W{r}))))))))"))
        ws.cell(row=r, column=27,
                value=f"=VLOOKUP($G{r},'คอมพิเศษ-คอมเงินสด3'!$A:$C,3,FALSE)")
        ws.cell(row=r, column=29,
                value=f"=VLOOKUP($G{r},Item!$A$2:$C$6,3,FALSE)")
        ws.cell(row=r, column=30,
                value=f"=IFERROR(VLOOKUP($G{r},Item!$A$2:C{r+4},3,FALSE),0%)")
        ws.cell(row=r, column=31,
                value=f"=IF(Z{r}>=2%,Z{r},IF(AB{r}=\"N\",0%,IF(AD{r}>0%,AD{r},X{r})))")
        ws.cell(row=r, column=32,
                value=f"=IF(ISNUMBER(SEARCH(\"ซันรูฟ\",I{r})),0,AE{r}*O{r})")

        # ไฮไลต์แถวซันรูฟ
        if desc_val and 'ซันรูฟ' in str(desc_val):
            for c in range(1, 33):
                ws.cell(row=r, column=c).fill = RED

        # number format
        for c in range(23, 32):
            ws.cell(row=r, column=c).number_format = (
                'General' if c in (27, 28) else '0.00%'
            )

        r += 1

print(f'CAL sheet: {r - 2} rows')

# จัดเรียงลำดับชีต
desired_order = []
for n in range(2, 19):
    if f'T{n}' in wb.sheetnames:
        desired_order.append(f'T{n}')
for name in [PPS_SHEET, 'ยอดขาย', 'Item', 'คอมพิเศษ-คอมเงินสด3',
             'sum', 'sum (ตัดO2O)', 'CAL']:
    if name in wb.sheetnames and name not in desired_order:
        desired_order.append(name)
for name in wb.sheetnames:
    if name not in desired_order:
        desired_order.append(name)
wb._sheets = [wb[name] for name in desired_order]

wb.save(OUT_PATH)
print('saved', OUT_PATH)
