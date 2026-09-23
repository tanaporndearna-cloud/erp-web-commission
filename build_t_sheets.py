"""
Build T2-T18 commission sheets from a decrypted/preprocessed Com_ERP xlsx
(sheets named "2".."18").

รองรับทั้ง 2 format:
  - format เก่า: header ที่ row 1 เป็น standard headers (สาขา, วันที่เอกสาร, ...)
  - format ใหม่ (ผ่าน preprocess_erp.py): header ที่ row 1 เป็น ERP headers
คอลัมน์สำคัญถูก detect ด้วย header name (ไม่ hardcode ตำแหน่ง)

Usage:
    python3 build_t_sheets.py <input.xlsx> <output.xlsx>
"""
import sys
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from copy import copy

IN_PATH = sys.argv[1]
OUT_PATH = sys.argv[2]

wb = openpyxl.load_workbook(IN_PATH)

# Rename sheets 2-18 -> T2-T18
for n in range(2, 19):
    if str(n) in wb.sheetnames:
        wb[str(n)].title = f'T{n}'

yellow   = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
orange   = PatternFill(start_color='FFFFC000', end_color='FFFFC000', fill_type='solid')
red_font = Font(color='FFFF0000')

if 'ยอดขาย' not in wb.sheetnames:
    print('WARNING: sheet "ยอดขาย" not found - cannot compute diff/ยอดขาย references')
    sys.exit(1)

yk = wb['ยอดขาย']


def get_header_map(ws):
    """Return dict: header_name -> col_index (1-based) จาก row 1"""
    return {str(ws.cell(row=1, column=c).value): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(row=1, column=c).value is not None}


def num(x):
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x)


def build_formula(old_val, sumQ, diff):
    """Build deposit adjustment formula: old_val + sumQ [+ diff if diff < 0]"""
    terms = []
    if old_val:
        terms.append(old_val)
    terms.append(sumQ)
    if diff < 0:
        terms.append(diff)
    s = num(terms[0])
    for t in terms[1:]:
        if t >= 0:
            s += '+' + num(t)
        else:
            s += '-' + num(-t)
    return '=' + s


for n in range(2, 19):
    sheet_name = f'T{n}'
    if sheet_name not in wb.sheetnames:
        print(f'{sheet_name}: sheet not found, skipping')
        continue
    ws = wb[sheet_name]

    # Detect column positions from header row
    hmap = get_header_map(ws)

    # คอลัมน์หลักที่ต้องการ — ลอง header ไทย ก่อน แล้ว fallback default
    COL_PROD  = (hmap.get('รหัสสินค้า')  or 7)   # Deposit check
    COL_TOTAL = (hmap.get('ราคารวม')     or 15)   # sumP (ราคารวม)
    COL_DEP   = (hmap.get('จำนวนเงินมัดจำ') or 16)  # sumQ (มัดจำ)
    # คอลัมน์หมายเหตุ: ถัดจากรหัสสินค้า 1 คอลัมน์
    COL_NOTE  = COL_PROD + 1
    # คอลัมน์ label: ก่อน ราคารวม 1 คอลัมน์
    COL_LABEL = COL_TOTAL - 1
    # คอลัมน์ "เช็ค sum>>": ก่อน ราคารวม 2 คอลัมน์
    COL_CHECK = COL_TOTAL - 2

    TOTAL_L = get_column_letter(COL_TOTAL)
    DEP_L   = get_column_letter(COL_DEP)

    # หา data_last: แถวสุดท้ายที่ col 1 (สาขา) มีค่า
    data_last = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value is None:
            break
        data_last = r

    sumP = round(sum((ws.cell(row=r, column=COL_TOTAL).value or 0)
                     for r in range(2, data_last + 1)), 2)
    sumQ = round(sum((ws.cell(row=r, column=COL_DEP).value or 0)
                     for r in range(2, data_last + 1)), 2)

    yk_row = n + 4
    yk_D   = yk.cell(row=yk_row, column=4).value
    diff   = round((yk_D or 0) - (sumP + sumQ), 2)

    # Deposit adjustment
    deposit_rows = [r for r in range(2, data_last + 1)
                    if ws.cell(row=r, column=COL_PROD).value == 'Deposit']
    if not deposit_rows:
        print(f'{sheet_name}: WARNING no Deposit row found, skipping adjustment')
        target = None
    else:
        target   = deposit_rows[-1]
        old_val  = ws.cell(row=target, column=COL_TOTAL).value or 0
        formula  = build_formula(old_val, sumQ, diff)
        dcell    = ws.cell(row=target, column=COL_TOTAL, value=formula)
        dcell.fill = copy(yellow)
        dcell.font = copy(red_font)

    # Unmerge ส่วน summary ก่อนเขียน
    r1 = data_last + 2
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= r1 - 1:
            ws.unmerge_cells(str(mc))

    # แถว r1: "เช็ค sum>>"
    ws.cell(row=r1, column=COL_CHECK, value='เช็ค sum>>')
    c_sumP = ws.cell(row=r1, column=COL_TOTAL,
                     value=f'=SUM({TOTAL_L}2:{TOTAL_L}{data_last})')
    c_sumP.fill = copy(orange)
    c_sumQ = ws.cell(row=r1, column=COL_DEP,
                     value=f'=SUM({DEP_L}2:{DEP_L}{data_last})')
    c_sumQ.fill = copy(orange)

    # แถว r1+1: ยอดERPไม่รวมมัดจำ
    ws.cell(row=r1 + 1, column=COL_LABEL, value='ยอดERPไม่รวมมัดจำ')
    ws.cell(row=r1 + 1, column=COL_TOTAL, value=sumP)

    # แถว r1+2: ยอดERPรวมมัดจำ
    ws.cell(row=r1 + 2, column=COL_LABEL, value='ยอดERPรวมมัดจำ')
    c_erp = ws.cell(row=r1 + 2, column=COL_TOTAL,
                    value=f'={TOTAL_L}{r1+1}+{DEP_L}{r1}')
    c_erp.fill = copy(orange)

    # แถว r1+3: ยอดขาย
    ws.cell(row=r1 + 3, column=COL_LABEL, value='ยอดขาย')
    ws.cell(row=r1 + 3, column=COL_TOTAL, value=f"='ยอดขาย'!D{yk_row}")

    # แถว r1+5: ยอดขาย - ยอดERPรวมมัดจำ
    ws.cell(row=r1 + 5, column=COL_LABEL, value='ยอดขาย-ยอดERPรวมมัดจำ')
    ws.cell(row=r1 + 5, column=COL_TOTAL,
            value=f'={TOTAL_L}{r1+3}-{TOTAL_L}{r1+2}')

    # หมายเหตุ
    if diff == 0:
        note = 'ยอดตรง'
    else:
        note = f'ปรับยอดมัดจำ {diff:+,.2f} บาท ให้ยอดตรง'
    note_cell = ws.cell(row=r1, column=COL_NOTE, value=f'หมายเหตุ: {note}')
    note_cell.font = copy(red_font)

    print(f'{sheet_name}: data_last={data_last} sumP={sumP} sumQ={sumQ} '
          f'yk_D={yk_D} diff={diff}'
          + (f' adj_row={target}' if target else ' (no deposit row)')
          + f' [ราคารวม=col{COL_TOTAL}, มัดจำ=col{COL_DEP}]')

wb.save(OUT_PATH)
print('saved', OUT_PATH)
