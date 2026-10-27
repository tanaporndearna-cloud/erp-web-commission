"""
อ่านไฟล์ PPS (รองรับ .xlsx, .xls binary, และ .xls ที่จริงเป็น TSV)
แล้ว inject เป็นชีต PPS มาตรฐานใน ERP xlsx

Usage:
    python3 build_pps.py <pps_source.xlsx|xls> <erp_working.xlsx> <output.xlsx>
"""
import sys
import csv
import io
import openpyxl
from openpyxl.styles import Font

PPS_PATH = sys.argv[1]
ERP_PATH = sys.argv[2]
OUT_PATH = sys.argv[3]


# ── Cell wrapper สำหรับข้อมูลที่อ่านจาก TSV / xlrd ──
class _Cell:
    def __init__(self, value):
        self.value = value


class ListSheetWrapper:
    """ห่อ list-of-lists ให้ใช้ .cell(row, col) แบบ openpyxl"""
    def __init__(self, rows):
        self._rows = rows
        self.max_row = len(rows)
        self.max_column = max((len(r) for r in rows), default=0)

    def cell(self, row, column):
        r, c = row - 1, column - 1
        if r < 0 or r >= len(self._rows):
            return _Cell(None)
        row_data = self._rows[r]
        if c < 0 or c >= len(row_data):
            return _Cell(None)
        v = row_data[c]
        # แปลง string ตัวเลขเป็น int/float
        if isinstance(v, str):
            v = v.strip()
            if v == '':
                v = None
            else:
                try:
                    v = int(v)
                except ValueError:
                    try:
                        v = float(v)
                    except ValueError:
                        pass
        return _Cell(v)


def read_pps(path):
    """อ่าน PPS รองรับ .xlsx, real-.xls, และ TSV-disguised-as-.xls"""
    is_xls = path.lower().endswith('.xls')

    # ── ลอง openpyxl ก่อน (ใช้ได้กับ .xlsx หรือ .xls ที่จริงเป็น xlsx) ──
    if not is_xls:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        print(f'  {ws.max_row} rows (openpyxl)')
        return ws

    # ── ลอง xlrd (real Excel binary .xls) ──
    try:
        import xlrd
        book = xlrd.open_workbook(path)
        sheet = book.sheets()[0]
        print(f'  {sheet.nrows} rows (xlrd binary)')

        class XlsWrapper:
            def __init__(self, s):
                self._s = s
                self.max_row = s.nrows
                self.max_column = s.ncols
            def cell(self, row, column):
                r, c = row - 1, column - 1
                if r < 0 or r >= self._s.nrows or c < 0 or c >= self._s.ncols:
                    return _Cell(None)
                v = self._s.cell_value(r, c)
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                if v == '':
                    v = None
                return _Cell(v)
        return XlsWrapper(sheet)

    except Exception:
        pass  # ไม่ใช่ binary xls → ลอง TSV

    # ── TSV / CSV text file ที่ตั้งนามสกุล .xls ──
    for enc in ('cp874', 'utf-8-sig', 'utf-8', 'iso-8859-1'):
        try:
            with open(path, encoding=enc, newline='') as f:
                sample = f.read(1024)
            # ตรวจว่าเป็น tab-separated หรือ comma-separated
            dialect = 'excel-tab' if sample.count('\t') > sample.count(',') else 'excel'
            rows = []
            with open(path, encoding=enc, newline='') as f:
                reader = csv.reader(f, dialect=dialect)
                for row in reader:
                    rows.append(row)
            if rows:
                print(f'  {len(rows)} rows (TSV/CSV, encoding={enc})')
                return ListSheetWrapper(rows)
        except Exception:
            continue

    raise RuntimeError(f'ไม่สามารถอ่านไฟล์ PPS: {path}')


# ── อ่าน PPS ──
print(f'Reading PPS: {PPS_PATH}')
pps_ws = read_pps(PPS_PATH)
pps_last = pps_ws.max_row

# ── auto-detect column จาก header ──
HEADER_MAP = {
    'Product Code':      'product_code',
    'Description':       'description',
    'Category':          'category',
    'Branch Province':   'branch_province',
    'Retail Price':      'retail_price',
    'Wholesale General': 'wholesale_gen',
}
col_idx = {}
for c in range(1, pps_ws.max_column + 1):
    h = pps_ws.cell(row=1, column=c).value
    if h and str(h).strip() in HEADER_MAP:
        col_idx[HEADER_MAP[str(h).strip()]] = c

SRC_PRODUCT_CODE    = col_idx.get('product_code',    3)
SRC_DESCRIPTION     = col_idx.get('description',     4)
SRC_CATEGORY        = col_idx.get('category',        2)
SRC_BRANCH_PROVINCE = col_idx.get('branch_province', 15)
SRC_RETAIL_PRICE    = col_idx.get('retail_price',    16)
SRC_WHOLESALE_GEN   = col_idx.get('wholesale_gen',   20)

print(f'  Columns: ProductCode={SRC_PRODUCT_CODE}, Desc={SRC_DESCRIPTION}, '
      f'Cat={SRC_CATEGORY}, BranchProv={SRC_BRANCH_PROVINCE}, '
      f'Retail={SRC_RETAIL_PRICE}, Wholesale={SRC_WHOLESALE_GEN}')

# ── โหลด ERP ──
print(f'Loading ERP: {ERP_PATH}')
wb = openpyxl.load_workbook(ERP_PATH)

for name in list(wb.sheetnames):
    if 'pps' in name.lower():
        del wb[name]
        print(f'  Removed old sheet: "{name}"')

ws = wb.create_sheet('PPS')

HEADERS = ['Product Code', 'Description', 'Category', '',
           'Branch Province', 'Retail Price', 'Wholesale General',
           'P Retail Price', 'Wholesale General P']
for c, h in enumerate(HEADERS, 1):
    cell = ws.cell(row=1, column=c, value=h if h else None)
    if h:
        cell.font = Font(bold=True)

out_row = 2
for r in range(2, pps_last + 1):
    if pps_ws.cell(row=r, column=SRC_PRODUCT_CODE).value is None:
        continue
    ws.cell(row=out_row, column=1, value=pps_ws.cell(row=r, column=SRC_PRODUCT_CODE).value)
    ws.cell(row=out_row, column=2, value=pps_ws.cell(row=r, column=SRC_DESCRIPTION).value)
    ws.cell(row=out_row, column=3, value=pps_ws.cell(row=r, column=SRC_CATEGORY).value)
    ws.cell(row=out_row, column=5, value=pps_ws.cell(row=r, column=SRC_BRANCH_PROVINCE).value)
    ws.cell(row=out_row, column=6, value=pps_ws.cell(row=r, column=SRC_RETAIL_PRICE).value)
    ws.cell(row=out_row, column=7, value=pps_ws.cell(row=r, column=SRC_WHOLESALE_GEN).value)
    ws.cell(row=out_row, column=8, value=f'=IF(E{out_row}>0,E{out_row},F{out_row})')
    ws.cell(row=out_row, column=9, value=f'=(G{out_row}+H{out_row})/2')
    out_row += 1

print(f'  Written {out_row - 2} rows to "PPS"')
wb.save(OUT_PATH)
print(f'Saved: {OUT_PATH}')
