"""
Rebuild the "sum" / "sum (ตัดO2O)" sheet's employee rows from scratch, using
ONLY the employee names that actually appear in the CAL sheet (column Q,
พนักงานขาย) for each branch (CAL column A, สาขา = 'T2'..'T18'). This is
important because the roster of employees can change month to month - we
must not reuse last month's employee list.

For each branch (in CAL order of first appearance), one row per distinct
employee is written, followed by a "รวม" (total) row and a small diff row,
matching the original sum-sheet layout:

  - col A = branch ('T2'..'T18')
  - col B = employee name (or 'รวม' for the total row)
  - cols C..AG (3..33) = one column per date (row 3 holds the date headers,
    kept from the original sheet) with
      =SUMIFS(CAL!$AF:$AF,CAL!$A:$A,sum!$A{r},CAL!$Q:$Q,sum!$B{r},CAL!$B:$B,sum!{col}$3)
  - col AI (35) = คอมรายคน = SUM(C{r}:AH{r})
  - รวม row: C..AG = SUM of the employee rows above, AI = SUM of their AI,
      AJ (36) คอมรวม = SUMIFS(CAL!$AF:$AF,CAL!$A:$A,sum!A{r})
      AK (37) ยอดขาย = SUMIFS(CAL!$O:$O,CAL!$A:$A,sum!$A{r})
      AL (38) % Com  = AJ/AK
  - diff row directly below รวม row: col AJ (36) = '=AJ{rvm}-AI{rvm}'

A final "Grand Total" row sums AF (commission) and O (ยอดขาย) directly across
all of CAL, so it's correct regardless of how many branches/employees exist.

Usage:
    python3 build_sum.py <base_with_CAL.xlsx> <final_output.xlsx> [sheet_name]

sheet_name defaults to 'sum'. Run again with 'sum (ตัดO2O)' for the second sheet.
NOTE: no O2O filter is applied yet - 'sum (ตัดO2O)' will currently be
identical to 'sum' until the O2O criterion is defined.

NEW: If the sheet does not exist in the file, it will be created automatically
with header rows built from CAL data (dates in row 3, labels in rows 1-2, 4).
"""
import sys
import re
import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font
from collections import OrderedDict, Counter

THAI_MONTHS = {
    1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
    5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
    9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
}

def get_report_month_year(dates):
    """หาเดือน/ปีหลักของรายงานจาก list ของวันที่ใน CAL
    ใช้เดือนที่โผล่มากที่สุด ถ้า tie ใช้เดือนล่าสุด"""
    month_counts = Counter()
    for d in dates:
        if hasattr(d, 'month'):
            month_counts[(d.year, d.month)] += 1
    if not month_counts:
        return None, None
    # เดือนที่มีวันมากสุด ถ้าเท่ากันเอาล่าสุด
    (year, month) = max(month_counts, key=lambda k: (month_counts[k], k[0], k[1]))
    be_year = year + 543
    return THAI_MONTHS.get(month, str(month)), be_year

IN_PATH = sys.argv[1]
OUT_PATH = sys.argv[2]
SHEET_NAME = sys.argv[3] if len(sys.argv) > 3 else 'sum'

wb = openpyxl.load_workbook(IN_PATH)

if 'CAL' not in wb.sheetnames:
    print('ERROR: CAL sheet not found - run build_cal05.py first')
    sys.exit(1)

cal = wb['CAL']
orange_fill = PatternFill(start_color='FFFFC000', end_color='FFFFC000', fill_type='solid')
NUM_FMT = '#,##0.00;-#,##0.00'
FONT_NORMAL = Font(bold=False)
FONT_BOLD = Font(bold=True)

sn = SHEET_NAME
sn_q = f"'{sn}'" if (' ' in sn or '(' in sn or ')' in sn) else sn

# ตรรกะ O2O: ชีต sum (ตัดO2O) จะใส่ 0 สำหรับพนักงานที่ไม่ใช่ของสาขานั้น
IS_O2O = (SHEET_NAME == 'sum (ตัดO2O)')

def emp_branch_num(name):
    """ดึงเลขสาขาจากชื่อพนักงาน เช่น 'อภิชญา (เบลT6)' -> 6, ไม่มี -> None"""
    m = re.search(r'T(\d+)', str(name))
    return int(m.group(1)) if m else None

# 1. branch -> ordered list of distinct employees, from CAL columns A (1) and Q (17)
branch_emps = OrderedDict()
cal_last = cal.max_row
for r in range(2, cal_last + 1):
    branch = cal.cell(row=r, column=1).value
    emp = cal.cell(row=r, column=17).value
    if not branch or not emp:
        continue
    branch_emps.setdefault(branch, OrderedDict())
    branch_emps[branch].setdefault(emp, True)

# 2. collect unique sorted dates from CAL column B for header row 3
#    แล้วขยายเป็น full range รวมวันอาทิตย์และวันที่ไม่มียอด
cal_dates_raw = []
seen_dates = set()
for r in range(2, cal_last + 1):
    d = cal.cell(row=r, column=2).value
    if d and d not in seen_dates:
        seen_dates.add(d)
        cal_dates_raw.append(d)
cal_dates_raw.sort(key=lambda x: (x if not hasattr(x, 'toordinal') else x.toordinal()))

# สร้าง full date range min→max รวมทุกวันในช่วง
if cal_dates_raw and hasattr(cal_dates_raw[0], 'date'):
    date_only = [d.date() for d in cal_dates_raw]
    min_d, max_d = min(date_only), max(date_only)
    cal_dates = []
    cur = min_d
    while cur <= max_d:
        cal_dates.append(datetime.datetime(cur.year, cur.month, cur.day))
        cur += datetime.timedelta(days=1)
    print(f'Date range: {min_d} → {max_d} ({len(cal_dates)} days, {len(cal_dates_raw)} with sales)')
else:
    cal_dates = cal_dates_raw

DATE_ROW = 3
FIRST_DATA_ROW = 5

# 3. Get or create the sheet
if SHEET_NAME not in wb.sheetnames:
    print(f'Sheet "{SHEET_NAME}" not found — creating from scratch')
    ws = wb.create_sheet(SHEET_NAME)

    # Row 1: title พร้อมชื่อเดือน/ปี
    month_name, be_year = get_report_month_year(cal_dates)
    if month_name:
        title = f'สรุปคอมมิชชั่น เดือน {month_name} {be_year}'
    else:
        title = 'สรุปคอมมิชชั่น'
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12)
    print(f'Title: {title}')

    # Row 2: ว่าง

    # Row 3: ใส่วันที่ทุกวันในช่วง (รวมวันอาทิตย์และวันที่ไม่มียอด)
    ws.cell(row=DATE_ROW, column=1, value='สาขา')
    ws.cell(row=DATE_ROW, column=2, value='ชื่อ-สกุล')
    date_cols = []
    for i, d in enumerate(cal_dates):
        c = 3 + i  # เริ่มที่ col C (3)
        cell = ws.cell(row=DATE_ROW, column=c, value=d)
        if hasattr(d, 'strftime'):
            cell.number_format = 'DD/MM/YYYY'
        date_cols.append(c)

    # Row 4: ว่าง

    print(f'Created new sheet "{SHEET_NAME}" with {len(date_cols)} date columns in row 3')
else:
    ws = wb[sn]
    # อัปเดต row 3 ด้วย full date range (ล้างของเก่าก่อน)
    for c in range(3, ws.max_column + 1):
        ws.cell(row=DATE_ROW, column=c).value = None
    date_cols = []
    for i, d in enumerate(cal_dates):
        c = 3 + i
        cell = ws.cell(row=DATE_ROW, column=c, value=d)
        if hasattr(d, 'strftime'):
            cell.number_format = 'DD/MM/YYYY'
        date_cols.append(c)
    print(f'Sheet "{SHEET_NAME}" found — updated to {len(date_cols)} date columns in row 3')

# 4. Clear everything from row 5 down
old_max_row = ws.max_row
old_max_col = ws.max_column

# remove merged cells that overlap the area we're about to rewrite
for mc in list(ws.merged_cells.ranges):
    if mc.min_row >= FIRST_DATA_ROW:
        ws.unmerge_cells(str(mc))

no_fill = PatternFill(fill_type=None)
for r in range(FIRST_DATA_ROW, old_max_row + 1):
    for c in range(1, old_max_col + 1):
        cell = ws.cell(row=r, column=c)
        cell.value = None
        cell.fill = no_fill  # clear any leftover highlight from old template
        cell.font = FONT_NORMAL  # employee rows should not be bold

# คำนวณ column ของ summary — เว้นว่าง 1 ช่องหลังวันสุดท้าย
AI_COL = max(date_cols) + 2   # คอมรายคน
AJ_COL = AI_COL + 1           # คอมรวม
AK_COL = AI_COL + 2           # ยอดขาย
AL_COL = AI_COL + 3           # %Com

ai_let        = openpyxl.utils.get_column_letter(AI_COL)
aj_let        = openpyxl.utils.get_column_letter(AJ_COL)
ak_let        = openpyxl.utils.get_column_letter(AK_COL)
al_let        = openpyxl.utils.get_column_letter(AL_COL)
last_date_let = openpyxl.utils.get_column_letter(max(date_cols))

# ใส่ label summary columns ใน row 3
ws.cell(row=DATE_ROW, column=AI_COL, value='คอมรายคน')
ws.cell(row=DATE_ROW, column=AJ_COL, value='คอมรวม')
ws.cell(row=DATE_ROW, column=AK_COL, value='ยอดขาย')
ws.cell(row=DATE_ROW, column=AL_COL, value='% Com')

# ตั้ง column width 15.00 ทุกคอลัมน์ (A ถึง AL_COL)
for c in range(1, AL_COL + 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 15.0

# branches in CAL appear in order T2..T18 already (since CAL is built T2..T18 in order)
branches = [b for b in branch_emps.keys()]

row = FIRST_DATA_ROW
for branch in branches:
    # หาเลขสาขา เช่น 'T6' -> 6
    branch_n = int(branch[1:]) if re.match(r'^T\d+$', branch) else None
    emps = list(branch_emps[branch].keys())
    emp_rows = []
    for emp in emps:
        ws.cell(row=row, column=1, value=branch)
        ws.cell(row=row, column=2, value=emp)

        # ตรวจว่าพนักงานนี้เป็นของสาขานี้จริงหรือเปล่า (สำหรับชีต O2O)
        # ถ้าไม่มีเลขสาขาในชื่อ → คำนวณปกติ
        # ถ้ามีเลขสาขาในชื่อ แต่ไม่ตรงกับสาขาปัจจุบัน → 0
        emp_n = emp_branch_num(emp)
        is_own_branch = (emp_n is None) or (emp_n == branch_n)
        use_formula = (not IS_O2O) or is_own_branch

        for c in date_cols:
            col_letter = openpyxl.utils.get_column_letter(c)
            if use_formula:
                val = (f"=SUMIFS(CAL!$AF:$AF,CAL!$A:$A,{sn_q}!$A{row},"
                       f"CAL!$Q:$Q,{sn_q}!$B{row},CAL!$B:$B,{sn_q}!{col_letter}$3)")
            else:
                val = 0
            dc = ws.cell(row=row, column=c, value=val)
            dc.number_format = NUM_FMT
        ai_cell = ws.cell(row=row, column=AI_COL,
                          value=f"=SUM(C{row}:{last_date_let}{row})")
        ai_cell.number_format = NUM_FMT
        emp_rows.append(row)
        row += 1

    # รวม row
    rvm = row
    ws.cell(row=rvm, column=1, value=branch)
    ws.cell(row=rvm, column=2, value='รวม')
    r0, r1 = emp_rows[0], emp_rows[-1]
    for c in date_cols:
        col_letter = openpyxl.utils.get_column_letter(c)
        sc = ws.cell(row=rvm, column=c, value=f"=SUM({col_letter}{r0}:{col_letter}{r1})")
        sc.number_format = NUM_FMT
    rvm_ai = ws.cell(row=rvm, column=AI_COL,
                     value=f"=SUM({ai_let}{r0}:{ai_let}{r1})")
    rvm_ai.number_format = NUM_FMT
    rvm_aj = ws.cell(row=rvm, column=AJ_COL,
                     value=f"=SUMIFS(CAL!$AF:$AF,CAL!$A:$A,{sn_q}!A{rvm})")
    rvm_aj.number_format = NUM_FMT
    rvm_ak = ws.cell(row=rvm, column=AK_COL,
                     value=f"=SUMIFS(CAL!$O:$O,CAL!$A:$A,{sn_q}!$A{rvm})")
    rvm_ak.number_format = NUM_FMT
    al_cell = ws.cell(row=rvm, column=AL_COL,
                      value=f"={aj_let}{rvm}/{ak_let}{rvm}")
    al_cell.number_format = '0.00%'
    for c in range(1, AL_COL + 1):
        rcell = ws.cell(row=rvm, column=c)
        rcell.fill = orange_fill
        rcell.font = FONT_BOLD
    row += 1

    # diff row
    cdiff = ws.cell(row=row, column=AJ_COL,
                    value=f"={aj_let}{rvm}-{ai_let}{rvm}")
    cdiff.number_format = NUM_FMT
    row += 1

    # blank separator row
    row += 1

# Grand Total row
gt_row = row
ws.cell(row=gt_row, column=AI_COL, value='Grand Total')
gc_aj = ws.cell(row=gt_row, column=AJ_COL,
                value=f"=SUM(CAL!AF2:AF{cal_last})")
gc_aj.number_format = NUM_FMT
gc_ak = ws.cell(row=gt_row, column=AK_COL,
                value=f"=SUM(CAL!O2:O{cal_last})")
gc_ak.number_format = NUM_FMT
gc_al = ws.cell(row=gt_row, column=AL_COL,
                value=f"={aj_let}{gt_row}/{ak_let}{gt_row}")
gc_al.number_format = '0.00%'

print(f'sheet "{sn}": {len(branches)} branches, '
      f'{sum(len(v) for v in branch_emps.values())} employee rows, '
      f'data rows {FIRST_DATA_ROW}-{gt_row}, '
      f'summary cols {ai_let}(คอมรายคน)..{al_let}(%Com)')

wb.save(OUT_PATH)
print('saved', OUT_PATH)
