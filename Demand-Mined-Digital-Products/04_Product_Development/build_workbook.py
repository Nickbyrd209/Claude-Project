"""Builds the pilot product: Home Daycare Tax-Ready Records Workbook (tax year 2026).

Run:  python3 build_workbook.py [output_path]
All calculations are live Excel formulas (Excel + Google Sheets compatible functions only).
Every hard-coded rate lives on the Setup sheet with its source noted beside it.
"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OUT = sys.argv[1] if len(sys.argv) > 1 else "Home_Daycare_Tax_Ready_Workbook_2026.xlsx"

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True, color="1F3864")
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_KEY = PatternFill("solid", fgColor="E2EFDA")
FILL_EX = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
USD = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.00%;(0.00%);"-"'
NUM = '#,##0;(#,##0);"-"'
DATE_FMT = "mm/dd/yyyy"

HDR = 4          # header row on log sheets
FIRST = 5        # first data row
N_INCOME = 300
N_EXP = 500
N_MILES = 300
N_FAM = 30
N_WEEKS = 53

INC_LAST = FIRST + N_INCOME - 1
EXP_LAST = FIRST + N_EXP - 1
MIL_LAST = FIRST + N_MILES - 1
FAM_LAST = FIRST + N_FAM - 1
WK_LAST = FIRST + N_WEEKS - 1

INCOME_TYPES = ["Parent Payment", "Subsidy / Voucher", "CACFP Reimbursement", "Other Business Income"]
USE_TYPES = ["Business only (100%)", "Shared (Time-Space %)", "Personal (not deductible)"]
BUSINESS_CATS = [
    "Advertising", "Car & Truck (actual, if not using mileage)", "Contract Labor / Substitutes",
    "Educational & Activity Supplies", "Toys & Play Equipment", "Furniture & Appliances (under $2,500)",
    "Food", "Household & Cleaning Supplies", "Insurance (liability, not homeowners)",
    "Licenses, Fees & Dues", "Office Supplies & Postage", "Professional Fees (tax prep, legal)",
    "Repairs (daycare-only items)", "Training & Conferences", "Wages Paid to Assistants",
    "Phone & Internet", "Other Business Expense",
]
HOME_CATS = [
    "Home: Mortgage Interest", "Home: Real Estate Taxes", "Home: Homeowners/Renters Insurance",
    "Home: Utilities (gas, electric, water, trash)", "Home: Rent", "Home: Repairs & Maintenance",
]
ALL_CATS = BUSINESS_CATS + HOME_CATS

wb = Workbook()


def style_range(ws, ref, font=None, fill=None, fmt=None, border=True, align=None):
    for row in ws[ref]:
        for c in row:
            if font: c.font = font
            if fill: c.fill = fill
            if fmt: c.number_format = fmt
            if border: c.border = BOX
            if align: c.alignment = align


def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = F_TITLE
    ws["A2"] = sub; ws["A2"].font = F_NOTE
    ws.sheet_view.showGridLines = False


def headers(ws, cols, widths):
    for i, (h, w) in enumerate(zip(cols, widths), start=1):
        c = ws.cell(row=HDR, column=i, value=h)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[HDR].height = 30
    ws.freeze_panes = ws.cell(row=FIRST, column=1)


def list_dv(ws, items_ref, cell_range):
    dv = DataValidation(type="list", formula1=items_ref, allow_blank=True, showErrorMessage=True,
                        errorTitle="Pick from the list", error="Please choose a value from the dropdown list.")
    ws.add_data_validation(dv); dv.add(cell_range)


# ---------------------------------------------------------------- Start Here
ws = wb.active; ws.title = "Start Here"
title(ws, "Home Daycare Tax-Ready Records Workbook — 2026",
      "For licensed / registered family child care providers who run child care in their own home (U.S.).")
ws.column_dimensions["A"].width = 4; ws.column_dimensions["B"].width = 110
lines = [
    ("WHAT THIS DOES", True),
    ("Turns the records you already keep (hours, meals, receipts, miles, payments) into the numbers your tax preparer asks for:", False),
    ("  • Your Time-Space Percentage (the % of shared home costs you may be able to deduct) — calculated automatically.", False),
    ("  • Standard meal allowance vs. actual food cost — side by side, so you can see which is larger.", False),
    ("  • A Schedule C-style expense summary and a Form 8829-style home expense summary.", False),
    ("  • A rough quarterly set-aside estimate so tax time is not a surprise.", False),
    ("  • Year-end payment statements for each family (parents need these for the child care credit).", False),
    ("", False),
    ("HOW TO USE IT (about 15 minutes a week)", True),
    ("  1. Setup tab: enter your name, address, home square footage and the square footage used for child care. Check the rates.", False),
    ("  2. Families tab: list each family once (name, parent, child, address).", False),
    ("  3. Each week: enter hours on Time Log, payments on Income, receipts on Expenses, trips on Mileage.", False),
    ("  4. Each month: enter meal and snack counts on Meals.", False),
    ("  5. Any time: open Tax Summary to see where you stand. In January: print Parent Statements and hand Tax Summary to your preparer.", False),
    ("", False),
    ("COLOR LEGEND", True),
    ("  Yellow cells with blue text = you type here.   White/green cells = calculated, do not type over them.", False),
    ("  Peach rows labelled EXAMPLE = sample entries showing the format. Delete or overwrite them before you start.", False),
    ("", False),
    ("IMPORTANT", True),
    ("This workbook is a record-keeping and organizing tool. It is not tax, legal or accounting advice and does not file anything.", False),
    ("Tax rules have conditions and limits (for example, the home expense limit on Form 8829, depreciation, and which meals may be counted).", False),
    ("Review results with a qualified tax professional and see IRS Publication 587 (Business Use of Your Home) and the Form 8829 instructions.", False),
    ("Rates on the Setup tab were gathered from public sources on 2026-11-08 and should be re-checked each year.", False),
]
for i, (t, bold) in enumerate(lines, start=4):
    c = ws.cell(row=i, column=2, value=t); c.font = F_BOLD if bold else F_BASE

# ---------------------------------------------------------------- Setup
ws = wb.create_sheet("Setup")
title(ws, "Setup", "Enter once per year. Yellow = your input. Rates come pre-filled with their source; verify before filing.")
ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 26; ws.column_dimensions["C"].width = 90
setup_rows = [
    # (label, value, fmt, note, is_input)
    ("Provider name", "Your Name", None, "Appears on parent statements.", True),
    ("Business name", "Little Acorns Family Child Care", None, "", True),
    ("Business address", "123 Example St, Springfield, ST 00000", None, "Appears on parent statements.", True),
    ("Provider EIN or SSN (last 4 only is fine here)", "XX-XXX1234", None, "Parents need your ID number for Form 2441; you may prefer to write it in by hand on printed statements.", True),
    ("Phone / email", "(555) 555-0100", None, "", True),
    ("Tax year", 2026, "0", "Used for week dates and mileage rate switch.", True),
    ("", None, None, "", False),
    ("SPACE", None, None, "", False),
    ("Total square feet of the home", 2000, NUM, "Include all finished rooms and areas (count the basement/garage only if regularly used for child care).", True),
    ("Square feet regularly used for child care", 1500, NUM, "Areas used regularly for the business (play rooms, kitchen, bathrooms, nap areas). See IRS Pub. 587.", True),
    ("Space percentage", "=IF(B13>0,MIN(1,B14/B13),0)", PCT, "Calculated.", False),
    ("", None, None, "", False),
    ("TIME", None, None, "", False),
    ("Days the home was available for child care this year", 365, NUM, "365 for a full year. Use fewer if you started or stopped mid-year (Form 8829 instructions).", True),
    ("Total hours available (days × 24)", "=B18*24", NUM, "Calculated.", False),
    ("Business hours this year (from Time Log)", "='Time Log'!E3", NUM, "Calculated from the Time Log tab.", False),
    ("Time percentage", "=IF(B19>0,MIN(1,B20/B19),0)", PCT, "Calculated.", False),
    ("", None, None, "", False),
    ("TIME-SPACE PERCENTAGE", "=B15*B21", PCT, "Time % × Space %. Applied to shared expenses and home expenses.", False),
    ("", None, None, "", False),
    ("FOOD METHOD", None, None, "", False),
    ("Food deduction method", "Standard Meal Allowance", None, "Choose 'Standard Meal Allowance' or 'Actual Food Cost'. Tax Summary shows both so you can compare.", True),
    ("Standard rate — breakfast", 1.70, USD, "CACFP Tier I day care home rate (contiguous states) for July 2025–June 2026, the rate in effect Jan 1, 2026. Source: NJ Dept. of Agriculture CACFP memo FY2025-26. Alaska/Hawaii differ. Verify with IRS Pub. 587 / your sponsor.", True),
    ("Standard rate — lunch or supper", 3.22, USD, "Same source as above.", True),
    ("Standard rate — snack", 0.96, USD, "Same source as above.", True),
    ("", None, None, "", False),
    ("MILEAGE", None, None, "", False),
    ("IRS business mileage rate — Jan 1 to Jun 30", 0.725, '$0.000', "IRS IR-2025-128 (72.5¢/mile for 2026). irs.gov", True),
    ("IRS business mileage rate — Jul 1 to Dec 31", 0.76, '$0.000', "Mid-year increase to 76¢ reported by CPA Practice Advisor (2026-07-16) citing IRS Announcement 2026-11. Verify on irs.gov.", True),
    ("Date the second rate starts", "=DATE(B10,7,1)", DATE_FMT, "Calculated from tax year.", False),
    ("", None, None, "", False),
    ("SET-ASIDE ESTIMATE (rough guide only)", None, None, "", False),
    ("Self-employment tax rate", 0.153, PCT, "12.4% Social Security + 2.9% Medicare, applied to 92.35% of net profit (Schedule SE).", True),
    ("Social Security wage base", 184500, USD, "2026 wage base (SSA). Earnings above this are not subject to the 12.4% part.", True),
    ("Income tax set-aside % (your choice)", 0.10, PCT, "A planning percentage you choose with your preparer. Not a calculation of your actual income tax.", True),
]
r = 5
for label, val, fmt, note, is_input in setup_rows:
    a = ws.cell(row=r, column=1, value=label or None)
    if label and label.isupper():
        a.font = F_BOLD
    else:
        a.font = F_BASE
    if val is not None:
        b = ws.cell(row=r, column=2, value=val)
        b.border = BOX
        if is_input:
            b.font = F_INPUT; b.fill = FILL_INPUT
        else:
            b.font = F_BOLD if label.startswith("TIME-SPACE") else (F_LINK if isinstance(val, str) and "!" in val else F_BASE)
            b.fill = FILL_KEY if label.startswith("TIME-SPACE") else PatternFill()
        if fmt: b.number_format = fmt
    if note:
        n = ws.cell(row=r, column=3, value=note); n.font = F_NOTE; n.alignment = WRAP
    r += 1

# Fixed addresses used elsewhere (assert layout)
S = {
    "provider": "Setup!$B$5", "business": "Setup!$B$6", "address": "Setup!$B$7", "id": "Setup!$B$8",
    "phone": "Setup!$B$9", "year": "Setup!$B$10", "space": "Setup!$B$15", "time": "Setup!$B$21",
    "ts": "Setup!$B$23", "food": "Setup!$B$26", "r_bfast": "Setup!$B$27", "r_lunch": "Setup!$B$28",
    "r_snack": "Setup!$B$29", "m1": "Setup!$B$32", "m2": "Setup!$B$33", "mswitch": "Setup!$B$34",
    "se_rate": "Setup!$B$37", "wage_base": "Setup!$B$38", "inc_pct": "Setup!$B$39",
}
assert ws["A23"].value == "TIME-SPACE PERCENTAGE" and ws["A26"].value == "Food deduction method"
assert ws["A39"].value.startswith("Income tax set-aside")
list_dv(ws, '"Standard Meal Allowance,Actual Food Cost"', "B26")

# Hidden lists sheet
wl = wb.create_sheet("Lists")
for i, v in enumerate(INCOME_TYPES, 1): wl.cell(row=i, column=1, value=v)
for i, v in enumerate(USE_TYPES, 1): wl.cell(row=i, column=2, value=v)
for i, v in enumerate(ALL_CATS, 1): wl.cell(row=i, column=3, value=v)
wl.sheet_state = "hidden"
L_INC = f"Lists!$A$1:$A${len(INCOME_TYPES)}"
L_USE = f"Lists!$B$1:$B${len(USE_TYPES)}"
L_CAT = f"Lists!$C$1:$C${len(ALL_CATS)}"

# ---------------------------------------------------------------- Families
ws = wb.create_sheet("Families")
title(ws, "Families", "List each family once. Use the exact same Family name on the Income tab. Totals fill in automatically.")
headers(ws, ["Family name (as used on Income)", "Parent / guardian name(s)", "Child name(s)", "Parent mailing address",
             "Paid by parent this year", "Subsidy / voucher received for this family"], [30, 30, 28, 40, 20, 24])
for rr in range(FIRST, FAM_LAST + 1):
    style_range(ws, f"A{rr}:D{rr}", font=F_INPUT, fill=FILL_INPUT)
    ws[f"E{rr}"] = f'=IF(A{rr}="","",SUMIFS(Income!$D${FIRST}:$D${INC_LAST},Income!$B${FIRST}:$B${INC_LAST},A{rr},Income!$C${FIRST}:$C${INC_LAST},"Parent Payment"))'
    ws[f"F{rr}"] = f'=IF(A{rr}="","",SUMIFS(Income!$D${FIRST}:$D${INC_LAST},Income!$B${FIRST}:$B${INC_LAST},A{rr},Income!$C${FIRST}:$C${INC_LAST},"Subsidy / Voucher"))'
    style_range(ws, f"E{rr}:F{rr}", font=F_BASE, fmt=USD)
for col, v in zip("ABCD", ["EXAMPLE Rivera", "Ana Rivera", "Leo Rivera", "45 Sample Ave, Springfield, ST 00000"]):
    ws[f"{col}{FIRST}"] = v; ws[f"{col}{FIRST}"].fill = FILL_EX

# ---------------------------------------------------------------- Time Log
ws = wb.create_sheet("Time Log")
title(ws, "Time Log", "Each week: hours children were in your care + hours you worked on the business in your home while they were not there.")
headers(ws, ["Week #", "Week starting", "Hours children present", "Hours of prep, cleaning, meal prep, paperwork, planning (in home, no children present)", "Total business hours", "Notes"],
        [8, 14, 16, 34, 16, 40])
ws["D3"] = "Total business hours this year:"; ws["D3"].font = F_BOLD; ws["D3"].alignment = Alignment(horizontal="right")
ws["E3"] = f"=SUM(E{FIRST}:E{WK_LAST})"; ws["E3"].font = F_BOLD; ws["E3"].number_format = NUM; ws["E3"].fill = FILL_KEY; ws["E3"].border = BOX
for i in range(N_WEEKS):
    rr = FIRST + i
    ws[f"A{rr}"] = i + 1
    ws[f"B{rr}"] = f"=DATE({S['year']},1,1)+7*(A{rr}-1)"
    ws[f"E{rr}"] = f"=IF(AND(C{rr}=\"\",D{rr}=\"\"),\"\",N(C{rr})+N(D{rr}))"
    style_range(ws, f"A{rr}:B{rr}", font=F_BASE); ws[f"B{rr}"].number_format = DATE_FMT
    style_range(ws, f"C{rr}:D{rr}", font=F_INPUT, fill=FILL_INPUT, fmt="0.0")
    style_range(ws, f"E{rr}:E{rr}", font=F_BASE, fmt="0.0")
    style_range(ws, f"F{rr}:F{rr}", font=F_INPUT, fill=FILL_INPUT)
ws[f"C{FIRST}"] = 55; ws[f"D{FIRST}"] = 12; ws[f"F{FIRST}"] = "EXAMPLE: 7am–6pm M–F; Sun meal prep + Mon–Fri cleanup"
for col in "CDF": ws[f"{col}{FIRST}"].fill = FILL_EX
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="168", allow_blank=True,
                    showErrorMessage=True, errorTitle="Check hours", error="Weekly hours must be between 0 and 168.")
ws.add_data_validation(dv); dv.add(f"C{FIRST}:D{WK_LAST}")
ws.conditional_formatting.add(f"A{FIRST}:F{WK_LAST}", FormulaRule(formula=[f"N($E{FIRST})>168"], fill=PatternFill("solid", fgColor="F8CBAD")))

# ---------------------------------------------------------------- Income
ws = wb.create_sheet("Income")
title(ws, "Income", "Every payment received. Type 'Parent Payment' for money parents paid you — that is what appears on their year-end statement.")
headers(ws, ["Date received", "Family name (match Families tab)", "Type", "Amount", "Method / check #", "Notes"], [14, 30, 24, 14, 18, 40])
for rr in range(FIRST, INC_LAST + 1):
    style_range(ws, f"A{rr}:F{rr}", font=F_INPUT, fill=FILL_INPUT)
    ws[f"A{rr}"].number_format = DATE_FMT; ws[f"D{rr}"].number_format = USD
list_dv(ws, f"={L_INC}", f"C{FIRST}:C{INC_LAST}")
list_dv(ws, f"=Families!$A${FIRST}:$A${FAM_LAST}", f"B{FIRST}:B{INC_LAST}")
for col, v in zip("ABCDEF", [date(2026, 1, 9), "EXAMPLE Rivera", "Parent Payment", 250, "Check 1042", "Week of Jan 5"]):
    ws[f"{col}{FIRST}"] = v; ws[f"{col}{FIRST}"].fill = FILL_EX

# ---------------------------------------------------------------- Expenses
ws = wb.create_sheet("Expenses")
title(ws, "Expenses", "Every receipt. Pick a category and how it is used. 'Home:' categories are always multiplied by your Time-Space %.")
headers(ws, ["Date", "Store / vendor", "What was it for?", "Category", "How is it used?", "Amount paid", "Counts toward deduction", "Notes"],
        [12, 22, 30, 36, 24, 13, 15, 30])
ts = S["ts"]
for rr in range(FIRST, EXP_LAST + 1):
    style_range(ws, f"A{rr}:F{rr}", font=F_INPUT, fill=FILL_INPUT)
    style_range(ws, f"H{rr}:H{rr}", font=F_INPUT, fill=FILL_INPUT)
    ws[f"A{rr}"].number_format = DATE_FMT; ws[f"F{rr}"].number_format = USD
    ws[f"G{rr}"] = (f'=IF(F{rr}="","",IF(LEFT(D{rr},5)="Home:",F{rr}*{ts},'
                    f'IF(AND(D{rr}="Food",{S["food"]}="Standard Meal Allowance"),0,'
                    f'IF(E{rr}="Business only (100%)",F{rr},IF(E{rr}="Shared (Time-Space %)",F{rr}*{ts},0)))))')
    style_range(ws, f"G{rr}:G{rr}", font=F_BASE, fmt=USD)
list_dv(ws, f"={L_CAT}", f"D{FIRST}:D{EXP_LAST}")
list_dv(ws, f"={L_USE}", f"E{FIRST}:E{EXP_LAST}")
ws.conditional_formatting.add(f"A{FIRST}:H{EXP_LAST}", FormulaRule(
    formula=[f'AND($F{FIRST}<>"",OR($D{FIRST}="",AND(LEFT($D{FIRST},5)<>"Home:",$E{FIRST}="")))'],
    fill=PatternFill("solid", fgColor="F8CBAD")))
ws["H3"] = "Red row = amount entered but category or use missing."; ws["H3"].font = F_NOTE
for col, v in zip("ABCDEFH", [date(2026, 1, 12), "Target", "Art supplies, paper plates", "Educational & Activity Supplies",
                              "Business only (100%)", 42.18, "EXAMPLE"]):
    ws[f"{col}{FIRST}"] = v; ws[f"{col}{FIRST}"].fill = FILL_EX

# ---------------------------------------------------------------- Meals
ws = wb.create_sheet("Meals")
title(ws, "Meals & Snacks", "Monthly counts of meals served to eligible children in your care. Max per child per day: 1 breakfast, 1 lunch, 1 supper, 3 snacks (Pub. 587).")
headers(ws, ["Month", "Breakfasts", "Lunches", "Suppers", "Snacks", "Standard allowance for month", "Notes"], [14, 12, 12, 12, 12, 22, 44])
months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
for i, m in enumerate(months):
    rr = FIRST + i
    ws[f"A{rr}"] = m; ws[f"A{rr}"].font = F_BASE; ws[f"A{rr}"].border = BOX
    style_range(ws, f"B{rr}:E{rr}", font=F_INPUT, fill=FILL_INPUT, fmt=NUM)
    ws[f"F{rr}"] = f"=N(B{rr})*{S['r_bfast']}+(N(C{rr})+N(D{rr}))*{S['r_lunch']}+N(E{rr})*{S['r_snack']}"
    style_range(ws, f"F{rr}:F{rr}", font=F_BASE, fmt=USD)
    style_range(ws, f"G{rr}:G{rr}", font=F_INPUT, fill=FILL_INPUT)
tot = FIRST + 12
ws[f"A{tot}"] = "TOTAL"; ws[f"A{tot}"].font = F_BOLD
for col in "BCDEF":
    ws[f"{col}{tot}"] = f"=SUM({col}{FIRST}:{col}{tot-1})"; ws[f"{col}{tot}"].font = F_BOLD; ws[f"{col}{tot}"].border = BOX
    ws[f"{col}{tot}"].number_format = USD if col == "F" else NUM
ws[f"F{tot}"].fill = FILL_KEY
ws[f"B{FIRST}"], ws[f"C{FIRST}"], ws[f"D{FIRST}"], ws[f"E{FIRST}"] = 84, 105, 0, 190
ws[f"G{FIRST}"] = "EXAMPLE counts — replace with your own"
for col in "BCDEG": ws[f"{col}{FIRST}"].fill = FILL_EX
MEALS_TOTAL = f"Meals!$F${tot}"

# ---------------------------------------------------------------- Mileage
ws = wb.create_sheet("Mileage")
title(ws, "Mileage", "Business trips only (grocery runs for daycare food, supply shopping, field trips, trainings). Rate switches automatically on the date set in Setup.")
headers(ws, ["Date", "Where to / purpose", "Round-trip miles", "Rate used", "Deduction", "Notes"], [12, 40, 14, 11, 14, 36])
for rr in range(FIRST, MIL_LAST + 1):
    style_range(ws, f"A{rr}:C{rr}", font=F_INPUT, fill=FILL_INPUT)
    style_range(ws, f"F{rr}:F{rr}", font=F_INPUT, fill=FILL_INPUT)
    ws[f"A{rr}"].number_format = DATE_FMT; ws[f"C{rr}"].number_format = "0.0"
    ws[f"D{rr}"] = f'=IF(OR(A{rr}="",C{rr}=""),"",IF(A{rr}>={S["mswitch"]},{S["m2"]},{S["m1"]}))'
    ws[f"E{rr}"] = f'=IF(D{rr}="","",C{rr}*D{rr})'
    style_range(ws, f"D{rr}:D{rr}", font=F_BASE, fmt="$0.000")
    style_range(ws, f"E{rr}:E{rr}", font=F_BASE, fmt=USD)
ws["D3"] = "Total miles / deduction:"; ws["D3"].font = F_BOLD
ws["C3"] = f"=SUM(C{FIRST}:C{MIL_LAST})"; ws["C3"].number_format = "#,##0.0"; ws["C3"].font = F_BOLD
ws["E3"] = f"=SUM(E{FIRST}:E{MIL_LAST})"; ws["E3"].number_format = USD; ws["E3"].font = F_BOLD; ws["E3"].fill = FILL_KEY
for col, v in zip("ABCF", [date(2026, 1, 10), "Costco — daycare groceries & supplies", 14.2, "EXAMPLE"]):
    ws[f"{col}{FIRST}"] = v; ws[f"{col}{FIRST}"].fill = FILL_EX
MILES_TOTAL = "Mileage!$E$3"

# ---------------------------------------------------------------- Tax Summary
ws = wb.create_sheet("Tax Summary")
title(ws, "Tax Summary (estimate)", "Live summary for you and your tax preparer. Figures are estimates from your entries — not a tax return.")
ws.column_dimensions["A"].width = 52; ws.column_dimensions["B"].width = 18; ws.column_dimensions["C"].width = 70
EXP_CAT = f"Expenses!$D${FIRST}:$D${EXP_LAST}"
EXP_DED = f"Expenses!$G${FIRST}:$G${EXP_LAST}"
EXP_AMT = f"Expenses!$F${FIRST}:$F${EXP_LAST}"
INC_TYPE = f"Income!$C${FIRST}:$C${INC_LAST}"
INC_AMT = f"Income!$D${FIRST}:$D${INC_LAST}"
rows = []  # (label, formula, fmt, note, style)
rows.append(("YOUR PERCENTAGES", None, None, "", "h"))
rows.append(("Space percentage", f"={S['space']}", PCT, "", ""))
rows.append(("Time percentage", f"={S['time']}", PCT, "", ""))
rows.append(("Time-Space percentage", f"={S['ts']}", PCT, "Give this to your preparer (Form 8829 / Pub. 587 daycare worksheet).", "k"))
rows.append(("", None, None, "", ""))
rows.append(("INCOME", None, None, "", "h"))
for t in INCOME_TYPES:
    rows.append((t, f'=SUMIFS({INC_AMT},{INC_TYPE},"{t}")', USD, "CACFP reimbursements are generally reported as income." if t.startswith("CACFP") else "", ""))
rows.append(("Total business income", "=SUM(B{a}:B{b})", USD, "", "k"))
inc_total_idx = len(rows) - 1
rows.append(("", None, None, "", ""))
rows.append(("BUSINESS EXPENSES (Schedule C-style)", None, None, "Shared items already multiplied by Time-Space %.", "h"))
cat_start = len(rows)
for c in BUSINESS_CATS:
    rows.append((c, f'=SUMIFS({EXP_DED},{EXP_CAT},"{c}")', USD, "Food receipts count as $0 while 'Standard Meal Allowance' is chosen on Setup." if c == "Food" else "", ""))
cat_end = len(rows) - 1
rows.append(("Food — standard meal allowance (from Meals tab)", f'=IF({S["food"]}="Standard Meal Allowance",{MEALS_TOTAL},0)', USD, "", ""))
rows.append(("Business mileage (from Mileage tab)", f"={MILES_TOTAL}", USD, "Don't also claim 'Car & Truck (actual)' for the same vehicle.", ""))
rows.append(("Total business expenses before home costs", "=SUM(B{c1}:B{c2})", USD, "", "k"))
exp_total_idx = len(rows) - 1
rows.append(("Tentative profit before home costs", "=B{inc}-B{exp}", USD, "", "k"))
tent_idx = len(rows) - 1
rows.append(("", None, None, "", ""))
rows.append(("HOME EXPENSES (Form 8829-style)", None, None, "Amounts below are your totals × Time-Space %.", "h"))
home_start = len(rows)
for c in HOME_CATS:
    rows.append((c.replace("Home: ", ""), f'=SUMIFS({EXP_DED},{EXP_CAT},"{c}")', USD, "", ""))
rows.append(("Home depreciation (from your preparer)", 0, USD, "Optional input. Depreciation has its own rules — ask your preparer.", "i"))
home_end = len(rows) - 1
rows.append(("Total home expenses (business share)", "=SUM(B{h1}:B{h2})", USD, "", "k"))
home_total_idx = len(rows) - 1
rows.append(("Check: home expenses vs. tentative profit", '=IF(B{home}>MAX(0,B{tent}),"Over limit — preparer must apply the Form 8829 limit","OK")', None,
             "Most home expenses can't create a loss; excess may carry forward. Mortgage interest & taxes have different treatment.", ""))
rows.append(("", None, None, "", ""))
rows.append(("ESTIMATED NET PROFIT", "=B{tent}-B{home}", USD, "Before the Form 8829 limit and any items not entered here.", "k"))
net_idx = len(rows) - 1
rows.append(("", None, None, "", ""))
rows.append(("FOOD METHOD COMPARISON", None, None, "", "h"))
rows.append(("Actual food receipts (business only + shared share)",
             f'=SUMPRODUCT(({EXP_CAT}="Food")*({EXP_AMT}<>"")*(N(+{EXP_AMT})))', USD, "placeholder", ""))
food_actual_idx = len(rows) - 1
rows.append(("Standard meal allowance (Meals tab)", f"={MEALS_TOTAL}", USD, "", ""))
rows.append(("Larger of the two", '=IF(B{fa}>=B{fs},"Actual Food Cost","Standard Meal Allowance")', None, "Higher is not automatically better: actual cost needs receipts and must exclude family food.", ""))
food_std_idx = len(rows) - 2
rows.append(("", None, None, "", ""))
rows.append(("ROUGH SET-ASIDE GUIDE", None, None, "Planning aid only — your actual tax depends on your whole return.", "h"))
rows.append(("Self-employment tax estimate", f"=IF(B{{net}}<=0,0,MIN(B{{net}}*0.9235,{S['wage_base']})*0.124+B{{net}}*0.9235*0.029)", USD,
             "Approximates Schedule SE: 92.35% of profit; 12.4% up to the wage base + 2.9% Medicare. Ignores wages from other jobs.", ""))
se_idx = len(rows) - 1
rows.append(("Income tax set-aside (your % on Setup)", f"=MAX(0,B{{net}})*{S['inc_pct']}", USD, "", ""))
rows.append(("Total suggested set-aside for the year", "=B{se}+B{it}", USD, "", "k"))
sa_idx = len(rows) - 1
rows.append(("Per quarter (÷ 4)", "=B{sa}/4", USD, "2026 estimated payment due dates: Apr 15, Jun 15, Sep 15 2026 and Jan 15 2027 (IRS Form 1040-ES).", "k"))

start_row = 5
rown = {i: start_row + i for i in range(len(rows))}
# Fix the actual-food formula: simple SUMIFS on amount is correct & portable
rows[food_actual_idx] = ("Actual food receipts (business only + shared share)",
                         f'=SUMIFS({EXP_AMT},{EXP_CAT},"Food",Expenses!$E${FIRST}:$E${EXP_LAST},"Business only (100%)")'
                         f'+SUMIFS({EXP_AMT},{EXP_CAT},"Food",Expenses!$E${FIRST}:$E${EXP_LAST},"Shared (Time-Space %)")*{S["ts"]}',
                         USD, "What food receipts would count if you chose 'Actual Food Cost'.", "")
inc_rows = [rown[i] for i in range(len(rows)) if rows[i][0] in INCOME_TYPES]
fmt_map = {
    "a": inc_rows[0], "b": inc_rows[-1], "inc": rown[inc_total_idx], "c1": rown[cat_start], "c2": rown[exp_total_idx - 1],
    "exp": rown[exp_total_idx], "tent": rown[tent_idx], "h1": rown[home_start], "h2": rown[home_end], "home": rown[home_total_idx],
    "net": rown[net_idx], "fa": rown[food_actual_idx], "fs": rown[food_actual_idx + 1], "se": rown[se_idx], "it": rown[se_idx + 1], "sa": rown[sa_idx],
}
for i, (label, f, fmt, note, st) in enumerate(rows):
    rr = rown[i]
    a = ws.cell(row=rr, column=1, value=label or None); a.font = F_BOLD if st in ("h", "k") else F_BASE
    if f is not None:
        val = f.format(**fmt_map) if isinstance(f, str) else f
        b = ws.cell(row=rr, column=2, value=val); b.border = BOX
        if st == "i":
            b.font = F_INPUT; b.fill = FILL_INPUT
        else:
            b.font = F_BOLD if st == "k" else F_BASE
            if st == "k": b.fill = FILL_KEY
        if fmt: b.number_format = fmt
    if note:
        n = ws.cell(row=rr, column=3, value=note); n.font = F_NOTE; n.alignment = WRAP
T = {k: f"'Tax Summary'!$B${v}" for k, v in fmt_map.items()}
ws.print_area = f"A1:C{rown[len(rows)-1]}"
ws.page_setup.fitToWidth = 1

# ---------------------------------------------------------------- Parent Statements
ws = wb.create_sheet("Parent Statements")
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 4; ws.column_dimensions["B"].width = 36; ws.column_dimensions["C"].width = 48
ws["B2"] = "Choose a family:"; ws["B2"].font = F_BOLD
ws["C2"] = "EXAMPLE Rivera"; ws["C2"].font = F_INPUT; ws["C2"].fill = FILL_INPUT; ws["C2"].border = BOX
list_dv(ws, f"=Families!$A${FIRST}:$A${FAM_LAST}", "C2")
ws["B3"] = "Pick a family, then print this page (File > Print > Selection/Print area). Repeat for each family."; ws["B3"].font = F_NOTE
ws["B5"] = f'="Child Care Payment Statement — Tax Year "&{S["year"]}'; ws["B5"].font = Font(name=FONT, size=14, bold=True)
fam_rng = f"Families!$A${FIRST}:$A${FAM_LAST}"
def fam(col):
    return f'=IFERROR(INDEX(Families!${col}${FIRST}:${col}${FAM_LAST},MATCH($C$2,{fam_rng},0)),"")'
items = [
    (7, "Provider", f"={S['provider']}"), (8, "Business name", f"={S['business']}"), (9, "Address", f"={S['address']}"),
    (10, "Taxpayer ID (EIN/SSN)", f"={S['id']}"), (11, "Phone / email", f"={S['phone']}"),
    (13, "Parent / guardian", fam("B")), (14, "Child(ren)", fam("C")), (15, "Parent address", fam("D")),
    (17, "Total you paid for child care in the year", f'=SUMIFS({INC_AMT},Income!$B${FIRST}:$B${INC_LAST},$C$2,{INC_TYPE},"Parent Payment")'),
    (18, "Paid on your behalf by subsidy/voucher (not paid by you)", f'=SUMIFS({INC_AMT},Income!$B${FIRST}:$B${INC_LAST},$C$2,{INC_TYPE},"Subsidy / Voucher")'),
    (19, "Number of payments recorded", f'=COUNTIFS(Income!$B${FIRST}:$B${INC_LAST},$C$2,{INC_TYPE},"Parent Payment")'),
]
for rr, lab, f in items:
    ws[f"B{rr}"] = lab; ws[f"B{rr}"].font = F_BASE
    ws[f"C{rr}"] = f; ws[f"C{rr}"].font = F_BOLD if rr == 17 else F_BASE; ws[f"C{rr}"].border = BOX
    if rr in (17, 18): ws[f"C{rr}"].number_format = USD
ws["C17"].fill = FILL_KEY
ws["B21"] = "Statement prepared from the provider's payment records. Keep for your records (e.g., Form 2441)."; ws["B21"].font = F_NOTE
ws["B24"] = "Provider signature: ______________________________"; ws["C24"] = "Date: ______________"
ws["B26"] = '=IF(COUNTIF(' + fam_rng + ',$C$2)=0,"⚠ Family not found on Families tab — check spelling.","")'
ws["B26"].font = Font(name=FONT, size=10, bold=True, color="C00000")
ws.print_area = "B5:C24"

# Sheet order
order = ["Start Here", "Setup", "Families", "Time Log", "Income", "Expenses", "Meals", "Mileage", "Tax Summary", "Parent Statements", "Lists"]
wb._sheets = [wb[n] for n in order]
for s in wb.worksheets:
    for row in s.iter_rows():
        for c in row:
            if c.font is None or c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size if c.font else 10, bold=c.font.bold if c.font else False,
                              italic=c.font.italic if c.font else False, color=c.font.color if c.font else None)
wb.save(OUT)

# Write the address map so QA can find cells without guessing
import json
json.dump({"setup": S, "tax_summary": T, "first": FIRST, "rows": {k: v for k, v in fmt_map.items()}},
          open(OUT.rsplit("/", 1)[0] + "/cell_map.json" if "/" in OUT else "cell_map.json", "w"), indent=2)
print("saved", OUT)
