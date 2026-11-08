"""Builds the Master Business Dashboard workbook from the opportunity CSV.
Run: python3 build_dashboard.py   (from this folder)
Financial Performance holds ACTUALS ONLY (blank until real sales). Scenarios are ILLUSTRATIVE.
"""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

F = "Arial"
B = Font(name=F, size=10); BB = Font(name=F, size=10, bold=True); BI = Font(name=F, size=10, color="0000FF")
H = Font(name=F, size=10, bold=True, color="FFFFFF"); T = Font(name=F, size=14, bold=True, color="1F3864")
N = Font(name=F, size=9, italic=True, color="595959")
HF = PatternFill("solid", fgColor="1F3864"); IF_ = PatternFill("solid", fgColor="FFF2CC"); KF = PatternFill("solid", fgColor="E2EFDA")
WARN = PatternFill("solid", fgColor="FCE4D6")
S = Side(style="thin", color="BFBFBF"); BOX = Border(left=S, right=S, top=S, bottom=S)
USD = '$#,##0.00;($#,##0.00);"-"'; PCT = '0.0%;(0.0%);"-"'
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()

def hdr(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical="center"); x.border = BOX
        if widths: ws.column_dimensions[x.column_letter].width = widths[i - 1]

# Market Intelligence
ws = wb.active; ws.title = "Market Intelligence"
ws["A1"] = "Market Intelligence — 15 niches (scores are analyst judgments; see 03_Opportunity_Scoring)"; ws["A1"].font = T
cols = ["Rank", "Niche", "Customer problem", "Demand /25", "Severity /20", "Gap /20", "Ease /15", "Acquisition /10", "Margin /10",
        "Total /100", "Confidence", "Evidence of demand", "Competitor examples", "Unmet need", "Competition", "Sources"]
hdr(ws, 3, cols, [6, 34, 40, 9, 9, 8, 8, 10, 9, 9, 14, 44, 44, 36, 13, 40])
rows = list(csv.DictReader(open("../03_Opportunity_Scoring/opportunity_scores.csv")))
for i, r in enumerate(rows):
    rr = 4 + i
    vals = [int(r["rank"]), r["niche"], r["customer_problem"], int(r["demand_25"]), int(r["severity_20"]), int(r["gap_20"]),
            int(r["ease_15"]), int(r["acquisition_10"]), int(r["margin_10"]), f"=SUM(D{rr}:I{rr})", r["confidence"],
            r["evidence_type"], r["competitor_examples"], r["unmet_need"], r["competition"], r["key_sources"]]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=rr, column=j, value=v); c.font = BB if j == 10 else B; c.border = BOX; c.alignment = WRAP
    ws.cell(row=rr, column=10).fill = KF
ws.freeze_panes = "C4"

# Product Pipeline
ws = wb.create_sheet("Product Pipeline")
ws["A1"] = "Product Pipeline"; ws["A1"].font = T
ws["A2"] = "Statuses: Researching · Opportunity Identified · Validating · Prototype Development · Customer Testing · Ready for Approval · Launching · Active · Optimizing · Paused · Rejected"; ws["A2"].font = N
cols = ["ID", "Product / opportunity", "Target customer", "Format", "Planned price", "Status", "Development", "QA status",
        "Validation status", "Decision", "Recommended next action", "Owner approval needed?"]
hdr(ws, 4, cols, [7, 34, 30, 18, 11, 20, 22, 30, 24, 14, 50, 26])
pipe = [
    ["P-001", "Home Daycare Tax-Ready Records Workbook 2026", "U.S. family child care providers", "Excel / Google Sheets", 29,
     "Prototype Development", "MVP built (v0.1)", "194/194 automated checks; Excel/Sheets/manual review pending", "Not started",
     "CONDITIONAL", "Owner opens in Excel + Google Sheets; approve CPA/EA review and validation outreach", "Yes — spend, outreach, accounts"],
    ["P-002", "Free Time-Space % Calculator (lead magnet)", "Same", "Excel / Google Sheets", 0, "Opportunity Identified",
     "Not started (derive from P-001)", "—", "—", "GO (internal)", "Build after P-001 Sheets check", "Yes — publishing"],
    ["O-003", "Cleaning business quote & job-profit calculator", "Solo/small residential cleaners", "Spreadsheet", 24,
     "Opportunity Identified", "Not started", "—", "Not started", "CONDITIONAL", "3 organic test videos in month 2 if capacity allows", "Yes — posting"],
    ["O-004", "Freelancer quarterly tax tracker", "1099 freelancers", "Spreadsheet", None, "Rejected", "—", "—", "—", "NO-GO",
     "Revisit only for a niche segment with special rules", "No"],
    ["O-005", "PTA/PTO treasurer handover kit", "Volunteer treasurers", "Spreadsheet + guide", None, "Paused", "—", "—", "—",
     "Watch", "Re-evaluate in June (Aug–Sep treasurer turnover)", "No"],
]
sv = DataValidation(type="list", formula1='"Researching,Opportunity Identified,Validating,Prototype Development,Customer Testing,Ready for Approval,Launching,Active,Optimizing,Paused,Rejected"')
ws.add_data_validation(sv); sv.add("F5:F60")
for i, p in enumerate(pipe):
    for j, v in enumerate(p, 1):
        c = ws.cell(row=5 + i, column=j, value=v); c.font = B; c.border = BOX; c.alignment = WRAP
        if j == 5 and v is not None: c.number_format = USD

# Financial Performance (actuals only)
ws = wb.create_sheet("Financial Performance")
ws["A1"] = "Financial Performance — ACTUALS ONLY"; ws["A1"].font = T
ws["A2"] = "No sales have occurred. Enter real monthly figures from the storefront/ads exports. Blank = no data (not zero)."; ws["A2"].font = N
cols = ["Month", "Product ID", "Units sold", "Gross revenue", "Platform & transaction fees", "Marketing spend", "Refunds",
        "Software expenses", "Other direct costs", "Contribution margin", "Store visitors", "Conversion rate", "Customer acquisition cost (paid)",
        "Other operating costs (one-time / overhead)", "Net operating profit before tax"]
hdr(ws, 4, cols, [10, 10, 9, 13, 14, 12, 10, 12, 12, 14, 11, 11, 14, 16, 16])
for rr in range(5, 41):
    for col in "ABCDEFGHIKN":
        c = ws[f"{col}{rr}"]; c.font = BI; c.fill = IF_; c.border = BOX
        if col in "DEFGHIN": c.number_format = USD
    ws[f"J{rr}"] = f'=IF(D{rr}="","",D{rr}-E{rr}-F{rr}-G{rr}-H{rr}-I{rr})'
    ws[f"L{rr}"] = f'=IF(OR(K{rr}="",K{rr}=0,C{rr}=""),"",C{rr}/K{rr})'
    ws[f"M{rr}"] = f'=IF(OR(C{rr}="",C{rr}=0,F{rr}=""),"",F{rr}/C{rr})'
    ws[f"O{rr}"] = f'=IF(J{rr}="","",J{rr}-N(N{rr}))'
    for col, fmt in zip("JLMO", [USD, PCT, USD, USD]):
        ws[f"{col}{rr}"].number_format = fmt; ws[f"{col}{rr}"].font = B; ws[f"{col}{rr}"].border = BOX
ws["A42"] = "TOTAL"; ws["A42"].font = BB
for col in "CDEFGHIJKNO":
    ws[f"{col}42"] = f"=SUM({col}5:{col}41)"; ws[f"{col}42"].font = BB; ws[f"{col}42"].number_format = USD if col not in "CK" else "#,##0"
ws["L42"] = '=IF(K42=0,"",C42/K42)'; ws["L42"].number_format = PCT
ws["M42"] = '=IF(C42=0,"",F42/C42)'; ws["M42"].number_format = USD

# Scenarios (illustrative)
ws = wb.create_sheet("Scenarios (Illustrative)")
ws["A1"] = "ILLUSTRATIVE scenarios — not predictions. Change blue inputs."; ws["A1"].font = T
ws.column_dimensions["A"].width = 44
for col in "BCD": ws.column_dimensions[col].width = 14
ws.column_dimensions["E"].width = 70
assump = [("Price per unit", 29, USD, "Recommended launch price (07_Sales)"),
          ("Platform % fee", 0.05, PCT, "Payhip Free plan (payhip help center)"),
          ("Processor % fee", 0.029, PCT, "Stripe standard card rate, approx. — verify"),
          ("Processor fixed fee per sale", 0.30, USD, "Stripe, approx. — verify"),
          ("Monthly software cost", 0, USD, "Payhip Free = $0/month")]
for i, (l, v, f, n) in enumerate(assump):
    r = 3 + i
    ws[f"A{r}"] = l; ws[f"A{r}"].font = B
    ws[f"B{r}"] = v; ws[f"B{r}"].font = BI; ws[f"B{r}"].fill = IF_; ws[f"B{r}"].number_format = f; ws[f"B{r}"].border = BOX
    ws[f"E{r}"] = n; ws[f"E{r}"].font = N
hdr(ws, 9, ["Monthly line item", "Low", "Moderate", "High", "Notes"])
ws.column_dimensions["A"].width = 44
lines = [("Units sold (assumption)", [5, 25, 75], "#,##0", "Assumed volumes, not forecasts", True),
         ("Marketing spend (assumption)", [0, 50, 150], USD, "Organic-first; paid only with approval", True),
         ("Refund rate (assumption)", [0.05, 0.04, 0.03], PCT, "Assumed; fees treated as not returned", True),
         ("Other direct costs", [0, 0, 0], USD, "e.g., annual CPA review amortized — enter if incurred", True)]
r = 10
for l, vals, f, n, inp in lines:
    ws[f"A{r}"] = l; ws[f"A{r}"].font = B
    for col, v in zip("BCD", vals):
        c = ws[f"{col}{r}"]; c.value = v; c.number_format = f; c.border = BOX; c.font = BI; c.fill = IF_
    ws[f"E{r}"] = n; ws[f"E{r}"].font = N; r += 1
calc = [("Gross sales", "={c}10*$B$3", "Units × price"),
        ("Platform & transaction fees", "={c}10*($B$3*($B$4+$B$5)+$B$6)", "Per-sale % + fixed"),
        ("Refunds", "={c}14*{c}12", "Gross × refund rate"),
        ("Advertising", "={c}11", ""),
        ("Software", "=$B$7", ""),
        ("Other direct costs", "={c}13", ""),
        ("CONTRIBUTION MARGIN", "={c}14-{c}15-{c}16-{c}17-{c}18-{c}19", "Gross − fees − refunds − ads − software − other direct"),
        ("Contribution margin %", '=IF({c}14=0,"",{c}20/{c}14)', ""),
        ("Paid CAC (ads ÷ units)", '=IF({c}10=0,"",{c}11/{c}10)', "Assumes all units from ads — conservative upper bound")]
assert r == 14
for l, f, n in calc:
    ws[f"A{r}"] = l; ws[f"A{r}"].font = BB if l.isupper() else B
    for col in "BCD":
        c = ws[f"{col}{r}"]; c.value = f.format(c=col); c.border = BOX; c.font = BB if l.isupper() else B
        c.number_format = PCT if "%" in l else USD
        if l.isupper(): c.fill = KF
    ws[f"E{r}"] = n; ws[f"E{r}"].font = N; r += 1
ws[f"A{r+1}"] = "Net operating profit before tax = contribution margin − one-time/overhead costs (see Budget). Revenue is never profit."; ws[f"A{r+1}"].font = N

# Budget tracker
ws = wb.create_sheet("Startup Budget")
ws["A1"] = "Startup Budget — nothing is spent without owner approval"; ws["A1"].font = T
hdr(ws, 3, ["ID", "Item", "Required?", "Estimated cost", "Approved? (Y/N)", "Actual spent", "Notes"], [6, 46, 14, 14, 14, 13, 60])
items = [("B1", "Payhip account (Free plan)", "Required", 0, "5% per sale only; no monthly fee"),
         ("B2", "Etsy shop setup + 3 listings", "Optional", 16, "Etsy has charged new US sellers a one-time setup fee (~$15, verify) + $0.20/listing"),
         ("B3", "Customer interview thank-you gift cards (5 × $10)", "Recommended", 50, "For 20-min calls with providers"),
         ("B4", "One-hour CPA/EA review of workbook rules & wording", "Strongly recommended", 150, "Reduces tax-tool liability risk; get 2 quotes"),
         ("B5", "Paid traffic test (Meta/Pinterest) for landing page", "Optional", 0, "Only if organic validation is inconclusive; cap $34 so total stays within $250"),
         ("B6", "Domain name", "Deferred", 0, "Not needed for validation; Payhip subdomain is fine")]
for i, (a, b, c_, d, n) in enumerate(items):
    rr = 4 + i
    for col, v in zip("ABCDG", [a, b, c_, d, n]):
        x = ws[f"{col}{rr}"]; x.value = v; x.font = B; x.border = BOX; x.alignment = WRAP
    ws[f"D{rr}"].number_format = USD
    for col in "EF":
        ws[f"{col}{rr}"].fill = IF_; ws[f"{col}{rr}"].font = BI; ws[f"{col}{rr}"].border = BOX
    ws[f"F{rr}"].number_format = USD
ws["B11"] = "Recommended total (B1–B4)"; ws["B11"].font = BB
ws["D11"] = "=SUM(D4:D7)"; ws["D11"].number_format = USD; ws["D11"].font = BB; ws["D11"].fill = KF
ws["B12"] = "Maximum incl. optional paid test (capped at $250 ceiling)"; ws["B12"].font = BB
ws["D12"] = "=MIN(250,D11+34)"; ws["D12"].number_format = USD; ws["D12"].font = BB
ws["B13"] = "Actual spent to date"; ws["B13"].font = BB
ws["F13"] = "=SUM(F4:F9)"; ws["F13"].number_format = USD; ws["F13"].font = BB
ws["B14"] = "Remaining of $250 ceiling"; ws["F14"] = "=250-F13"; ws["F14"].number_format = USD

wb.save("Master_Business_Dashboard.xlsx")
print("saved")
