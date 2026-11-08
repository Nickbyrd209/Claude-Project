"""QA for the Home Daycare Tax-Ready Records Workbook.

Fills a copy of the workbook with synthetic TEST data (not customer data), recalculates it in
LibreOffice, and compares every key output with an independent Python re-implementation.

Usage: python3 test_workbook.py <workbook.xlsx> <recalc.py> <scratch_dir>
"""
import json, random, shutil, subprocess, sys, os
from datetime import date, timedelta
from openpyxl import load_workbook

SRC, RECALC, SCRATCH = sys.argv[1], sys.argv[2], sys.argv[3]
FIRST = 5
results = []


def check(name, got, want, tol=0.01):
    if isinstance(want, str):
        ok = got == want
    else:
        ok = got is not None and abs(float(got) - float(want)) <= tol
    results.append((name, ok, got, want))


def recalc(path):
    out = subprocess.run([sys.executable, RECALC, path, "90"], capture_output=True, text=True)
    info = json.loads(out.stdout[out.stdout.index("{"):])
    return info


def blank_examples(wb):
    for sh, cols in [("Families", "ABCD"), ("Time Log", "CDF"), ("Income", "ABCDEF"), ("Expenses", "ABCDEFH"),
                     ("Meals", "BCDEG"), ("Mileage", "ABCF")]:
        for c in cols:
            wb[sh][f"{c}{FIRST}"] = None


def build_case(name, mutate):
    path = os.path.join(SCRATCH, f"qa_{name}.xlsx")
    shutil.copy(SRC, path)
    wb = load_workbook(path)
    blank_examples(wb)
    model = mutate(wb)
    wb.save(path)
    info = recalc(path)
    vals = load_workbook(path, data_only=True)
    return info, vals, model


# ------------------------------------------------------------------ independent model
def model_summary(m):
    space = min(1, m["dc_sqft"] / m["home_sqft"]) if m["home_sqft"] > 0 else 0
    hours = sum(c + p for c, p in m["weeks"])
    time = min(1, hours / (m["days"] * 24)) if m["days"] > 0 else 0
    ts = time * space
    inc = {}
    for d, fam, t, amt in m["income"]:
        inc[t] = inc.get(t, 0) + amt
    total_inc = sum(inc.values())
    biz, home, food_actual = {}, {}, 0.0
    for d, cat, use, amt in m["expenses"]:
        if cat.startswith("Home:"):
            home[cat] = home.get(cat, 0) + amt * ts
            continue
        if cat == "Food":
            if use == "Business only (100%)": food_actual += amt
            elif use == "Shared (Time-Space %)": food_actual += amt * ts
            if m["food"] == "Standard Meal Allowance":
                continue
        f = 1 if use == "Business only (100%)" else ts if use == "Shared (Time-Space %)" else 0
        biz[cat] = biz.get(cat, 0) + amt * f
    meals = sum(b * m["rb"] + (l + s) * m["rl"] + sn * m["rs"] for b, l, s, sn in m["meals"])
    switch = date(2026, 7, 1)
    miles = sum(mi * (m["m2"] if d >= switch else m["m1"]) for d, mi in m["miles"])
    exp_total = sum(biz.values()) + (meals if m["food"] == "Standard Meal Allowance" else 0) + miles
    tent = total_inc - exp_total
    home_total = sum(home.values()) + m.get("depr", 0)
    net = tent - home_total
    se = 0 if net <= 0 else min(net * 0.9235, 184500) * 0.124 + net * 0.9235 * 0.029
    it = max(0, net) * m["inc_pct"]
    fam_paid = {}
    for d, fam, t, amt in m["income"]:
        if t == "Parent Payment":
            fam_paid[fam] = fam_paid.get(fam, 0) + amt
    return dict(space=space, hours=hours, time=time, ts=ts, inc=inc, total_inc=total_inc, biz=biz, home=home,
                meals=meals, miles=miles, exp_total=exp_total, tent=tent, home_total=home_total, net=net, se=se,
                it=it, sa=se + it, q=(se + it) / 4, food_actual=food_actual, fam_paid=fam_paid,
                limit="Over limit — preparer must apply the Form 8829 limit" if home_total > max(0, tent) else "OK")


FAMILIES = ["Rivera", "Okafor", "Nguyen", "Schmidt"]
CATS = ["Educational & Activity Supplies", "Toys & Play Equipment", "Household & Cleaning Supplies", "Food",
        "Insurance (liability, not homeowners)", "Licenses, Fees & Dues", "Office Supplies & Postage"]
HOME = ["Home: Mortgage Interest", "Home: Real Estate Taxes", "Home: Utilities (gas, electric, water, trash)",
        "Home: Homeowners/Renters Insurance"]
USES = ["Business only (100%)", "Shared (Time-Space %)", "Personal (not deductible)"]


def fill(wb, m):
    st = wb["Setup"]
    st["B13"], st["B14"], st["B18"], st["B26"], st["B39"] = m["home_sqft"], m["dc_sqft"], m["days"], m["food"], m["inc_pct"]
    for i, f in enumerate(FAMILIES):
        r = FIRST + i
        wb["Families"][f"A{r}"], wb["Families"][f"B{r}"] = f, f"Parent {f}"
        wb["Families"][f"C{r}"], wb["Families"][f"D{r}"] = f"Child {f}", "1 Test Rd"
    for i, (c, p) in enumerate(m["weeks"]):
        wb["Time Log"][f"C{FIRST+i}"], wb["Time Log"][f"D{FIRST+i}"] = c, p
    for i, (d, fam, t, amt) in enumerate(m["income"]):
        r = FIRST + i
        ws = wb["Income"]; ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = d, fam, t, amt
    for i, (d, cat, use, amt) in enumerate(m["expenses"]):
        r = FIRST + i
        ws = wb["Expenses"]; ws[f"A{r}"], ws[f"D{r}"], ws[f"E{r}"], ws[f"F{r}"] = d, cat, use, amt
    for i, (b, l, s, sn) in enumerate(m["meals"]):
        r = FIRST + i
        ws = wb["Meals"]; ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"], ws[f"E{r}"] = b, l, s, sn
    for i, (d, mi) in enumerate(m["miles"]):
        r = FIRST + i
        ws = wb["Mileage"]; ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = d, "test trip", mi
    if "depr" in m:
        wb["Tax Summary"]["B47"] = m["depr"]
    wb["Parent Statements"]["C2"] = m.get("stmt_family", "Rivera")


def base_model(seed=7, food="Standard Meal Allowance", scale=1.0):
    rnd = random.Random(seed)
    m = dict(home_sqft=2000, dc_sqft=1500, days=365, food=food, inc_pct=0.10, rb=1.70, rl=3.22, rs=0.96,
             m1=0.725, m2=0.76, depr=850.0)
    m["weeks"] = [(55.0, 12.0)] * 50 + [(0.0, 4.0), (0.0, 0.0)]  # 2 vacation weeks, week 53 blank
    inc, d0 = [], date(2026, 1, 2)
    for w in range(50):
        d = d0 + timedelta(days=7 * w)
        for f in FAMILIES[:3]:
            inc.append((d, f, "Parent Payment", round(250 * scale, 2)))
        inc.append((d, "Schmidt", "Subsidy / Voucher", round(180 * scale, 2)))
    for mo in range(1, 13):
        inc.append((date(2026, mo, 20), "", "CACFP Reimbursement", 412.35))
    inc.append((date(2026, 6, 1), "", "Other Business Income", 300.0))
    m["income"] = inc[:300]
    exp = []
    for i in range(120):
        exp.append((date(2026, 1, 1) + timedelta(days=rnd.randint(0, 364)), rnd.choice(CATS), rnd.choice(USES), round(rnd.uniform(5, 300), 2)))
    for mo in range(1, 13):
        for h, amt in zip(HOME, [900.0, 350.0, 240.0, 110.0]):
            exp.append((date(2026, mo, 5), h, "", amt))
    m["expenses"] = exp
    m["meals"] = [(rnd.randint(60, 90), rnd.randint(80, 110), rnd.randint(0, 10), rnd.randint(150, 220)) for _ in range(12)]
    m["miles"] = [(date(2026, 1, 1) + timedelta(days=rnd.randint(0, 364)), round(rnd.uniform(2, 40), 1)) for _ in range(80)]
    return m


def compare(label, vals, m):
    e = model_summary(m)
    ts = vals["Tax Summary"]
    check(f"{label}: space %", vals["Setup"]["B15"].value, e["space"], 1e-9)
    check(f"{label}: business hours", vals["Time Log"]["E3"].value, e["hours"], 1e-6)
    check(f"{label}: time %", vals["Setup"]["B21"].value, e["time"], 1e-9)
    check(f"{label}: time-space %", vals["Setup"]["B23"].value, e["ts"], 1e-9)
    check(f"{label}: total income", ts["B15"].value, e["total_inc"])
    check(f"{label}: meal allowance", vals["Meals"]["F17"].value, e["meals"])
    check(f"{label}: mileage deduction", vals["Mileage"]["E3"].value, e["miles"])
    check(f"{label}: expenses before home", ts["B37"].value, e["exp_total"])
    check(f"{label}: tentative profit", ts["B38"].value, e["tent"])
    check(f"{label}: home expenses (business share)", ts["B48"].value, e["home_total"])
    check(f"{label}: 8829 limit flag", ts["B49"].value, e["limit"])
    check(f"{label}: net profit", ts["B51"].value, e["net"])
    check(f"{label}: actual food comparison", ts["B54"].value, e["food_actual"])
    check(f"{label}: SE tax estimate", ts["B59"].value, e["se"])
    check(f"{label}: quarterly set-aside", ts["B62"].value, e["q"])
    fam = m.get("stmt_family", "Rivera")
    check(f"{label}: parent statement total ({fam})", vals["Parent Statements"]["C17"].value, e["fam_paid"].get(fam, 0))
    for i, f in enumerate(FAMILIES):
        got = vals["Families"][f"E{FIRST+i}"].value
        check(f"{label}: Families paid {f}", got, e["fam_paid"].get(f, 0))
    # each Schedule C row vs model
    for r in range(18, 35):
        cat = ts[f"A{r}"].value
        check(f"{label}: category {cat}", ts[f"B{r}"].value, e["biz"].get(cat, 0))


def run():
    cases = []

    def case_a(wb):
        m = base_model(); fill(wb, m); return m
    cases.append(("A_full_year_standard_meals", case_a))

    def case_b(wb):
        m = base_model(food="Actual Food Cost"); fill(wb, m); return m
    cases.append(("B_actual_food", case_b))

    def case_c(wb):  # edge: no sq ft, loss year, family with no payments selected
        m = base_model(scale=0.1); m["home_sqft"] = 0; m["stmt_family"] = "Schmidt"; fill(wb, m); return m
    cases.append(("C_zero_sqft_loss_year", case_c))

    def case_d(wb):  # high income -> SS wage base cap
        m = base_model(scale=20.0); fill(wb, m); return m
    cases.append(("D_high_income_wage_base_cap", case_d))

    def case_e(wb):  # partial year + over-168 guard is visual only; check time % cap at 100%
        m = base_model(); m["days"] = 100; m["stmt_family"] = "Nobody"; fill(wb, m); return m
    cases.append(("E_partial_year_unknown_family", case_e))

    for name, fn in cases:
        info, vals, m = build_case(name, fn)
        results.append((f"{name}: recalc status", info.get("status") == "success" and info.get("total_errors") == 0, info.get("total_errors"), 0))
        compare(name, vals, m)
        if name.startswith("E_"):
            check(f"{name}: unknown-family warning shown", vals["Parent Statements"]["B26"].value,
                  "⚠ Family not found on Families tab — check spelling.")
        if name.startswith("B_"):
            check(f"{name}: meal allowance row zero under actual method", vals["Tax Summary"]["B35"].value, 0)
        if name.startswith("C_"):
            check(f"{name}: SE tax zero on loss", vals["Tax Summary"]["B59"].value, 0)

    # shipping file: example rows only, no errors
    info = recalc(SRC)
    results.append(("Shipping file recalculates with zero errors", info.get("status") == "success" and info.get("total_errors") == 0, info.get("total_errors"), 0))

    passed = sum(1 for r in results if r[1])
    lines = [f"# Automated QA results\n", f"Passed {passed} of {len(results)} checks.\n", "| Check | Result | Workbook | Independent model |", "|---|---|---|---|"]
    for n, ok, got, want in results:
        g = f"{got:.4f}" if isinstance(got, float) else got
        w = f"{want:.4f}" if isinstance(want, float) else want
        lines.append(f"| {n} | {'PASS' if ok else 'FAIL'} | {g} | {w} |")
    print("\n".join(lines))
    return passed == len(results)


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
