"""
Generate a realistic, fully fictional 12-month dataset for one Saudi FMCG wholesaler:
  data/invoices.xlsx        - sales invoices + credit notes (what the SME would upload)
  data/bank_statement.csv   - matching bank transactions (stand-in for parsed statement PDFs)
  data/bank_statement.pdf   - the same transactions rendered as 12 monthly statements
  data/sample-data.js       - both datasets as JSON, loaded by the prototype

Every name, VAT number and IBAN is invented. Deterministic: same seed -> same files.
Run:  python3 scripts/generate_data.py
"""
import csv, json, math, random
from datetime import date, timedelta
from pathlib import Path

random.seed(20261024)
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data"
OUT.mkdir(exist_ok=True)

START, END = date(2025, 10, 1), date(2026, 9, 30)   # 12 months; END is the "as of" date
VAT = 0.15
COMPANY = {
    "name": "Rimal Al-Sharq Trading Co. (fictional)",
    "cr": "1010999001", "vat": "310999000100003", "city": "Riyadh",
    "sector": "FMCG / food wholesale", "founded": 2017,
    "bank": "Sahra Bank (fictional)", "iban": "SA00 9900 0000 0012 3456 7890",
}

# --- Buyers: name, city, segment, weight (share of sales), terms (days), behaviour ------------
# behaviour: avg days paid AFTER due date (negative = early), sd, share of invoices never collected
BUYERS = [
    ("Al-Waha Hypermarkets", "Riyadh", "Hypermarket chain", 17, 60, 12, 6, 0.0),
    ("Sanabel Supermarkets", "Riyadh", "Supermarket chain", 11, 45, 4, 4, 0.0),
    ("Dar Al-Nakhla Grocery Group", "Al-Kharj", "Supermarket chain", 8, 45, 9, 6, 0.0),
    ("Ghaf Tree Markets", "Riyadh", "Supermarket chain", 6, 30, 2, 3, 0.0),
    ("Najd Fresh Minimarkets", "Riyadh", "Minimarket chain", 5, 30, 18, 10, 0.01),
    ("Qimam HoReCa Supplies", "Riyadh", "HoReCa distributor", 5, 30, 3, 4, 0.0),
    ("Sadu Catering Co.", "Riyadh", "Catering", 4, 30, 6, 5, 0.0),
    ("Bayt Al-Tamr Stores", "Buraydah", "Supermarket", 4, 30, 1, 3, 0.0),
    ("Areej Family Market", "Riyadh", "Supermarket", 3, 30, 5, 5, 0.0),
    ("Nakheel Corner Mini Mart", "Riyadh", "Minimarket", 3, 30, 0, 3, 0.0),   # seeded: 'paid' but no bank trace
    ("Fanar Restaurants Group", "Riyadh", "HoReCa", 3, 30, 35, 15, 0.04),    # seeded: chronic late payer
    ("Hadeel Grocery", "Al-Majma'ah", "Grocery", 2, 15, 2, 3, 0.0),
    ("Marafi Supermarket", "Riyadh", "Supermarket", 2, 30, 4, 4, 0.0),
    ("Liwan Hotel Supplies", "Riyadh", "HoReCa", 2, 30, 8, 6, 0.0),
    ("Shams Al-Wadi Market", "Al-Kharj", "Supermarket", 2, 30, 3, 4, 0.0),
    ("Yasmin Bakery Chain", "Riyadh", "Bakery chain", 2, 15, 1, 2, 0.0),
    ("Tuwaiq Mini Markets", "Riyadh", "Minimarket chain", 2, 30, 10, 8, 0.0),
    ("Rimth Grocery", "Shaqra", "Grocery", 1.5, 15, 2, 3, 0.0),
    ("Al-Bustan Corner", "Riyadh", "Minimarket", 1.5, 15, 3, 3, 0.0),
    ("Mawj Coffee Roasters", "Riyadh", "HoReCa", 1.5, 30, 0, 3, 0.0),
    ("Sidra Market", "Riyadh", "Grocery", 1, 15, 4, 4, 0.0),
    ("Diyar Al-Khair Store", "Diriyah", "Grocery", 1, 15, 2, 3, 0.0),
    ("Nadeem Restaurants", "Riyadh", "HoReCa", 1, 30, 20, 10, 0.02),
    ("Lamar Mini Mart", "Riyadh", "Minimarket", 1, 0, 0, 0, 0.0),
    ("Barakah Groceries", "Riyadh", "Grocery", 1, 0, 0, 0, 0.0),
    ("Sahl Al-Thumama Store", "Riyadh", "Grocery", 1, 0, 0, 0, 0.0),
    ("Wadi Hanifa Market", "Riyadh", "Grocery", 1, 0, 0, 0, 0.0),
    ("Jood Corner Shop", "Riyadh", "Grocery", 0.8, 0, 0, 0, 0.0),
    ("Fajr Al-Riyadh Store", "Riyadh", "Grocery", 0.8, 0, 0, 0, 0.0),
    ("Oud Al-Qasr Market", "Al-Kharj", "Grocery", 0.8, 0, 0, 0, 0.0),
    ("Samar Minimarket", "Riyadh", "Minimarket", 0.7, 15, 5, 4, 0.0),
    ("Hanaa Grocery", "Riyadh", "Grocery", 0.7, 15, 3, 3, 0.0),
    ("Raed Event Catering", "Riyadh", "Catering", 0.7, 30, 14, 8, 0.0),
    ("Masa Sweets", "Riyadh", "Bakery", 0.5, 15, 2, 2, 0.0),
]
def vat_no(i): return f"3{100200300 + i * 7919:09d}0000{(i % 9) + 1}"[:15].ljust(15, "3")
buyers = []
for i, (n, c, seg, w, terms, late, sd, bad) in enumerate(BUYERS):
    buyers.append(dict(id=f"B{i+1:03d}", name=n, city=c, segment=seg, weight=w, terms=terms,
                       late=late, sd=sd, bad=bad, vat=vat_no(i + 1), cash=(terms == 0)))
W = sum(b["weight"] for b in buyers)

CATS = [("Rice & grains", .24), ("Cooking oil & ghee", .16), ("Dates", .12), ("Beverages", .14),
        ("Long-life dairy", .12), ("Canned food", .10), ("Sugar & flour", .07), ("Cleaning supplies", .05)]

# --- Monthly sales targets (net of VAT, SAR) ------------------------------------------------
# Seasonality: Ramadan 1447 ~ 18 Feb - 19 Mar 2026 (stocking up in Jan/Feb), Eid al-Adha ~ 27 May 2026,
# summer slowdown Jul-Aug, back-to-school Sep. Underlying growth ~1.2%/month.
SEASON = {10: 0.97, 11: 0.98, 12: 1.00, 1: 1.12, 2: 1.38, 3: 1.18, 4: 0.92, 5: 1.10, 6: 0.95, 7: 0.84, 8: 0.86, 9: 1.03}
BASE = 1_250_000

def months(pre=0):
    y, m = START.year, START.month
    for _ in range(pre):
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    for _ in range(12 + pre):
        yield y, m
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)

def month_days(y, m):
    d = date(y, m, 1)
    while d.month == m:
        yield d
        d += timedelta(days=1)

invoices, seq = [], 0
# Invoices start 2 months BEFORE the window so that collections are already flowing in Oct 2025
# (opening receivables). Those pre-window invoices are not in the upload, only their payments in the bank.
for k, (y, m) in enumerate(months(pre=2)):
    target = BASE * SEASON[m] * (1.012 ** (k - 2)) * random.uniform(0.97, 1.03)
    days = [d for d in month_days(y, m) if d.weekday() != 4]          # no deliveries on Friday
    spent = 0
    while spent < target:
        b = random.choices(buyers, weights=[x["weight"] for x in buyers])[0]
        # bigger buyers place bigger orders
        size = random.lognormvariate(math.log(4200 + 260 * b["weight"]), 0.5)
        net = round(size / 5) * 5 + random.choice([0, 0.25, 0.5, 0.75])
        d = random.choice(days)
        seq += 1
        cat = random.choices([c for c, _ in CATS], weights=[w for _, w in CATS])[0]
        invoices.append(dict(invoice_no=f"INV-{d.year}-{seq:05d}", issue_date=d, buyer=b, category=cat,
                             net=round(net, 2)))
        spent += net

invoices.sort(key=lambda r: (r["issue_date"], r["invoice_no"]))
# renumber in date order so numbers are sequential like a real ERP
for i, r in enumerate(invoices, 1):
    r["invoice_no"] = f"INV-{r['issue_date'].year}-{i:05d}"

# --- Payment behaviour (pre-window invoices included: their payments land in the window) --------------------------------------------------------------------
for r in invoices:
    b = r["buyer"]
    r["vat"] = round(r["net"] * VAT, 2)
    r["total"] = round(r["net"] + r["vat"], 2)
    r["terms"] = b["terms"]
    r["due_date"] = r["issue_date"] + timedelta(days=b["terms"])
    if b["cash"]:
        r["pay_date"] = r["issue_date"]
    elif random.random() < b["bad"]:
        r["pay_date"] = None                                          # never collected
    else:
        lag = max(-5, int(random.gauss(b["late"], b["sd"])))
        r["pay_date"] = r["due_date"] + timedelta(days=lag)
    r["paid_amount"] = r["total"]
    # ~3% of invoices are partly paid (short delivery / damaged goods deducted by buyer)
    if r["pay_date"] and not b["cash"] and random.random() < 0.03:
        r["paid_amount"] = round(r["total"] * random.uniform(0.85, 0.97), 2)
    if r["pay_date"] and r["pay_date"] > END:
        r["pay_date"], r["paid_amount"] = None, 0.0                   # not yet paid at as-of date
    if r["pay_date"] is None:
        r["paid_amount"] = 0.0

def status(r):
    if r["paid_amount"] >= r["total"] - 0.01: return "Paid"
    if r["paid_amount"] > 0: return "Partially paid"
    return "Overdue" if r["due_date"] < END else "Unpaid"

pre_invoices = [r for r in invoices if r["issue_date"] < START]
invoices = [r for r in invoices if r["issue_date"] >= START]

# --- Credit notes (returns / expired stock), ~1.8% of sales -------------------------------
credit_notes = []
for r in random.sample(invoices, k=int(len(invoices) * 0.04)):
    if r["buyer"]["cash"]: continue
    amt = round(r["net"] * random.uniform(0.2, 0.6), 2)
    d = min(END, r["issue_date"] + timedelta(days=random.randint(5, 25)))
    credit_notes.append(dict(invoice_no=f"CN-{d.year}-{len(credit_notes)+1:04d}", issue_date=d, buyer=r["buyer"],
                             category=r["category"], net=-amt, vat=-round(amt * VAT, 2),
                             total=-round(amt * (1 + VAT), 2), terms=0, due_date=d, pay_date=None,
                             paid_amount=0.0, ref=r["invoice_no"]))

# --- Bank statement -----------------------------------------------------------------------
tx = []
def add(d, desc, amount, kind, counterparty=""):
    if d < START or d > END: return
    tx.append(dict(date=d, description=desc, amount=round(amount, 2), kind=kind, counterparty=counterparty))

# Seeded issue: Nakheel Corner Mini Mart's invoices are marked paid in the sheet, but none of
# that money ever reaches this account (could be cash kept off-book, another account, or fake invoices).
GHOST = "Nakheel Corner Mini Mart"

# Credit-term buyers: one transfer per (buyer, pay date) - often covers several invoices.
from collections import defaultdict
by_pay = defaultdict(list)
for r in pre_invoices + invoices:
    if r["pay_date"] and not r["buyer"]["cash"] and r["buyer"]["name"] != GHOST:
        pd = r["pay_date"]
        while pd.weekday() in (4, 5): pd += timedelta(days=1)     # banks: transfers land Sun-Thu
        if pd > END: continue
        by_pay[(r["buyer"]["name"], pd)].append(r)
for (name, pd), rows in by_pay.items():
    amt = sum(x["paid_amount"] for x in rows)
    refs = " ".join(x["invoice_no"][-5:] for x in rows[:4]) + (" +" if len(rows) > 4 else "")
    style = random.random()
    if style < 0.6:
        desc = f"SARIE TRF IN {name.upper()} INV {refs}"
    elif style < 0.85:
        desc = f"SARIE TRF IN {name.upper()}"
    else:                                                          # truncated / abbreviated payer name
        desc = f"IPS IN {name.upper()[:14]}"
    add(pd, desc, amt, "customer_payment", name)

# Cash buyers: driver collects cash, deposited at branch every Sun & Wed (lumped).
cash_by_day = defaultdict(float)
for r in pre_invoices + invoices:
    if r["buyer"]["cash"]:
        d = r["issue_date"] + timedelta(days=1)
        while d.weekday() not in (6, 2): d += timedelta(days=1)    # Sun=6, Wed=2
        if d <= END: cash_by_day[d] += r["total"]
for d, amt in cash_by_day.items():
    add(d, "CASH DEPOSIT - BRANCH 214 OLAYA", amt * random.uniform(0.985, 1.0), "cash_deposit")

# Non-sales credits (must NOT be counted as revenue)
add(date(2026, 1, 12), "TRF FROM OWN ACCT SA00..5521 TOP UP", 300000, "own_transfer")
add(date(2026, 2, 3), "OWNER CAPITAL INJECTION - F. ALHARBI", 250000, "owner_injection")
add(date(2025, 11, 18), "VAT REFUND ZATCA", 18450.00, "tax_refund")

# Debits
monthly_sales = defaultdict(float)
for r in invoices: monthly_sales[(r["issue_date"].year, r["issue_date"].month)] += r["net"]
prev3 = lambda y, m: [((y, m - i) if m - i > 0 else (y - 1, m - i + 12)) for i in (1, 2, 3)]
for (y, m) in months():
    s = monthly_sales[(y, m)]
    # supplier payments ~79% of sales, a few per week; heavier before Ramadan
    pays = [d for d in month_days(y, m) if d.weekday() in (6, 1, 3)]
    suppliers = ["NAJD FOOD INDUSTRIES (F)", "GULF GRAINS IMPORT (F)", "AL-SAFA OILS (F)", "TAMR VALLEY FARMS (F)",
                 "HAYAT BEVERAGES (F)", "SAHARA PACKAGING (F)"]
    total_sup = s * 1.15 * random.uniform(0.745, 0.765)
    shares = [random.random() for _ in pays]
    for d, sh in zip(pays, shares):
        add(d, f"SARIE TRF OUT {random.choice(suppliers)}", -total_sup * sh / sum(shares), "supplier_payment")
    add(date(y, m, 27) if date(y, m, 27).weekday() not in (4, 5) else date(y, m, 25), "PAYROLL WPS 23 EMPLOYEES",
        -random.uniform(93000, 96500), "payroll")
    add(date(y, m, 10), "GOSI CONTRIBUTIONS", -10900.00, "gosi")
    add(date(y, m, 5), "LOAN INSTALMENT REF TL-77812", -42500.00, "loan_repayment")   # existing term loan
    add(date(y, m, 15), "FUEL & FLEET - PETROL CARD", -random.uniform(21000, 26000), "operating")
    add(date(y, m, 20), "STC BUSINESS / UTILITIES", -random.uniform(6800, 9200), "operating")
    add(date(y, m, 28), "BANK CHARGES", -random.uniform(180, 420), "bank_fee")
    if m in (10, 1, 4, 7):
        add(date(y, m, 3), "WAREHOUSE RENT SULAY Q", -120000.00, "rent")
    if m in (10, 1, 4, 7):   # quarterly VAT for previous quarter
        out_vat = sum(monthly_sales.get(k, BASE) for k in prev3(y, m)) * VAT
        add(date(y, m, 25), "ZATCA VAT PAYMENT", -out_vat * 0.20, "tax")   # output VAT less input VAT

tx.sort(key=lambda t: (t["date"], -t["amount"]))
bal = 450_118.40
for t in tx:
    bal = round(bal + t["amount"], 2)
    t["balance"] = bal

# --- Seeded spreadsheet data-quality issues (documented in README) --------------------------
rows = []
def inv_row(r, buyer_name=None):
    return {
        "Invoice No": r["invoice_no"], "Issue Date": r["issue_date"], "Due Date": r["due_date"],
        "Customer": buyer_name or r["buyer"]["name"], "Customer VAT No": r["buyer"]["vat"],
        "City": r["buyer"]["city"], "Category": r["category"], "Payment Terms (days)": r["terms"],
        "Net Amount (SAR)": r["net"], "VAT 15% (SAR)": r["vat"], "Total (SAR)": r["total"],
        "Status": status(r) if not r["invoice_no"].startswith("CN") else "Credit note",
        "Paid Date": r["pay_date"], "Paid Amount (SAR)": r["paid_amount"],
        "Reference": r.get("ref", ""),
    }
for r in invoices:
    name = r["buyer"]["name"]
    # 1) the largest buyer appears under a second spelling on ~30% of rows (same VAT no.)
    if name == "Al-Waha Hypermarkets" and random.random() < 0.3:
        name = "Al Waha Hypermarket Co."
    rows.append(inv_row(r, name))
# 2) duplicated rows (re-exported)
for r in random.sample(invoices, 5):
    rows.append(inv_row(r))
# 3) VAT not equal to 15% of net on a few rows (typed manually)
for row in random.sample(rows, 4):
    row["VAT 15% (SAR)"] = round(row["Net Amount (SAR)"] * 0.05, 2)
    row["Total (SAR)"] = round(row["Net Amount (SAR)"] + row["VAT 15% (SAR)"], 2)
for c in credit_notes:
    rows.append(inv_row(c))
rows.sort(key=lambda x: (x["Issue Date"], x["Invoice No"]))

# --- Write Excel ----------------------------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
wb = Workbook(); ws = wb.active; ws.title = "Invoices"
hdr = list(rows[0].keys())
ws.append(hdr)
for row in rows:
    ws.append([row[h] for h in hdr])
for c in ws[1]:
    c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E46")
    c.alignment = Alignment(wrap_text=True, vertical="center")
for col, h in enumerate(hdr, 1):
    ws.column_dimensions[get_column_letter(col)].width = 14 if "Date" in h else (30 if h == "Customer" else 16)
    for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
        for x in cell:
            if "Date" in h: x.number_format = "yyyy-mm-dd"
            if "SAR" in h: x.number_format = "#,##0.00"
ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
info = wb.create_sheet("Company")
for k, v in COMPANY.items(): info.append([k, v])
info.append(["export_note", "Exported from ERP on 2026-10-02 (fictional data)"])
wb.save(OUT / "invoices.xlsx")

# --- Write bank CSV -------------------------------------------------------------------------
with open(OUT / "bank_statement.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["date", "description", "debit", "credit", "balance"])
    for t in tx:
        w.writerow([t["date"].isoformat(), t["description"], f"{-t['amount']:.2f}" if t["amount"] < 0 else "",
                    f"{t['amount']:.2f}" if t["amount"] > 0 else "", f"{t['balance']:.2f}"])

# --- Write bank PDF (12 monthly statements) -------------------------------------------------
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
st = getSampleStyleSheet()
doc = SimpleDocTemplate(str(OUT / "bank_statement.pdf"), pagesize=A4, leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28)
story = []
opening = 450_118.40
for (y, m) in months():
    mt = [t for t in tx if t["date"].year == y and t["date"].month == m]
    story.append(Paragraph(f"<b>{COMPANY['bank']}</b> &mdash; Account statement (FICTIONAL SAMPLE)", st["Title"]))
    story.append(Paragraph(f"{COMPANY['name']} &middot; IBAN {COMPANY['iban']} &middot; Period {date(y,m,1):%d %b %Y} to {list(month_days(y,m))[-1]:%d %b %Y}", st["Normal"]))
    story.append(Paragraph(f"Opening balance: SAR {opening:,.2f} &middot; Closing balance: SAR {mt[-1]['balance']:,.2f}", st["Normal"]))
    story.append(Spacer(1, 8))
    data = [["Date", "Description", "Debit", "Credit", "Balance"]]
    for t in mt:
        data.append([t["date"].isoformat(), t["description"][:52], f"{-t['amount']:,.2f}" if t["amount"] < 0 else "",
                     f"{t['amount']:,.2f}" if t["amount"] > 0 else "", f"{t['balance']:,.2f}"])
    tb = Table(data, colWidths=[58, 270, 72, 72, 80], repeatRows=1)
    tb.setStyle(TableStyle([("FONTSIZE", (0, 0), (-1, -1), 6.5), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E46")),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F5F4")]),
                            ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    story += [tb, PageBreak()]
    opening = mt[-1]["balance"]
doc.build(story)

# --- JSON for the prototype ---------------------------------------------------------------
def j(v): return v.isoformat() if isinstance(v, date) else v
payload = {
    "company": COMPANY, "asOf": END.isoformat(), "periodStart": START.isoformat(),
    "invoices": [{k: j(v) for k, v in r.items()} for r in rows],
    "bank": [{"date": t["date"].isoformat(), "description": t["description"], "amount": t["amount"], "balance": t["balance"]} for t in tx],
}
(OUT / "sample-data.js").write_text("window.SAMPLE_DATA = " + json.dumps(payload) + ";\n")

inv_only = [r for r in rows if r["Invoice No"].startswith("INV")]
print(f"invoice rows: {len(rows)} (invoices {len(inv_only)}, credit notes {len(credit_notes)})")
print(f"net sales: {sum(r['Net Amount (SAR)'] for r in inv_only):,.0f} SAR")
print(f"bank tx: {len(tx)}, closing balance {tx[-1]['balance']:,.2f}, min balance {min(t['balance'] for t in tx):,.2f}")
for (y, m) in months(): print(y, m, f"{monthly_sales[(y, m)]:,.0f}")
