#!/usr/bin/env python3
"""
Vendor delivery ranking + Excel scorecard
Data: USAID SCMS Delivery History (public, CC BY-ND). Do NOT commit the CSV, only this script + results.

Run:   python build_vendor_scorecard.py path/to/SCMS_Delivery_History_Dataset.csv
Needs: pip install pandas openpyxl
Makes: supplier.db (clean table + SQL ranking), vendor_scorecard.xlsx, cleaning_log.txt

Metric definitions (same as the README, written before looking at results):
  On-time (OTD)  : delivered date <= scheduled date
  Lead time      : delivered date - PO sent date (days)
  Min shipments  : 10 per vendor before a vendor is ranked
"""
import sqlite3
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CSV_PATH = sys.argv[1] if len(sys.argv) > 1 else "SCMS_Delivery_History_Dataset.csv"
MIN_SHIPMENTS = 10
PSEUDO_VENDOR = "SCMS from RDC"          # looks like a distribution centre, not a supplier

log = []
def say(msg):
    print(msg)
    log.append(msg)

# ---------------------------------------------------------------- 1. LOAD
df = pd.read_csv(CSV_PATH)
df.columns = [c.strip() for c in df.columns]
say(f"Loaded {len(df):,} rows, {df.shape[1]} columns, {df['Vendor'].nunique()} distinct vendors")

need = ["Vendor", "PO Sent to Vendor Date", "Scheduled Delivery Date", "Delivered to Client Date",
        "Freight Cost (USD)", "Line Item Value", "Product Group"]
missing = [c for c in need if c not in df.columns]
if missing:
    sys.exit(f"Missing columns {missing}. Columns found: {list(df.columns)}")

# ---------------------------------------------------------------- 2. DATES
# Each date column gets its OWN format list (the PO column uses a different format).
# Anything that is not a date ("N/A - From RDC", "Date Not Captured") becomes NaT (empty), not 0.
def parse_dates(col, formats):
    s = df[col].astype(str).str.strip()
    out = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    for f in formats:
        todo = out.isna()
        out.loc[todo] = pd.to_datetime(s[todo], format=f, errors="coerce")
    bad = s[out.isna()].value_counts()
    say(f"{col}: {out.notna().sum():,} parsed, {out.isna().sum():,} not a date. Top non-date values: "
        + "; ".join(f"'{k}' x{v}" for k, v in bad.head(4).items()))
    return out

df["po_date"]   = parse_dates("PO Sent to Vendor Date",   ["%m/%d/%y", "%m/%d/%Y", "%d-%b-%y"])
df["sched"]     = parse_dates("Scheduled Delivery Date",  ["%d-%b-%y", "%m/%d/%y", "%m/%d/%Y"])
df["delivered"] = parse_dates("Delivered to Client Date", ["%d-%b-%y", "%m/%d/%y", "%m/%d/%Y"])

# ---------------------------------------------------------------- 3. METRICS PER SHIPMENT
both = df["sched"].notna() & df["delivered"].notna()
df["days_late"] = (df["delivered"] - df["sched"]).dt.days.where(both)
df["on_time"]   = (df["days_late"] <= 0).astype("Int64").where(both)        # 1 = on time, 0 = late

df["lead_time_days"] = (df["delivered"] - df["po_date"]).dt.days
impossible = df["lead_time_days"] < 0
say(f"Lead time: {df['lead_time_days'].notna().sum():,} rows have both dates; "
    f"{impossible.sum()} impossible (delivered before PO) set to empty; "
    f"max {df['lead_time_days'].max():.0f} days")
df.loc[impossible, "lead_time_days"] = pd.NA

# ---------------------------------------------------------------- 4. FREIGHT + LANDED COST
# Freight is a mix of numbers and text. Keep the text; apply ONE stated rule:
#   number -> use it | text containing "included" -> $0 (cost is inside the price) | other text -> unknown (empty)
raw = df["Freight Cost (USD)"].astype(str).str.strip()
num = pd.to_numeric(raw, errors="coerce")
included = num.isna() & raw.str.contains("included", case=False, na=False)
df["freight_text"]      = raw.where(num.isna())
df["freight_usd_known"] = num.mask(included, 0.0)           # NaN where text we cannot price
df["line_value"]        = pd.to_numeric(df["Line Item Value"], errors="coerce")
say(f"Freight: {num.isna().sum():,} text values ({num.isna().mean():.0%}); "
    f"{included.sum():,} say included -> treated as $0; "
    f"{df['freight_usd_known'].isna().sum():,} unknown (left empty)")

df["vendor"] = df["Vendor"].astype(str).str.strip()
df["product_group"] = df["Product Group"]
is_rdc = df["vendor"].str.contains(PSEUDO_VENDOR, case=False)
say(f"'{PSEUDO_VENDOR}': {is_rdc.sum():,} rows ({is_rdc.mean():.0%}), excluded from vendor ranking, "
    f"kept in the overall rate")

overall_otd = round(100 * df["on_time"].dropna().astype(int).mean(), 1)
say(f"Overall on-time delivery, all rows: {overall_otd}%")

# ---------------------------------------------------------------- 5. SQLITE + SQL RANKING
cols = ["vendor", "product_group", "po_date", "sched", "delivered", "days_late", "on_time",
        "lead_time_days", "freight_text", "freight_usd_known", "line_value"]
out = df[cols].copy()
for c in ["po_date", "sched", "delivered"]:
    out[c] = out[c].dt.strftime("%Y-%m-%d")
con = sqlite3.connect("supplier.db")
out.to_sql("shipments", con, if_exists="replace", index=False)

RANKING_SQL = f"""
SELECT vendor,
       COUNT(*)                                                    AS shipments,
       ROUND(100.0 * SUM(on_time) / COUNT(*), 1)                   AS otd_pct,
       ROUND(AVG(CASE WHEN days_late > 0 THEN days_late END), 1)   AS avg_days_late_when_late,
       ROUND(AVG(lead_time_days), 1)                               AS avg_lead_time_days,
       COUNT(lead_time_days)                                       AS lead_time_rows,
       ROUND(100.0 * SUM(freight_usd_known) /
             SUM(CASE WHEN freight_usd_known IS NOT NULL THEN line_value END), 1) AS freight_pct_of_value
FROM shipments
WHERE on_time IS NOT NULL
  AND vendor NOT LIKE '%{PSEUDO_VENDOR}%'
GROUP BY vendor
HAVING COUNT(*) >= {MIN_SHIPMENTS}
ORDER BY otd_pct DESC, shipments DESC;
"""
rank = pd.read_sql(RANKING_SQL, con)
con.close()
say(f"Vendors ranked (>= {MIN_SHIPMENTS} shipments, excluding {PSEUDO_VENDOR}): {len(rank)}")
rank.to_csv("vendor_ranking.csv", index=False)
print("\n", rank.to_string(index=False), "\n")

# ---------------------------------------------------------------- 6. EXCEL SCORECARD (live formulas)
wb = Workbook()
sc, asm, lg = wb.active, wb.create_sheet("Assumptions"), wb.create_sheet("Cleaning log")
sc.title = "Scorecard"

blue = Font(color="0000FF")                     # blue = typed input, black = formula
hdr_fill = PatternFill("solid", fgColor="1F3864")
hdr_font = Font(bold=True, color="FFFFFF")

# Assumptions: change a weight and the whole scorecard re-ranks.
asm["A1"], asm["B1"] = "Setting", "Value"
rows = [("Weight: on-time delivery",              0.60),
        ("Weight: avg days late when late",       0.25),
        ("Weight: freight % of value",            0.15),
        ("High-volume threshold (shipments)",     100),
        ("Overall on-time delivery % (all rows)", overall_otd),
        ("Min shipments to rank",                 MIN_SHIPMENTS)]
for i, (k, v) in enumerate(rows, start=2):
    asm[f"A{i}"], asm[f"B{i}"] = k, v
    asm[f"B{i}"].font = blue
asm["A9"]  = "Check: weights must add to 1"; asm["B9"] = "=B2+B3+B4"
asm["A11"] = "Rules"
notes = ["Score parts are 0-100, min-max across ranked vendors. Higher = better for all three.",
         "Vendor never late: avg days late counted as 0. No known freight: freight score set to 50 (neutral).",
         "Freight 'included in commodity cost' counted as $0; other text freight left unknown.",
         f"'{PSEUDO_VENDOR}' excluded from ranking (appears to be a distribution centre, not a supplier).",
         "Weights are a judgment call, not a law. Say so in the interview and show the sensitivity."]
for i, t in enumerate(notes, start=12):
    asm[f"A{i}"] = t
asm.column_dimensions["A"].width = 60

heads = ["Vendor", "Shipments", "OTD %", "Avg days late (when late)", "Avg lead time (days)",
         "Lead-time rows", "Freight % of value",
         "OTD score", "Days-late score", "Freight score", "Composite score", "Rank", "Flag"]
for j, h in enumerate(heads, start=1):
    c = sc.cell(row=1, column=j, value=h)
    c.fill, c.font = hdr_fill, hdr_font
    c.alignment = Alignment(wrap_text=True, vertical="center")

first, last = 2, 1 + len(rank)
rng = lambda col: f"${col}${first}:${col}${last}"
for i, r in enumerate(rank.itertuples(index=False), start=first):
    vals = [r.vendor, r.shipments, r.otd_pct,
            0 if pd.isna(r.avg_days_late_when_late) else r.avg_days_late_when_late,
            None if pd.isna(r.avg_lead_time_days) else r.avg_lead_time_days,
            r.lead_time_rows,
            None if pd.isna(r.freight_pct_of_value) else r.freight_pct_of_value]
    for j, v in enumerate(vals, start=1):
        c = sc.cell(row=i, column=j, value=v)
        if j > 1: c.font = blue
    sc[f"H{i}"] = f"=IFERROR(100*(C{i}-MIN({rng('C')}))/(MAX({rng('C')})-MIN({rng('C')})),50)"
    sc[f"I{i}"] = f"=IFERROR(100*(MAX({rng('D')})-D{i})/(MAX({rng('D')})-MIN({rng('D')})),50)"
    sc[f"J{i}"] = f"=IF(G{i}=\"\",50,IFERROR(100*(MAX({rng('G')})-G{i})/(MAX({rng('G')})-MIN({rng('G')})),50))"
    sc[f"K{i}"] = f"=ROUND(H{i}*Assumptions!$B$2+I{i}*Assumptions!$B$3+J{i}*Assumptions!$B$4,1)"
    sc[f"L{i}"] = f"=RANK(K{i},{rng('K')},0)"
    sc[f"M{i}"] = (f"=IF(AND(B{i}>=Assumptions!$B$5,C{i}<Assumptions!$B$6),\"Watch first: high volume, below overall OTD\","
                   f"IF(B{i}<30,\"Low volume: read with care\",\"\"))")
for j in range(1, 14):
    sc.column_dimensions[get_column_letter(j)].width = 14
sc.column_dimensions["A"].width = 38
sc.column_dimensions["M"].width = 42
sc.row_dimensions[1].height = 45
sc.freeze_panes = "B2"

ch = BarChart(); ch.type = "bar"; ch.title = "On-time delivery % by vendor"
ch.add_data(Reference(sc, min_col=3, min_row=1, max_row=last), titles_from_data=True)
ch.set_categories(Reference(sc, min_col=1, min_row=first, max_row=last))
ch.legend = None; ch.height = max(8, 0.55 * len(rank)); ch.width = 22
sc.add_chart(ch, "O2")

for i, line in enumerate(log, start=1):
    lg.cell(row=i, column=1, value=line)
lg.column_dimensions["A"].width = 150

wb.save("vendor_scorecard.xlsx")
open("cleaning_log.txt", "w").write("\n".join(log))
say("Wrote supplier.db, vendor_ranking.csv, vendor_scorecard.xlsx, cleaning_log.txt")
