"""
Step 3: clean the USAID SCMS shipment file and load it into SQLite.
Run from the repo root:  python3 python/01_load_scms.py
Reads the CSV in data/, writes data/scms.db (table: shipments).
Every cleaning decision is printed so you can paste it into the README.
"""
import glob, sqlite3, sys
import pandas as pd

files = glob.glob("data/*.csv")
if len(files) != 1:
    sys.exit(f"Expected exactly one CSV in data/, found {len(files)}: {files}")
df = pd.read_csv(files[0], encoding="utf-8-sig")
print(f"Loaded {files[0]}: {len(df):,} rows, {len(df.columns)} columns")

# normalise column names: lowercase, snake_case, no units in parentheses
df.columns = (df.columns.str.strip().str.lower()
              .str.replace(r"\(.*?\)", "", regex=True)
              .str.replace(r"[^a-z0-9]+", "_", regex=True).str.strip("_"))
need = ["vendor", "po_sent_to_vendor_date", "scheduled_delivery_date",
        "delivered_to_client_date", "unit_price", "freight_cost"]
missing = [c for c in need if c not in df.columns]
if missing:
    sys.exit(f"Columns not found: {missing}\nColumns available: {list(df.columns)}")

log = []

# dates: text -> real dates; anything unreadable (N/A, Date Not Captured) -> empty.
# The three columns use DIFFERENT formats in this file, so each gets its own format.
date_formats = {
    "po_sent_to_vendor_date": "%m/%d/%y",        # e.g. 8/27/14
    "scheduled_delivery_date": "%d-%b-%y",       # e.g. 2-Jun-06
    "delivered_to_client_date": "%d-%b-%y",
}
for col, fmt in date_formats.items():
    parsed = pd.to_datetime(df[col], format=fmt, errors="coerce")
    bad = int(parsed.isna().sum())
    log.append(f"{col}: {bad:,} rows not a valid date, set to empty")
    df[col] = parsed

# freight: numbers stay numbers, text ("See ASN...", "Freight Included in Commodity Cost") is kept as a note
df["freight_note"] = df["freight_cost"].where(pd.to_numeric(df["freight_cost"], errors="coerce").isna())
df["freight_usd"] = pd.to_numeric(df["freight_cost"], errors="coerce")
log.append(f"freight_cost: {int(df['freight_usd'].isna().sum()):,} rows are text, "
           f"kept in freight_note, freight_usd left empty")
df["freight_included"] = df["freight_note"].fillna("").str.contains("included", case=False)
log.append(f"freight_included: {int(df['freight_included'].sum()):,} rows say freight is included in "
           f"commodity cost (treated as 0 in landed cost, a rule you can change)")

df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")

# the metrics, exactly as defined in the README
df["lead_days"] = (df["delivered_to_client_date"] - df["po_sent_to_vendor_date"]).dt.days
df["late_days"] = (df["delivered_to_client_date"] - df["scheduled_delivery_date"]).dt.days
df["on_time"] = (df["late_days"] <= 0).where(df["late_days"].notna()).astype("Int64")

for c in ["po_sent_to_vendor_date", "scheduled_delivery_date", "delivered_to_client_date"]:
    df[c] = df[c].dt.strftime("%Y-%m-%d")

con = sqlite3.connect("data/scms.db")
df.to_sql("shipments", con, if_exists="replace", index=False)
con.close()

print("\nCleaning log (paste into your README):")
for line in log:
    print(" -", line)
ok = df["on_time"].dropna()
print(f"\nShipments with scheduled and delivered dates: {len(ok):,}")
print(f"Overall on-time delivery: {100 * ok.mean():.1f}%")
print("Saved data/scms.db, table 'shipments'")
