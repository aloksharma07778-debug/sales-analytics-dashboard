"""
Sales Analytics Project - ETL Script
Reproduces the Power Query cleaning steps from Step_by_step_approach.docx
in Python/pandas, so the same logic can run outside Power BI (and the
cleaned output can still be imported straight into Power BI).
"""
import pandas as pd
import numpy as np
import re

RAW = "../data/raw/"
OUT = "../data/cleaned/"

# ---------- STEP 1: LOAD ----------
def load_messy_excel_as_csv(path):
    """These sample xlsx files store one comma-joined text column per row
    (header row is the column list). Parse them like a CSV."""
    raw = pd.read_excel(path, header=None)[0].astype(str)
    from io import StringIO
    text = "\n".join(raw.tolist())
    return pd.read_csv(StringIO(text))

customers = load_messy_excel_as_csv(RAW + "Customers.xlsx")
products = load_messy_excel_as_csv(RAW + "Products.xlsx")
targets = load_messy_excel_as_csv(RAW + "Targets.xlsx")
def load_quoted_row_csv(path):
    """Each line in these sample CSVs is wrapped entirely in one set of
    quotes (commas inside are the real delimiters, not escaped commas),
    so parse it as plain text rather than standard CSV quoting."""
    with open(path, "r", encoding="utf-8-sig") as f:
        lines = [l.strip().strip('"') for l in f if l.strip()]
    from io import StringIO
    return pd.read_csv(StringIO("\n".join(lines)))

inventory = load_quoted_row_csv(RAW + "Inventory.csv")
sales = load_quoted_row_csv(RAW + "Sales_Transactions.csv")

# Strip stray inline comments accidentally left in sample data (e.g. "<-- duplicate")
def strip_comment(val):
    if isinstance(val, str) and "<--" in val:
        return val.split("<--")[0].strip()
    return val

for df in (customers, products, sales):
    for col in df.columns:
        df[col] = df[col].apply(strip_comment)

# ---------- STEP 2: DATA CLEANING (ETL) ----------

# Fix mixed date formats in Sales (2024-01-05, 2024/01/07, 05-02-2024 ...)
def parse_mixed_date(val):
    val = str(val).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return pd.to_datetime(val, format=fmt)
        except ValueError:
            continue
    return pd.to_datetime(val, errors="coerce", dayfirst=True)

sales["OrderDate"] = sales["OrderDate"].apply(parse_mixed_date)
customers["JoinDate"] = pd.to_datetime(customers["JoinDate"], errors="coerce")
inventory["StockDate"] = pd.to_datetime(inventory["StockDate"], errors="coerce")

# Remove duplicates (exact duplicate rows, e.g. order 1006/1007)
before = len(sales)
sales = sales.drop_duplicates(subset=[c for c in sales.columns if c != "OrderID"], keep="first")
dupes_removed = before - len(sales)

# Handle nulls
customers["Region"] = customers["Region"].fillna("Unknown")
products["Category"] = products["Category"].fillna("Others")
sales["Discount"] = pd.to_numeric(sales["Discount"], errors="coerce").fillna(0)
sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce")
sales["UnitPrice"] = pd.to_numeric(sales["UnitPrice"], errors="coerce")
sales["Cost"] = pd.to_numeric(sales["Cost"], errors="coerce")

# Handle returns: negative Quantity = return, flag it, keep for revenue math (Power BI does the same)
sales["IsReturn"] = sales["Quantity"] < 0

# Calculated columns
sales["SalesAmount"] = sales["Quantity"] * sales["UnitPrice"] * (1 - sales["Discount"])
sales["Profit"] = (sales["UnitPrice"] - sales["Cost"]) * sales["Quantity"]

# Merge tables (Sales + Products -> Category/SubCategory, Sales + Customers -> Region)
sales_full = sales.merge(products, on="ProductID", how="left") \
                   .merge(customers[["CustomerID", "CustomerName", "Region"]], on="CustomerID", how="left")

# ---------- STEP 3: DATE TABLE ----------
date_table = pd.DataFrame({"Date": pd.date_range("2024-01-01", "2024-12-31", freq="D")})
date_table["Year"] = date_table["Date"].dt.year
date_table["Month"] = date_table["Date"].dt.strftime("%b")
date_table["MonthNumber"] = date_table["Date"].dt.month

# ---------- WRITE CLEANED WORKBOOK ----------
with pd.ExcelWriter(OUT + "Cleaned_Sales_Data.xlsx", engine="openpyxl") as writer:
    sales.to_excel(writer, sheet_name="Sales", index=False)
    sales_full.to_excel(writer, sheet_name="Sales_Full (merged)", index=False)
    customers.to_excel(writer, sheet_name="Customers", index=False)
    products.to_excel(writer, sheet_name="Products", index=False)
    inventory.to_excel(writer, sheet_name="Inventory", index=False)
    targets.to_excel(writer, sheet_name="Targets", index=False)
    date_table.to_excel(writer, sheet_name="DateTable", index=False)

print(f"Duplicates removed: {dupes_removed}")
print(f"Rows -> Sales: {len(sales)}, Customers: {len(customers)}, Products: {len(products)}, Inventory: {len(inventory)}, Targets: {len(targets)}")
print("Cleaned workbook written to", OUT + "Cleaned_Sales_Data.xlsx")
