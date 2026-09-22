"""
data_processor.py
=================
Complete Data Pipeline & ETL for E-Commerce Sales Analytics.
- Loads and cleans raw data from ecommerce_sales_dataset.xlsx.
- Imputes missing values, removes duplicates, and standardizes formats.
- Performs feature engineering (Gross Sales, Discount Amount, Net Revenue, Dates).
- Performs Customer RFM Segmentation.
- Generates Star Schema (Fact_Sales, Dim_Customers, Dim_Products, Dim_Locations, Dim_Dates).
- Ingests all tables into a local SQLite database (ecommerce_sales.db).
- Exports CSVs for Power BI, Tableau, and Excel reporting.
"""

import os
import sys
import sqlite3
import pandas as pd
import numpy as np

# Ensure UTF-8 output on Windows consoles
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

def run_etl():
    print("=" * 60)
    print("      E-COMMERCE DATA PIPELINE & STAR SCHEMA ETL")
    print("=" * 60)

    # 1. Directories setup
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data")
    processed_dir = os.path.join(data_dir, "processed")
    raw_dir = os.path.join(data_dir, "raw")
    os.makedirs(processed_dir, exist_ok=True)
    os.makedirs(raw_dir, exist_ok=True)

    # 2. File paths
    raw_file = os.path.join(base_dir, "ecommerce_sales_dataset.xlsx")
    if not os.path.exists(raw_file):
        raw_file = os.path.join(raw_dir, "ecommerce_sales_dataset.xlsx")

    if not os.path.exists(raw_file):
        raise FileNotFoundError(f"Source file not found at {raw_file}")

    print(f"\n[1/6] Loading raw dataset: {raw_file}")
    df_raw = pd.read_excel(raw_file)
    print(f"      Initial shape: {df_raw.shape[0]} rows, {df_raw.shape[1]} columns")

    # 3. Data Cleaning
    print("\n[2/6] Cleaning data and handling missing values...")
    df = df_raw.copy()

    # Deduplicate
    dup_count = df.duplicated().sum()
    if dup_count > 0:
        print(f"      Removing {dup_count} exact duplicate rows.")
        df = df.drop_duplicates()

    # Fill missing Categoricals
    if "Customer_Name" in df.columns:
        df["Customer_Name"] = df["Customer_Name"].fillna("Unknown Customer")
    if "City" in df.columns:
        df["City"] = df["City"].fillna("Bengaluru")
    if "State" in df.columns:
        df["State"] = df["State"].fillna("Karnataka")
    if "Payment_Method" in df.columns:
        df["Payment_Method"] = df["Payment_Method"].fillna("UPI")
    if "Order_Status" in df.columns:
        df["Order_Status"] = df["Order_Status"].fillna("Delivered")

    # Category imputation from known products
    product_category_map = df.dropna(subset=["Category"]).drop_duplicates("Product").set_index("Product")["Category"].to_dict()
    df["Category"] = df["Category"].fillna(df["Product"].map(product_category_map)).fillna("General")

    # Numeric imputation
    median_qty = df["Quantity"].median()
    df["Quantity"] = df["Quantity"].fillna(median_qty).astype(int)

    # Date normalization
    df["Order_Date"] = pd.to_datetime(df["Order_Date"])
    df = df.sort_values("Order_Date").reset_index(drop=True)

    # 4. Feature Engineering
    print("\n[3/6] Performing Feature Engineering & Financial Calculations...")
    df["Gross_Sales"] = df["Quantity"] * df["Unit_Price"]
    df["Discount_Amount"] = (df["Gross_Sales"] * df["Discount"]).round(2)
    df["Revenue"] = (df["Gross_Sales"] - df["Discount_Amount"]).round(2)

    # Date dimensions
    df["Year"] = df["Order_Date"].dt.year
    df["Quarter"] = "Q" + df["Order_Date"].dt.quarter.astype(str)
    df["Month"] = df["Order_Date"].dt.month
    df["Month_Name"] = df["Order_Date"].dt.strftime("%B")
    df["YearMonth"] = df["Order_Date"].dt.strftime("%Y-%m")
    df["Day_Of_Week"] = df["Order_Date"].dt.day_name()
    df["Day_Num"] = df["Order_Date"].dt.dayofweek

    # Discount Tier
    def get_discount_tier(d):
        if d == 0:
            return "No Discount (0%)"
        elif d <= 0.10:
            return "Low (<10%)"
        elif d <= 0.20:
            return "Moderate (10-20%)"
        else:
            return "High (>20%)"

    df["Discount_Tier"] = df["Discount"].apply(get_discount_tier)

    # 5. Customer RFM Segmentation
    print("\n[4/6] Computing Customer RFM & Segmentation...")
    max_date = df["Order_Date"].max()
    cust_agg = df.groupby("Customer_ID").agg(
        Total_Revenue=("Revenue", "sum"),
        Total_Orders=("Order_ID", "nunique"),
        Total_Quantity=("Quantity", "sum"),
        First_Order=("Order_Date", "min"),
        Last_Order=("Order_Date", "max")
    ).reset_index()

    cust_agg["AOV"] = (cust_agg["Total_Revenue"] / cust_agg["Total_Orders"]).round(2)
    cust_agg["Recency_Days"] = (max_date - cust_agg["Last_Order"]).dt.days

    # Segmentation by spend quantiles
    q1 = cust_agg["Total_Revenue"].quantile(0.33)
    q2 = cust_agg["Total_Revenue"].quantile(0.66)

    def segment_cust(rev):
        if rev > q2:
            return "High Value"
        elif rev > q1:
            return "Medium Value"
        else:
            return "Low Value"

    cust_agg["Customer_Segment"] = cust_agg["Total_Revenue"].apply(segment_cust)

    # Merge customer segment back to main df
    df = df.merge(cust_agg[["Customer_ID", "Customer_Segment"]], on="Customer_ID", how="left")

    # 6. Build Star Schema Tables
    print("\n[5/6] Generating Star Schema Tables (Fact & Dimensions)...")

    # Dim_Customers
    dim_customers = df.groupby("Customer_ID").agg(
        Customer_Name=("Customer_Name", "first"),
        Gender=("Gender", "first"),
        Age=("Age", "first"),
        Customer_Segment=("Customer_Segment", "first")
    ).reset_index().merge(cust_agg[["Customer_ID", "Total_Revenue", "Total_Orders", "AOV", "Recency_Days"]], on="Customer_ID", how="left")
    dim_customers.rename(columns={"Total_Revenue": "Lifetime_Spend"}, inplace=True)

    # Dim_Products
    dim_products = df.groupby(["Product", "Category"]).agg(
        Unit_Price=("Unit_Price", "median"),
        Total_Units_Sold=("Quantity", "sum"),
        Total_Revenue=("Revenue", "sum")
    ).reset_index()

    # Dim_Locations
    dim_locations = df[["City", "State"]].drop_duplicates().reset_index(drop=True)
    dim_locations["Location_ID"] = dim_locations.index + 1
    dim_locations = dim_locations[["Location_ID", "City", "State"]]

    # Dim_Dates
    unique_dates = pd.date_range(start=df["Order_Date"].min(), end=df["Order_Date"].max(), freq="D")
    dim_dates = pd.DataFrame({"Date": unique_dates})
    dim_dates["Date_Key"] = dim_dates["Date"].dt.strftime("%Y%m%d").astype(int)
    dim_dates["Year"] = dim_dates["Date"].dt.year
    dim_dates["Quarter"] = "Q" + dim_dates["Date"].dt.quarter.astype(str)
    dim_dates["Month"] = dim_dates["Date"].dt.month
    dim_dates["Month_Name"] = dim_dates["Date"].dt.strftime("%B")
    dim_dates["Day"] = dim_dates["Date"].dt.day
    dim_dates["Day_Of_Week"] = dim_dates["Date"].dt.day_name()
    dim_dates["Is_Weekend"] = dim_dates["Date"].dt.dayofweek.isin([5, 6]).astype(int)

    # Fact_Sales
    fact_sales = df[[
        "Order_ID", "Order_Date", "Customer_ID", "Product", "Category",
        "City", "State", "Quantity", "Unit_Price", "Discount",
        "Gross_Sales", "Discount_Amount", "Revenue", "Order_Status",
        "Payment_Method", "Discount_Tier", "Customer_Segment"
    ]].copy()
    fact_sales["Date_Key"] = fact_sales["Order_Date"].dt.strftime("%Y%m%d").astype(int)

    # 7. Save Flat Cleaned Data and Star Schema CSVs
    cleaned_csv = os.path.join(processed_dir, "ecommerce_cleaned.csv")
    root_cleaned_csv = os.path.join(base_dir, "ecommerce_cleaned.csv")
    df.to_csv(cleaned_csv, index=False)
    df.to_csv(root_cleaned_csv, index=False)

    fact_sales.to_csv(os.path.join(processed_dir, "fact_sales.csv"), index=False)
    dim_customers.to_csv(os.path.join(processed_dir, "dim_customers.csv"), index=False)
    dim_products.to_csv(os.path.join(processed_dir, "dim_products.csv"), index=False)
    dim_locations.to_csv(os.path.join(processed_dir, "dim_locations.csv"), index=False)
    dim_dates.to_csv(os.path.join(processed_dir, "dim_dates.csv"), index=False)

    # Also put in powerbi_tableau directory for immediate drag-and-drop
    pbi_dir = os.path.join(base_dir, "powerbi_tableau")
    os.makedirs(pbi_dir, exist_ok=True)
    fact_sales.to_csv(os.path.join(pbi_dir, "fact_sales.csv"), index=False)
    dim_customers.to_csv(os.path.join(pbi_dir, "dim_customers.csv"), index=False)
    dim_products.to_csv(os.path.join(pbi_dir, "dim_products.csv"), index=False)
    dim_locations.to_csv(os.path.join(pbi_dir, "dim_locations.csv"), index=False)
    dim_dates.to_csv(os.path.join(pbi_dir, "dim_dates.csv"), index=False)

    # 8. Ingest into SQLite Database (ecommerce_sales.db)
    print("\n[6/6] Creating & Populating SQLite Database: ecommerce_sales.db...")
    db_path = os.path.join(base_dir, "ecommerce_sales.db")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Drop existing if needed
    cur.execute("DROP TABLE IF EXISTS fact_sales")
    cur.execute("DROP TABLE IF EXISTS dim_customers")
    cur.execute("DROP TABLE IF EXISTS dim_products")
    cur.execute("DROP TABLE IF EXISTS dim_locations")
    cur.execute("DROP TABLE IF EXISTS dim_dates")
    cur.execute("DROP TABLE IF EXISTS ecommerce_orders")

    # Ingest tables
    fact_sales.to_sql("fact_sales", conn, if_exists="replace", index=False)
    dim_customers.to_sql("dim_customers", conn, if_exists="replace", index=False)
    dim_products.to_sql("dim_products", conn, if_exists="replace", index=False)
    dim_locations.to_sql("dim_locations", conn, if_exists="replace", index=False)
    dim_dates.to_sql("dim_dates", conn, if_exists="replace", index=False)
    df.to_sql("ecommerce_orders", conn, if_exists="replace", index=False)

    # Create Indexes for fast analytical queries
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fact_date ON fact_sales(Order_Date)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fact_cust ON fact_sales(Customer_ID)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fact_prod ON fact_sales(Product)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fact_cat ON fact_sales(Category)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fact_status ON fact_sales(Order_Status)")

    conn.commit()
    conn.close()

    print(f"\nSUCCESS: ETL Completed Successfully!")
    print(f"Cleaned CSV: {cleaned_csv}")
    print(f"SQLite DB: {db_path}")
    print(f"Star Schema tables created in SQLite: fact_sales, dim_customers, dim_products, dim_locations, dim_dates")

    # Return key stats
    total_rev = df["Revenue"].sum()
    total_orders = df["Order_ID"].nunique()
    total_cust = df["Customer_ID"].nunique()
    aov = total_rev / total_orders
    print(f"\nKey Metrics:")
    print(f"  Total Revenue : ₹{total_rev:,.2f}")
    print(f"  Total Orders  : {total_orders:,}")
    print(f"  Total Custs   : {total_cust:,}")
    print(f"  Average Order : ₹{aov:,.2f}")
    return df

if __name__ == "__main__":
    run_etl()
