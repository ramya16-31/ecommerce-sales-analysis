"""
ecommerce_analysis.py
=====================
Comprehensive E-Commerce Data Analysis & Visual Chart Generator.
Uses Python, Pandas, Matplotlib, and OpenPyXL.
- Cleans data, calculates financial KPIs and customer segmentation.
- Generates 10 high-resolution analytical charts saved to assets/charts/.
- Exports multi-sheet Excel report: ecommerce_cleaned_analysis.xlsx.
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# Ensure UTF-8 output on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Configure clean matplotlib styles
plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")
plt.rcParams["font.sans-serif"] = "Arial, DejaVu Sans, Helvetica"
plt.rcParams["axes.edgecolor"] = "#E2E8F0"
plt.rcParams["axes.linewidth"] = 0.8
plt.rcParams["grid.color"] = "#F1F5F9"
plt.rcParams["grid.linestyle"] = "--"

# Palette
PRIMARY_COLOR = "#2563EB"   # Royal Blue
SECONDARY_COLOR = "#0D9488" # Teal
ACCENT_COLOR = "#F59E0B"    # Amber
DANGER_COLOR = "#EF4444"    # Red
PURPLE_COLOR = "#7C3AED"    # Purple
PALETTE = ["#2563EB", "#0D9488", "#F59E0B", "#7C3AED", "#EC4899", "#10B981", "#6366F1", "#F97316"]

def format_currency(val):
    if val >= 10000000:
        return f"₹{val/10000000:.2f} Cr"
    elif val >= 100000:
        return f"₹{val/100000:.2f} L"
    elif val >= 1000:
        return f"₹{val/1000:.1f} K"
    else:
        return f"₹{val:,.0f}"

def run_analysis():
    print("=" * 60)
    print("       E-COMMERCE DATA ANALYSIS & VISUALIZATION ENGINE")
    print("=" * 60)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    charts_dir = os.path.join(base_dir, "assets", "charts")
    os.makedirs(charts_dir, exist_ok=True)

    file_name = os.path.join(base_dir, "ecommerce_sales_dataset.xlsx")
    if not os.path.exists(file_name):
        print(f"\nERROR: Dataset file not found at {file_name}")
        return

    # 1. Load Data
    print("\n[1/5] Loading and inspecting dataset...")
    df = pd.read_excel(file_name)
    initial_rows = len(df)

    # 2. Data Cleaning
    print("\n[2/5] Cleaning data...")
    df = df.drop_duplicates()
    
    text_cols = ["Customer_Name", "City", "State", "Category", "Payment_Method", "Order_Status"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna("Unknown")

    median_qty = df["Quantity"].median()
    df["Quantity"] = df["Quantity"].fillna(median_qty)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"])

    # Calculations
    df["Gross_Sales"] = df["Quantity"] * df["Unit_Price"]
    df["Discount_Amount"] = (df["Gross_Sales"] * df["Discount"]).round(2)
    df["Revenue"] = (df["Gross_Sales"] - df["Discount_Amount"]).round(2)
    df["Month_Year"] = df["Order_Date"].dt.to_period("M").astype(str)
    df["Day_Of_Week"] = df["Order_Date"].dt.day_name()

    print(f"      Rows after deduplication: {len(df)} (Removed {initial_rows - len(df)})")

    # 3. KPIs
    total_revenue = df["Revenue"].sum()
    total_gross = df["Gross_Sales"].sum()
    total_discount = df["Discount_Amount"].sum()
    total_orders = df["Order_ID"].nunique()
    total_customers = df["Customer_ID"].nunique()
    total_quantity = df["Quantity"].sum()
    aov = total_revenue / total_orders

    delivered_df = df[df["Order_Status"] == "Delivered"]
    realized_revenue = delivered_df["Revenue"].sum()
    delivered_orders = delivered_df["Order_ID"].nunique()

    print("\n" + "=" * 45)
    print("               BUSINESS KPIs")
    print("=" * 45)
    print(f"Total Revenue        : ₹{total_revenue:,.2f}")
    print(f"Gross Sales          : ₹{total_gross:,.2f}")
    print(f"Total Discount       : ₹{total_discount:,.2f}")
    print(f"Total Orders         : {total_orders:,}")
    print(f"Total Customers      : {total_customers:,}")
    print(f"Total Units Sold     : {total_quantity:,.0f}")
    print(f"Average Order Value  : ₹{aov:,.2f}")
    print(f"Delivered Revenue    : ₹{realized_revenue:,.2f} ({realized_revenue/total_revenue*100:.1f}%)")
    print("=" * 45)

    # Aggregations
    category_revenue = df.groupby("Category")["Revenue"].sum().sort_values(ascending=False)
    product_revenue = df.groupby("Product")["Revenue"].sum().sort_values(ascending=False)
    city_revenue = df.groupby("City")["Revenue"].sum().sort_values(ascending=False)
    payment_analysis = df["Payment_Method"].value_counts()
    order_status = df["Order_Status"].value_counts()
    monthly_rev = df.groupby("Month_Year")["Revenue"].sum()
    monthly_orders = df.groupby("Month_Year")["Order_ID"].nunique()

    # Customer Segmentation
    cust_analysis = df.groupby("Customer_ID").agg(
        Total_Revenue=("Revenue", "sum"),
        Total_Orders=("Order_ID", "nunique"),
        Total_Quantity=("Quantity", "sum")
    )
    cust_analysis["Average_Order_Value"] = cust_analysis["Total_Revenue"] / cust_analysis["Total_Orders"]

    q1 = cust_analysis["Total_Revenue"].quantile(0.33)
    q2 = cust_analysis["Total_Revenue"].quantile(0.66)

    def segment_cust(rev):
        if rev <= q1:
            return "Low Value"
        elif rev <= q2:
            return "Medium Value"
        else:
            return "High Value"

    cust_analysis["Segment"] = cust_analysis["Total_Revenue"].apply(segment_cust)

    # 4. Generating Publication-Quality Charts
    print("\n[3/5] Generating 10 high-resolution analytical charts...")

    # Chart 1: Monthly Revenue Trend & Orders (Dual Axis)
    fig, ax1 = plt.subplots(figsize=(12, 6))
    months = monthly_rev.index.tolist()
    ax1.plot(months, monthly_rev.values / 100000, marker="o", color=PRIMARY_COLOR, linewidth=2.5, label="Revenue (₹ Lakhs)")
    ax1.set_ylabel("Revenue (₹ in Lakhs)", color=PRIMARY_COLOR, fontsize=12, fontweight="bold")
    ax1.tick_params(axis="y", labelcolor=PRIMARY_COLOR)
    ax1.set_xticks(range(len(months)))
    ax1.set_xticklabels(months, rotation=45, ha="right")

    ax2 = ax1.twinx()
    ax2.bar(months, monthly_orders.values, color=SECONDARY_COLOR, alpha=0.3, width=0.4, label="Order Count")
    ax2.set_ylabel("Total Orders Placed", color=SECONDARY_COLOR, fontsize=12, fontweight="bold")
    ax2.tick_params(axis="y", labelcolor=SECONDARY_COLOR)
    ax2.grid(False)

    plt.title("Monthly Revenue & Order Volume Trend (2025 - 2026)", fontsize=14, fontweight="bold", pad=15)
    fig.tight_layout()
    chart1_path = os.path.join(charts_dir, "01_monthly_revenue_trend.png")
    plt.savefig(chart1_path, dpi=300)
    plt.close()
    print("  -> Saved: 01_monthly_revenue_trend.png")

    # Chart 2: Category Performance
    plt.figure(figsize=(10, 5))
    bars = plt.barh(category_revenue.index[::-1], category_revenue.values[::-1] / 100000, color=PRIMARY_COLOR, edgecolor="#1D4ED8", height=0.6)
    plt.xlabel("Revenue (₹ in Lakhs)", fontsize=11, fontweight="bold")
    plt.title("Revenue Contribution by Category", fontsize=13, fontweight="bold", pad=15)
    for bar in bars:
        w = bar.get_width()
        pct = (w * 100000 / total_revenue) * 100
        plt.text(w + 0.3, bar.get_y() + bar.get_height()/2, f"₹{w:.1f}L ({pct:.1f}%)", va="center", fontsize=10, fontweight="bold", color="#1E293B")
    plt.xlim(0, max(category_revenue.values / 100000) * 1.2)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "02_category_performance.png"), dpi=300)
    plt.close()
    print("  -> Saved: 02_category_performance.png")

    # Chart 3: Top 10 Products by Revenue
    top10_prod = product_revenue.head(10)
    plt.figure(figsize=(11, 6))
    bars = plt.bar(top10_prod.index, top10_prod.values / 100000, color=SECONDARY_COLOR, edgecolor="#0F766E", width=0.6)
    plt.ylabel("Revenue (₹ in Lakhs)", fontsize=11, fontweight="bold")
    plt.title("Top 10 Revenue-Generating Products", fontsize=13, fontweight="bold", pad=15)
    plt.xticks(rotation=35, ha="right", fontsize=10)
    for bar in bars:
        h = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, h + 0.5, f"₹{h:.1f}L", ha="center", fontsize=9, fontweight="bold", color="#0F172A")
    plt.ylim(0, max(top10_prod.values / 100000) * 1.15)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "03_top_10_products.png"), dpi=300)
    plt.close()
    print("  -> Saved: 03_top_10_products.png")

    # Chart 4: City Sales Performance (Top 10)
    top10_cities = city_revenue.head(10)
    plt.figure(figsize=(10, 5))
    bars = plt.barh(top10_cities.index[::-1], top10_cities.values[::-1] / 100000, color=PURPLE_COLOR, edgecolor="#6D28D9", height=0.6)
    plt.xlabel("Revenue (₹ in Lakhs)", fontsize=11, fontweight="bold")
    plt.title("Top 10 Cities by Revenue", fontsize=13, fontweight="bold", pad=15)
    for bar in bars:
        w = bar.get_width()
        plt.text(w + 0.2, bar.get_y() + bar.get_height()/2, f"₹{w:.1f}L", va="center", fontsize=9, fontweight="bold")
    plt.xlim(0, max(top10_cities.values / 100000) * 1.18)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "04_city_sales_analysis.png"), dpi=300)
    plt.close()
    print("  -> Saved: 04_city_sales_analysis.png")

    # Chart 5: Payment Method Share (Donut Chart)
    plt.figure(figsize=(7, 7))
    wedges, texts, autotexts = plt.pie(
        payment_analysis.values,
        labels=payment_analysis.index,
        autopct="%1.1f%%",
        startangle=140,
        colors=PALETTE[:len(payment_analysis)],
        wedgeprops=dict(width=0.45, edgecolor="white", linewidth=2),
        textprops=dict(fontsize=11)
    )
    for at in autotexts:
        at.set_color("white")
        at.set_fontweight("bold")
    plt.title("Payment Channel Distribution (Transaction Volume)", fontsize=13, fontweight="bold", pad=20)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "05_payment_distribution.png"), dpi=300)
    plt.close()
    print("  -> Saved: 05_payment_distribution.png")

    # Chart 6: Order Fulfillment Breakdown
    plt.figure(figsize=(8, 5))
    status_colors = {"Delivered": "#10B981", "Shipped": "#3B82F6", "Cancelled": "#EF4444", "Returned": "#F59E0B"}
    cols = [status_colors.get(s, "#6B7280") for s in order_status.index]
    bars = plt.bar(order_status.index, order_status.values, color=cols, edgecolor="#334155", width=0.55)
    plt.ylabel("Number of Orders", fontsize=11, fontweight="bold")
    plt.title("Order Fulfillment Status Distribution", fontsize=13, fontweight="bold", pad=15)
    for bar in bars:
        h = bar.get_height()
        pct = h / total_orders * 100
        plt.text(bar.get_x() + bar.get_width()/2, h + 10, f"{h} ({pct:.1f}%)", ha="center", fontsize=10, fontweight="bold")
    plt.ylim(0, max(order_status.values) * 1.15)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "06_order_status_breakdown.png"), dpi=300)
    plt.close()
    print("  -> Saved: 06_order_status_breakdown.png")

    # Chart 7: Customer Segmentation (Count vs Spend)
    seg_stats = cust_analysis.groupby("Segment").agg(
        Cust_Count=("Total_Revenue", "count"),
        Total_Spend=("Total_Revenue", "sum")
    ).loc[["High Value", "Medium Value", "Low Value"]]

    fig, ax1 = plt.subplots(figsize=(9, 5))
    x = np.arange(len(seg_stats.index))
    w = 0.35
    ax1.bar(x - w/2, seg_stats["Cust_Count"], width=w, color=PRIMARY_COLOR, label="Customer Count")
    ax1.set_ylabel("Number of Customers", color=PRIMARY_COLOR, fontsize=11, fontweight="bold")
    ax1.tick_params(axis="y", labelcolor=PRIMARY_COLOR)
    ax1.set_xticks(x)
    ax1.set_xticklabels(seg_stats.index, fontsize=11, fontweight="bold")

    ax2 = ax1.twinx()
    ax2.bar(x + w/2, seg_stats["Total_Spend"] / 100000, width=w, color=ACCENT_COLOR, label="Total Spend (₹ Lakhs)")
    ax2.set_ylabel("Revenue (₹ in Lakhs)", color=ACCENT_COLOR, fontsize=11, fontweight="bold")
    ax2.tick_params(axis="y", labelcolor=ACCENT_COLOR)
    ax2.grid(False)

    plt.title("Customer Segmentation: Distribution vs Spend Contribution", fontsize=13, fontweight="bold", pad=15)
    fig.tight_layout()
    plt.savefig(os.path.join(charts_dir, "07_customer_rfm_segments.png"), dpi=300)
    plt.close()
    print("  -> Saved: 07_customer_rfm_segments.png")

    # Chart 8: Discount Strategy Impact
    df["Discount_Pct"] = (df["Discount"] * 100).astype(int)
    disc_summary = df.groupby("Discount_Pct").agg(
        Total_Orders=("Order_ID", "count"),
        Total_Rev=("Revenue", "sum")
    ).reset_index()

    plt.figure(figsize=(9, 5))
    plt.scatter(df["Discount"] * 100, df["Revenue"] / 1000, alpha=0.6, color=PRIMARY_COLOR, edgecolors="none", s=50)
    plt.xlabel("Discount Percentage (%)", fontsize=11, fontweight="bold")
    plt.ylabel("Order Revenue (₹ in Thousands)", fontsize=11, fontweight="bold")
    plt.title("Discount Rate vs Order Revenue Distribution", fontsize=13, fontweight="bold", pad=15)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "08_discount_impact_scatter.png"), dpi=300)
    plt.close()
    print("  -> Saved: 08_discount_impact_scatter.png")

    # Chart 9: Day of Week Velocity
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    dow_rev = df.groupby("Day_Of_Week")["Revenue"].sum().reindex(day_order)
    plt.figure(figsize=(9, 5))
    bars = plt.bar(dow_rev.index, dow_rev.values / 100000, color=SECONDARY_COLOR, edgecolor="#0D9488", width=0.55)
    plt.ylabel("Revenue (₹ in Lakhs)", fontsize=11, fontweight="bold")
    plt.title("Weekly Sales Velocity (Revenue by Day of Week)", fontsize=13, fontweight="bold", pad=15)
    plt.xticks(fontsize=10)
    for bar in bars:
        h = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, h + 0.3, f"₹{h:.1f}L", ha="center", fontsize=9, fontweight="bold")
    plt.ylim(0, max(dow_rev.values / 100000) * 1.15)
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "09_day_of_week_velocity.png"), dpi=300)
    plt.close()
    print("  -> Saved: 09_day_of_week_velocity.png")

    # Chart 10: Unified Executive Dashboard Infographic (4 Subplots)
    fig, axs = plt.subplots(2, 2, figsize=(16, 11))
    fig.suptitle(f"E-COMMERCE SALES EXECUTIVE SUMMARY (TOTAL REVENUE: ₹{total_revenue:,.2f})", fontsize=16, fontweight="bold", y=0.98)

    # 10.1 Category Breakdown
    axs[0, 0].bar(category_revenue.index, category_revenue.values / 100000, color=PRIMARY_COLOR, edgecolor="#1D4ED8")
    axs[0, 0].set_title("Revenue by Category (₹ Lakhs)", fontsize=12, fontweight="bold")
    axs[0, 0].tick_params(axis="x", rotation=25)

    # 10.2 Monthly Trend
    axs[0, 1].plot(monthly_rev.index, monthly_rev.values / 100000, marker="o", color=SECONDARY_COLOR, linewidth=2)
    axs[0, 1].set_title("Monthly Revenue Trajectory (₹ Lakhs)", fontsize=12, fontweight="bold")
    axs[0, 1].tick_params(axis="x", rotation=45)

    # 10.3 Payment Methods
    axs[1, 0].pie(payment_analysis.values, labels=payment_analysis.index, autopct="%1.0f%%", colors=PALETTE[:len(payment_analysis)], textprops=dict(fontsize=9))
    axs[1, 0].set_title("Payment Method Share", fontsize=12, fontweight="bold")

    # 10.4 Order Status
    axs[1, 1].bar(order_status.index, order_status.values, color=cols)
    axs[1, 1].set_title("Order Fulfillment Status", fontsize=12, fontweight="bold")

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(charts_dir, "10_executive_dashboard_infographic.png"), dpi=300)
    plt.close()
    print("  -> Saved: 10_executive_dashboard_infographic.png")

    # 5. Export multi-sheet Excel Report
    print("\n[4/5] Exporting multi-sheet Excel report...")
    output_excel = os.path.join(base_dir, "ecommerce_cleaned_analysis.xlsx")

    kpi_summary = pd.DataFrame({
        "Metric": [
            "Total Revenue", "Gross Sales", "Total Discount Given",
            "Total Orders", "Total Unique Customers", "Total Units Sold",
            "Average Order Value (AOV)", "Delivered Revenue", "Delivered Orders",
            "Cancellation Rate (%)", "Return Rate (%)"
        ],
        "Value": [
            total_revenue, total_gross, total_discount,
            total_orders, total_customers, total_quantity,
            aov, realized_revenue, delivered_orders,
            (order_status.get("Cancelled", 0) / total_orders) * 100,
            (order_status.get("Returned", 0) / total_orders) * 100
        ]
    })

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned_Transactions", index=False)
        kpi_summary.to_excel(writer, sheet_name="Executive_KPIs", index=False)
        category_revenue.reset_index().rename(columns={"Revenue": "Total_Revenue"}).to_excel(writer, sheet_name="Category_Analysis", index=False)
        product_revenue.reset_index().rename(columns={"Revenue": "Total_Revenue"}).to_excel(writer, sheet_name="Product_Performance", index=False)
        city_revenue.reset_index().rename(columns={"Revenue": "Total_Revenue"}).to_excel(writer, sheet_name="City_Performance", index=False)
        payment_analysis.reset_index().rename(columns={"count": "Order_Count"}).to_excel(writer, sheet_name="Payment_Methods", index=False)
        order_status.reset_index().rename(columns={"count": "Order_Count"}).to_excel(writer, sheet_name="Fulfillment_Status", index=False)
        monthly_rev.reset_index().rename(columns={"Revenue": "Monthly_Revenue"}).to_excel(writer, sheet_name="Monthly_Trends", index=False)
        cust_analysis.reset_index().to_excel(writer, sheet_name="Customer_RFM_Segments", index=False)

    print(f"      Excel report exported: {output_excel}")
    print("\n[5/5] Visual and statistical analysis completed successfully!")
    print("=" * 60)

if __name__ == "__main__":
    run_analysis()
