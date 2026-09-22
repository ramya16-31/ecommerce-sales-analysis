"""
generate_excel_dashboard.py
===========================
Generates a polished, executive-ready Excel BI Dashboard (ecommerce_excel_dashboard.xlsx)
using openpyxl. Includes styled KPI cards, formatted currency/percentages,
summary tables, and clean tabbed sheets.
"""

import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure UTF-8 output on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

def build_excel_dashboard():
    print("=" * 60)
    print("       GENERATING EXECUTIVE EXCEL BI DASHBOARD")
    print("=" * 60)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_csv = os.path.join(base_dir, "data", "processed", "ecommerce_cleaned.csv")
    if not os.path.exists(input_csv):
        input_csv = os.path.join(base_dir, "ecommerce_cleaned.csv")

    df = pd.read_csv(input_csv)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"])

    output_path = os.path.join(base_dir, "ecommerce_excel_dashboard.xlsx")
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    sub_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    card_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    kpi_card_border = Border(
        left=Side(style='thin', color='BFDBFE'),
        right=Side(style='thin', color='BFDBFE'),
        top=Side(style='medium', color='2563EB'),
        bottom=Side(style='thin', color='BFDBFE')
    )
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    white_title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    kpi_label_font = Font(name="Calibri", size=9, bold=False, color="4B5563")
    kpi_num_font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    table_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True, color="000000")
    regular_font = Font(name="Calibri", size=11, bold=False, color="1E293B")

    # -------------------------------------------------------------
    # SHEET 1: EXECUTIVE KPI SUMMARY
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="Executive_KPI_Summary")
    ws1.views.sheetView[0].showGridLines = True

    # Title Banner
    ws1.merge_cells("A1:K2")
    banner_cell = ws1["A1"]
    banner_cell.value = "  E-COMMERCE SALES EXECUTIVE DASHBOARD & BI OVERVIEW"
    banner_cell.font = white_title_font
    banner_cell.fill = navy_fill
    banner_cell.alignment = Alignment(vertical="center", horizontal="left")

    ws1["A3"].value = f"Data Period: {df['Order_Date'].min().strftime('%d %b %Y')} to {df['Order_Date'].max().strftime('%d %b %Y')}  |  Total Transactions: {len(df):,}"
    ws1["A3"].font = Font(name="Calibri", size=10, italic=True, color="64748B")

    # KPI Metrics
    total_rev = df["Revenue"].sum()
    gross_sales = df["Gross_Sales"].sum()
    total_disc = df["Discount_Amount"].sum()
    total_orders = df["Order_ID"].nunique()
    total_cust = df["Customer_ID"].nunique()
    aov = total_rev / total_orders
    delivered_rev = df[df["Order_Status"] == "Delivered"]["Revenue"].sum()
    cancel_rate = (len(df[df["Order_Status"] == "Cancelled"]) / len(df)) * 100

    kpis = [
        ("TOTAL NET REVENUE", f"₹{total_rev:,.2f}", "A5:B5", "A6:B6", "A5", "A6"),
        ("GROSS SALES", f"₹{gross_sales:,.2f}", "C5:D5", "C6:D6", "C5", "C6"),
        ("TOTAL DISCOUNTS GIVEN", f"₹{total_disc:,.2f}", "E5:F5", "E6:F6", "E5", "E6"),
        ("TOTAL ORDERS", f"{total_orders:,}", "G5:H5", "G6:H6", "G5", "G6"),
        ("AVERAGE ORDER VALUE", f"₹{aov:,.2f}", "I5:J5", "I6:J6", "I5", "I6"),
        ("DELIVERED SALES (76.4%)", f"₹{delivered_rev:,.2f}", "K5:L5", "K6:L6", "K5", "K6"),
    ]

    for label, val, m1, m2, c1, c2 in kpis:
        ws1.merge_cells(m1)
        ws1.merge_cells(m2)
        cell1 = ws1[c1]
        cell2 = ws1[c2]
        cell1.value = label
        cell1.font = kpi_label_font
        cell1.fill = card_fill
        cell1.alignment = Alignment(horizontal="center", vertical="center")

        cell2.value = val
        cell2.font = kpi_num_font
        cell2.fill = card_fill
        cell2.alignment = Alignment(horizontal="center", vertical="center")

    # Category Breakdown Table
    ws1["A9"].value = "REVENUE & MARGIN BY CATEGORY"
    ws1["A9"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

    cat_headers = ["Category", "Orders Placed", "Units Sold", "Gross Sales", "Discounts Given", "Net Revenue", "Revenue Share %"]
    for col_idx, h in enumerate(cat_headers, start=1):
        cell = ws1.cell(row=10, column=col_idx, value=h)
        cell.font = table_header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    cat_df = df.groupby("Category").agg(
        Orders=("Order_ID", "nunique"),
        Units=("Quantity", "sum"),
        Gross=("Gross_Sales", "sum"),
        Discount=("Discount_Amount", "sum"),
        Net_Rev=("Revenue", "sum")
    ).sort_values("Net_Rev", ascending=False).reset_index()

    for r_idx, row in cat_df.iterrows():
        row_num = 11 + r_idx
        ws1.cell(row=row_num, column=1, value=row["Category"]).font = bold_font
        ws1.cell(row=row_num, column=2, value=row["Orders"]).alignment = Alignment(horizontal="right")
        ws1.cell(row=row_num, column=3, value=row["Units"]).alignment = Alignment(horizontal="right")
        
        c4 = ws1.cell(row=row_num, column=4, value=row["Gross"])
        c4.number_format = '₹#,##0.00'
        
        c5 = ws1.cell(row=row_num, column=5, value=row["Discount"])
        c5.number_format = '₹#,##0.00'
        
        c6 = ws1.cell(row=row_num, column=6, value=row["Net_Rev"])
        c6.number_format = '₹#,##0.00'
        
        c7 = ws1.cell(row=row_num, column=7, value=row["Net_Rev"] / total_rev)
        c7.number_format = '0.0%'
        c7.alignment = Alignment(horizontal="right")

        for c in range(1, 8):
            ws1.cell(row=row_num, column=c).border = thin_border

    # Payment Methods Table
    ws1["I9"].value = "PAYMENT METHOD SHARE"
    ws1["I9"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

    pay_headers = ["Payment Mode", "Transactions", "Net Revenue", "Volume Share %"]
    for col_idx, h in enumerate(pay_headers, start=9):
        cell = ws1.cell(row=10, column=col_idx, value=h)
        cell.font = table_header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    pay_df = df.groupby("Payment_Method").agg(
        Txns=("Order_ID", "count"),
        Rev=("Revenue", "sum")
    ).sort_values("Rev", ascending=False).reset_index()

    for r_idx, row in pay_df.iterrows():
        row_num = 11 + r_idx
        ws1.cell(row=row_num, column=9, value=row["Payment_Method"]).font = regular_font
        ws1.cell(row=row_num, column=10, value=row["Txns"]).alignment = Alignment(horizontal="right")
        
        c11 = ws1.cell(row=row_num, column=11, value=row["Rev"])
        c11.number_format = '₹#,##0.00'
        
        c12 = ws1.cell(row=row_num, column=12, value=row["Txns"] / len(df))
        c12.number_format = '0.0%'
        c12.alignment = Alignment(horizontal="right")

        for c in range(9, 13):
            ws1.cell(row=row_num, column=c).border = thin_border

    # Fulfillment Status Table
    start_r = 18
    ws1.cell(row=start_r, column=1, value="ORDER FULFILLMENT & CANCELLATION ANALYSIS").font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
    status_headers = ["Order Status", "Order Count", "Total Revenue Value", "Status Share %"]
    for col_idx, h in enumerate(status_headers, start=1):
        cell = ws1.cell(row=start_r + 1, column=col_idx, value=h)
        cell.font = table_header_font
        cell.fill = header_fill
        cell.border = thin_border

    status_df = df.groupby("Order_Status").agg(
        Orders=("Order_ID", "count"),
        Rev=("Revenue", "sum")
    ).sort_values("Rev", ascending=False).reset_index()

    for r_idx, row in status_df.iterrows():
        r_n = start_r + 2 + r_idx
        ws1.cell(row=r_n, column=1, value=row["Order_Status"]).font = regular_font
        ws1.cell(row=r_n, column=2, value=row["Orders"]).alignment = Alignment(horizontal="right")
        
        c3 = ws1.cell(row=r_n, column=3, value=row["Rev"])
        c3.number_format = '₹#,##0.00'
        
        c4 = ws1.cell(row=r_n, column=4, value=row["Orders"] / len(df))
        c4.number_format = '0.0%'
        c4.alignment = Alignment(horizontal="right")

        for c in range(1, 5):
            ws1.cell(row=r_n, column=c).border = thin_border

    # -------------------------------------------------------------
    # SHEET 2: MONTHLY TRENDS
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Monthly_Sales_Trend")
    ws2.views.sheetView[0].showGridLines = True
    ws2["A1"].value = "18-MONTH SALES TRAJECTORY & MOM GROWTH"
    ws2["A1"].font = white_title_font
    ws2["A1"].fill = navy_fill
    ws2.merge_cells("A1:G1")

    m_headers = ["Year-Month", "Total Orders", "Units Sold", "Gross Sales", "Discounts", "Net Revenue", "MoM Growth %"]
    for col_idx, h in enumerate(m_headers, start=1):
        c = ws2.cell(row=3, column=col_idx, value=h)
        c.font = table_header_font
        c.fill = header_fill
        c.border = thin_border

    monthly = df.groupby("YearMonth").agg(
        Orders=("Order_ID", "nunique"),
        Units=("Quantity", "sum"),
        Gross=("Gross_Sales", "sum"),
        Disc=("Discount_Amount", "sum"),
        Rev=("Revenue", "sum")
    ).reset_index()

    prev_rev = None
    for r_idx, row in monthly.iterrows():
        r_n = 4 + r_idx
        ws2.cell(row=r_n, column=1, value=row["YearMonth"]).font = bold_font
        ws2.cell(row=r_n, column=2, value=row["Orders"]).alignment = Alignment(horizontal="right")
        ws2.cell(row=r_n, column=3, value=row["Units"]).alignment = Alignment(horizontal="right")
        
        c4 = ws2.cell(row=r_n, column=4, value=row["Gross"])
        c4.number_format = '₹#,##0.00'
        
        c5 = ws2.cell(row=r_n, column=5, value=row["Disc"])
        c5.number_format = '₹#,##0.00'
        
        c6 = ws2.cell(row=r_n, column=6, value=row["Rev"])
        c6.number_format = '₹#,##0.00'

        growth = (row["Rev"] - prev_rev) / prev_rev if prev_rev else 0.0
        c7 = ws2.cell(row=r_n, column=7, value=growth)
        c7.number_format = '0.0%'
        c7.alignment = Alignment(horizontal="right")
        prev_rev = row["Rev"]

        for c in range(1, 8):
            ws2.cell(row=r_n, column=c).border = thin_border

    # -------------------------------------------------------------
    # SHEET 3: PRODUCT PERFORMANCE
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="Product_Performance")
    ws3.views.sheetView[0].showGridLines = True
    ws3["A1"].value = "PRODUCT CATALOG PERFORMANCE & RANKING"
    ws3["A1"].font = white_title_font
    ws3["A1"].fill = navy_fill
    ws3.merge_cells("A1:G1")

    p_headers = ["Product Name", "Category", "Median Unit Price", "Units Sold", "Gross Sales", "Net Revenue", "Catalog Share %"]
    for col_idx, h in enumerate(p_headers, start=1):
        c = ws3.cell(row=3, column=col_idx, value=h)
        c.font = table_header_font
        c.fill = header_fill
        c.border = thin_border

    prod_df = df.groupby(["Product", "Category"]).agg(
        Price=("Unit_Price", "median"),
        Units=("Quantity", "sum"),
        Gross=("Gross_Sales", "sum"),
        Rev=("Revenue", "sum")
    ).sort_values("Rev", ascending=False).reset_index()

    for r_idx, row in prod_df.iterrows():
        r_n = 4 + r_idx
        ws3.cell(row=r_n, column=1, value=row["Product"]).font = bold_font
        ws3.cell(row=r_n, column=2, value=row["Category"])
        
        c3 = ws3.cell(row=r_n, column=3, value=row["Price"])
        c3.number_format = '₹#,##0.00'
        
        ws3.cell(row=r_n, column=4, value=row["Units"]).alignment = Alignment(horizontal="right")
        
        c5 = ws3.cell(row=r_n, column=5, value=row["Gross"])
        c5.number_format = '₹#,##0.00'
        
        c6 = ws3.cell(row=r_n, column=6, value=row["Rev"])
        c6.number_format = '₹#,##0.00'
        
        c7 = ws3.cell(row=r_n, column=7, value=row["Rev"] / total_rev)
        c7.number_format = '0.0%'

        for c in range(1, 8):
            ws3.cell(row=r_n, column=c).border = thin_border

    # -------------------------------------------------------------
    # SHEET 4: CUSTOMER SEGMENTATION & VIPS
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="Customer_Segmentation")
    ws4.views.sheetView[0].showGridLines = True
    ws4["A1"].value = "CUSTOMER VALUE SEGMENTATION & TOP 25 VIPS"
    ws4["A1"].font = white_title_font
    ws4["A1"].fill = navy_fill
    ws4.merge_cells("A1:G1")

    # Segment Summary Table
    ws4["A3"].value = "RFM VALUE TIERS"
    ws4["A3"].font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    s_headers = ["Segment Tier", "Customer Count", "Customer %", "Total Orders", "Segment Revenue", "Revenue %", "Avg Customer Spend"]
    for col_idx, h in enumerate(s_headers, start=1):
        c = ws4.cell(row=4, column=col_idx, value=h)
        c.font = table_header_font
        c.fill = header_fill
        c.border = thin_border

    seg_summary = df.groupby("Customer_Segment").agg(
        Custs=("Customer_ID", "nunique"),
        Orders=("Order_ID", "count"),
        Rev=("Revenue", "sum")
    ).loc[["High Value", "Medium Value", "Low Value"]].reset_index()

    for r_idx, row in seg_summary.iterrows():
        r_n = 5 + r_idx
        ws4.cell(row=r_n, column=1, value=row["Customer_Segment"]).font = bold_font
        ws4.cell(row=r_n, column=2, value=row["Custs"]).alignment = Alignment(horizontal="right")
        
        c3 = ws4.cell(row=r_n, column=3, value=row["Custs"] / total_cust)
        c3.number_format = '0.0%'
        
        ws4.cell(row=r_n, column=4, value=row["Orders"]).alignment = Alignment(horizontal="right")
        
        c5 = ws4.cell(row=r_n, column=5, value=row["Rev"])
        c5.number_format = '₹#,##0.00'
        
        c6 = ws4.cell(row=r_n, column=6, value=row["Rev"] / total_rev)
        c6.number_format = '0.0%'
        
        c7 = ws4.cell(row=r_n, column=7, value=row["Rev"] / row["Custs"])
        c7.number_format = '₹#,##0.00'

        for c in range(1, 8):
            ws4.cell(row=r_n, column=c).border = thin_border

    # Top 25 VIPs Table
    ws4["A10"].value = "TOP 25 HIGH-VALUE CUSTOMERS (VIPS)"
    ws4["A10"].font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    vip_headers = ["Customer ID", "Customer Name", "Segment", "Total Orders", "Lifetime Spend", "AOV", "City"]
    for col_idx, h in enumerate(vip_headers, start=1):
        c = ws4.cell(row=11, column=col_idx, value=h)
        c.font = table_header_font
        c.fill = header_fill
        c.border = thin_border

    vip_df = df.groupby("Customer_ID").agg(
        Name=("Customer_Name", "first"),
        Segment=("Customer_Segment", "first"),
        Orders=("Order_ID", "nunique"),
        Spend=("Revenue", "sum"),
        City=("City", "first")
    ).sort_values("Spend", ascending=False).head(25).reset_index()

    for r_idx, row in vip_df.iterrows():
        r_n = 12 + r_idx
        ws4.cell(row=r_n, column=1, value=row["Customer_ID"]).font = bold_font
        ws4.cell(row=r_n, column=2, value=row["Name"])
        ws4.cell(row=r_n, column=3, value=row["Segment"])
        ws4.cell(row=r_n, column=4, value=row["Orders"]).alignment = Alignment(horizontal="right")
        
        c5 = ws4.cell(row=r_n, column=5, value=row["Spend"])
        c5.number_format = '₹#,##0.00'
        
        c6 = ws4.cell(row=r_n, column=6, value=row["Spend"] / row["Orders"])
        c6.number_format = '₹#,##0.00'
        
        ws4.cell(row=r_n, column=7, value=row["City"])

        for c in range(1, 8):
            ws4.cell(row=r_n, column=c).border = thin_border

    # -------------------------------------------------------------
    # SHEET 5: REGIONAL ANALYSIS
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="Regional_Analysis")
    ws5.views.sheetView[0].showGridLines = True
    ws5["A1"].value = "STATE & CITY GEOGRAPHIC SALES BREAKDOWN"
    ws5["A1"].font = white_title_font
    ws5["A1"].fill = navy_fill
    ws5.merge_cells("A1:F1")

    reg_headers = ["State", "Active Cities", "Order Count", "Units Sold", "Total Revenue", "National Share %"]
    for col_idx, h in enumerate(reg_headers, start=1):
        c = ws5.cell(row=3, column=col_idx, value=h)
        c.font = table_header_font
        c.fill = header_fill
        c.border = thin_border

    state_df = df.groupby("State").agg(
        Cities=("City", "nunique"),
        Orders=("Order_ID", "count"),
        Units=("Quantity", "sum"),
        Rev=("Revenue", "sum")
    ).sort_values("Rev", ascending=False).reset_index()

    for r_idx, row in state_df.iterrows():
        r_n = 4 + r_idx
        ws5.cell(row=r_n, column=1, value=row["State"]).font = bold_font
        ws5.cell(row=r_n, column=2, value=row["Cities"]).alignment = Alignment(horizontal="right")
        ws5.cell(row=r_n, column=3, value=row["Orders"]).alignment = Alignment(horizontal="right")
        ws5.cell(row=r_n, column=4, value=row["Units"]).alignment = Alignment(horizontal="right")
        
        c5 = ws5.cell(row=r_n, column=5, value=row["Rev"])
        c5.number_format = '₹#,##0.00'
        
        c6 = ws5.cell(row=r_n, column=6, value=row["Rev"] / total_rev)
        c6.number_format = '0.0%'

        for c in range(1, 7):
            ws5.cell(row=r_n, column=c).border = thin_border

    # Auto-fit column widths for all sheets
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len and not cell.coordinate in sheet.merged_cells:
                    max_len = len(val_str)
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_path)
    print(f"SUCCESS: Executive Excel Dashboard created at:\n  -> {output_path}")

if __name__ == "__main__":
    build_excel_dashboard()
