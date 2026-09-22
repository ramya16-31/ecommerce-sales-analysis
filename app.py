"""
app.py
======
Interactive E-Commerce Sales Business Intelligence Dashboard.
Built with 100% Pure Python & Streamlit.
Powered by SQLite (ecommerce_sales.db) & Cleaned Data.
"""

import os
import sys
import sqlite3
import pandas as pd
import numpy as np
import altair as alt
import streamlit as st

# Configure Streamlit page
st.set_page_config(
    page_title="E-Commerce Sales BI Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Modern BI Styling
st.markdown("""
<style>
    /* Metric Card Styling */
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        color: #1E3A8A;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #64748B;
        text-transform: uppercase;
    }
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        background-color: #F8FAFC;
        border-radius: 6px 6px 0 0;
        padding-top: 10px;
        padding-bottom: 10px;
        font-weight: 600;
        font-size: 0.95rem;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2563EB !important;
        color: white !important;
    }
    /* Section headers */
    h3 {
        color: #1E293B;
        padding-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# DATA LOADING & CACHING
# -------------------------------------------------------------
@st.cache_data
def load_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(base_dir, "data", "processed", "ecommerce_cleaned.csv")
    if not os.path.exists(csv_path):
        csv_path = os.path.join(base_dir, "ecommerce_cleaned.csv")
    
    if not os.path.exists(csv_path):
        # Auto-run ETL if not yet generated
        import data_processor
        data_processor.run_etl()
    
    df = pd.read_csv(csv_path)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"])
    return df

df_full = load_data()

# -------------------------------------------------------------
# SIDEBAR FILTERS
# -------------------------------------------------------------
st.sidebar.title("🔍 BI Filters")
st.sidebar.markdown("Filter all visualizations and KPIs dynamically:")

# 1. Date Range
min_date = df_full["Order_Date"].min().date()
max_date = df_full["Order_Date"].max().date()

date_range = st.sidebar.date_input(
    "Select Date Range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

# 2. Category Filter
categories = sorted(df_full["Category"].dropna().unique().tolist())
selected_categories = st.sidebar.multiselect(
    "Product Category",
    options=categories,
    default=categories
)

# 3. Order Status Filter
statuses = sorted(df_full["Order_Status"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect(
    "Order Status",
    options=statuses,
    default=statuses
)

# 4. State Filter
states = sorted(df_full["State"].dropna().unique().tolist())
selected_states = st.sidebar.multiselect(
    "State / Region",
    options=states,
    default=states
)

# 5. Payment Method Filter
payments = sorted(df_full["Payment_Method"].dropna().unique().tolist())
selected_payments = st.sidebar.multiselect(
    "Payment Method",
    options=payments,
    default=payments
)

# 6. Customer Segment
segments = ["High Value", "Medium Value", "Low Value"]
selected_segments = st.sidebar.multiselect(
    "Customer Segment",
    options=segments,
    default=segments
)

# Apply Filters
if len(date_range) == 2:
    start_d, end_d = date_range
    filtered_df = df_full[
        (df_full["Order_Date"].dt.date >= start_d) &
        (df_full["Order_Date"].dt.date <= end_d)
    ]
else:
    filtered_df = df_full.copy()

if selected_categories:
    filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]
if selected_statuses:
    filtered_df = filtered_df[filtered_df["Order_Status"].isin(selected_statuses)]
if selected_states:
    filtered_df = filtered_df[filtered_df["State"].isin(selected_states)]
if selected_payments:
    filtered_df = filtered_df[filtered_df["Payment_Method"].isin(selected_payments)]
if selected_segments:
    filtered_df = filtered_df[filtered_df["Customer_Segment"].isin(selected_segments)]

# -------------------------------------------------------------
# TOP HEADER & BANNER
# -------------------------------------------------------------
st.title("🛒 E-Commerce Sales Analytics & BI Dashboard")
st.caption(f"Showing **{len(filtered_df):,}** of **{len(df_full):,}** orders | Date Range: **{filtered_df['Order_Date'].min().strftime('%d %b %Y') if len(filtered_df)>0 else 'N/A'}** to **{filtered_df['Order_Date'].max().strftime('%d %b %Y') if len(filtered_df)>0 else 'N/A'}**")

if len(filtered_df) == 0:
    st.warning("⚠️ No orders match your selected filter criteria. Please adjust your sidebar filters.")
    st.stop()

# -------------------------------------------------------------
# EXECUTIVE KPI SUMMARY CARDS
# -------------------------------------------------------------
total_rev = filtered_df["Revenue"].sum()
gross_rev = filtered_df["Gross_Sales"].sum()
total_disc = filtered_df["Discount_Amount"].sum()
total_orders = filtered_df["Order_ID"].nunique()
total_custs = filtered_df["Customer_ID"].nunique()
units_sold = filtered_df["Quantity"].sum()
aov = total_rev / total_orders if total_orders > 0 else 0

delivered_df = filtered_df[filtered_df["Order_Status"] == "Delivered"]
deliv_rev = delivered_df["Revenue"].sum()
deliv_rate = (len(delivered_df) / len(filtered_df)) * 100 if len(filtered_df) > 0 else 0
cancel_count = len(filtered_df[filtered_df["Order_Status"] == "Cancelled"])
cancel_rate = (cancel_count / len(filtered_df)) * 100 if len(filtered_df) > 0 else 0

kpi_c1, kpi_c2, kpi_c3, kpi_c4, kpi_c5, kpi_c6 = st.columns(6)

kpi_c1.metric("Net Revenue", f"₹{total_rev:,.0f}", delta=f"-₹{total_disc:,.0f} Disc")
kpi_c2.metric("Total Orders", f"{total_orders:,}", delta=f"{units_sold:,} Units")
kpi_c3.metric("Avg Order Value", f"₹{aov:,.0f}")
kpi_c4.metric("Active Customers", f"{total_custs:,}")
kpi_c5.metric("Delivered Sales", f"₹{deliv_rev:,.0f}", delta=f"{deliv_rate:.1f}% Fulfilled")
kpi_c6.metric("Cancellation Rate", f"{cancel_rate:.1f}%", delta=f"{cancel_count} Orders", delta_color="inverse")

st.divider()

# -------------------------------------------------------------
# TABBED ANALYTICAL SECTIONS
# -------------------------------------------------------------
tab_overview, tab_products, tab_geo, tab_customers, tab_operations = st.tabs([
    "📈 Executive Overview & Trends",
    "🛍️ Products & Categories",
    "🗺️ Geographic Insights",
    "👥 Customer RFM Segments",
    "⚡ Operations & SQL Explorer"
])

# -------------------------------------------------------------
# TAB 1: EXECUTIVE OVERVIEW & TRENDS
# -------------------------------------------------------------
with tab_overview:
    col_t1, col_t2 = st.columns([7, 5])
    
    with col_t1:
        st.subheader("Monthly Revenue Trajectory")
        monthly_df = filtered_df.groupby("YearMonth").agg(
            Revenue=("Revenue", "sum"),
            Orders=("Order_ID", "nunique")
        ).reset_index()

        line_chart = alt.Chart(monthly_df).mark_line(point=True, color="#2563EB", strokeWidth=3).encode(
            x=alt.X("YearMonth:N", title="Year-Month", axis=alt.Axis(labelAngle=-45)),
            y=alt.Y("Revenue:Q", title="Net Revenue (₹)"),
            tooltip=["YearMonth", alt.Tooltip("Revenue:Q", format=",.2f"), alt.Tooltip("Orders:Q", format=",")]
        ).properties(height=350)
        st.altair_chart(line_chart, use_container_width=True)

    with col_t2:
        st.subheader("Revenue by Category")
        cat_df = filtered_df.groupby("Category")["Revenue"].sum().reset_index().sort_values("Revenue", ascending=False)
        bar_cat = alt.Chart(cat_df).mark_bar(color="#0D9488").encode(
            x=alt.X("Revenue:Q", title="Revenue (₹)"),
            y=alt.Y("Category:N", sort="-x", title=""),
            tooltip=["Category", alt.Tooltip("Revenue:Q", format=",.2f")]
        ).properties(height=350)
        st.altair_chart(bar_cat, use_container_width=True)

    st.subheader("Weekly Sales Velocity (Revenue by Day of Week)")
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    dow_df = filtered_df.groupby("Day_Of_Week").agg(
        Revenue=("Revenue", "sum"),
        Orders=("Order_ID", "nunique")
    ).reindex(day_order).reset_index()

    dow_chart = alt.Chart(dow_df).mark_bar(color="#6366F1").encode(
        x=alt.X("Day_Of_Week:N", sort=day_order, title="Day of Week"),
        y=alt.Y("Revenue:Q", title="Total Revenue (₹)"),
        tooltip=["Day_Of_Week", alt.Tooltip("Revenue:Q", format=",.2f"), "Orders"]
    ).properties(height=280)
    st.altair_chart(dow_chart, use_container_width=True)

# -------------------------------------------------------------
# TAB 2: PRODUCTS & CATEGORIES
# -------------------------------------------------------------
with tab_products:
    p_col1, p_col2 = st.columns([6, 6])

    with p_col1:
        st.subheader("Top 10 Best-Selling Products by Revenue")
        top_prod = filtered_df.groupby(["Product", "Category"]).agg(
            Revenue=("Revenue", "sum"),
            Units=("Quantity", "sum")
        ).reset_index().sort_values("Revenue", ascending=False).head(10)

        top_prod_chart = alt.Chart(top_prod).mark_bar(color="#F59E0B").encode(
            x=alt.X("Revenue:Q", title="Revenue (₹)"),
            y=alt.Y("Product:N", sort="-x", title=""),
            color=alt.Color("Category:N", scale=alt.Scale(scheme="category10")),
            tooltip=["Product", "Category", alt.Tooltip("Revenue:Q", format=",.2f"), "Units"]
        ).properties(height=380)
        st.altair_chart(top_prod_chart, use_container_width=True)

    with p_col2:
        st.subheader("Category Performance Breakdown")
        cat_perf = filtered_df.groupby("Category").agg(
            Total_Orders=("Order_ID", "nunique"),
            Units_Sold=("Quantity", "sum"),
            Gross_Sales=("Gross_Sales", "sum"),
            Total_Discount=("Discount_Amount", "sum"),
            Net_Revenue=("Revenue", "sum")
        ).reset_index()
        cat_perf["Revenue_Share_%"] = (cat_perf["Net_Revenue"] / total_rev * 100).round(1)
        cat_perf["Avg_Unit_Price"] = (cat_perf["Gross_Sales"] / cat_perf["Units_Sold"]).round(0)
        st.dataframe(cat_perf.sort_values("Net_Revenue", ascending=False), use_container_width=True, hide_index=True)

        st.subheader("Discount vs Net Revenue Distribution")
        scatter_chart = alt.Chart(filtered_df).mark_circle(size=60, opacity=0.6, color="#2563EB").encode(
            x=alt.X("Discount:Q", title="Discount Rate (%)", axis=alt.Axis(format="%")),
            y=alt.Y("Revenue:Q", title="Order Revenue (₹)"),
            color="Category:N",
            tooltip=["Order_ID", "Product", "Category", "Discount", alt.Tooltip("Revenue:Q", format=",.2f")]
        ).properties(height=260)
        st.altair_chart(scatter_chart, use_container_width=True)

# -------------------------------------------------------------
# TAB 3: GEOGRAPHIC INSIGHTS
# -------------------------------------------------------------
with tab_geo:
    g_col1, g_col2 = st.columns([6, 6])

    with g_col1:
        st.subheader("Top States by Revenue")
        state_df = filtered_df.groupby("State").agg(
            Revenue=("Revenue", "sum"),
            Orders=("Order_ID", "count")
        ).reset_index().sort_values("Revenue", ascending=False).head(10)

        state_chart = alt.Chart(state_df).mark_bar(color="#7C3AED").encode(
            x=alt.X("Revenue:Q", title="Revenue (₹)"),
            y=alt.Y("State:N", sort="-x", title=""),
            tooltip=["State", alt.Tooltip("Revenue:Q", format=",.2f"), "Orders"]
        ).properties(height=380)
        st.altair_chart(state_chart, use_container_width=True)

    with g_col2:
        st.subheader("Top 10 Cities by Revenue")
        city_df = filtered_df.groupby(["City", "State"]).agg(
            Revenue=("Revenue", "sum"),
            Orders=("Order_ID", "count")
        ).reset_index().sort_values("Revenue", ascending=False).head(10)

        city_chart = alt.Chart(city_df).mark_bar(color="#EC4899").encode(
            x=alt.X("Revenue:Q", title="Revenue (₹)"),
            y=alt.Y("City:N", sort="-x", title=""),
            tooltip=["City", "State", alt.Tooltip("Revenue:Q", format=",.2f"), "Orders"]
        ).properties(height=380)
        st.altair_chart(city_chart, use_container_width=True)

    st.subheader("Complete State & City Performance Table")
    geo_table = filtered_df.groupby(["State", "City"]).agg(
        Orders=("Order_ID", "count"),
        Units=("Quantity", "sum"),
        Revenue=("Revenue", "sum")
    ).reset_index().sort_values("Revenue", ascending=False)
    st.dataframe(geo_table, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# TAB 4: CUSTOMER RFM SEGMENTATION
# -------------------------------------------------------------
with tab_customers:
    c_col1, c_col2 = st.columns([5, 7])

    with c_col1:
        st.subheader("Customer Segment Revenue Contribution")
        seg_df = filtered_df.groupby("Customer_Segment").agg(
            Customers=("Customer_ID", "nunique"),
            Revenue=("Revenue", "sum")
        ).reset_index()

        seg_pie = alt.Chart(seg_df).mark_arc(innerRadius=50).encode(
            theta=alt.Theta(field="Revenue", type="quantitative"),
            color=alt.Color(field="Customer_Segment", type="nominal", scale=alt.Scale(
                domain=["High Value", "Medium Value", "Low Value"],
                range=["#2563EB", "#0D9488", "#94A3B8"]
            )),
            tooltip=["Customer_Segment", "Customers", alt.Tooltip("Revenue:Q", format=",.2f")]
        ).properties(height=320)
        st.altair_chart(seg_pie, use_container_width=True)

    with c_col2:
        st.subheader("Customer Value Tier Summary")
        seg_summary = filtered_df.groupby("Customer_Segment").agg(
            Customers=("Customer_ID", "nunique"),
            Orders=("Order_ID", "count"),
            Revenue=("Revenue", "sum")
        ).loc[["High Value", "Medium Value", "Low Value"]].reset_index()
        seg_summary["Revenue_Share_%"] = (seg_summary["Revenue"] / total_rev * 100).round(1)
        seg_summary["Avg_Spend_Per_Customer"] = (seg_summary["Revenue"] / seg_summary["Customers"]).round(2)
        st.dataframe(seg_summary, use_container_width=True, hide_index=True)

    st.subheader("Top 20 High-Value VIP Customers")
    vip_table = filtered_df.groupby("Customer_ID").agg(
        Name=("Customer_Name", "first"),
        Segment=("Customer_Segment", "first"),
        Orders=("Order_ID", "nunique"),
        Total_Spend=("Revenue", "sum"),
        City=("City", "first")
    ).reset_index().sort_values("Total_Spend", ascending=False).head(20)
    vip_table["AOV"] = (vip_table["Total_Spend"] / vip_table["Orders"]).round(2)
    st.dataframe(vip_table, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# TAB 5: OPERATIONS & SQL EXPLORER
# -------------------------------------------------------------
with tab_operations:
    op_col1, op_col2 = st.columns(2)

    with op_col1:
        st.subheader("Order Fulfillment Status")
        status_chart_df = filtered_df["Order_Status"].value_counts().reset_index()
        status_chart_df.columns = ["Status", "Orders"]
        status_bar = alt.Chart(status_chart_df).mark_bar().encode(
            x=alt.X("Orders:Q", title="Number of Orders"),
            y=alt.Y("Status:N", sort="-x", title=""),
            color=alt.Color("Status:N", scale=alt.Scale(
                domain=["Delivered", "Shipped", "Cancelled", "Returned"],
                range=["#10B981", "#3B82F6", "#EF4444", "#F59E0B"]
            )),
            tooltip=["Status", "Orders"]
        ).properties(height=260)
        st.altair_chart(status_bar, use_container_width=True)

    with op_col2:
        st.subheader("Payment Channel Distribution")
        pay_df = filtered_df["Payment_Method"].value_counts().reset_index()
        pay_df.columns = ["Payment_Method", "Count"]
        pay_donut = alt.Chart(pay_df).mark_arc(innerRadius=45).encode(
            theta="Count:Q",
            color=alt.Color("Payment_Method:N", scale=alt.Scale(scheme="tableau10")),
            tooltip=["Payment_Method", "Count"]
        ).properties(height=260)
        st.altair_chart(pay_donut, use_container_width=True)

    st.subheader("🔍 Live SQL Query Console (SQLite)")
    st.markdown("Execute SQL queries against `ecommerce_sales.db` in real time:")

    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ecommerce_sales.db")
    sample_sql = "SELECT Category, COUNT(*) AS Orders, ROUND(SUM(Revenue), 2) AS Total_Sales FROM fact_sales GROUP BY Category ORDER BY Total_Sales DESC;"
    user_sql = st.text_area("SQL Query", value=sample_sql, height=80)

    if st.button("Run SQL Query"):
        try:
            conn = sqlite3.connect(db_path)
            res_df = pd.read_sql_query(user_sql, conn)
            conn.close()
            st.success(f"Query returned {len(res_df)} rows:")
            st.dataframe(res_df, use_container_width=True)
        except Exception as e:
            st.error(f"SQL Error: {e}")

    st.subheader("Filtered Transaction Records")
    st.dataframe(filtered_df.head(100), use_container_width=True)
    
    # Download Button
    csv_data = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Filtered Data as CSV",
        data=csv_data,
        file_name="ecommerce_filtered_orders.csv",
        mime="text/csv"
    )
