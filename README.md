# 🛒 E-Commerce Sales Analytics & Business Intelligence Project

An end-to-end, enterprise-grade Business Intelligence (BI) and Data Analytics platform built strictly with **Python, SQL, Excel, and Power BI / Tableau**.

---

## 🌟 Tech Stack & Components

| Pillar | Technology | Implementation & Deliverables |
| :--- | :--- | :--- |
| **Python** | Python 3.12, Pandas, Matplotlib, Altair, Streamlit | Automated ETL pipeline (`data_processor.py`), 10 publication-quality charts (`assets/charts/`), interactive live dashboard (`app.py`). |
| **SQL** | SQLite / ANSI SQL | Relational database (`ecommerce_sales.db`), Star Schema DDL (`sql/schema_star.sql`), 12 advanced business analytical queries (`sql/business_analysis_queries.sql`, `run_sql_analysis.py`). |
| **Excel** | OpenPyXL, Microsoft Excel | Formatted multi-tab Executive BI Dashboard (`ecommerce_excel_dashboard.xlsx`) with styled KPI cards, pivot-style tables, and monthly timelines. |
| **Power BI / Tableau** | Star Schema, DAX, Dimensions | Modeled fact and dimension tables (`Fact_Sales`, `Dim_Customers`, `Dim_Products`, `Dim_Locations`, `Dim_Dates`), DAX measures library (`powerbi_dax_measures.dax`), visual blueprint guide (`POWERBI_TABLEAU_GUIDE.md`). |

---

## 📁 Project Architecture & Directory Structure

```
Ecommerce project/
├── data/
│   ├── raw/
│   │   └── ecommerce_sales_dataset.xlsx        # Raw transactional dataset (1,005 rows)
│   └── processed/
│       ├── ecommerce_cleaned.csv               # Cleaned master dataset
│       ├── fact_sales.csv                      # Star schema Fact table
│       ├── dim_customers.csv                   # Customer RFM dimension
│       ├── dim_products.csv                    # Product catalog dimension
│       ├── dim_locations.csv                   # Geographic dimension (City, State)
│       └── dim_dates.csv                       # Calendar dimension
├── sql/
│   ├── schema_star.sql                         # Star schema DDL with foreign keys and indexes
│   └── business_analysis_queries.sql           # 12 Advanced SQL queries (CTEs, Window functions)
├── assets/
│   └── charts/                                 # 10 High-Resolution (300 DPI) charts
│       ├── 01_monthly_revenue_trend.png        # Dual-axis monthly revenue and order trajectory
│       ├── 02_category_performance.png         # Horizontal category revenue and share %
│       ├── 03_top_10_products.png              # Top revenue-generating products
│       ├── 04_city_sales_analysis.png          # Top 10 cities performance
│       ├── 05_payment_distribution.png         # Donut chart of payment channel share
│       ├── 06_order_status_breakdown.png       # Delivery success vs cancellation/returns
│       ├── 07_customer_rfm_segments.png        # Customer count vs spend contribution
│       ├── 08_discount_impact_scatter.png      # Discount rate vs order revenue scatter
│       ├── 09_day_of_week_velocity.png         # Weekly sales velocity
│       └── 10_executive_dashboard_infographic.png # Unified 4-panel executive summary
├── powerbi_tableau/
│   ├── fact_sales.csv                          # Pre-modeled Fact table
│   ├── dim_customers.csv                       # Customer dimension
│   ├── dim_products.csv                        # Product dimension
│   ├── dim_locations.csv                       # Location dimension
│   ├── dim_dates.csv                           # Calendar dimension
│   ├── powerbi_dax_measures.dax                # Copy-paste DAX measures (Total Sales, AOV, MoM Growth)
│   └── POWERBI_TABLEAU_GUIDE.md                # Step-by-step dashboard layout blueprint
├── ecommerce_sales.db                          # Relational SQLite database
├── ecommerce_excel_dashboard.xlsx              # Formatted Executive Excel Dashboard
├── data_processor.py                           # Python ETL pipeline & Star Schema builder
├── run_sql_analysis.py                         # SQL runner executing the 12 business queries
├── ecommerce_analysis.py                       # Python chart & statistical analysis generator
├── generate_excel_dashboard.py                 # Script generating the styled Excel dashboard
├── app.py                                      # Pure Python interactive Streamlit dashboard
├── run_dashboard.bat                           # 1-Click batch launcher for the web dashboard
├── run_all.bat                                 # 1-Click batch launcher for the entire pipeline
└── README.md                                   # Comprehensive documentation
```

---

## 📊 Core Business KPIs & Findings

| Business Metric | Value | Takeaway & Insight |
| :--- | :--- | :--- |
| **Total Net Revenue** | **₹18,032,395.00** | Net revenue realized after deducting ₹1.43M in discounts. |
| **Gross Sales** | **₹19,464,100.00** | Total order value before discounts. |
| **Total Discounts Given** | **₹1,431,705.00** | 7.36% effective average discount rate across all catalog items. |
| **Total Orders** | **1,000** | 1,000 unique orders placed across 18 active months. |
| **Active Customers** | **198** | Healthy customer base with an average of 5.05 orders per customer. |
| **Average Order Value (AOV)** | **₹18,032.40** | Driven primarily by premium electronics (Laptops, Smartphones, Tablets). |
| **Delivered Sales (Fulfillment)** | **₹13,778,020.00 (76.4%)** | 789 orders successfully delivered. |
| **Cancellation & Return Rate** | **6.3% Cancelled, 4.9% Returned** | Represents ₹2.35M in unrealized or lost revenue. |
| **Top Payment Channel** | **UPI (40.8% Share)** | UPI generated ₹7.35M across 393 orders, followed by Credit Cards (24.0%). |
| **Top State by Revenue** | **Tamil Nadu (₹4.92M | 27.3%)** | Top sales state, followed by Karnataka (₹3.48M) and Maharashtra (₹2.50M). |
| **High-Value Customer Tier** | **63.1% Revenue Share** | Top 33.8% of customers generate nearly two-thirds of all revenue. |

---

## 🚀 How to Run the Project

### 1. Launch the Interactive Python Dashboard (Streamlit)
Double-click **`run_dashboard.bat`** or run:
```bash
streamlit run app.py
```
This opens the interactive browser dashboard at `http://localhost:8501` with:
- Dynamic sidebar filters (Date Range, Categories, States, Order Statuses, Payment Methods, Segments).
- 5 tabbed deep-dive dashboards (Overview, Products, Geography, Customer RFM, Operations & Live SQL Console).
- Real-time CSV export of filtered records.

### 2. Run the SQL Business Analytics Engine
Run the SQL queries against `ecommerce_sales.db`:
```bash
python run_sql_analysis.py
```
This outputs formatted query results answering the 12 core business questions (CTEs, Window functions, MoM Growth, VIP Customers, Category ranking).

### 3. Re-generate High-Resolution Charts
```bash
python ecommerce_analysis.py
```
Saves 10 publication-quality 300 DPI charts into `assets/charts/`.

### 4. Build the Formatted Executive Excel Workbook
```bash
python generate_excel_dashboard.py
```
Generates `ecommerce_excel_dashboard.xlsx` complete with formatted KPI cards, pivot-style tables, and regional breakdowns.

### 5. Run the Entire End-to-End Pipeline
Double-click **`run_all.bat`** or execute:
```bash
python data_processor.py
python run_sql_analysis.py
python ecommerce_analysis.py
python generate_excel_dashboard.py
```

---

## 📐 Power BI & Tableau Integration

All files in `powerbi_tableau/` are formatted in a clean Star Schema:
1. Open **Power BI Desktop** or **Tableau Desktop**.
2. Connect to the CSV files in `powerbi_tableau/` (`fact_sales.csv`, `dim_customers.csv`, `dim_products.csv`, `dim_locations.csv`, `dim_dates.csv`).
3. Establish relationships as documented in [`POWERBI_TABLEAU_GUIDE.md`](file:///c:/Users/ramya/Desktop/Ecommerce%20project/powerbi_tableau/POWERBI_TABLEAU_GUIDE.md).
4. Copy and paste the DAX formulas from [`powerbi_dax_measures.dax`](file:///c:/Users/ramya/Desktop/Ecommerce%20project/powerbi_tableau/powerbi_dax_measures.dax).

---

## 📜 12 Analytical SQL Queries Included

1. **Executive KPI Overview**: Total orders, revenue, gross sales, discounts, units, and AOV.
2. **Month-over-Month (MoM) Growth**: Monthly sales trajectory and percentage growth using `LAG()`.
3. **Category Contribution**: Revenue share and average unit price per category.
4. **Top 3 Products per Category**: Ranked using `DENSE_RANK() OVER (PARTITION BY Category ORDER BY Revenue DESC)`.
5. **Order Fulfillment Health**: Delivery rate vs cancellation/return revenue impact.
6. **Payment Channel Performance**: Revenue, ticket size, and average discount per payment mode.
7. **Customer RFM Segmentation**: Revenue contribution across High, Medium, and Low value tiers.
8. **Top 10 VIP Spenders**: Customer name, lifetime spend, order count, and recency.
9. **Geographic State Ranking**: State sales volume, average order size, and national share %.
10. **Discount Tier Effectiveness**: Volume and revenue across 0%, <10%, and 10-20% discount bands.
11. **Weekly Sales Velocity**: Revenue distribution by day of the week.
12. **Repeat Purchase Behavior**: One-time buyers vs occasional vs loyal repeat buyers (5+ orders).
