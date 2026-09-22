"""
run_sql_analysis.py
===================
Executes the analytical SQL queries in sql/business_analysis_queries.sql
against the SQLite database (ecommerce_sales.db) and prints clean,
formatted reports with business insights.
"""

import os
import sys
import sqlite3
import pandas as pd

# Windows UTF-8 console output setup
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

def run_queries():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(base_dir, "ecommerce_sales.db")
    sql_file = os.path.join(base_dir, "sql", "business_analysis_queries.sql")

    if not os.path.exists(db_path):
        print("Database not found! Running data_processor.py first...")
        import data_processor
        data_processor.run_etl()

    conn = sqlite3.connect(db_path)

    queries = {
        "1. Executive KPI Summary": """
            SELECT 
                COUNT(DISTINCT Order_ID) AS Total_Orders,
                COUNT(DISTINCT Customer_ID) AS Total_Customers,
                SUM(Quantity) AS Total_Units_Sold,
                ROUND(SUM(Gross_Sales), 2) AS Gross_Sales,
                ROUND(SUM(Discount_Amount), 2) AS Total_Discounts,
                ROUND(SUM(Revenue), 2) AS Net_Revenue,
                ROUND(SUM(Revenue) / COUNT(DISTINCT Order_ID), 2) AS Average_Order_Value,
                ROUND(SUM(Discount_Amount) * 100.0 / SUM(Gross_Sales), 2) AS Effective_Discount_Rate_Pct
            FROM fact_sales;
        """,
        "2. Month-over-Month (MoM) Revenue Growth": """
            WITH Monthly_Sales AS (
                SELECT 
                    strftime('%Y-%m', Order_Date) AS Sales_Month,
                    COUNT(DISTINCT Order_ID) AS Order_Count,
                    SUM(Quantity) AS Units_Sold,
                    ROUND(SUM(Revenue), 2) AS Monthly_Revenue
                FROM fact_sales
                GROUP BY strftime('%Y-%m', Order_Date)
            )
            SELECT 
                Sales_Month,
                Order_Count,
                Units_Sold,
                Monthly_Revenue,
                LAG(Monthly_Revenue, 1) OVER (ORDER BY Sales_Month) AS Prev_Month_Revenue,
                ROUND(
                    (Monthly_Revenue - LAG(Monthly_Revenue, 1) OVER (ORDER BY Sales_Month)) * 100.0 / 
                    LAG(Monthly_Revenue, 1) OVER (ORDER BY Sales_Month), 
                    2
                ) AS MoM_Growth_Pct
            FROM Monthly_Sales
            ORDER BY Sales_Month;
        """,
        "3. Category Performance & Revenue Contribution": """
            SELECT 
                Category,
                COUNT(DISTINCT Order_ID) AS Total_Orders,
                SUM(Quantity) AS Units_Sold,
                ROUND(SUM(Revenue), 2) AS Total_Revenue,
                ROUND(AVG(Unit_Price), 2) AS Avg_Unit_Price,
                ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM fact_sales), 2) AS Revenue_Share_Pct
            FROM fact_sales
            GROUP BY Category
            ORDER BY Total_Revenue DESC;
        """,
        "4. Top 3 Best-Selling Products per Category": """
            WITH Ranked_Products AS (
                SELECT 
                    Category,
                    Product,
                    SUM(Quantity) AS Units_Sold,
                    ROUND(SUM(Revenue), 2) AS Product_Revenue,
                    DENSE_RANK() OVER (PARTITION BY Category ORDER BY SUM(Revenue) DESC) as Category_Rank
                FROM fact_sales
                GROUP BY Category, Product
            )
            SELECT 
                Category,
                Category_Rank,
                Product,
                Units_Sold,
                Product_Revenue
            FROM Ranked_Products
            WHERE Category_Rank <= 3
            ORDER BY Category, Category_Rank;
        """,
        "5. Order Fulfillment & Lost Revenue Analysis": """
            SELECT 
                Order_Status,
                COUNT(Order_ID) AS Order_Count,
                ROUND(COUNT(Order_ID) * 100.0 / (SELECT COUNT(*) FROM fact_sales), 2) AS Status_Share_Pct,
                ROUND(SUM(Revenue), 2) AS Total_Value,
                ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM fact_sales), 2) AS Value_Share_Pct
            FROM fact_sales
            GROUP BY Order_Status
            ORDER BY Total_Value DESC;
        """,
        "6. Payment Channel Share & Average Ticket Size": """
            SELECT 
                Payment_Method,
                COUNT(Order_ID) AS Transaction_Count,
                ROUND(SUM(Revenue), 2) AS Channel_Revenue,
                ROUND(AVG(Revenue), 2) AS Avg_Ticket_Size,
                ROUND(AVG(Discount) * 100, 2) AS Avg_Discount_Pct,
                ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM fact_sales), 2) AS Revenue_Share_Pct
            FROM fact_sales
            GROUP BY Payment_Method
            ORDER BY Channel_Revenue DESC;
        """,
        "7. Customer Segmentation (RFM Value Tiers)": """
            SELECT 
                Customer_Segment,
                COUNT(DISTINCT Customer_ID) AS Customer_Count,
                ROUND(COUNT(DISTINCT Customer_ID) * 100.0 / (SELECT COUNT(DISTINCT Customer_ID) FROM fact_sales), 2) AS Customer_Share_Pct,
                COUNT(Order_ID) AS Total_Orders,
                ROUND(SUM(Revenue), 2) AS Segment_Revenue,
                ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM fact_sales), 2) AS Revenue_Share_Pct,
                ROUND(SUM(Revenue) / COUNT(DISTINCT Customer_ID), 2) AS Avg_Spend_Per_Customer
            FROM fact_sales
            GROUP BY Customer_Segment
            ORDER BY Segment_Revenue DESC;
        """,
        "8. Top 10 VIP Customers by Lifetime Spend": """
            SELECT 
                c.Customer_ID,
                c.Customer_Name,
                c.Customer_Segment,
                c.Total_Orders,
                ROUND(c.Lifetime_Spend, 2) AS Lifetime_Spend,
                ROUND(c.AOV, 2) AS Average_Order_Value,
                c.Recency_Days AS Days_Since_Last_Order
            FROM dim_customers c
            ORDER BY c.Lifetime_Spend DESC
            LIMIT 10;
        """,
        "9. Top 10 States by Revenue": """
            SELECT 
                State,
                COUNT(DISTINCT City) AS Active_Cities,
                COUNT(Order_ID) AS Total_Orders,
                ROUND(SUM(Revenue), 2) AS State_Revenue,
                ROUND(SUM(Revenue) / COUNT(Order_ID), 2) AS Avg_Order_Value,
                ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM fact_sales), 2) AS National_Share_Pct
            FROM fact_sales
            GROUP BY State
            ORDER BY State_Revenue DESC
            LIMIT 10;
        """,
        "10. Discount Strategy Effectiveness": """
            SELECT 
                Discount_Tier,
                COUNT(Order_ID) AS Order_Volume,
                SUM(Quantity) AS Units_Sold,
                ROUND(SUM(Gross_Sales), 2) AS Gross_Sales,
                ROUND(SUM(Discount_Amount), 2) AS Total_Discount_Given,
                ROUND(SUM(Revenue), 2) AS Net_Revenue,
                ROUND(AVG(Revenue), 2) AS Avg_Order_Revenue
            FROM fact_sales
            GROUP BY Discount_Tier
            ORDER BY Net_Revenue DESC;
        """,
        "11. Day of Week Sales Velocity": """
            SELECT 
                CASE CAST(strftime('%w', Order_Date) AS INTEGER)
                    WHEN 0 THEN 'Sunday'
                    WHEN 1 THEN 'Monday'
                    WHEN 2 THEN 'Tuesday'
                    WHEN 3 THEN 'Wednesday'
                    WHEN 4 THEN 'Thursday'
                    WHEN 5 THEN 'Friday'
                    WHEN 6 THEN 'Saturday'
                END AS Day_Of_Week,
                COUNT(Order_ID) AS Total_Orders,
                SUM(Quantity) AS Total_Units,
                ROUND(SUM(Revenue), 2) AS Total_Revenue,
                ROUND(AVG(Revenue), 2) AS Avg_Order_Value
            FROM fact_sales
            GROUP BY strftime('%w', Order_Date)
            ORDER BY Total_Revenue DESC;
        """,
        "12. Customer Repeat Purchase Behavior": """
            WITH Customer_Order_Frequency AS (
                SELECT 
                    Customer_ID,
                    COUNT(Order_ID) AS Order_Count,
                    SUM(Revenue) AS Customer_Spend
                FROM fact_sales
                GROUP BY Customer_ID
            )
            SELECT 
                CASE 
                    WHEN Order_Count = 1 THEN '1 Order (One-time Buyer)'
                    WHEN Order_Count BETWEEN 2 AND 4 THEN '2-4 Orders (Occasional Buyer)'
                    ELSE '5+ Orders (Loyal / Repeat Buyer)'
                END AS Buyer_Type,
                COUNT(Customer_ID) AS Customer_Count,
                SUM(Order_Count) AS Total_Orders_Placed,
                ROUND(SUM(Customer_Spend), 2) AS Total_Revenue_Generated,
                ROUND(SUM(Customer_Spend) * 100.0 / (SELECT SUM(Revenue) FROM fact_sales), 2) AS Revenue_Share_Pct
            FROM Customer_Order_Frequency
            GROUP BY Buyer_Type
            ORDER BY Total_Revenue_Generated DESC;
        """
    }

    print("=" * 80)
    print("       E-COMMERCE SQL BUSINESS INTELLIGENCE ANALYSIS")
    print("=" * 80)

    for title, query in queries.items():
        print(f"\n>>> {title}")
        print("-" * 80)
        df_result = pd.read_sql_query(query, conn)
        print(df_result.to_string(index=False))
        print("-" * 80)

    conn.close()
    print("\nAll 12 SQL queries executed successfully against ecommerce_sales.db!")

if __name__ == "__main__":
    run_queries()
