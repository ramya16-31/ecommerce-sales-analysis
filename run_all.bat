@echo off
title E-Commerce Analytics Pipeline Runner
echo ============================================================
echo   RUNNING FULL E-COMMERCE ANALYTICS PIPELINE
echo   Python + SQL + Excel + Charts
echo ============================================================
echo.

echo [1/4] Running Data Cleaning, Feature Engineering & SQLite ETL...
python data_processor.py
if %ERRORLEVEL% NEQ 0 (
    echo Error in data_processor.py!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/4] Executing 12 Business Analytics SQL Queries...
python run_sql_analysis.py
if %ERRORLEVEL% NEQ 0 (
    echo Error in run_sql_analysis.py!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/4] Generating 10 High-Resolution Python Charts...
python ecommerce_analysis.py
if %ERRORLEVEL% NEQ 0 (
    echo Error in ecommerce_analysis.py!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [4/4] Building Polished Executive Excel Dashboard...
python generate_excel_dashboard.py
if %ERRORLEVEL% NEQ 0 (
    echo Error in generate_excel_dashboard.py!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ============================================================
echo   ALL TASKS COMPLETED SUCCESSFULLY!
echo   - SQLite Database: ecommerce_sales.db
echo   - Charts: assets\charts\
echo   - Excel Dashboard: ecommerce_excel_dashboard.xlsx
echo   - Star Schema CSVs: powerbi_tableau\
echo ============================================================
echo.
pause
