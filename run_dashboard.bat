@echo off
title E-Commerce Sales BI Dashboard
echo ============================================================
echo   STARTING E-COMMERCE SALES BI DASHBOARD (STREAMLIT)
echo ============================================================
echo.
echo Launching local interactive dashboard...
python -m streamlit run app.py --browser.gatherUsageStats false
pause
