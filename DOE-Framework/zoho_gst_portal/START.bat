@echo off
title Zoho Billing to GST Portal
color 1F

echo.
echo  ============================================================
echo    Zoho Billing to GST Portal  -  Monthly Upload Tool
echo  ============================================================
echo.
echo  Starting the app... please wait.
echo.

:: Install required packages if not already installed
python -m pip install flask requests python-dotenv --quiet

echo  App is ready!
echo.
echo  *** Open your browser and go to:  http://localhost:5000 ***
echo.
echo  (Keep this window open while using the app. Close it to stop.)
echo.

cd /d "%~dp0"
python app.py

pause
