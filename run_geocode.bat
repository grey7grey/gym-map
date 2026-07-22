@echo off
echo ============================================
echo   Gym store coords baking tool
echo   Reads AMAP_KEY from your .env automatically.
echo   Parsing 106 stores, please wait ~30 seconds...
echo ============================================
"C:\Users\prote\.workbuddy\binaries\python\envs\default\Scripts\python.exe" "%~dp0geocode_once.py"
echo.
echo Done. Refresh the browser page to use built-in coords (no Key needed).
pause
