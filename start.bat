@echo off
REM Cambri Demand Plan — Launch Script (Windows)
REM Double-click to start; keep this window open.

echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo  Cambri Demand Plan Dashboard
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

cd /d "%~dp0"

REM Install dependencies if missing
python -c "import flask, requests, openpyxl" 2>nul || (
  echo ^> Installing dependencies...
  pip install flask requests openpyxl --quiet
)

echo ^> Starting server on http://localhost:5050
echo ^> Opening dashboard in browser...

timeout /t 1 /nobreak >nul
start "" "demand_dashboard.html"

python server.py
pause
