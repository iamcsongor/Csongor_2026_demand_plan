#!/bin/bash
# Cambri Demand Plan — Launch Script (macOS / Linux)
# Run this once; the dashboard will stay live in your browser.

set -e
cd "$(dirname "$0")"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo " Cambri Demand Plan Dashboard"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Install dependencies silently if missing
python3 -c "import flask, requests, openpyxl" 2>/dev/null || {
  echo "→ Installing dependencies…"
  pip3 install flask requests openpyxl --quiet
}

echo "→ Starting server on http://localhost:5050"
echo "→ Opening dashboard in browser…"

# Open the dashboard file
sleep 1
open "demand_dashboard.html" 2>/dev/null || xdg-open "demand_dashboard.html" 2>/dev/null || true

# Start the server (keeps running until you Ctrl+C)
python3 server.py
