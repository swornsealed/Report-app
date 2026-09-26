#!/bin/bash
cd "$(dirname "$0")"

echo ""
echo "  ╔══════════════════════════════════════╗"
echo "  ║    Radiometer ABL Report Generator  ║"
echo "  ║    Pathology Queensland              ║"
echo "  ╚══════════════════════════════════════╝"
echo ""

if ! command -v python3 &>/dev/null; then
  echo "  [!] Python 3 not found."
  echo "      Install via: https://python.org  or  brew install python"
  read -p "Press Enter to exit..."
  exit 1
fi

echo "  [1/3] Checking packages..."
# Only reach the internet on FIRST-TIME setup; normal runs are fully offline.
if python3 -c "import flask, docx, openpyxl, pandas, matplotlib" 2>/dev/null; then
  echo "        All packages present - no internet needed."
else
  echo "        First-time setup: installing packages (internet required once)..."
  pip3 install flask python-docx openpyxl pandas matplotlib --quiet 2>/dev/null
fi

echo "  [2/3] Starting app..."
echo "  [3/3] Opening browser..."
echo ""
echo "  ─────────────────────────────────────────"
echo "  App running at: http://localhost:5758"
echo "  Close this window to stop."
echo "  ─────────────────────────────────────────"
echo ""

python3 app.py
