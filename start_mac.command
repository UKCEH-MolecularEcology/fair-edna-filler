#!/bin/bash
# FAIRe Metadata Filler – macOS launcher
# Double-click this file to start the tool.
# If macOS asks "Are you sure you want to open this?" click Open.
set -e
cd "$(dirname "$0")"

PYTHON=""
for candidate in python3 python; do
    if command -v "$candidate" &>/dev/null; then
        VER=$("$candidate" -c "import sys; print(sys.version_info[:2] >= (3,9))" 2>/dev/null)
        if [ "$VER" = "True" ]; then
            PYTHON="$candidate"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo ""
    echo "============================================================"
    echo "  Python 3.9+ not found."
    echo ""
    echo "  Please install Python from https://www.python.org/downloads/"
    echo "  then double-click this file again."
    echo "============================================================"
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

echo ""
echo "  FAIRe Metadata Filler"
echo "  Installing / checking dependencies..."
echo ""
MISSING=$("$PYTHON" -c "import importlib.util as u; print(' '.join(p for p in ['pdfplumber','anthropic'] if u.find_spec(p) is None))")
if [ -n "$MISSING" ]; then
    "$PYTHON" -m pip install --quiet $MISSING 2>/dev/null \
    || "$PYTHON" -m pip install --quiet --user $MISSING 2>/dev/null \
    || "$PYTHON" -m pip install --quiet --user --break-system-packages $MISSING 2>/dev/null \
    || echo "  (Could not install $MISSING – PDF/AI features may be unavailable)"
fi

echo "  Starting server at http://127.0.0.1:8787"
echo "  Opening browser..."
echo ""
echo "  To stop the tool close this window (or press Ctrl+C)."
echo ""

# Open browser after a short delay
(sleep 2 && open "http://127.0.0.1:8787") &

"$PYTHON" server.py
