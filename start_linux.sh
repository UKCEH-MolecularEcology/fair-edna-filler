#!/bin/bash
# FAIRe Metadata Filler – Linux launcher
# Run: bash start_linux.sh
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
    echo "Python 3.9+ not found. Install with:"
    echo "  sudo apt install python3   # Debian/Ubuntu"
    echo "  sudo dnf install python3   # Fedora/RHEL"
    echo ""
    exit 1
fi

echo ""
echo "  FAIRe Metadata Filler"
echo "  Installing / checking dependencies..."
"$PYTHON" -m pip install --quiet --upgrade pdfplumber anthropic 2>/dev/null || true

echo "  Starting server at http://127.0.0.1:8787"
echo "  Press Ctrl+C to stop."
echo ""

# Try to open browser
(sleep 2 && (xdg-open "http://127.0.0.1:8787" 2>/dev/null || true)) &

"$PYTHON" server.py
