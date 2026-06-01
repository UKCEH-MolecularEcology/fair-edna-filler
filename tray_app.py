"""
FAIRe Metadata Filler — system tray launcher.

Runs the web server in a background thread, shows a tray icon, and opens
the browser on click.  No terminal window required.

Run once:
    pip install pystray pillow pdfplumber anthropic

Then just double-click this file (or the platform launcher).
"""

from __future__ import annotations
import os
import sys
import threading
import webbrowser
from pathlib import Path

# ── bundled-app path fix ──────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).resolve().parent

os.chdir(APP_DIR)
sys.path.insert(0, str(APP_DIR))

PORT   = int(os.environ.get('FAIRE_TOOL_PORT', '8787'))
URL    = f'http://127.0.0.1:{PORT}'
TITLE  = 'FAIRe Metadata Filler'


# ── build tray icon from UKCEH Air colour ────────────────────────────────────
def _make_icon():
    try:
        from PIL import Image, ImageDraw, ImageFont
        size   = 64
        img    = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw   = ImageDraw.Draw(img)
        # Circle in UKCEH Water blue
        draw.ellipse([4, 4, size-4, size-4], fill='#4679E1')
        # White 'F' letter
        try:
            font = ImageFont.truetype('/System/Library/Fonts/Helvetica.ttc', 36)
        except Exception:
            font = ImageFont.load_default()
        draw.text((size//2, size//2), 'F', fill='white', font=font, anchor='mm')
        return img
    except Exception:
        # Minimal 16×16 fallback
        from PIL import Image
        img = Image.new('RGB', (16, 16), '#4679E1')
        return img


# ── start web server in background ───────────────────────────────────────────
def _start_server():
    import server as srv
    from http.server import ThreadingHTTPServer
    httpd = ThreadingHTTPServer(('127.0.0.1', PORT), srv.Handler)
    httpd.serve_forever()


# ── tray menu actions ─────────────────────────────────────────────────────────
def _open_browser(_icon=None, _item=None):
    webbrowser.open(URL)


def _quit(icon, _item=None):
    icon.stop()
    os._exit(0)


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    # Start server thread
    t = threading.Thread(target=_start_server, daemon=True)
    t.start()

    # Give server a moment then open browser
    import time
    time.sleep(1.2)
    webbrowser.open(URL)

    # Try tray icon (pystray + pillow required)
    try:
        import pystray
        from pystray import MenuItem as Item, Menu

        icon = pystray.Icon(
            name   = 'faire_filler',
            icon   = _make_icon(),
            title  = TITLE,
            menu   = Menu(
                Item('Open FAIRe Filler', _open_browser, default=True),
                Menu.SEPARATOR,
                Item('Quit', _quit),
            ),
        )
        print(f'{TITLE} running at {URL}  (look for the tray icon)')
        icon.run()

    except ImportError:
        # pystray not installed — just keep the server alive
        print(f'{TITLE} running at {URL}')
        print('Install pystray and pillow for a system-tray icon.')
        print('Press Ctrl+C to stop.')
        try:
            while True:
                time.sleep(60)
        except KeyboardInterrupt:
            pass


if __name__ == '__main__':
    main()
