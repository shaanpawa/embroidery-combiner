"""
Micro Automation — Desktop Entry Point

Starts the FastAPI backend, opens the browser, and shows a system tray icon.
Used by PyInstaller to create the Windows EXE.
"""

import os
import sys
import socket
import threading
import time
import webbrowser

# ---------------------------------------------------------------------------
# Resolve paths for frozen (PyInstaller) vs development mode
# ---------------------------------------------------------------------------

if getattr(sys, "frozen", False):
    # Running as PyInstaller bundle
    APP_ROOT = os.path.dirname(sys.executable)
else:
    # Running as script (development)
    APP_ROOT = os.path.dirname(os.path.abspath(__file__))

# Data directory: %LOCALAPPDATA%\MicroAutomation\data on Windows, fallback to ./data
if sys.platform == "win32":
    local_appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    DATA_DIR = os.path.join(local_appdata, "MicroAutomation", "data")
else:
    DATA_DIR = os.path.join(APP_ROOT, "data")

os.makedirs(DATA_DIR, exist_ok=True)

# Set environment variables BEFORE importing the app
os.environ["DESKTOP_MODE"] = "true"
os.environ["AUTH_DISABLED"] = "true"
os.environ["MICRO_DATA_DIR"] = DATA_DIR
os.environ["MICRO_APP_ROOT"] = APP_ROOT


# ---------------------------------------------------------------------------
# Port discovery
# ---------------------------------------------------------------------------

def find_free_port(start: int = 8000, end: int = 8010) -> int:
    """Find an available port in the given range."""
    for port in range(start, end + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port))
                return port
        except OSError:
            continue
    return start  # Fallback


def wait_for_server(port: int, timeout: float = 15.0) -> bool:
    """Wait for the server to start accepting connections."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                s.connect(("127.0.0.1", port))
                return True
        except (ConnectionRefusedError, OSError):
            time.sleep(0.3)
    return False


# ---------------------------------------------------------------------------
# Server thread
# ---------------------------------------------------------------------------

def run_server(port: int):
    """Start uvicorn in the current thread."""
    import uvicorn
    uvicorn.run(
        "api.server:app",
        host="127.0.0.1",
        port=port,
        log_level="warning",
        access_log=False,
    )


# ---------------------------------------------------------------------------
# System tray (optional — graceful fallback if pystray unavailable)
# ---------------------------------------------------------------------------

def run_tray(port: int, stop_event: threading.Event):
    """Show a system tray icon with Open / Quit options."""
    try:
        import pystray
        from PIL import Image, ImageDraw

        # Create a simple icon (blue square with white M)
        img = Image.new("RGB", (64, 64), color=(59, 73, 133))
        draw = ImageDraw.Draw(img)
        draw.text((18, 12), "M", fill="white")

        def on_open(icon, item):
            webbrowser.open(f"http://localhost:{port}")

        def on_quit(icon, item):
            icon.stop()
            stop_event.set()

        icon = pystray.Icon(
            "MicroAutomation",
            img,
            "Micro Automation",
            menu=pystray.Menu(
                pystray.MenuItem("Open Micro Automation", on_open, default=True),
                pystray.MenuItem("Quit", on_quit),
            ),
        )
        icon.run()
    except ImportError:
        # pystray not available — just wait for stop event
        stop_event.wait()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    port = find_free_port()

    # Start server in background thread
    server_thread = threading.Thread(target=run_server, args=(port,), daemon=True)
    server_thread.start()

    # Wait for server to be ready
    if wait_for_server(port):
        webbrowser.open(f"http://localhost:{port}")
    else:
        print(f"Warning: Server did not start within timeout on port {port}")
        webbrowser.open(f"http://localhost:{port}")

    # Run system tray (blocks until quit)
    stop_event = threading.Event()
    run_tray(port, stop_event)

    # Clean exit
    sys.exit(0)


if __name__ == "__main__":
    main()
