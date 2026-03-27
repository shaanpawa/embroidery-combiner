"""
Micro Automation — Desktop Entry Point

Starts the FastAPI backend, opens the browser, and shows a system tray icon.
Used by PyInstaller to create the Windows EXE.
"""

import logging
import logging.handlers
import os
import socket
import sys
import tempfile
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
    LOG_DIR = os.path.join(local_appdata, "MicroAutomation", "logs")
else:
    DATA_DIR = os.path.join(APP_ROOT, "data")
    LOG_DIR = os.path.join(APP_ROOT, "logs")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

LOG_FILE = os.path.join(LOG_DIR, "micro_automation.log")

# Set environment variables BEFORE importing the app
os.environ["DESKTOP_MODE"] = "true"
os.environ["AUTH_DISABLED"] = "true"
os.environ["MICRO_DATA_DIR"] = DATA_DIR
os.environ["MICRO_APP_ROOT"] = APP_ROOT


# ---------------------------------------------------------------------------
# Logging setup — captures ALL output (critical since console=False in EXE)
# ---------------------------------------------------------------------------

def setup_logging():
    """Configure file-based logging. Redirects stdout/stderr to log file."""
    handler = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8",
    )
    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S",
    ))

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    root.addHandler(handler)

    # Redirect stdout/stderr to log file (EXE has no console)
    class _LogStream:
        def __init__(self, logger, level):
            self._logger = logger
            self._level = level
            self._buf = ""

        def write(self, msg):
            if msg and msg.strip():
                self._logger.log(self._level, msg.rstrip())

        def flush(self):
            pass

    sys.stdout = _LogStream(logging.getLogger("stdout"), logging.INFO)
    sys.stderr = _LogStream(logging.getLogger("stderr"), logging.ERROR)


log = logging.getLogger("MicroAutomation")


# ---------------------------------------------------------------------------
# Error dialog — visible feedback when something goes wrong
# ---------------------------------------------------------------------------

def show_error_dialog(title: str, message: str):
    """Show a native error dialog on Windows, fall back to print elsewhere."""
    log.error(f"ERROR DIALOG: {title} — {message}")
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)  # MB_ICONERROR
            return
        except Exception:
            pass
    # Fallback for non-Windows or if ctypes fails
    print(f"[{title}] {message}")


# ---------------------------------------------------------------------------
# Port discovery
# ---------------------------------------------------------------------------

def find_free_port(start: int = 8000, end: int = 8020):
    """Find an available port in the given range. Returns None if all taken."""
    for port in range(start, end + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port))
                return port
        except OSError:
            log.debug(f"Port {port} is in use, trying next...")
            continue
    return None


def wait_for_server(port: int, timeout: float = 20.0) -> bool:
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
# Startup splash page — instant visual feedback
# ---------------------------------------------------------------------------

def create_splash_page(port: int) -> str:
    """Create a temporary HTML loading page that polls the server health endpoint.

    Returns the path to the temp HTML file.
    """
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Micro Automation — Starting</title>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    color: #e0e0e0;
    display: flex; align-items: center; justify-content: center;
    min-height: 100vh;
  }}
  .container {{ text-align: center; }}
  .logo {{ font-size: 48px; font-weight: 700; color: #7c83db; margin-bottom: 8px; }}
  .subtitle {{ font-size: 14px; color: #888; margin-bottom: 40px; }}
  .spinner {{
    width: 48px; height: 48px; margin: 0 auto 24px;
    border: 4px solid rgba(124, 131, 219, 0.2);
    border-top-color: #7c83db;
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
  }}
  @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
  .status {{ font-size: 16px; color: #aaa; }}
  .error {{ color: #ff6b6b; margin-top: 20px; display: none; max-width: 500px; line-height: 1.6; }}
  .error-title {{ font-size: 18px; font-weight: 700; margin-bottom: 12px; }}
  .error-hint {{ font-size: 13px; color: #aaa; margin-top: 12px; }}
  .log-path {{
    font-size: 12px; color: #7c83db; margin-top: 12px; word-break: break-all;
    background: rgba(124, 131, 219, 0.1); padding: 8px 12px; border-radius: 6px;
    cursor: pointer; user-select: all;
  }}
  .log-path:hover {{ background: rgba(124, 131, 219, 0.2); }}
  .copy-btn {{
    margin-top: 8px; padding: 6px 16px; background: #7c83db; color: white;
    border: none; border-radius: 4px; cursor: pointer; font-size: 12px;
  }}
  .copy-btn:hover {{ background: #6b72c4; }}
</style>
</head>
<body>
<div class="container">
  <div class="logo">Micro</div>
  <div class="subtitle">Automation by Ossia</div>
  <div class="spinner" id="spinner"></div>
  <div class="status" id="status">Starting server... / กำลังเริ่มเซิร์ฟเวอร์...</div>
  <div class="error" id="error">
    <div class="error-title">Server failed to start / เซิร์ฟเวอร์เริ่มไม่ได้</div>
    <div id="error-detail"></div>
    <div class="error-hint">
      If blocked by antivirus, add MicroAutomation to exceptions.<br>
      หากถูกบล็อกโดยโปรแกรมป้องกันไวรัส ให้เพิ่ม MicroAutomation เป็นข้อยกเว้น
    </div>
    <div class="log-path" id="log-path" onclick="copyLogPath()" title="Click to copy / คลิกเพื่อคัดลอก">{LOG_FILE.replace(os.sep, '/')}</div>
    <button class="copy-btn" onclick="copyLogPath()">Copy log path / คัดลอกที่อยู่ล็อก</button>
  </div>
</div>
<script>
  const port = {port};
  const url = 'http://localhost:' + port;
  let attempts = 0;
  const maxAttempts = 60; // 30 seconds at 500ms intervals

  function checkHealth() {{
    attempts++;
    const secs = Math.round(attempts * 0.5);
    document.getElementById('status').textContent =
      'Starting server... (' + secs + 's) / กำลังเริ่มเซิร์ฟเวอร์... (' + secs + ' วินาที)';

    fetch(url + '/api/health', {{ mode: 'cors' }})
      .then(r => {{ if (r.ok) window.location.href = url; else retry(); }})
      .catch(() => retry());
  }}

  function retry() {{
    if (attempts >= maxAttempts) {{
      document.getElementById('spinner').style.display = 'none';
      document.getElementById('status').style.display = 'none';
      document.getElementById('error').style.display = 'block';
      document.getElementById('error-detail').innerHTML =
        'The server did not respond after ' + Math.round(maxAttempts * 0.5) + ' seconds.<br>' +
        'เซิร์ฟเวอร์ไม่ตอบกลับหลังจาก ' + Math.round(maxAttempts * 0.5) + ' วินาที<br><br>' +
        'Please send the log file to your administrator for help.<br>' +
        'กรุณาส่งไฟล์ล็อกให้ผู้ดูแลระบบเพื่อขอความช่วยเหลือ';
      return;
    }}
    setTimeout(checkHealth, 500);
  }}

  function copyLogPath() {{
    const path = document.getElementById('log-path').textContent;
    navigator.clipboard.writeText(path).then(() => {{
      const btn = document.querySelector('.copy-btn');
      btn.textContent = 'Copied! / คัดลอกแล้ว!';
      setTimeout(() => {{ btn.textContent = 'Copy log path / คัดลอกที่อยู่ล็อก'; }}, 2000);
    }});
  }}

  // Start polling after a brief initial delay
  setTimeout(checkHealth, 500);
</script>
</body>
</html>"""

    splash_path = os.path.join(tempfile.gettempdir(), "micro_loading.html")
    with open(splash_path, "w", encoding="utf-8") as f:
        f.write(html)
    return splash_path


# ---------------------------------------------------------------------------
# Server thread
# ---------------------------------------------------------------------------

_server_error = None  # Shared variable for error reporting to main thread


def run_server(port: int, max_retries: int = 3):
    """Start uvicorn with retry logic."""
    global _server_error

    # Pre-validate that the server module is importable
    try:
        log.info("Validating server imports...")
        import uvicorn
        from api.server import app  # noqa: F811
        log.info("Server imports OK")
    except ImportError as e:
        _server_error = f"Missing dependency: {e}"
        log.critical(f"Import validation failed: {e}", exc_info=True)
        return

    for attempt in range(1, max_retries + 1):
        try:
            log.info(f"Server startup attempt {attempt}/{max_retries} on port {port}")
            uvicorn.run(
                app,
                host="127.0.0.1",
                port=port,
                log_level="warning",
                access_log=False,
            )
            return  # uvicorn.run() blocks until shutdown — if we get here, clean exit
        except Exception as e:
            log.error(f"Server attempt {attempt} failed: {e}", exc_info=True)
            _server_error = str(e)
            if attempt < max_retries:
                log.info(f"Retrying in 2 seconds...")
                time.sleep(2)

    log.error("All server startup attempts exhausted.")


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
            f"Micro Automation — localhost:{port}",
            menu=pystray.Menu(
                pystray.MenuItem("Open Micro Automation", on_open, default=True),
                pystray.MenuItem("Quit", on_quit),
            ),
        )
        icon.run()
    except ImportError:
        # pystray not available — just wait for stop event
        log.warning("pystray not available, running without system tray")
        stop_event.wait()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    setup_logging()
    log.info("=" * 60)
    log.info("Micro Automation starting")
    log.info(f"  APP_ROOT: {APP_ROOT}")
    log.info(f"  DATA_DIR: {DATA_DIR}")
    log.info(f"  LOG_FILE: {LOG_FILE}")
    log.info(f"  Frozen: {getattr(sys, 'frozen', False)}")
    log.info("=" * 60)

    # Verify static frontend exists (PyInstaller bundle)
    if getattr(sys, "frozen", False):
        static_dir = os.path.join(APP_ROOT, "static_web")
        if not os.path.isdir(static_dir) or not os.listdir(static_dir):
            log.error(f"static_web/ directory missing or empty at {static_dir}")
            show_error_dialog(
                "Micro Automation",
                "Installation may be corrupted — frontend files are missing.\n"
                "การติดตั้งอาจเสียหาย — ไม่พบไฟล์ frontend\n\n"
                "Please reinstall the application.\n"
                "กรุณาติดตั้งแอปพลิเคชันใหม่\n\n"
                f"Expected: {static_dir}\n"
                f"Log file: {LOG_FILE}",
            )
            sys.exit(1)
        log.info(f"static_web/ verified: {len(os.listdir(static_dir))} files")

    # Find a free port
    port = find_free_port()
    if port is None:
        show_error_dialog(
            "Micro Automation",
            "Could not find a free port (8000-8020).\n"
            "ไม่พบพอร์ตว่าง (8000-8020)\n\n"
            "Please close other applications using these ports and try again.\n"
            "กรุณาปิดแอปพลิเคชันอื่นที่ใช้พอร์ตเหล่านี้แล้วลองอีกครั้ง\n\n"
            f"Log file: {LOG_FILE}",
        )
        sys.exit(1)
    log.info(f"Using port {port}")

    # Open splash page immediately — gives instant visual feedback
    splash_path = create_splash_page(port)
    log.info(f"Opening splash page: {splash_path}")
    webbrowser.open(f"file:///{splash_path}")

    # Start server in background thread
    server_thread = threading.Thread(target=run_server, args=(port,), daemon=True)
    server_thread.start()

    # Wait for server to be ready
    if wait_for_server(port):
        log.info("Server is ready — splash page will auto-redirect")
        # Write URL fallback file to desktop (user can open manually if browser redirect fails)
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            if os.path.isdir(desktop):
                url_file = os.path.join(desktop, "MicroAutomation.url")
                with open(url_file, "w") as f:
                    f.write(f"[InternetShortcut]\nURL=http://localhost:{port}\n")
                log.info(f"URL shortcut written to {url_file}")
        except Exception as e:
            log.warning(f"Could not write URL shortcut: {e}")
    else:
        log.error("Server did not start within timeout")
        show_error_dialog(
            "Micro Automation",
            f"The server failed to start. / เซิร์ฟเวอร์เริ่มไม่ได้\n\n"
            f"Error: {_server_error or 'Unknown (timeout)'}\n\n"
            f"If blocked by antivirus, add MicroAutomation to exceptions.\n"
            f"หากถูกบล็อกโดยโปรแกรมป้องกันไวรัส ให้เพิ่ม MicroAutomation เป็นข้อยกเว้น\n\n"
            f"Log file: {LOG_FILE}",
        )
        sys.exit(1)

    # Run system tray (blocks until quit)
    stop_event = threading.Event()
    run_tray(port, stop_event)

    # Clean exit
    log.info("Shutting down")
    sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Last-resort error handling — log and show dialog
        try:
            setup_logging()
        except Exception:
            pass
        log.critical(f"Fatal error: {e}", exc_info=True)
        show_error_dialog(
            "Micro Automation",
            f"An unexpected error occurred. / เกิดข้อผิดพลาดที่ไม่คาดคิด\n\n{e}\n\n"
            f"Please send the log file to your administrator.\n"
            f"กรุณาส่งไฟล์ล็อกให้ผู้ดูแลระบบ\n\n"
            f"Log file: {LOG_FILE}",
        )
        sys.exit(1)
