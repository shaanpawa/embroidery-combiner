"""Starts both the FastAPI backend and Next.js frontend for local dev."""
import os
import subprocess
import sys
import signal
import time

root = os.path.dirname(os.path.abspath(__file__))
web_dir = os.path.join(root, "web")
venv_python = os.path.join(root, ".venv", "bin", "python3")

# Start API server in background (uses venv Python for dependencies)
api_env = os.environ.copy()
api_env["AUTH_DISABLED"] = "true"  # Skip auth in dev mode
api_proc = subprocess.Popen(
    [venv_python, "-m", "uvicorn", "api.server:app", "--reload", "--port", "8000"],
    cwd=root,
    env=api_env,
)

# Start Next.js frontend (blocks — this is the "main" process for the preview tool)
try:
    subprocess.run(["npm", "run", "dev", "--", "-p", "5123"], cwd=web_dir)
finally:
    api_proc.terminate()
    api_proc.wait(timeout=5)
