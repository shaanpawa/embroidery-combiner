#!/bin/bash
# ============================================================
#  Test Desktop Mode on Mac
#  Tests the full factory experience locally
# ============================================================
#
#  Usage:
#    ./test_desktop.sh          # Build static frontend + run
#    ./test_desktop.sh --quick  # Skip frontend build (use existing static_web/)
#
# ============================================================

set -e
cd "$(dirname "$0")"

PORT=8099
YELLOW='\033[1;33m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
NC='\033[0m'

echo ""
echo -e "${BLUE}╔══════════════════════════════════════════════╗${NC}"
echo -e "${BLUE}║   Micro Automation — Desktop Mode Test       ║${NC}"
echo -e "${BLUE}╚══════════════════════════════════════════════╝${NC}"
echo ""

# Step 1: Build static frontend (unless --quick)
if [ "$1" != "--quick" ]; then
  echo -e "${YELLOW}Step 1: Building static frontend...${NC}"
  .venv/bin/python build_desktop.py --skip-installer 2>&1 | grep -E "(===|Moving|Restored|Copied|✓|Step)" || true

  # PyInstaller will fail on Mac (expected) — we only need the static_web/ build
  if [ ! -d "static_web" ]; then
    echo "ERROR: static_web/ not created. Check build_desktop.py output."
    exit 1
  fi
  echo -e "${GREEN}✅ Static frontend built → static_web/${NC}"
else
  echo -e "${YELLOW}Step 1: Skipping frontend build (--quick)${NC}"
  if [ ! -d "static_web" ]; then
    echo "ERROR: static_web/ doesn't exist. Run without --quick first."
    exit 1
  fi
fi

echo ""
echo -e "${YELLOW}Step 2: Starting desktop mode on port ${PORT}...${NC}"
echo -e "   Open your browser to: ${GREEN}http://localhost:${PORT}${NC}"
echo -e "   Press ${YELLOW}Ctrl+C${NC} to stop"
echo ""

# Start in desktop mode
export DESKTOP_MODE=true
export AUTH_DISABLED=true
export MICRO_APP_ROOT="$(pwd)"
export MICRO_DATA_DIR="$(pwd)/data"

.venv/bin/python -m uvicorn api.server:app --host 127.0.0.1 --port $PORT --log-level info
