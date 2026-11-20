#!/bin/bash
# ============================================================
# Sahana Eden — Startup Script (web2py + FastAPI)
# ============================================================
# Usage:
#   ./start_eden.sh                # foreground, random password
#   ./start_eden.sh mypass         # background, custom password
#   ./start_eden.sh mypass daemon  # background, daemon mode
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
WEB2PY_DIR="$PROJECT_DIR/eden_install/web2py"
PYTHON="$PROJECT_DIR/.venv/bin/python"
FASTAPI_DIR="$WEB2PY_DIR/applications/eden/fastapi_app"
FASTAPI_SERVER="$PROJECT_DIR/.venv/bin/uvicorn"
PORT="${EDEN_PORT:-8000}"
API_PORT="${EDEN_API_PORT:-8001}"
IP="${EDEN_IP:-127.0.0.1}"

# If no password provided, generate one
if [ -z "$1" ]; then
  ADMIN_PW="Eden$(date +%s | tail -c 6)"
  BACKGROUND=""
else
  ADMIN_PW="$1"
  if [ "$2" = "daemon" ] || [ "$2" = "bg" ]; then
    BACKGROUND="daemon"
  fi
fi

echo "============================================================"
echo "  Sahana Eden — Starting"
echo "============================================================"
echo "  Directory : $WEB2PY_DIR"
echo "  Listen IP : $IP"
echo "  Port      : $PORT"
echo "  API Port  : $API_PORT"
echo "  Admin PW  : $ADMIN_PW"
echo "  Mode      : ${BACKGROUND:-foreground}"
echo "============================================================"
echo ""
echo "  URL to access Eden:"
echo "    http://localhost:$PORT/eden"
echo ""
echo "  Demo accounts (already in database):"
echo "    admin@example.com       / testing     (system administrator)"
echo "    normaluser@example.com  / testing     (regular user)"
echo ""
echo "  Web2py admin (for error tickets, app management):"
echo "    http://localhost:$PORT/admin    (password = the admin PW above)"
echo ""
echo "============================================================"
echo ""

cd "$WEB2PY_DIR"
"$PYTHON" "$SCRIPT_DIR/configure_eden_home.py" "$WEB2PY_DIR"

start_api() {
  if curl -fsS "http://$IP:$API_PORT/eden/health/" >/dev/null 2>&1; then
    echo "✓ FastAPI already ready at http://localhost:$API_PORT"
    return
  fi

  if [ ! -x "$FASTAPI_SERVER" ] || [ ! -f "$FASTAPI_DIR/main.py" ]; then
    echo "FastAPI is not installed. Set up the project .venv and Eden runtime first." >&2
    exit 1
  fi

  DATABASE_URL="sqlite+aiosqlite:///$PROJECT_DIR/eden_fastapi.db" \
    nohup "$FASTAPI_SERVER" main:app --app-dir "$FASTAPI_DIR" \
      --host="$IP" --port="$API_PORT" > /tmp/eden-fastapi.log 2>&1 &
  API_PID=$!
  echo "$API_PID" > /tmp/eden-fastapi.pid
  echo "✓ FastAPI started (PID: $API_PID), log: /tmp/eden-fastapi.log"
}

if [ "$BACKGROUND" = "daemon" ]; then
  start_api
  nohup "$PYTHON" web2py.py --no_gui -a "$ADMIN_PW" --ip="$IP" --port="$PORT" \
       > /tmp/web2py.log 2>&1 &
  PID=$!
  echo $PID > /tmp/web2py.pid
  echo "✓ Server started in background (PID: $PID)"
  echo "  Log: /tmp/web2py.log"
  echo "  Stop: ./stop_eden.sh"
  echo ""
  echo "Waiting for server to be ready..."
  for i in $(seq 1 30); do
    if curl -sS -o /dev/null "http://$IP:$PORT/" 2>/dev/null; then
      echo "✓ Server ready after ${i}s — open http://localhost:$PORT/eden"
      break
    fi
    sleep 1
  done
  for i in $(seq 1 30); do
    if curl -fsS -o /dev/null "http://$IP:$API_PORT/eden/health/" 2>/dev/null; then
      echo "✓ API ready after ${i}s — http://localhost:$API_PORT/eden/docs"
      API_READY=1
      break
    fi
    sleep 1
  done
  if [ "${API_READY:-0}" -ne 1 ]; then
    echo "FastAPI did not become ready; inspect /tmp/eden-fastapi.log." >&2
    exit 1
  fi
else
  # Foreground — Ctrl+C to stop
  start_api
  echo "Starting in foreground. Press Ctrl+C to stop."
  echo ""
  exec "$PYTHON" web2py.py --no_gui -a "$ADMIN_PW" --ip="$IP" --port="$PORT"
fi
