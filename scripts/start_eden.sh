#!/bin/bash
# ============================================================
# Sahana Eden — Startup Script
# ============================================================
# Usage:
#   ./start_eden.sh                # foreground, random password
#   ./start_eden.sh mypass         # background, custom password
#   ./start_eden.sh mypass daemon  # background, daemon mode
# ============================================================

set -e

WEB2PY_DIR="/home/z/my-project/eden_install/web2py"
PORT="${EDEN_PORT:-8000}"
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

if [ "$BACKGROUND" = "daemon" ]; then
  nohup python3 web2py.py --no_gui -a "$ADMIN_PW" --ip="$IP" --port="$PORT" \
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
else
  # Foreground — Ctrl+C to stop
  echo "Starting in foreground. Press Ctrl+C to stop."
  echo ""
  exec python3 web2py.py --no_gui -a "$ADMIN_PW" --ip="$IP" --port="$PORT"
fi
