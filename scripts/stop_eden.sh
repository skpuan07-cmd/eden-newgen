#!/bin/bash
# Stop the running Sahana Eden / web2py server
PIDFILE="/tmp/web2py.pid"
if [ -f "$PIDFILE" ]; then
  PID=$(cat "$PIDFILE")
  if kill -0 "$PID" 2>/dev/null; then
    echo "Stopping web2py (PID: $PID)..."
    kill -SIGTERM "$PID"
    sleep 2
    if kill -0 "$PID" 2>/dev/null; then
      echo "Force killing..."
      kill -9 "$PID"
    fi
    echo "✓ Server stopped."
  else
    echo "Process $PID not running (stale pidfile)."
  fi
  rm -f "$PIDFILE"
else
  # Fallback: kill any web2py process
  PIDS=$(pgrep -f "web2py.py --no_gui")
  if [ -n "$PIDS" ]; then
    echo "Killing web2py processes: $PIDS"
    pkill -f "web2py.py --no_gui"
    sleep 2
    echo "✓ All web2py processes stopped."
  else
    echo "No web2py process found."
  fi
fi
