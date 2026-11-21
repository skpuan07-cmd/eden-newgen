#!/bin/bash
# Stop the running Sahana Eden services.

stop_service() {
  NAME="$1"
  PIDFILE="$2"
  EXPECTED="$3"

  if [ ! -f "$PIDFILE" ]; then
    echo "$NAME pidfile not found."
    return
  fi

  PID=$(cat "$PIDFILE")
  case "$PID" in
    ''|*[!0-9]*)
      echo "Invalid $NAME pidfile."
      rm -f "$PIDFILE"
      return
      ;;
  esac

  COMMAND=$(ps -p "$PID" -o args= 2>/dev/null || true)
  case "$COMMAND" in
    *"$EXPECTED"*) ;;
    *)
      echo "$NAME process not running (stale pidfile)."
      rm -f "$PIDFILE"
      return
      ;;
  esac

  echo "Stopping $NAME (PID: $PID)..."
  kill -SIGTERM "$PID"
  for _ in $(seq 1 10); do
    if ! kill -0 "$PID" 2>/dev/null; then
      echo "✓ $NAME stopped."
      rm -f "$PIDFILE"
      return
    fi
    sleep 1
  done

  echo "$NAME did not stop gracefully; sending SIGKILL."
  kill -SIGKILL "$PID"
  rm -f "$PIDFILE"
}

stop_service "web2py" "/tmp/web2py.pid" "web2py.py --no_gui"
stop_service "FastAPI" "/tmp/eden-fastapi.pid" "uvicorn main:app"
