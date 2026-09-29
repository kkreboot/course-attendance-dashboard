#!/usr/bin/env bash
# Starts the Streamlit dashboard server in the background, waits for it to
# come up, then opens it in the default browser. Double-click this (or the
# .desktop entry pointing at it) to go from a cold start to the dashboard
# in a browser tab with no terminal typing -- the Linux equivalent of
# launch_dashboard_chrome.bat.
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

PORT="${DASHBOARD_PORT:-8501}"
URL="http://localhost:${PORT}"
LOG="/tmp/course-dashboard.log"

# Prefer this machine's own venv (see setup.sh -- .venv-Darwin / .venv-Linux,
# one per OS because the project folder is Dropbox-synced and a venv is not
# portable across machines). Fall back to a legacy shared .venv for a copy
# that hasn't re-run setup.sh yet, then to whatever is on PATH.
VENV=".venv-$(uname -s)"
if [ -x "$VENV/bin/streamlit" ]; then
    STREAMLIT="$VENV/bin/streamlit"
elif [ -x ".venv/bin/streamlit" ] && .venv/bin/python -c "import streamlit" >/dev/null 2>&1; then
    STREAMLIT=".venv/bin/streamlit"
elif command -v streamlit >/dev/null 2>&1; then
    STREAMLIT="streamlit"
else
    echo "streamlit not found. Run ./setup.sh first to create a venv and install requirements.txt." >&2
    exit 1
fi

# If something's already listening on the port, assume it's a dashboard
# already running (e.g. from a previous launch) and just open a tab.
if ! curl -s -o /dev/null "$URL"; then
    echo "Starting dashboard server (log: $LOG)..."
    nohup "$STREAMLIT" run dashboard.py \
        --server.port "$PORT" --server.headless true \
        --browser.gatherUsageStats false \
        >"$LOG" 2>&1 &
    disown

    # Poll instead of a fixed sleep -- first run after a reboot (cold
    # filesystem cache, venv not yet warmed up) is slower than a re-launch.
    for _ in $(seq 1 60); do
        if curl -s -o /dev/null "$URL"; then
            break
        fi
        sleep 1
    done
fi

if [ "$(uname -s)" = "Darwin" ] && command -v open >/dev/null 2>&1; then
    open "$URL" >/dev/null 2>&1 &
elif command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$URL" >/dev/null 2>&1 &
elif command -v google-chrome >/dev/null 2>&1; then
    google-chrome --new-window "$URL" >/dev/null 2>&1 &
elif command -v chromium >/dev/null 2>&1; then
    chromium --new-window "$URL" >/dev/null 2>&1 &
elif command -v chromium-browser >/dev/null 2>&1; then
    chromium-browser --new-window "$URL" >/dev/null 2>&1 &
else
    echo "No browser launcher found (open/xdg-open/google-chrome/chromium). Open $URL manually."
fi
