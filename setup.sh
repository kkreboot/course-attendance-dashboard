#!/usr/bin/env bash
# One-time setup on a fresh machine: creates a per-machine venv and installs
# requirements.txt into it. Run this once after transferring the project,
# then use ./launch_dashboard.sh (or run.py / streamlit directly) from then on.
#
# The venv lives in .venv-<OS> (.venv-Darwin, .venv-Linux, .venv-Windows)
# rather than a single shared .venv, because this project folder is synced
# with Dropbox across macOS/Linux/Windows. A venv is not portable -- its
# interpreter symlinks and console-script shebangs are absolute paths from
# the machine that built it -- so one shared .venv means whichever machine
# ran setup.sh last is the only one that works. One dir per OS lets all
# three coexist in the same synced folder.
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

VENV=".venv-$(uname -s)"   # .venv-Darwin on macOS, .venv-Linux on Linux
PYTHON="${PYTHON:-python3}"

if ! command -v "$PYTHON" >/dev/null 2>&1; then
    echo "$PYTHON not found. Install Python 3.10+ first (e.g. sudo apt install python3 python3-venv)." >&2
    exit 1
fi

# A venv copied in from another machine can still *start* (bin/python may
# resolve to a working system interpreter via the pyvenv.cfg home path), so
# checking that bin/python runs is not enough -- also require the
# site-packages dir for the interpreter's own version to exist. That is what
# a Linux 3.8 venv opened on macOS lacks.
venv_ok() {
    [ -x "$VENV/bin/python" ] || return 1
    "$VENV/bin/python" -c "import os, sys, sysconfig
assert sys.prefix != sys.base_prefix
assert os.path.isdir(sysconfig.get_paths()['purelib'])" >/dev/null 2>&1
}

if [ -d "$VENV" ] && ! venv_ok; then
    echo "Existing $VENV is not usable on this machine; rebuilding..."
    rm -rf "$VENV"
fi

if [ ! -d "$VENV" ]; then
    echo "Creating venv in $VENV..."
    "$PYTHON" -m venv "$VENV"
fi

# Keep the venv out of Dropbox sync where the client supports it: it is
# hundreds of MB of machine-specific binaries that no other machine can use.
# Best-effort only -- an older Dropbox client or a non-Dropbox copy of the
# project just ignores this.
if command -v xattr >/dev/null 2>&1; then
    xattr -w com.dropbox.ignored 1 "$VENV" 2>/dev/null || true
elif command -v attr >/dev/null 2>&1; then
    attr -s com.dropbox.ignored -V 1 "$VENV" >/dev/null 2>&1 || true
fi

echo "Installing requirements.txt..."
"$VENV/bin/pip" install --upgrade pip
"$VENV/bin/pip" install -r requirements.txt

# Fail here rather than at the first click if the install came up short.
"$VENV/bin/python" -c "
import pandas, streamlit, reportlab, openpyxl, xlrd
print('OK: pandas', pandas.__version__, '| streamlit', streamlit.__version__)
"

chmod +x launch_dashboard.sh setup.sh

PROJECT_DIR="$(pwd)"

if [ "$(uname -s)" = "Darwin" ]; then
    # macOS has no .desktop entries; the equivalent is a real .app bundle in
    # ~/Applications, so the dashboard is in Spotlight and can be pinned to
    # the Dock. It only wraps launch_dashboard.sh -- see make_mac_app.py.
    "$VENV/bin/python" make_mac_app.py || \
        echo "App bundle not created; ./launch_dashboard.sh still works." >&2
    echo
    echo "Setup complete (venv: $VENV)."
    echo "  - Launch the '<COURSE_CODE> Dashboard' app from Spotlight, ~/Applications, or the Dock"
    echo "  - Or run ./launch_dashboard.sh directly"
    echo "  - Or $VENV/bin/python run.py --help for the CLI"
    exit 0
fi

# ---- desktop launcher: paths are only known once the project has actually
# landed somewhere on this machine, so the .desktop entry is generated here
# rather than shipped with a guessed path baked in. ----
DESKTOP_FILE="$HOME/.local/share/applications/course-dashboard.desktop"
mkdir -p "$HOME/.local/share/applications"
cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Name=Course Dashboard
Comment=Launch the course attendance & seating control panel
Exec=${PROJECT_DIR}/launch_dashboard.sh
Path=${PROJECT_DIR}
Terminal=false
Categories=Education;Office;
EOF
chmod +x "$DESKTOP_FILE"

# A copy on the Desktop too, for a literal double-click -- most file
# managers refuse to run a .desktop file that isn't marked trusted/executable
# yet, hence the chmod above and (where the tool exists) marking it trusted.
if [ -d "$HOME/Desktop" ]; then
    cp "$DESKTOP_FILE" "$HOME/Desktop/Course Dashboard.desktop"
    chmod +x "$HOME/Desktop/Course Dashboard.desktop"
    if command -v gio >/dev/null 2>&1; then
        gio set "$HOME/Desktop/Course Dashboard.desktop" metadata::trusted true 2>/dev/null || true
    fi
fi

echo
echo "Setup complete (venv: $VENV)."
echo "  - Double-click 'Course Dashboard' on your Desktop (or in your app launcher/menu)"
echo "  - Or run ./launch_dashboard.sh directly"
echo "  - Or $VENV/bin/python run.py --help for the CLI"
