#!/bin/bash
# ─────────────────────────────────────────────────────────
# build_mac_app.sh
# Run this script ONCE to create Font Audit.app on your Mac.
# After building, drag any .pptx file onto the app icon.
# ─────────────────────────────────────────────────────────

set -e

APP_NAME="Font Audit"
APP_DIR="$HOME/Desktop/${APP_NAME}.app"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "Building ${APP_NAME}.app on your Desktop..."

# Verify required files exist
if [ ! -f "$SCRIPT_DIR/font_audit.py" ]; then
    echo "  ✗ ERROR: font_audit.py not found next to this script."
    exit 1
fi
if [ ! -f "$SCRIPT_DIR/FontAudit.applescript" ]; then
    echo "  ✗ ERROR: FontAudit.applescript not found next to this script."
    exit 1
fi

# Install font_audit.py to ~/scripts/
mkdir -p "$HOME/scripts"
cp "$SCRIPT_DIR/font_audit.py" "$HOME/scripts/font_audit.py"
chmod +x "$HOME/scripts/font_audit.py"
echo "  ✓ Installed font_audit.py to ~/scripts/"

# Remove old app
rm -rf "$APP_DIR"

# Compile the AppleScript into a proper droplet .app
osacompile -o "$APP_DIR" "$SCRIPT_DIR/FontAudit.applescript"

echo ""
echo "═══════════════════════════════════════════════════"
echo "  ✓ ${APP_NAME}.app created on your Desktop!"
echo "═══════════════════════════════════════════════════"
echo ""
echo "  Usage:"
echo "    • Drag any .pptx file onto the app icon"
echo "    • Or double-click the app and pick a file"
echo ""
echo "  Prerequisites (one-time):"
echo "    python3 -m venv ~/font-audit-env"
echo "    source ~/font-audit-env/bin/activate"
echo "    pip install python-pptx"
echo ""
echo "  First launch: macOS may block it —"
echo "    Right-click → Open → Open"
echo ""
