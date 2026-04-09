#!/bin/bash
# ──────────────────────────────────────────────
# EDL to CSV — Installer
# Run this once to allow the app on your Mac.
# ──────────────────────────────────────────────

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_NAME="EDL to CSV.app"
APP_PATH="$SCRIPT_DIR/$APP_NAME"

clear
echo ""
echo "  ╔══════════════════════════════════════╗"
echo "  ║   EDL to CSV — Installer             ║"
echo "  ║   by Chad Littlepage                 ║"
echo "  ╚══════════════════════════════════════╝"
echo ""

if [ ! -d "$APP_PATH" ]; then
    echo "  ✖ Could not find '$APP_NAME' in the same folder as this script."
    echo "    Make sure both files are in the same directory."
    echo ""
    read -n 1 -s -r -p "  Press any key to close..."
    exit 1
fi

echo "  This will allow '$APP_NAME' to run on this Mac."
echo "  (macOS blocks unsigned apps by default.)"
echo ""
read -n 1 -s -r -p "  Press any key to continue..."
echo ""
echo ""

# Remove quarantine flag
xattr -cr "$APP_PATH" 2>/dev/null

# Ad-hoc sign
codesign --force --deep --sign - "$APP_PATH" 2>/dev/null

echo "  ✔ Done! '$APP_NAME' is now ready to use."
echo ""
echo "  How to use:"
echo "    • Drag and drop .edl files onto the app to convert"
echo "    • Double-click the app to change preferences (CSV or XLSX)"
echo ""
read -n 1 -s -r -p "  Press any key to close..."
