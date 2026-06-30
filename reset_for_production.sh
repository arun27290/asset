#!/bin/bash
# ═══════════════════════════════════════════════════════════════
#  Server Asset Inventory — Reset to clean state before go-live
#  Run this ONCE after testing, before sharing the app with users.
#  It deletes all test data: users, tags, notes, flags, corrections,
#  decommissions, new server suggestions, uploaded inventory, and
#  archived snapshots. The app code itself (asset_inventory.py) is
#  never touched.
# ═══════════════════════════════════════════════════════════════

set -e
cd "$(dirname "$0")"

echo "This will permanently delete all test data:"
echo "  - All user accounts (you'll get a new admin password on next run)"
echo "  - All tags, notes, flags"
echo "  - All correction / decommission / new-server submissions"
echo "  - The uploaded inventory file and its archive history"
echo "  - The audit log"
echo ""
read -p "Type YES to confirm: " confirm
if [ "$confirm" != "YES" ]; then
    echo "Cancelled. Nothing was deleted."
    exit 0
fi

# User accounts + session key — forces fresh admin password on next run
rm -f users.json .flask_secret

# Activity data
rm -f tags.json notes.json flags.json
rm -f corrections.json decommissions.json new_servers.json

# Uploaded inventory + its metadata
rm -f current_inventory.xlsx current_inventory_meta.json
rm -f prev_inventory.xlsx prev_inventory_meta.json

# Archive of past uploads
rm -rf archive

# Audit trail
rm -f audit.log

# Python cache (harmless but tidy)
rm -rf __pycache__

echo ""
echo "✅ Reset complete. Folder is now clean."
echo "   Run 'python asset_inventory.py' to start fresh —"
echo "   a new one-time admin password will be printed to the console."
