#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Recon Daily - macOS/Linux launcher (call this from cron)
# Edit RECON_ROOT to your share folder, then this runs today's reconciliation.
# ---------------------------------------------------------------------------
set -euo pipefail

# Path to the share folder that holds the dated sub-folders (YYYY-MM-DD).
# For a synced SharePoint/OneDrive library this is a normal local path, e.g.:
#   RECON_ROOT="$HOME/Hubpay/Recon - Documents/Daily"
RECON_ROOT="/path/to/your/share/folder"

# Days to look back when aging a break (default 10).
RECON_LOOKBACK="${RECON_LOOKBACK:-10}"

cd "$(dirname "$0")"
python3 -m recon --root "$RECON_ROOT" --lookback "$RECON_LOOKBACK"
