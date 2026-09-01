#!/usr/bin/env bash
#
# SheetHub -> Sheetlux CD GitOps Sync Bridge
#
# Pulls declarative manifests from the SheetHub `Files` tab into a local directory
# (e.g. ./gitops) for Sheetlux CD to reconcile into the Sheeternetes cluster.
#
# Usage:
#   ./sheetlux-sync.sh [repo] [target_dir] [manifest_name]
#   ./sheetlux-sync.sh sncf/hello-web ./gitops app.json
#
set -euo pipefail

here="$(cd "$(dirname "$0")" && pwd)"
REPO="${1:-sncf/hello-web}"
DIR="${2:-$here/gitops}"
MANIFEST="${3:-app.json}"
SHEETHUB_BIN="$here/sheethub"

mkdir -p "$DIR"

echo "[sheetlux-sync] Fetching manifest '$MANIFEST' from SheetHub repo: $REPO..."
content="$("$SHEETHUB_BIN" file get "$REPO" "$MANIFEST" 2>/dev/null || true)"

if [ -n "$content" ] && [ "$content" != "null" ]; then
  echo "$content" > "$DIR/$MANIFEST"
  echo "[sheetlux-sync] Synced $MANIFEST from SheetHub -> $DIR/$MANIFEST"
  echo "[sheetlux-sync] Ready for Sheetlux CD to reconcile into Sheeternetes."
else
  echo "[sheetlux-sync] Warning: manifest $MANIFEST not found in SheetHub repo $REPO."
fi
