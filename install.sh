#!/bin/sh
# Headway macOS installer — per-user, no admin required.
# Usage: curl -fsSL https://raw.githubusercontent.com/smo-key/headway/main/install.sh | sh
set -eu

REPO="smo-key/headway"
DEST="$HOME/Applications"

echo "Fetching the latest Headway release…"
URL=$(curl -fsSL "https://api.github.com/repos/$REPO/releases/latest" |
  grep -oE '"browser_download_url": *"[^"]+\.app\.tar\.gz"' |
  head -1 | sed 's/.*"\(https[^"]*\)"/\1/')

if [ -z "$URL" ]; then
  echo "Could not find a macOS build in the latest release of $REPO." >&2
  exit 1
fi

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

echo "Downloading $(basename "$URL")…"
curl -fL --progress-bar "$URL" -o "$TMP/headway.app.tar.gz"

mkdir -p "$DEST"
rm -rf "$DEST/Headway.app"
tar -xzf "$TMP/headway.app.tar.gz" -C "$DEST"

# unsigned build: clear the quarantine flag so Gatekeeper lets it open
xattr -dr com.apple.quarantine "$DEST/Headway.app" 2>/dev/null || true

echo "Installed $DEST/Headway.app"
open "$DEST/Headway.app" 2>/dev/null || true
