#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="$ROOT_DIR/dist"
OUT_ZIP="$DIST_DIR/kanban-for-outlook.zip"

mkdir -p "$DIST_DIR"
rm -f "$OUT_ZIP"

cd "$ROOT_DIR"

# Package the app for local installation.
# Excludes git metadata and previously built archives.
zip -r "$OUT_ZIP" \
  ACKNOWLEDGEMENTS.md \
  CHANGELOG.md \
  DISCLAIMER.md \
  LICENSE \
  PRIVACY.md \
  README.md \
  START_HERE.html \
  ROADMAP.md \
  SECURITY.md \
  THIRD_PARTY_NOTICES.md \
  docs \
  css \
  fonts \
  js \
  kanban.html \
  vendor \
  themes \
  install.cmd \
  install-local.cmd \
  uninstall.cmd \
  upgrade.html \
  whatsnew.html \
  -x "*.zip" \
  -x ".git/*" \
  -x ".github/*" \
  -x "dist/*"

sha256sum "$OUT_ZIP" > "$OUT_ZIP.sha256"

echo "Created: $OUT_ZIP"
echo "Created: $OUT_ZIP.sha256"
