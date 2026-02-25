#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="$ROOT_DIR/dist"
OUT_ZIP="$DIST_DIR/kanban-for-outlook.zip"

STAGE_DIR="$DIST_DIR/stage"
PKG_NAME="kanban-for-outlook"
PKG_DIR="$STAGE_DIR/$PKG_NAME"

mkdir -p "$DIST_DIR"
rm -f "$OUT_ZIP"

rm -rf "$STAGE_DIR"
mkdir -p "$PKG_DIR"

# Stage a clean, user-friendly release folder.
cp "$ROOT_DIR/START_HERE.html" "$PKG_DIR/"
cp "$ROOT_DIR/kanban.html" "$PKG_DIR/"
cp "$ROOT_DIR/upgrade.html" "$PKG_DIR/"
cp "$ROOT_DIR/whatsnew.html" "$PKG_DIR/"

cp "$ROOT_DIR/install.cmd" "$PKG_DIR/"
cp "$ROOT_DIR/install-local.cmd" "$PKG_DIR/"
cp "$ROOT_DIR/uninstall.cmd" "$PKG_DIR/"

cp "$ROOT_DIR/LICENSE" "$PKG_DIR/"
cp "$ROOT_DIR/THIRD_PARTY_NOTICES.md" "$PKG_DIR/"

cp -R "$ROOT_DIR/css" "$PKG_DIR/"
cp -R "$ROOT_DIR/fonts" "$PKG_DIR/"
cp -R "$ROOT_DIR/js" "$PKG_DIR/"
cp -R "$ROOT_DIR/vendor" "$PKG_DIR/"
cp -R "$ROOT_DIR/themes" "$PKG_DIR/"

mkdir -p "$PKG_DIR/docs"
cp "$ROOT_DIR/docs/README.md" "$PKG_DIR/docs/"
cp "$ROOT_DIR/docs/SETUP.md" "$PKG_DIR/docs/"
cp "$ROOT_DIR/docs/USAGE.md" "$PKG_DIR/docs/"
cp "$ROOT_DIR/docs/THEMES.md" "$PKG_DIR/docs/"
cp "$ROOT_DIR/docs/ACCESSIBILITY.md" "$PKG_DIR/docs/"

# Create zip with a single top-level folder.
(cd "$STAGE_DIR" && zip -r "$OUT_ZIP" "$PKG_NAME" -x "*.zip")

rm -rf "$STAGE_DIR"

sha256sum "$OUT_ZIP" > "$OUT_ZIP.sha256"

echo "Created: $OUT_ZIP"
echo "Created: $OUT_ZIP.sha256"
