#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
BUILD_DIR="$ROOT_DIR/fc-deploy/build"
ZIP_PATH="$ROOT_DIR/fc-deploy/talent-ninebox-fc.zip"

rm -rf "$ROOT_DIR/fc-deploy"
mkdir -p "$BUILD_DIR"

rsync -a \
  --exclude ".venv/" \
  --exclude ".pytest_cache/" \
  --exclude "__pycache__/" \
  --exclude "*.pyc" \
  --exclude ".vercel/" \
  --exclude "tmp/" \
  --exclude "fc-deploy/" \
  --exclude "tests/" \
  "$ROOT_DIR/" "$BUILD_DIR/"

cp "$ROOT_DIR/fc/bootstrap" "$BUILD_DIR/bootstrap"
chmod +x "$BUILD_DIR/bootstrap"

python3 -m pip install \
  --target "$BUILD_DIR" \
  --no-cache-dir \
  -r "$ROOT_DIR/requirements.txt"

(
  cd "$BUILD_DIR"
  zip -qr "$ZIP_PATH" .
)

echo "FC deployment package created:"
echo "$ZIP_PATH"
