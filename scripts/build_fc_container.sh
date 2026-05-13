#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
IMAGE_NAME="${1:-talent-ninebox-fc:latest}"

docker build \
  --platform linux/amd64 \
  -f "$ROOT_DIR/fc/Dockerfile" \
  -t "$IMAGE_NAME" \
  "$ROOT_DIR"

echo "FC container image built:"
echo "$IMAGE_NAME"
