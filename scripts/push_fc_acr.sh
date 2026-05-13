#!/usr/bin/env bash
set -euo pipefail

LOCAL_IMAGE="${1:-talent-ninebox-fc:latest}"
REMOTE_IMAGE="${2:-}"

if [[ -z "$REMOTE_IMAGE" ]]; then
  echo "Usage: bash scripts/push_fc_acr.sh talent-ninebox-fc:latest registry.cn-hangzhou.aliyuncs.com/<namespace>/<repo>:latest"
  exit 1
fi

docker tag "$LOCAL_IMAGE" "$REMOTE_IMAGE"
docker push "$REMOTE_IMAGE"

echo "FC image pushed:"
echo "$REMOTE_IMAGE"
