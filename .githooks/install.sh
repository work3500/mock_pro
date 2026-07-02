#!/usr/bin/env bash
#
# KTNET Git 훅 설치 스크립트
# 저장소가 .githooks 디렉터리를 훅 경로로 사용하도록 설정합니다.
#
# 사용법:  ./.githooks/install.sh
#
set -euo pipefail

repo_root=$(git rev-parse --show-toplevel)
cd "$repo_root"

git config core.hooksPath .githooks
chmod +x .githooks/pre-commit 2>/dev/null || true

echo "✔ core.hooksPath 를 '.githooks' 로 설정했습니다."
echo "  이제 커밋 시 .githooks/pre-commit 훅이 실행됩니다."
