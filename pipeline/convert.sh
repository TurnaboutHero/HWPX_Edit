#!/usr/bin/env bash
# HWPX ↔ Markdown 변환 파이프라인 래퍼
# 사용법:
#   ./convert.sh to-md  문서.hwpx [-o output.md]
#   ./convert.sh to-hwpx 문서.md [-o output.hwpx] [-r 원본.hwpx]
#   ./convert.sh smart  원본.hwpx 편집된.md [-o output.hwpx]

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python "$SCRIPT_DIR/convert.py" "$@"
