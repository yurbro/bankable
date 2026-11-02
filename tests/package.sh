#!/usr/bin/env bash
# 打包成 claude.ai 可上传的 bankable.skill（只含 SKILL.md、references、scripts、assets）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="${1:-$ROOT/bankable.skill}"
TMP="$(mktemp -d)"
mkdir -p "$TMP/bankable"
cp -r "$ROOT/SKILL.md" "$ROOT/references" "$ROOT/scripts" "$ROOT/assets" "$TMP/bankable/"
find "$TMP" -name "__pycache__" -type d -prune -exec rm -rf {} +
rm -f "$OUT"
(cd "$TMP" && zip -qr "$OUT" bankable)
rm -rf "$TMP"
echo "[OK] $OUT ($(du -h "$OUT" | cut -f1))"
