#!/usr/bin/env bash
# Bankable 脚本冒烟测试：用示例假设跑通 模型 → 重算三情景 → 敏感性 → 图表 → Word → 一致性核对。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="${1:-$(mktemp -d)}"
mkdir -p "$WORK"
echo "[smoke] work dir: $WORK"
python "$ROOT/scripts/build_financial_model.py" "$ROOT/assets/assumptions-example.json" -o "$WORK/model.xlsx"
python "$ROOT/scripts/recalc_and_extract.py" "$WORK/model.xlsx" -o "$WORK/model_outputs.json" --assumptions "$ROOT/assets/assumptions-example.json"
python - "$WORK/model_outputs.json" <<'PY'
import json, sys
o = json.load(open(sys.argv[1], encoding="utf-8"))
assert o["checks"]["balance_sheet_balances"], "balance sheet does not balance"
assert set(o["scenarios"]) == {"base", "downside", "upside"}, o["scenarios"].keys()
assert o["scenarios"]["base"]["summary"]["irr_with_tv"] is not None
print("[smoke] model OK:", o["narrative_numbers"]["revenue_by_year"][0], "→", o["narrative_numbers"]["revenue_by_year"][-1])
PY
python "$ROOT/scripts/run_sensitivity.py" "$ROOT/assets/assumptions-example.json" -o "$WORK/sensitivity_outputs.json" --md "$WORK/sensitivity.md" --chart "$WORK/charts/sensitivity.png" > /dev/null
python "$ROOT/scripts/make_charts.py" "$WORK/model_outputs.json" -o "$WORK/charts" > /dev/null
cp "$ROOT/examples/示例-plan（仅3章，测试排版用）.md" "$WORK/plan.md"
python "$ROOT/scripts/build_docx.py" "$WORK/plan.md" -o "$WORK/plan.docx" --title "冒烟测试" --org "测试" --date "2026年11月" > /dev/null
python "$ROOT/scripts/build_docx.py" "$WORK/plan.md" -o "$WORK/plan_stripped.docx" --title "冒烟测试" --strip-refs > /dev/null
python "$ROOT/scripts/build_pptx.py" "$ROOT/examples/示例-deck（3页）.md" -o "$WORK/deck.pptx" > /dev/null
# 示例 plan 的数字对应示例假设；核对脚本退出码 0 表示必提数字齐全
python "$ROOT/scripts/check_consistency.py" "$WORK/plan.md" "$WORK/model_outputs.json" -o "$WORK/consistency_report.md" --extra "$WORK/sensitivity_outputs.json" > /dev/null || { echo "[smoke] consistency check reported missing numbers (see $WORK/consistency_report.md)"; exit 1; }
ls -1 "$WORK"
echo "[smoke] all steps passed"
