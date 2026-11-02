# 示例（虚构案例）

- `示例-财务模型（虚构案例）.xlsx`：由 `assets/assumptions-example.json` 生成并重算后的三表模型，含情景对比与第 1 年月度现金流。
- `示例-plan（仅3章，测试排版用）.md` 与 `示例-Word排版样例（仅3章）.docx`：演示 plan.md 的写法（标题层级、表题、图、引用框、分页）及 `build_docx.py` 的排版效果；末尾故意放了一个与模型不一致的数字，用于演示 `check_consistency.py` 的提示。
- `示例-deck（3页）.md`：`build_pptx.py` 的输入格式示例。

`tests/smoke_test.sh` 用这些文件跑通整条脚本流水线。
