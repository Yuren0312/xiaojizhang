# -*- coding: utf-8 -*-
"""从「输出 - 5月会计报表.xlsx」反推版式，生成：

  accounting_report/templates/报表模板.xlsx   —— 清空数值后的空白模板
  accounting_report/rules/layout.yaml         —— 行次 ↔ 单元格坐标

这样版式（合并区、列宽、字体、边框、数字格式）与现有报表完全一致，
也不必手抄一百多行坐标。只需运行一次；样表换了再跑一次即可。

    python tools/build_template.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import yaml

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from accounting_report.config import (  # noqa: E402
    SHEET_BS, SHEET_IS, SHEET_CF, SHEET_IS_Q, SHEET_CF_Q, RULES_DIR, TEMPLATE_PATH,
)

SAMPLE = ROOT / "输出 - 5月会计报表.xlsx"

# 每张表：行次列 → 值列映射；date_cell / date_kind；空值写法
SPEC = {
    SHEET_BS: {
        "date_cell": "E3", "date_kind": "month_end", "company_cell": "A3", "blank": "",
        "groups": [
            {"line_col": "B", "cols": {"end": "C", "year_start": "D"}},
            {"line_col": "F", "cols": {"end": "G", "year_start": "H"}},
        ],
    },
    SHEET_IS: {
        "date_cell": "B2", "date_kind": "month", "company_cell": "A2", "blank": "",
        "groups": [{"line_col": "B", "cols": {"ytd": "C", "month": "D"}}],
    },
    SHEET_CF: {
        "date_cell": "B3", "date_kind": "month", "company_cell": "A3", "blank": "",
        "groups": [{"line_col": "B", "cols": {"ytd": "C", "month": "D"}}],
    },
    SHEET_IS_Q: {
        "date_cell": "B2", "date_kind": "month", "company_cell": "A2", "blank": "",
        "groups": [{"line_col": "B",
                    "cols": {"ytd": "C", "q1": "D", "q2": "E", "q3": "F", "q4": "G"}}],
    },
    # 现金流量表季报在样表里把 0 显式写成 0.00，其余表留空——照搬，保证逐格一致
    SHEET_CF_Q: {
        "date_cell": "B3", "date_kind": "month", "company_cell": "A3", "blank": 0.0,
        "groups": [{"line_col": "B",
                    "cols": {"ytd": "C", "q1": "D", "q2": "E", "q3": "F", "q4": "G"}}],
    },
}

# 模板里不留具体公司名，出表时由 config.COMPANY_NAME 填入
COMPANY_PREFIX = "企业名称: "


def build() -> dict:
    wb = openpyxl.load_workbook(SAMPLE)
    layout: dict = {}

    for sheet, spec in SPEC.items():
        ws = wb[sheet]
        lines: dict[int, dict] = {}

        for group in spec["groups"]:
            lcol = group["line_col"]
            for row in range(1, ws.max_row + 1):
                raw = ws[f"{lcol}{row}"].value
                if not isinstance(raw, str) or not raw.strip().isdigit():
                    continue
                lineno = int(raw.strip())
                if lineno in lines:
                    raise ValueError(f"{sheet} 行次 {lineno} 重复出现（第 {row} 行）")
                entry = {"row": row}
                entry.update(group["cols"])
                lines[lineno] = entry
                # 清空该行的数值格
                for col in group["cols"].values():
                    ws[f"{col}{row}"] = spec["blank"]

        layout[sheet] = {
            "date_cell": spec["date_cell"],
            "date_kind": spec["date_kind"],
            "company_cell": spec["company_cell"],
            "blank": spec["blank"],
            "lines": {k: lines[k] for k in sorted(lines)},
        }
        # 清空日期格，公司名也从模板里抹掉（出表时按配置填）
        ws[spec["date_cell"]] = ""
        ws[spec["company_cell"]] = COMPANY_PREFIX

    TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(TEMPLATE_PATH)
    return layout


def main() -> int:
    if not SAMPLE.exists():
        print(f"找不到样表：{SAMPLE}", file=sys.stderr)
        return 1

    layout = build()

    # 校验行次连续完整
    expected = {SHEET_BS: 53, SHEET_IS: 32, SHEET_CF: 22, SHEET_IS_Q: 32, SHEET_CF_Q: 22}
    ok = True
    for sheet, n in expected.items():
        got = set(layout[sheet]["lines"])
        missing = set(range(1, n + 1)) - got
        extra = got - set(range(1, n + 1))
        status = "OK" if not missing and not extra else "异常"
        if missing or extra:
            ok = False
        print(f"  {sheet:<14} 行次 {len(got):>2} 个  {status}"
              + (f"  缺失={sorted(missing)}" if missing else "")
              + (f"  多余={sorted(extra)}" if extra else ""))

    out = RULES_DIR / "layout.yaml"
    header = (
        "# 由 tools/build_template.py 从「输出 - 5月会计报表.xlsx」自动生成，请勿手改。\n"
        "# 行次 → {row: Excel行号, <列名>: Excel列字母}\n"
    )
    with open(out, "w", encoding="utf-8") as f:
        f.write(header)
        yaml.safe_dump(layout, f, allow_unicode=True, sort_keys=False, default_flow_style=False)

    print(f"\n模板  → {TEMPLATE_PATH}")
    print(f"版式  → {out}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
