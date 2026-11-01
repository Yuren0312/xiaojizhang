#!/usr/bin/env python3
"""
小虾记账 — 示例数据生成工具

用法:
    python tools/generate_sample_data.py              # 生成完整示例数据到 输入/
    python tools/generate_sample_data.py --clean      # 清理后重新生成
    python tools/generate_sample_data.py --period 2026-08  # 指定账期

数据说明:
    - 所有数据均为虚构，用于验证系统全管线是否正常工作
    - 生成后将覆盖已有同名文件
    - 企业信息（资产台账/员工名册）生成在 输入/企业信息/，跨账期复用
    - 其他数据生成在 输入/{period}/ 对应模块子目录中
"""
import argparse
import os
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
INPUT = ROOT / "输入"
COMPANY_INFO = INPUT / "企业信息"


def _money_fmt(ws_obj, col_letter, start_row, end_row):
    """为金额列设置千分位两位小数格式"""
    from openpyxl.styles import numbers
    for r in range(start_row, end_row + 1):
        ws_obj[f"{col_letter}{r}"].number_format = '#,##0.00'


def _add_header_style(ws_obj, row=1, max_col=14):
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for c in range(1, max_col + 1):
        cell = ws_obj.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill


def generate_invoice_ledger(period: str) -> Path:
    """发票台账 → 输入/{period}/收入发票/发票台账.xlsx"""
    dest_dir = INPUT / period / "收入发票"
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / "发票台账.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发票台账"
    ws.append(["开票日期", "发票代码", "发票号码", "发票类型", "购方名称",
               "货物服务名称", "不含税金额", "税额", "含税金额", "备注"])

    invoices = [
        ("2026-08-03", "", "A0000001", "专用",   "客户A", "软件开发服务",   60000, 600, 60600, ""),
        ("2026-08-05", "", "A0000002", "普通",   "客户B", "技术咨询服务",   30000, 300, 30300, ""),
        ("2026-08-10", "", "A0000003", "电子普通", "客户C", "UI界面设计",    50000, 500, 50500, ""),
        ("2026-08-15", "", "A0000004", "专用",   "客户D", "系统集成开发",   80000, 800, 80800, ""),
        ("2026-08-20", "", "A0000005", "普通",   "客户E", "职业技能培训",   20000,   0, 20000, "免税-教育培训"),
    ]
    for inv in invoices:
        ws.append(inv)

    _add_header_style(ws, max_col=10)
    _money_fmt(ws, "G", 2, len(invoices) + 1)
    _money_fmt(ws, "H", 2, len(invoices) + 1)
    _money_fmt(ws, "I", 2, len(invoices) + 1)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 16
    for c in "GHI":
        ws.column_dimensions[c].width = 14

    wb.save(path)
    return path


def generate_payroll(period: str) -> Path:
    """工资表 → 输入/{period}/工资社保/工资表.xlsx"""
    dest_dir = INPUT / period / "工资社保"
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / "工资表.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工资表"
    ws.append(["姓名", "身份证号", "应发工资", "个人养老保险", "个人医疗保险",
               "个人失业保险", "个人公积金", "企业年金(个人)", "子女教育",
               "继续教育", "住房贷款利息", "住房租金", "赡养老人", "婴幼儿照护"])

    employees = [
        ("张三", "110101199001010001", 15000, 1200, 300, 75, 1800, 0, 1000, 0, 1000, 0, 1000, 0),
        ("李四", "110101199202020002", 12000,  960, 240, 60, 1440, 0,    0, 400, 0, 1500, 1000, 0),
        ("王五", "110101198503030003", 20000, 1600, 400, 100, 2400, 0, 1000, 0, 0, 0, 2000, 1000),
        ("赵六", "110101199504040004",  8000,  640, 160, 40,  960, 0,    0, 0, 0, 0, 0, 0),
    ]
    for emp in employees:
        ws.append(emp)

    _add_header_style(ws, max_col=14)
    for c in "CDEFGHIJKLMN":
        _money_fmt(ws, c, 2, len(employees) + 1)

    wb.save(path)
    return path


def generate_asset_ledger() -> Path:
    """资产台账 → 输入/企业信息/资产台账.xlsx"""
    COMPANY_INFO.mkdir(parents=True, exist_ok=True)
    path = COMPANY_INFO / "资产台账.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "资产台账"
    ws.append(["资产名称", "类别", "原值", "累计折旧", "净值", "购置日期"])

    assets = [
        ("笔记本电脑 ThinkPad", "电子设备",   8000,  2000,  6000, "2025-06-01"),
        ("开发服务器",       "电子设备",   20000,  5000, 15000, "2025-08-01"),
        ("办公家具一套",      "办公设备",   12000,  3000,  9000, "2025-09-01"),
        ("软件著作权",       "无形资产",   150000,     0, 150000, "2026-01-01"),
        ("显示器×3",        "电子设备",    9000,  2000,  7000, "2025-10-01"),
    ]
    for a in assets:
        ws.append(a)

    _add_header_style(ws, max_col=6)
    for c in "CDE":
        _money_fmt(ws, c, 2, len(assets) + 1)

    wb.save(path)
    return path


def generate_employee_roster() -> Path:
    """员工名册 → 输入/企业信息/员工名册.xlsx"""
    COMPANY_INFO.mkdir(parents=True, exist_ok=True)
    path = COMPANY_INFO / "员工名册.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "员工名册"
    ws.append(["姓名", "身份证号", "入职日期", "部门", "岗位"])

    staff = [
        ("张三", "110101199001010001", "2025-06-01", "技术部", "高级开发工程师"),
        ("李四", "110101199202020002", "2025-08-01", "技术部", "前端开发工程师"),
        ("王五", "110101198503030003", "2025-06-01", "管理部", "项目经理"),
        ("赵六", "110101199504040004", "2026-01-01", "运营部", "运营专员"),
    ]
    for s in staff:
        ws.append(s)

    _add_header_style(ws, max_col=5)
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["E"].width = 18

    wb.save(path)
    return path


def generate_business_events(period: str) -> Path:
    """业务事项 → 输入/业务事项.xlsx

    列: 日期 / 摘要 / 借方科目 / 贷方科目 / 金额 / 来源
    这是人工录入的当月业务，需要按借贷对录入。
    收入类: 借银行存款，贷营业收入
    费用类: 借费用科目，贷银行存款
    """
    path = INPUT / "业务事项.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "业务事项"
    ws.append(["日期", "摘要", "借方科目", "贷方科目", "金额", "来源"])

    events = [
        # 收入 — 借银行 贷收入
        (f"{period}-03", "技术咨询收入",  "银行存款_中国工商银行", "营业收入",           30300, "业务回单/回单1.pdf"),
        (f"{period}-05", "UI设计收入",    "银行存款_中国工商银行", "营业收入",           50500, "业务回单/回单2.pdf"),
        (f"{period}-10", "系统开发收入",  "银行存款_中国工商银行", "营业收入",           80800, "业务回单/回单3.pdf"),
        # 费用 — 借费用 贷银行
        (f"{period}-01", "阿里云服务费",  "营业成本",            "银行存款_中国工商银行",  1500, "发票/阿里云.pdf"),
        (f"{period}-01", "域名续费",      "管理费用",            "银行存款_中国工商银行",   500, "发票/域名.pdf"),
        (f"{period}-01", "银行手续费",    "财务费用",            "银行存款_中国工商银行",   200, "银行对账单"),
        (f"{period}-05", "发放工资",      "应付职工薪酬",        "银行存款_中国工商银行", 10000, "工资表"),
        (f"{period}-05", "社保公积金代缴","应付职工薪酬",        "银行存款_中国工商银行",  3000, "工资表"),
        (f"{period}-10", "办公用品采购",  "管理费用",            "银行存款_中国工商银行",   500, "发票/办公用品.pdf"),
    ]
    for ev in events:
        ws.append(ev)

    _add_header_style(ws, max_col=6)
    _money_fmt(ws, "E", 2, len(events) + 1)
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["F"].width = 24

    wb.save(path)
    return path


def generate_input_readme() -> Path:
    """输入目录说明文件"""
    path = INPUT / "README.md"
    content = """# 输入数据目录

> 本目录存放企业每月交付的原始业务资料。
> **银行账号、回单验证码等敏感数据请勿提交到版本库。**

## 目录结构

```
输入/
├── README.md                    # 本文件
├── 业务事项.xlsx                # 当月人工录入的分录（日期/摘要/借方科目/贷方科目/金额/来源）
├── 企业信息/                    # 跨账期复用的基础信息
│   ├── 资产台账.xlsx            # 固定资产清单（名称/类别/原值/累计折旧/净值/购置日期）
│   └── 员工名册.xlsx            # 在册员工（姓名/身份证号/入职日期/部门/岗位）
├── 2026-08/                     # 各月资料，按 YYYY-MM 命名
│   ├── 收入发票/
│   │   └── 发票台账.xlsx        # 当月开票记录（开票日期/发票号码/发票类型/不含税金额/税额/备注）
│   ├── 工资社保/
│   │   └── 工资表.xlsx          # 当月工资+社保公积金+专项附加扣除
│   ├── 银行对账单/
│   │   └── 2026.08对账单.pdf    # 工行企业网银导出的PDF/XLS对账单
│   └── 业务回单/
│       └── 回单*.pdf            # 银行电子回单 PDF
└── 票据记录.csv                 # （可选）OCR识别导出的票据记录
```

## 首次使用

运行以下命令生成示例数据，了解各文件格式：

```bash
python tools/generate_sample_data.py
```

然后将示例数据替换为真实企业数据即可开始试运行。
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return path


def clean_sample_data():
    """清理示例数据但不删除已有真实数据"""
    print("清理已有示例数据...")
    paths_to_check = [
        INPUT / "业务事项.xlsx",
        INPUT / "README.md",
        COMPANY_INFO / "资产台账.xlsx",
        COMPANY_INFO / "员工名册.xlsx",
    ]
    for p in paths_to_check:
        if p.exists():
            p.unlink()
            print(f"  removed: {p}")

    # 按账期目录清理
    for d in INPUT.iterdir():
        if d.is_dir() and d.name != "企业信息" and not d.name.startswith("."):
            # 只删示例文件，保留银行对账单等大文件
            target = d / "收入发票" / "发票台账.xlsx"
            if target.exists():
                target.unlink()
                print(f"  removed: {target}")
            target = d / "工资社保" / "工资表.xlsx"
            if target.exists():
                target.unlink()
                print(f"  removed: {target}")


def main():
    parser = argparse.ArgumentParser(description="小虾记账 — 示例数据生成")
    parser.add_argument("--clean", action="store_true", help="清理旧数据后重新生成")
    parser.add_argument("--period", default="2026-08", help="账期 (默认 2026-08)")
    args = parser.parse_args()

    if args.clean:
        clean_sample_data()

    period = args.period

    print(f"生成示例数据 — 账期 {period}")
    print("=" * 50)

    files = []

    p = generate_invoice_ledger(period)
    files.append(p)
    print(f"  ✅ 发票台账 (5 张) → {p.relative_to(ROOT)}")

    p = generate_payroll(period)
    files.append(p)
    print(f"  ✅ 工资表 (4 人) → {p.relative_to(ROOT)}")

    p = generate_asset_ledger()
    files.append(p)
    print(f"  ✅ 资产台账 (5 项) → {p.relative_to(ROOT)}")

    p = generate_employee_roster()
    files.append(p)
    print(f"  ✅ 员工名册 (4 人) → {p.relative_to(ROOT)}")

    p = generate_business_events(period)
    files.append(p)
    print(f"  ✅ 业务事项 (9 条) → {p.relative_to(ROOT)}")

    p = generate_input_readme()
    files.append(p)
    print(f"  ✅ 输入目录说明 → {p.relative_to(ROOT)}")

    print(f"\n共生成 {len(files)} 个文件")
    print(f"\n⚠️  以上为虚构测试数据，请替换为真实企业数据后开始试运行。")
    print(f"    python -m accounting_report --full {period} --input 输入 --output 输出")


if __name__ == "__main__":
    main()
