# -*- coding: utf-8 -*-
"""生成税务底稿所需的示例数据源（体验官白班轮新增）

解决「输出6 所得税预缴 / 输出7 个税扣缴」长期停留在模板级输出的体验问题：
企业只需一次性维护资产台账与员工名册（静态配置，跨期复用），每月提供工资表，
三个底稿即可从「[模板] 输出」升级为「正式输出」。

用法：
    python tools/generate_ledger_samples.py [输出根目录]

默认输出到 ./输入 目录（工作副本），不会触碰原始资料目录。

生成文件：
    输入/企业信息/资产台账.xlsx     （Sheet「资产台账」）
    输入/企业信息/员工名册.xlsx     （Sheet「员工名册」）
    输入/{year}-{month}/工资社保/工资表.xlsx （Sheet「工资表」）
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)


def _style_header(ws, headers):
    for j, h in enumerate(headers, 1):
        c = ws.cell(1, j, h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
    for col, w in zip("ABCDEFGHIJKLMN", (12, 20, 12, 12, 12, 12, 16, 14, 12, 12, 12, 12, 12, 12)):
        ws.column_dimensions[col].width = w


def write_asset_ledger(out: Path) -> Path:
    """资产台账：静态配置，跨期复用"""
    ws_name = "资产台账"
    wb = Workbook()
    ws = wb.active
    ws.title = ws_name
    headers = ["资产类别", "资产名称", "购置日期", "原值", "累计折旧", "净值", "备注"]
    _style_header(ws, headers)
    rows = [
        # 小微企业示例：办公设备（净值合计 46,000）
        ("电子设备", "笔记本电脑", "2025-06-15", "12000.00", "6000.00", "6000.00", "办公用"),
        ("电子设备", "台式电脑",   "2025-06-15", "8000.00",  "4000.00", "4000.00", "办公用"),
        ("办公设备", "打印机",     "2025-08-01", "3000.00",  "1500.00", "1500.00", ""),
        ("办公设备", "办公桌椅",   "2025-08-01", "10000.00", "5000.00", "5000.00", ""),
        ("电子设备", "手机",       "2026-01-10", "10000.00", "2000.00", "8000.00", ""),
        ("电子设备", "服务器",     "2026-03-01", "30000.00", "4500.00", "25500.00", "账务数据服务器"),
        ("运输工具", "电动车",     "2026-04-01", "6000.00",  "1000.00", "5000.00", ""),
        ("其他",     "办公用品",   "2026-05-01", "2000.00",  "500.00",  "1500.00", ""),
    ]
    for r in rows:
        ws.append(r)
    # 合计行
    ws.append(("合计", "", "", "81000.00", "24500.00", "56500.00", "净值=资产总额参考值"))
    wb.save(out)
    return out


def write_employee_roster(out: Path) -> Path:
    """员工名册：静态配置，跨期复用"""
    wb = Workbook()
    ws = wb.active
    ws.title = "员工名册"
    headers = ["姓名", "身份证号", "入职日期", "岗位", "社保基数", "公积金基数", "备注"]
    _style_header(ws, headers)
    rows = [
        ("张伟", "110101199001011234", "2025-11-01", "负责人", "12000.00", "12000.00", "法定代表人"),
        ("李娜", "110101199203054567", "2025-11-01", "会计",   "8000.00",  "8000.00",  "财务岗"),
        ("王芳", "110101199505067890", "2026-01-01", "行政",   "7000.00",  "7000.00",  ""),
    ]
    for r in rows:
        ws.append(r)
    wb.save(out)
    return out


def write_payroll(out: Path, year: int, month: int) -> Path:
    """当月工资表（Sheet「工资表」）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "工资表"
    headers = ["姓名", "身份证号", "应发工资", "个人养老保险", "个人医疗保险", "个人失业保险",
               "个人公积金", "企业年金(个人)", "子女教育", "继续教育", "住房贷款利息",
               "住房租金", "赡养老人", "婴幼儿照护"]
    _style_header(ws, headers)
    # 3 名员工示例（个人五险一金按比例，专项附加扣除示例）
    rows = [
        ("张伟", "110101199001011234", "15000.00", "960.00",  "240.00", "30.00",
         "1800.00", "0.00", "0.00", "0.00", "0.00", "1500.00", "2000.00", "0.00"),
        ("李娜", "110101199203054567", "10000.00", "640.00",  "160.00", "20.00",
         "1200.00", "0.00", "1000.00", "0.00", "0.00", "0.00", "0.00", "0.00"),
        ("王芳", "110101199505067890", "8000.00",  "512.00",  "128.00", "16.00",
         "960.00", "0.00", "0.00", "400.00", "0.00", "0.00", "1000.00", "0.00"),
    ]
    for r in rows:
        ws.append(r)
    wb.save(out)
    return out


def main() -> None:
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("输入")
    root = Path(root)
    company = root / "企业信息"
    company.mkdir(parents=True, exist_ok=True)

    p1 = write_asset_ledger(company / "资产台账.xlsx")
    p2 = write_employee_roster(company / "员工名册.xlsx")

    # 工资表：生成当期月份（默认取当前月份；也支持显式年份月份参数）
    import datetime
    now = datetime.date.today()
    year, month = now.year, now.month
    if len(sys.argv) >= 4:
        year, month = int(sys.argv[2]), int(sys.argv[3])
    period_dir = root / f"{year:04d}-{month:02d}" / "工资社保"
    period_dir.mkdir(parents=True, exist_ok=True)
    p3 = write_payroll(period_dir / "工资表.xlsx", year, month)

    print(f"✅ 资产台账: {p1}")
    print(f"✅ 员工名册: {p2}")
    print(f"✅ 工资表:   {p3}  ({year}-{month:02d})")
    print("重新运行 run_chain 后，输出6/输出7 将升为正式输出。")


if __name__ == "__main__":
    main()
