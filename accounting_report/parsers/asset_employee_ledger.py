# -*- coding: utf-8 -*-
"""资产台账与员工名册解析器（静态配置级，跨期复用）

两个 parser 都读一个配置格式的文件，企业一次性维护，无需每月重新上传。

资产台账格式（Sheet「资产台账」）：
| 资产类别 | 资产名称 | 购置日期 | 原值 | 累计折旧 | 净值 | 备注 |

员工名册格式（Sheet「员工名册」）：
| 姓名 | 身份证号 | 入职日期 | 岗位 | 社保基数 | 公积金基数 | 备注 |
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import List, Optional

from openpyxl import load_workbook

from ..models import money, ZERO


@dataclass
class AssetItem:
    category: str           # 资产类别
    name: str               # 资产名称
    purchase_date: str      # 购置日期
    original_value: Decimal # 原值
    accumulated_depr: Decimal  # 累计折旧
    net_value: Decimal      # 净值
    note: str = ""


@dataclass
class EmployeeInfo:
    name: str               # 姓名
    id_card: str            # 身份证号
    hire_date: str          # 入职日期
    position: str           # 岗位
    social_base: Decimal    # 社保基数
    fund_base: Decimal      # 公积金基数
    note: str = ""


def _safe_dec(val) -> Decimal:
    if val is None:
        return ZERO
    if isinstance(val, Decimal):
        return val
    if isinstance(val, (int, float)):
        return money(str(val))
    s = str(val).strip().replace(",", "").replace("，", "")
    if not s:
        return ZERO
    try:
        return money(s)
    except Exception:
        return ZERO


def _safe_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


# ── 资产台账 ───────────────────────────────────────────────────────────

def parse_asset_ledger(path: str | Path) -> dict:
    """解析资产台账 → 资产总额

    Args:
        path: 资产台账 Excel 路径

    Returns:
        {
            "total_assets": Decimal,         # 资产总额（净值合计）
            "total_original": Decimal,       # 原值合计
            "items": [AssetItem, ...],       # 明细
        }
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"资产台账不存在: {path}")

    wb = load_workbook(path, data_only=True)
    ws = wb["资产台账"] if "资产台账" in wb.sheetnames else wb.active
    header = [_safe_str(c.value) for c in ws[1]]

    def col_idx(key: str) -> int:
        for i, h in enumerate(header):
            if key in h:
                return i
        return -1

    ci_cat = col_idx("类别")
    ci_name = col_idx("名称")
    ci_date = col_idx("购置")
    ci_orig = col_idx("原值")
    ci_depr = col_idx("折旧")
    ci_net = col_idx("净值")
    ci_note = col_idx("备注")

    items = []
    total_net = ZERO
    total_original = ZERO

    for row in ws.iter_rows(min_row=2, values_only=False):
        name = _safe_str(row[ci_name].value) if ci_name >= 0 and ci_name < len(row) else ""
        if not name:
            continue
        item = AssetItem(
            category=_safe_str(row[ci_cat].value) if ci_cat >= 0 and ci_cat < len(row) else "",
            name=name,
            purchase_date=_safe_str(row[ci_date].value) if ci_date >= 0 and ci_date < len(row) else "",
            original_value=_safe_dec(row[ci_orig].value) if ci_orig >= 0 and ci_orig < len(row) else ZERO,
            accumulated_depr=_safe_dec(row[ci_depr].value) if ci_depr >= 0 and ci_depr < len(row) else ZERO,
            net_value=_safe_dec(row[ci_net].value) if ci_net >= 0 and ci_net < len(row) else ZERO,
            note=_safe_str(row[ci_note].value) if ci_note >= 0 and ci_note < len(row) else "",
        )
        items.append(item)
        total_net += item.net_value
        total_original += item.original_value

    wb.close()
    return {
        "total_assets": total_net,
        "total_original": total_original,
        "items": items,
    }


# ── 员工名册 ───────────────────────────────────────────────────────────

def parse_employee_roster(path: str | Path) -> dict:
    """解析员工名册 → 从业人数 + 员工底表

    Args:
        path: 员工名册 Excel 路径

    Returns:
        {
            "count": int,                    # 从业人数
            "employees": [EmployeeInfo, ...],
        }
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"员工名册不存在: {path}")

    wb = load_workbook(path, data_only=True)
    ws = wb["员工名册"] if "员工名册" in wb.sheetnames else wb.active
    header = [_safe_str(c.value) for c in ws[1]]

    def col_idx(key: str) -> int:
        for i, h in enumerate(header):
            if key in h:
                return i
        return -1

    ci_name = col_idx("姓名")
    ci_id = col_idx("身份证")
    ci_hire = col_idx("入职")
    ci_pos = col_idx("岗位")
    ci_ss = col_idx("社保")
    ci_fund = col_idx("公积金")
    ci_note = col_idx("备注")

    employees = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        name = _safe_str(row[ci_name].value) if ci_name >= 0 and ci_name < len(row) else ""
        if not name:
            continue
        emp = EmployeeInfo(
            name=name,
            id_card=_safe_str(row[ci_id].value) if ci_id >= 0 and ci_id < len(row) else "",
            hire_date=_safe_str(row[ci_hire].value) if ci_hire >= 0 and ci_hire < len(row) else "",
            position=_safe_str(row[ci_pos].value) if ci_pos >= 0 and ci_pos < len(row) else "",
            social_base=_safe_dec(row[ci_ss].value) if ci_ss >= 0 and ci_ss < len(row) else ZERO,
            fund_base=_safe_dec(row[ci_fund].value) if ci_fund >= 0 and ci_fund < len(row) else ZERO,
            note=_safe_str(row[ci_note].value) if ci_note >= 0 and ci_note < len(row) else "",
        )
        employees.append(emp)

    wb.close()
    return {
        "count": len(employees),
        "employees": employees,
    }
