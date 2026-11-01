# -*- coding: utf-8 -*-
"""工资表解析器

读取企业每月提供的工资表 Excel，返回结构化员工工资数据供个税扣缴明细使用。

Excel 模板格式（Sheet「工资表」）：
| 姓名 | 身份证号 | 应发工资 | 个人养老保险 | 个人医疗保险 | 个人失业保险 | 个人公积金 | 企业年金(个人) | 子女教育 | 继续教育 | 住房贷款利息 | 住房租金 | 赡养老人 | 婴幼儿照护 |
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import List, Optional

from openpyxl import load_workbook

from ..models import money, ZERO


@dataclass
class EmployeePayroll:
    """单员工月度工资数据"""
    name: str
    id_card: str
    monthly_income: Decimal       # 本期税前工资
    pension: Decimal              # 基本养老保险（个人）
    medical: Decimal              # 基本医疗保险（个人）
    unemployment: Decimal         # 失业保险（个人）
    housing_fund: Decimal         # 住房公积金（个人）
    enterprise_annuity: Decimal   # 企业年金（个人）
    # 专项附加扣除
    child_edu: Decimal            # 子女教育
    continue_edu: Decimal         # 继续教育
    housing_loan: Decimal         # 住房贷款利息
    housing_rent: Decimal         # 住房租金
    elderly_care: Decimal         # 赡养老人
    infant_care: Decimal          # 婴幼儿照护


def _safe_dec(val) -> Decimal:
    if val is None:
        return ZERO
    if isinstance(val, Decimal):
        return val
    if isinstance(val, (int, float)):
        return money(str(val))
    s = str(val).strip().replace(",", "").replace("，", "")
    if s == "":
        return ZERO
    try:
        return money(s)
    except Exception:
        return ZERO


def _safe_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def parse_payroll_ledger(path: str | Path) -> List[EmployeePayroll]:
    """解析工资表 Excel → 员工列表

    Args:
        path: 工资表 xlsx 文件路径

    Returns:
        EmployeePayroll 列表

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 格式错误
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"工资表不存在: {path}")

    wb = load_workbook(path, data_only=True)

    ws = wb["工资表"] if "工资表" in wb.sheetnames else wb.active

    header = [_safe_str(c.value) for c in ws[1]]

    # 建立列名→列号映射（模糊匹配）
    def find_col(key: str) -> int:
        for i, h in enumerate(header):
            if key in h:
                return i
        return -1

    col_name = find_col("姓名")
    col_id = find_col("身份证")
    col_income = find_col("应发")
    col_pension = find_col("养老")
    col_medical = find_col("医疗")
    col_unemp = find_col("失业")
    col_fund = find_col("公积金")
    col_annuity = find_col("企业年金")
    col_child = find_col("子女")
    col_continue = find_col("继续")
    col_loan = find_col("住房贷款")
    col_rent = find_col("住房租金")
    col_elderly = find_col("赡养")
    col_infant = find_col("婴幼儿")

    if col_name < 0 or col_income < 0:
        wb.close()
        raise ValueError(f"工资表缺少必要列（姓名/应发工资）。表头: {header}")

    def cell(idx: int) -> str | Decimal:
        if idx < 0 or idx >= len(row):
            return ZERO if idx >= 0 else ""
        val = row[idx].value
        if idx in (col_income, col_pension, col_medical, col_unemp,
                    col_fund, col_annuity, col_child, col_continue,
                    col_loan, col_rent, col_elderly, col_infant):
            return _safe_dec(val)
        return _safe_str(val)

    employees = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        name = _safe_str(row[col_name].value) if col_name < len(row) else ""
        if not name:
            continue
        emp = EmployeePayroll(
            name=name,
            id_card=cell(col_id) if isinstance(cell(col_id), str) else "",
            monthly_income=_safe_dec(row[col_income].value) if col_income < len(row) else ZERO,
            pension=cell(col_pension) if isinstance(cell(col_pension), Decimal) else ZERO,
            medical=cell(col_medical) if isinstance(cell(col_medical), Decimal) else ZERO,
            unemployment=cell(col_unemp) if isinstance(cell(col_unemp), Decimal) else ZERO,
            housing_fund=cell(col_fund) if isinstance(cell(col_fund), Decimal) else ZERO,
            enterprise_annuity=cell(col_annuity) if isinstance(cell(col_annuity), Decimal) else ZERO,
            child_edu=cell(col_child) if isinstance(cell(col_child), Decimal) else ZERO,
            continue_edu=cell(col_continue) if isinstance(cell(col_continue), Decimal) else ZERO,
            housing_loan=cell(col_loan) if isinstance(cell(col_loan), Decimal) else ZERO,
            housing_rent=cell(col_rent) if isinstance(cell(col_rent), Decimal) else ZERO,
            elderly_care=cell(col_elderly) if isinstance(cell(col_elderly), Decimal) else ZERO,
            infant_care=cell(col_infant) if isinstance(cell(col_infant), Decimal) else ZERO,
        )
        employees.append(emp)

    wb.close()
    return employees


def employees_to_p7_dict(
    employees: List[EmployeePayroll],
    prior_results: Optional[dict] = None,  # 上月 p7 输出中每人的累计值
) -> list:
    """将 EmployeePayroll 列表转为 p7_payroll_tax.generate 可用的 dict 列表

    Args:
        employees: 当月员工列表
        prior_results: {"张三": {"cumulative_income": ..., "cumulative_prepaid": ...}, ...}
                       从上一期输出 7 读取的累计值

    Returns:
        员工 dict 列表，格式匹配 p7.generate 的 employees 参数
    """
    result = []
    for emp in employees:
        prior = (prior_results or {}).get(emp.name, {})
        social_insurance = emp.pension + emp.medical + emp.unemployment

        cumulative_income = emp.monthly_income + prior.get("cumulative_income", ZERO)
        cumulative_special = (social_insurance + emp.housing_fund + emp.enterprise_annuity
                             + prior.get("cumulative_special_deduct", ZERO))
        cumulative_additional = (
            emp.child_edu + emp.continue_edu + emp.housing_loan +
            emp.housing_rent + emp.elderly_care + emp.infant_care
            + prior.get("cumulative_special_additional", ZERO)
        )

        result.append({
            "name": emp.name,
            "id_card": emp.id_card,
            "monthly_income": emp.monthly_income,
            "tax_free_income": ZERO,
            "pension": emp.pension,
            "medical": emp.medical,
            "unemployment": emp.unemployment,
            "housing_fund": emp.housing_fund,
            "cumulative_income": cumulative_income,
            "cumulative_special_deduct": cumulative_special,
            "cumulative_special_additional": cumulative_additional,
            "cumulative_other_deduct": emp.enterprise_annuity,
            "cumulative_prepaid": prior.get("cumulative_prepaid", ZERO),
            "child_edu": emp.child_edu,
            "continue_edu": emp.continue_edu,
            "housing_loan": emp.housing_loan,
            "housing_rent": emp.housing_rent,
            "elderly_care": emp.elderly_care,
            "infant_care": emp.infant_care,
        })
    return result
