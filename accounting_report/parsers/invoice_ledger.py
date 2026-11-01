# -*- coding: utf-8 -*-
"""发票台账解析器

读取企业维护的发票台账 Excel，返回结构化发票数据供增值税申报底稿使用。

Excel 模板格式（Sheet「发票台账」）：
| 开票日期 | 发票代码 | 发票号码 | 发票类型 | 购方名称 | 货物/服务名称 | 不含税金额 | 税额 | 含税金额 | 备注 |

发票类型（invoice_type）映射：
  - 增值税普通发票 → normal
  - 增值税专用发票 → special
  - 增值税电子普通发票 → e_normal
  - 增值税电子专用发票 → e_special
  - 其他 → other

返回 dict 结构：
{
    "invoiced": Decimal,         # 开票不含税收入合计
    "uninvoiced": Decimal,       # 无票收入（由外部传入或从营业收入推算）
    "tax_exempt": Decimal,       # 免税收入（从备注判定）
    "special_invoice": Decimal,  # 专票不含税金额
    "details": [                 # 发票明细列表
        {
            "code": str,
            "number": str,
            "date": str,
            "buyer": str,
            "service": str,
            "amount": Decimal,
            "tax": Decimal,
        },
        ...
    ],
}
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import List

from openpyxl import load_workbook

from ..models import money, ZERO


@dataclass
class InvoiceRecord:
    """单张发票记录"""
    date: str             # YYYY-MM-DD
    code: str             # 发票代码
    number: str           # 发票号码
    invoice_type: str     # normal / special / e_normal / e_special / other
    buyer: str            # 购方名称
    service: str          # 货物/服务名称
    amount: Decimal       # 不含税金额
    tax: Decimal          # 税额
    total: Decimal        # 含税金额
    note: str = ""        # 备注


def _safe_decimal(val) -> Decimal:
    """将单元格值转为 Decimal，None / 空 / 非数字 → ZERO"""
    if val is None:
        return ZERO
    if isinstance(val, (int, float)):
        return money(str(val))
    if isinstance(val, Decimal):
        return val
    s = str(val).strip().replace(",", "").replace("，", "")
    if s == "" or s == "-":
        return ZERO
    try:
        return money(s)
    except Exception:
        return ZERO


def _safe_str(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return s


def _normalize_invoice_type(raw: str) -> str:
    """标准化发票类型（兼容中文标签与英文码）"""
    r = raw.strip().lower()
    if r in ("special", "e_special", "增值税专用发票", "增值税电子专用发票"):
        return "e_special" if "电子" in r or r.startswith("e_") else "special"
    if r in ("normal", "e_normal", "增值税普通发票", "增值税电子普通发票"):
        return "e_normal" if "电子" in r or r.startswith("e_") else "normal"
    if "专用" in r:
        if "电子" in r:
            return "e_special"
        return "special"
    if "普通" in r:
        if "电子" in r:
            return "e_normal"
        return "normal"
    if "电子" in r:
        return "e_normal"
    return "other"


def parse_invoice_ledger(path: str | Path) -> dict:
    """解析发票台账 Excel → 发票汇总数据

    Args:
        path: 发票台账 xlsx 文件路径。

    Returns:
        dict: 发票汇总数据（供 p5_vat_declaration.generate 使用）。

    Raises:
        FileNotFoundError: 文件不存在。
        ValueError: 格式错误（缺少关键列）。
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"发票台账不存在: {path}")

    wb = load_workbook(path, data_only=True)

    # 优先找「发票台账」sheet，否则取第一个
    if "发票台账" in wb.sheetnames:
        ws = wb["发票台账"]
    else:
        ws = wb.active

    # 读表头行（第1行），建立列索引
    header = []
    for cell in ws[1]:
        header.append(_safe_str(cell.value))

    required = ["发票号码", "不含税金额"]
    col_map = {}
    for req in required:
        found = False
        for i, h in enumerate(header):
            if req in h:
                col_map[req] = i
                found = True
                break
        if not found:
            raise ValueError(f"发票台账缺少必要列: {req}（表头: {header}）")

    # 可选列：模糊匹配列名（兼容"购方名称"/"购方 名称"等写法）
    optional = {
        "开票日期": "date",
        "发票代码": "code",
        "发票号码": "number",
        "发票类型": "type",
        "购方名称": "buyer",
        "货物/服务名称": "service",
        "不含税金额": "amount",
        "税额": "tax",
        "含税金额": "total",
        "备注": "note",
    }
    for h, key in optional.items():
        if h not in col_map:
            for i, ch in enumerate(header):
                if h in ch:
                    col_map[h] = i
                    break

    # 可选列缺失返回 None（调用方必须判空，不能用默认 0——第 0 列可能是「日期」，
    # 缺列静默读第 0 列会把日期当购方名称、导致发票方向误判）
    def _col(name: str) -> int | None:
        return col_map.get(name)

    # 读数据行
    records: List[InvoiceRecord] = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        # 跳过全空行
        vals = [_safe_str(c.value) for c in row]
        if all(v == "" for v in vals):
            continue
        # 跳过发票号码为空的行
        ci = _col("发票号码")
        number = _safe_str(row[ci].value) if ci is not None and ci < len(row) else ""
        if not number:
            continue

        def _cell(name: str):
            i = _col(name)
            if i is None or i >= len(row):
                return None
            return row[i].value

        date_val = _safe_str(_cell("开票日期"))
        # 日期处理：可能是 datetime
        if not date_val:
            cell = _cell("开票日期")
            if cell is not None and hasattr(cell, "strftime"):
                date_val = cell.strftime("%Y-%m-%d")

        record = InvoiceRecord(
            date=date_val,
            code=_safe_str(_cell("发票代码")),
            number=number,
            invoice_type=_normalize_invoice_type(_safe_str(_cell("发票类型"))),
            buyer=_safe_str(_cell("购方名称")),
            service=_safe_str(_cell("货物/服务名称")),
            amount=_safe_decimal(_cell("不含税金额")),
            tax=_safe_decimal(_cell("税额")),
            total=_safe_decimal(_cell("含税金额")),
            note=_safe_str(_cell("备注")),
        )
        records.append(record)

    wb.close()

    # 汇总（区分进项/销项：购方为本公司的发票是费用票/进项，不计入销售额）
    try:
        from ..config import COMPANY_NAME
    except Exception:
        COMPANY_NAME = ""
    invoiced = ZERO
    uninvoiced = ZERO
    tax_exempt = ZERO
    special_invoice = ZERO
    input_amount = ZERO      # 进项发票不含税合计
    details = []

    for r in records:
        is_special = r.invoice_type in ("special", "e_special")
        is_tax_exempt = "免税" in r.note or "免税" in r.service
        # 进项判定：购方名称含本公司名。COMPANY_NAME 为空时不能判进项
        # （空字符串 in 任意字符串恒真 → 全部发票误判进项，销项销售额被清空）
        company = (COMPANY_NAME or "").strip()
        is_input = bool(company) and bool(r.buyer) and company in r.buyer
        unknown = not bool(r.buyer)  # 购方未识别 → 方向待确认，不计入销售额（宁少勿错）
        if is_input or unknown:
            if is_input:
                input_amount += r.amount
            details.append({
                "code": r.code, "number": r.number, "date": r.date,
                "buyer": r.buyer, "service": r.service,
                "amount": r.amount, "tax": r.tax,
                "direction": "进项" if is_input else "待确认",
            })
            continue
        if is_tax_exempt:
            tax_exempt += r.amount
        else:
            invoiced += r.amount
        if is_special:
            special_invoice += r.amount
        details.append({
            "code": r.code,
            "number": r.number,
            "date": r.date,
            "buyer": r.buyer,
            "service": r.service,
            "amount": r.amount,
            "tax": r.tax,
            "direction": "销项",
        })

    return {
        "invoiced": invoiced,
        "uninvoiced": uninvoiced,  # 无票收入需从营业收入推算，此处预留
        "tax_exempt": tax_exempt,
        "special_invoice": special_invoice,
        "input_amount": input_amount,  # 进项不含税合计（费用票）
        "details": details,
    }
