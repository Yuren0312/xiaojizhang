# -*- coding: utf-8 -*-
"""全账务核算汇总表

生成一个 xlsx 工作簿，包含 4 个 sheet：
  1. 科目余额汇总 — 资产/负债/权益/损益 大类汇总
  2. 短期理财明细 — 理财申购、赎回、余额统计
  3. 往来明细 — 阿里云/腾讯/股东A等应付账款、其他应付款明细
  4. 费用分类统计 — 办公费/业务招待/福利费/银行手续费/利息等分类汇总
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ..models import ZERO, money

# ── 样式常量 ────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_FONT = Font(name="微软雅黑", size=10)
BOLD_FONT = Font(name="微软雅黑", bold=True, size=10)
TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
SECTION_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SECTION_FONT = Font(name="微软雅黑", bold=True, size=10, color="1F3864")
MONEY_FMT = '#,##0.00'
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _write_row(ws, row: int, values: list, money_cols: set | None = None,
               bold: bool = False) -> None:
    if money_cols is None:
        money_cols = set()
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BOLD_FONT if bold else CELL_FONT
        cell.alignment = Alignment(vertical="center")
        cell.border = THIN_BORDER
        if i in money_cols and isinstance(v, (Decimal, int, float)):
            cell.number_format = MONEY_FMT


def _auto_width(ws, min_width: int = 8, max_width: int = 40) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            v = str(cell.value) if cell.value is not None else ""
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            lengths.append(length)
        best = max(lengths) if lengths else min_width
        ws.column_dimensions[col_letter].width = max(min(best + 2, max_width), min_width)


# ── Sheet 1: 科目余额汇总 ───────────────────────────────────────────────

def _tb_rows(trial_balances: list) -> list[dict]:
    """把科目余额表输入归一成 {code,name,open_d,open_c,pd,pc,end_d,end_c} 行。

    兼容 dict 与 TrialBalance 对象两种输入；金额列缺失时按 0 处理。
    """
    rows = []
    for item in trial_balances:
        if isinstance(item, dict):
            rows.append({
                "code": str(item.get("account_code", item.get("code", item.get("科目编码", ""))) or ""),
                "name": str(item.get("account_name", item.get("name", item.get("科目", ""))) or ""),
                "type": item.get("type", item.get("类型", "")),
                "open_d": money(item.get("open_debit", item.get("期初借方", ZERO))),
                "open_c": money(item.get("open_credit", item.get("期初贷方", ZERO))),
                "pd": money(item.get("period_debit", item.get("借方发生额", item.get("debit", ZERO)))),
                "pc": money(item.get("period_credit", item.get("贷方发生额", item.get("credit", ZERO)))),
                "end_d": money(item.get("end_debit", item.get("期末借方", ZERO))),
                "end_c": money(item.get("end_credit", item.get("期末贷方", ZERO))),
            })
        else:
            rows.append({
                "code": str(getattr(item, "account_code", "") or ""),
                "name": str(getattr(item, "account_name", "") or getattr(item, "name", "")),
                "type": getattr(item, "type", ""),
                "open_d": money(getattr(item, "open_debit", ZERO)),
                "open_c": money(getattr(item, "open_credit", ZERO)),
                "pd": money(getattr(item, "period_debit", ZERO)),
                "pc": money(getattr(item, "period_credit", ZERO)),
                "end_d": money(getattr(item, "end_debit", ZERO)),
                "end_c": money(getattr(item, "end_credit", ZERO)),
            })
    return [r for r in rows if r["name"] or r["code"]]


def _code_class(row: dict) -> str | None:
    """科目编码首段分类（1 资产 / 2 负债 / 3 权益；5/6 等损益类无法仅凭编码区分）"""
    code = (row["code"] or "").strip().split(".")[0]
    if not code.isdigit():
        return None
    return {"1": "资产", "2": "负债", "3": "权益", "4": "权益"}.get(code[0])


def _name_class(name: str) -> str | None:
    """科目名称启发式兜底（收入/费用带方向词；不可解析时返回 None）"""
    if not name:
        return None
    if any(kw in name for kw in ("收入", "收益", "利得", "利息收入")):
        return "收入"
    if any(kw in name for kw in ("成本", "费用", "支出", "损失", "税", "折旧", "摊销", "招待", "福利", "办公")):
        return "费用"
    return None


def _resolve_type(name: str) -> str | None:
    """按科目名称解析类别（走 rules 别名归一）；未定义返回 None"""
    if not name:
        return None
    try:
        from ..rules import resolve_account
        return resolve_account(name).type
    except KeyError:
        return None


TYPE_TO_GROUP = {
    "asset": "资产", "asset_contra": "资产",
    "liability": "负债",
    "equity": "权益",
    "revenue": "收入",
    "expense": "费用",
}


def _trial_balance_summary(ws, trial_balances: list) -> None:
    """资产/负债/权益/损益 大类汇总（借贷分列，无负数误导）

    此前把「期初/期末净额」合并成一列：负债/权益类余额以负数呈现（
    -100513.01），非财务用户极易误解；且子科目行（如 100201 中国工商银行）
    因名称不在科目表内被归入默认「费用」类。本版：
    - 期初/期末按 借方/贷方 分列（与源科目余额表/试算平衡表格式一致）；
    - 科目归类：科目表名称解析 → 编码首段（1资产/2负债/3权益/4权益）→
      名称关键词兜底，不再把未知科目一律扔进「费用」；
    - 父子科目去重：科目余额表常同时导出 汇总科目+明细科目 两行（1002 银行存款
      与 100201 中国工商银行 数值完全相同），按编码前缀保留父科目行，避免
      小计把同一余额重复相加。
    """
    rows = _tb_rows(trial_balances)

    # 父子去重：编码为其他编码严格前缀者视为汇总父行下的明细行，丢弃明细行
    # （父行含完整数值且名称可解析；丢父行则汇总科目整体消失，丢子行更安全）。
    codes = [r["code"] for r in rows if r["code"]]
    children = {
        c for c in codes
        if any(c.startswith(o) and len(c) > len(o) for o in codes if o and o != c)
    }
    rows = [r for r in rows if r["code"] not in children]

    groups: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        # 显式 type（dict 自带 asset/expense 等）优先；否则科目表名称解析（返回
        # 英文 type，再映射中文）；仍失败则编码首段 → 名称关键词兜底；都不中默认「费用」。
        raw = r.get("type") or _resolve_type(r["name"])
        group = TYPE_TO_GROUP.get(raw, None)
        if group is None:
            group = _code_class(r) or _name_class(r["name"]) or "费用"
        groups[group].append(r)

    headers = ["科目名称", "期初借方", "期初贷方", "本期借方", "本期贷方",
               "期末借方", "期末贷方"]
    money_cols = {2, 3, 4, 5, 6, 7}
    col_count = len(headers)
    for ci, h in enumerate(headers, start=1):
        ws.cell(row=2, column=ci, value=h)
    _style_header(ws, 2, col_count)

    # 总计（默认全部为零；保证「总计」行永远存在）
    totals = {k: ZERO for k in ("open_d", "open_c", "pd", "pc", "end_d", "end_c")}

    row = 3
    for group_name, group_rows in groups.items():
        if not group_rows:
            continue
        # 大类标题行
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_count)
        ws.cell(row=row, column=1, value=f"── {group_name}类 ──").font = SECTION_FONT
        for c in range(1, col_count + 1):
            ws.cell(row=row, column=c).fill = SECTION_FILL
        row += 1

        group_totals = {k: ZERO for k in totals}

        for r in sorted(group_rows, key=lambda x: x["name"]):
            vals = [r["name"], r["open_d"], r["open_c"], r["pd"], r["pc"],
                    r["end_d"], r["end_c"]]
            _write_row(ws, row, [float(v) if isinstance(v, Decimal) and i else v
                                 for i, v in enumerate(vals)], money_cols)
            for i, key in enumerate(("open_d", "open_c", "pd", "pc", "end_d", "end_c")):
                group_totals[key] += r[key]
                totals[key] += r[key]
            row += 1

        # 小计行
        _write_row(ws, row, [
            f"{group_name}类小计",
            float(group_totals["open_d"]), float(group_totals["open_c"]),
            float(group_totals["pd"]), float(group_totals["pc"]),
            float(group_totals["end_d"]), float(group_totals["end_c"]),
        ], money_cols, bold=True)
        row += 1
        row += 1  # 空行分隔

    # 总计行
    _write_row(ws, row, [
        "总计",
        float(totals["open_d"]), float(totals["open_c"]),
        float(totals["pd"]), float(totals["pc"]),
        float(totals["end_d"]), float(totals["end_c"]),
    ], money_cols, bold=True)

    # 借贷平衡提示（有数据时）
    if totals["pd"] != ZERO or totals["pc"] != ZERO:
        ok = (totals["open_d"] == totals["open_c"] and totals["pd"] == totals["pc"]
              and totals["end_d"] == totals["end_c"])
        hint = "✓ 期初/本期/期末 借贷分别平衡" if ok else "✗ 借贷不平衡，请核查源科目余额表"
        _write_row(ws, row + 2, [hint, "", "", "", "", "", ""], set())
        ws.cell(row=row + 2, column=1).font = BOLD_FONT

    _auto_width(ws, min_width=12)


# ── Sheet 2: 短期理财明细 ───────────────────────────────────────────────

def _short_term_investment_detail(ws, ledger_entries: list) -> None:
    """理财申购、赎回、余额统计"""
    headers = ["日期", "摘要", "对方科目", "申购金额", "赎回金额", "余额", "备注"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {4, 5, 6}
    row = 3

    # 筛选短期投资相关分录
    inv_entries = []
    for entry in ledger_entries:
        if isinstance(entry, dict):
            acct = entry.get("account_name", entry.get("科目名称", ""))
            dr = entry.get("debit", 0)
            cr = entry.get("credit", 0)
            summary = entry.get("summary", entry.get("摘要", ""))
        else:
            acct = str(getattr(entry, 'account_name', getattr(entry, 'debit', '')))
            dr = getattr(entry, 'debit', 0)
            cr = getattr(entry, 'credit', 0)
            summary = getattr(entry, 'summary', '')

        # 只按科目筛选：摘要含"理财"会把银行存款等科目的流水误拉进来
        if "短期投资" in str(acct):
            inv_entries.append(entry)

    if not inv_entries:
        _write_row(ws, row, ["", "暂无短期理财交易记录", "", "", "", "", ""], money_cols)
        _auto_width(ws)
        return

    # 按日期排序
    inv_entries.sort(key=lambda e: getattr(e, 'date', '') if not isinstance(e, dict)
                     else e.get("date", e.get("日期", "")))

    running_balance = ZERO
    for entry in inv_entries:
        if isinstance(entry, dict):
            date = entry.get("date", entry.get("日期", ""))
            summary = entry.get("summary", entry.get("摘要", ""))
            acct = str(entry.get("account_name", entry.get("科目名称", "")))
            debit_amt = money(entry.get("debit", ZERO))
            credit_amt = money(entry.get("credit", ZERO))
            amount = debit_amt if debit_amt > ZERO else credit_amt
        else:
            date = getattr(entry, 'date', '')
            summary = getattr(entry, 'summary', '')
            acct = str(getattr(entry, 'account_name', ''))
            debit_amt = money(getattr(entry, 'debit', ZERO))
            credit_amt = money(getattr(entry, 'credit', ZERO))
            amount = debit_amt if debit_amt > ZERO else credit_amt

        is_purchase = "短期投资" in acct and debit_amt > ZERO   # 借短期投资 = 申购
        is_redeem = "短期投资" in acct and credit_amt > ZERO   # 贷短期投资 = 赎回

        purchase = float(amount) if is_purchase else 0.0
        redeem = float(amount) if is_redeem else 0.0
        if is_purchase:
            running_balance += amount
        elif is_redeem:
            running_balance -= amount

        counterparty = "货币资金"  # 明细账无对方科目信息，统一显示货币资金
        _write_row(ws, row, [
            date, summary, counterparty,
            purchase if purchase else "",
            redeem if redeem else "",
            float(running_balance),
            "申购" if is_purchase else "赎回",
        ], money_cols)
        row += 1

    # 汇总
    row += 1
    _write_row(ws, row, ["", "期末余额", "", "", "", float(running_balance), ""],
               money_cols, bold=True)

    _auto_width(ws)


# ── Sheet 3: 往来明细 ───────────────────────────────────────────────────

def _aging_ref_date(period_end, ledger_entries) -> str:
    """账龄计算的参考日期（YYYY-MM-DD）。

    period_end: 报表账期（'YYYY-MM' 或 'YYYY-MM-DD'）；缺省取明细账最大日期。
    此前直接 date.today() 导致账龄随运行日期漂移、历史报表不可复现（白班20），
    报表账龄应始终以「账期末日」为基准。
    """
    import calendar
    from datetime import date, datetime
    raw = str(period_end or "").strip()
    if len(raw) >= 7 and raw[:4].isdigit() and raw[5:7].isdigit():
        y, m = int(raw[:4]), int(raw[5:7])
        if len(raw) >= 10 and raw[8:10].isdigit():
            return f"{raw[:10]}"
        last = calendar.monthrange(y, m)[1]
        return f"{y:04d}-{m:02d}-{last:02d}"
    # 兜底：以明细账中的最大交易日期作为截至日（确定性，不依赖系统时钟）
    latest = ""
    for e in ledger_entries or []:
        d = str(e.get("date", "") or "")[:10] if isinstance(e, dict) \
            else str(getattr(e, "date", "") or "")[:10]
        if d and (not latest or d > latest):
            latest = d
    return latest or ""


def _payable_detail(ws, ledger_entries: list, period_end: str | None = None) -> None:
    """应付账款、其他应付款明细"""
    headers = ["日期", "摘要", "对方", "科目", "借方", "贷方", "余额", "账龄（天）"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {5, 6, 7}
    row = 3

    # 筛选往来科目
    payable_entries = []
    target_accounts = ("应付账款", "其他应付款", "预付账款", "应收账款", "其他应收款")

    for entry in ledger_entries:
        if isinstance(entry, dict):
            acct = str(entry.get("account_name", entry.get("科目名称", "")))
        else:
            acct = str(getattr(entry, 'account_name', ''))

        hits = any(acct in a for a in target_accounts) or any(a in acct for a in target_accounts)
        if hits:
            payable_entries.append(entry)

    if not payable_entries:
        _write_row(ws, row, ["", "", "暂无往来明细", "", "", "", "", ""], money_cols)
        _auto_width(ws)
        return

    # 按对方分组统计余额
    counterparties: dict[str, dict[str, Decimal]] = defaultdict(
        lambda: {"debit": ZERO, "credit": ZERO, "entries": [], "account": ""}
    )

    for entry in payable_entries:
        if isinstance(entry, dict):
            date = entry.get("date", entry.get("日期", ""))
            summary = entry.get("summary", entry.get("摘要", ""))
            acct = str(entry.get("account_name", entry.get("科目名称", "")))
            debit_amt = money(entry.get("debit", ZERO))
            credit_amt = money(entry.get("credit", ZERO))
            amount = debit_amt if debit_amt > ZERO else credit_amt
        else:
            date = getattr(entry, 'date', '')
            summary = getattr(entry, 'summary', '')
            acct = str(getattr(entry, 'account_name', ''))
            debit_amt = money(getattr(entry, 'debit', ZERO))
            credit_amt = money(getattr(entry, 'credit', ZERO))
            amount = debit_amt if debit_amt > ZERO else credit_amt

        # 确定往来科目（借/贷方向由金额方向决定）
        is_payable_dr = debit_amt > ZERO
        is_payable_cr = credit_amt > ZERO
        acct = next((a for a in target_accounts if a in acct), acct)

        # 对方：无对应科目名，用摘要推断
        counterparty = summary

        cp = counterparty or "（未识别）"
        counterparties[cp]["entries"].append(entry)
        counterparties[cp]["account"] = acct
        if is_payable_dr:
            counterparties[cp]["debit"] += amount
        elif is_payable_cr:
            counterparties[cp]["credit"] += amount

    for cp in sorted(counterparties.keys()):
        c = counterparties[cp]
        balance = c["credit"] - c["debit"]
        # 取最早日期计算账龄
        dates = []
        for e in c["entries"]:
            d = getattr(e, 'date', '') if not isinstance(e, dict) else e.get("date", e.get("日期", ""))
            dates.append(str(d))
        first_date = min(dates, default="")
        last_date = max(dates, default="")
        # 账龄基准 = 报表账期末日（确定性；此前 date.today() 输出随时间漂移）
        ref = _aging_ref_date(period_end, [])
        aging = ""
        if ref and last_date and ref >= last_date:
            try:
                from datetime import datetime
                d_ref = datetime.strptime(ref, "%Y-%m-%d").date()
                d = datetime.strptime(last_date, "%Y-%m-%d").date()
                aging = (d_ref - d).days
            except (ValueError, OverflowError):
                aging = ""

        _write_row(ws, row, [
            last_date if len(dates) > 1 else first_date,
            cp, cp, c["account"],
            float(c["debit"]) if c["debit"] > ZERO else "",
            float(c["credit"]) if c["credit"] > ZERO else "",
            float(balance),
            str(aging) if aging != "" else "",
        ], money_cols)
        row += 1

    _auto_width(ws)


# ── Sheet 4: 费用分类统计 ───────────────────────────────────────────────

def _expense_category_stats(ws, ledger_entries: list) -> None:
    """办公费/业务招待/福利费/银行手续费/利息等分类汇总"""
    headers = ["月份", "费用类别", "金额", "占费用总额比例", "备注"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {3}
    row = 3

    # 按科目和月份归类费用
    # 预定义的分类映射
    category_map = {
        "财务费用": "银行手续费",
        "财务费用-利息": "利息收支",
        "营业成本": "营业成本",
        "业务招待费": "业务招待费",
        "管理费用": "管理费用",
        "销售费用": "销售费用",
        "税金及附加": "税金及附加",
        "城市维护建设税": "税金及附加",
        "所得税费用": "所得税费用",
        "开办费": "开办费",
        "研究费用": "研究费用",
        "营业外支出": "营业外支出",
    }

    monthly_categories: dict[str, dict[str, Decimal]] = defaultdict(
        lambda: defaultdict(lambda: ZERO)
    )
    expense_total = ZERO

    for entry in ledger_entries:
        if isinstance(entry, dict):
            date_str = entry.get("date", entry.get("日期", ""))
            acct = str(entry.get("account_name", entry.get("科目名称", "")))
            debit_amt = money(entry.get("debit", ZERO))
            credit_amt = money(entry.get("credit", ZERO))
        else:
            date_str = getattr(entry, 'date', '')
            acct = str(getattr(entry, 'account_name', ''))
            debit_amt = money(getattr(entry, 'debit', ZERO))
            credit_amt = money(getattr(entry, 'credit', ZERO))

        # 判断科目是否为费用类（费用借方增加）
        cat = category_map.get(acct, "")
        if not cat and acct:
            # 对于没在映射里的科目，看是否是费用类
            if any(kw in acct for kw in ("费", "税", "支出", "成本", "折旧")):
                cat = acct

        if cat:
            month_key = str(date_str)[:7] if date_str else ""
            # 费用类科目按借贷净额计入（借方增加、贷方冲减）：
            # 利息收入通常「贷 财务费用」冲减费用，若只取单侧金额会把冲减
            # 当正向费用累加，费用总额虚高、占比失真（白班20）。
            monthly_categories[month_key][cat] += (debit_amt - credit_amt)
            expense_total += (debit_amt - credit_amt)

    if not monthly_categories or not any(
            v for cats in monthly_categories.values() for v in cats.values()):
        _write_row(ws, row, ["", "", "暂无费用记录", "", ""], money_cols)
        _auto_width(ws)
        return

    # 输出
    for month_key in sorted(monthly_categories.keys()):
        cats = monthly_categories[month_key]
        for cat in sorted(cats.keys()):
            amt = float(cats[cat])
            pct = f"{amt / float(expense_total) * 100:.1f}%" if expense_total > ZERO else "0%"
            _write_row(ws, row, [month_key, cat, amt, pct, ""], money_cols)
            row += 1

    # 汇总
    row += 1
    _write_row(ws, row, ["合计", "费用总额", float(expense_total), "100%", ""],
               money_cols, bold=True)

    _auto_width(ws)


# ── 主函数 ──────────────────────────────────────────────────────────────

def write_account_summary(
    trial_balances: list,   # 科目余额条目
    ledger_entries: list,   # 明细账条目（可按科目分组）
    out_path: Path,
    period_end: str | None = None,  # 报表账期 'YYYY-MM'/'YYYY-MM-DD'：账龄按账期末日算
) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Sheet 1: 科目余额汇总
    ws1 = wb.active
    ws1.title = "科目余额汇总"
    ws1.cell(row=1, column=1, value="科目余额汇总表").font = TITLE_FONT
    ws1.cell(row=1, column=3, value=f"单位：元").font = Font(name="微软雅黑", size=9, color="666666")
    _trial_balance_summary(ws1, trial_balances)

    # Sheet 2: 短期理财明细
    ws2 = wb.create_sheet("短期理财明细")
    ws2.cell(row=1, column=1, value="短期理财明细").font = TITLE_FONT
    _short_term_investment_detail(ws2, ledger_entries)

    # Sheet 3: 往来明细
    ws3 = wb.create_sheet("往来明细")
    ws3.cell(row=1, column=1, value="往来款项明细").font = TITLE_FONT
    _payable_detail(ws3, ledger_entries, period_end=period_end)

    # Sheet 4: 费用分类统计
    ws4 = wb.create_sheet("费用分类统计")
    ws4.cell(row=1, column=1, value="费用分类统计").font = TITLE_FONT
    _expense_category_stats(ws4, ledger_entries)

    wb.save(out_path)
    return out_path
