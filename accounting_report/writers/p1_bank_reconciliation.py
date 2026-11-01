# -*- coding: utf-8 -*-
"""银行对账报告

生成一个 xlsx 工作簿，包含 5 个 sheet：
  1. 对账明细 — 银行流水与账面银行逐笔核对
  2. 未达账项 — 银行已记企业未记 + 企业已记银行未记
  3. 余额调节表 — 银行对账单余额调节 + 企业账面余额调节
  4. 手续费利息统计 — 按月汇总银行手续费、利息收入、利息支出
  5. 往来分类汇总 — 按对方户名分类汇总

对账口径（白班19 统一，三个 sheet 共用同一份配对结果，杜绝互相矛盾）：
  - 配对规则：金额相等（±0.02）、方向正确（银行借方=付款 ↔ 账面贷方为
    银行存款/货币资金；银行贷方=收款 ↔ 账面借方为银行存款/货币资金）、
    日期同一月份（引擎按月切分数据；凭证常以月末批次入账、银行流水为实际
    入账日，同一笔跨数日属正常——逐日窗口会让「对账明细/未达账项/余额调节
    表」对同一笔给出不同判定，调节表两侧失衡、结论失真）。
  - 消耗式配对：每笔银行流水至多匹配一笔账面分录、反之亦然（不消耗的话，
    同额多笔会被重复拿去「证明存在」，真实未达被掩盖）。
  - 只有真正缺对方侧配对的项才算未达；零金额行（借=贷=0）不是真实动账，
    不参与配对、也不计入未达。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Font, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ..models import ZERO, money
from ..parsers.icbc_statement import BankTxn

# ── 样式 ──────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_FONT = Font(name="微软雅黑", size=10)
MONEY_FMT = '#,##0.00'
TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
SUBTITLE_FONT = Font(name="微软雅黑", bold=True, size=11)
SECTION_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _write_row(ws, row: int, values: list, money_cols: set | None = None) -> None:
    """写一行数据，金额列用数字格式"""
    if money_cols is None:
        money_cols = set()
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = CELL_FONT
        cell.alignment = Alignment(vertical="center")
        if i in money_cols and isinstance(v, (Decimal, int, float)):
            cell.number_format = MONEY_FMT


def _auto_width(ws, min_width: int = 8, max_width: int = 36) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            v = str(cell.value) if cell.value is not None else ""
            # 中文字符按 2 个宽度估算
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            lengths.append(length)
        best = max(lengths) if lengths else min_width
        ws.column_dimensions[col_letter].width = max(min(best + 2, max_width), min_width)


# ── 账面分录通用取值（兼容 JournalEntry / 带 debit_account 对象 / dict）────

def _dr(e) -> str:
    if hasattr(e, 'debit_account'):
        return str(e.debit_account)
    if isinstance(e, dict):
        return str(e.get('debit_account') or e.get('debit', ''))
    return str(getattr(e, 'debit', ''))


def _cr(e) -> str:
    if hasattr(e, 'credit_account'):
        return str(e.credit_account)
    if isinstance(e, dict):
        return str(e.get('credit_account') or e.get('credit', ''))
    return str(getattr(e, 'credit', ''))


def _date(e) -> str:
    return str(getattr(e, 'date', ''))


def _summary(e) -> str:
    return str(getattr(e, 'summary', ''))


def _amt(e) -> Decimal:
    return money(getattr(e, 'amount', ZERO))


def _is_bank_out_book(e) -> bool:
    """账面分录是否表现为银行存款减少（贷方是银行）"""
    return "货币资金" in _cr(e) or "银行存款" in _cr(e)


def _is_bank_in_book(e) -> bool:
    """账面分录是否表现为银行存款增加（借方是银行）"""
    return "货币资金" in _dr(e) or "银行存款" in _dr(e)


def _is_cash_entry(e) -> bool:
    return _is_bank_in_book(e) or _is_bank_out_book(e)


def _txn_amt(t: BankTxn) -> Decimal:
    return t.debit if t.debit > ZERO else t.credit


def _txn_is_out(t: BankTxn) -> bool:
    """银行借方 = 账户支出"""
    return t.debit > ZERO


def _same_month(a: str, b: str) -> bool:
    return str(a)[:7] == str(b)[:7]


# ── 单一消耗式配对（全报告唯一对账依据）────────────────────────────────

def _match_bank_book(bank_txns: list[BankTxn], book_txns: list) -> list[int | None]:
    """消耗式配对：bank_txns[i] → 匹配到的 book_txns 下标（None = 未匹配）。

    规则：
      - 金额相等（差 < 0.02），方向正确；
      - 日期同一月份（跨数日属月末批次入账的正常现象，不做逐日窗口）；
      - 零金额银行行不参与配对；
      - 每笔账面分录至多配一笔银行流水（consumptive）。
    """
    pairs: list[int | None] = [None] * len(bank_txns)
    used_book: set[int] = set()
    for bi, t in enumerate(bank_txns):
        amt = _txn_amt(t)
        if amt <= ZERO:
            continue
        want_out = _txn_is_out(t)      # 银行支出 → 账面贷银行
        for ebi, e in enumerate(book_txns):
            if ebi in used_book:
                continue
            if not _is_cash_entry(e):
                continue
            dir_ok = (_is_bank_out_book(e) if want_out else _is_bank_in_book(e))
            if not dir_ok:
                continue
            if not _same_month(_date(e), t.date):
                continue
            if abs(amt - _amt(e)) >= Decimal("0.02"):
                continue
            pairs[bi] = ebi
            used_book.add(ebi)
            break
    return pairs


def _unmatched_bank(bank_txns: list[BankTxn],
                    pairs: list[int | None]) -> list[tuple[int, BankTxn]]:
    """银行已记、企业未记（含方向与金额信息供 Sheet2/3 使用）"""
    out = []
    for i, t in enumerate(bank_txns):
        if pairs[i] is None and _txn_amt(t) > ZERO:
            out.append((i, t))
    return out


def _unmatched_book(book_txns: list, used_book: set[int]) -> list[tuple[int, object]]:
    """企业已记、银行未记（仅涉银行分录）"""
    return [(i, e) for i, e in enumerate(book_txns)
            if i not in used_book and _is_cash_entry(e)]


# ── 对账明细 ────────────────────────────────────────────────────────────

def _reconciliation_detail(ws, bank_txns: list[BankTxn], book_txns: list) -> None:
    """银行流水与账面银行逐笔核对（配对结果与 Sheet2/3 完全同源）"""
    pairs = _match_bank_book(bank_txns, book_txns)
    used_book = {p for p in pairs if p is not None}

    headers = ["日期", "摘要", "对方", "银行借方", "银行贷方",
               "账面借方", "账面贷方", "匹配状态", "差异说明"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {4, 5, 6, 7}
    row = 3

    for idx, txn in enumerate(bank_txns):
        bank_debit = float(txn.debit) if txn.debit else 0.0
        bank_credit = float(txn.credit) if txn.credit else 0.0
        ebi = pairs[idx]
        matched = ebi is not None
        status = "已匹配" if matched else "未匹配"
        book_debit_val = None
        book_credit_val = None
        diff_note = ""
        if matched:
            e = book_txns[ebi]
            amt = float(_amt(e))
            # 银行借方（企业付款）→ 账面：银行存款减少（贷方）
            if txn.debit > ZERO:
                book_credit_val = amt
            else:
                book_debit_val = amt
            diff_note = f"匹配: 分录#{ebi + 1}"
        else:
            diff_note = "银行有、账面无 — 待查明"
            if _txn_amt(txn) > ZERO:
                diff_note += f" | 银行金额: {bank_debit:.2f}/{bank_credit:.2f}"
            else:
                diff_note = "零金额行（非动账）"

        _write_row(ws, row, [
            txn.date, txn.summary, txn.counterparty,
            bank_debit if bank_debit else "",
            bank_credit if bank_credit else "",
            book_debit_val if book_debit_val else "",
            book_credit_val if book_credit_val else "",
            status, diff_note,
        ], money_cols)
        row += 1

    # 账面有、银行无的（与上面银行行无关的剩余分录）
    for ebi, e in _unmatched_book(book_txns, used_book):
        amt = float(_amt(e))
        _write_row(ws, row, [
            _date(e), _summary(e), "",
            "", "",
            amt if _is_bank_in_book(e) else "",
            amt if _is_bank_out_book(e) else "",
            "未匹配", "账面有、银行无 — 可能未达或POS延时",
        ], money_cols)
        row += 1

    _auto_width(ws)


# ── 未达账项 ────────────────────────────────────────────────────────────

def _outstanding_items(ws, bank_txns: list[BankTxn], book_txns: list) -> None:
    """未达账项：银行已记企业未记 / 企业已记银行未记（与调节表同一配对）"""
    headers = ["类型", "日期", "摘要", "金额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {4}
    row = 3

    pairs = _match_bank_book(bank_txns, book_txns)
    used_book = {p for p in pairs if p is not None}

    # 企业已记、银行未记
    for ebi, e in _unmatched_book(book_txns, used_book):
        _write_row(ws, row, [
            "企业已记、银行未记", _date(e), _summary(e), float(_amt(e))
        ], money_cols)
        row += 1

    # 银行已记、企业未记
    for bi, txn in _unmatched_bank(bank_txns, pairs):
        _write_row(ws, row, [
            "银行已记、企业未记", txn.date,
            f"{txn.summary} | 对方:{txn.counterparty}",
            float(_txn_amt(txn)),
        ], money_cols)
        row += 1

    if row == 3:
        _write_row(ws, row, ["", "", "无未达账项", ""], money_cols)

    _auto_width(ws)


# ── 余额调节表 ──────────────────────────────────────────────────────────

def _balance_reconciliation(ws, bank_txns: list[BankTxn], book_txns: list,
                            opening_cash: Decimal | int | float | None = None) -> None:
    """余额调节表

    opening_cash：企业账面银行存款期末余额（来自本期报表 BS 货币资金），
    直接作为调节表右侧账面余额；None 时回退为按本月分录借贷净额推算。

    未达账项与「未达账项」sheet 共用同一配对结果（白班19 统一口径）：
    Sheet2 逐笔列示的就是调节表里计入的净额，两表永不打架。
    """
    # 银行对账单余额 = 最后一笔流水（按日期）的余额。
    # 不能用 bank_txns[-1] 盲取：输入目录常有多份对账单 PDF 合并（对账单1.pdf +
    # 对账单1(1).pdf 等），rglob 顺序与解析顺序不保证按日期递增；若某份 PDF
    # 覆盖了整月而另一份只覆盖月初，[-1] 会拿到「月初」的余额，余额调节表
    # 银行侧起点错误、调节结论失真。显式取日期最大的最后一笔（同日多笔时
    # 按列表中最后一条——单 PDF 内即对账单顺序）。
    if bank_txns:
        latest_date = max(t.date for t in bank_txns)
        last_of_day = [t for t in bank_txns if t.date == latest_date][-1]
        bank_balance = float(last_of_day.balance)
    else:
        bank_balance = 0.0

    # 统一配对 → 两侧未达项同源
    pairs = _match_bank_book(bank_txns, book_txns)
    used_book = {p for p in pairs if p is not None}

    bank_only_in = ZERO   # 银行已收、企业未收
    bank_only_out = ZERO  # 银行已付、企业未付
    for bi, txn in _unmatched_bank(bank_txns, pairs):
        if txn.debit > ZERO:
            bank_only_out += txn.debit
        else:
            bank_only_in += txn.credit

    company_only_in = ZERO   # 企业已收、银行未收
    company_only_out = ZERO  # 企业已付、银行未付
    for ebi, e in _unmatched_book(book_txns, used_book):
        if _is_bank_in_book(e):
            company_only_in += _amt(e)
        elif _is_bank_out_book(e):
            company_only_out += _amt(e)

    # 企业账面余额：优先用传入的期末余额；无则从本月涉及货币资金的分录净额推算
    if opening_cash is not None:
        book_cash_balance = money(opening_cash)
    else:
        book_cash_balance = ZERO
        for e in book_txns:
            amt = _amt(e)
            if _is_bank_in_book(e):
                book_cash_balance += amt
            elif _is_bank_out_book(e):
                book_cash_balance -= amt

    # 标准余额调节表口径（企业银行存款余额调节表）：
    #   银行对账单余额 ＋ 企业已收、银行未收（在途存款） － 企业已付、银行未付（未达支票）＝ 调节后余额
    #   企业账面余额　　＋ 银行已收、企业未收（银行代收）　　－ 银行已付、企业未付（银行代付）＝ 调节后余额
    # 两侧调节项必须用对方侧的未达项，方向反了会把存在未达账项的正常账套误判为「调节不一致」。
    bank_adjusted = Decimal(str(bank_balance)) + company_only_in - company_only_out

    book_adjusted = book_cash_balance + bank_only_in - bank_only_out

    # 写调节表
    headers = ["项目", "金额", "项目", "金额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {2, 4}
    row = 3

    # 左侧：银行对账单侧（加在途存款、减未兑现支票 → 调节后余额）
    items = [
        ("银行对账单余额", float(bank_balance)),
        ("加：企业已收、银行未收", float(company_only_in)),
        ("减：企业已付、银行未付", float(company_only_out)),
        ("调节后余额", float(bank_adjusted)),
    ]
    for label, val in items:
        _write_row(ws, row, [label, val, "", ""], money_cols)
        row += 1

    # 右侧：企业账面侧（加银行代收、减银行代付 → 调节后余额）
    row = 3
    items = [
        ("", None, "企业账面银行存款余额", float(book_cash_balance)),
        ("", None, "加：银行已收、企业未收", float(bank_only_in)),
        ("", None, "减：银行已付、企业未付", float(bank_only_out)),
        ("", None, "调节后余额", float(book_adjusted)),
    ]
    for _, _, label, val in items:
        # 先不覆盖已有左侧数据，在右侧列写
        ws.cell(row=row, column=3, value=label).font = CELL_FONT
        ws.cell(row=row, column=4, value=val if val is not None else "").font = CELL_FONT
        if isinstance(val, (int, float)):
            ws.cell(row=row, column=4).number_format = MONEY_FMT
        ws.cell(row=row, column=3).font = CELL_FONT
        row += 1

    # 添加调节结论
    row += 1
    ws.cell(row=row, column=1, value="调节结论：").font = Font(name="微软雅黑", bold=True, size=10)
    if abs(bank_adjusted - book_adjusted) < Decimal("0.02"):
        ws.cell(row=row, column=2, value="银行与企业余额调节一致 ✓").font = Font(
            name="微软雅黑", size=10, color="008000"
        )
    else:
        ws.cell(row=row, column=2, value=f"调节不一致！差额: {float(bank_adjusted - book_adjusted):.2f}").font = Font(
            name="微软雅黑", size=10, color="FF0000"
        )

    # 口径说明（与 Sheet2 同源，用户可对照）
    note_row = row + 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=4)
    ws.cell(row=note_row, column=1,
            value="口径：未达账项 = 单方动账且缺对方配对（配对按 金额相等+方向一致+同月，"
                  "消耗式一一对应），与「未达账项」sheet 完全一致").font = Font(
        name="微软雅黑", size=9, color="808080")

    _auto_width(ws)


# ── 手续费利息统计 ──────────────────────────────────────────────────────

def _fee_interest_stats(ws, bank_txns: list[BankTxn]) -> None:
    """按月汇总银行手续费、利息收入、利息支出"""
    from collections import defaultdict

    # 按月份归类
    monthly: dict[str, dict[str, Decimal]] = defaultdict(
        lambda: {"手续费": ZERO, "利息收入": ZERO, "利息支出": ZERO}
    )

    for txn in bank_txns:
        month_key = txn.date[:7]  # YYYY-MM
        text = f"{txn.summary} {txn.counterparty}"
        if "手续费" in text:
            monthly[month_key]["手续费"] += txn.debit
        elif "利息" in text:
            if txn.debit > ZERO:
                monthly[month_key]["利息支出"] += txn.debit
            if txn.credit > ZERO:
                monthly[month_key]["利息收入"] += txn.credit
        elif "认证" in text or "小微" in text:
            monthly[month_key]["利息收入"] += txn.credit

    headers = ["月份", "手续费（借方）", "利息收入（贷方）", "利息支出（借方）", "净额"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {2, 3, 4, 5}
    row = 3
    for month_key in sorted(monthly.keys()):
        m = monthly[month_key]
        fee = float(m["手续费"])
        int_in = float(m["利息收入"])
        int_out = float(m["利息支出"])
        net = -fee + int_in - int_out
        _write_row(ws, row, [month_key, fee, int_in, int_out, net], money_cols)
        row += 1

    # 合计行
    if row > 3:
        tot_fee = sum(float(m["手续费"]) for m in monthly.values())
        tot_in = sum(float(m["利息收入"]) for m in monthly.values())
        tot_out = sum(float(m["利息支出"]) for m in monthly.values())
        tot_net = -tot_fee + tot_in - tot_out
        _write_row(ws, row, ["合计", tot_fee, tot_in, tot_out, tot_net], money_cols)
        for c in range(1, 6):
            ws.cell(row=row, column=c).font = Font(name="微软雅黑", bold=True, size=10)

    _auto_width(ws)


# ── 往来分类汇总 ────────────────────────────────────────────────────────

def _counterparty_summary(ws, bank_txns: list[BankTxn]) -> None:
    """按对方户名分类汇总"""
    from collections import defaultdict

    categories: dict[str, dict[str, Decimal]] = defaultdict(
        lambda: {"借方": ZERO, "贷方": ZERO, "笔数": ZERO}
    )

    for txn in bank_txns:
        cp = txn.counterparty.strip() or "（空）"
        categories[cp]["借方"] += txn.debit
        categories[cp]["贷方"] += txn.credit
        categories[cp]["笔数"] += Decimal("1")

    headers = ["对方户名", "借方发生额", "贷方发生额", "净额", "笔数", "分类"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    _style_header(ws, 2, len(headers))

    money_cols = {2, 3, 4}

    # 自动分类规则
    def classify(name: str) -> str:
        n = name.lower()
        if "财付通" in n or "微信" in n or "wechat" in n:
            return "平台认证款"
        if "阿里云" in n or "aliyun" in n:
            return "云服务商"
        if "腾讯" in n or "tencent" in n:
            return "云服务商"
        if "华为" in n or "huawei" in n:
            return "云服务商"
        if "银行" in n or "bank" in n or "金融" in n:
            return "金融机构"
        if "税" in n or "税务" in n or "国税" in n:
            return "税务"
        if "社保" in n or "公积金" in n:
            return "社保公积金"
        if "股东" in n or "借款" in n:
            return "股东借款"
        if "利息" in n:
            return "利息相关"
        return "其他"

    row = 3
    for cp in sorted(categories.keys()):
        c = categories[cp]
        net = float(c["贷方"] - c["借方"])
        _write_row(ws, row, [
            cp,
            float(c["借方"]) if c["借方"] > ZERO else "",
            float(c["贷方"]) if c["贷方"] > ZERO else "",
            net,
            int(c["笔数"]),
            classify(cp),
        ], money_cols)
        row += 1

    # 合计行
    tot_debit = sum(float(c["借方"]) for c in categories.values())
    tot_credit = sum(float(c["贷方"]) for c in categories.values())
    tot_net = tot_credit - tot_debit
    tot_cnt = sum(int(c["笔数"]) for c in categories.values())
    _write_row(ws, row, ["合计", tot_debit, tot_credit, tot_net, tot_cnt, ""], money_cols)
    for c in range(1, 7):
        ws.cell(row=row, column=c).font = Font(name="微软雅黑", bold=True, size=10)

    _auto_width(ws)


# ── 主函数 ──────────────────────────────────────────────────────────────

def write_bank_reconciliation(
    bank_txns: list,     # BankTxn 列表（从 icbc_statement.parse 来的）
    book_txns: list,     # 账面银行交易（从明细账解析来的）
    out_path: Path,
    opening_cash: Decimal | int | float | None = None,  # 企业账面银行存款期末余额
) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Sheet 1: 对账明细
    ws1 = wb.active
    ws1.title = "对账明细"
    ws1.cell(row=1, column=1, value="银行对账明细").font = TITLE_FONT
    _reconciliation_detail(ws1, bank_txns, book_txns)

    # Sheet 2: 未达账项
    ws2 = wb.create_sheet("未达账项")
    ws2.cell(row=1, column=1, value="未达账项").font = TITLE_FONT
    _outstanding_items(ws2, bank_txns, book_txns)

    # Sheet 3: 余额调节表
    ws3 = wb.create_sheet("余额调节表")
    ws3.cell(row=1, column=1, value="银行存款余额调节表").font = TITLE_FONT
    _balance_reconciliation(ws3, bank_txns, book_txns, opening_cash=opening_cash)

    # Sheet 4: 手续费利息统计
    ws4 = wb.create_sheet("手续费利息统计")
    ws4.cell(row=1, column=1, value="手续费与利息统计").font = TITLE_FONT
    _fee_interest_stats(ws4, bank_txns)

    # Sheet 5: 往来分类汇总
    ws5 = wb.create_sheet("往来分类汇总")
    ws5.cell(row=1, column=1, value="往来分类汇总").font = TITLE_FONT
    _counterparty_summary(ws5, bank_txns)

    wb.save(out_path)
    return out_path
