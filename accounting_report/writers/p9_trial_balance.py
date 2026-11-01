# -*- coding: utf-8 -*-
"""试算平衡表导出（体验官白班2轮新增 · 输出 9）

会计核心文件之一，此前产品缺失。从**会计分录**（而非科目余额表的
"本期发生额"列——该列含结转冲销负数，借贷合计不相等）汇总生成标准试算平衡表：

  科目编码 | 科目名称 | 期初借方 | 期初贷方 | 本期借方 | 本期贷方 | 期末借方 | 期末贷方

- 本期发生额：按科目逐笔累加分录的借贷方 → **借贷合计必然相等**（每笔分录借贷平衡）
- 期末余额：期初 + 本期净额，按自然方向归位（资产/费用→借方余额，负债/权益/收入→贷方余额）
- 底部输出借贷合计 + 试算平衡校验（期初/本期/期末 借贷必须相等）

输出：{year}-{month} 试算平衡表.xlsx（Sheet「试算平衡表」）
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook import Workbook

from ..models import ZERO, money
from ..rules import resolve_account

TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
CELL_FONT = Font(name="微软雅黑", size=10)
BOLD_FONT = Font(name="微软雅黑", bold=True, size=10)
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = '#,##0.00'
WARN_FILL = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
OK_FILL = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")


def _canonical(name: str) -> str:
    """科目名规范化：别名 → 标准科目名（银行存款_中国工商银行→货币资金 等）。

    p9 期初/期末余额按 BS 行次归位，但分录里的科目常用别名（财务软件导出、
    业务事项手填），若按原始字符串汇总流量，同一逻辑科目会被拆成多行
    （如 其他应付款 与 其他应付款_预付报销款 并存），且 BS 行 owner 科目
    （如 货币资金）拿不到别名科目上的流量，期末余额恒等于期初——试算平衡表
    每一行的余额都不对（总计仍平，但逐行误导）。这里统一解析为规范名再汇总。
    """
    if not name:
        return name
    try:
        return resolve_account(name).name
    except KeyError:
        return name


def generate(
    entries: list,
    trial_balances: list | None = None,
    prior=None,
    year: int = 0,
    month: int = 0,
    out_path: Path | str | None = None,
) -> Path:
    """生成试算平衡表 xlsx。

    Args:
        entries: JournalEntry 列表（分录）。
        trial_balances: TrialBalance 列表（科目余额表，用于科目编码/名称，可空）。
        prior: PriorPeriod（上月报表，期初余额按报表口径取，链跑时准确）。
        year/month: 账期。
        out_path: 输出路径（目录或文件）。
    """
    out_path = Path(out_path) if out_path else Path(".")
    if out_path.is_dir():
        prefix = f"{year:04d}-{month:02d} " if year else ""
        out_path = out_path / f"{prefix}试算平衡表.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # ── 1. 分录汇总：各科目 借方/贷方 本期发生额（借贷必然相等） ──
    #    科目名先做别名归一（_canonical），保证同一逻辑科目只出现一行，
    #    且 BS 行 owner 科目能拿到别名科目上的全部流量。
    debit_flow: dict[str, Decimal] = defaultdict(lambda: ZERO)
    credit_flow: dict[str, Decimal] = defaultdict(lambda: ZERO)
    for e in entries:
        dr_raw = str(getattr(e, "debit", "") or "").strip()
        cr_raw = str(getattr(e, "credit", "") or "").strip()
        amt = money(getattr(e, "amount", ZERO))
        if dr_raw:
            debit_flow[_canonical(dr_raw)] += amt
        if cr_raw:
            credit_flow[_canonical(cr_raw)] += amt

    # ── 2. 科目清单（分录科目 ∪ 余额表科目 ∪ 期初余额行次主科目） ──
    names = set(debit_flow) | set(credit_flow)
    tb_map = {}
    for tb in trial_balances or []:
        raw_name = str(getattr(tb, "account_name", "") or "")
        # 余额表科目名同样做别名归一（如 银行存款-工行 → 货币资金），
        # 避免与分录科目重复成行；解析失败（未定义科目）保持原名。
        name = _canonical(raw_name)
        tb_map[name] = tb
    names.update(tb_map.keys())

    prior_bs0 = getattr(prior, "bs_end", None) if prior is not None else None
    if prior_bs0:
        # 期初有余额但本期无发生额的科目也要进清单（如实收资本），否则期初不平
        from ..rules import chart_of_accounts
        accounts, _ = chart_of_accounts()
        for name0, acct in accounts.items():
            if acct.bs is not None and acct.bs in prior_bs0:
                names.add(name0)
                tb_map.setdefault(name0, None)

    # ── 3. 期初/期末余额（按自然方向归位） ──
    def _side(v: Decimal):
        """(借方, 贷方)：正→借方，负→贷方"""
        return (v, ZERO) if v >= ZERO else (ZERO, -v)

    rows = []
    totals = {k: ZERO for k in ("open_d", "open_c", "pd", "pc", "end_d", "end_c")}
    prior_bs = getattr(prior, "bs_end", None) if prior is not None else None

    # 期初按 BS 行次去重：同一行次可能有 汇总科目+明细科目（如 1002/100201），
    # 期初只赋给一个科目，避免重复累加导致借贷不平。
    # 明细科目（bs_sub 非空）映射「其中」子行次，主科目映射主行次，各自取期初。
    def _bs_row(acct) -> int | None:
        return acct.bs_sub if acct.bs_sub is not None else acct.bs

    owner_candidates: dict[int, list[str]] = {}
    for name in sorted(names):
        try:
            row = _bs_row(resolve_account(name))
        except KeyError:
            continue
        if row is not None:
            owner_candidates.setdefault(row, []).append(name)

    line_owner: dict = {}
    for row, cands in owner_candidates.items():
        # 主科目（bs_sub 为空）优先做 owner，明细科目不抢整行期初
        primary = next(
            (c for c in cands if resolve_account(c).bs_sub is None), None
        )
        line_owner[row] = primary if primary is not None else cands[0]

    for name in sorted(names):
        tb = tb_map.get(name)
        code = str(getattr(tb, "account_code", "") or "") if tb else ""
        pd = debit_flow.get(name, ZERO)
        pc = credit_flow.get(name, ZERO)

        # 科目方向
        try:
            acct = resolve_account(name)
            natural_debit = acct.type in ("asset", "expense")
            bs_line = _bs_row(acct)
        except KeyError:
            natural_debit = True
            bs_line = None

        # 期初余额：优先上月报表对应行次（链跑准确，按行次去重），否则余额表文件值
        if prior_bs is not None and bs_line is not None:
            if line_owner.get(bs_line) == name:
                open_net = money(prior_bs.get(bs_line, ZERO))
                od_, oc_ = _side(open_net) if natural_debit else _side(-open_net)
                open_d, open_c = od_, oc_
            else:
                open_d = open_c = ZERO  # 明细科目：期初已在汇总科目体现，不重复
        else:
            open_d = money(getattr(tb, "open_debit", ZERO)) if tb else ZERO
            open_c = money(getattr(tb, "open_credit", ZERO)) if tb else ZERO

        # 期末余额 = 期初 + 本期净额。方向：借方余额科目 net 正→借方；
        # 贷方余额科目（负债/权益/收入）net 负→贷方，_side(net) 天然处理符号
        net = (open_d - open_c) + (pd - pc)
        end_d, end_c = _side(net)

        totals["open_d"] += open_d
        totals["open_c"] += open_c
        totals["pd"] += pd
        totals["pc"] += pc
        totals["end_d"] += end_d
        totals["end_c"] += end_c
        rows.append((code, name, open_d, open_c, pd, pc, end_d, end_c))

    # ── 4. 渲染 ──
    wb = Workbook()
    ws = wb.active
    ws.title = "试算平衡表"

    ws.merge_cells("A1:H1")
    c = ws.cell(1, 1, f"{year}年{month}月 试算平衡表" if year else "试算平衡表")
    c.font = TITLE_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["科目编码", "科目名称", "期初借方", "期初贷方", "本期借方",
               "本期贷方", "期末借方", "期末贷方"]
    for j, h in enumerate(headers, 1):
        cell = ws.cell(3, j, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    money_cols = {3, 4, 5, 6, 7, 8}
    row = 4
    for code, name, od, oc, pd, pc, ed, ec in rows:
        vals = [code, name, od, oc, pd, pc, ed, ec]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(row, j, float(v) if isinstance(v, Decimal) and j >= 3 else v)
            cell.border = BORDER
            cell.font = CELL_FONT
            cell.alignment = Alignment(vertical="center")
            if j in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = MONEY_FMT
        row += 1

    ws.cell(row, 1, "").border = BORDER
    ws.cell(row, 2, "合计").font = BOLD_FONT
    ws.cell(row, 2).border = BORDER
    for j, key in enumerate(("open_d", "open_c", "pd", "pc", "end_d", "end_c"), 3):
        cell = ws.cell(row, j, float(totals[key]))
        cell.number_format = MONEY_FMT
        cell.font = BOLD_FONT
        cell.border = BORDER
    row += 2

    checks = [
        ("期初余额", totals["open_d"], totals["open_c"]),
        ("本期发生额", totals["pd"], totals["pc"]),
        ("期末余额", totals["end_d"], totals["end_c"]),
    ]
    ws.cell(row, 1, "试算平衡校验").font = BOLD_FONT
    row += 1
    all_ok = True
    for label, d, c_ in checks:
        ok = d == c_
        all_ok = all_ok and ok
        cell = ws.cell(row, 1,
                       f"{label}：借方 {d:,.2f} ＝ 贷方 {c_:,.2f}  {'✓ 平衡' if ok else '✗ 不平！'}")
        cell.font = BOLD_FONT if ok else Font(name="微软雅黑", bold=True, size=10, color="C62828")
        cell.fill = OK_FILL if ok else WARN_FILL
        row += 1
    ws.cell(row, 1, "✓ 试算平衡表借贷平衡" if all_ok
            else "✗ 试算不平衡，请核查数据！").font = BOLD_FONT

    for col, w in zip("ABCDEFGH", (12, 28, 14, 14, 14, 14, 14, 14)):
        ws.column_dimensions[col].width = w

    wb.save(out_path)
    return out_path
