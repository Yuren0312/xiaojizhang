# -*- coding: utf-8 -*-
"""工商银行企业存款对账单 PDF → 现金流水

以日期行为锚：找全部日期行 → 相邻日期行之间的空间归入该笔 → 按列组装。
支持日期拆行、对方户名多行折行、表头合并词等多种版式。

核心改进（2026-08）：
- 多页读取：读取 PDF 全部页面，后续页无表头时整页作为数据。
- 动态列检测：
  - 文本列（日期/对方户名/摘要等）通过表头词位置检测，各列边界取
    相邻表头 x 坐标的中点，兼容不同月份 PDF 列宽差异。
  - 金额列（借方/贷方/余额）通过聚类 amount-pattern 词语的 x 坐标
    自动检测，兼容金额列左右偏移。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from ..models import money, ZERO

DATE_RE = re.compile(r"^(20\d{2})-(\d{1,2})-(\d{1,2})$")
AMOUNT_RE = re.compile(r"^[\d,]+\.\d{2}$")
# 拆行日期的两种模式
DATE_YEAR_RE = re.compile(r"^(20\d{2})-$")         # "2025-"
DATE_MONTHDAY_RE = re.compile(r"^\d{1,2}-\d{1,2}$") # "12-18" / "2-5"
DATE_PREFIX_RE = re.compile(r"^(20\d{2})-\d{1,2}-$")  # "2026-02-" / "2026-2-"
DATE_DAY_RE = re.compile(r"^\d{1,2}$")              # "06"

# 金额列聚类间距（pt）：x 坐标差小于此值视为同一列
_CLUSTER_GAP = 15


@dataclass
class BankTxn:
    date: str
    summary: str
    counterparty: str
    debit: Decimal
    credit: Decimal
    balance: Decimal
    source: str = ""

    @property
    def inflow(self) -> Decimal:
        return self.credit

    @property
    def outflow(self) -> Decimal:
        return self.debit


def _x_c(w) -> float:
    return (float(w[0]) + float(w[2])) / 2


def _y_c(w) -> float:
    return (float(w[1]) + float(w[3])) / 2


def _money(text: str) -> Decimal:
    text = text.strip().replace(",", "").replace("，", "")
    try:
        return money(text)
    except Exception:
        return ZERO


def parse(pdf_path: str | Path) -> list[BankTxn]:
    try:
        return _parse_inner(pdf_path)
    except Exception as e:
        raise ValueError(
            f"银行对账单解析失败（{Path(pdf_path).name}）：{e}。"
            f"PDF 版式可能与预期不一致，请检查列对齐。"
        ) from e


# ------------------------------------------------------------------ #
#  表头列检测
# ------------------------------------------------------------------ #

# 需要跳过的表头行噪声词（账户信息、元数据等）
_HEADER_SKIP = ("：", "北京", "网点", "0045", "币种", "人民币")


def _detect_header_cols(
    header_band: list[tuple],
) -> tuple[list[tuple[float, str]], list[tuple[float, str]]]:
    """从表头词中检测列位置。

    返回 (text_headers, amount_headers)，均为 (xc, role) 列表按 x 排序。
    text_headers 中的 role: date / trade / voucher_type / voucher_no /
    cp_account / counterparty / summary
    amount_headers 中的 role: debit / credit / balance
    """
    text_headers: list[tuple[float, str]] = []
    amount_headers: list[tuple[float, str]] = []

    for w in header_band:
        txt = w[4].strip()
        xc = _x_c(w)
        if any(s in txt for s in _HEADER_SKIP):
            continue
        role = None
        if "日期" in txt:
            role = "date"
        elif "交易" in txt:
            role = "trade"
        elif "凭证种类" in txt or "凭证种" in txt:
            role = "voucher_type"
        elif "凭证号" in txt:
            role = "voucher_no"
        elif "对方账号" in txt:
            role = "cp_account"
        elif "对方" in txt:
            role = "counterparty"
        elif "摘要" in txt:
            role = "summary"
        elif "借方" in txt:
            role = "debit"
        elif "贷方" in txt:
            role = "credit"
        elif "余额" in txt:
            role = "balance"
        if role is None:
            continue
        if role in ("debit", "credit", "balance"):
            amount_headers.append((xc, role))
        else:
            text_headers.append((xc, role))

    text_headers.sort()
    amount_headers.sort()
    return text_headers, amount_headers


def _build_text_zones(
    text_headers: list[tuple[float, str]],
    amt_start: float,
) -> list[tuple[str, float, float]]:
    """从表头位置构建文本列区间 (role, lo, hi)。

    最后一个文本列的右边界为 amt_start（金额区域左边界）。
    """
    zones: list[tuple[str, float, float]] = []
    for i, (xc, role) in enumerate(text_headers):
        lo = 0.0 if i == 0 else (text_headers[i - 1][0] + xc) / 2
        if i < len(text_headers) - 1:
            hi = (xc + text_headers[i + 1][0]) / 2
        else:
            hi = amt_start
        zones.append((role, lo, hi))
    return zones


# ------------------------------------------------------------------ #
#  金额列检测（聚类）
# ------------------------------------------------------------------ #

def _cluster_amounts(
    data: list[tuple],
) -> tuple[list[tuple[float, float]], bool]:
    """聚类金额词的 x 坐标，返回 (列边界列表, 是否借方贷方分离)。

    列边界列表: [(lo, hi), ...]，从左到右依次为借方、贷方、余额。
    """
    amount_xcs = sorted(set(round(_x_c(w), 1) for w in data
                           if AMOUNT_RE.match(w[4].strip())))
    clusters: list[list[float]] = []
    for x in amount_xcs:
        if clusters and x - clusters[-1][-1] < _CLUSTER_GAP:
            clusters[-1].append(x)
        else:
            clusters.append([x])

    if len(clusters) >= 3:
        d0, d1, d2 = clusters[0], clusters[1], clusters[2]
        deb_hi = (d0[-1] + d1[0]) / 2
        cre_hi = (d1[-1] + d2[0]) / 2
        return [(d0[0] - 10, deb_hi),
                (deb_hi, cre_hi),
                (cre_hi, 9999.0)], True
    if len(clusters) == 2:
        mid = (clusters[0][-1] + clusters[1][0]) / 2
        return [(clusters[0][0] - 10, mid),
                (mid, 9999.0)], False
    # 回退
    return [(380.0, 9999.0)], False


# ------------------------------------------------------------------ #
#  主解析
# ------------------------------------------------------------------ #

def _parse_inner(pdf_path: str | Path) -> list[BankTxn]:
    import fitz

    pdf_path = Path(pdf_path)
    doc = fitz.open(str(pdf_path))

    # --- 读取全部页面 ---
    all_data: list[tuple] = []
    header_band: list[tuple] = []

    for page_idx in range(doc.page_count):
        page = doc[page_idx]
        raw = [(float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4]))
               for w in page.get_text("words")]
        y_off = page_idx * 1000.0

        # 找本页表头
        page_header_lo = 0.0
        for w in raw:
            if "借方" in w[4]:
                page_header_lo = max(page_header_lo, _y_c(w) + 3)

        if page_idx == 0:
            if page_header_lo == 0:
                doc.close()
                raise ValueError(f"{pdf_path.name} 找不到表头")
            header_band = [w for w in raw
                          if page_header_lo - 35 < _y_c(w) <= page_header_lo]

        # 收集数据词（过滤页脚噪声）
        def _is_footer(txt: str) -> bool:
            txt = txt.strip()
            if "工商银行" in txt or "版权所有" in txt:
                return True
            if re.match(r"^\d{1,2}:\d{2}:\d{2}$", txt):
                return True
            return False

        if page_header_lo > 0:
            for w in raw:
                if _y_c(w) > page_header_lo and not _is_footer(w[4]):
                    all_data.append((w[0], w[1] + y_off,
                                    w[2], w[3] + y_off, w[4]))
        else:
            # 后续页无表头：整页作为数据
            for w in raw:
                if not _is_footer(w[4]):
                    all_data.append((w[0], w[1] + y_off,
                                    w[2], w[3] + y_off, w[4]))

    doc.close()

    if not all_data:
        raise ValueError(f"{pdf_path.name} 无可提取文本")

    # --- 金额列聚类 ---
    amt_bounds, separate_dc = _cluster_amounts(all_data)
    amt_start = amt_bounds[0][0]

    # --- 表头列检测 ---
    text_headers, _ = _detect_header_cols(header_band)

    # 如果表头检测失败，使用回退值
    if not text_headers:
        text_headers = [(47.0, "date"), (340.0, "counterparty"),
                       (390.0, "summary")]

    text_zones = _build_text_zones(text_headers, amt_start)

    # --- 列判定函数 ---
    def col_of(xc: float) -> str:
        # 金额区域
        if xc >= amt_start:
            if separate_dc:
                if amt_bounds[0][0] <= xc < amt_bounds[0][1]:
                    return "debit"
                if amt_bounds[1][0] <= xc < amt_bounds[1][1]:
                    return "credit"
            else:
                if amt_bounds[0][0] <= xc < amt_bounds[0][1]:
                    return "combined"
            if amt_bounds[-1][0] <= xc < amt_bounds[-1][1]:
                return "balance"
            return "other"
        # 文本区域
        for role, lo, hi in text_zones:
            if lo <= xc < hi:
                return role
        return "other"

    # --- 日期锚点 ---
    # 两步法：先按 y 紧分组（5pt），再按模式合并拆行日期
    _TIGHT_Y = 5  # 同视觉行的 y 容差

    date_words: list[tuple[float, tuple]] = []
    for w in all_data:
        if col_of(_x_c(w)) != "date":
            continue
        txt = w[4].strip()
        if not txt:
            continue
        date_words.append((_y_c(w), w))
    date_words.sort(key=lambda x: x[0])

    # Step 1: 紧分组
    tight_groups: list[list[tuple]] = []  # [[w, w, ...], ...]
    for yc, w in date_words:
        if tight_groups and abs(yc - _y_c(tight_groups[-1][0])) <= _TIGHT_Y:
            tight_groups[-1].append(w)
        else:
            tight_groups.append([w])

    # Step 2: 模式合并拆行日期
    # 模式1："2025-" + "12-18" → "2025-12-18"
    # 模式2："2026-02-" + "06" → "2026-02-06"
    anchors: list[tuple[float, str]] = []
    i = 0
    while i < len(tight_groups):
        group = tight_groups[i]
        text = "".join(w[4] for w in sorted(group, key=lambda w: w[1]))
        ay = _y_c(group[0])

        # 模式1："YYYY-" + "MM-DD"
        if DATE_YEAR_RE.match(text) and i + 1 < len(tight_groups):
            next_text = "".join(w[4] for w in sorted(
                tight_groups[i + 1], key=lambda w: w[1]))
            if DATE_MONTHDAY_RE.match(next_text):
                anchors.append((ay, text + next_text))
                i += 2
                continue

        # 模式2："YYYY-MM-" + "DD"
        if DATE_PREFIX_RE.match(text) and i + 1 < len(tight_groups):
            next_text = "".join(w[4] for w in sorted(
                tight_groups[i + 1], key=lambda w: w[1]))
            if DATE_DAY_RE.match(next_text):
                anchors.append((ay, text + next_text))
                i += 2
                continue

        # 完整日期或不可识别的片段
        m = DATE_RE.match(text)
        if m:
            # 统一补零为 YYYY-MM-DD：单数字月日版式（如 "2026-2-5"）若不
            # 规范化，与分录日期（"2026-02-05"）字符串比对将永不匹配
            anchors.append((ay, f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"))
        # 否则丢弃不完整片段（如孤立的 "2025-"）
        i += 1

    if not anchors:
        raise ValueError(f"{pdf_path.name} 无日期行")

    max_y = max(_y_c(w) for w in all_data)

    # --- 每笔流水 ---
    txns: list[BankTxn] = []
    for idx, (ay, date_str) in enumerate(anchors):
        lo = ay - 5
        hi = anchors[idx + 1][0] - 5 if idx + 1 < len(anchors) else max_y + 10

        zone_words = [w for w in all_data if lo <= _y_c(w) <= hi]

        # --- 金额 ---
        if separate_dc:
            debit_words = [w[4].strip() for w in zone_words
                          if col_of(_x_c(w)) == "debit"
                          and AMOUNT_RE.match(w[4].strip())]
            credit_words = [w[4].strip() for w in zone_words
                           if col_of(_x_c(w)) == "credit"
                           and AMOUNT_RE.match(w[4].strip())]
            debit = _money(debit_words[-1]) if debit_words else ZERO
            credit = _money(credit_words[-1]) if credit_words else ZERO
        else:
            combined_words = [w[4].strip() for w in zone_words
                             if col_of(_x_c(w)) == "combined"
                             and AMOUNT_RE.match(w[4].strip())]
            cp_text = "".join(w[4] for w in zone_words
                             if col_of(_x_c(w)) == "counterparty"
                             and not AMOUNT_RE.match(w[4].strip()))
            sm_text = "".join(w[4] for w in zone_words
                             if col_of(_x_c(w)) == "summary"
                             and not AMOUNT_RE.match(w[4].strip()))
            combined = cp_text + sm_text
            if combined_words:
                if any(kw in combined for kw in ("利息", "报销", "借款", "认证",
                                                  "退款", "投资", "注册",
                                                  "打款", "支付宝")):
                    credit = _money(combined_words[-1])
                    debit = ZERO
                else:
                    debit = _money(combined_words[-1])
                    credit = ZERO
            else:
                debit = ZERO
                credit = ZERO

        balance_words = [w[4].strip() for w in zone_words
                        if col_of(_x_c(w)) == "balance"
                        and AMOUNT_RE.match(w[4].strip())]
        balance = _money(balance_words[-1]) if balance_words else ZERO

        # --- 摘要 ---
        summary = "".join(w[4] for w in zone_words
                         if col_of(_x_c(w)) == "summary"
                         and not AMOUNT_RE.match(w[4].strip()))

        # --- 对方户名 ---
        counterparty = "".join(w[4] for w in zone_words
                              if col_of(_x_c(w)) == "counterparty"
                              and not AMOUNT_RE.match(w[4].strip()))

        txns.append(BankTxn(
            date=date_str,
            summary=summary,
            counterparty=counterparty,
            debit=debit,
            credit=credit,
            balance=balance,
            source=pdf_path.name,
        ))

    if not txns:
        raise ValueError(f"{pdf_path.name} 未解析出任何流水")

    _reconcile_zero_amount_by_balance_continuity(txns)
    return txns


# ------------------------------------------------------------------ #
#  余额连续性兜底（修复版式边界：PDF 借方/贷方列空白但余额隐含动账）
# ------------------------------------------------------------------ #

# 摘要里不该出现 "按余额连续性推断" 的标记关键字，避免循环
_BALANCE_INFERRED_TAG = "[按余额连续性推断动账]"


def _reconcile_zero_amount_by_balance_continuity(
    txns: list[BankTxn],
    max_infer: Decimal = Decimal("100000"),
) -> None:
    """对借贷均为 0 的行检查余额变化，按差额补借/贷并标记推断来源。

    适用场景：银行对账单 PDF 版式中部分行的金额列打印空白（如「借方 -
    贷方 0.00 余额 98,160.16」），但通过余额连续性可知实际发生了 0.34 动账。
    不补足会导致该笔动账在对账报告中被标为「零金额行（非动账）」，
    余额调节表结论偏差。

    防误判：金额阈值 max_infer；连续同源；跳过两笔都为零的零金额行；
    同方向标记摘要便于对账报告区分。
    """
    prev: BankTxn | None = None
    for txn in txns:
        if prev is not None and txn.debit == ZERO and txn.credit == ZERO \
                and prev.balance != ZERO and txn.balance != ZERO:
            diff = txn.balance - prev.balance
            if diff != ZERO and abs(diff) <= max_infer \
                    and _BALANCE_INFERRED_TAG not in txn.summary:
                if diff < ZERO:
                    # 余额下降 → 借方(支出)
                    txn.debit = -diff
                    txn.summary = (
                        f"{txn.summary}{_BALANCE_INFERRED_TAG}" if txn.summary
                        else _BALANCE_INFERRED_TAG
                    )
                else:
                    # 余额上升 → 贷方(收入)
                    txn.credit = diff
                    txn.summary = (
                        f"{txn.summary}{_BALANCE_INFERRED_TAG}" if txn.summary
                        else _BALANCE_INFERRED_TAG
                    )
        prev = txn
