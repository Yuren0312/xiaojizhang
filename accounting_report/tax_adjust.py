# -*- coding: utf-8 -*-
"""税务调整项：纳税调增/调减、弥补亏损、减免税额、已预缴税额。

为什么不做 Excel 解析
---------------------
预缴税款台账、纳税调整台账、税收优惠台账这三样**没有任何统一格式** ——
每家代账公司、每个老会计的表格都不一样，强行解析只会把数字读错，
而税表读错一位数就是申报错误。所以这里的定位是：
**数字由会计确认后录入，程序只负责代入计算和校验勾稽关系。**

数据来源（后者覆盖前者）
------------------------
1. ``输入/<账期>/税务调整.json``          —— 机器可读，供以后对接台账
2. ``输入/<账期>/_tax-adjust_补充说明.txt`` —— 前端「税务调整」页保存的文本

文本格式
--------
每行一条「项目 + 金额」，分隔符可以是空格、冒号、等号、全角冒号；
金额支持千分位、¥ 符号、括号负数、末尾「元」。 ``#`` 开头为注释。

::

    纳税调增额：12,000
    纳税调减额 3000
    弥补以前年度亏损=0
    已预缴所得税额 ¥45,000.00
    已缴增值税额：(1,200)      # 括号表示负数

未出现的项目一律按 0 处理 —— 与税局表「无则填 0」的要求一致。
无法识别的行不会静默丢弃，会记进 ``problems`` 供界面提示。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

ZERO = Decimal("0")

# 文本补充说明的模块名（与前端 modules.ts 的 id 对应）
TEXT_MODULE = "tax-adjust"
TEXT_FILENAME = f"_{TEXT_MODULE}_补充说明.txt"
JSON_FILENAME = "税务调整.json"

# 字段 -> 可识别的中文别名（先长后短，避免「调增」抢掉「纳税调增额」）
_ALIASES: dict[str, tuple[str, ...]] = {
    "increase": ("纳税调增额", "应纳税所得额调增", "调增额", "调增"),
    "decrease": ("纳税调减额", "应纳税所得额调减", "调减额", "调减"),
    "loss_offset": ("弥补以前年度亏损", "弥补以前年度亏损额", "弥补亏损", "亏损弥补"),
    "other_exemption": ("其他减免所得税额", "其他减免税额", "减免所得税额", "减免税额"),
    "income_tax_prepaid": ("已预缴所得税额", "已预缴所得税", "预缴所得税额", "已缴所得税额", "已缴所得税"),
    "vat_prepaid": ("已缴增值税额", "已预缴增值税额", "已缴增值税", "预缴增值税额", "已缴税额"),
    "prior_overpaid": ("以前年度多缴在本期抵缴", "以前年度多缴抵缴", "多缴抵缴"),
}

# 金额：可选负数，可选千分位，可选括号负数
_NUM = r"(?:\(?\s*-?\s*[\d,]+(?:\.\d+)?\s*\)?)"
_LINE_RE = re.compile(
    rf"^(?P<name>.*?)\s*[:：=]?\s*[¥￥]?\s*(?P<num>{_NUM})\s*(?:元)?\s*$"
)


@dataclass
class TaxAdjust:
    """税务调整项。全部为 Decimal，未填写 = 0。"""

    increase: Decimal = ZERO           # 纳税调增额
    decrease: Decimal = ZERO           # 纳税调减额
    loss_offset: Decimal = ZERO        # 弥补以前年度亏损
    other_exemption: Decimal = ZERO    # 其他减免所得税额
    income_tax_prepaid: Decimal = ZERO  # 已预缴所得税额（本年累计）
    vat_prepaid: Decimal = ZERO        # 已缴增值税额（本年累计）
    prior_overpaid: Decimal = ZERO     # 以前年度多缴在本期抵缴
    source: str = ""                   # 数据来自哪个文件（空 = 未提供）
    problems: list[str] = field(default_factory=list)  # 无法识别的行

    @property
    def is_empty(self) -> bool:
        """七个数字全 0 且没有来源文件。"""
        return not self.source

    @property
    def taxable_adjustment(self) -> Decimal:
        """应纳税所得额净调整 = 调增 − 调减 − 弥补亏损。"""
        return self.increase - self.decrease - self.loss_offset

    def as_dict(self) -> dict[str, str]:
        return {
            "纳税调增额": str(self.increase),
            "纳税调减额": str(self.decrease),
            "弥补以前年度亏损": str(self.loss_offset),
            "其他减免所得税额": str(self.other_exemption),
            "已预缴所得税额": str(self.income_tax_prepaid),
            "已缴增值税额": str(self.vat_prepaid),
            "以前年度多缴在本期抵缴": str(self.prior_overpaid),
            "来源": self.source or "（未提供，按 0 计）",
        }


def _to_decimal(raw: str) -> Decimal | None:
    """把 '1,200.50' / '(1,200)' / '¥300' 转 Decimal，失败返回 None。"""
    s = raw.strip().replace("¥", "").replace("￥", "").replace(",", "").replace(" ", "")
    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]
    try:
        d = Decimal(s)
    except (InvalidOperation, ValueError):
        return None
    return -d if negative else d


def _match_field(name: str) -> str | None:
    """把「纳税调增额（业务招待费）」这类名字映射到字段名。"""
    compact = re.sub(r"[\s　()（）\[\]【】:：]", "", name or "")
    if not compact:
        return None
    # 长别名优先，避免「调增」命中「纳税调增额」的歧义（同一字段，无害；
    # 但「已缴增值税额」与「已缴所得税额」必须靠长度区分）
    candidates = sorted(
        ((alias, fld) for fld, aliases in _ALIASES.items() for alias in aliases),
        key=lambda x: len(x[0]),
        reverse=True,
    )
    for alias, fld in candidates:
        if alias in compact:
            return fld
    return None


def parse_text(text: str) -> tuple[dict[str, Decimal], list[str]]:
    """解析补充说明文本，返回 (字段->金额, 无法识别的行)。"""
    values: dict[str, Decimal] = {}
    problems: list[str] = []
    for raw_line in (text or "").splitlines():
        line = raw_line.strip().lstrip("　")
        if not line or line.startswith("#") or line.startswith("//"):
            continue
        m = _LINE_RE.match(line)
        if not m:
            problems.append(f"无法解析：{line}")
            continue
        fld = _match_field(m.group("name"))
        if fld is None:
            problems.append(f"未知项目：{line}")
            continue
        amount = _to_decimal(m.group("num"))
        if amount is None:
            problems.append(f"金额无法识别：{line}")
            continue
        values[fld] = amount
    return values, problems


def load_from_text_file(path: Path) -> TaxAdjust:
    """从一个补充说明 txt 构造 TaxAdjust（文件不存在则返回空白对象）。"""
    adj = TaxAdjust()
    try:
        text = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return adj
    values, problems = parse_text(text)
    for key, val in values.items():
        setattr(adj, key, val)
    adj.problems = problems
    adj.source = Path(path).name
    return adj


def load_from_json_file(path: Path) -> TaxAdjust:
    """从 JSON 构造 TaxAdjust。键名支持英文与中文两种写法。"""
    adj = TaxAdjust()
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return adj
    if not isinstance(data, dict):
        return adj
    cn_to_en = {alias: fld for fld, aliases in _ALIASES.items() for alias in aliases}
    for key, raw in data.items():
        fld = key if key in _ALIASES else cn_to_en.get(str(key).strip())
        if fld is None:
            adj.problems.append(f"未知项目：{key}")
            continue
        amount = _to_decimal(str(raw))
        if amount is None:
            adj.problems.append(f"金额无法识别：{key}={raw}")
            continue
        setattr(adj, fld, amount)
    adj.source = Path(path).name
    return adj


def load(period_dir: Path) -> TaxAdjust:
    """读取某账期目录下的税务调整项。

    优先 JSON（结构化、字段明确），其次前端文本框保存的 txt。
    两者都没有时返回``is_empty``为 True 的对象，调用方据此提示「未提供」。
    """
    period_dir = Path(period_dir)
    json_path = period_dir / JSON_FILENAME
    if json_path.exists():
        return load_from_json_file(json_path)
    return load_from_text_file(period_dir / TEXT_FILENAME)
