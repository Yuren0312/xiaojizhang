# -*- coding: utf-8 -*-
"""规则加载：把 YAML 读成程序用的结构，并做一致性校验"""
from __future__ import annotations

import functools

import yaml

from ..config import RULES_DIR
from ..models import Account


def _load(name: str) -> dict:
    with open(RULES_DIR / name, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


@functools.lru_cache(maxsize=1)
def chart_of_accounts() -> tuple[dict[str, Account], dict[str, str]]:
    """返回 (科目名 → Account, 别名 → 正式科目名)"""
    raw = _load("accounts.yaml")
    accounts: dict[str, Account] = {}
    for name, spec in raw["accounts"].items():
        spec = dict(spec or {})
        accounts[name] = Account(
            name=name,
            type=spec.get("type", "asset"),
            bs=spec.get("bs"),
            bs_sub=spec.get("bs_sub"),
            is_=spec.get("is"),
            is_sub=spec.get("is_sub"),
            cash=bool(spec.get("cash", False)),
            cf_in=spec.get("cf_in"),
            cf_out=spec.get("cf_out"),
        )
    aliases = dict(raw.get("aliases") or {})
    # 别名指向的科目必须存在
    for alias, target in aliases.items():
        if target not in accounts:
            raise ValueError(f"accounts.yaml 别名 {alias!r} 指向未定义的科目 {target!r}")
    return accounts, aliases


def resolve_account(name: str) -> Account:
    """科目名（含别名）→ Account，找不到就报错，绝不猜"""
    accounts, aliases = chart_of_accounts()
    key = (name or "").strip()
    key = aliases.get(key, key)
    if key not in accounts:
        raise KeyError(
            f"未定义的科目：{name!r}。"
            f"请在 accounting_report/rules/accounts.yaml 中补充，或使用已有科目。"
        )
    return accounts[key]


@functools.lru_cache(maxsize=1)
def balance_sheet_rules() -> dict:
    return _load("balance_sheet.yaml")


@functools.lru_cache(maxsize=1)
def income_statement_rules() -> dict:
    return _load("income_statement.yaml")


@functools.lru_cache(maxsize=1)
def cash_flow_rules() -> dict:
    return _load("cash_flow.yaml")


@functools.lru_cache(maxsize=1)
def bank_mapping() -> dict:
    return _load("mapping.yaml")


@functools.lru_cache(maxsize=1)
def layout() -> dict:
    """五张表的行次 ↔ 单元格坐标（由 tools/build_template.py 生成）"""
    return _load("layout.yaml")
