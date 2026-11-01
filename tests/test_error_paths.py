# -*- coding: utf-8 -*-
"""失败路径：出错时必须给出对方看得懂、照着就能改的提示

对方不写代码，报错信息就是唯一的说明书。这里锁定几条关键提示的内容，
避免以后重构时把提示改烂。
"""
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from accounting_report.engine.journalize import UnmappedTransaction, from_bank  # noqa: E402
from accounting_report.parsers.icbc_statement import BankTxn  # noqa: E402
from accounting_report.rules import resolve_account  # noqa: E402


def _txn(summary: str, counterparty: str = "", credit="0.00", debit="0.00") -> BankTxn:
    return BankTxn(date="2026-06-10", summary=summary, counterparty=counterparty,
                   debit=Decimal(debit), credit=Decimal(credit),
                   balance=Decimal("0.00"), source="test.pdf")


def test_unmapped_transaction_lists_the_offender():
    """摘要匹配不上时，要把是哪笔流水、该改哪个文件说清楚"""
    with pytest.raises(UnmappedTransaction) as e:
        from_bank([_txn("某某科技杂费", "某某科技有限公司", credit="5000.00")])

    msg = str(e.value)
    assert "mapping.yaml" in msg          # 告诉他改哪个文件
    assert "bank_rules" in msg            # 告诉他改哪一节
    assert "某某科技杂费" in msg           # 告诉他是哪笔
    assert "2026-06-10" in msg


def test_sales_payment_maps_to_revenue():
    """销售货款流水应识别为营业收入（原缺陷：只能归营业外收入/报未映射）"""
    entries = from_bank([_txn("某某科技货款", "某某科技有限公司", credit="5000.00")])
    assert len(entries) == 1
    assert entries[0].debit == "货币资金"
    assert entries[0].credit == "营业收入"
    assert entries[0].amount == Decimal("5000.00")


def test_known_summary_is_mapped():
    """已配规则的摘要能正常入账，方向也对：进账 = 借货币资金"""
    entries = from_bank([_txn("微信支付认证7281", "财付通支付科技有限公司",
                              credit="0.01")])
    assert len(entries) == 1
    assert entries[0].debit == "货币资金"
    assert entries[0].credit == "财务费用-利息"


def test_outflow_direction():
    """出账 = 贷货币资金"""
    entries = from_bank([_txn("阿里云服务费", "阿里云计算有限公司", debit="115.00")])
    assert len(entries) == 1
    assert entries[0].debit == "营业成本"
    assert entries[0].credit == "货币资金"


def test_unknown_account_is_rejected():
    """业务事项表里写了没定义的科目，要报错并指路，不能猜一个近似的"""
    with pytest.raises(KeyError) as e:
        resolve_account("研发费用")  # 正确科目名是「研究费用」
    assert "accounts.yaml" in str(e.value)


def test_account_alias_works():
    assert resolve_account("银行存款").name == "货币资金"
    assert resolve_account("主营业务收入").name == "营业收入"
