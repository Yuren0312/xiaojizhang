# -*- coding: utf-8 -*-
"""路径与全局配置"""
import os
from pathlib import Path

PKG_DIR = Path(__file__).resolve().parent

# 数据根目录（输入/输出/归档 所在位置）
# - 开发模式：项目根目录（accounting_report 的上一级）
# - 桌面版：由 Electron 通过环境变量 XJZ_DATA_DIR 指定用户可写目录
_ENV_ROOT = os.environ.get("XJZ_DATA_DIR")
if _ENV_ROOT:
    ROOT = Path(_ENV_ROOT)
else:
    ROOT = Path(__file__).resolve().parent.parent

INPUT_DIR = ROOT / "输入"
OUTPUT_DIR = ROOT / "输出"

# 内置规则（随程序分发，只读）
_BUNDLED_RULES_DIR = PKG_DIR / "rules"
# 运行时规则（用户可编辑，位于数据目录，首次启动时从内置规则复制）
_RUNTIME_RULES_DIR = ROOT / "accounting_report" / "rules"

RULES_DIR = _RUNTIME_RULES_DIR if _RUNTIME_RULES_DIR.exists() else _BUNDLED_RULES_DIR
TEMPLATE_PATH = PKG_DIR / "templates" / "报表模板.xlsx"

# 报表表头的企业名称。换一家主体只改这一行，模板不用动。
COMPANY_NAME = "示例科技有限公司"

# 上期报表 PDF 目录（首次使用时放上月 5 表 PDF 的位置）。
# 设为 None 则不启用此 fallback，仅依赖输出目录中的上月 xlsx。
PRIOR_REPORT_DIR = None

# 五张表的 sheet 名
SHEET_BS = "资产负债表"
SHEET_IS = "利润表"
SHEET_CF = "现金流量表"
SHEET_IS_Q = "利润表季报"
SHEET_CF_Q = "现金流量表季报"

SHEETS = [SHEET_BS, SHEET_IS, SHEET_CF, SHEET_IS_Q, SHEET_CF_Q]

# 金额保留位数（元至角分）
MONEY_DP = 2
