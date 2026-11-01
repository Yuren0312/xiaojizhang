# 参与贡献

欢迎提交 Issue 与 PR。这个项目面向真实记账场景，**数据正确性优先于功能数量**。

## 开发环境

```bash
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt

npm --prefix web-ui install
npm --prefix web-ui run build
npm start
```

## 提交前请自检

```bash
# 后端回归（含解析、分录、报表勾稽）
python -m pytest tests/ -q

# 前端类型检查与构建
npm --prefix web-ui run build
```

测试必须全绿再提 PR。若你的改动涉及报表输出，请在 PR 描述里说明**改了哪张表、影响哪些单元格**。

## 代码约定

- **金额一律用 `Decimal`**，禁止 `float`。会计报表要求到分，`0.1 + 0.2` 的浮点误差会直接让表不平。
- 新增科目映射改 YAML（`accounting_report/rules/`），不要硬编码在 Python 里。
- 解析器只做「把文件变成结构化数据」，业务判断放在分录/规则层，别混在一起。
- 中文注释可以直接写，代码标识符用英文。

## 千万不要提交

- 真实账套、银行回单、工资表、发票台账（`输入/` `输出/` 已在 `.gitignore` 里）
- 任何 API Key、Token、证书、`.env`
- 企业真实名称、纳税人识别号、个人信息

需要复现问题请用 `tools/generate_sample_data.py` 生成合成数据。

## 提交信息

```
<type>: <一句话说明>

type: feat | fix | refactor | docs | test | chore | build
```

例：`fix: 现金流量表经营项下「支付其他」漏计手续费`

## PR

- 一个 PR 只做一件事。
- 附上复现步骤或测试用例。
- 涉及界面的改动请附截图。
