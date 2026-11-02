# 小虾记账 · XiaoJiZhang

给小微企业用的**月度会计报表自动编制**桌面工具。把上月报表和本月的银行对账单、回单、业务事项丢进去，自动生成本月的五张会计报表。

Electron 桌面端 + React 前端 + Python 账务引擎，**全部在本机运行，不上传任何账务数据**。

> 状态：v1.1.0，Windows 10/11 x64。目前只支持《小企业会计准则》主体。

---

## 它解决什么问题

小微企业每月要出五张表：资产负债表、利润表、现金流量表，加上季度版的利润表和现金流量表。手工做法是拿上月的表，逐格推算这个月的变化——勾稽关系靠人脑维护，改一处忘一处，表就不平了，而且不平的时候不知道错在哪。

这个工具在中间加了一层**会计分录**：

```
银行对账单 ─┐
业务回单   ─┼─→ 会计分录（借贷平衡） ─┬─→ 资产负债表
业务事项表 ─┘                        ├─→ 利润表
                                    └─→ 现金流量表
```

所有输入先转成借贷分录，三张表都从同一批分录派生。资产负债表平衡因此是**结构性成立**的，不是靠人工对齐。设计思路详见 [docs/DESIGN.md](docs/DESIGN.md)。

## 主要能力

| 能力 | 说明 |
| --- | --- |
| 银行对账单解析 | 工行等电子回单 PDF，提取日期、金额、摘要、对方户名 |
| 业务单据解析 | 增值税发票 PDF（文本层），自动生成发票台账 |
| 分录生成 | 按规则把每笔业务映射成借贷分录，科目表可改 |
| 五张报表 | 资产负债表 / 利润表 / 现金流量表 / 两张季报，输出 xlsx，版式对齐财务软件 |
| 税务底稿 | 增值税申报底稿、企业所得税 A 类底稿、工资薪金个税 |
| 税务调整项 | 纳税调增/调减、弥补亏损、减免所得税额、已预缴税额，录入后代入所得税底稿 |
| 电子税务局报送 | 会小企 01/02 表，以税局模板为底稿填数，下载即可上传，无需誊抄 |
| 其他附表 | 银行余额调节、往来明细、核算汇总、分析简报 |
| 桌面封装 | 本地后端随包分发，双击即用，无需装 Python |

## 快速开始

### 方式一：用现成的安装包

到 Releases 下载 `小虾记账-<版本>-x64.exe`（安装版）或 `便携版.exe`。

首次启动会在 `文档/小虾记账/` 建好数据目录（`归档` `输入` `输出` `accounting_report`），把单据放进对应账期的 `输入` 目录即可。

### 方式二：从源码跑

需要 **Python 3.11 / 3.12**（3.13 上 PyMuPDF 的轮子目前有兼容问题）、**Node 18+**。
Windows 还需安装 [Microsoft Visual C++ 2015-2022 Redistributable (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe)，否则 PyMuPDF 会报 `DLL load failed while importing _extra`。

```bash
git clone https://github.com/Yuren0312/xiaojizhang.git
cd xiaojizhang

# 1. 后端依赖
python -m venv .venv
.venv\Scripts\activate           # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt

# 2. 前端依赖并构建
npm --prefix web-ui install
npm --prefix web-ui run build    # 产出 web-ui/dist → 由 prepare 脚本复制到 dist/

# 3. 组装并启动桌面端
node scripts/prepare-build.js
set VENV_PYTHON=.venv\Scripts\python.exe   # macOS/Linux: export VENV_PYTHON=.../python
npm start
```

只想调界面：`npm run dev`（Vite 开发服务器，默认 5173），Electron 会加载 `http://localhost:5173`。

### 方式三：只用账务引擎（命令行）

```bash
# 生成一套合成示例账套（不含任何真实数据）
python tools/generate_sample_data.py

# 出指定月份的报表
python -m accounting_report --month 2026-09
```

## 目录结构

```
electron/            Electron 主进程（窗口、后端进程托管、IPC）
backend/             FastAPI 服务（REST API，供前端调用）
accounting_report/   账务引擎（解析 → 分录 → 报表），纯 Python，可单独用
  parsers/           PDF / xlsx 解析器
  rules/             科目映射、账务规则（YAML，用户可改）
  writers/           各张报表的生成器
  templates/         xlsx 报表模板
web-ui/              React + Vite + Tailwind 前端
scripts/             构建期脚本
tools/               示例数据生成、模板构建、图标生成等辅助工具
tests/               回归测试（pytest）
docs/                设计文档、架构、打包说明
```

## 数据与隐私

- 解析、记账、出表**全部在本机完成**，运行期不联网。
- 账套默认落在 `我的文档/小虾记账/`，卸载软件不会删除该目录。
- 仓库里**不含任何真实账套、银行回单、工资表**；示例数据由 `tools/generate_sample_data.py` 合成。

## 已知限制

- PDF 解析依赖**文本层**，扫描件（图片型 PDF）需要先做 OCR，当前未内置。
- 只实现了《小企业会计准则》一套口径；一般纳税人与小规模纳税人的差异按配置区分。
- 报表模板为固定版式，跨软件导入时可能需要微调列宽。
- 生成的税务底稿是**辅助计算**，正式申报前请人工复核。

## 文档

| 文档 | 内容 |
| --- | --- |
| [docs/DESIGN.md](docs/DESIGN.md) | 账务引擎的设计思路（为什么以分录为枢纽） |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | 三层架构、进程模型、端口与数据流 |
| [docs/USAGE.md](docs/USAGE.md) | 面向使用者的操作说明 |
| [docs/PACKAGING.md](docs/PACKAGING.md) | 怎么打出 Windows 安装包 |
| [CONTRIBUTING.md](CONTRIBUTING.md) | 开发约定、测试、提交规范 |

## 许可证

[MIT](LICENSE)。使用时请自行判断是否适用于你的业务场景；本项目不构成税务或审计意见。
