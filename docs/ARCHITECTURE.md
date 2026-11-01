# 架构说明

三层：**Electron 外壳 → React 前端 → Python 账务引擎**。三层都跑在同一台机器上，没有服务端。

```
┌─────────────────────────────────────────────────────────────┐
│ Electron 主进程 (electron/main.js)                          │
│  · 建窗口、加载启动页 → 前端首页                             │
│  · 随手启动 Python 后端子进程（随机端口）                     │
│  · 原生文件对话框、外部链接交给系统浏览器                     │
│  · 退出时 taskkill 后端进程树                                │
└───────────────┬─────────────────────────────┬───────────────┘
                │ IPC (preload.js, contextIsolation)
┌───────────────▼──────────────┐   ┌──────────▼───────────────┐
│ 渲染进程：React + Vite        │──►│ FastAPI (backend/server.py)│
│ web-ui/                       │HTTP│ 127.0.0.1:<随机端口>       │
└──────────────────────────────┘   └──────────┬───────────────┘
                                              │ import
                                   ┌──────────▼───────────────┐
                                   │ accounting_report/        │
                                   │ 解析 → 分录 → 报表        │
                                   └───────────────────────────┘
```

## 进程与启动顺序

1. Electron `whenReady` **先建窗口**，加载 `electron/loading.html`（本地启动页，秒开）。
2. 并行 `startBackend()`：挑一个可用端口（首选 18432，被占用则随机），`spawn` 后端进程。
3. 轮询 `GET /api/health`，200 后 `loadFrontend()` 切到前端首页。

这个顺序是为了避开「首次启动后端要 30 多秒、期间屏幕上什么都没有」的体验问题。**不要改回「先 await 后端再建窗口」**。

后端可执行文件在包里是 `resources/backend/server.exe`（PyInstaller 产物，放在 asar 外面，因为 asar 内的 exe 无法直接执行）。

## 端口与环境变量

| 变量 | 作用 |
| --- | --- |
| `XJZ_PORT` | 后端监听端口（Electron 传入） |
| `XJZ_HOST` | 监听地址，默认 `127.0.0.1` |
| `XJZ_DATA_DIR` | 数据根目录（账套所在），桌面版指向 `文档/小虾记账` |
| `VENV_PYTHON` | 开发模式下指定用哪个 Python 启动 uvicorn |

## 数据流

```
输入/年-月/
  ├─ 银行对账单/*.pdf   ──parse──┐
  ├─ 收入发票/*.pdf      ──parse──┤
  ├─ 工资社保/*.xlsx     ──parse──┼─→ 会计分录（借贷平衡）
  └─ 业务事项.xlsx       ──parse──┘        │
                                           ├─→ writers/ 各报表 → 输出/*.xlsx
上月 输出/*.xlsx ──读期初──┘
```

账务引擎不依赖 Electron，也不依赖 FastAPI，`python -m accounting_report` 可以独立跑，这是刻意的：桌面端只是它的一个外壳。

## 目录约定

| 路径 | 是否进版本库 | 说明 |
| --- | --- | --- |
| `electron/` `backend/` `accounting_report/` `web-ui/` | ✅ | 源码 |
| `dist/` `web-ui/dist/` `release/` | ❌ | 构建产物，可重建 |
| `backend-server/` | ❌ | PyInstaller 产物，可重建 |
| `输入/` `输出/` | ❌ | **真实账套，绝不能进仓库** |
| `backend/config.json` | ❌ | 运行期配置，用 `config.example.json` 代替 |

## 安全边界

- 后端只监听 `127.0.0.1`，不接受外部连接。曾有把 `host` 改成 `0.0.0.0` 的写法，那等于把本地账务 API 暴露到局域网（配合回环地址白名单缺失可被 RCE），已改回。
- 渲染进程开 `contextIsolation`、关 `nodeIntegration`，只通过 preload 暴露白名单 IPC。
- 应用内导航被限制为本地 `file://` 与 `http://127.0.0.1:<port>`，外链一律交给系统浏览器。
