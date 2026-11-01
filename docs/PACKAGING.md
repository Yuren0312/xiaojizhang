# 打包说明（Windows）

产物：NSIS 安装包 + 便携版单文件 exe。构建在 Windows 上完成，macOS/Linux 交叉打 Windows 包未验证。

## 前置

- Node 18+、Python 3.10+
- 首次构建会从网络拉 Electron 与 NSIS 资源；国内网络建议走镜像（见下）

## 一条命令

```bash
bash build.sh          # 仓库根目录；等价于下面三步
```

等价手工流程：

```bash
# 1. 前端构建（产出 web-ui/dist）
npm run build:frontend

# 2. 组装：检查前端产物 + 后端 server.exe 是否齐备，并把前端复制到 dist/
npm run prepare:app

# 3. 出包（安装版 + 便携版）
npx electron-builder --win nsis portable
```

后端 `server.exe` 由 PyInstaller 生成：

```bash
python build_backend.py     # 产物在 backend-server/server/server.exe
```

## 需要留意的四件事

### 1. 镜像

GitHub Release 直连常常超时，构建前设好镜像：

```bash
export ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/
export ELECTRON_BUILDER_BINARIES_MIRROR=https://npmmirror.com/mirrors/electron-builder-binaries/
```

### 2. `NODE_OPTIONS` 可能让构建诡异地失败

某些安全软件或宿主环境会往 `NODE_OPTIONS` 注入脚本，导致 npm 删除失败时把依赖文件改名成 `xxx.DELETE.<hash>`，构建报 `Cannot find module`。构建前清空它：

```bash
export NODE_OPTIONS=
```

`build.sh` 里已经清过。

### 3. GPU 崩溃会让双击没反应

低配机、远程桌面、虚拟机上 Chromium 的 GPU 进程可能连续崩溃并 FATAL 退出，表现是「双击了但什么都没发生」。`electron/main.js` 里已经：

```js
app.disableHardwareAcceleration();
app.commandLine.appendSwitch('in-process-gpu');   // 光第一行不够
```

### 4. 测启动时留意两个假象

- `ELECTRON_RUN_AS_NODE=1` 会让 Electron 退化成纯 Node：无窗口、退出码 0。自动化测试前要 `env.pop("ELECTRON_RUN_AS_NODE")`。
- 新生成的 exe 第一次运行会被杀软全盘扫描，冷启动可能 6~19 秒才有窗口，这不是卡死。自动化等待请给 25 秒以上。

## 版本号与文件名

- 版本号只在 `package.json` 的 `version` 维护（electron-builder 25 不接受配置文件里的根级 `version`）。
- 产物命名模板在 `electron-builder.yml`：安装版 `${productName}-${version}-${arch}.exe`，便携版后缀改为 `便携版`。

## 卸载与数据

`deleteAppDataOnUninstall: false`：卸载只删程序文件，`文档/小虾记账` 原样保留。这一点有实测记录，改动时请重新验证。
