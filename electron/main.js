/**
 * 小虾记账 - Electron 主进程
 * 职责：窗口管理、Python 后端进程托管、原生文件对话框、安全策略
 */
const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const net = require('net');
const http = require('http');
const { spawn } = require('child_process');

// 低配机 / 远程桌面 / 虚拟机常见显卡驱动不可用，Electron 会因
// "GPU process isn't usable. Goodbye." 直接崩溃退出（表现为双击无反应）。
// 本应用是纯 2D 表单界面，关掉硬件加速更稳，且启动更快。
app.disableHardwareAcceleration();

// 注意：disableHardwareAcceleration 只关掉硬件加速合成，Chromium 仍可能为其他
// 用途单独拉起 GPU 进程；该进程在无独显 / 远程桌面 / 虚拟机上会连续崩溃并 FATAL
// 退出（实测连崩 6 次后 exit 0x80000003）。把 GPU 放进主进程可稳定启动。
// 复现：不带本开关时 exe 双击即退且 stderr 出现 gpu_data_manager_impl_private.cc(423)。
app.commandLine.appendSwitch('in-process-gpu');

// ==================== 配置 ====================
const CONFIG = {
  appName: '小虾记账',
  width: 1200,
  height: 820,
  minWidth: 960,
  minHeight: 640,
  backendPort: 0,          // 运行时确定（随机可用端口，避免端口冲突）
  defaultPort: 18432,      // 首选端口
};

let mainWindow = null;
let backendProcess = null;
let isQuitting = false;

// 数据目录：我的文档\小虾记账（用户可见、可备份）
function getDataDir() {
  const dir = path.join(app.getPath('documents'), '小虾记账');
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

// ==================== 窗口 ====================
function createWindow() {
  // 打包后 __dirname = resources/app.asar/electron，图标在 asar 根 assets/
  const iconPath = path.join(__dirname, '..', 'assets', 'icon.ico');

  mainWindow = new BrowserWindow({
    width: CONFIG.width,
    height: CONFIG.height,
    minWidth: CONFIG.minWidth,
    minHeight: CONFIG.minHeight,
    title: CONFIG.appName,
    icon: fs.existsSync(iconPath) ? iconPath : undefined,
    backgroundColor: '#ffffff',
    show: false,
    autoHideMenuBar: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: true,
      spellcheck: false,
    },
  });

  // 先加载秒开的本地启动页，后端就绪后再切到真正的前端首页。
  // 原实现是「await startBackend() 再建窗口」，首次启动后端要 30 多秒，
  // 这段时间屏幕上什么都没有，用户会以为双击没反应。
  const loadingPath = path.join(__dirname, 'loading.html');
  if (fs.existsSync(loadingPath)) {
    mainWindow.loadFile(loadingPath);
  }

  mainWindow.once('ready-to-show', () => mainWindow.show());
  // 兜底：极低配/软件渲染异常时首帧可能迟迟不到，超时也要把窗口显出来，
  // 否则用户看到的是"双击没反应"（进程在跑但窗口始终 show:false）。
  setTimeout(() => {
    if (mainWindow && !mainWindow.isVisible()) mainWindow.show();
  }, 8000);
  mainWindow.on('closed', () => { mainWindow = null; });

  // 外部链接交给系统浏览器，禁止应用内打开新窗口
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\//i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  // 禁止跳转到外部域名（只允许本地页面）
  mainWindow.webContents.on('will-navigate', (event, url) => {
    const isLocal = url.startsWith('file://') || url.startsWith(`http://127.0.0.1:${CONFIG.backendPort}`);
    if (!isLocal) event.preventDefault();
  });
}

// 后端就绪后，把启动页换成真正的前端首页
function loadFrontend() {
  if (!mainWindow) return; // 启动过程中用户已经关了窗口
  // 打包后 __dirname = resources/app.asar/electron，前端产物在 asar 根 dist/
  const frontendPath = path.join(__dirname, '..', 'dist', 'index.html');
  if (fs.existsSync(frontendPath)) {
    mainWindow.loadFile(frontendPath);
  } else {
    mainWindow.loadURL('http://localhost:5173');
  }
}

function buildMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        {
          label: '打开数据目录',
          click: () => shell.openPath(getDataDir()),
        },
        {
          label: '打开输出报表目录',
          click: () => {
            const out = path.join(getDataDir(), '输出');
            fs.mkdirSync(out, { recursive: true });
            shell.openPath(out);
          },
        },
        { type: 'separator' },
        { label: '退出', role: 'quit' },
      ],
    },
    {
      label: '编辑',
      submenu: [
        { label: '撤销', role: 'undo' },
        { label: '重做', role: 'redo' },
        { type: 'separator' },
        { label: '剪切', role: 'cut' },
        { label: '复制', role: 'copy' },
        { label: '粘贴', role: 'paste' },
        { label: '全选', role: 'selectAll' },
      ],
    },
    {
      label: '视图',
      submenu: [
        { label: '重新加载', role: 'reload' },
        { type: 'separator' },
        { label: '实际大小 (100%)', role: 'resetZoom' },
        { label: '放大', role: 'zoomIn' },
        { label: '缩小', role: 'zoomOut' },
        { type: 'separator' },
        { label: '全屏', role: 'togglefullscreen' },
        {
          label: '窗口置顶',
          type: 'checkbox',
          checked: false,
          click: (item) => {
            if (mainWindow) mainWindow.setAlwaysOnTop(item.checked);
          },
        },
        { type: 'separator' },
        { label: '开发者工具', role: 'toggleDevTools' },
      ],
    },
    {
      label: '帮助',
      submenu: [
        {
          label: `关于 ${CONFIG.appName}`,
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: `关于 ${CONFIG.appName}`,
              message: `${CONFIG.appName} v${app.getVersion()}`,
              detail:
                '小微企业智能代账桌面工具\n\n' +
                '· 全部账务解析与报表运算在本机完成\n' +
                '· 不上传任何财务数据到云端\n' +
                `· 数据目录：${getDataDir()}\n\n` +
                '示例科技有限公司',
              buttons: ['确定'],
            });
          },
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ==================== 后端进程 ====================

function findAvailablePort(preferred) {
  const tryListen = (port, fallbackRandom) =>
    new Promise((resolve) => {
      const srv = net.createServer();
      srv.once('error', () => {
        if (fallbackRandom) return resolve(0);
        tryListen(0, true).then(resolve);
      });
      srv.listen(port, '127.0.0.1', () => {
        const actual = srv.address().port;
        srv.close(() => resolve(actual));
      });
    });
  return tryListen(preferred, false);
}

function waitForBackend(port, maxRetries = 60) {
  return new Promise((resolve, reject) => {
    let tries = 0;
    const check = () => {
      const req = http.request(
        { hostname: '127.0.0.1', port, path: '/api/health', timeout: 1500 },
        (res) => {
          res.resume();
          if (res.statusCode === 200) return resolve(true);
          retry();
        }
      );
      req.on('error', retry);
      req.setTimeout(1500, () => req.destroy());
      req.end();
    };
    const retry = () => {
      tries += 1;
      if (tries > maxRetries) return reject(new Error('后端服务启动超时'));
      setTimeout(check, 500);
    };
    check();
  });
}

async function startBackend() {
  const isDev = !app.isPackaged;
  const dataDir = getDataDir();
  const port = await findAvailablePort(CONFIG.defaultPort);
  CONFIG.backendPort = port;

  let cmd;
  let args;
  let cwd;

  if (isDev) {
    cmd = process.env.VENV_PYTHON || 'python';
    args = [
      '-m', 'uvicorn', 'server:app',
      '--host', '127.0.0.1',
      '--port', String(port),
      '--app-dir', path.join(__dirname, '..', 'backend'),
    ];
    cwd = path.join(__dirname, '..');
  } else {
    cmd = path.join(process.resourcesPath, 'backend', 'server.exe');
    args = [];
    cwd = path.dirname(cmd);
  }

  console.log(`[backend] spawn ${cmd} port=${port} data=${dataDir}`);

  backendProcess = spawn(cmd, args, {
    cwd,
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
    env: {
      ...process.env,
      XJZ_PORT: String(port),
      XJZ_HOST: '127.0.0.1',
      XJZ_DATA_DIR: dataDir,
      PYTHONIOENCODING: 'utf-8',
      PYTHONUTF8: '1',
    },
  });

  backendProcess.stdout.on('data', (d) => process.stdout.write(`[backend] ${d}`));
  backendProcess.stderr.on('data', (d) => process.stderr.write(`[backend] ${d}`));
  backendProcess.on('close', (code) => {
    console.log(`[backend] exited code=${code}`);
    backendProcess = null;
    // 后端异常退出：非主动退出时提示用户
    if (!isQuitting && mainWindow) {
      dialog.showErrorBox(
        '服务异常',
        `本地账务引擎已退出（代码 ${code}）。请重新启动小虾记账。`
      );
    }
  });

  await waitForBackend(port);
  console.log('[backend] ready');
}

function stopBackend() {
  if (!backendProcess) return;
  try {
    if (process.platform === 'win32') {
      spawn('taskkill', ['/pid', String(backendProcess.pid), '/f', '/t'], { windowsHide: true });
    } else {
      backendProcess.kill('SIGTERM');
    }
  } catch (e) {
    console.error('[backend] stop failed', e);
  }
  backendProcess = null;
}

// ==================== IPC ====================

ipcMain.on('get-backend-port-sync', (e) => {
  e.returnValue = CONFIG.backendPort || 0;
});

ipcMain.handle('get-backend-port', () => CONFIG.backendPort);

ipcMain.handle('dialog:open-file', async (_e, options) => {
  return dialog.showOpenDialog(mainWindow, options || {});
});

ipcMain.handle('dialog:save-file', async (_e, options) => {
  return dialog.showSaveDialog(mainWindow, options || {});
});

ipcMain.handle('dialog:select-dir', async () => {
  return dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: '选择包含银行回单/明细的文件夹',
  });
});

ipcMain.handle('get-app-info', () => ({
  name: CONFIG.appName,
  version: app.getVersion(),
  dataDir: getDataDir(),
  isPackaged: app.isPackaged,
}));

ipcMain.handle('open-data-dir', () => shell.openPath(getDataDir()));
ipcMain.handle('open-external', (_e, url) => shell.openExternal(url));

// ==================== 生命周期 ====================

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });

  app.whenReady().then(async () => {
    buildMenu();
    // 先把窗口开出来（加载秒开的启动页），再并行启动后端；
    // 后端就绪后切到前端首页。这样双击后几秒内就有可见界面。
    createWindow();
    try {
      await startBackend();
      console.log('[startup] backend ready, loading frontend');
      loadFrontend();
    } catch (err) {
      dialog.showErrorBox(
        '启动失败',
        `无法启动本地账务引擎：\n${err.message}\n\n` +
          '请确认安装目录未被安全软件拦截，或尝试重新安装。'
      );
      app.quit();
    }
  });
}

app.on('window-all-closed', () => {
  stopBackend();
  app.quit();
});

app.on('before-quit', () => {
  isQuitting = true;
  stopBackend();
});
