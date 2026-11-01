/**
 * 小虾记账 - Preload
 * 通过 contextBridge 暴露受限 API，渲染进程无法访问 Node
 */
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // 后端端口：同步读取，保证前端模块加载时即可拿到
  getBackendPortSync: () => ipcRenderer.sendSync('get-backend-port-sync'),
  getBackendPort: () => ipcRenderer.invoke('get-backend-port'),

  // 原生对话框
  openFile: (options) => ipcRenderer.invoke('dialog:open-file', options),
  saveFile: (options) => ipcRenderer.invoke('dialog:save-file', options),
  selectDir: () => ipcRenderer.invoke('dialog:select-dir'),

  // 应用信息 / 目录
  getAppInfo: () => ipcRenderer.invoke('get-app-info'),
  openDataDir: () => ipcRenderer.invoke('open-data-dir'),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),

  isElectron: true,
  platform: process.platform,
});
