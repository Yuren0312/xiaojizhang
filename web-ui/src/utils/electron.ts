/**
 * 小虾记账 - 桌面端适配工具
 * 仅在 Electron 环境生效；Web 环境自动降级
 */

export const isElectron = typeof window !== 'undefined' && 'electronAPI' in window;

/** 应用信息（含数据目录） */
export async function getAppInfo() {
  if (isElectron && window.electronAPI) {
    return window.electronAPI.getAppInfo();
  }
  return { name: '小虾记账', version: '1.0.0 (Web)', dataDir: '', isPackaged: false };
}

/** 打开账套数据目录（桌面版） */
export async function openDataDir(): Promise<boolean> {
  if (isElectron && window.electronAPI) {
    await window.electronAPI.openDataDir();
    return true;
  }
  return false;
}

/** 选择文件夹（导入银行 PDF 目录） */
export async function selectDirectory(): Promise<string | null> {
  if (isElectron && window.electronAPI) {
    const result = await window.electronAPI.selectDir();
    if (!result.canceled && result.filePaths.length > 0) return result.filePaths[0];
  }
  return null;
}

/** 打开外部链接 */
export async function openExternal(url: string): Promise<void> {
  if (isElectron && window.electronAPI) {
    await window.electronAPI.openExternal(url);
    return;
  }
  window.open(url, '_blank', 'noopener');
}
