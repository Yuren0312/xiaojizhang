/**
 * Electron API 类型声明（preload 暴露）
 */
interface ElectronAPI {
  getBackendPortSync: () => number;
  getBackendPort: () => Promise<number>;

  openFile: (options?: Electron.OpenDialogOptions) => Promise<Electron.OpenDialogReturnValue>;
  saveFile: (options?: Electron.SaveDialogOptions) => Promise<Electron.SaveDialogReturnValue>;
  selectDir: () => Promise<Electron.OpenDialogReturnValue>;

  getAppInfo: () => Promise<{
    name: string;
    version: string;
    dataDir: string;
    isPackaged: boolean;
  }>;

  openDataDir: () => Promise<string>;
  openExternal: (url: string) => Promise<void>;

  isElectron: true;
  platform: NodeJS.Platform;
}

declare global {
  interface Window {
    electronAPI?: ElectronAPI;
  }
}

export {};
