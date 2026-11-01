import type { AccountingPeriod, ModuleId, ModuleStatus, MissingFlag, StoredFile, ResultFile, UserPreferences } from '../types';

const LS_PREFIX = 'lbs:';

function lsKey(...parts: string[]): string {
  return LS_PREFIX + parts.join(':');
}

export function getModuleStatusKey(period: AccountingPeriod, module: ModuleId): string {
  return lsKey(period, module, 'status');
}

export function getModuleTextKey(period: AccountingPeriod, module: ModuleId): string {
  return lsKey(period, module, 'text');
}

export function getModuleMissingKey(period: AccountingPeriod, module: ModuleId): string {
  return lsKey(period, module, 'missing');
}

export function getModuleFilesKey(period: AccountingPeriod, module: ModuleId): string {
  return lsKey(period, module, 'files');
}

export function getResultFilesKey(period: AccountingPeriod): string {
  return lsKey(period, 'result', 'files');
}

export function getResultGeneratedAtKey(period: AccountingPeriod): string {
  return lsKey(period, 'result', 'generatedAt');
}

export function getPreferencesKey(): string {
  return lsKey('preferences');
}

export function getItem<T>(key: string, defaultValue: T): T {
  try {
    const raw = localStorage.getItem(key);
    if (raw === null) return defaultValue;
    return JSON.parse(raw) as T;
  } catch {
    return defaultValue;
  }
}

export function setItem<T>(key: string, value: T): void {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.error('localStorage write error:', e);
  }
}

export function removeItem(key: string): void {
  try {
    localStorage.removeItem(key);
  } catch (e) {
    console.error('localStorage remove error:', e);
  }
}

export function getModuleStatus(period: AccountingPeriod, module: ModuleId): ModuleStatus {
  return getItem<ModuleStatus>(getModuleStatusKey(period, module), 'none');
}

export function setModuleStatus(period: AccountingPeriod, module: ModuleId, status: ModuleStatus): void {
  setItem(getModuleStatusKey(period, module), status);
}

export function getModuleText(period: AccountingPeriod, module: ModuleId): string {
  return getItem<string>(getModuleTextKey(period, module), '');
}

export function setModuleText(period: AccountingPeriod, module: ModuleId, text: string): void {
  setItem(getModuleTextKey(period, module), text);
}

export function getModuleMissing(period: AccountingPeriod, module: ModuleId): MissingFlag {
  return getItem<MissingFlag>(getModuleMissingKey(period, module), null);
}

export function setModuleMissing(period: AccountingPeriod, module: ModuleId, flag: MissingFlag): void {
  setItem(getModuleMissingKey(period, module), flag);
}

export function getModuleFiles(period: AccountingPeriod, module: ModuleId): StoredFile[] {
  return getItem<StoredFile[]>(getModuleFilesKey(period, module), []);
}

export function setModuleFiles(period: AccountingPeriod, module: ModuleId, files: StoredFile[]): void {
  setItem(getModuleFilesKey(period, module), files);
}

export function getResultFiles(period: AccountingPeriod): ResultFile[] {
  return getItem<ResultFile[]>(getResultFilesKey(period), []);
}

export function setResultFiles(period: AccountingPeriod, files: ResultFile[]): void {
  setItem(getResultFilesKey(period), files);
}

export function getResultGeneratedAt(period: AccountingPeriod): string | null {
  return getItem<string | null>(getResultGeneratedAtKey(period), null);
}

export function setResultGeneratedAt(period: AccountingPeriod, date: string): void {
  setItem(getResultGeneratedAtKey(period), date);
}

export function getPreferences(): UserPreferences {
  const now = new Date();
  const defaultPeriod = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  return getItem<UserPreferences>(getPreferencesKey(), {
    lastPeriod: defaultPeriod,
    language: 'zh-CN',
  });
}

export function setPreferences(prefs: UserPreferences): void {
  setItem(getPreferencesKey(), prefs);
}

export function getAvailablePeriods(): AccountingPeriod[] {
  const periods: Set<string> = new Set();
  const prefix = LS_PREFIX;
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith(prefix)) {
      const parts = key.replace(prefix, '').split(':');
      if (parts.length >= 2) {
        const period = parts[0];
        if (/^\d{4}-\d{2}$/.test(period)) {
          periods.add(period);
        }
      }
    }
  }
  const current = getPreferences().lastPeriod;
  periods.add(current);
  return Array.from(periods).sort();
}

/** 清理指定账期的所有前端缓存键 */
export function clearPeriodCache(period: AccountingPeriod): number {
  let count = 0;
  const prefix = lsKey(period);
  const keysToRemove: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith(prefix)) {
      keysToRemove.push(key);
    }
  }
  for (const key of keysToRemove) {
    removeItem(key);
    count++;
  }
  return count;
}

/** 输入资料变更后失效结果缓存（localStorage 结果键 + IndexedDB 结果 blob）。
 *
 * 上传/删除/保存输入资料后调用，避免进入结果页时静默展示基于旧输入的缓存结果。 */
export async function invalidateResultCache(period: AccountingPeriod): Promise<void> {
  // 1) localStorage：lbs:{period}:result:*（files / generatedAt）
  const prefix = lsKey(period, 'result');
  const keysToRemove: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith(prefix)) {
      keysToRemove.push(key);
    }
  }
  keysToRemove.forEach(removeItem);

  // 2) IndexedDB：结果 blob（key = /api/results/{period}/...）
  try {
    const { getAllFileKeys, deleteFile } = await import('./indexedDB');
    const keys = await getAllFileKeys('results');
    const idbPrefix = `/api/results/${period}/`;
    for (const key of keys) {
      if (key.startsWith(idbPrefix)) {
        await deleteFile(key, 'results');
      }
    }
  } catch {
    /* 清理失败不阻断 */
  }
}
