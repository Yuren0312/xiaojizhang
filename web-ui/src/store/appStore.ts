import { create } from 'zustand';
import type { AccountingPeriod, ToastMessage, ModuleStatus, StoredFile, ModuleId, MissingFlag } from '../types';
import {
  getPreferences,
  setPreferences,
  getModuleStatus as lsGetModuleStatus,
  getModuleText as lsGetModuleText,
  getModuleMissing as lsGetModuleMissing,
  getModuleFiles as lsGetModuleFiles,
  getAvailablePeriods,
} from '../services/localStorage';

function getCurrentDefaultPeriod(): string {
  const prefs = getPreferences();
  return prefs.lastPeriod;
}

interface AppState {
  currentPeriod: AccountingPeriod;
  availablePeriods: AccountingPeriod[];
  toasts: ToastMessage[];
  showSettings: boolean;

  setCurrentPeriod: (period: AccountingPeriod) => void;
  refreshAvailablePeriods: () => void;
  addToast: (toast: Omit<ToastMessage, 'id'>) => void;
  removeToast: (id: string) => void;
  setShowSettings: (show: boolean) => void;

  getModuleStatus: (period: AccountingPeriod, module: ModuleId) => ModuleStatus;
  getModuleText: (period: AccountingPeriod, module: ModuleId) => string;
  getModuleMissing: (period: AccountingPeriod, module: ModuleId) => MissingFlag;
  getModuleFiles: (period: AccountingPeriod, module: ModuleId) => StoredFile[];
}

let toastCounter = 0;

export const useAppStore = create<AppState>((set, get) => ({
  currentPeriod: getCurrentDefaultPeriod(),
  availablePeriods: getAvailablePeriods(),
  toasts: [],
  showSettings: false,

  setCurrentPeriod: (period) => {
    const prefs = getPreferences();
    prefs.lastPeriod = period;
    setPreferences(prefs);
    set({ currentPeriod: period, availablePeriods: getAvailablePeriods() });
  },

  refreshAvailablePeriods: () => {
    set({ availablePeriods: getAvailablePeriods() });
  },

  addToast: (toast) => {
    const id = `toast-${++toastCounter}`;
    set((state) => ({
      toasts: [...state.toasts, { ...toast, id }],
    }));
    const duration = toast.duration ?? 3000;
    setTimeout(() => {
      get().removeToast(id);
    }, duration);
  },

  removeToast: (id) => {
    set((state) => ({
      toasts: state.toasts.filter((t) => t.id !== id),
    }));
  },

  setShowSettings: (show) => {
    set({ showSettings: show });
  },

  getModuleStatus: (period, module) => lsGetModuleStatus(period, module),
  getModuleText: (period, module) => lsGetModuleText(period, module),
  getModuleMissing: (period, module) => lsGetModuleMissing(period, module),
  getModuleFiles: (period, module) => lsGetModuleFiles(period, module),
}));
