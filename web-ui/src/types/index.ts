// Module status types
export type ModuleStatus = 'none' | 'partial' | 'completed' | 'na';

// Missing flag types
export type MissingFlag = 'delayed' | 'not_applicable' | null;

// Module IDs
export type ModuleId = 'revenue' | 'expenses' | 'bank' | 'payroll' | 'company-info';

// Module definition
export interface ModuleDefinition {
  id: ModuleId;
  icon: string;
  title: string;
  subtitle: string;
  description: string;
  route: string;
  uploadHint: string;
  uploadAccept: string;
  textLabel: string;
  textPlaceholder: string;
}

// Stored file metadata
export interface StoredFile {
  id: string;
  name: string;
  size: number;
  type: string;
  indexedDBKey: string;
  uploadedAt: string;
  /** 后端实际保存的文件名（重名时可能被后端改为 发票(1).pdf 等），
   * 用于删除后端文件时准确定位；缺失时回退到 name。 */
  backendName?: string;
}

// Result file metadata
export interface ResultFile {
  id: string;
  name: string;
  size: number;
  type: string;
  label: string;
  description: string;
  indexedDBKey: string;
}

// Extension file (future)
export interface ExtensionFile {
  id: string;
  name: string;
  description: string;
  status: 'development' | 'planned' | 'ready';
  eta: string;
}

// Computation step
export interface ComputationStep {
  id: string;
  label: string;
  status: 'waiting' | 'running' | 'completed' | 'failed';
}

// Toast notification
export interface ToastMessage {
  id: string;
  type: 'success' | 'error' | 'warning' | 'info';
  message: string;
  duration?: number;
}

// User preferences
export interface UserPreferences {
  lastPeriod: string;
  language: string;
  entityType?: 'corp' | 'individual';
}

// Accounting period (YYYY-MM format)
export type AccountingPeriod = string;
