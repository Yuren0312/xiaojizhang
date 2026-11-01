/**
 * 小虾记账 API 客户端
 * 封装对后端 FastAPI 服务的 HTTP 调用
 *
 * 运行环境适配：
 * - 开发模式：走 Vite proxy（/api -> localhost:8100）
 * - 桌面版（Electron）：从主进程同步读取本地引擎端口
 * - 生产 Web 版：通过环境变量 VITE_API_BASE 配置
 */

const isElectron = typeof window !== 'undefined' && 'electronAPI' in window;

// 同步解析：Electron 创建窗口前主进程已确定端口，避免首屏请求的竞态
export const API_BASE = (() => {
  if (import.meta.env.VITE_API_BASE) return import.meta.env.VITE_API_BASE as string;
  if (isElectron && window.electronAPI?.getBackendPortSync) {
    try {
      const port = window.electronAPI.getBackendPortSync();
      if (port) return `http://127.0.0.1:${port}`;
    } catch {
      /* 忽略，回落到同源 */
    }
  }
  return '';
})();

// ---- Types ----

export interface OutputFile {
  label: string;
  filename: string;
  size: number;
  download_url: string;
}

export interface ComputeResponse {
  success: boolean;
  month_label: string;
  outputs: OutputFile[];
  errors: string[];
  warnings: string[];
  stats: {
    voucher_count: number;
    bank_txn_count: number;
    journal_entry_count: number;
  };
}

export interface PeriodInfo {
  period: string;
  has_data: boolean;
  output_count: number;
}

export interface ResultFileItem {
  name: string;
  size: number;
  ext: string;
}

export interface ArchiveResponse {
  success: boolean;
  message: string;
  period: string;
  archive_path: string;
}

export interface ArchivedPeriod {
  period: string;
  output_count: number;
  archived_at: string;
}

// ---- API ----

async function request<T>(
  path: string,
  options: RequestInit = {},
): Promise<T> {
  const url = `${API_BASE}${path}`;
  const res = await fetch(url, {
    signal: AbortSignal.timeout(60000),
    ...options,
    headers: {
      ...(options.headers || {}),
    },
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`API ${res.status}: ${body || res.statusText}`);
  }

  return res.json();
}

/** 健康检查 */
export async function checkHealth(): Promise<{ status: string; company: string }> {
  return request('/api/health');
}

/** 列出账期 */
export async function listPeriods(): Promise<PeriodInfo[]> {
  return request('/api/periods');
}

/** 上传文件到指定模块 */
export async function deleteInputFile(
  period: string,
  module: string,
  filename: string,
): Promise<{ success: boolean }> {
  const url = `${API_BASE}/api/input/${encodeURIComponent(period)}/${encodeURIComponent(module)}/${encodeURIComponent(filename)}`;
  const res = await fetch(url, { method: 'DELETE', signal: AbortSignal.timeout(30000) });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data?.detail || `删除失败 (${res.status})`);
  }
  return res.json();
}

export async function uploadFile(
  period: string,
  module: string,
  file: File,
  onProgress?: (pct: number) => void,
): Promise<{ success: boolean; filename: string; path: string; size: number }> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    const formData = new FormData();
    formData.append('period', period);
    formData.append('module', module);
    formData.append('file', file);

    xhr.upload.addEventListener('progress', (e) => {
      if (e.lengthComputable && onProgress) {
        onProgress(Math.round((e.loaded / e.total) * 100));
      }
    });

    xhr.addEventListener('load', () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          resolve(JSON.parse(xhr.responseText));
        } catch {
          reject(new Error('解析响应失败'));
        }
      } else {
        reject(new Error(`上传失败 (${xhr.status})`));
      }
    });

    xhr.addEventListener('error', () => reject(new Error('网络错误')));
    xhr.addEventListener('abort', () => reject(new Error('上传已取消')));

    xhr.open('POST', `${API_BASE}/api/upload`);
    xhr.send(formData);
  });
}

/** 上传文本说明 */
export async function uploadText(
  period: string,
  module: string,
  text: string,
): Promise<{ success: boolean; path: string }> {
  const formData = new FormData();
  formData.append('period', period);
  formData.append('module', module);
  formData.append('text', text);

  const res = await fetch(`${API_BASE}/api/upload-text`, {
    method: 'POST',
    body: formData,
  });

  if (!res.ok) throw new Error(`上传文本失败 (${res.status})`);
  return res.json();
}

/** 触发账务运算 */
export async function compute(
  year: number,
  month: number,
  entityType: 'corp' | 'individual' = 'corp',
): Promise<ComputeResponse> {
  const res = await fetch(`${API_BASE}/api/compute`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ year, month, entity_type: entityType }),
    signal: AbortSignal.timeout(180000),  // 运算为最长调用，180s 超时
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`运算失败: ${body || res.statusText}`);
  }

  return res.json();
}

export interface ComputeTaskStatus {
  task_id: string;
  status: 'pending' | 'running' | 'done' | 'done_with_errors' | 'error';
  progress: number;
  message: string;
  result?: ComputeResponse;
  error?: string;
}

/** 创建异步计算任务（后端后台线程执行，返回 task_id 供轮询真实进度） */
export async function startComputeTask(
  year: number,
  month: number,
  entityType: 'corp' | 'individual' = 'corp',
  signal?: AbortSignal,
): Promise<{ task_id: string }> {
  const res = await fetch(`${API_BASE}/api/compute/task`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ year, month, entity_type: entityType }),
    signal: signal || AbortSignal.timeout(30000),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`创建计算任务失败: ${body || res.statusText}`);
  }
  return res.json();
}

/** 轮询任务进度（0-100 真实进度 + 阶段文案） */
export async function getComputeTask(taskId: string, signal?: AbortSignal): Promise<ComputeTaskStatus> {
  const res = await fetch(`${API_BASE}/api/compute/task/${encodeURIComponent(taskId)}`, {
    signal: signal || AbortSignal.timeout(30000),
  });
  if (!res.ok) {
    throw new Error(`查询任务失败 (${res.status})`);
  }
  return res.json();
}

/** 列出结果文件 */
export async function listResults(period: string): Promise<{ period: string; files: ResultFileItem[] }> {
  return request(`/api/results/${period}`);
}

/** 获取下载 URL */
export function getDownloadUrl(downloadUrl: string): string {
  return `${API_BASE}${downloadUrl}`;
}

/** 流式下载结果文件 */
export async function downloadResult(downloadUrl: string): Promise<{ blob: Blob; filename: string }> {
  const fullUrl = `${API_BASE}${downloadUrl}`;
  const res = await fetch(fullUrl);

  if (!res.ok) throw new Error(`下载失败 (${res.status})`);

  const disposition = res.headers.get('content-disposition') || '';
  const match = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
  let filename = 'download';
  if (match?.[1]) {
    filename = match[1].replace(/['"]/g, '');
  } else {
    filename = downloadUrl.split('/').pop() || 'download';
  }

  const blob = await res.blob();
  return { blob, filename };
}

/** 归档指定账期 */
export async function archivePeriod(period: string): Promise<ArchiveResponse> {
  const res = await fetch(`${API_BASE}/api/archive`, {
    signal: AbortSignal.timeout(60000),
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ period }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`归档失败: ${body || res.statusText}`);
  }

  return res.json();
}

/** 列出已归档账期 */
export async function listArchived(): Promise<{ archived: ArchivedPeriod[] }> {
  return request('/api/archive');
}
