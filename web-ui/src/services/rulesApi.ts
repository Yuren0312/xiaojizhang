/**
 * 规则配置 API 客户端
 */

const API_BASE = import.meta.env.VITE_API_BASE || '';

export interface BankRule {
  match: string;
  account: string;
  note: string;
}

export interface MappingRules {
  rules: BankRule[];
  cash_account: string;
  source_file: string;
}

export interface AccountDef {
  type: string;
  bs?: number;
  bs_sub?: number;
  is?: number;
  is_sub?: number;
  cash?: boolean;
  cf_in?: number;
  cf_out?: number;
}

export interface AccountsData {
  aliases: Record<string, string>;
  accounts: Record<string, AccountDef>;
}

export interface CompanyConfig {
  entity_type: 'corp' | 'individual';
  company_name: string;
  tax_rate: number;
}

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const url = `${API_BASE}${path}`;
  const res = await fetch(url, {
    ...options,
    headers: { ...(options.headers || {}) },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`API ${res.status}: ${body || res.statusText}`);
  }
  return res.json();
}

/** 获取映射规则 */
export async function getMappingRules(): Promise<MappingRules> {
  return request('/api/rules/mapping');
}

/** 更新映射规则（整表替换） */
export async function updateMappingRules(
  data: { bank_rules: { match: string; account: string; note?: string }[]; cash_account?: string },
): Promise<{ success: boolean; rules_count: number }> {
  return request('/api/rules/mapping', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
}

/** 获取科目表 */
export async function getAccounts(): Promise<AccountsData> {
  return request('/api/rules/accounts');
}

/** 获取企业配置 */
export async function getConfig(): Promise<CompanyConfig> {
  return request('/api/config');
}

/** 保存企业配置 */
export async function saveConfig(config: Partial<CompanyConfig>): Promise<{ success: boolean }> {
  return request('/api/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(config),
  });
}
