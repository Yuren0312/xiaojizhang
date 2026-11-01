import type { ModuleDefinition } from '../types';

export const MODULE_DEFINITIONS: ModuleDefinition[] = [
  {
    id: 'revenue',
    icon: '💰',
    title: '收入信息录入',
    subtitle: '收入信息',
    description: '已开具发票、无票收入',
    route: '/input/revenue',
    uploadHint: '上传发票PDF/图片/Excel',
    uploadAccept: '.pdf,.jpg,.jpeg,.png,.xlsx,.xls',
    textLabel: '无票收入信息',
    textPlaceholder: '请描述本月未开票的收入业务，例如：\n2026-07-15 零售现金收入 5000元\n2026-07-20 服务费收入（未开票） 3000元\n\n建议格式：日期 + 业务摘要 + 金额（每行一条）',
  },
  {
    id: 'expenses',
    icon: '🧾',
    title: '成本费用发票上传',
    subtitle: '成本费用',
    description: '进项发票、各类费用发票',
    route: '/input/expenses',
    uploadHint: '上传进项发票/费用发票（PDF/图片/Excel）',
    uploadAccept: '.pdf,.jpg,.jpeg,.png,.xlsx,.xls',
    textLabel: '无票费用信息',
    textPlaceholder: '请描述未取得发票的费用，例如：\n2026-07-10 办公用品采购 2000元\n2026-07-25 快递物流费 500元\n\n建议格式：日期 + 业务摘要 + 金额（每行一条）',
  },
  {
    id: 'bank',
    icon: '🏦',
    title: '银行单据上传',
    subtitle: '银行单据',
    description: '银行对账单、银行回单',
    route: '/input/bank',
    uploadHint: '上传银行对账单/回单（PDF）',
    uploadAccept: '.pdf,.jpg,.jpeg,.png,.xlsx,.xls',
    textLabel: '现金交易补充说明',
    textPlaceholder: '大额现金收付说明，例如：\n2026-07-08 大额现金收款 50000元（来自XX客户）\n2026-07-12 现金支付XX货款 30000元\n\n建议格式：日期 + 业务摘要 + 金额（每行一条）',
  },
  {
    id: 'payroll',
    icon: '👥',
    title: '工资&社保信息录入',
    subtitle: '工资社保',
    description: '工资数据、社保变动信息',
    route: '/input/payroll',
    uploadHint: '上传工资表/社保变动表（Excel/图片）',
    uploadAccept: '.pdf,.jpg,.jpeg,.png,.xlsx,.xls',
    textLabel: '工资数据手动录入（如未上传文件）',
    textPlaceholder: '请填写员工工资数据，例如：\n张三 基本工资10000 社保2000 实发8000\n李四 基本工资8000 社保1600 实发6400\n\n建议格式：姓名 + 基本工资 + 社保 + 实发（每行一人）',
  },
  {
    id: 'company-info',
    icon: '⚙',
    title: '企业基础信息变更登记',
    subtitle: '企业信息',
    description: '银行账户地址、联系人、预留手机号变更',
    route: '/input/company-info',
    uploadHint: '上传变更证明文件（PDF/图片）',
    uploadAccept: '.pdf,.jpg,.jpeg,.png',
    textLabel: '企业信息变更说明',
    textPlaceholder: '描述变更内容，例如：\n变更日期：2026-07-15\n变更事项：银行账户地址变更\n变更后信息：XX市XX区XX路XX号\n\n建议格式：变更日期 + 变更事项 + 变更后信息',
  },
];

export function getModuleById(id: string): ModuleDefinition | undefined {
  return MODULE_DEFINITIONS.find((m) => m.id === id);
}
