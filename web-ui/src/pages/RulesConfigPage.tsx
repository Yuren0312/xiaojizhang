import React, { useState, useEffect, useCallback } from 'react';
import { useNavigate } from 'react-router-dom';
import { ArrowLeft, Plus, Trash2, Save, RotateCcw } from 'lucide-react';
import { getMappingRules, updateMappingRules, type BankRule } from '../services/rulesApi';
import { useAppStore } from '../store/appStore';
import ConfirmDialog from '../components/ui/ConfirmDialog';

interface RuleWithId extends BankRule {
  _id: string;
}

const RulesConfigPage: React.FC = () => {
  const navigate = useNavigate();
  const { addToast } = useAppStore();

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [rules, setRules] = useState<RuleWithId[]>([]);
  const [cashAccount, setCashAccount] = useState('');
  const [sourceFile, setSourceFile] = useState('');
  const [isDirty, setIsDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [originalRules, setOriginalRules] = useState<RuleWithId[]>([]);
  const [originalCashAccount, setOriginalCashAccount] = useState('');

  const loadRules = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const data = await getMappingRules();
      const loaded = data.rules.map((r, i) => ({
        match: r.match,
        account: r.account,
        note: r.note || '',
        _id: `rule-${i}`,
      }));
      setRules(loaded);
      setCashAccount(data.cash_account || '');
      setSourceFile(data.source_file || '');
      setOriginalRules(JSON.parse(JSON.stringify(loaded)));
      setOriginalCashAccount(data.cash_account || '');
      setIsDirty(false);
    } catch (err: any) {
      setError(err?.message || '加载规则失败');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadRules();
  }, [loadRules]);

  const markDirty = () => setIsDirty(true);

  const handleMatchChange = (index: number, value: string) => {
    setRules((prev) => prev.map((r, i) => (i === index ? { ...r, match: value } : r)));
    markDirty();
  };

  const handleAccountChange = (index: number, value: string) => {
    setRules((prev) => prev.map((r, i) => (i === index ? { ...r, account: value } : r)));
    markDirty();
  };

  const handleNoteChange = (index: number, value: string) => {
    setRules((prev) => prev.map((r, i) => (i === index ? { ...r, note: value } : r)));
    markDirty();
  };

  const handleAddRule = () => {
    setRules((prev) => [...prev, { match: '', account: '', note: '', _id: `rule-new-${Date.now()}` }]);
    markDirty();
  };

  const handleDeleteRule = (index: number) => {
    setRules((prev) => prev.filter((_, i) => i !== index));
    markDirty();
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      const payload = {
        bank_rules: rules.map((r) => ({
          match: r.match.trim(),
          account: r.account.trim(),
          note: r.note.trim(),
        })),
        cash_account: cashAccount.trim() || '货币资金',
      };

      // 校验
      const emptyMatch = payload.bank_rules.findIndex((r) => !r.match);
      if (emptyMatch >= 0) {
        addToast({ type: 'error', message: `第 ${emptyMatch + 1} 条规则的匹配关键词不能为空` });
        setSaving(false);
        return;
      }
      const emptyAccount = payload.bank_rules.findIndex((r) => !r.account);
      if (emptyAccount >= 0) {
        addToast({ type: 'error', message: `第 ${emptyAccount + 1} 条规则的对方科目不能为空` });
        setSaving(false);
        return;
      }

      await updateMappingRules(payload);
      // 快照用 trim 后的 payload，与落盘 YAML 一致（避免撤销恢复出带空格版本）
      setOriginalRules(payload.bank_rules.map((r, i) => ({ ...r, _id: rules[i]?._id || `rule-${i}` })));
      setOriginalCashAccount(payload.cash_account);
      setIsDirty(false);
      addToast({ type: 'success', message: `✅ 已保存 ${rules.length} 条映射规则` });
    } catch (err: any) {
      addToast({ type: 'error', message: `保存失败: ${err?.message}` });
    } finally {
      setSaving(false);
    }
  };

  const handleReset = () => {
    setRules(JSON.parse(JSON.stringify(originalRules)));
    setCashAccount(originalCashAccount);
    setIsDirty(false);
    setShowResetConfirm(false);
    addToast({ type: 'info', message: '已放弃修改，返回首页' });
    navigate('/');  // 修复：确认撤销后离开页面，避免用户被困
  };

  return (
    <div className="max-w-[1200px] mx-auto px-6 py-8 flex flex-col gap-6">
      <div className="flex items-center gap-4">
        <button
          onClick={() => {
            if (isDirty) {
              setShowResetConfirm(true);
            } else {
              navigate('/');
            }
          }}
          className="flex items-center gap-1 text-sm text-text-secondary hover:text-primary-500 
                     transition-fast cursor-pointer"
        >
          <ArrowLeft size={16} />
          <span>返回首页</span>
        </button>
        <h1 className="text-xl font-semibold text-text-primary">⚙️ 银行流水映射规则</h1>
      </div>

      {error && (
        <div className="bg-error-light border border-error rounded-lg p-4 text-sm text-error">
          {error}
          <button onClick={loadRules} className="ml-3 underline cursor-pointer">重试</button>
        </div>
      )}

      <div className="bg-warning-light border border-warning rounded-lg px-4 py-3 text-sm text-text-primary">
        <strong>⚠ 规则说明：</strong> 银行流水按「摘要 + 对方户名」中的关键词自上而下匹配，
        第一个匹配到的规则决定对方科目。匹配不上的流水程序会报错并列出待归类清单。
        规则文件：<code className="bg-gray-200 px-1 rounded">{sourceFile}</code>
      </div>

      {loading ? (
        <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
          <div className="animate-pulse space-y-4">
            {[1, 2, 3, 4, 5].map((i) => (
              <div key={i} className="flex gap-3">
                <div className="h-10 bg-gray-200 rounded flex-1" />
                <div className="h-10 bg-gray-200 rounded w-32" />
                <div className="h-10 bg-gray-200 rounded w-40" />
                <div className="h-10 bg-gray-200 rounded w-10" />
              </div>
            ))}
          </div>
        </div>
      ) : (
        <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
          {/* Header */}
          <div className="px-6 py-3 border-b border-gray-200 bg-gray-50 rounded-t-lg grid grid-cols-12 gap-3 text-xs font-medium text-text-secondary">
            <div className="col-span-1">#</div>
            <div className="col-span-3">匹配关键词 *</div>
            <div className="col-span-2">对方科目 *</div>
            <div className="col-span-4">备注说明</div>
            <div className="col-span-2">操作</div>
          </div>

          {/* Rule Rows */}
          <div className="divide-y divide-gray-100 max-h-[60vh] overflow-y-auto">
            {rules.map((rule, index) => (
              <div
                key={rule._id || index}
                className="px-6 py-3 grid grid-cols-12 gap-3 items-center hover:bg-gray-50 transition-fast"
              >
                <span className="col-span-1 text-xs text-text-disabled">{index + 1}</span>
                <div className="col-span-3">
                  <input
                    type="text"
                    value={rule.match}
                    onChange={(e) => handleMatchChange(index, e.target.value)}
                    placeholder="如：微信支付认证"
                    className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded 
                               focus:outline-none focus:ring-2 focus:ring-primary-300 focus:border-primary-400"
                  />
                </div>
                <div className="col-span-2">
                  <input
                    type="text"
                    value={rule.account}
                    onChange={(e) => handleAccountChange(index, e.target.value)}
                    placeholder="如：财务费用-利息"
                    className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded 
                               focus:outline-none focus:ring-2 focus:ring-primary-300 focus:border-primary-400"
                  />
                </div>
                <div className="col-span-4">
                  <input
                    type="text"
                    value={rule.note}
                    onChange={(e) => handleNoteChange(index, e.target.value)}
                    placeholder="对应什么业务场景..."
                    className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded 
                               focus:outline-none focus:ring-2 focus:ring-primary-300 focus:border-primary-400"
                  />
                </div>
                <div className="col-span-2">
                  <button
                    onClick={() => handleDeleteRule(index)}
                    className="inline-flex items-center gap-1 px-2 py-1.5 text-xs text-error 
                               hover:bg-error-light rounded transition-fast cursor-pointer"
                  >
                    <Trash2 size={12} />
                    <span>删除</span>
                  </button>
                </div>
              </div>
            ))}
          </div>

          {/* Add Rule */}
          <div className="px-6 py-3 border-t border-gray-200">
            <button
              onClick={handleAddRule}
              className="inline-flex items-center gap-1 text-sm text-primary-500 hover:text-primary-600 
                         font-medium transition-fast cursor-pointer"
            >
              <Plus size={16} />
              <span>新增规则</span>
            </button>
          </div>
        </div>
      )}

      {/* Cash Account */}
      <div className="bg-white rounded-lg border border-gray-200 p-4 shadow-sm">
        <label className="text-sm font-medium text-text-primary">
          现金及现金等价物科目
          <span className="text-text-disabled font-normal ml-2">
            （对账单的借贷方向据此判断流入/流出）
          </span>
        </label>
        <input
          type="text"
          value={cashAccount}
          onChange={(e) => { setCashAccount(e.target.value); markDirty(); }}
          placeholder="货币资金"
          className="mt-2 w-full max-w-sm px-3 py-2 text-sm border border-gray-300 rounded 
                     focus:outline-none focus:ring-2 focus:ring-primary-300 focus:border-primary-400"
        />
      </div>

      {/* Buttons */}
      <div className="flex justify-end gap-3">
        <button
          onClick={() => navigate('/')}
          className="px-4 py-2 text-sm text-text-secondary bg-white border border-gray-300 
                     hover:bg-gray-50 rounded-lg transition-fast cursor-pointer"
        >
          返回首页
        </button>
        {isDirty && (
          <button
            onClick={() => setShowResetConfirm(true)}
            className="px-4 py-2 text-sm text-text-secondary bg-white border border-gray-300 
                       hover:bg-gray-50 rounded-lg transition-fast cursor-pointer flex items-center gap-2"
          >
            <RotateCcw size={14} />
            <span>撤销修改</span>
          </button>
        )}
        <button
          onClick={handleSave}
          disabled={saving || !isDirty}
          className={`px-5 py-2 text-sm font-medium text-white rounded-lg transition-fast 
            cursor-pointer flex items-center gap-2
            ${saving || !isDirty
              ? 'bg-gray-300 cursor-not-allowed'
              : 'bg-primary-500 hover:bg-primary-600 shadow-sm'
            }`}
        >
          {saving ? (
            <>
              <div className="w-4 h-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
              <span>保存中...</span>
            </>
          ) : (
            <>
              <Save size={14} />
              <span>保存规则</span>
            </>
          )}
        </button>
      </div>

      <ConfirmDialog
        isOpen={showResetConfirm}
        title="确认撤销修改"
        message="将放弃当前所有未保存的修改，恢复为上次保存的版本。确定继续？"
        confirmText="确认撤销"
        cancelText="继续编辑"
        onConfirm={handleReset}
        onCancel={() => setShowResetConfirm(false)}
        variant="danger"
      />
    </div>
  );
};

export default RulesConfigPage;
