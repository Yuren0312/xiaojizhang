import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAppStore } from '../../store/appStore';
import Modal from '../ui/Modal';
import { getConfig, saveConfig, type CompanyConfig } from '../../services/rulesApi';
import { getPreferences, setPreferences } from '../../services/localStorage';

const SettingsModal: React.FC = () => {
  const navigate = useNavigate();
  const { showSettings, setShowSettings, addToast } = useAppStore();

  const [config, setConfig] = useState<CompanyConfig>({
    entity_type: 'corp',
    company_name: '',
    tax_rate: 5.0,
  });
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [isDirty, setIsDirty] = useState(false);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    if (showSettings && !loaded) {
      setLoading(true);
      getConfig()
        .then((cfg) => {
          setConfig(cfg);
          setLoaded(true);
        })
        .catch(() => {
          addToast({ type: 'warning', message: '加载企业配置失败，使用默认值' });
          setLoaded(true);
        })
        .finally(() => setLoading(false));
    }
  }, [showSettings, loaded, addToast]);

  const handleChange = (field: keyof CompanyConfig, value: string | number) => {
    setConfig((prev) => ({ ...prev, [field]: value }));
    setIsDirty(true);
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      await saveConfig(config);
      // 同步 entity_type 到 localStorage，供 ResultPage 读取
      const prefs = getPreferences();
      prefs.entityType = config.entity_type;
      setPreferences(prefs);
      setIsDirty(false);
      addToast({ type: 'success', message: '✅ 企业配置已保存' });
      setShowSettings(false);
    } catch (err: any) {
      addToast({ type: 'error', message: `保存失败: ${err?.message}` });
    } finally {
      setSaving(false);
    }
  };

  return (
    <Modal
      isOpen={showSettings}
      onClose={() => {
        if (isDirty) {
          // 保留脏数据但不保存的提示可以简化
        }
        setShowSettings(false);
      }}
      title="⚙️ 设置"
      size="md"
    >
      <div className="space-y-6">
        {loading ? (
          <div className="animate-pulse space-y-4 py-4">
            {[1, 2, 3].map((i) => (
              <div key={i} className="h-10 bg-gray-200 rounded" />
            ))}
          </div>
        ) : (
          <>
            {/* 纳税人类型 */}
            <div>
              <label className="block text-sm font-medium text-text-primary mb-1.5">
                纳税人类型
              </label>
              <div className="flex gap-3">
                <label
                  className={`flex-1 px-3 py-2 text-sm rounded-lg border cursor-pointer transition-fast text-center
                    ${config.entity_type === 'corp'
                      ? 'bg-primary-50 border-primary-400 text-primary-600 font-medium'
                      : 'border-gray-300 text-text-secondary hover:bg-gray-50'
                    }`}
                >
                  <input
                    type="radio"
                    className="sr-only"
                    name="entity_type"
                    value="corp"
                    checked={config.entity_type === 'corp'}
                    onChange={() => handleChange('entity_type', 'corp')}
                  />
                  🏢 有限公司
                </label>
                <label
                  className={`flex-1 px-3 py-2 text-sm rounded-lg border cursor-pointer transition-fast text-center
                    ${config.entity_type === 'individual'
                      ? 'bg-primary-50 border-primary-400 text-primary-600 font-medium'
                      : 'border-gray-300 text-text-secondary hover:bg-gray-50'
                    }`}
                >
                  <input
                    type="radio"
                    className="sr-only"
                    name="entity_type"
                    value="individual"
                    checked={config.entity_type === 'individual'}
                    onChange={() => handleChange('entity_type', 'individual')}
                  />
                  👤 个体户 / 个独
                </label>
              </div>
              <p className="text-xs text-text-disabled mt-1">
                {config.entity_type === 'corp' ? '适用企业所得税（5%优惠税率）' : '适用经营所得个人所得税（五级累进）'}
              </p>
            </div>

            {/* 企业名称 */}
            <div>
              <label className="block text-sm font-medium text-text-primary mb-1.5">
                企业名称
              </label>
              <input
                type="text"
                value={config.company_name}
                onChange={(e) => handleChange('company_name', e.target.value)}
                placeholder="请输入企业全称"
                className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg 
                           focus:outline-none focus:ring-2 focus:ring-primary-300 focus:border-primary-400"
              />
            </div>

            {/* 税率 */}
            {config.entity_type === 'corp' && (
              <div>
                <label className="block text-sm font-medium text-text-primary mb-1.5">
                  企业所得税优惠税率 (%)
                </label>
                <input
                  type="number"
                  value={config.tax_rate}
                  onChange={(e) => handleChange('tax_rate', parseFloat(e.target.value) || 0)}
                  step="0.5"
                  min="0"
                  max="25"
                  className="w-full max-w-[120px] px-3 py-2 text-sm border border-gray-300 rounded-lg 
                             focus:outline-none focus:ring-2 focus:ring-primary-300 focus:border-primary-400"
                />
                <p className="text-xs text-text-disabled mt-1">
                  小微企业标准税率 5%（年应纳税所得额 ≤ 300万），一般企业 25%
                </p>
              </div>
            )}

            {/* 规则管理入口 */}
            <div className="border-t border-gray-200 pt-4">
              <h4 className="text-sm font-semibold text-text-primary mb-2">🛠 高级配置</h4>
              <button
                onClick={() => {
                  setShowSettings(false);
                  navigate('/rules-config');
                }}
                className="w-full py-2.5 text-sm font-medium text-primary-500 bg-primary-50 
                           hover:bg-primary-100 rounded-lg transition-fast cursor-pointer"
              >
                📋 管理银行流水映射规则
              </button>
              <p className="text-xs text-text-disabled mt-1 text-center">
                配置银行摘要关键词 → 会计科目的对应关系
              </p>
            </div>

            {/* 关于 */}
            <div className="border-t border-gray-200 pt-4">
              <h4 className="text-sm font-semibold text-text-primary mb-2">关于小虾记账</h4>
              <p className="text-sm text-text-secondary">
                小虾记账 v1.0 — 专为小微企业打造的智能代账工具。
                账套数据安全存储，支持本地与云端部署。
              </p>
              <div className="mt-2 text-xs text-text-disabled space-y-0.5">
                <p>准则：小企业会计准则</p>
                <p>技术：Python 引擎 + React 前端</p>
              </div>
            </div>
          </>
        )}

        {/* Buttons */}
        <div className="flex gap-3 pt-2">
          <button
            onClick={() => setShowSettings(false)}
            className="flex-1 py-2 text-sm font-medium text-text-secondary bg-gray-100 
                       hover:bg-gray-200 rounded-lg transition-fast cursor-pointer"
          >
            取消
          </button>
          <button
            onClick={handleSave}
            disabled={saving || !isDirty}
            className={`flex-1 py-2 text-sm font-medium text-white rounded-lg transition-fast cursor-pointer
              ${saving || !isDirty
                ? 'bg-gray-300 cursor-not-allowed'
                : 'bg-primary-500 hover:bg-primary-600'
              }`}
          >
            {saving ? '保存中...' : '保存配置'}
          </button>
        </div>
      </div>
    </Modal>
  );
};

export default SettingsModal;
