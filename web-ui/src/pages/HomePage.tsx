import React, { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAppStore } from '../store/appStore';
import { MODULE_DEFINITIONS } from '../services/modules';
import { getModuleMissing, getModuleFiles, getModuleText } from '../services/localStorage';
import ConfirmDialog from '../components/ui/ConfirmDialog';
import type { ModuleId, ModuleStatus } from '../types';


const STATUS_CONFIG: Record<ModuleStatus, { label: string; color: string; bg: string; icon: string }> = {
  completed: { label: '已完成', color: 'text-success', bg: 'bg-success-light', icon: '✅' },
  partial: { label: '部分完成', color: 'text-warning', bg: 'bg-warning-light', icon: '🟡' },
  none: { label: '未填写', color: 'text-text-disabled', bg: 'bg-gray-100', icon: '❌' },
  na: { label: '无此业务', color: 'text-orange', bg: 'bg-orange-light', icon: '⚠' },
};

function getEffectiveStatus(period: string, moduleId: ModuleId): ModuleStatus {
  const missing = getModuleMissing(period, moduleId);
  if (missing === 'not_applicable') return 'na';

  const files = getModuleFiles(period, moduleId);
  const text = getModuleText(period, moduleId);

  if (files.length > 0 || text.trim().length > 0) return 'completed';
  if (missing === 'delayed') return 'partial';
  return 'none';
}

const HomePage: React.FC = () => {
  const navigate = useNavigate();
  const { currentPeriod } = useAppStore();
  const [showConfirm, setShowConfirm] = useState(false);
  const [generating, setGenerating] = useState(false);

  const moduleStates = useMemo(() => {
    return MODULE_DEFINITIONS.map((mod) => ({
      ...mod,
      status: getEffectiveStatus(currentPeriod, mod.id),
    }));
  }, [currentPeriod]);

  const completedCount = moduleStates.filter(
    (m) => m.status === 'completed' || m.status === 'na'
  ).length;
  const totalModules = MODULE_DEFINITIONS.length;
  const progressPercent = Math.round((completedCount / totalModules) * 100);
  const allNone = moduleStates.every((m) => m.status === 'none');
  const hasIncomplete = moduleStates.some(
    (m) => m.status === 'none' || m.status === 'partial'
  );

  const handleGenerate = () => {
    if (allNone || generating) return;
    if (hasIncomplete) {
      setShowConfirm(true);
      return;
    }
    setGenerating(true);
    navigate('/result');
  };

  const handleConfirmGenerate = () => {
    setShowConfirm(false);
    navigate('/result');
  };

  const incompleteModules = moduleStates
    .filter((m) => m.status === 'none' || m.status === 'partial')
    .map((m) => m.title);

  return (
    <div className="max-w-[1200px] mx-auto px-6 py-8 flex flex-col gap-8">
      {/* Progress Section */}
      <section className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
        <h2 className="text-lg font-semibold text-text-primary mb-1">📊 账务处理进度</h2>
        <p className="text-sm text-text-secondary mb-4">
          请逐项上传本月原始资料，完成后一键生成账务文件
        </p>

        <div className="relative h-3 bg-gray-200 rounded-full overflow-hidden mb-3">
          <div
            className="absolute left-0 top-0 h-full bg-primary-500 rounded-full transition-all duration-500 ease-out"
            style={{ width: `${progressPercent}%` }}
          />
        </div>

        <div className="flex flex-wrap items-center gap-2 text-sm">
          <span className="text-text-secondary">
            {completedCount}/{totalModules} 模块已填报
          </span>
          <span className="text-text-disabled">|</span>
          {moduleStates.map((mod) => {
            const config = STATUS_CONFIG[mod.status];
            return (
              <span key={mod.id} className="flex items-center gap-1 text-xs">
                <span>{config.icon}</span>
                <span className={config.color}>{mod.subtitle}</span>
              </span>
            );
          })}
        </div>
      </section>

      {/* Module Cards Grid */}
      <section className="grid grid-cols-1 md:grid-cols-3 xl:grid-cols-5 gap-4">
        {moduleStates.map((mod) => {
              const config = STATUS_CONFIG[mod.status];
              return (
                <div
                  key={mod.id}
                  className="bg-white rounded-lg border border-gray-200 p-5 flex flex-col gap-3
                             card-hover cursor-pointer shadow-sm"
                  onClick={() => navigate(mod.route)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') navigate(mod.route);
                  }}
                >
                  <div className="flex items-center gap-3">
                    <span className="text-3xl" role="img" aria-label={mod.title}>
                      {mod.icon}
                    </span>
                    <div>
                      <h3 className="text-base font-semibold text-text-primary">{mod.title}</h3>
                    </div>
                  </div>

                  <p className="text-sm text-text-secondary">{mod.description}</p>

                  <div className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium 
                                   ${config.bg} ${config.color} w-fit`}>
                    <span>{config.icon}</span>
                    <span>{config.label}</span>
                  </div>

                  {/* 视觉按钮样式，但保持单一可交互元素（外层 div 承担点击），
                      避免 role=button 内嵌真实 button 造成嵌套交互与可访问性问题 */}
                  <span
                    className="block w-full mt-auto py-2 text-sm font-medium text-primary-500 bg-primary-50 
                               text-center rounded-lg transition-fast"
                  >
                    进入录入
                  </span>
                </div>
              );
            })}
      </section>

      {/* Generate Button */}
      <div className="flex justify-end">
        <button
          onClick={handleGenerate}
          disabled={allNone || generating}
          className={`px-8 py-3 rounded-lg font-semibold text-white text-base transition-fast
            flex items-center gap-2 cursor-pointer
            ${allNone
              ? 'bg-gray-300 cursor-not-allowed'
              : 'bg-primary-500 hover:bg-primary-600 shadow-md hover:shadow-lg'
            }`}
          title={allNone ? '请先上传至少一个模块的资料' : '点击生成账务文件'}
        >
          <span>{generating ? '⏳' : '🚀'}</span>
          <span>{generating ? '正在生成…' : '开始生成账务文件'}</span>
        </button>
      </div>

      <ConfirmDialog
        isOpen={showConfirm}
        title="部分模块尚未填写"
        message={`以下模块尚未填写：\n${incompleteModules.join('、')}\n\n是否忽略并继续？`}
        confirmText="忽略继续"
        cancelText="返回补充"
        onConfirm={handleConfirmGenerate}
        onCancel={() => setShowConfirm(false)}
      />
    </div>
  );
};

export default HomePage;
