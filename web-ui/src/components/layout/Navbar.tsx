import React, { useEffect, useState } from 'react';
import { Link, useLocation } from 'react-router-dom';
import { useAppStore } from '../../store/appStore';
import { Settings, FolderArchive } from 'lucide-react';
import { listArchived, listPeriods } from '../../services/api';

interface NavbarProps {
  title?: string;
}

const NEW_PERIOD_OPTION = '__new_period__';

const Navbar: React.FC<NavbarProps> = ({ title = '小虾记账' }) => {
  const { currentPeriod, availablePeriods, setCurrentPeriod, mergeAvailablePeriods, addToast, setShowSettings } = useAppStore();
  const [archivedCount, setArchivedCount] = useState(0);
  const location = useLocation();

  useEffect(() => {
    listArchived()
      .then((data) => setArchivedCount(data.archived.length))
      .catch(() => {});
  }, [location.pathname]);

  // 合并后端 /api/periods 发现的账期（含仅有输出/输入目录但本地无缓存键的账期）
  useEffect(() => {
    listPeriods()
      .then((data) => mergeAvailablePeriods(data.map((p) => p.period)))
      .catch(() => {});
  }, [mergeAvailablePeriods]);

  const handlePeriodChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const value = e.target.value;
    if (value === NEW_PERIOD_OPTION) {
      const input = window.prompt('输入新账期（格式 YYYY-MM，例如 2026-09）：');
      if (input && /^\d{4}-\d{2}$/.test(input.trim())) {
        const period = input.trim();
        // 校验合理性（年份/月份），不合法时提示
        const [, , m] = period.match(/^(\d{4})-(\d{2})$/) || [];
        const month = parseInt(m, 10);
        if (month >= 1 && month <= 12) {
          setCurrentPeriod(period);
          mergeAvailablePeriods([period]);
        } else {
          addToast({ type: 'warning', message: `账期月份无效：${period}` });
        }
      } else if (input) {
        addToast({ type: 'warning', message: '账期格式应为 YYYY-MM，例如 2026-09' });
      }
      // 恢复 select 显示为当前账期
      e.target.value = currentPeriod;
      return;
    }
    setCurrentPeriod(value);
  };

  return (
    <nav className="bg-primary-700 text-white shadow-md sticky top-0 z-50">
      <div className="max-w-[1200px] mx-auto px-6 h-14 flex items-center justify-between">
        {/* Logo & Title */}
        <Link to="/" className="flex items-center gap-2 hover:opacity-90 transition-fast">
          <span className="text-2xl" role="img" aria-label="lobster">🦞</span>
          <span className="text-lg font-semibold tracking-wide">{title}</span>
        </Link>

        {/* Right side */}
        <div className="flex items-center gap-4">
          {/* Archived count */}
          {archivedCount > 0 && (
            <div className="flex items-center gap-1 text-sm text-primary-100">
              <FolderArchive size={16} />
              <span>{archivedCount} 期已归档</span>
            </div>
          )}

          {/* Period Selector（全页面可见：结果页可直接切换账期查看不同月份） */}
          <div className="flex items-center gap-2">
            <label htmlFor="period-select" className="text-sm text-primary-100 hidden sm:inline">
              当前账期:
            </label>
            <select
              id="period-select"
              value={currentPeriod}
              onChange={handlePeriodChange}
              className="bg-primary-600 text-white text-sm px-3 py-1.5 rounded-md border border-primary-500 
                         focus:outline-none focus:ring-2 focus:ring-primary-300 cursor-pointer
                         hover:bg-primary-500 transition-fast appearance-none"
              aria-label="选择账期"
            >
              {availablePeriods.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
              <option value={NEW_PERIOD_OPTION}>➕ 新建账期...</option>
            </select>
          </div>

          {/* Settings */}
          <button
            onClick={() => setShowSettings(true)}
            className="p-2 rounded-md hover:bg-primary-600 transition-fast cursor-pointer"
            aria-label="设置"
            title="设置"
          >
            <Settings size={20} />
          </button>
        </div>
      </div>
    </nav>
  );
};

export default Navbar;
