import React from 'react';
import { useAppStore } from '../../store/appStore';
import { CheckCircle, XCircle, AlertTriangle, Info, X } from 'lucide-react';

const iconMap = {
  success: CheckCircle,
  error: XCircle,
  warning: AlertTriangle,
  info: Info,
};

const colorMap = {
  success: 'bg-success text-white',
  error: 'bg-error text-white',
  warning: 'bg-warning text-gray-900',
  info: 'bg-primary-500 text-white',
};

const Toast: React.FC = () => {
  const { toasts, removeToast } = useAppStore();

  if (toasts.length === 0) return null;

  return (
    <div className="fixed top-4 right-4 z-[9999] flex flex-col gap-2 pointer-events-none">
      {toasts.map((toast) => {
        const Icon = iconMap[toast.type];
        // 错误用 role="alert"（打断式提示）；成功/警告/信息用 role="status"，
        // 避免非关键 toast 每次都触发屏幕阅读器打断播报
        const roleAttr = toast.type === 'error' ? 'alert' : 'status';
        return (
          <div
            key={toast.id}
            className={`${colorMap[toast.type]} rounded-lg px-4 py-3 shadow-lg flex items-center gap-2 
                        min-w-[280px] max-w-[400px] animate-slide-in-right pointer-events-auto`}
            role={roleAttr}
          >
            <Icon size={18} className="shrink-0" />
            <span className="text-sm flex-1">{toast.message}</span>
            <button
              onClick={() => removeToast(toast.id)}
              className="shrink-0 opacity-70 hover:opacity-100 transition-fast cursor-pointer"
              aria-label="关闭通知"
            >
              <X size={14} />
            </button>
          </div>
        );
      })}
    </div>
  );
};

export default Toast;
