import React, { useState, useRef, useEffect, useCallback } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { getModuleById } from '../services/modules';
import {
  getModuleFiles,
  setModuleFiles,
  getModuleText,
  setModuleText,
  getModuleMissing,
  setModuleMissing,
  setModuleStatus,
  invalidateResultCache,
} from '../services/localStorage';
import { useAppStore } from '../store/appStore';
import { storeFile, deleteFile, generateFileKey } from '../services/indexedDB';
import { uploadFile as uploadToBackend, uploadText as uploadTextToBackend, deleteInputFile } from '../services/api';
import ConfirmDialog from '../components/ui/ConfirmDialog';
import { Eye, Trash2, ArrowLeft, Upload } from 'lucide-react';
import FilePreview from '../components/ui/FilePreview';
import type { StoredFile, MissingFlag, ModuleId } from '../types';

const ALLOWED_EXTENSIONS = ['.pdf', '.jpg', '.jpeg', '.png', '.xlsx', '.xls'];
const MAX_FILE_SIZE = 20 * 1024 * 1024;

const CategoryInputPage: React.FC = () => {
  const { module: moduleId } = useParams<{ module: string }>();
  const navigate = useNavigate();
  const { currentPeriod, addToast } = useAppStore();

  const moduleDef = getModuleById(moduleId || '');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [files, setFiles] = useState<StoredFile[]>([]);
  const [textInput, setTextInput] = useState('');
  const [missingFlag, setMissingFlag] = useState<MissingFlag>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<Record<string, number>>({});
  const [saving, setSaving] = useState(false);
  const [showLeaveConfirm, setShowLeaveConfirm] = useState(false);
  const [pendingNavigation, setPendingNavigation] = useState<string | null>(null);
  const [previewFile, setPreviewFile] = useState<{ file: StoredFile; blob: Blob } | null>(null);
  const [isDirty, setIsDirty] = useState(false);
  const [fileToDelete, setFileToDelete] = useState<StoredFile | null>(null);

  useEffect(() => {
    if (!moduleId) return;
    setFiles(getModuleFiles(currentPeriod, moduleId as ModuleId));
    setTextInput(getModuleText(currentPeriod, moduleId as ModuleId));
    setMissingFlag(getModuleMissing(currentPeriod, moduleId as ModuleId));
  }, [currentPeriod, moduleId]);

  useEffect(() => {
    const handler = (e: BeforeUnloadEvent) => {
      if (isDirty) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [isDirty]);

  const markDirty = () => {
    if (!isDirty) setIsDirty(true);
  };

  const handleBack = () => {
    if (isDirty) {
      setPendingNavigation('/');
      setShowLeaveConfirm(true);
    } else {
      navigate('/');
    }
  };

  const handleConfirmLeave = () => {
    setShowLeaveConfirm(false);
    setIsDirty(false);
    if (pendingNavigation) {
      navigate(pendingNavigation);
    }
  };

  const isValidFile = (file: File): string | null => {
    const ext = '.' + file.name.split('.').pop()?.toLowerCase();
    if (!ALLOWED_EXTENSIONS.includes(ext)) {
      return `不支持"${ext}"格式，支持：${ALLOWED_EXTENSIONS.join(', ')}`;
    }
    if (file.size > MAX_FILE_SIZE) {
      return '文件过大（最大20MB），请压缩后上传';
    }
    return null;
  };

  const handleUploadFiles = useCallback(
    async (fileList: FileList | File[]) => {
      const fileArray = Array.from(fileList);
      const validFiles: File[] = [];
      const errors: string[] = [];

      for (const file of fileArray) {
        const error = isValidFile(file);
        if (error) {
          errors.push(`${file.name}: ${error}`);
        } else {
          validFiles.push(file);
        }
      }

      if (errors.length > 0) {
        errors.forEach((e) => addToast({ type: 'error', message: e }));
      }

      if (validFiles.length === 0) return;


      // 并行上传（每文件独立：后端 + 本地 IndexedDB）
      const tasks = validFiles.map(async (file) => {
        const key = generateFileKey(currentPeriod, moduleId || 'unknown', file.name);
        setUploadProgress((prev) => ({ ...prev, [key]: 0 }));
        try {
          const result = await uploadToBackend(currentPeriod, moduleId || 'unknown', file, (pct) => {
            setUploadProgress((prev) => ({ ...prev, [key]: pct / 100 }));
          });
          await storeFile(key, file, 'files');
          setUploadProgress((prev) => ({ ...prev, [key]: 1 }));
          // 用后端返回的实际文件名（重名会被后端改名如 发票(1).pdf），
          // 元数据展示与删除键都用它，避免改名文件残留不可删
          const backendName = result.filename || file.name;
          return {
            id: key,
            name: backendName,
            size: file.size,
            type: file.type,
            indexedDBKey: key,
            backendName,
            uploadedAt: new Date().toISOString(),
          } as StoredFile;
        } catch {
          addToast({ type: 'error', message: `上传失败: ${file.name}` });
          return null;
        } finally {
          setUploadProgress((prev) => {
            const next = { ...prev };
            delete next[key];
            return next;
          });
        }
      });
      const results = await Promise.all(tasks);
      const newFiles: StoredFile[] = results.filter((x): x is StoredFile => x !== null);

      if (newFiles.length > 0) {
        setFiles((prev) => {
          const next = [...prev, ...newFiles];
          // 立即落盘 localStorage，避免导航（Logo/账期下拉/返回）绕过保存导致 UI 与后端不一致
          setModuleFiles(currentPeriod, moduleId as ModuleId, next);
          return next;
        });
        markDirty();
        // 输入资料变更 → 失效本账期结果缓存，避免进入结果页展示旧结果
        invalidateResultCache(currentPeriod);
        addToast({
          type: 'success',
          message: `成功上传 ${newFiles.length} 个文件`,
        });
      }
    },
    [currentPeriod, moduleId, addToast]
  );

  const handleDeleteFile = async (file: StoredFile) => {
    try {
      // 先删后端文件（引擎读取的原始资料），再删本地 IndexedDB 缓存
      // period/module 从 indexedDBKey 提取；文件名用后端实际保存名（backendName）
      // 而不是拼接 key 的最后一段，重名改名场景才能准确删除
      const parts = String(file.indexedDBKey || '').split('/');
      if (parts.length >= 3) {
        const period = parts[0];
        const module = parts[1];
        const fname = file.backendName || file.name;
        try {
          await deleteInputFile(period, module, fname);
        } catch { /* 后端删除失败不阻断本地删除 */ }
      }
      await deleteFile(file.indexedDBKey, 'files');
      setFiles((prev) => {
        const next = prev.filter((f) => f.id !== file.id);
        // 立即落盘 localStorage，避免导航绕过保存导致 UI 与后端不一致
        setModuleFiles(currentPeriod, moduleId as ModuleId, next);
        return next;
      });
      markDirty();
      // 输入资料变更 → 失效结果缓存
      invalidateResultCache(currentPeriod);
      addToast({ type: 'success', message: `已删除 ${file.name}` });
    } catch {
      addToast({ type: 'error', message: '删除文件失败' });
    }
  };

  const handlePreviewFile = async (file: StoredFile) => {
    try {
      const { getFile } = await import('../services/indexedDB');
      const fileData = await getFile(file.indexedDBKey, 'files');
      if (fileData) {
        setPreviewFile({ file, blob: fileData.data });
      } else {
        addToast({ type: 'warning', message: `本地缓存中没有 ${file.name}，无法预览` });
      }
    } catch {
      addToast({ type: 'error', message: '预览文件失败' });
    }
  };

  const handleDropZoneClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      handleUploadFiles(e.target.files);
      e.target.value = '';
    }
  };

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.currentTarget === e.target) {
      setIsDragging(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    if (e.dataTransfer.files.length > 0) {
      handleUploadFiles(e.dataTransfer.files);
    }
  };

  const handleMissingChange = (flag: MissingFlag) => {
    setMissingFlag(flag);
    markDirty();
    // 立即落盘，避免导航绕过保存导致状态丢失
    if (moduleId) {
      setModuleMissing(currentPeriod, moduleId as ModuleId, flag);
    }
    // 缺失标记（无此业务/延迟）影响引擎输入口径 → 失效结果缓存，
    // 避免结果页继续展示基于旧标记生成的旧结果
    invalidateResultCache(currentPeriod);
  };

  const handleSave = async () => {
    if (!moduleId) return;
    setSaving(true);
    try {
      // 保存到本地
      setModuleFiles(currentPeriod, moduleId as ModuleId, files);
      setModuleText(currentPeriod, moduleId as ModuleId, textInput);
      setModuleMissing(currentPeriod, moduleId as ModuleId, missingFlag);

      // 同步文本到后端（引擎也可能用到补充说明）。text 为空也调用，
      // 用空串覆盖旧文件，避免清空文本后后端仍残留旧内容。
      uploadTextToBackend(currentPeriod, moduleId, textInput).catch(() => {
        addToast({ type: 'warning', message: '补充说明已保存在本地，但同步后端失败' });
      });

      // 文本/缺失标记变更同样影响运算输入 → 失效结果缓存，
      // 与上传/删除文件保持一致，避免结果页展示基于旧输入的缓存结果
      invalidateResultCache(currentPeriod);

      if (missingFlag === 'not_applicable') {
        setModuleStatus(currentPeriod, moduleId as ModuleId, 'na');
      } else if (files.length > 0 || textInput.trim().length > 0) {
        setModuleStatus(currentPeriod, moduleId as ModuleId, 'completed');
      } else if (missingFlag === 'delayed') {
        setModuleStatus(currentPeriod, moduleId as ModuleId, 'partial');
      } else {
        setModuleStatus(currentPeriod, moduleId as ModuleId, 'none');
      }

      setIsDirty(false);
      addToast({
        type: 'success',
        message: `✅ ${moduleDef?.subtitle || moduleId} 已保存`,
      });

      // 成功路径不再复位 saving：保持按钮禁用直到 1.5s 后跳转完成，
      // 避免跳转延迟期间按钮恢复可点导致重复保存。
      setTimeout(() => {
        navigate('/');
      }, 1500);
    } catch {
      setSaving(false);
      addToast({ type: 'error', message: '保存失败，请检查后重试' });
    }
  };

  const handleCancel = () => {
    if (isDirty) {
      setPendingNavigation('/');
      setShowLeaveConfirm(true);
    } else {
      navigate('/');
    }
  };

  if (!moduleDef) {
    return (
      <div className="max-w-[1200px] mx-auto px-6 py-12 text-center">
        <p className="text-text-disabled text-lg">模块未找到: {moduleId}</p>
        <button
          onClick={() => navigate('/')}
          className="mt-4 text-primary-500 hover:text-primary-600 transition-fast cursor-pointer"
        >
          ← 返回首页
        </button>
      </div>
    );
  }

  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const getFileIcon = (name: string): string => {
    const ext = name.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'pdf': return '📄';
      case 'xlsx':
      case 'xls': return '📊';
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif': return '🖼';
      default: return '📁';
    }
  };

  return (
    <div className="max-w-[1200px] mx-auto px-6 py-8 flex flex-col gap-6">
      <button
        onClick={handleBack}
        className="flex items-center gap-1 text-sm text-text-secondary hover:text-primary-500 
                   transition-fast cursor-pointer w-fit"
      >
        <ArrowLeft size={16} />
        <span>返回首页</span>
      </button>

      <div>
        <h1 className="text-xl font-bold text-text-primary flex items-center gap-2">
          <span className="text-2xl">{moduleDef.icon}</span>
          {moduleDef.title}
        </h1>
        <p className="text-sm text-text-secondary mt-1">{moduleDef.description}</p>
      </div>

      {/* Upload Zone */}
      <div
        className={`border-2 border-dashed rounded-xl p-10 text-center transition-all duration-200 cursor-pointer
          ${isDragging
            ? 'border-primary-500 bg-primary-50'
            : 'border-gray-300 bg-gray-50 hover:border-primary-300 hover:bg-primary-50/30'
          }`}
        onClick={handleDropZoneClick}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') fileInputRef.current?.click();
        }}
      >
        <Upload size={40} className={`mx-auto mb-3 ${isDragging ? 'text-primary-500' : 'text-gray-400'}`} />
        <p className="text-text-secondary text-sm mb-1">
          拖拽文件到此处，或<span className="text-primary-500 font-medium">点击选择文件</span>
        </p>
        <p className="text-xs text-text-disabled">
          支持 PDF / JPG / PNG / Excel 批量上传
        </p>
        <p className="text-xs text-text-disabled">
          （单文件≤20MB，可批量上传）
        </p>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept={moduleDef.uploadAccept}
          onChange={handleFileInputChange}
          className="hidden"
          aria-label="选择文件上传"
        />
      </div>

      {/* Upload Progress */}
      {Object.entries(uploadProgress).length > 0 && (
        <div className="space-y-2">
          {Object.entries(uploadProgress).map(([key, progress]) => {
            // key = {period}/{module}/{timestamp}-{random}-{fileName}，取文件名需跳过两个前缀段
            const segments = key.split('/');
            const last = segments[segments.length - 1] || key;
            const fileName = last.split('-').slice(2).join('-') || last || key;
            return (
              <div key={key} className="flex items-center gap-3 bg-white border border-gray-200 rounded-md px-4 py-2">
                <div className="w-4 h-4 rounded-full border-2 border-primary-500 border-t-transparent animate-spin" />
                <span className="text-sm text-text-secondary flex-1 truncate">{fileName}</span>
                <div className="w-24 h-1.5 bg-gray-200 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-primary-500 rounded-full transition-all duration-300"
                    style={{ width: `${Math.round(progress * 100)}%` }}
                  />
                </div>
                <span className="text-xs text-text-disabled w-10 text-right">
                  {Math.round(progress * 100)}%
                </span>
              </div>
            );
          })}
        </div>
      )}

      {/* File List */}
      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <h3 className="text-base font-semibold text-text-primary mb-4">已上传文件列表</h3>
        {files.length === 0 ? (
          <div
            className={`text-center py-8 text-text-disabled text-sm bg-gray-50 rounded-md 
              ${missingFlag === 'not_applicable' ? 'opacity-50' : ''}`}
          >
            {missingFlag === 'not_applicable'
              ? '已标记为无此业务，文件可删除'
              : '暂无文件，请拖拽上传或点击选择'}
          </div>
        ) : (
          <div className="space-y-2">
            {files.map((file) => (
              <div
                key={file.id}
                className={`flex items-center gap-3 px-4 py-3 border border-gray-100 rounded-md 
                  hover:bg-gray-50 transition-fast
                  ${missingFlag === 'not_applicable' ? 'opacity-50' : ''}`}
              >
                <span className="text-xl">{getFileIcon(file.name)}</span>
                <div className="flex-1 min-w-0">
                  <p className="text-sm text-text-primary truncate">{file.name}</p>
                  <p className="text-xs text-text-disabled">{formatSize(file.size)}</p>
                </div>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    handlePreviewFile(file);
                  }}
                  className="p-1.5 text-text-secondary hover:text-primary-500 hover:bg-primary-50 
                             rounded-md transition-fast cursor-pointer"
                  title="预览"
                  aria-label={`预览 ${file.name}`}
                >
                  <Eye size={16} />
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setFileToDelete(file);
                  }}
                  className="p-1.5 text-text-secondary hover:text-error hover:bg-error-light 
                             rounded-md transition-fast cursor-pointer"
                  title="删除"
                  aria-label={`删除 ${file.name}`}
                >
                  <Trash2 size={16} />
                </button>
              </div>
            ))}
          </div>
        )}
      </section>

      {/* Text Input */}
      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-base font-semibold text-text-primary">{moduleDef.textLabel}</h3>
          <span className="text-xs text-text-disabled bg-gray-100 px-2 py-0.5 rounded">选填</span>
        </div>
        <textarea
          value={textInput}
          onChange={(e) => {
            const value = e.target.value;
            setTextInput(value);
            markDirty();
            // 立即落盘，避免导航绕过保存导致文本丢失
            if (moduleId) {
              setModuleText(currentPeriod, moduleId as ModuleId, value);
            }
          }}
          placeholder={moduleDef.textPlaceholder}
          rows={6}
          className="w-full px-4 py-3 text-sm text-text-primary border border-gray-200 rounded-lg 
                     focus:outline-none focus:ring-2 focus:ring-primary-500 focus:border-transparent
                     placeholder:text-text-disabled resize-y transition-fast"
          aria-label={moduleDef.textLabel}
        />
        <p className="text-xs text-text-disabled mt-2">建议格式：日期 + 业务摘要 + 金额（每行一条）</p>
      </section>

      {/* Missing Flag */}
      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <div className="flex items-start gap-2 mb-3">
          <span className="text-lg">⚠</span>
          <div>
            <h3 className="text-base font-semibold text-text-primary">缺失标记</h3>
            <p className="text-sm text-text-secondary">如果本模块没有资料需要上传：</p>
          </div>
        </div>
        <div className="flex flex-col gap-3 ml-7">
          <label className="flex items-center gap-3 cursor-pointer">
            <input
              type="radio"
              name="missing-flag"
              value="delayed"
              checked={missingFlag === 'delayed'}
              onChange={() => handleMissingChange('delayed')}
              className="w-4 h-4 text-primary-500 focus:ring-primary-500"
            />
            <span className="text-sm text-text-primary">资料暂未提供（后续补充）</span>
          </label>
          <label className="flex items-center gap-3 cursor-pointer">
            <input
              type="radio"
              name="missing-flag"
              value="not_applicable"
              checked={missingFlag === 'not_applicable'}
              onChange={() => handleMissingChange('not_applicable')}
              className="w-4 h-4 text-primary-500 focus:ring-primary-500"
            />
            <span className="text-sm text-text-primary">本企业本月无该项业务</span>
          </label>
        </div>
      </section>

      {/* Action Buttons */}
      <div className="flex gap-4 justify-end">
        <button
          onClick={handleCancel}
          className="px-6 py-2.5 text-sm font-medium text-text-secondary bg-white 
                     border border-gray-300 hover:bg-gray-50 rounded-lg transition-fast cursor-pointer"
        >
          取消返回
        </button>
        <button
          onClick={handleSave}
          disabled={saving}
          className="px-6 py-2.5 text-sm font-medium text-white bg-primary-500 
                     hover:bg-primary-600 disabled:bg-primary-300 disabled:cursor-not-allowed 
                     rounded-lg transition-fast flex items-center gap-2 cursor-pointer"
        >
          {saving ? (
            <>
              <div className="w-4 h-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
              <span>保存中...</span>
            </>
          ) : (
            <span>💾 保存</span>
          )}
        </button>
      </div>

      {/* 删除文件确认 */}
      <ConfirmDialog
        isOpen={!!fileToDelete}
        title="删除文件"
        message={`确认永久删除「${fileToDelete?.name || ''}」？\n文件将从服务端和本地缓存中一并移除，无法恢复。`}
        confirmText="确认删除"
        cancelText="取消"
        variant="danger"
        onConfirm={() => {
          const target = fileToDelete;
          setFileToDelete(null);
          if (target) {
            handleDeleteFile(target);
          }
        }}
        onCancel={() => setFileToDelete(null)}
      />

      <ConfirmDialog
        isOpen={showLeaveConfirm}
        title="未保存的内容"
        message="您有未保存的内容，是否放弃？"
        confirmText="放弃离开"
        cancelText="继续编辑"
        variant="danger"
        onConfirm={handleConfirmLeave}
        onCancel={() => {
          setShowLeaveConfirm(false);
          setPendingNavigation(null);
        }}
      />

      {/* File Preview Modal */}
      <FilePreview
        isOpen={!!previewFile}
        onClose={() => setPreviewFile(null)}
        fileName={previewFile?.file.name || ''}
        blob={previewFile?.blob || new Blob()}
      />
    </div>
  );
};

export default CategoryInputPage;
