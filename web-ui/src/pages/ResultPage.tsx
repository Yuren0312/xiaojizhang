import React, { useState, useEffect, useCallback, useRef } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAppStore } from '../store/appStore';
import { setResultFiles, setResultGeneratedAt, invalidateResultCache, clearPeriodCache, getResultFilesKey, getResultGeneratedAtKey, getPreferences } from '../services/localStorage';
import {
  API_BASE,
  downloadResult,
  archivePeriod,
  startComputeTask,
  getComputeTask,
  type OutputFile,
  type ComputeResponse,
} from '../services/api';
import { ArrowLeft, Eye, Download, Package, RotateCcw, Archive } from 'lucide-react';
import FilePreview from '../components/ui/FilePreview';
import ConfirmDialog from '../components/ui/ConfirmDialog';
import type { ResultFile, ExtensionFile } from '../types';

const EXTENSION_FILES: ExtensionFile[] = [
  { id: 'ext-1', name: '增值税申报底稿', description: '已支持生成，本期需发票数据源', status: 'ready', eta: '需补充发票台账/PDF' },
  { id: 'ext-2', name: '个税扣缴明细表', description: '已支持生成，本期需工资数据源', status: 'ready', eta: '需补充工资表' },
  { id: 'ext-3', name: '企业所得税预缴底稿', description: '已支持生成，本期需资产/员工数据源', status: 'ready', eta: '需补充资产台账/员工名册' },
];

type PagePhase = 'loading' | 'computing' | 'success' | 'error';

function apiOutputsToResultFiles(outputs: OutputFile[]): ResultFile[] {
  return outputs.map((o) => ({
    id: o.filename,
    name: o.filename,
    size: o.size,
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    label: o.label,
    description: '',
    indexedDBKey: o.download_url,
  }));
}

const ResultPage: React.FC = () => {
  const navigate = useNavigate();
  const { currentPeriod, addToast } = useAppStore();
  const computeRef = useRef<AbortController | null>(null);

  const [phase, setPhase] = useState<PagePhase>('loading');
  const [progress, setProgress] = useState(0);
  const [progressText, setProgressText] = useState('正在连接引擎...');
  const [errorMessage, setErrorMessage] = useState('');
  const [errorDetail, setErrorDetail] = useState('');
  const [resultFiles, setResultFilesState] = useState<ResultFile[]>([]);
  const [stats, setStats] = useState<ComputeResponse['stats'] | null>(null);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [isStale, setIsStale] = useState(false);
  const [previewFile, setPreviewFile] = useState<{ file: ResultFile; blob: Blob } | null>(null);
  const [packaging, setPackaging] = useState(false);
  const [showArchiveConfirm, setShowArchiveConfirm] = useState(false);
  const [archiving, setArchiving] = useState(false);

  const runComputation = useCallback(async () => {
    setPhase('computing');
    setProgress(0);
    setProgressText('正在连接引擎...');
    setErrorMessage('');
    setErrorDetail('');
    setResultFilesState([]);
    setIsStale(false);
    setStats(null);
    setWarnings([]);

    const abortController = new AbortController();
    computeRef.current = abortController;

    try {
      const [yearStr, monthStr] = currentPeriod.split('-');
      const year = parseInt(yearStr, 10);
      const month = parseInt(monthStr, 10);

      // 异步任务化：创建任务 → 轮询真实进度（不再模拟假进度）
      // 传入外部 signal，账期切换/取消时能中断在途请求与轮询
      const { task_id } = await startComputeTask(year, month, getPreferences().entityType || 'corp', abortController.signal);
      if (abortController.signal.aborted) return;

      let finalState: Awaited<ReturnType<typeof getComputeTask>> | null = null;
      // eslint-disable-next-line no-constant-condition
      while (true) {
        if (abortController.signal.aborted) return;
        const state = await getComputeTask(task_id, abortController.signal);
        if (abortController.signal.aborted) return; // 轮询返回途中被取消，丢弃旧任务结果
        setProgress(state.progress);
        setProgressText(state.message);
        if (state.status === 'done' || state.status === 'done_with_errors' || state.status === 'error') {
          finalState = state;
          break;
        }
        await new Promise((r) => setTimeout(r, 800));
      }

      if (abortController.signal.aborted) return; // 循环退出后再确认一次，防止污染新账期

      if (finalState.status === 'error') {
        throw new Error(finalState.error || '运算失败');
      }
      const result = finalState.result as ComputeResponse;

      if (result.outputs.length > 0) {
        setProgress(100);
        setProgressText('运算完成！');

        const files = apiOutputsToResultFiles(result.outputs);
        setResultFilesState(files);
        setStats(result.stats);
        setWarnings(result.warnings);

        // 缓存到本地
        setResultFiles(currentPeriod, files);
        setResultGeneratedAt(currentPeriod, new Date().toISOString());

        if (result.errors.length > 0) {
          setErrorMessage(result.errors.join('；'));
          addToast({ type: 'warning', message: '部分输出存在错误，请查看详情' });
        } else {
          addToast({ type: 'success', message: '✅ 账务文件生成完成！' });
        }

        setPhase('success');
      } else if (result.errors.length > 0) {
        setPhase('error');
        setErrorMessage(result.errors.join('；'));
        setErrorDetail('请检查输入资料是否正确，或联系管理员查看引擎日志。');
      } else {
        setPhase('error');
        setErrorMessage('未生成任何输出文件');
        setErrorDetail('当月可能无业务数据或输入资料缺失');
      }
    } catch (err: any) {
      if (abortController.signal.aborted) return;

      const msg = err?.message || err?.toString() || '未知错误';
      setPhase('error');
      setErrorMessage(`运算失败：${msg}`);
      setErrorDetail('请检查后端服务是否正常运行，或查看引擎日志后重试');
      addToast({ type: 'error', message: '运算失败' });
    }
  }, [currentPeriod, addToast]);

  // 页面加载：先检查是否有缓存
  useEffect(() => {
    const cached = localStorage.getItem(getResultFilesKey(currentPeriod));
    const cachedAt = localStorage.getItem(getResultGeneratedAtKey(currentPeriod));

    if (cached && cachedAt) {
      try {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setResultFilesState(parsed);
          setPhase('success');
          setProgress(100);
          // Stale if cached > 30 min ago
          const ageMs = Date.now() - new Date(cachedAt).getTime();
          setIsStale(ageMs > 30 * 60 * 1000);
          return;
        }
      } catch { /* ignore */ }
    }

    // 没有有效缓存，自动触发运算
    runComputation();

    return () => {
      computeRef.current?.abort();
    };
  }, [currentPeriod, runComputation]);

  const handlePreview = async (file: ResultFile) => {
    try {
      // 先从 IndexedDB 本地找（缓存），没有再从后端下载
      const localStore = await import('../services/indexedDB');
      const local = await localStore.getFile(file.indexedDBKey, 'results');
      let blob: Blob;

      if (local) {
        blob = local.data;
      } else {
        // 从后端下载
        const result = await downloadResult(file.indexedDBKey);
        blob = result.blob;
        // 存到本地缓存
        await localStore.storeFile(file.indexedDBKey, blob, 'results');
      }

      setPreviewFile({ file, blob });
    } catch {
      addToast({ type: 'error', message: '预览文件失败' });
    }
  };

  const handleClosePreview = () => {
    setPreviewFile(null);
  };

  const handleDownloadZip = async (url: string, filename: string) => {
    try {
      const fullUrl = `${API_BASE || ''}${url}`;
      const res = await fetch(fullUrl);
      if (!res.ok) throw new Error(`下载失败 (${res.status})`);
      const blob = await res.blob();
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(a.href);
      addToast({ type: 'success', message: '备份下载完成' });
    } catch {
      addToast({ type: 'error', message: '备份下载失败，请检查后端服务' });
    }
  };

  const handleBackupAll = () => {
    handleDownloadZip('/api/export/all', `小虾记账_全量备份_${new Date().toISOString().slice(0,10)}.zip`);
  };

  const handleDownload = async (file: ResultFile) => {
    try {
      const result = await downloadResult(file.indexedDBKey);
      const url = URL.createObjectURL(result.blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `[${currentPeriod}] ${file.name}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch {
      addToast({ type: 'error', message: '下载失败' });
    }
  };

  const handleDownloadAll = async () => {
    setPackaging(true);
    try {
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();

      for (const file of resultFiles) {
        try {
          const result = await downloadResult(file.indexedDBKey);
          zip.file(file.name, result.blob);
        } catch {
          addToast({ type: 'warning', message: `下载失败: ${file.name}` });
        }
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `[${currentPeriod}] 账务文件包.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      addToast({ type: 'success', message: '打包下载完成' });
    } catch {
      addToast({ type: 'error', message: '打包下载失败' });
    } finally {
      setPackaging(false);
    }
  };

  const handleRegenerate = async () => {
    setResultFiles(currentPeriod, []);
    // 只失效本账期的结果缓存（localStorage result 键 + IndexedDB 结果 blob），
    // 不清除各模块的 files/text/missing/status 元数据，避免首页状态被误清空。
    // 注意：结果 blob 的 key 是 download_url（/api/results/{period}/...），
    // invalidateResultCache 内部按该前缀清理。
    await invalidateResultCache(currentPeriod);
    runComputation();
  };

  const handleArchive = async () => {
    setShowArchiveConfirm(false);
    setArchiving(true);
    try {
      const res = await archivePeriod(currentPeriod);
      if (res.success) {
        clearPeriodCache(currentPeriod);
        addToast({ type: 'success', message: `✅ 账期 ${currentPeriod} 归档完成` });
        navigate('/');
      } else {
        addToast({ type: 'error', message: res.message });
      }
    } catch (err: any) {
      const msg = err?.message || err?.toString() || '未知错误';
      addToast({ type: 'error', message: `归档失败：${msg}` });
    } finally {
      setArchiving(false);
    }
  };

  const formatSize = (bytes: number): string => {
    if (!bytes) return '-';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const coreOutputs = resultFiles.filter((f) =>
    ['会计报表', '银行对账报告', '核算汇总表', '分析简报', '记账凭证', '试算平衡表'].includes(f.label)
  );
  const extOutputs = resultFiles.filter((f) =>
    ['增值税申报底稿', '所得税预缴底稿', '个税扣缴明细表'].includes(f.label)
  );

  return (
    <div className="max-w-[1200px] mx-auto px-6 py-8 flex flex-col gap-6">
      <button
        onClick={() => navigate('/')}
        className="flex items-center gap-1 text-sm text-text-secondary hover:text-primary-500 
                   transition-fast cursor-pointer w-fit"
      >
        <ArrowLeft size={16} />
        <span>返回首页</span>
      </button>

      {/* Stale Warning */}
      {isStale && (
        <div className="bg-warning-light border border-warning rounded-lg px-4 py-3 flex items-center gap-2">
          <span className="text-warning">⚠</span>
          <p className="text-sm text-text-primary">
            输入资料可能已更新，当前结果可能是旧的，建议重新生成
          </p>
          <button
            onClick={handleRegenerate}
            className="ml-auto text-sm text-primary-500 hover:text-primary-600 font-medium 
                       flex items-center gap-1 transition-fast cursor-pointer"
          >
            <RotateCcw size={14} />
            <span>重新生成</span>
          </button>
        </div>
      )}

      {/* Error Display */}
      {phase === 'error' && (
        <div className="bg-error-light border border-error rounded-lg p-6">
          <div className="flex items-start gap-3 mb-4">
            <span className="text-error text-xl">❌</span>
            <div>
              <h3 className="text-lg font-semibold text-text-primary">运算失败</h3>
              <p className="text-sm text-text-secondary mt-1">{errorMessage}</p>
            </div>
          </div>
          {errorDetail && (
            <div className="text-sm text-text-secondary ml-10 mb-4">
              <p className="text-xs text-text-disabled">{errorDetail}</p>
            </div>
          )}
          <div className="text-sm text-text-secondary ml-10 mb-4">
            <p className="font-medium">修复建议：</p>
            <ul className="list-disc list-inside mt-1 space-y-1 text-xs">
              <li>确认后端服务正常运行（可访问本系统根路径验证）</li>
              <li>检查输入资料是否完整（至少需要银行对账单或凭证）</li>
              <li>查看「输入」目录下对应账期的文件</li>
              <li>如果是摘要匹配失败，需要在 rules/mapping.yaml 中添加映射</li>
            </ul>
          </div>
          <div className="flex gap-3 ml-10">
            <button
              onClick={handleRegenerate}
              className="px-4 py-2 text-sm font-medium text-white bg-error hover:bg-red-600 
                         rounded-lg transition-fast cursor-pointer flex items-center gap-2"
            >
              <RotateCcw size={14} />
              <span>重新运算</span>
            </button>
            <button
              onClick={() => navigate('/')}
              className="px-4 py-2 text-sm font-medium text-text-secondary bg-white 
                         border border-gray-300 hover:bg-gray-50 rounded-lg transition-fast cursor-pointer"
            >
              返回首页修改输入
            </button>
          </div>
        </div>
      )}

      {/* Loading */}
      {phase === 'loading' && (
        <section className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm flex items-center justify-center py-20">
          <div className="text-center space-y-3">
            <div className="w-10 h-10 rounded-full border-[3px] border-primary-200 border-t-primary-500 animate-spin mx-auto" />
            <p className="text-text-secondary text-sm">正在检查已有结果...</p>
          </div>
        </section>
      )}

      {/* Computing Progress */}
      {phase === 'computing' && (
        <section className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
          <div className="flex items-center justify-between mb-1">
            <h2 className="text-lg font-semibold text-text-primary">📊 正在生成账务文件...</h2>
            <button
              onClick={() => {
                computeRef.current?.abort();
                addToast({ type: 'info', message: '已取消本次运算' });
                navigate('/');
              }}
              className="px-3 py-1.5 text-xs font-medium text-text-secondary bg-gray-100 
                         hover:bg-gray-200 rounded-md transition-fast cursor-pointer"
            >
              取消运算
            </button>
          </div>
          <p className="text-sm text-text-secondary">{progressText}</p>

          <div className="relative h-3 bg-gray-200 rounded-full overflow-hidden mt-4 mb-3">
            <div
              className="absolute left-0 top-0 h-full bg-primary-500 rounded-full transition-all duration-500 ease-out animate-progress-pulse"
              style={{ width: `${progress}%` }}
            />
          </div>

          <div className="text-right text-sm text-text-secondary">{progress}%</div>

          {progress > 80 && (
            <p className="text-xs text-warning mt-3">运算进入最后阶段，请耐心等待...</p>
          )}
        </section>
      )}

      {/* Success Results */}
      {phase === 'success' && (
        <>
          {/* Summary Banner */}
          <section className="bg-success-light border border-success rounded-lg px-6 py-4 shadow-sm">
            <div className="flex items-center gap-2">
              <span className="text-xl">✅</span>
              <div>
                <p className="font-semibold text-text-primary">
                  {currentPeriod} 账务文件生成完成！
                </p>
                <p className="text-sm text-text-secondary">
                  {stats ? (
                    <>凭证 {stats.voucher_count} 笔，银行流水 {stats.bank_txn_count} 笔，分录 {stats.journal_entry_count} 笔</>
                  ) : (
                    '从缓存加载'
                  )}
                </p>
              </div>
            </div>
            {warnings.length > 0 && (
              <div className="mt-2 text-xs text-warning">
                {warnings.map((w, i) => (
                  <p key={i}>⚠ {w}</p>
                ))}
              </div>
            )}
            {errorMessage && (
              <div className="mt-2 text-xs text-error bg-error-light border border-error rounded-md px-3 py-2">
                <p className="font-medium">部分输出存在问题：</p>
                {errorMessage.split('；').map((e, i) => (
                  <p key={i}>• {e}</p>
                ))}
              </div>
            )}
          </section>

          {/* Core Output Files */}
          <section className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
            <h3 className="text-base font-semibold text-text-primary mb-4">核心输出文件</h3>
            {coreOutputs.length === 0 ? (
              <div className="text-center py-6 text-text-disabled text-sm">
                暂未生成核心文件
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {coreOutputs.map((file) => (
                  <div
                    key={file.id}
                    className="border border-gray-200 rounded-lg p-4 flex flex-col gap-3
                               hover:shadow-md transition-normal"
                  >
                    <div className="flex items-start gap-3">
                      <span className="text-3xl">📊</span>
                      <div className="flex-1 min-w-0">
                        <p className="font-medium text-text-primary">{file.label}</p>
                        <p className="text-xs text-text-disabled truncate">{file.name}</p>
                        <p className="text-xs text-text-disabled">{formatSize(file.size)}</p>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <button
                        onClick={() => handlePreview(file)}
                        className="flex-1 py-2 text-sm text-primary-500 bg-primary-50 
                                   hover:bg-primary-100 rounded-md transition-fast cursor-pointer
                                   flex items-center justify-center gap-1"
                      >
                        <Eye size={14} />
                        <span>在线预览</span>
                      </button>
                      <button
                        onClick={() => handleDownload(file)}
                        className="flex-1 py-2 text-sm text-primary-500 bg-primary-50 
                                   hover:bg-primary-100 rounded-md transition-fast cursor-pointer
                                   flex items-center justify-center gap-1"
                      >
                        <Download size={14} />
                        <span>下载</span>
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* Download All */}
            {resultFiles.length > 0 && (
              <div className="mt-6 p-4 bg-gray-50 rounded-lg flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <Package size={20} className="text-text-secondary" />
                  <span className="text-sm text-text-primary">导出与备份</span>
                </div>
                <div className="flex gap-3 flex-wrap">
                  <button
                    onClick={handleDownloadAll}
                    disabled={packaging}
                    className="px-4 py-2 text-sm font-medium text-white bg-primary-500 
                               hover:bg-primary-600 disabled:bg-primary-300 disabled:cursor-not-allowed
                               rounded-lg transition-fast cursor-pointer flex items-center gap-2"
                  >
                    {packaging ? (
                      <>
                        <div className="w-4 h-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                        <span>打包中...</span>
                      </>
                    ) : (
                      <>
                        <Download size={14} />
                        <span>当月报表.zip</span>
                      </>
                    )}
                  </button>
                  <button
                    onClick={handleBackupAll}
                    disabled={packaging}
                    className="px-4 py-2 text-sm font-medium text-primary-500 bg-primary-50 
                               hover:bg-primary-100 disabled:bg-gray-100 disabled:text-text-disabled disabled:cursor-not-allowed
                               border border-primary-200 rounded-lg transition-fast cursor-pointer flex items-center gap-2"
                  >
                    <Download size={14} />
                    <span>全量备份.zip</span>
                  </button>
                  <button
                    onClick={() => setShowArchiveConfirm(true)}
                    disabled={archiving}
                    className="px-4 py-2 text-sm font-medium text-orange bg-orange-light 
                               hover:bg-orange-100 disabled:bg-gray-100 disabled:text-text-disabled disabled:cursor-not-allowed
                               border border-orange-200 rounded-lg transition-fast cursor-pointer flex items-center gap-2"
                  >
                    {archiving ? (
                      <>
                        <div className="w-4 h-4 rounded-full border-2 border-orange border-t-transparent animate-spin" />
                        <span>归档中...</span>
                      </>
                    ) : (
                      <>
                        <Archive size={14} />
                        <span>归档本账期</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}
          </section>

          {/* Extension Outputs */}
          <section className={`bg-white rounded-lg border border-gray-200 p-6 shadow-sm ${extOutputs.length === 0 ? 'opacity-60' : ''}`}>
            <h3 className="text-base font-semibold text-text-primary mb-4">
              扩展输出区
              <span className="text-xs text-text-disabled font-normal ml-2">
                {extOutputs.length > 0 ? '税务申报底稿' : '暂无数据'}
              </span>
            </h3>
            {extOutputs.length > 0 ? (
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {extOutputs.map((file) => {
                  const isTemplate = file.name.includes('[模板]');
                  return (
                  <div
                    key={file.id}
                    className={`border rounded-lg p-3 flex flex-col gap-2
                               hover:shadow-md transition-normal
                               ${isTemplate ? 'border-amber-300 bg-amber-50/50' : 'border-gray-200'}`}
                  >
                    <div className="flex items-center gap-2">
                      <span className="text-xl">📋</span>
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-medium text-text-primary flex items-center gap-1.5">
                          {file.label}
                          {isTemplate && (
                            <span className="text-[10px] font-medium text-amber-700 bg-amber-100 
                                           border border-amber-200 rounded-full px-1.5 py-0.5 whitespace-nowrap">
                              模板版
                            </span>
                          )}
                        </p>
                        <p className="text-xs text-text-disabled truncate">{file.name}</p>
                        <p className="text-xs text-text-disabled">{formatSize(file.size)}</p>
                      </div>
                    </div>
                    {isTemplate && (
                      <p className="text-[11px] text-amber-700 leading-snug">
                        缺少数据源（发票台账/工资表/资产台账），当前为模板输出，仅供参考，不可直接申报
                      </p>
                    )}
                    <div className="flex gap-2">
                      <button
                        onClick={() => handlePreview(file)}
                        className="flex-1 py-1.5 text-xs text-primary-500 bg-primary-50 
                                   hover:bg-primary-100 rounded-md transition-fast cursor-pointer"
                      >
                        预览
                      </button>
                      <button
                        onClick={() => handleDownload(file)}
                        className="flex-1 py-1.5 text-xs text-primary-500 bg-primary-50 
                                   hover:bg-primary-100 rounded-md transition-fast cursor-pointer"
                      >
                        下载
                      </button>
                    </div>
                  </div>
                  );
                })}
              </div>
            ) : (
              <div className="space-y-3">
                {EXTENSION_FILES.map((ext) => (
                  <div
                    key={ext.id}
                    className="flex items-center gap-4 px-4 py-3 border border-gray-100 rounded-md bg-gray-50"
                  >
                    <span className="text-xl">📋</span>
                    <div className="flex-1">
                      <p className="text-sm font-medium text-text-primary">{ext.name}</p>
                      <p className="text-xs text-text-disabled">{ext.description}</p>
                    </div>
                    <span className="text-xs text-warning bg-warning-light px-2 py-1 rounded-full">
                      ⏳ {ext.eta}
                    </span>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* Regenerate */}
          <div className="flex justify-end">
            <button
              onClick={handleRegenerate}
              className="px-5 py-2 text-sm font-medium text-text-secondary bg-white 
                         border border-gray-300 hover:bg-gray-50 rounded-lg transition-fast 
                         cursor-pointer flex items-center gap-2"
            >
              <RotateCcw size={14} />
              <span>重新生成</span>
            </button>
          </div>
        </>
      )}

      {/* Preview Modal */}
      <FilePreview
        isOpen={!!previewFile}
        onClose={handleClosePreview}
        fileName={previewFile?.file.name || ''}
        blob={previewFile?.blob || new Blob()}
        onDownload={() => previewFile && handleDownload(previewFile.file)}
      />

      {/* Archive Confirm Dialog */}
      <ConfirmDialog
        isOpen={showArchiveConfirm}
        title="归档本账期"
        message={`确认将 ${currentPeriod} 的输入资料和输出文件归档？\n归档后文件将移动到「归档」目录，当前账期缓存将被清理。`}
        confirmText="确认归档"
        cancelText="取消"
        onConfirm={handleArchive}
        onCancel={() => setShowArchiveConfirm(false)}
        variant="default"
      />
    </div>
  );
};

export default ResultPage;
