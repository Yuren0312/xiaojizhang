/**
 * 共享文件预览组件
 *
 * 支持格式：
 *   - xlsx/xls → 表格预览（前50行）
 *   - pdf → iframe 嵌入预览（Chrome/Edge 内置 PDF 查看器）
 *   - jpg/jpeg/png/gif/webp/bmp/svg → 图片预览
 *   - txt/csv/md → 文本预览
 *   - 其他 → 下载引导
 */
import React, { useState, useEffect } from 'react';
import Modal from './Modal';
import { Download } from 'lucide-react';

type PreviewFormat = 'excel' | 'pdf' | 'image' | 'text' | 'unknown';

interface FilePreviewProps {
  /** 是否显示 */
  isOpen: boolean;
  /** 关闭回调 */
  onClose: () => void;
  /** 文件名 */
  fileName: string;
  /** 文件 Blob */
  blob: Blob;
  /** 可选：下载回调 */
  onDownload?: () => void;
}

function detectFormat(name: string): PreviewFormat {
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (ext === 'xlsx' || ext === 'xls') return 'excel';
  if (ext === 'pdf') return 'pdf';
  if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg'].includes(ext)) return 'image';
  if (['txt', 'csv', 'md', 'log', 'json', 'xml', 'yaml', 'yml'].includes(ext)) return 'text';
  return 'unknown';
}

const FilePreview: React.FC<FilePreviewProps> = ({
  isOpen,
  onClose,
  fileName,
  blob,
  onDownload,
}) => {
  const format = detectFormat(fileName);
  const [objectUrl, setObjectUrl] = useState('');
  const [excelData, setExcelData] = useState<{
    headers: string[];
    rows: string[][];
  } | null>(null);
  const [textContent, setTextContent] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  useEffect(() => {
    if (!isOpen || !blob) return;

    const url = URL.createObjectURL(blob);
    setObjectUrl(url);
    setLoading(true);
    setError('');
    setExcelData(null);
    setTextContent('');

    if (format === 'excel') {
      loadExcel(blob);
    } else if (format === 'text') {
      loadText(blob);
    } else if (format === 'pdf' || format === 'image') {
      // 直接用 blob URL 渲染，无需额外解析
      setLoading(false);
    } else {
      setLoading(false);
    }

    return () => {
      URL.revokeObjectURL(url);
    };
  }, [isOpen, blob, format]);

  // react-hooks/exhaustive-deps 解禁：依赖稳定引用
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const stableLoader = React.useRef(false);
  useEffect(() => { stableLoader.current = true; }, []);

  const loadExcel = async (b: Blob) => {
    try {
      const XLSX = await import('xlsx');
      const buffer = await b.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      if (firstSheet) {
        const sheet = workbook.Sheets[firstSheet];
        const data = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1 });
        const headers = (data[0] as string[]) || [];
        const rows = data.slice(1, 51) as string[][];
        setExcelData({ headers, rows });
      }
      setLoading(false);
    } catch (e: any) {
      setError(`Excel 解析失败: ${e?.message || '未知错误'}`);
      setLoading(false);
    }
  };

  const loadText = async (b: Blob) => {
    try {
      const text = await b.text();
      setTextContent(text.slice(0, 100000)); // 最多 100KB
      setLoading(false);
    } catch (e: any) {
      setError(`文本读取失败: ${e?.message || '未知错误'}`);
      setLoading(false);
    }
  };

  const handleDownload = () => {
    if (onDownload) {
      onDownload();
    } else {
      const a = document.createElement('a');
      a.href = objectUrl;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
  };

  const renderContent = () => {
    if (loading) {
      return (
        <div className="flex items-center justify-center py-20">
          <div className="w-8 h-8 rounded-full border-[3px] border-primary-200 border-t-primary-500 animate-spin" />
        </div>
      );
    }

    if (error) {
      return (
        <div className="flex items-center justify-center py-16">
          <div className="text-center space-y-3">
            <p className="text-error text-sm">{error}</p>
            <button
              onClick={handleDownload}
              className="inline-flex items-center gap-2 text-primary-500 hover:text-primary-600 font-medium text-sm cursor-pointer"
            >
              <Download size={16} />
              <span>下载文件</span>
            </button>
          </div>
        </div>
      );
    }

    switch (format) {
      case 'excel':
        return excelData ? (
          <div>
            <div className="overflow-x-auto max-h-[65vh]">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-gray-100 sticky top-0">
                    {excelData.headers.map((header, idx) => (
                      <th
                        key={idx}
                        className="border border-gray-300 px-3 py-2 text-left font-medium text-text-primary whitespace-nowrap"
                      >
                        {header || `列${idx + 1}`}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {excelData.rows.map((row, rowIdx) => (
                    <tr key={rowIdx} className="hover:bg-gray-50">
                      {excelData.headers.map((_, colIdx) => (
                        <td
                          key={colIdx}
                          className="border border-gray-200 px-3 py-1.5 text-text-secondary whitespace-nowrap"
                        >
                          {row[colIdx] != null ? String(row[colIdx]) : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="text-xs text-text-disabled mt-2 text-center">
              仅显示第一个工作表的1~50行预览。完整数据请下载查看。
            </p>
          </div>
        ) : null;

      case 'pdf':
        return (
          <div className="flex flex-col" style={{ height: '75vh' }}>
            <iframe
              src={objectUrl}
              className="flex-1 w-full border-0 rounded-md"
              title={fileName}
              style={{ minHeight: 500 }}
            />
            <p className="text-xs text-text-disabled mt-2 text-center">
              使用浏览器内置 PDF 查看器预览
            </p>
          </div>
        );

      case 'image':
        return (
          <div className="flex flex-col items-center" style={{ maxHeight: '75vh' }}>
            <img
              src={objectUrl}
              alt={fileName}
              className="max-w-full max-h-[70vh] object-contain rounded-md"
            />
            <p className="text-xs text-text-disabled mt-2 text-center">{fileName}</p>
          </div>
        );

      case 'text':
        return (
          <div>
            <pre className="bg-gray-50 border border-gray-200 rounded-md p-4 text-sm text-text-primary overflow-auto max-h-[65vh] whitespace-pre-wrap font-mono">
              {textContent || '(空文件)'}
            </pre>
            <p className="text-xs text-text-disabled mt-2 text-center">
              仅显示前 100KB 内容
            </p>
          </div>
        );

      default:
        return (
          <div className="flex items-center justify-center py-16">
            <div className="text-center space-y-4">
              <p className="text-text-secondary text-sm">
                此格式暂不支持在线预览，请下载后查看
              </p>
              <button
                onClick={handleDownload}
                className="inline-flex items-center gap-2 text-primary-500 hover:text-primary-600 font-medium text-sm cursor-pointer"
              >
                <Download size={16} />
                <span>下载文件</span>
              </button>
            </div>
          </div>
        );
    }
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} title={fileName} size="full">
      {renderContent()}
    </Modal>
  );
};

export default FilePreview;
