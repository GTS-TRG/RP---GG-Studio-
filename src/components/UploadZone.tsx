import React, { useRef } from 'react';
import { Upload, FileSpreadsheet, Play, CheckCircle2, RefreshCw, FolderOpen, Save } from 'lucide-react';
import { supportsFileSystemAccess } from '../utils/fileSystemAccess';

interface UploadZoneProps {
  onFileUpload: (file: File) => void;
  onOpenWithFileSystem?: () => void;
  onLoadDemo: () => void;
  fileName?: string;
  totalPanels?: number;
  totalCircuits?: number;
  specSheetName?: string | null;
  isLoading?: boolean;
  /** sidebar = gọn trong task panel trái */
  variant?: 'banner' | 'sidebar';
  canWriteOriginal?: boolean;
  pendingPatchCount?: number;
  onSaveToExcel?: () => void;
  onOverwriteOriginal?: () => void;
  isSaving?: boolean;
  /** Doc lai chinh file da mo tren dia */
  onReload?: () => void;
  /** True khi con giu file handle -> Reload duoc */
  canReload?: boolean;
  /** Tu dong theo doi file tren dia */
  autoWatch?: boolean;
  onToggleAutoWatch?: (next: boolean) => void;
  /** File tren dia da doi nhung chua nap (dang co sua chua luu) */
  diskChanged?: boolean;
  lastSyncAt?: Date | null;
}

export const UploadZone: React.FC<UploadZoneProps> = ({
  onFileUpload,
  onOpenWithFileSystem,
  onLoadDemo,
  fileName,
  totalPanels,
  totalCircuits,
  specSheetName,
  isLoading,
  variant = 'sidebar',
  canWriteOriginal = false,
  pendingPatchCount = 0,
  onSaveToExcel,
  onOverwriteOriginal,
  isSaving = false,
  onReload,
  canReload = false,
  autoWatch = false,
  onToggleAutoWatch,
  diskChanged = false,
  lastSyncAt = null,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isSidebar = variant === 'sidebar';
  const fsSupported = supportsFileSystemAccess();

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.match(/\.(xlsx|xlsm|xls)$/i)) {
        onFileUpload(file);
      } else {
        alert('Please select an Excel file (.xlsx, .xlsm, or .xls).');
      }
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFileUpload(e.target.files[0]);
    }
  };

  return (
    <div
      className={`bg-[#FFFFFF] border border-[#D5DEE8] shadow-sm ${
        isSidebar ? 'rounded-2xl p-3.5' : 'rounded-3xl p-3.5'
      }`}
    >
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept=".xlsx, .xlsm, .xls"
        className="hidden"
      />

      {!fileName ? (
        <div
          onDragOver={handleDragOver}
          onDrop={handleDrop}
          className={`border-2 border-dashed border-[#C5D0DC] hover:border-[#2F6F4E] text-center transition-all bg-[#F5F8FB]/60 hover:bg-[#F5F8FB] group ${
            isSidebar ? 'rounded-xl p-4' : 'rounded-2xl p-8'
          }`}
        >
          <div
            className={`mx-auto rounded-2xl bg-[#E8EEF4] border border-[#C5D0DC] flex items-center justify-center text-[#1B7A45] group-hover:scale-105 transition-all shadow-xs ${
              isSidebar ? 'w-10 h-10 mb-2' : 'w-12 h-12 mb-3'
            }`}
          >
            <Upload className={isSidebar ? 'w-5 h-5' : 'w-6 h-6'} />
          </div>
          <h3 className={`font-bold text-[#1A2332] mb-1 ${isSidebar ? 'text-sm leading-snug' : 'text-base'}`}>
            {isSidebar
              ? 'Nạp file Excel bảng tính tải'
              : 'Kéo thả hoặc Bấm vào đây để nạp File Excel Bảng Tính Tải Điện'}
          </h3>
          <p className={`text-[#5A6A7A] leading-relaxed ${isSidebar ? 'text-[11px] mb-3' : 'text-xs max-w-md mx-auto mb-5'}`}>
            Hỗ trợ <span className="text-[#1B7A45] font-mono font-semibold">.xlsx / .xlsm / .xls</span>
            {isSidebar
              ? '. Open workbook to edit, then Save as a new copy.'
              : '. Tự động tìm Spec. Cable và các tủ điện (MSB, DB, LP, ...).'}
          </p>

          <div className={`flex flex-col gap-2 ${isSidebar ? '' : 'flex-wrap items-center justify-center sm:flex-row gap-3 pt-1'}`}>
            {fsSupported && onOpenWithFileSystem && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  onOpenWithFileSystem();
                }}
                className={`bg-[#0369A1] hover:bg-[#075985] text-white font-semibold text-xs rounded-full shadow-xs transition-colors flex items-center justify-center space-x-2 ${
                  isSidebar ? 'w-full px-3 py-2' : 'px-5 py-2.5'
                }`}
              >
                <FolderOpen className="w-4 h-4" />
                <span>Open workbook</span>
              </button>
            )}

            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                fileInputRef.current?.click();
              }}
              className={`bg-[#2F6F4E] hover:bg-[#255A3F] text-white font-semibold text-xs rounded-full shadow-xs transition-colors flex items-center justify-center space-x-2 ${
                isSidebar ? 'w-full px-3 py-2' : 'px-5 py-2.5'
              }`}
            >
              <FileSpreadsheet className="w-4 h-4" />
              <span>Chọn file</span>
            </button>

            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                onLoadDemo();
              }}
              className={`bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] font-medium text-xs rounded-full border border-[#C5D0DC] transition-colors flex items-center justify-center space-x-2 ${
                isSidebar ? 'w-full px-3 py-2' : 'px-5 py-2.5'
              }`}
            >
              <Play className="w-3.5 h-3.5 text-[#2D8A55]" />
              <span>Dùng file Demo</span>
            </button>
          </div>
        </div>
      ) : (
        <div
          className={`bg-[#F5F8FB] border border-[#D5DEE8] ${
            isSidebar ? 'rounded-xl p-3 space-y-3' : 'rounded-2xl p-3 flex flex-col gap-2.5'
          }`}
        >
          <div
            className={`flex min-w-0 ${
              isSidebar ? 'items-start space-x-2.5' : 'flex-1 items-start space-x-3'
            }`}
          >
            <div
              className={`bg-[#E6F4EC] border border-[#A8D4B8] text-[#1B7A45] shrink-0 ${
                isSidebar ? 'p-2 rounded-lg' : 'p-3 rounded-xl'
              }`}
            >
              <FileSpreadsheet className={isSidebar ? 'w-5 h-5' : 'w-6 h-6'} />
            </div>
            <div className="min-w-0 flex-1">
              <div className={`flex ${isSidebar ? 'flex-col gap-1' : 'items-center space-x-2'}`}>
                <span
                  className={`font-bold text-[#1A2332] break-words ${
                    isSidebar ? 'text-sm leading-snug' : 'text-sm truncate max-w-xs sm:max-w-md'
                  }`}
                  title={fileName}
                >
                  {fileName}
                </span>
                <span className="inline-flex items-center self-start px-2 py-0.5 rounded-full text-[10px] font-bold bg-[#E6F4EC] text-[#1B7A45] border border-[#A8D4B8]">
                  <CheckCircle2 className="w-3 h-3 mr-1" /> Đã nạp
                </span>
                {canWriteOriginal && (
                  <span className="inline-flex items-center self-start px-2 py-0.5 rounded-full text-[10px] font-bold bg-[#E0F2FE] text-[#0369A1] border border-[#7DD3FC]">
                    Editable
                  </span>
                )}
              </div>
              <div
                className={`text-[#5A6A7A] mt-1.5 ${
                  isSidebar ? 'flex flex-col gap-0.5 text-[11px]' : 'flex flex-wrap items-center gap-x-4 gap-y-1 text-xs'
                }`}
              >
                <span>
                  Số tủ: <strong className="text-[#1A2332]">{totalPanels}</strong>
                </span>
                {!isSidebar && <span>•</span>}
                <span>
                  Số mạch: <strong className="text-[#1A2332]">{totalCircuits}</strong>
                </span>
                {!isSidebar && <span>•</span>}
                <span className={isSidebar ? 'leading-snug' : ''}>
                  Spec:{' '}
                  <strong className={specSheetName ? 'text-[#1B7A45] font-mono' : 'text-[#5A6A7A]'}>
                    {specSheetName ? `[${specSheetName}]` : 'Mặc định TCVN/IEC'}
                  </strong>
                </span>
                {pendingPatchCount > 0 && (
                  <span className="text-[#0369A1] font-semibold">
                    {pendingPatchCount} unsaved cell(s)
                  </span>
                )}
              </div>
            </div>
          </div>

          <div
            className={`flex gap-2 ${
              isSidebar
                ? 'flex-col'
                : 'flex-row flex-wrap items-center justify-start min-w-0'
            }`}
          >
            {pendingPatchCount > 0 && onSaveToExcel && (
              <button
                onClick={onSaveToExcel}
                disabled={isSaving}
                className={`bg-[#0369A1] hover:bg-[#075985] disabled:opacity-60 text-white text-xs font-semibold rounded-full transition-colors flex items-center justify-center space-x-1.5 ${
                  isSidebar ? 'w-full px-3 py-2' : 'px-4 py-2'
                }`}
              >
                <Save className="w-3.5 h-3.5" />
                <span>{isSaving ? 'Saving…' : `Save as copy (${pendingPatchCount})`}</span>
              </button>
            )}

            {pendingPatchCount > 0 && canWriteOriginal && onOverwriteOriginal && (
              <button
                onClick={onOverwriteOriginal}
                disabled={isSaving}
                className={`bg-[#F0F4F8] hover:bg-[#FEE2E2] text-[#B91C1C] text-xs font-medium rounded-full border border-[#FECACA] transition-colors flex items-center justify-center space-x-1.5 ${
                  isSidebar ? 'w-full px-3 py-2' : 'px-4 py-2'
                }`}
              >
                <span>Overwrite original…</span>
              </button>
            )}

            {fsSupported && onOpenWithFileSystem && (
              <button
                onClick={onOpenWithFileSystem}
                disabled={isLoading}
                className={`bg-[#E0F2FE] hover:bg-[#BAE6FD] text-[#0369A1] text-xs font-medium rounded-full border border-[#7DD3FC] transition-colors flex items-center justify-center space-x-1.5 ${
                  isSidebar ? 'w-full px-3 py-2' : 'px-4 py-2'
                }`}
              >
                <FolderOpen className="w-3.5 h-3.5" />
                <span>Open workbook</span>
              </button>
            )}

            {canReload && onReload ? (
              <button
                onClick={onReload}
                disabled={isLoading}
                title="Đọc lại file từ đĩa và kiểm tra lại ngay"
                className={`disabled:opacity-60 text-white text-xs font-semibold rounded-full transition-colors flex items-center justify-center space-x-1.5 shadow-xs ${
                  diskChanged
                    ? 'bg-[#B45309] hover:bg-[#92400E] animate-pulse'
                    : 'bg-[#2F6F4E] hover:bg-[#255A3F]'
                } ${isSidebar ? 'w-full px-3 py-2' : 'px-4 py-2'}`}
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
                <span>
                  {isLoading
                    ? 'Đang tải lại…'
                    : diskChanged
                    ? 'File đã đổi — Reload'
                    : 'Reload'}
                </span>
              </button>
            ) : (
              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={isLoading}
                className={`bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] text-xs font-medium rounded-full border border-[#C5D0DC] transition-colors flex items-center justify-center space-x-1.5 shadow-2xs ${
                  isSidebar ? 'w-full px-3 py-2' : 'px-4 py-2'
                }`}
              >
                <RefreshCw className={`w-3.5 h-3.5 text-[#2D8A55] ${isLoading ? 'animate-spin' : ''}`} />
                <span>Đổi file khác</span>
              </button>
            )}
          </div>

          {canReload && onToggleAutoWatch && (
            <label
              className="flex items-center gap-2 text-[11px] text-[#5A6A7A] cursor-pointer select-none"
              title="Tự động đọc lại file mỗi khi bạn lưu trong Excel"
            >
              <input
                type="checkbox"
                checked={autoWatch}
                onChange={(e) => onToggleAutoWatch(e.target.checked)}
                className="w-3.5 h-3.5 accent-[#2F6F4E] cursor-pointer"
              />
              <span className="font-medium text-[#1A2332]">Tự động cập nhật khi file thay đổi</span>
              {autoWatch && !diskChanged && (
                <span className="inline-flex items-center gap-1 text-[#1B7A45]">
                  <span className="w-1.5 h-1.5 rounded-full bg-[#2D8A55] animate-pulse" />
                  {lastSyncAt
                    ? `đã đồng bộ ${lastSyncAt.toLocaleTimeString('vi-VN')}`
                    : 'đang theo dõi'}
                </span>
              )}
            </label>
          )}
        </div>
      )}
    </div>
  );
};
