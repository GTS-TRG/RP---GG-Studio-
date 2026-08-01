import React from 'react';
import { ShieldCheck, Settings, BookOpen, Download, FileSpreadsheet, Save } from 'lucide-react';

interface HeaderProps {
  onOpenConfig: () => void;
  onOpenRules: () => void;
  onDownloadSample: () => void;
  onExportReport: () => void;
  onSaveToExcel?: () => void;
  hasIssues: boolean;
  fileName?: string;
  pendingPatchCount?: number;
  canSaveOriginal?: boolean;
  isSaving?: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  onOpenConfig,
  onOpenRules,
  onDownloadSample,
  onExportReport,
  onSaveToExcel,
  hasIssues,
  fileName,
  pendingPatchCount = 0,
  canSaveOriginal = false,
  isSaving = false,
}) => {
  return (
    <header className="bg-[#FFFFFF]/95 backdrop-blur-sm border-b border-[#D5DEE8] text-[#1A2332] sticky top-0 z-30 shadow-sm">
      <div className="w-full px-3 sm:px-4 lg:px-4 h-16 flex items-center justify-between">
        {/* Brand */}
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 bg-[#2F6F4E] rounded-xl flex items-center justify-center text-white shadow-xs">
            <ShieldCheck className="w-5 h-5 text-white" />
          </div>
          <div>
            <h1 className="text-base sm:text-lg font-bold tracking-tight text-[#1A2332] flex items-center gap-2">
              <span>PanelReview</span>
              <span className="text-xs font-bold text-[#1B7A45] uppercase tracking-widest bg-[#E6F4EC] border border-[#A8D4B8] px-2 py-0.5 rounded-full">
                Analyzer
              </span>
            </h1>
            <p className="text-xs text-[#5A6A7A] hidden sm:block">
              Kiểm tra Bảng tính tải điện Excel theo Tiêu chuẩn Kỹ thuật Điện
            </p>
          </div>
        </div>

        {/* Actions */}
        <div className="flex items-center space-x-2 sm:space-x-2.5">
          <button
            onClick={onDownloadSample}
            title="Tải file Excel mẫu (.xlsx)"
            className="flex items-center space-x-1.5 text-xs font-medium bg-[#E8EEF4] text-[#1A2332] hover:bg-[#334155] border border-[#C5D0DC] px-3 py-1.5 rounded-full transition-colors"
          >
            <Download className="w-3.5 h-3.5 text-[#2D8A55]" />
            <span className="hidden lg:inline">Tải Mẫu Excel</span>
          </button>

          <button
            onClick={onOpenRules}
            title="Xem 13 quy chuẩn kiểm tra"
            className="flex items-center space-x-1.5 text-xs font-medium bg-[#E8EEF4] text-[#1A2332] hover:bg-[#334155] border border-[#C5D0DC] px-3 py-1.5 rounded-full transition-colors"
          >
            <BookOpen className="w-3.5 h-3.5 text-[#2D8A55]" />
            <span className="hidden sm:inline">Tiêu chuẩn VBA</span>
          </button>

          <button
            onClick={onOpenConfig}
            title="Cấu hình thông số dự án"
            className="flex items-center space-x-1.5 text-xs font-medium bg-[#E8EEF4] text-[#1A2332] hover:bg-[#334155] border border-[#C5D0DC] px-3 py-1.5 rounded-full transition-colors"
          >
            <Settings className="w-3.5 h-3.5 text-[#2D8A55]" />
            <span className="hidden sm:inline">Thông số</span>
          </button>

          {onSaveToExcel && pendingPatchCount > 0 && (
            <button
              onClick={onSaveToExcel}
              disabled={isSaving}
              title="Save changes as a new Excel copy (original file stays safe)"
              className="flex items-center space-x-1.5 text-xs font-semibold bg-[#0369A1] hover:bg-[#075985] disabled:opacity-60 text-white px-3.5 py-1.5 rounded-full shadow-xs transition-all"
            >
              <Save className={`w-3.5 h-3.5 ${isSaving ? 'animate-pulse' : ''}`} />
              <span>
                {isSaving ? 'Saving…' : `Save as copy (${pendingPatchCount})`}
              </span>
            </button>
          )}

          {hasIssues && (
            <button
              onClick={onExportReport}
              className="flex items-center space-x-1.5 text-xs font-semibold bg-[#2F6F4E] hover:bg-[#255A3F] text-white px-3.5 py-1.5 rounded-full shadow-xs transition-all"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Xuất Excel Report</span>
            </button>
          )}
        </div>
      </div>
    </header>
  );
};
