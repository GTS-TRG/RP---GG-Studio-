import React, { useState } from 'react';
import {
  Search,
  Filter,
  AlertCircle,
  AlertTriangle,
  CheckCircle,
  XCircle,
  HelpCircle,
  CheckCheck,
  RotateCcw,
  MessageSquare
} from 'lucide-react';
import { ReviewIssue } from '../types';

interface ReportTableProps {
  issues: ReviewIssue[];
  panelNames: string[];
  onUpdateIssueStatus: (issueId: string, status: ReviewIssue['status'], remarks: string) => void;
  onBatchUpdateStatus: (status: ReviewIssue['status']) => void;
}

export const ReportTable: React.FC<ReportTableProps> = ({
  issues,
  panelNames,
  onUpdateIssueStatus,
  onBatchUpdateStatus,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedPanel, setSelectedPanel] = useState<string>('ALL');
  const [selectedSeverity, setSelectedSeverity] = useState<'ALL' | 'ERROR' | 'WARNING'>('ALL');
  const [selectedStatus, setSelectedStatus] = useState<string>('ALL');

  const filteredIssues = issues.filter((issue) => {
    if (selectedPanel !== 'ALL' && issue.sheetName !== selectedPanel) return false;
    if (selectedSeverity === 'ERROR' && issue.isWarning) return false;
    if (selectedSeverity === 'WARNING' && !issue.isWarning) return false;
    if (selectedStatus !== 'ALL' && issue.status !== selectedStatus) return false;

    if (searchQuery.trim() !== '') {
      const q = searchQuery.toLowerCase();
      const matchName = issue.lineName.toLowerCase().includes(q);
      const matchDesc = issue.description.toLowerCase().includes(q);
      const matchSheet = issue.sheetName.toLowerCase().includes(q);
      const matchRemarks = issue.remarks.toLowerCase().includes(q);
      return matchName || matchDesc || matchSheet || matchRemarks;
    }

    return true;
  });

  return (
    <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm space-y-0">
      {/* Table Header Controls */}
      <div className="p-4 bg-[#F5F8FB] border-b border-[#D5DEE8] flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        <div className="flex items-center space-x-2">
          <div className="p-2 bg-[#E6F4EC] text-[#1B7A45] rounded-xl border border-[#A8D4B8]">
            <Filter className="w-4 h-4" />
          </div>
          <div>
            <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider">
              Bảng Báo Cáo Thẩm Tra Hồ Sơ (Audit Trail)
            </h2>
            <p className="text-xs text-[#5A6A7A]">
              Hiển thị {filteredIssues.length} / {issues.length} vấn đề phát hiện
            </p>
          </div>
        </div>

        {/* Filters */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Search Input */}
          <div className="relative min-w-[180px] sm:min-w-[220px]">
            <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-[#5A6A7A]" />
            <input
              type="text"
              placeholder="Tìm theo mã mạch, mô tả..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] placeholder-[#5A6A7A] text-xs rounded-full pl-8 pr-3 py-1.5 focus:outline-none focus:border-[#1B7A45]"
            />
          </div>

          {/* Panel Selector */}
          <select
            value={selectedPanel}
            onChange={(e) => setSelectedPanel(e.target.value)}
            className="bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] text-xs rounded-full px-3 py-1.5 focus:outline-none focus:border-[#1B7A45]"
          >
            <option value="ALL">Tất cả Tủ điện ({panelNames.length})</option>
            {panelNames.map((p) => (
              <option key={p} value={p}>
                Tủ {p}
              </option>
            ))}
          </select>

          {/* Severity Selector */}
          <select
            value={selectedSeverity}
            onChange={(e) => setSelectedSeverity(e.target.value as any)}
            className="bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] text-xs rounded-full px-3 py-1.5 focus:outline-none focus:border-[#1B7A45]"
          >
            <option value="ALL">Tất cả Mức độ</option>
            <option value="ERROR">Chỉ Lỗi [ERROR]</option>
            <option value="WARNING">Chỉ Cảnh báo [WARNING]</option>
          </select>

          {/* Status Selector */}
          <select
            value={selectedStatus}
            onChange={(e) => setSelectedStatus(e.target.value)}
            className="bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] text-xs rounded-full px-3 py-1.5 focus:outline-none focus:border-[#1B7A45]"
          >
            <option value="ALL">Tất cả Trạng thái</option>
            <option value="UNRESOLVED">Chưa xử lý (Unresolved)</option>
            <option value="OK">Đã duyệt (OK)</option>
            <option value="IGNORE">Bỏ qua (Ignore)</option>
            <option value="APPROVED">Đã chấp thuận (Approved)</option>
          </select>

          {/* Batch Actions */}
          <div className="flex items-center space-x-1 pl-1 border-l border-[#C5D0DC]">
            <button
              onClick={() => onBatchUpdateStatus('OK')}
              title="Đánh dấu tất cả hiển thị là OK"
              className="p-1.5 bg-[#E6F4EC] hover:bg-[#384C27] text-[#1B7A45] border border-[#A8D4B8] rounded-full text-xs transition-colors flex items-center space-x-1"
            >
              <CheckCheck className="w-3.5 h-3.5" />
              <span className="hidden xl:inline text-[11px] font-semibold">Duyệt tất cả OK</span>
            </button>

            <button
              onClick={() => onBatchUpdateStatus('UNRESOLVED')}
              title="Đặt lại tất cả hiển thị thành Chưa xử lý"
              className="p-1.5 bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] border border-[#C5D0DC] rounded-full text-xs transition-colors"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#2D8A55]" />
            </button>
          </div>
        </div>
      </div>

      {/* Issues Table */}
      {filteredIssues.length === 0 ? (
        <div className="p-12 text-center text-[#5A6A7A]">
          <CheckCircle className="w-12 h-12 text-[#1B7A45] mx-auto mb-3 opacity-90" />
          <h3 className="text-base font-bold text-[#1A2332]">Không tìm thấy vấn đề vi phạm nào!</h3>
          <p className="text-xs text-[#5A6A7A] max-w-sm mx-auto mt-1">
            Không có lỗi hoặc cảnh báo phù hợp với bộ lọc hiện tại. Bảng tính đạt yêu cầu thẩm tra.
          </p>
        </div>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs text-[#1A2332] border-collapse">
            <thead className="bg-[#E8EEF4] text-[#1B7A45] font-bold uppercase border-b border-[#D5DEE8] text-[11px]">
              <tr>
                <th className="py-3 px-3 w-12 text-center font-bold">STT</th>
                <th className="py-3 px-3 min-w-[160px] w-48 font-bold whitespace-nowrap">Tủ Điện</th>
                <th className="py-3 px-3 min-w-[200px] font-bold">Mạch & Mô Tả</th>
                <th className="py-3 px-4 min-w-[380px] font-bold">Nội Dung Kiểm Tra Vi Phạm / Chi Tiết</th>
                <th className="py-3 px-3 w-36 min-w-[150px] text-center font-bold">Trạng Thái Thẩm Tra</th>
                <th className="py-3 px-3 min-w-[260px] w-64 font-bold">Ghi Chú Kỹ Sư (Remarks)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#D5DEE8]">
              {filteredIssues.map((issue, idx) => {
                const isResolved = issue.status !== 'UNRESOLVED';

                return (
                  <tr
                    key={issue.id}
                    className={`transition-colors ${
                      isResolved
                        ? 'bg-[#F0F4F8] text-[#6B7C8C] hover:bg-[#E8EEF4]'
                        : issue.isWarning
                        ? 'bg-[#FFFBEB] hover:bg-[#FEF3C7]'
                        : 'bg-[#FEF2F2] hover:bg-[#FEE2E2]'
                    }`}
                  >
                    {/* Index */}
                    <td className="py-3 px-3 text-center text-[#5A6A7A] font-mono font-bold">
                      {idx + 1}
                    </td>

                    {/* Sheet Name */}
                    <td className="py-3 px-3 font-bold text-[#1A2332] min-w-[160px] max-w-[220px]">
                      <span className="inline-block px-2.5 py-1 rounded-md bg-[#F0F4F8] border border-[#C5D0DC] text-[#1B7A45] font-mono text-[12px] font-bold break-all whitespace-normal leading-snug">
                        {issue.sheetName}
                      </span>
                    </td>

                    {/* Line Name & Short Desc */}
                    <td className="py-3 px-3">
                      <div className="font-bold text-[#1A2332] font-mono">{issue.lineName}</div>
                      {issue.shortDesc && (
                        <div className="text-[11px] text-[#5A6A7A] italic break-words whitespace-normal" title={issue.shortDesc}>
                          ({issue.shortDesc})
                        </div>
                      )}
                      <div className="text-[10px] text-[#6B7C8C]">Dòng Excel: #{issue.rowIndex}</div>
                    </td>

                    {/* Issue Description */}
                    <td className="py-3 px-4 min-w-[280px] break-words whitespace-normal">
                      <div className="flex items-start space-x-2">
                        {issue.isWarning ? (
                          <AlertTriangle className="w-4 h-4 text-[#B45309] shrink-0 mt-0.5" />
                        ) : (
                          <AlertCircle className="w-4 h-4 text-[#DC2626] shrink-0 mt-0.5" />
                        )}
                        <div className="break-words whitespace-normal">
                          <span
                            className={`font-bold text-xs ${
                              issue.isWarning ? 'text-[#B45309]' : 'text-[#DC2626]'
                            }`}
                          >
                            {issue.isWarning ? '[WARNING]' : '[ERROR]'}
                          </span>{' '}
                          <span className="text-[#1A2332] leading-relaxed break-words whitespace-normal font-medium">{issue.description}</span>
                        </div>
                      </div>
                    </td>

                    {/* Status Selector */}
                    <td className="py-3 px-3 text-center min-w-[140px]">
                      <select
                        value={issue.status}
                        onChange={(e) =>
                          onUpdateIssueStatus(issue.id, e.target.value as ReviewIssue['status'], issue.remarks)
                        }
                        className={`text-xs font-bold px-2.5 py-1.5 rounded-full border focus:outline-none transition-colors w-full ${
                          issue.status === 'OK'
                            ? 'bg-[#E6F4EC] text-[#1B7A45] border-[#A8D4B8]'
                            : issue.status === 'APPROVED'
                            ? 'bg-[#1E3040] text-[#60A5FA] border-[#2563EB]'
                            : issue.status === 'IGNORE'
                            ? 'bg-[#F0F4F8] text-[#5A6A7A] border-[#C5D0DC]'
                            : issue.isWarning
                            ? 'bg-[#FEF3C7] text-[#B45309] border-[#D97706]'
                            : 'bg-[#FEE2E2] text-[#DC2626] border-[#FECACA]'
                        }`}
                      >
                        <option value="UNRESOLVED">🔴 Chưa xử lý</option>
                        <option value="OK">🟢 OK (Đã đạt)</option>
                        <option value="APPROVED">🔵 Chấp thuận (Approved)</option>
                        <option value="IGNORE">⚪ Bỏ qua (Ignore)</option>
                      </select>
                    </td>

                    {/* Ô ghi chú: textarea tự xuống dòng + giãn cao theo nội dung */}
                    <td className="py-3 px-3 min-w-[240px] align-top">
                      <div className="relative">
                        <MessageSquare className="w-3 h-3 absolute left-2.5 top-2.5 text-[#5A6A7A] pointer-events-none z-10" />
                        <textarea
                          placeholder="Nhập ghi chú..."
                          value={issue.remarks}
                          rows={2}
                          onChange={(e) => {
                            const el = e.target;
                            el.style.height = 'auto';
                            el.style.height = `${Math.max(el.scrollHeight, 40)}px`;
                            onUpdateIssueStatus(issue.id, issue.status, el.value);
                          }}
                          ref={(el) => {
                            if (el) {
                              el.style.height = 'auto';
                              el.style.height = `${Math.max(el.scrollHeight, 40)}px`;
                            }
                          }}
                          className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] text-xs rounded-lg pl-7 pr-2 py-1.5 focus:outline-none focus:border-[#1B7A45] placeholder-[#5A6A7A] resize-y min-h-[40px] leading-relaxed break-words whitespace-pre-wrap overflow-hidden"
                        />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};
