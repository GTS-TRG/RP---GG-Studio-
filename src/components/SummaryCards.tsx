import React from 'react';
import { AlertTriangle, AlertCircle, CheckCircle2, Layers, Cpu, ShieldAlert } from 'lucide-react';
import { ReviewSummary } from '../types';

interface SummaryCardsProps {
  summary: ReviewSummary;
  /** sidebar = xếp dọc trong task panel trái */
  variant?: 'grid' | 'sidebar';
}

export const SummaryCards: React.FC<SummaryCardsProps> = ({ summary, variant = 'sidebar' }) => {
  const totalIssues = summary.errorCount + summary.warningCount;
  const unresolvedIssues = totalIssues - summary.resolvedCount;
  const complianceRate = summary.totalCircuits > 0
    ? Math.max(0, Math.round(((summary.totalCircuits - summary.errorCount) / summary.totalCircuits) * 100))
    : 100;

  const cards = [
    {
      key: 'panels',
      label: 'Tủ điện đã quét',
      value: summary.totalPanels,
      unit: 'Sheet',
      icon: Layers,
      wrap: 'bg-[#FFFFFF] border-[#D5DEE8]',
      labelCls: 'text-[#5A6A7A]',
      valueCls: 'text-[#1A2332]',
      unitCls: 'text-[#5A6A7A]',
      iconCls: 'text-[#2D8A55]',
    },
    {
      key: 'circuits',
      label: 'Tổng số mạch',
      value: summary.totalCircuits,
      unit: 'mạch',
      icon: Cpu,
      wrap: 'bg-[#FFFFFF] border-[#D5DEE8]',
      labelCls: 'text-[#5A6A7A]',
      valueCls: 'text-[#1A2332]',
      unitCls: 'text-[#5A6A7A]',
      iconCls: 'text-[#2D8A55]',
    },
    {
      key: 'errors',
      label: 'Lỗi nghiêm trọng',
      value: summary.errorCount,
      unit: 'Lỗi [ERROR]',
      icon: AlertCircle,
      wrap: 'bg-[#FEF2F2] border-[#FECACA]',
      labelCls: 'text-[#DC2626]',
      valueCls: 'text-[#B91C1C]',
      unitCls: 'text-[#DC2626] font-bold',
      iconCls: 'text-[#DC2626]',
    },
    {
      key: 'warnings',
      label: 'Cảnh báo kỹ thuật',
      value: summary.warningCount,
      unit: '[WARNING]',
      icon: AlertTriangle,
      wrap: 'bg-[#FFFBEB] border-[#FCD34D]',
      labelCls: 'text-[#B45309]',
      valueCls: 'text-[#B45309]',
      unitCls: 'text-[#B45309] font-bold',
      iconCls: 'text-[#B45309]',
    },
    {
      key: 'resolved',
      label: 'Đã xử lý / Duyệt',
      value: summary.resolvedCount,
      unit: totalIssues > 0 ? `(${Math.round((summary.resolvedCount / totalIssues) * 100)}%)` : '100%',
      icon: CheckCircle2,
      wrap: 'bg-[#ECFDF5] border-[#A7F3D0]',
      labelCls: 'text-[#1B7A45]',
      valueCls: 'text-[#1A2332]',
      unitCls: 'text-[#1B7A45] font-semibold',
      iconCls: 'text-[#1B7A45]',
    },
  ];

  const renderCompliance = (extraCls = '') => (
    <div
      className={`bg-[#FFFFFF] border border-[#D5DEE8] rounded-2xl p-3.5 flex flex-col justify-center gap-2 text-xs ${extraCls}`}
    >
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center space-x-2 min-w-0">
          <ShieldAlert className="w-4 h-4 text-[#2D8A55] shrink-0" />
          <span className="font-semibold text-[#1A2332] leading-snug">Tuân thủ quy chuẩn</span>
        </div>
        <span className="font-bold text-[#1B7A45] text-sm shrink-0">{complianceRate}%</span>
      </div>

      <div className="w-full bg-[#F0F4F8] rounded-full h-2 overflow-hidden border border-[#C5D0DC]">
        <div
          className={`h-full transition-all duration-500 rounded-full ${
            complianceRate >= 90
              ? 'bg-[#1B7A45]'
              : complianceRate >= 70
              ? 'bg-[#B45309]'
              : 'bg-[#DC2626]'
          }`}
          style={{ width: `${complianceRate}%` }}
        />
      </div>

      <div className="text-[#5A6A7A] text-[11px] leading-relaxed">
        {unresolvedIssues === 0 ? (
          <span className="text-[#1B7A45] font-semibold">✓ Tất cả vấn đề đã được xử lý.</span>
        ) : (
          <span>
            Còn <strong className="text-[#DC2626]">{unresolvedIssues}</strong> vấn đề cần rà soát.
          </span>
        )}
      </div>
    </div>
  );

  if (variant === 'sidebar') {
    return (
      <div className="space-y-2.5">
        {cards.map((c) => {
          const Icon = c.icon;
          return (
            <div
              key={c.key}
              className={`border rounded-2xl px-3.5 py-3 flex items-center justify-between gap-2 shadow-xs ${c.wrap}`}
            >
              <div className="min-w-0">
                <div className={`flex items-center gap-1.5 text-[10px] font-bold uppercase tracking-wider mb-1 ${c.labelCls}`}>
                  <Icon className={`w-3.5 h-3.5 shrink-0 ${c.iconCls}`} />
                  <span className="truncate">{c.label}</span>
                </div>
                <div className="flex items-baseline gap-1.5">
                  <span className={`text-2xl font-bold tabular-nums leading-none ${c.valueCls}`}>{c.value}</span>
                  <span className={`text-[10px] ${c.unitCls}`}>{c.unit}</span>
                </div>
              </div>
            </div>
          );
        })}
        {renderCompliance()}
      </div>
    );
  }

  return (
    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-7 gap-2.5 sm:gap-3">
      {cards.map((c) => {
        const Icon = c.icon;
        return (
          <div
            key={c.key}
            className={`border rounded-2xl px-3.5 py-3 flex flex-col justify-between shadow-xs ${c.wrap}`}
          >
            <div className={`flex items-center justify-between gap-2 mb-1.5 ${c.labelCls}`}>
              <span className="text-[10px] font-bold uppercase tracking-wider truncate">{c.label}</span>
              <Icon className={`w-3.5 h-3.5 shrink-0 ${c.iconCls}`} />
            </div>
            <div className="flex items-baseline justify-between gap-1.5">
              <span className={`text-2xl font-bold tabular-nums leading-none ${c.valueCls}`}>{c.value}</span>
              <span className={`text-[10px] ${c.unitCls}`}>{c.unit}</span>
            </div>
          </div>
        );
      })}
      {renderCompliance('col-span-2')}
    </div>
  );
};
