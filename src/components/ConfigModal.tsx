import React, { useEffect, useState } from 'react';
import { X, Settings, RotateCcw, Check } from 'lucide-react';
import { ProjectConfig } from '../types';
import { DEFAULT_CONFIG } from '../utils/panelReviewer';

interface ConfigModalProps {
  isOpen: boolean;
  onClose: () => void;
  config: ProjectConfig;
  onSaveConfig: (newConfig: ProjectConfig) => void;
}

export const ConfigModal: React.FC<ConfigModalProps> = ({
  isOpen,
  onClose,
  config,
  onSaveConfig,
}) => {
  const [formConfig, setFormConfig] = useState<ProjectConfig>({ ...config });

  // Modal luon mounted (chi return null khi dong) nen phai dong bo lai moi lan mo,
  // neu khong se hien gia tri cu khi config bi doi tu noi khac (vd: ti le lap day o tab Specs)
  useEffect(() => {
    if (isOpen) setFormConfig({ ...config });
  }, [isOpen, config]);

  if (!isOpen) return null;

  const handleChange = (field: keyof ProjectConfig, val: string) => {
    setFormConfig({
      ...formConfig,
      [field]: Number(val),
    });
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSaveConfig(formConfig);
    onClose();
  };

  const handleReset = () => {
    setFormConfig({ ...DEFAULT_CONFIG });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-xs animate-fade-in">
      <div className="bg-[#FFFFFF] border border-[#D5DEE8] rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden">
        {/* Header */}
        <div className="p-4 bg-[#F5F8FB] border-b border-[#D5DEE8] flex items-center justify-between">
          <div className="flex items-center space-x-2">
            <Settings className="w-5 h-5 text-[#1B7A45]" />
            <h3 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider">
              Cấu Hình Thông Số Thẩm Tra Hồ Sơ
            </h3>
          </div>
          <button
            onClick={onClose}
            className="p-1 text-[#5A6A7A] hover:text-[#1A2332] hover:bg-[#F0F4F8] rounded-full transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Form Body */}
        <form onSubmit={handleSubmit} className="p-5 space-y-4 text-xs text-[#1A2332]">
          <div className="bg-[#F5F8FB] p-3.5 rounded-2xl border border-[#D5DEE8] space-y-3">
            <h4 className="font-bold text-[#1B7A45] uppercase tracking-wide text-[11px]">
              Quy chuẩn Tủ Điện Tổng (MSB) - Rule 13 — chỉ MSB
            </h4>
            <p className="text-[10px] text-[#6B7C8C] leading-snug">
              Isc ≥ ngưỡng dưới chỉ kiểm trên tủ MSB. Tủ DB / LP / EM không bị yêu cầu 65kA.
            </p>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">MSB — Isc tối thiểu (kA):</label>
                <input
                  type="number"
                  value={formConfig.minIscMsb}
                  onChange={(e) => handleChange('minIscMsb', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">MSB — In tối thiểu (A):</label>
                <input
                  type="number"
                  value={formConfig.minAmpMsb}
                  onChange={(e) => handleChange('minAmpMsb', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
            </div>
          </div>

          <div className="bg-[#F5F8FB] p-3.5 rounded-2xl border border-[#D5DEE8] space-y-3">
            <h4 className="font-bold text-[#1B7A45] uppercase tracking-wide text-[11px]">
              Ngưỡng Isc tủ phân phối (không phải MSB)
            </h4>
            <p className="text-[10px] text-[#6B7C8C] leading-snug">
              Không dùng ngưỡng 65kA. Chỉ giới hạn MCB tối đa / MCCB tối thiểu cho tủ nhánh.
            </p>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">MCB Isc tối đa cho phép (kA):</label>
                <input
                  type="number"
                  value={formConfig.maxMcbIsc}
                  onChange={(e) => handleChange('maxMcbIsc', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">MCCB Isc tối thiểu yêu cầu (kA):</label>
                <input
                  type="number"
                  value={formConfig.minMccbIsc}
                  onChange={(e) => handleChange('minMccbIsc', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
            </div>
          </div>

          <div className="bg-[#F5F8FB] p-3.5 rounded-2xl border border-[#D5DEE8] space-y-3">
            <h4 className="font-bold text-[#1B7A45] uppercase tracking-wide text-[11px]">
              Hệ Số An Toàn Chọn CB & Cấu Trúc File
            </h4>
            <div className="grid grid-cols-3 gap-3">
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">Tải thường:</label>
                <input
                  type="number"
                  step="0.05"
                  value={formConfig.safetyFactorNormal}
                  onChange={(e) => handleChange('safetyFactorNormal', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">Tải PCCC:</label>
                <input
                  type="number"
                  step="0.05"
                  value={formConfig.safetyFactorFire}
                  onChange={(e) => handleChange('safetyFactorFire', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">Dòng bắt đầu (Row):</label>
                <input
                  type="number"
                  value={formConfig.startRow}
                  onChange={(e) => handleChange('startRow', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
              <div>
                <label className="block text-[#5A6A7A] mb-1 font-medium">
                  Tỉ lệ lấp đầy ống tối đa (%):
                </label>
                <input
                  type="number"
                  min={1}
                  max={100}
                  value={formConfig.conduitFillPercent}
                  onChange={(e) => handleChange('conduitFillPercent', e.target.value)}
                  className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
                />
              </div>
            </div>
          </div>

          {/* Footer buttons */}
          <div className="pt-3 flex items-center justify-between border-t border-[#D5DEE8]">
            <button
              type="button"
              onClick={handleReset}
              className="px-3.5 py-1.5 bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] rounded-full border border-[#C5D0DC] transition-colors flex items-center space-x-1"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#2D8A55]" />
              <span>Khôi phục mặc định</span>
            </button>

            <div className="flex items-center space-x-2">
              <button
                type="button"
                onClick={onClose}
                className="px-3.5 py-1.5 bg-[#F5F8FB] hover:bg-[#F0F4F8] text-[#1A2332] rounded-full transition-colors border border-[#C5D0DC]"
              >
                Hủy
              </button>

              <button
                type="submit"
                className="px-4 py-1.5 bg-[#2F6F4E] hover:bg-[#255A3F] text-white font-bold rounded-full transition-colors flex items-center space-x-1 shadow-xs"
              >
                <Check className="w-4 h-4" />
                <span>Lưu & Áp Dụng</span>
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
};
