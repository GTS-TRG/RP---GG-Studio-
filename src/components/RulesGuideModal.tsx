import React from 'react';
import { X, BookOpen, CheckCircle, ShieldAlert, Cpu, Zap } from 'lucide-react';

interface RulesGuideModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const RulesGuideModal: React.FC<RulesGuideModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  const rulesList = [
    {
      num: '13',
      title: 'Dòng Cắt & Định Mức Tủ Tổng MSB (Rule 13)',
      desc: 'CHỈ áp dụng cho CB trên tủ tổng MSB: Isc ≥ 65kA và In ≥ 32A. Tủ khác (DB, LP, EM, …) KHÔNG bắt buộc Isc ≥ 65kA.',
      badge: 'CHỈ MSB',
      color: 'border-rose-500/50 bg-rose-950/20 text-rose-300',
    },
    {
      num: '01',
      title: 'Giới Hạn Dòng Cắt Isc Cho Tủ Điện Phân Phối',
      desc: 'Áp dụng cho tủ không phải MSB: MCB Isc tối đa ≤ 15kA; MCCB Isc tối thiểu ≥ 18kA. Không dùng ngưỡng 65kA cho các tủ này.',
      badge: 'KHÔNG PHẢI MSB',
      color: 'border-amber-500/50 bg-amber-950/20 text-amber-300',
    },
    {
      num: '02',
      title: 'Mạch Dự Phòng (Spare) & Mạch Điều Khiển (CT)',
      desc: 'Mạch dự phòng (Spare / Prefix SP) và Mạch điều khiển (Prefix CT / Mạch điều khiển) KHÔNG yêu cầu khai báo thông số cáp pha và cáp PE. Mạch dự phòng kiểm tra bắt buộc thông số CB dưới dạng Cảnh Báo [WARNING].',
      badge: 'CẬP NHẬT: SPARE & CT',
      color: 'border-blue-500/50 bg-blue-950/20 text-blue-300',
    },
    {
      num: '03',
      title: 'Số Cực CB Theo Số Pha & Mạch Ổ Cắm (Prefix S)',
      desc: 'Mạch 3 pha (phụ tải trên ≥ 2 pha) phải chọn CB 3P/4P. Mạch 1 pha phải dùng 1P, 1P+N, 2P. Riêng mạch Ổ cắm (tiền tố S) bắt buộc dùng 1P+N bảo vệ dây nguội.',
      badge: 'SỐ CỰC POLE',
      color: 'border-emerald-500/50 bg-emerald-950/20 text-emerald-300',
    },
    {
      num: '04',
      title: 'Dòng Định Mức CB (In) vs Tải Tính Toán',
      desc: 'Tải thường tính an toàn K = 1.25, tải PCCC (Bơm chữa cháy, hút khói, tạo áp) tính K = 1.50. So sánh In được chọn với chuẩn trong Bảng Cáp Spec. Cable.',
      badge: 'TẢI TÍNH TOÁN',
      color: 'border-slate-700 bg-slate-900/60 text-slate-200',
    },
    {
      num: '05',
      title: 'Cảnh Báo CB Quá Cỡ (Oversized CB)',
      desc: 'Nếu dòng định mức In của CB được chọn vượt quá 1 cấp so với cấp quy chuẩn cần thiết từ tải tính toán, hệ thống sẽ đưa ra Cảnh báo chọn CB quá cỡ.',
      badge: 'OVERSIZED',
      color: 'border-slate-700 bg-slate-900/60 text-slate-200',
    },
    {
      num: '06',
      title: 'Kiểm Tra Tiết Diện Cáp Theo Rating CB (Cable Cross-Section)',
      desc: 'Chỉ yêu cầu kiểm tra tiết diện cáp (mm²) tương ứng với dòng định mức rating của CB. Không yêu cầu bắt buộc quy định về chủng loại cáp/vỏ cách điện (Cu/PVC, Cu/XLPE/PVC, ...).',
      badge: 'TIẾT DIỆN CÁP',
      color: 'border-purple-500/50 bg-purple-950/20 text-purple-300',
    },
    {
      num: '07',
      title: 'Kiểm Tra Tiết Diện & Số Lượng Dây Pha (Phase Cable)',
      desc: 'Đối chiếu tiết diện dây pha với bảng Spec. Cable theo CB In và kiểm tra số lượng sợi/lõi cáp: Mạch 3 pha yêu cầu tối thiểu 3 dây pha (3x1C hoặc cáp đa lõi 3C/4C/5C); Động cơ Sao/Tam giác (S/D) cần 6 dây pha (2x(3x1C) hoặc 2x3C); Mạch 1 pha cần tối thiểu 2 dây (2x1C hoặc cáp 2C/3C).',
      badge: 'CẬP NHẬT: DÂY PHA & SỐ CÁP',
      color: 'border-[#1B7A45] bg-[#E6F4EC] text-[#1B7A45]',
    },
    {
      num: '08',
      title: 'Kiểm Tra Dây Nối Đất PE (PE Cable)',
      desc: 'Đối chiếu tiết diện dây PE với bảng Spec. Cable. Ngoại lệ: Quạt chạy tốc độ thấp (Low speed fan) hoặc cáp đã có lõi PE tích hợp (3C cho 1 pha, 5C cho 3 pha).',
      badge: 'DÂY PE',
      color: 'border-slate-700 bg-slate-900/60 text-slate-200',
    },
    {
      num: '09',
      title: 'Sai Lệch Liên Kết Công Thức (Cross-Link Mismatch)',
      desc: 'Cảnh báo nếu công thức ô Công Suất Tải liên kết từ một sheet tủ điện khác nhưng ô chọn CB lại liên kết từ một sheet tủ điện khác.',
      badge: 'LIÊN KẾT',
      color: 'border-amber-500/50 bg-amber-950/20 text-amber-300',
    },
    {
      num: '10',
      title: 'Bỏ Qua Kiểm Tra Đối Với Mạch Cấp Nguồn (CP) & Comment OK',
      desc: 'Mạch tiền tố CP (Cấp nguồn) hoặc các ô có Comment / Cột 26 chứa từ khóa "OK" / "IGNORE" sẽ được tự động bỏ qua không báo lỗi.',
      badge: 'BYPASS',
      color: 'border-slate-700 bg-slate-900/60 text-slate-200',
    },
    {
      num: '11',
      title: 'Tự Động Trích Xuất Mô Tả Ngắn (Short Description)',
      desc: 'Tự động làm sạch mô tả dài (bỏ các từ CẤP NGUỒN CHO, TỦ ĐIỆN, ...) để ghép thành tên mạch gọn đẹp: LineName (ShortDesc) trong báo cáo.',
      badge: 'MÔ TẢ NGẮN',
      color: 'border-slate-700 bg-slate-900/60 text-slate-200',
    },
    {
      num: '12',
      title: 'Bảo Toàn Lịch Sử Kiểm Tra (Audit Trail)',
      desc: 'Khi kiểm tra lại file Excel hoặc cập nhật file, các trạng thái đã duyệt (OK, IGNORE, APPROVED) và Ghi chú kỹ sư (Remarks) sẽ được lưu giữ nguyên vẹn.',
      badge: 'AUDIT TRAIL',
      color: 'border-emerald-500/50 bg-emerald-950/20 text-emerald-300',
    },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-xs animate-fade-in">
      <div className="bg-[#FFFFFF] border border-[#D5DEE8] rounded-3xl shadow-2xl max-w-2xl w-full max-h-[85vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="p-4 bg-[#F5F8FB] border-b border-[#D5DEE8] flex items-center justify-between shrink-0">
          <div className="flex items-center space-x-2">
            <BookOpen className="w-5 h-5 text-[#1B7A45]" />
            <h3 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider">
              Hệ Thống 13 Quy Chuẩn Kiểm Tra Bảng Tính Tải Điện
            </h3>
          </div>
          <button
            onClick={onClose}
            className="p-1 text-[#5A6A7A] hover:text-[#1A2332] hover:bg-[#F0F4F8] rounded-full transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* List Body */}
        <div className="p-5 overflow-y-auto space-y-3 divide-y divide-[#D5DEE8] text-xs">
          {rulesList.map((rule) => (
            <div key={rule.num} className="pt-3 first:pt-0 flex items-start space-x-3">
              <div className="w-7 h-7 rounded-xl bg-[#E6F4EC] border border-[#A8D4B8] text-[#1B7A45] font-bold font-mono flex items-center justify-center shrink-0 text-xs">
                {rule.num}
              </div>
              <div className="flex-1">
                <div className="flex items-center justify-between gap-2 mb-1">
                  <h4 className="font-bold text-[#1A2332] text-xs">{rule.title}</h4>
                  <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-[#F0F4F8] border border-[#C5D0DC] text-[#1B7A45]">
                    {rule.badge}
                  </span>
                </div>
                <p className="text-[#5A6A7A] leading-relaxed text-[11px]">{rule.desc}</p>
              </div>
            </div>
          ))}
        </div>

        {/* Footer */}
        <div className="p-3 bg-[#F5F8FB] border-t border-[#D5DEE8] text-right shrink-0">
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-[#2F6F4E] hover:bg-[#255A3F] text-white font-bold text-xs rounded-full transition-colors"
          >
            Đã Hiểu & Đóng
          </button>
        </div>
      </div>
    </div>
  );
};
