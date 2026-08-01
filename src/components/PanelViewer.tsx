import React, { useEffect, useMemo, useState } from 'react';
import { PanelSheetData, PanelTotalRow, RawCircuitRow, ReviewIssue } from '../types';
import {
  Cpu,
  AlertCircle,
  AlertTriangle,
  Check,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  LogIn,
} from 'lucide-react';
import { PanelEditableField } from '../utils/excelPatch';
import { CbRatingItem, formatCbRatingLabel } from '../utils/cbRatingTable';
import { IssueField, groupIssuesByField } from '../utils/issueFieldMap';
import {
  SpecListItem,
  formatCbTypeLabel,
  formatIscLabel,
  formatPoleLabel,
  withCurrentOption,
} from '../utils/cbOptionLists';

interface PanelViewerProps {
  panels: PanelSheetData[];
  issues: ReviewIssue[];
  /** Danh sách In từ Spec. Cable — CB_Rating */
  cbRatings?: CbRatingItem[];
  cbTypes?: SpecListItem[];
  poleOptions?: SpecListItem[];
  iscOptions?: SpecListItem[];
  /** Cập nhật field mạch + đăng ký patch Excel */
  onUpdateCircuit?: (
    sheetName: string,
    rowIndex: number,
    field: PanelEditableField,
    value: string
  ) => void;
}

const inputCls =
  'w-full min-w-[4.5rem] bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded px-1.5 py-1 text-[12px] focus:outline-none focus:border-[#1B7A45]';
/** Select Loại CB — đủ hiện RCBO(30mA), không dư khoảng trống */
const cbTypeSelectCls =
  'w-full min-w-[6.5rem] bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded px-1 py-1 text-[12px] focus:outline-none focus:border-[#1B7A45]';

/** Bỏ dấu tiếng Việt + hạ chữ thường, để tìm tủ không cần gõ đúng dấu */
const COMBINING_DIACRITICS_RE = new RegExp(
  '[' + String.fromCharCode(0x0300) + '-' + String.fromCharCode(0x036f) + ']',
  'g'
);

function normalizeSearchText(s: string): string {
  return s
    .normalize('NFD')
    .replace(COMBINING_DIACRITICS_RE, '')
    .replace(/đ/gi, 'd')
    .toLowerCase()
    .trim();
}

/** Vị trí + nội dung tooltip đang hiển thị */
interface TipState {
  issues: ReviewIssue[];
  x: number;
  y: number;
  pinned: boolean;
}

interface BadgeProps {
  list: ReviewIssue[];
  className?: string;
  /** Mở tooltip tại vị trí icon */
  onShow: (el: HTMLElement, list: ReviewIssue[], pinned: boolean) => void;
  onHide: () => void;
  /** Đang có tooltip bị ghim -> hover không đổi nội dung */
  isPinned: boolean;
}

/**
 * Icon cảnh báo đặt ngay trên ô thông số bị lỗi.
 * Định nghĩa ở cấp module — nếu đặt trong component cha thì mỗi lần render
 * React coi là component mới và remount cả ô (mất con trỏ khi đang gõ).
 */
const IssueBadge: React.FC<BadgeProps> = ({
  list,
  className = '',
  onShow,
  onHide,
  isPinned,
}) => {
  if (list.length === 0) return null;
  const hasError = list.some((i) => !i.isWarning);
  const Icon = hasError ? AlertCircle : AlertTriangle;

  return (
    <button
      type="button"
      aria-label={list.map((i) => i.description).join(' | ')}
      onMouseEnter={(e) => {
        if (!isPinned) onShow(e.currentTarget, list, false);
      }}
      onMouseLeave={() => {
        if (!isPinned) onHide();
      }}
      onClick={(e) => {
        e.stopPropagation();
        onShow(e.currentTarget, list, true);
      }}
      className={`inline-flex items-center justify-center rounded-full bg-white shadow-xs cursor-pointer transition-transform hover:scale-110 ${className}`}
    >
      <Icon
        className={`w-4 h-4 ${hasError ? 'text-[#DC2626]' : 'text-[#B45309]'}`}
        strokeWidth={2.5}
      />
      {list.length > 1 && (
        <span
          className={`ml-0.5 text-[10px] font-bold leading-none ${
            hasError ? 'text-[#DC2626]' : 'text-[#B45309]'
          }`}
        >
          {list.length}
        </span>
      )}
    </button>
  );
};

/** Bọc nội dung ô + gắn icon lỗi ở góc trên phải */
const CellWithBadge: React.FC<
  Omit<BadgeProps, 'className'> & { children: React.ReactNode }
> = ({ list, children, onShow, onHide, isPinned }) => (
  <div className="relative isolate">
    {children}
    <IssueBadge
      list={list}
      // z thấp hơn thead sticky — tránh icon tràn đè lên header khi cuộn
      className="absolute -top-2.5 -right-1 z-[1] px-0.5"
      onShow={onShow}
      onHide={onHide}
      isPinned={isPinned}
    />
  </div>
);

/** Hiển thị số lỗi / cảnh báo bằng icon thay vì chữ */
const CountBadges: React.FC<{ err: number; warn: number; compact?: boolean }> = ({
  err,
  warn,
  compact = false,
}) => {
  if (err === 0 && warn === 0) {
    return (
      <span className="inline-flex items-center gap-1 text-[#1B7A45]">
        <Check className="w-3.5 h-3.5" strokeWidth={3} />
        {!compact && <span className="text-[11px] font-semibold">đạt</span>}
      </span>
    );
  }
  return (
    <span className="inline-flex items-center gap-2 shrink-0">
      {err > 0 && (
        <span className="inline-flex items-center gap-0.5 text-[#DC2626]" title={`${err} lỗi`}>
          <AlertCircle className="w-3.5 h-3.5" strokeWidth={2.5} />
          <span className="text-[12px] font-bold tabular-nums">{err}</span>
        </span>
      )}
      {warn > 0 && (
        <span className="inline-flex items-center gap-0.5 text-[#B45309]" title={`${warn} cảnh báo`}>
          <AlertTriangle className="w-3.5 h-3.5" strokeWidth={2.5} />
          <span className="text-[12px] font-bold tabular-nums">{warn}</span>
        </span>
      )}
    </span>
  );
};

/* ── Khối tổng kết chân bảng (tfoot) ──────────────────────────────────────
   Nền xanh đậm hơn thân bảng để tách bạch với dòng mạch; mỗi ô phải tự mang
   nền vì tfoot dùng position:sticky (nền của <tr> không phủ khi dính đáy). */
const footRowCls = 'bg-[#D8EDE1]';
const footCellCls = 'py-1 px-2 bg-[#D8EDE1] align-middle';
const footNumCls = `${footCellCls} px-1.5 text-center font-bold tabular-nums text-[#1A2332] whitespace-nowrap`;

/**
 * Đường kẻ của tfoot phải vẽ bằng box-shadow, KHÔNG dùng border:
 * bảng đang border-collapse nên viền do <table> vẽ, không dính theo ô sticky
 * -> khi cuộn, vạch bị bỏ lại đúng chỗ cũ giữa bảng.
 */
const FOOT_TOP_LINE = '0 -2px 0 0 #1B7A45';
const FOOT_LEFT_LINE = 'inset 2px 0 0 0 #1B7A45';

/** Nhãn chữ của dòng tổng — số nằm riêng ở cột Full để thẳng hàng */
const FootLabel: React.FC<{ text: string }> = ({ text }) => (
  <span className="font-sans text-[12px] font-bold uppercase tracking-wide text-[#1B7A45] whitespace-nowrap">
    {text}
  </span>
);

/**
 * Bốn ô Full/R/Y/B — thẳng cột với công suất của từng mạch phía trên.
 * Trả về Fragment 4 <td> nên phải đặt trực tiếp trong <tr>.
 */
const FootPhaseCells: React.FC<{ total?: PanelTotalRow; shadow?: string }> = ({
  total,
  shadow,
}) => {
  const style = shadow ? { boxShadow: shadow } : undefined;
  const cell = (v?: number) => (v === undefined ? '' : v.toFixed(2));
  return (
    <>
      <td className={`${footNumCls} text-[12px] text-[#1B7A45]`} style={style}>
        {cell(total?.full)}
      </td>
      <td className={footNumCls} style={style}>{cell(total?.r)}</td>
      <td className={footNumCls} style={style}>{cell(total?.y)}</td>
      <td className={footNumCls} style={style}>{cell(total?.b)}</td>
    </>
  );
};

/** Một ô thông số lộ vào, đặt đúng dưới cột tương ứng của bảng */
const FootIncomingCell: React.FC<{
  value?: string;
  align?: 'left' | 'center' | 'right';
  bold?: boolean;
  wrap?: boolean;
  sans?: boolean;
  shadow?: string;
  /** Lỗi thẩm tra riêng của ô này (vd sai tiết diện/số lượng dây pha) — hiện icon giống ô mạch thường */
  badgeList?: ReviewIssue[];
  onShowBadge?: BadgeProps['onShow'];
  onHideBadge?: () => void;
  isPinned?: boolean;
}> = ({
  value,
  align = 'left',
  bold = false,
  wrap = false,
  sans = false,
  shadow,
  badgeList,
  onShowBadge,
  onHideBadge,
  isPinned,
}) => (
  <td
    style={shadow ? { boxShadow: shadow } : undefined}
    // Tailwind quét chuỗi tĩnh — không dùng `text-${align}` vì lớp sẽ không được sinh ra
    className={`${footCellCls} ${
      align === 'center' ? 'text-center' : align === 'right' ? 'text-right' : 'text-left'
    } ${
      wrap ? 'break-words whitespace-normal leading-snug' : 'whitespace-nowrap'
    } ${sans ? 'font-sans font-medium' : ''} ${bold ? 'font-bold text-[#1A2332]' : 'text-[#334155]'}`}
  >
    {badgeList && onShowBadge && onHideBadge ? (
      <div className="relative isolate inline-block max-w-full align-middle">
        {value}
        <IssueBadge
          list={badgeList}
          className="absolute -top-2.5 -right-3 z-[1] px-0.5"
          onShow={onShowBadge}
          onHide={onHideBadge}
          isPinned={!!isPinned}
        />
      </div>
    ) : (
      value
    )}
  </td>
);

export const PanelViewer: React.FC<PanelViewerProps> = ({
  panels,
  issues,
  cbRatings = [],
  cbTypes = [],
  poleOptions = [],
  iscOptions = [],
  onUpdateCircuit,
}) => {
  const [selectedSheet, setSelectedSheet] = useState<string>(panels[0]?.sheetName || '');
  const [tip, setTip] = useState<TipState | null>(null);
  const [onlyIssues, setOnlyIssues] = useState(false);
  /** Vị trí danh sách tủ (dùng fixed để không bị khung card cắt mất) */
  const [listPos, setListPos] = useState<{ x: number; y: number; w: number } | null>(null);
  /** Từ khoá tìm nhanh tên tủ trong danh sách sổ xuống */
  const [panelSearch, setPanelSearch] = useState('');

  // Đóng danh sách -> xoá luôn từ khoá tìm, để lần mở sau bắt đầu sạch
  useEffect(() => {
    if (!listPos) setPanelSearch('');
  }, [listPos]);

  const currentPanel = panels.find((p) => p.sheetName === selectedSheet) || panels[0];

  const typeOpts = useMemo(() => cbTypes.map((r) => r.label), [cbTypes]);
  const poleOpts = useMemo(() => poleOptions.map((r) => r.label), [poleOptions]);
  const iscOpts = useMemo(() => iscOptions.map((r) => r.label), [iscOptions]);
  const inOpts = useMemo(() => cbRatings.map((r) => r.label), [cbRatings]);

  // Danh sách tủ: click ra ngoài hoặc Esc thì đóng
  useEffect(() => {
    if (!listPos) return;
    const close = () => setListPos(null);
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setListPos(null);
    };
    document.addEventListener('click', close);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('click', close);
      document.removeEventListener('keydown', onKey);
    };
  }, [listPos]);

  // Tooltip đang ghim: click ra ngoài / cuộn / Esc thì đóng
  useEffect(() => {
    if (!tip?.pinned) return;
    const close = () => setTip(null);
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setTip(null);
    };
    document.addEventListener('click', close);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('click', close);
      document.removeEventListener('keydown', onKey);
    };
  }, [tip?.pinned]);

  if (!currentPanel) {
    return (
      <div className="bg-[#FFFFFF] border border-[#D5DEE8] rounded-3xl p-8 text-center text-[#5A6A7A]">
        Chưa có dữ liệu tủ điện nào được nạp.
      </div>
    );
  }

  // Lỗi của tủ hiện tại, gom theo dòng
  const panelIssuesMap = new Map<number, ReviewIssue[]>();
  for (const issue of issues) {
    if (issue.sheetName === currentPanel.sheetName) {
      const list = panelIssuesMap.get(issue.rowIndex) || [];
      list.push(issue);
      panelIssuesMap.set(issue.rowIndex, list);
    }
  }

  const edit = (row: RawCircuitRow, field: PanelEditableField, value: string) => {
    onUpdateCircuit?.(currentPanel.sheetName, row.rowIndex, field, value);
  };

  // Đếm lỗi/cảnh báo cho từng tủ — dùng cho nhãn trong danh sách chọn
  const countsByPanel = new Map<string, { err: number; warn: number }>();
  for (const p of panels) countsByPanel.set(p.sheetName, { err: 0, warn: 0 });
  for (const i of issues) {
    const c = countsByPanel.get(i.sheetName);
    if (!c) continue;
    if (i.isWarning) c.warn += 1;
    else c.err += 1;
  }
  // CHỈ đếm lỗi thật — cảnh báo không phải lỗi
  const panelsWithErrors = panels.filter(
    (p) => (countsByPanel.get(p.sheetName)?.err ?? 0) > 0
  ).length;

  /**
   * Danh sách hiển thị trong dropdown.
   * Luôn giữ tủ đang xem trong danh sách, kể cả khi nó không còn lỗi —
   * nếu không, danh sách sẽ trống sau khi người dùng sửa hết lỗi của tủ đó.
   */
  const navList = onlyIssues
    ? panels.filter(
        (p) =>
          (countsByPanel.get(p.sheetName)?.err ?? 0) > 0 ||
          p.sheetName === currentPanel.sheetName
      )
    : panels;

  const navIndex = navList.findIndex((p) => p.sheetName === currentPanel.sheetName);

  /** Danh sách hiển thị trong dropdown sau khi lọc thêm theo ô tìm kiếm (không ảnh hưởng nút </>) */
  const searchedNavList = panelSearch.trim()
    ? navList.filter((p) => normalizeSearchText(p.sheetName).includes(normalizeSearchText(panelSearch)))
    : navList;

  /** Chuyển sang tủ trước/sau trong danh sách đang lọc (quay vòng) */
  const goRelative = (step: number) => {
    if (navList.length < 2) return;
    const from = navIndex >= 0 ? navIndex : 0;
    const next = (from + step + navList.length) % navList.length;
    setSelectedSheet(navList[next].sheetName);
  };

  /** Lộ vào — trải trên 2 dòng của tfoot nên tách ra cho gọn */
  const inc = currentPanel.footer?.incoming;
  // Issue của lộ vào không gắn với dòng mạch nào (panel.circuits) nên phải tra riêng
  // theo rowIndex đã gán khi parse (xem reviewIncomingCable trong panelReviewer.ts).
  const incomingRowIndex = inc?.rowIndex ?? currentPanel.endRow;
  const incomingIssues = inc ? panelIssuesMap.get(incomingRowIndex) || [] : [];
  const incByField = groupIssuesByField(incomingIssues);
  const incAt = (f: IssueField) => incByField.get(f) || [];

  const hideTip = () => setTip(null);

  /** Mở tooltip tại vị trí icon (dùng fixed để không bị khung cuộn cắt mất) */
  const openTip = (el: HTMLElement, list: ReviewIssue[], pinned: boolean) => {
    const r = el.getBoundingClientRect();
    const half = 150;
    const x = Math.min(Math.max(r.left + r.width / 2, half + 8), window.innerWidth - half - 8);
    setTip({ issues: list, x, y: r.bottom + 8, pinned });
  };

  return (
    <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm flex flex-col max-h-[calc(100vh-19rem)] min-h-[320px]">
      {/* Thanh chọn tủ điện — dùng danh sách sổ xuống vì dự án có thể tới hàng chục tủ,
          xếp thành hàng tab sẽ tràn và bóp vỡ bố cục tiêu đề */}
      <div className="shrink-0 z-20 p-3 bg-[#F5F8FB] border-b border-[#D5DEE8] flex flex-col lg:flex-row lg:items-center gap-3">
        <div className="flex items-center gap-2 min-w-0 lg:w-[300px] xl:w-[360px] shrink-0">
          <div className="p-2 bg-[#E6F4EC] text-[#1B7A45] rounded-xl border border-[#A8D4B8] shrink-0">
            <Cpu className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider truncate">
              {currentPanel.sheetName}
            </h2>
            <p className="text-[11px] text-[#5A6A7A] truncate">
              {currentPanel.isMSB ? 'Tủ điện tổng (MSB)' : 'Tủ phân phối/chiếu sáng'} •{' '}
              {currentPanel.circuits.length} mạch
              {onUpdateCircuit ? ' • sửa được' : ''}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap lg:flex-nowrap lg:ml-auto min-w-0">
          <button
            type="button"
            onClick={() => goRelative(-1)}
            disabled={navList.length < 2}
            title="Tủ trước"
            className="p-1.5 rounded-lg border border-[#C5D0DC] bg-white text-[#5A6A7A] hover:bg-[#E8EEF4] disabled:opacity-40 shrink-0"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>

          {/* Nút mở danh sách tủ — dựng riêng vì <select> gốc chỉ nhận text, không hiện được icon */}
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              if (listPos) {
                setListPos(null);
                return;
              }
              const r = e.currentTarget.getBoundingClientRect();
              setListPos({ x: r.left, y: r.bottom + 4, w: Math.max(r.width, 280) });
            }}
            className="min-w-[260px] max-w-[440px] flex-1 bg-white border border-[#C5D0DC] hover:border-[#1B7A45] text-[#1A2332] rounded-lg px-2.5 py-2 text-[13px] font-semibold flex items-center justify-between gap-2 transition-colors"
          >
            <span className="truncate">{currentPanel.sheetName}</span>
            <span className="flex items-center gap-2 shrink-0">
              <CountBadges
                err={countsByPanel.get(currentPanel.sheetName)?.err ?? 0}
                warn={countsByPanel.get(currentPanel.sheetName)?.warn ?? 0}
              />
              <ChevronDown className="w-4 h-4 text-[#5A6A7A]" />
            </span>
          </button>

          <button
            type="button"
            onClick={() => goRelative(1)}
            disabled={navList.length < 2}
            title="Tủ tiếp theo"
            className="p-1.5 rounded-lg border border-[#C5D0DC] bg-white text-[#5A6A7A] hover:bg-[#E8EEF4] disabled:opacity-40 shrink-0"
          >
            <ChevronRight className="w-4 h-4" />
          </button>

          <span className="text-[11px] text-[#5A6A7A] font-mono whitespace-nowrap shrink-0">
            {navIndex >= 0 ? navIndex + 1 : '-'}/{navList.length}
          </span>

          <label
            className="flex items-center gap-1.5 text-[11px] text-[#1A2332] cursor-pointer select-none whitespace-nowrap shrink-0"
            title="Chỉ liệt kê tủ có LỖI (cảnh báo không tính)"
          >
            <input
              type="checkbox"
              checked={onlyIssues}
              onChange={(e) => setOnlyIssues(e.target.checked)}
              className="w-3.5 h-3.5 accent-[#2F6F4E] cursor-pointer"
            />
            <span className="font-medium">Chỉ tủ có lỗi</span>
            <span className="font-mono text-[#DC2626]">({panelsWithErrors})</span>
          </label>
        </div>
      </div>

      {/* Circuit Spreadsheet Table — cuộn trong khung; thead sticky */}
      <div className="overflow-auto flex-1 min-h-0">
        <table className="w-full text-left text-sm text-[#1A2332] border-collapse">
          <thead className="sticky top-0 z-30 bg-[#E8EEF4] text-[#1B7A45] font-bold uppercase border-b border-[#D5DEE8] text-[12px] shadow-[0_1px_0_0_#D5DEE8]">
            <tr>
              <th className="py-3 px-2 w-14 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Row</th>
              <th className="py-3 px-2 min-w-[100px] text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Mã Mạch</th>
              <th className="py-3 px-2 min-w-[240px] text-center font-bold bg-[#E8EEF4]">Mô Tả Phụ Tải</th>
              <th className="py-3 px-1.5 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Full (kVA)</th>
              <th className="py-3 px-1.5 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">R</th>
              <th className="py-3 px-1.5 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Y</th>
              <th className="py-3 px-1.5 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">B</th>
              <th className="py-3 px-2 min-w-[88px] text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Itt (A)</th>
              <th className="py-3 px-2 min-w-[130px] text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Loại CB</th>
              <th className="py-3 px-2 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Số Cực</th>
              <th className="py-3 px-2 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">In (A)</th>
              <th className="py-3 px-2 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Isc (kA)</th>
              <th className="py-3 px-2 min-w-[180px] text-center font-bold bg-[#E8EEF4]">Dây Pha (Phase)</th>
              <th className="py-3 px-2 min-w-[150px] text-center font-bold bg-[#E8EEF4]">Dây PE</th>
              <th className="py-3 px-2 min-w-[220px] text-center font-bold bg-[#E8EEF4]">Giải Pháp Lắp Đặt (Installation)</th>
              <th className="py-3 px-2 w-20 text-center font-bold whitespace-nowrap bg-[#E8EEF4]">Kiểm Tra</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#D5DEE8] font-mono text-[12px]">
            {currentPanel.circuits.map((row) => {
              const rowIssues = panelIssuesMap.get(row.rowIndex) || [];
              const byField = groupIssuesByField(rowIssues);
              const at = (f: IssueField) => byField.get(f) || [];
              const hasRowError = rowIssues.some((i) => !i.isWarning);
              const hasRowWarning = rowIssues.some((i) => i.isWarning);
              const canEdit = !!onUpdateCircuit;

              return (
                <tr
                  key={row.rowIndex}
                  className={`transition-colors ${
                    hasRowError
                      ? 'bg-[#FEF2F2] hover:bg-[#FEE2E2]'
                      : hasRowWarning
                      ? 'bg-[#FFFBEB] hover:bg-[#FEF3C7]'
                      : 'hover:bg-[#F5F8FB]'
                  }`}
                >
                  <td className="py-3 px-2 text-center text-[#5A6A7A]">#{row.rowIndex}</td>
                  <td className="py-3 px-2 font-bold text-[#1B7A45] whitespace-nowrap text-[12px]">{row.lineName}</td>
                  <td className="py-3 px-2 font-sans text-[#1A2332] min-w-[240px] break-words whitespace-normal font-medium text-[12px] leading-snug">
                    {row.description}
                  </td>
                  <td className="py-3 px-1.5 text-center whitespace-nowrap font-semibold">{row.fullLoad > 0 ? row.fullLoad.toFixed(1) : '-'}</td>
                  <td className="py-3 px-1.5 text-center whitespace-nowrap">{row.rLoad > 0 ? row.rLoad.toFixed(1) : '-'}</td>
                  <td className="py-3 px-1.5 text-center whitespace-nowrap">{row.yLoad > 0 ? row.yLoad.toFixed(1) : '-'}</td>
                  <td className="py-3 px-1.5 text-center whitespace-nowrap">{row.bLoad > 0 ? row.bLoad.toFixed(1) : '-'}</td>
                  <td className="py-3 px-2 min-w-[88px] text-right font-bold text-[#1B7A45] whitespace-nowrap text-[12px]">
                    {row.hasExcelError && row.iCalc === 0 ? (
                      <span className="text-[#DC2626]">#REF!</span>
                    ) : row.iCalc > 0 ? (
                      `${row.iCalc.toFixed(1)} A`
                    ) : (
                      '-'
                    )}
                  </td>

                  {/* Loại CB */}
                  <td className="py-2 px-1.5 font-bold text-[#1A2332] whitespace-nowrap min-w-[110px]">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('cbType')}>
                      {canEdit ? (
                        <select
                          className={cbTypeSelectCls}
                          value={formatCbTypeLabel(row.cbType) || ''}
                          onChange={(e) => edit(row, 'cbType', e.target.value)}
                        >
                          <option value="">—</option>
                          {withCurrentOption(typeOpts, row.cbType, formatCbTypeLabel).map((opt) => (
                            <option key={opt} value={opt}>
                              {opt}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <span>{formatCbTypeLabel(row.cbType) || '-'}</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* Số cực */}
                  <td className="py-2 px-1.5 text-center whitespace-nowrap font-semibold">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('poleVal')}>
                      {canEdit ? (
                        <select
                          className={`${inputCls} text-center`}
                          value={formatPoleLabel(row.poleVal) || ''}
                          onChange={(e) => edit(row, 'poleVal', e.target.value)}
                        >
                          <option value="">—</option>
                          {withCurrentOption(poleOpts, row.poleVal, formatPoleLabel).map((opt) => (
                            <option key={opt} value={opt}>
                              {opt}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <span>{row.poleVal || '-'}</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* In (A) */}
                  <td className="py-2 px-1.5 text-right font-bold text-[#1B7A45] whitespace-nowrap text-[12px]">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('cbText')}>
                      {canEdit ? (
                        <select
                          className={`${inputCls} text-right font-bold text-[#1B7A45]`}
                          value={formatCbRatingLabel(row.cbText) || ''}
                          onChange={(e) => edit(row, 'cbText', e.target.value)}
                        >
                          <option value="">—</option>
                          {withCurrentOption(inOpts, row.cbText, formatCbRatingLabel).map((opt) => (
                            <option key={opt} value={opt}>
                              {opt}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <span>{row.cbAmp > 0 ? formatCbRatingLabel(row.cbText || row.cbAmp) : '-'}</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* Isc (kA) */}
                  <td className="py-2 px-1.5 text-center whitespace-nowrap font-semibold">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('cbIsc')}>
                      {canEdit ? (
                        <select
                          className={`${inputCls} text-center`}
                          value={formatIscLabel(row.cbIsc) || ''}
                          onChange={(e) => edit(row, 'cbIsc', e.target.value)}
                        >
                          <option value="">—</option>
                          {withCurrentOption(iscOpts, row.cbIsc, formatIscLabel).map((opt) => (
                            <option key={opt} value={opt}>
                              {opt}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <span>{row.cbIsc ? formatIscLabel(row.cbIsc) : '-'}</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* Dây pha */}
                  <td className="py-2 px-1.5 text-[#1A2332] min-w-[180px]">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('phaseCableText')}>
                      {canEdit ? (
                        <input
                          className={inputCls}
                          value={row.phaseCableText}
                          onChange={(e) => edit(row, 'phaseCableText', e.target.value)}
                        />
                      ) : (
                        <span className="break-words whitespace-normal leading-snug">{row.phaseCableText || '-'}</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* Dây PE */}
                  <td className="py-2 px-1.5 text-[#5A6A7A] min-w-[150px]">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('peCableText')}>
                      {canEdit ? (
                        <input
                          className={inputCls}
                          value={row.peCableText}
                          onChange={(e) => edit(row, 'peCableText', e.target.value)}
                        />
                      ) : (
                        <span className="break-words whitespace-normal leading-snug">{row.peCableText || '-'}</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* Giải pháp lắp đặt */}
                  <td className="py-2 px-1.5 text-[#334155] font-sans font-medium min-w-[220px]">
                    <CellWithBadge onShow={openTip} onHide={hideTip} isPinned={!!tip?.pinned} list={at('installMethod')}>
                      {canEdit ? (
                        <input
                          className={inputCls}
                          value={row.installMethod || ''}
                          onChange={(e) => edit(row, 'installMethod', e.target.value)}
                        />
                      ) : row.installMethod ? (
                        <span className="inline-block px-2.5 py-1 rounded bg-[#F0F4F8] border border-[#C5D0DC] text-[12px]">
                          {row.installMethod}
                        </span>
                      ) : (
                        <span>-</span>
                      )}
                    </CellWithBadge>
                  </td>

                  {/* Thẩm tra — tổng hợp gọn, chỉ 1 icon */}
                  <td className="py-2 px-2 text-center align-middle">
                    {rowIssues.length === 0 ? (
                      <span
                        className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-[#E6F4EC] border border-[#A8D4B8]"
                        title="Đạt tiêu chuẩn"
                      >
                        <Check className="w-3.5 h-3.5 text-[#1B7A45]" strokeWidth={3} />
                      </span>
                    ) : (
                      <IssueBadge
                        list={rowIssues}
                        className="px-1 py-0.5"
                        onShow={openTip}
                        onHide={hideTip}
                        isPinned={!!tip?.pinned}
                      />
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>

          {/* Khối tổng kết + lộ vào — nằm TRONG bảng để mọi giá trị thẳng cột với mạch con.
              sticky bottom: luôn thấy khi cuộn danh sách mạch dài. */}
          {currentPanel.footer && (
            <tfoot className="sticky bottom-0 z-30 font-mono text-[12px]">
              {/* Dòng 1: tổng CS định mức (trái) + nhãn LỘ VÀO (phải, tận dụng vùng trống) */}
              <tr className={footRowCls}>
                <td colSpan={3} className={footCellCls} style={{ boxShadow: FOOT_TOP_LINE }}>
                  {currentPanel.footer.ratedPower && <FootLabel text="Tổng CS định mức (kVA)" />}
                </td>
                <FootPhaseCells total={currentPanel.footer.ratedPower} shadow={FOOT_TOP_LINE} />
                <td className={footCellCls} style={{ boxShadow: FOOT_TOP_LINE }} />
                <td
                  colSpan={7}
                  className={footCellCls}
                  style={{ boxShadow: `${FOOT_TOP_LINE}, ${FOOT_LEFT_LINE}` }}
                >
                  {inc && (
                    <span className="inline-flex items-center gap-2 whitespace-nowrap">
                      <LogIn className="w-4 h-4 text-[#1B7A45]" strokeWidth={2.5} />
                      <span className="font-sans text-[12px] font-bold uppercase tracking-wide text-[#1B7A45]">
                        Lộ vào:
                      </span>
                      <span className="font-sans text-[12px] font-bold text-[#1A2332]">
                        {inc.source || '—'}
                      </span>
                    </span>
                  )}
                </td>
                <td className={footCellCls} style={{ boxShadow: FOOT_TOP_LINE }} />
              </tr>

              {/* Dòng 2: CS tính toán + dòng tính toán (trái) + thông số CB lộ vào (phải) */}
              <tr className={footRowCls}>
                <td colSpan={3} className={footCellCls}>
                  {currentPanel.footer.calcPower && <FootLabel text="CS tính toán (kVA)" />}
                  {currentPanel.footer.diversityFactor !== undefined && (
                    <span className="ml-2 font-sans text-[12px] font-semibold text-[#5A6A7A] whitespace-nowrap">
                      (Ks {currentPanel.footer.diversityFactor})
                    </span>
                  )}
                </td>
                <FootPhaseCells total={currentPanel.footer.calcPower} />
                {/* Dòng tính toán nằm đúng cột Itt (A) */}
                <td
                  className={`${footCellCls} text-right font-bold text-[#1B7A45] text-[12px] whitespace-nowrap`}
                >
                  {currentPanel.footer.calcCurrent !== undefined
                    ? `${currentPanel.footer.calcCurrent.toFixed(2)} A`
                    : ''}
                </td>
                <FootIncomingCell value={inc?.cbType} bold shadow={FOOT_LEFT_LINE} />
                <FootIncomingCell value={inc?.poleVal} align="center" />
                <FootIncomingCell value={inc?.cbText} align="right" bold />
                <FootIncomingCell value={inc?.cbIsc} align="center" />
                <FootIncomingCell
                  value={inc?.phaseCableText}
                  wrap
                  badgeList={incAt('phaseCableText')}
                  onShowBadge={openTip}
                  onHideBadge={hideTip}
                  isPinned={!!tip?.pinned}
                />
                <FootIncomingCell
                  value={inc?.peCableText}
                  wrap
                  badgeList={incAt('peCableText')}
                  onShowBadge={openTip}
                  onHideBadge={hideTip}
                  isPinned={!!tip?.pinned}
                />
                <FootIncomingCell
                  value={inc?.installMethod}
                  wrap
                  sans
                  badgeList={incAt('installMethod')}
                  onShowBadge={openTip}
                  onHideBadge={hideTip}
                  isPinned={!!tip?.pinned}
                />
                <td className={`${footCellCls} text-center align-middle`}>
                  {inc ? (
                    incomingIssues.length === 0 ? (
                      <span
                        className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-[#E6F4EC] border border-[#A8D4B8]"
                        title="Đạt tiêu chuẩn"
                      >
                        <Check className="w-3.5 h-3.5 text-[#1B7A45]" strokeWidth={3} />
                      </span>
                    ) : (
                      <IssueBadge
                        list={incomingIssues}
                        className="px-1 py-0.5"
                        onShow={openTip}
                        onHide={hideTip}
                        isPinned={!!tip?.pinned}
                      />
                    )
                  ) : null}
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      {/* Danh sách tủ — fixed để không bị khung card (overflow-hidden) cắt mất */}
      {listPos && (
        <div
          className="fixed z-50 max-h-[400px] flex flex-col bg-white border border-[#D5DEE8] rounded-xl shadow-lg overflow-hidden"
          style={{ left: listPos.x, top: listPos.y, width: listPos.w }}
          onClick={(e) => e.stopPropagation()}
        >
          {/* Ô tìm nhanh tên tủ — sticky trên đầu, không cuộn theo danh sách */}
          <div className="shrink-0 p-2 border-b border-[#E8EEF4] bg-white">
            <input
              type="text"
              autoFocus
              value={panelSearch}
              onChange={(e) => setPanelSearch(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Escape') {
                  e.stopPropagation();
                  setListPos(null);
                }
              }}
              placeholder="Tìm tên tủ..."
              className="w-full bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 text-[13px] focus:outline-none focus:border-[#1B7A45]"
            />
          </div>

          <div className="overflow-auto py-1">
            {searchedNavList.map((p) => {
              const c = countsByPanel.get(p.sheetName) || { err: 0, warn: 0 };
              const active = p.sheetName === currentPanel.sheetName;
              return (
                <button
                  key={p.sheetName}
                  type="button"
                  onClick={() => {
                    setSelectedSheet(p.sheetName);
                    setListPos(null);
                  }}
                  className={`w-full px-3 py-2 flex items-center justify-between gap-3 text-left text-[13px] transition-colors ${
                    active
                      ? 'bg-[#E6F4EC] text-[#1B7A45] font-bold'
                      : 'text-[#1A2332] hover:bg-[#F0F4F8]'
                  }`}
                >
                  <span className="truncate">{p.sheetName}</span>
                  <CountBadges err={c.err} warn={c.warn} />
                </button>
              );
            })}
            {searchedNavList.length === 0 && (
              <div className="px-3 py-3 text-[12px] text-[#5A6A7A]">
                {panelSearch.trim()
                  ? `Không tìm thấy tủ khớp "${panelSearch.trim()}".`
                  : 'Không có tủ nào.'}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Tooltip dùng chung — position fixed nên không bị khung cuộn cắt */}
      {tip && (
        <div
          className="fixed z-50 w-[300px] max-w-[90vw]"
          style={{ left: tip.x, top: tip.y, transform: 'translateX(-50%)' }}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="bg-white border border-[#D5DEE8] rounded-xl shadow-lg overflow-hidden">
            {tip.issues.map((iss, idx) => (
              <div
                key={iss.id}
                className={`flex items-start gap-2 p-2.5 font-sans text-[13px] leading-snug ${
                  idx > 0 ? 'border-t border-[#E8EEF4]' : ''
                } ${iss.isWarning ? 'bg-[#FFFBEB] text-[#B45309]' : 'bg-[#FEF2F2] text-[#DC2626]'}`}
              >
                {iss.isWarning ? (
                  <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                )}
                <span className="font-medium">{iss.description}</span>
              </div>
            ))}
            {tip.pinned && (
              <div className="px-2.5 py-1.5 bg-[#F5F8FB] border-t border-[#E8EEF4] text-[11px] text-[#5A6A7A] font-sans">
                Bấm ra ngoài hoặc Esc để đóng
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
