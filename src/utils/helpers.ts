/**
 * Helper utility functions mirroring VBA logic
 */

/**
 * Mã lỗi Excel (cell.t === 'e') — SheetJS lưu v = error code
 * #REF! = 23 → dễ bị nhầm thành dòng điện / công suất nếu không kiểm tra.
 */
export const EXCEL_ERROR_CODES: Record<number, string> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
};

/** Cell kiểu SheetJS tối thiểu */
export interface ExcelLikeCell {
  t?: string;
  v?: unknown;
  w?: string;
  f?: string;
}

/** Kiểm tra ô là lỗi công thức Excel (#REF!, #N/A, ...) */
export function isExcelErrorCell(cell: ExcelLikeCell | null | undefined): boolean {
  if (!cell) return false;
  if (cell.t === 'e') return true;
  const text = String(cell.w ?? cell.v ?? '').toUpperCase();
  return (
    text.includes('#REF!') ||
    text.includes('#N/A') ||
    text.includes('#VALUE!') ||
    text.includes('#DIV/0!') ||
    text.includes('#NAME?') ||
    text.includes('#NUM!') ||
    text.includes('#NULL!')
  );
}

/** Nhãn lỗi hiển thị (#REF!, ...) */
export function getExcelErrorLabel(cell: ExcelLikeCell | null | undefined): string {
  if (!cell) return '';
  if (cell.t === 'e' && typeof cell.v === 'number' && EXCEL_ERROR_CODES[cell.v]) {
    return EXCEL_ERROR_CODES[cell.v];
  }
  const text = String(cell.w ?? cell.v ?? '').trim();
  if (text.startsWith('#')) return text.toUpperCase();
  return '#ERR!';
}

/**
 * Đọc số từ ô Excel — trả 0 nếu ô lỗi (#REF! = mã 23 không được coi là số).
 */
export function extractNumberFromCell(cell: ExcelLikeCell | null | undefined): number {
  if (!cell || isExcelErrorCell(cell)) return 0;
  return extractNumber(cell.v);
}

/** Đọc text từ ô — ô lỗi trả chuỗi rỗng (hoặc label nếu wantErrorLabel) */
export function extractTextFromCell(
  cell: ExcelLikeCell | null | undefined,
  options?: { keepErrorLabel?: boolean }
): string {
  if (!cell) return '';
  if (isExcelErrorCell(cell)) {
    return options?.keepErrorLabel ? getExcelErrorLabel(cell) : '';
  }
  if (cell.v === null || cell.v === undefined) return '';
  return String(cell.v).trim();
}

export function extractNumber(val: unknown): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') {
    // Tránh dùng nhầm mã lỗi Excel nếu caller quên check cell.t
    return val;
  }
  const s = String(val).replace(/,/g, '').trim();
  // Chuỗi lỗi Excel
  if (s.startsWith('#')) return 0;

  let tmp = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if ((ch >= '0' && ch <= '9') || ch === '.') {
      tmp += ch;
    } else if (tmp.length > 0) {
      break;
    }
  }
  return tmp === '' ? 0 : parseFloat(tmp);
}

export function getShortDesc(desc: string): string {
  if (!desc) return '';
  let firstLine = desc.split('\n')[0].trim();
  const uLine = firstLine.toUpperCase();

  if (uLine.includes('DỰ PHÒNG') || uLine.includes('SPARE') || uLine.includes('DU PHONG')) {
    return 'Dự phòng';
  }

  if (uLine.startsWith('CẤP NGUỒN ')) {
    firstLine = firstLine.substring(10).trim();
  } else if (uLine.startsWith('CAP NGUON ')) {
    firstLine = firstLine.substring(10).trim();
  }

  if (firstLine.toUpperCase().startsWith('CHO ')) {
    firstLine = firstLine.substring(4).trim();
  }

  if (firstLine.toUpperCase().startsWith('TỦ ĐIỆN ')) {
    firstLine = firstLine.substring(8).trim();
  } else if (firstLine.toUpperCase().startsWith('TU DIEN ')) {
    firstLine = firstLine.substring(8).trim();
  }

  if (firstLine.length > 35) {
    firstLine = firstLine.substring(0, 32) + '...';
  }

  return firstLine;
}

export function getPrefix(lineName: string): string {
  const s = lineName.trim().toUpperCase();
  if (s.startsWith('SP')) return 'SP';
  if (s.startsWith('CP')) return 'CP';
  if (s.startsWith('CT')) return 'CT';
  if (s.startsWith('ST')) return 'ST';
  if (s.startsWith('SD')) return 'SD';
  if (s.startsWith('SOL')) return 'SOL';
  if (s.startsWith('L')) return 'L';
  if (s.startsWith('S')) return 'S';
  if (s.startsWith('P')) return 'P';
  return '';
}

export function isSummaryRow(s: string): boolean {
  if (!s) return false;
  const t = s.trim().toLowerCase();
  const keywords = [
    'công suất tính toán',
    'power calculations',
    'công suất kết nối',
    'total connected load',
    'hệ số đồng thời',
    'diversity factor',
    'dòng điện tính toán',
    'current (a)',
    'tổng công suất',
    'dòng điện tổng',
    'cos phi',
    'hệ số công suất'
  ];
  if (keywords.some(k => t.includes(k))) return true;

  // Nhan dang chinh xac cac nhan dong tong dang ngan.
  // Phai so khop tuyet doi: 'sum' neu dung includes() se dinh nham 'consumption'.
  const exact = ['sum', 'total', 'tong', 'tong cong'];
  if (exact.includes(t)) return true;

  return /^total\s+(rate|cal|calculated|connected)\b/.test(t);
}

export function isFireCircuit(desc: string, lineName: string): boolean {
  const text = (desc + ' ' + lineName).toLowerCase();
  const keywords = [
    'sumpit',
    'hut khoi',
    'hút khói',
    'bu khi',
    'bù khí',
    'tao ap',
    'tạo áp',
    'fire pump',
    'smoke',
    'fcc',
    'pccc',
    'bơm chữa cháy',
    'chữa cháy'
  ];
  return keywords.some(k => text.includes(k));
}

export function isLowSpeedFan(desc: string): boolean {
  const t = desc.toLowerCase();
  return t.includes('low speed') || t.includes('low-speed');
}

export function normalizePoleValue(val: unknown): string {
  if (val === null || val === undefined) return '';
  return String(val).toUpperCase().replace(/\s+/g, '');
}

export function isThreePhaseByLoads(rVal: number, yVal: number, bVal: number): boolean {
  let cnt = 0;
  if (rVal > 0) cnt++;
  if (yVal > 0) cnt++;
  if (bVal > 0) cnt++;
  return cnt >= 2;
}

/**
 * Tải 3 pha có cân bằng không (R ≈ Y ≈ B, sai lệch trong ngưỡng tolerance).
 * Cân bằng (vd tải động cơ 3 pha thuần) -> không bắt buộc dây trung tính, 3 dây (L1,L2,L3) là đủ.
 * Không cân bằng (vd tủ tổng gộp nhiều tải 1 pha rải trên 3 pha) -> dòng trung tính khác 0,
 * bắt buộc phải có dây N -> tối thiểu 4 dây (L1,L2,L3,N).
 */
export function isBalancedThreePhaseLoad(
  rVal: number,
  yVal: number,
  bVal: number,
  tolerance = 0.03
): boolean {
  if (rVal <= 0 || yVal <= 0 || bVal <= 0) return false;
  const maxV = Math.max(rVal, yVal, bVal);
  const minV = Math.min(rVal, yVal, bVal);
  if (maxV <= 0) return false;
  return (maxV - minV) / maxV <= tolerance;
}

export function isValidSinglePhasePole(poleVal: string): boolean {
  const p = normalizePoleValue(poleVal);
  return p === '1P' || p === '1P+N' || p === '2P';
}

export function isValidThreePhasePole(poleVal: string): boolean {
  const p = normalizePoleValue(poleVal);
  return p === '3P' || p === '4P';
}

export function containsAny(textValue: string, keywords: string[]): boolean {
  const lower = textValue.toLowerCase();
  return keywords.some(k => lower.includes(k.toLowerCase()));
}

/**
 * Strips cable specification brand/material prefixes (e.g. "Cu/PVC 1x", "Cu/XLPE/PVC 1x")
 * and leaves only the cable cross-section value (e.g. "1.5", "2.5", "4", "2x150").
 */
export function cleanCableSectionOnly(str: string): string {
  if (!str) return '';
  let s = str.trim();

  // Strip common conductor/insulation/sheath material prefixes
  s = s.replace(/^(?:CU|AL)\s*[\/.\s]*(?:XLPE|PVC|CXV|CV|LSZH|FR)*[\/.\s]*(?:PVC|LSZH)*\s*/i, '');
  
  // Strip core multipliers like 1x, 1C-, 1x1C-
  s = s.replace(/^(?:1[xX*]|1C[-_.\s]*)+/i, '');

  // Strip trailing mm2 or sqmm
  s = s.replace(/\s*(?:mm2|sqmm)\b/i, '');

  return s.trim() || str.trim();
}

export function getSheetNameFromFormula(formula?: string): string {
  if (!formula || !formula.startsWith('=')) return '';
  const exclPos = formula.indexOf('!');
  if (exclPos > 0) {
    let sheetName = formula.substring(1, exclPos).replace(/'/g, '');
    const bracketPos = sheetName.indexOf(']');
    if (bracketPos > 0) {
      sheetName = sheetName.substring(bracketPos + 1);
    }
    return sheetName.trim();
  }
  return '';
}
