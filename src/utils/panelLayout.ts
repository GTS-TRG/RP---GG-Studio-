import * as XLSX from 'xlsx';

/**
 * Nhận diện bố cục sheet tủ điện theo TIÊU ĐỀ CỘT thay vì gán cứng vị trí.
 * Lý do: mỗi dự án có form khác nhau — cột và dòng bắt đầu không cố định.
 */

export interface PanelCols {
  lineName?: number;
  description?: number;
  rLoad?: number;
  yLoad?: number;
  bLoad?: number;
  iCalc?: number;
  cbType?: number;
  poleVal?: number;
  cbText?: number;
  cbIsc?: number;
  phaseCableText?: number;
  peCableText?: number;
  installMethod?: number;
}

export interface PanelLayout {
  cols: PanelCols;
  /** Dòng dữ liệu đầu tiên (0-based) */
  dataStartRow: number;
  /** Dòng tiêu đề đã dùng để dò (0-based) */
  headerRow: number;
  /** Số trường nhận ra được — dùng để đánh giá độ tin cậy */
  matched: number;
  /** true = dò được theo tiêu đề; false = dùng bố cục mặc định */
  detected: boolean;
}

/** Bố cục mặc định (form cũ): A=Tên mạch, B=Mô tả, D/E/F=R/Y/B, G=Itt, H..N=CB & cáp */
export const LEGACY_COLS: PanelCols = {
  lineName: 0,
  description: 1,
  rLoad: 3,
  yLoad: 4,
  bLoad: 5,
  iCalc: 6,
  cbType: 7,
  poleVal: 8,
  cbText: 9,
  cbIsc: 10,
  phaseCableText: 11,
  peCableText: 12,
  installMethod: 13,
};

/** Bỏ dấu tiếng Việt + chuẩn hoá để so khớp tiêu đề */
export function normHeader(s: unknown): string {
  return String(s ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/Đ/g, 'D')
    .replace(/đ/g, 'd')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

type FieldKey = keyof PanelCols;

interface Rule {
  field: FieldKey;
  test: (t: string) => boolean;
}

/**
 * Thứ tự QUAN TRỌNG: luật hẹp đặt trước luật rộng.
 * Ví dụ "DONG NGAT DM CURRENT (IN)" phải vào cbText, không được rơi vào iCalc.
 */
const RULES: Rule[] = [
  // Dòng ngắt định mức In — phải xét trước cột dòng điện tính toán
  {
    field: 'cbText',
    test: (t) =>
      /\(\s*IN\s*\)/.test(t) ||
      /DONG NGAT/.test(t) ||
      /NGAT CURRENT/.test(t) ||
      /RATED CURRENT/.test(t),
  },
  { field: 'cbIsc', test: (t) => /\bISC\b/.test(t) || /BREAKING CAPACITY/.test(t) },
  { field: 'poleVal', test: (t) => /\bPOLE\b/.test(t) || /SO CUC/.test(t) },
  {
    field: 'cbType',
    test: (t) =>
      (/\bLOAI\b/.test(t) || /\bTYPE\b/.test(t)) && !/CABLE|CAP|DAY/.test(t),
  },
  {
    field: 'installMethod',
    test: (t) => /INSTALLATION/.test(t) || /GIAI PHAP LAP DAT/.test(t),
  },
  {
    field: 'peCableText',
    test: (t) =>
      /EARTHING/.test(t) ||
      /DAY NOI DAT/.test(t) ||
      /DAY PE/.test(t) ||
      /CABEL E\b/.test(t) ||
      /CABLE E\b/.test(t) ||
      /^E$/.test(t),
  },
  {
    // Bắt buộc có ngữ cảnh "cáp" (CABLE / mm²), nếu không sẽ dính nhầm cột "Phase R (kW)"
    field: 'phaseCableText',
    test: (t) =>
      /CABLE PHASE/.test(t) ||
      /DAY PHA/.test(t) ||
      (/\bPHASE\b/.test(t) && /(CABLE|CAP|MM2|MM²)/.test(t)),
  },
  // Dòng điện tính toán — loại trừ các biến thể thuộc về In
  {
    field: 'iCalc',
    test: (t) =>
      (/DONG DIEN/.test(t) || /CURRENT/.test(t) || /\bI\s*CALC\b/.test(t)) &&
      !/NGAT|\(\s*IN\s*\)|RATED/.test(t),
  },
  {
    field: 'lineName',
    test: (t) => /TEN MACH/.test(t) || (/\bLINE\b/.test(t) && !/DESCRIPTION/.test(t)),
  },
  {
    field: 'description',
    test: (t) => /DESCRIPTION/.test(t) || /CONG NANG/.test(t) || /MO TA/.test(t),
  },
];

/** Cột công suất theo pha: "R" / "Y" / "B" hoặc "Phase R (kW)" */
function phaseLetterOf(sub: string): 'rLoad' | 'yLoad' | 'bLoad' | null {
  if (sub === 'R' || /^PHASE R\b/.test(sub)) return 'rLoad';
  if (sub === 'Y' || /^PHASE Y\b/.test(sub)) return 'yLoad';
  if (sub === 'B' || /^PHASE B\b/.test(sub)) return 'bLoad';
  return null;
}

/**
 * Gộp text tiêu đề của một cột từ dải nhiều dòng.
 * Cần thiết vì form thường tách tiêu đề nhóm (dòng trên) và tiêu đề con (dòng dưới).
 */
function columnHeaderText(
  ws: XLSX.WorkSheet,
  rowFrom: number,
  rowTo: number,
  col: number
): string {
  const parts: string[] = [];
  for (let r = rowFrom; r <= rowTo; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
    if (!cell || cell.v == null || cell.v === '') continue;
    parts.push(normHeader(cell.v));
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

/** Dòng đánh số thứ tự cột (1,2,3...) — cần bỏ qua, không phải dữ liệu */
function isNumberingRow(ws: XLSX.WorkSheet, r: number, maxCol: number): boolean {
  let markers = 0;
  let others = 0;
  for (let c = 0; c <= maxCol; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    if (!cell || cell.v == null || cell.v === '') continue;
    const raw = String(cell.v).trim();
    const n = Number(raw);
    // Dong danh so cot: 1,2,3... hoac nhan cot dang (A), (B), (C)
    if (Number.isInteger(n) && n >= 1 && n <= 40) markers++;
    else if (/^\(?[A-Za-z]{1,2}\)?$/.test(raw)) markers++;
    else others++;
  }
  return markers >= 4 && others === 0;
}

/**
 * Dò bố cục của một sheet tủ điện.
 * fallbackStartRow: config.startRow (1-based) dùng khi không dò được.
 */
export function detectPanelLayout(ws: XLSX.WorkSheet, fallbackStartRow: number): PanelLayout {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z200');
  const maxCol = Math.min(range.e.c, 40);
  const maxScanRow = Math.min(range.e.r, 40);

  let best: { row: number; cols: PanelCols; matched: number } | null = null;

  for (let r = range.s.r; r <= maxScanRow; r++) {
    const bandFrom = Math.max(range.s.r, r - 2);
    const cols: PanelCols = {};
    let matched = 0;

    // bandTexts: gộp tiêu đề nhóm + tiêu đề con (dùng cho hầu hết các trường)
    // subTexts: chỉ tiêu đề con của đúng dòng r (R/Y/B luôn nằm ở tiêu đề con)
    const bandTexts: string[] = [];
    const subTexts: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      bandTexts[c] = columnHeaderText(ws, bandFrom, r, c);
      subTexts[c] = columnHeaderText(ws, r, r, c);
    }

    // B1: giữ chỗ cột công suất theo pha (R/Y/B) TRƯỚC, để các luật khác
    // không chiếm nhầm — ví dụ "Phase R (kW)" có chứa chữ PHASE.
    const anchorCol = bandTexts.findIndex(
      (t) => t && (/FULL/.test(t) || /RATE POWER/.test(t))
    );
    const reserved = new Set<number>();
    const winStart = anchorCol >= 0 ? anchorCol : 0;
    const winEnd = anchorCol >= 0 ? Math.min(maxCol, anchorCol + 6) : maxCol;

    for (let c = winStart; c <= winEnd; c++) {
      const field = phaseLetterOf(subTexts[c] || '');
      if (field && cols[field] === undefined) {
        cols[field] = c;
        reserved.add(c);
        matched++;
      }
    }
    if (anchorCol >= 0) reserved.add(anchorCol);

    // B2: các trường còn lại dò theo tiêu đề gộp
    for (let c = 0; c <= maxCol; c++) {
      if (reserved.has(c)) continue;
      const text = bandTexts[c];
      if (!text) continue;

      for (const rule of RULES) {
        if (cols[rule.field] !== undefined) continue; // mỗi trường chỉ lấy cột đầu tiên khớp
        if (rule.test(text)) {
          cols[rule.field] = c;
          matched++;
          break; // mỗi cột chỉ gán 1 trường
        }
      }
    }

    if (!best || matched > best.matched) {
      best = { row: r, cols, matched };
    }
  }

  // Cần đủ các trường cốt lõi mới coi là dò được
  const okCore =
    best !== null &&
    best.matched >= 6 &&
    best.cols.cbType !== undefined &&
    best.cols.poleVal !== undefined &&
    best.cols.cbText !== undefined &&
    best.cols.phaseCableText !== undefined;

  if (!best || !okCore) {
    return {
      cols: { ...LEGACY_COLS },
      dataStartRow: Math.max(0, fallbackStartRow - 1),
      headerRow: Math.max(0, fallbackStartRow - 2),
      matched: 0,
      detected: false,
    };
  }

  // Dòng dữ liệu đầu tiên: sau tiêu đề, bỏ qua dòng trống và dòng đánh số cột
  const keyCol = best.cols.lineName ?? best.cols.description ?? 0;

  /** Dòng dữ liệu thật phải có ít nhất một thông số kỹ thuật, không chỉ có chữ ở cột tên */
  const looksLikeDataRow = (r: number): boolean => {
    const key = ws[XLSX.utils.encode_cell({ r, c: keyCol })];
    if (!key || key.v == null || String(key.v).trim() === '') return false;

    const probes = [
      best!.cols.cbType,
      best!.cols.cbText,
      best!.cols.poleVal,
      best!.cols.iCalc,
      best!.cols.phaseCableText,
    ];
    return probes.some((c) => {
      if (c === undefined) return false;
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return cell && cell.v != null && String(cell.v).trim() !== '';
    });
  };

  let dataStart = best.row + 1;
  for (let r = best.row + 1; r <= Math.min(range.e.r, best.row + 15); r++) {
    if (isNumberingRow(ws, r, maxCol)) continue;
    if (looksLikeDataRow(r)) {
      dataStart = r;
      break;
    }
  }

  return {
    cols: best.cols,
    dataStartRow: dataStart,
    headerRow: best.row,
    matched: best.matched,
    detected: true,
  };
}
