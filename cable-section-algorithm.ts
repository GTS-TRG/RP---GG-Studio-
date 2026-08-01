/**
 * ============================================================================
 * THUẬT TOÁN NHẬN DIỆN TIẾT DIỆN CÁP & SỐ LƯỢNG CÁP TRÊN 1 PHA
 * ============================================================================
 * Trích xuất từ dự án "GG Studio - Panel Schedule Reviewer".
 * File này KHÔNG phụ thuộc gì vào phần còn lại của dự án gốc — có thể copy
 * nguyên văn sang project khác (Node/TS, hoặc xoá type annotation để dùng
 * như .js thuần).
 *
 * ĐẦU VÀO: một chuỗi text mô tả dây/cáp pha, ví dụ lấy từ ô Excel bảng tủ điện:
 *    "3x1C-4 Cu/XLPE/PVC"
 *    "1x4C-2.5"
 *    "2x(3x1C-6) Cu/PVC"
 *    "4C-10 Cu/XLPE/PVC + 1C-4"
 *
 * ĐẦU RA:
 *  1. parseTextToDict()   -> Map liệt kê từng "loại cáp" (tiết diện, vỏ, số lõi,
 *                             số lượng sợi) đã nhận diện được trong chuỗi.
 *  2. getCableCoreStats() -> Thống kê tổng số lõi/dây pha, số cáp 1 lõi, số cáp
 *                             đa lõi... để so sánh với yêu cầu (mạch 1 pha cần
 *                             >=2 lõi, mạch 3 pha cần >=3 lõi, v.v).
 *
 * Ý TƯỞNG THUẬT TOÁN (parseTextToDict):
 *  - Cho phép nhiều cụm cáp nối nhau bằng dấu "+" (mỗi cụm parse riêng).
 *  - Dò theo 4 nhóm regex, ưu tiên từ đặc thù -> tổng quát, sau khi 1 nhóm
 *    match được đoạn nào thì "xoá" đoạn đó khỏi chuỗi (thay bằng khoảng trắng)
 *    để nhóm sau không match trùng:
 *      Nhóm 1: (n)x1C-size[vỏ]         -> cáp 1 lõi, có thể có hệ số nhân/bó
 *      Nhóm 2: (n)x kC-size[vỏ], k>=2  -> cáp đa lõi (2C,3C,4C,5C...)
 *      Nhóm 3: N x size[vỏ] (không chữ C) -> suy luận N=2 => 2 lõi, ngược lại N sợi 1 lõi
 *      Nhóm 3b: các biến thể còn sót của "1C-size" (vd "(1C-1.5)")
 *  - Gộp (merge) các cụm trùng {tiết diện, vỏ, số lõi} bằng cách cộng dồn số lượng.
 *  - Nếu không parse được gì bằng regex chính, có fallback regex đơn giản hơn
 *    (bắt số trước "C") để tránh trả về rỗng hoàn toàn.
 *
 * Ý TƯỞNG THUẬT TOÁN (getCableCoreStats):
 *  - Duyệt dict kết quả ở trên, với mỗi loại cáp:
 *      totalCores += count * coreCount
 *      nếu coreCount === 1 -> cộng vào singleCoreCableCount (đếm SỢI cáp 1 lõi)
 *      nếu coreCount  > 1 -> cộng vào multiCoreCableCount (đếm SỢI cáp đa lõi)
 *      theo dõi coreCount lớn nhất từng gặp (maxSingleCableCoreCount)
 *  - Từ các số liệu này, người dùng có thể tự áp quy tắc, ví dụ:
 *      Mạch 1 pha (L+N)  cần totalPhaseCores >= 2
 *      Mạch 3 pha (L1L2L3) cần totalPhaseCores >= 3 (nếu tải CÂN BẰNG)
 *                                              >= 4 (nếu tải KHÔNG cân bằng — cần thêm dây N)
 *
 * 3. isBalancedThreePhaseLoad() + checkThreePhaseCoreCount() (mục 4 bên dưới):
 *    Bổ sung quan trọng — 1 mạch/lộ 3 pha (CB 3P) không mặc định chỉ cần 3 dây.
 *    Nếu tải giữa 3 pha R/Y/B KHÔNG bằng nhau (vd tủ tổng gộp nhiều tải 1 pha
 *    rải trên 3 pha để cân tải), dòng điện chạy về trung tính khác 0 -> BẮT BUỘC
 *    có dây N -> cần tối thiểu 4 dây (L1,L2,L3,N), không phải 3.
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// 1. DANH SÁCH QUY CÁCH VỎ CÁP (SHEATH TYPES) — có thể tuỳ biến theo dự án
// ----------------------------------------------------------------------------

export const CABLE_SHEATH_TYPES = [
  'CU/PVC',
  'CU/XLPE/PVC',
  'CU/MICA/XLPE/FR-PVC',
  'CU/MICA/XLPE/LSZH',
  'CU/PVC/PVC',
] as const;

export type CableSheathType = string;

/** Chuẩn hoá tên vỏ để so khớp không phân biệt khoảng trắng/hoa-thường */
export function normalizeSheathKey(raw: string): string {
  return raw.toUpperCase().replace(/\s+/g, '');
}

/** Chuẩn hoá tên cột vỏ hiển thị (dùng khi lưu/khoá Map) */
export function formatSheathColumnName(raw: string): string {
  return raw.trim().replace(/\s+/g, '').toUpperCase();
}

/**
 * Nhận diện quy cách vỏ cáp từ 1 đoạn text tự do.
 * VD: "Cu/XLPE/PVC 3C-2.5mm2" -> "CU/XLPE/PVC"
 *     "Mica LSZH"             -> "CU/MICA/XLPE/LSZH"
 * Ưu tiên khớp theo tên dài nhất trước (tránh "XLPE" nuốt mất "MICA/XLPE").
 */
export function detectSheathType(
  cableText: string,
  availableSheaths?: string[]
): CableSheathType | undefined {
  if (!cableText) return undefined;
  const t = normalizeSheathKey(cableText);

  const pool = (availableSheaths?.length ? availableSheaths : [...CABLE_SHEATH_TYPES])
    .slice()
    .sort((a, b) => normalizeSheathKey(b).length - normalizeSheathKey(a).length);

  for (const s of pool) {
    const ns = normalizeSheathKey(s);
    if (ns.length >= 3 && t.includes(ns)) return formatSheathColumnName(s);
  }

  if (t.includes('MICA') && (t.includes('LSZH') || t.includes('LSOH'))) return 'CU/MICA/XLPE/LSZH';
  if (t.includes('MICA') && (t.includes('FR-PVC') || t.includes('FRPVC') || t.includes('FR/'))) {
    return 'CU/MICA/XLPE/FR-PVC';
  }
  if (t.includes('MICA')) return 'CU/MICA/XLPE/FR-PVC';
  if (t.includes('XLPE')) return 'CU/XLPE/PVC';
  if (t.includes('PVC/PVC') || t.includes('PVCPVC')) return 'CU/PVC/PVC';
  if (t.includes('PVC') || t.includes('CXV') || t.includes('/CV')) return 'CU/PVC';
  return undefined;
}

/** Chuẩn hoá 1 chuỗi vỏ "thô" bắt được từ regex về CableSheathType đã biết */
export function resolveSheathType(
  raw: string,
  fallbackText?: string,
  availableSheaths?: string[]
): CableSheathType | undefined {
  const pool = availableSheaths?.length ? availableSheaths : [...CABLE_SHEATH_TYPES];
  const cleaned = raw.replace(/\bmm2\b/gi, '').replace(/\bsqmm\b/gi, '').trim();

  if (cleaned) {
    const n = normalizeSheathKey(cleaned);
    for (const s of pool) {
      if (normalizeSheathKey(s) === n) return formatSheathColumnName(s);
    }
    const detected = detectSheathType(cleaned, pool);
    if (detected) return detected;
  }
  if (fallbackText) return detectSheathType(fallbackText, pool);
  return undefined;
}

// ----------------------------------------------------------------------------
// 2. NHẬN DIỆN TIẾT DIỆN + SỐ LƯỢNG CÁP TỪ CHUỖI TEXT (parseTextToDict)
// ----------------------------------------------------------------------------

/** Khoá gộp: "tiết_diện|vỏ|số_lõi" */
export type CableDictKey = string;

export interface CableDictEntry {
  /** Tiết diện dây (mm2), ví dụ 1.5, 2.5, 4, 6, 10, 16, 25... */
  sectionMM2: number;
  /** Chuỗi vỏ "thô" bắt được (trước khi resolve) */
  sheathRaw: string;
  /** Loại vỏ đã chuẩn hoá, undefined nếu không nhận diện được */
  sheathType?: CableSheathType;
  /** Số lõi trong 1 sợi cáp: 1 = cáp đơn (1C), >=2 = cáp nhiều lõi (2C/3C/4C/5C...) */
  coreCount: number;
  /** Số SỢI cáp loại này (đã cộng dồn nếu trùng loại) */
  count: number;
}

function makeDictKey(sectionMM2: number, sheathRaw: string, coreCount: number): CableDictKey {
  return `${sectionMM2}|${sheathRaw.toLowerCase()}|${coreCount}`;
}

function addToDict(
  dict: Map<CableDictKey, CableDictEntry>,
  entry: Omit<CableDictEntry, 'count'> & { count: number }
) {
  const key = makeDictKey(entry.sectionMM2, entry.sheathRaw, entry.coreCount);
  const existing = dict.get(key);
  if (existing) {
    existing.count += entry.count;
  } else {
    dict.set(key, { ...entry });
  }
}

/**
 * Parse 1 chuỗi mô tả cáp (pha hoặc PE) thành Map các loại cáp đã nhận diện.
 *
 * Hỗ trợ các dạng viết phổ biến trong bảng tủ điện VN:
 *   "3x1C-4"                 -> 3 sợi, 1 lõi, 4mm2
 *   "3x1C-4 Cu/XLPE/PVC"     -> như trên + vỏ Cu/XLPE/PVC
 *   "2x(3x1C-6)"             -> 2 bó x 3 sợi = 6 sợi, 1 lõi, 6mm2
 *   "1x4C-2.5"                -> 1 sợi cáp 4 lõi, 2.5mm2
 *   "2x4C-10 Cu/XLPE/PVC"    -> 2 sợi cáp 4 lõi, 10mm2
 *   "4x4"  (không chữ C)      -> N=4 -> hiểu là 4 sợi 1 lõi, 4mm2
 *   "2x4"  (không chữ C)      -> N=2 -> hiểu là 1 sợi cáp 2 lõi, 4mm2
 *   "4C-10 + 1C-4"           -> nhiều cụm cáp nối bằng dấu "+"
 *
 * @param srcText          chuỗi cần parse
 * @param availableSheaths danh sách vỏ hợp lệ (mặc định dùng CABLE_SHEATH_TYPES)
 */
export function parseTextToDict(
  srcText: string,
  availableSheaths?: string[]
): Map<CableDictKey, CableDictEntry> {
  const dict = new Map<CableDictKey, CableDictEntry>();
  if (!srcText || !String(srcText).trim()) return dict;

  const fullText = String(srcText).toLowerCase().trim();
  const textParts = fullText.split('+');

  for (const rawPart of textParts) {
    let textPart = rawPart.trim();
    if (!textPart) continue;

    textPart = textPart.replace(/\s*(?:mm2|sqmm)\b/gi, '');
    const partFallbackSheath = detectSheathType(rawPart, availableSheaths);

    // Nhóm 1: (cụm)?(n)x1C-size[vỏ]  -> cáp 1 lõi
    const re1 =
      /(\d+)?x?\(?(\d+)\s*x\s*1\s*[cC]\s*[-_x×\s]\s*(\d+(?:[.,]\d+)?)(?:\))?\s*([\w/\-]*)/gi;
    let tmp = textPart;
    let m: RegExpExecArray | null;
    const matched1: string[] = [];

    while ((m = re1.exec(textPart)) !== null) {
      const soCum1 = m[1] ? parseInt(m[1], 10) : 1;
      const soNhan1 = parseInt(m[2], 10) || 1;
      const td1 = parseFloat(m[3].replace(',', '.'));
      let vo1 = (m[4] || '').trim();
      if (!vo1) vo1 = partFallbackSheath?.toLowerCase() || '';
      if (td1 <= 0) continue;

      const sheathType = resolveSheathType(vo1, rawPart, availableSheaths);
      addToDict(dict, {
        sectionMM2: td1,
        sheathRaw: vo1 || sheathType || 'unknown',
        sheathType,
        coreCount: 1,
        count: soCum1 * soNhan1,
      });
      matched1.push(m[0]);
    }
    for (const span of matched1) tmp = tmp.replace(span, ' ');

    // Nhóm 2: (n)x kC-size[vỏ], k>=2  -> cáp đa lõi
    const re2 =
      /(\d+)?\s*[x]?\s*\(?\s*(?:(\d+)\s*[x]\s*)?(\d+)\s*[cC]\s*[-_x×]?\s*(\d+(?:[.,]\d+)?)(?:\))?\s*([\w/\-]*)/gi;
    const matched2: string[] = [];
    while ((m = re2.exec(tmp)) !== null) {
      const outerCnt = m[1] ? parseInt(m[1], 10) : 1;
      const innerCnt = m[2] ? parseInt(m[2], 10) : 1;
      const soCore2 = parseInt(m[3], 10);
      const td2 = parseFloat(m[4].replace(',', '.'));
      let voBoc2 = (m[5] || '').trim();
      if (!voBoc2) voBoc2 = partFallbackSheath?.toLowerCase() || '';
      if (td2 <= 0 || soCore2 <= 0) continue;
      if (soCore2 === 1) continue; // 1C đã xử lý ở nhóm 1

      const sheathType = resolveSheathType(voBoc2, rawPart, availableSheaths);
      addToDict(dict, {
        sectionMM2: td2,
        sheathRaw: voBoc2 || sheathType || 'unknown',
        sheathType,
        coreCount: soCore2,
        count: outerCnt * innerCnt,
      });
      matched2.push(m[0]);
    }
    for (const span of matched2) tmp = tmp.replace(span, ' ');

    // Nhóm 3: Nxsize không có chữ C -> N=2 hiểu là cáp 2 lõi, khác 2 hiểu là N sợi 1 lõi
    const re3 = /(\d+)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*([\w/\-]*)/gi;
    while ((m = re3.exec(tmp)) !== null) {
      const count = parseInt(m[1], 10);
      const section = parseFloat(m[2].replace(',', '.'));
      let vo = (m[3] || '').trim();
      if (!vo) vo = partFallbackSheath?.toLowerCase() || '';
      if (section <= 0 || count <= 0) continue;

      const coreCount = count === 2 ? 2 : 1;
      const cableCount = coreCount === 2 ? 1 : count;
      const sheathType = resolveSheathType(vo, rawPart, availableSheaths);
      addToDict(dict, {
        sectionMM2: section,
        sheathRaw: vo || sheathType || 'unknown',
        sheathType,
        coreCount,
        count: cableCount,
      });
    }

    // Nhóm 3b: bắt các biến thể "1C-size" còn sót (vd "(1C-1.5)", "2x(1Cx1.5)")
    if (matched1.length === 0) {
      const re1b =
        /(?:(\d+)\s*[x×]\s*)?\(?\s*1\s*[cC]\s*[-_x×]\s*(\d+(?:[.,]\d+)?)\s*\)?\s*([\w/\-]*)/gi;
      while ((m = re1b.exec(tmp)) !== null) {
        const cnt1 = m[1] ? parseInt(m[1], 10) : 1;
        const td1 = parseFloat(m[2].replace(',', '.'));
        let vo1 = (m[3] || '').trim();
        if (!vo1) vo1 = partFallbackSheath?.toLowerCase() || '';
        if (td1 <= 0 || cnt1 <= 0) continue;
        const sheathType = resolveSheathType(vo1, rawPart, availableSheaths);
        addToDict(dict, {
          sectionMM2: td1,
          sheathRaw: vo1 || sheathType || 'unknown',
          sheathType,
          coreCount: 1,
          count: cnt1,
        });
      }
    }
  }

  return dict;
}

/** Merge 2 dict (thường dùng để gộp dây pha + dây PE khi tính diện tích ống) */
export function mergeCableDicts(
  a: Map<CableDictKey, CableDictEntry>,
  b: Map<CableDictKey, CableDictEntry>
): Map<CableDictKey, CableDictEntry> {
  const all = new Map<CableDictKey, CableDictEntry>();
  for (const [k, v] of a) all.set(k, { ...v });
  for (const [k, v] of b) {
    const existing = all.get(k);
    if (existing) existing.count += v.count;
    else all.set(k, { ...v });
  }
  return all;
}

// ----------------------------------------------------------------------------
// 3. THỐNG KÊ SỐ LƯỢNG CÁP / SỐ LÕI TRÊN 1 PHA (getCableCoreStats)
// ----------------------------------------------------------------------------

export interface CableCoreStats {
  /** Tổng số lõi/dây pha quy đổi = sum(count * coreCount) trên toàn bộ chuỗi */
  totalPhaseCores: number;
  /** Tổng số SỢI cáp loại 1 lõi (vd "3x1C-4" -> 3) */
  singleCoreCableCount: number;
  /** Tổng số SỢI cáp loại nhiều lõi (vd "2x4C-10" -> 2 sợi cáp 4C) */
  multiCoreCableCount: number;
  /** Số lõi lớn nhất từng gặp trong 1 sợi cáp (vd có cả 3C và 4C thì lấy 4) */
  maxSingleCableCoreCount: number;
}

/**
 * Phân tích chuỗi text tiết diện dây pha -> số lõi & số lượng cáp.
 * Dùng để kiểm tra: mạch 1 pha cần >=2 lõi (L+N), mạch 3 pha cần >=3 lõi (L1L2L3).
 */
export function getCableCoreStats(phaseCableText: string): CableCoreStats {
  const dict = parseTextToDict(phaseCableText);
  let totalPhaseCores = 0;
  let singleCoreCableCount = 0;
  let multiCoreCableCount = 0;
  let maxSingleCableCoreCount = 0;

  for (const entry of dict.values()) {
    const totalCoresForEntry = entry.count * entry.coreCount;
    totalPhaseCores += totalCoresForEntry;
    if (entry.coreCount === 1) {
      singleCoreCableCount += entry.count;
    } else {
      multiCoreCableCount += entry.count;
    }
    if (entry.coreCount > maxSingleCableCoreCount) {
      maxSingleCableCoreCount = entry.coreCount;
    }
  }

  // Fallback: nếu regex chính không bắt được gì nhưng chuỗi vẫn có nội dung,
  // thử bắt nhanh theo mẫu đơn giản "(n)C" / "(n)x...1C" để không trả về rỗng.
  if (totalPhaseCores === 0 && phaseCableText && phaseCableText.trim()) {
    const textUpper = phaseCableText.toUpperCase();
    const match1C = textUpper.match(/(\d+)\s*X?\s*1\s*C/);
    if (match1C) {
      const cnt = parseInt(match1C[1], 10);
      if (cnt > 0) {
        singleCoreCableCount = cnt;
        totalPhaseCores = cnt;
        maxSingleCableCoreCount = 1;
      }
    } else {
      const matchKC = textUpper.match(/(\d+)\s*C/);
      if (matchKC) {
        const cores = parseInt(matchKC[1], 10);
        if (cores > 0) {
          multiCoreCableCount = 1;
          totalPhaseCores = cores;
          maxSingleCableCoreCount = cores;
        }
      }
    }
  }

  return { totalPhaseCores, singleCoreCableCount, multiCoreCableCount, maxSingleCableCoreCount };
}

// ----------------------------------------------------------------------------
// 4. SỐ DÂY TỐI THIỂU CHO MẠCH 3 PHA: 3 DÂY (cân bằng) HAY 4 DÂY (cần trung tính)?
// ----------------------------------------------------------------------------
/**
 * Bài học thực tế (bug đã gặp): 1 tủ điện 3 pha (CB 3P) nhưng tải R/Y/B KHÔNG
 * bằng nhau (vd tủ tổng gộp nhiều tải 1 pha rải trên 3 pha để cân tải) — trường
 * hợp này dòng điện chạy về dây trung tính khác 0 -> BẮT BUỘC phải có dây N,
 * tức tối thiểu 4 dây (L1, L2, L3, N), không phải 3 dây như mạch 3 pha thuần
 * (vd động cơ 3 pha cân bằng, chỉ cần L1,L2,L3, không cần N).
 *
 * Trước khi có bản cập nhật này, thuật toán luôn coi mạch 3 pha chỉ cần >=3 dây,
 * nên trường hợp tủ 3 pha không cân tải mà chỉ đi 2 sợi cáp 1 lõi (thiếu cả dây
 * pha thứ 3 lẫn dây N) không bị phát hiện là lỗi.
 */

/**
 * Tải 3 pha có cân bằng không (R ≈ Y ≈ B, sai lệch trong ngưỡng tolerance mặc định 3%).
 * - true  (cân bằng, vd động cơ 3 pha thuần)         -> chỉ cần 3 dây (L1,L2,L3), KHÔNG cần N.
 * - false (không cân bằng, vd tủ tổng/lộ vào gộp      -> dòng trung tính khác 0,
 *          nhiều tải 1 pha rải trên 3 pha)               bắt buộc có dây N -> tối thiểu 4 dây.
 * Nếu 1 trong 3 giá trị <= 0 (không phải tải 3 pha đầy đủ) -> coi là KHÔNG cân bằng (an toàn hơn).
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

/**
 * Kiểm tra số lượng dây/lõi cáp pha cho 1 mạch/lộ 3 pha, trả về thông báo lỗi
 * (string) nếu thiếu dây, hoặc null nếu đạt.
 *
 * @param phaseCableText  chuỗi cáp pha gốc (chỉ dùng để hiển thị trong thông báo lỗi)
 * @param stats           kết quả getCableCoreStats(phaseCableText)
 * @param isBalanced      kết quả isBalancedThreePhaseLoad(rVal, yVal, bVal)
 */
export function checkThreePhaseCoreCount(
  phaseCableText: string,
  stats: CableCoreStats,
  isBalanced: boolean
): string | null {
  const requiredCores = isBalanced ? 3 : 4;
  const neutralNote = isBalanced
    ? ''
    : ' (tải giữa các pha R/Y/B không cân bằng → dòng trung tính khác 0, bắt buộc phải có dây trung tính N)';

  if (stats.singleCoreCableCount > 0 && stats.multiCoreCableCount === 0 && stats.singleCoreCableCount < requiredCores) {
    return `Sai số lượng cáp 1 lõi cho mạch 3 pha${neutralNote}: cần tối thiểu ${requiredCores} sợi cáp 1 lõi (${requiredCores}x1C) hoặc cáp đa lõi (${requiredCores}C trở lên). Hiện tại chỉ có ${stats.singleCoreCableCount} sợi 1C (${phaseCableText}).`;
  }
  if (stats.multiCoreCableCount > 0 && stats.singleCoreCableCount === 0 && stats.maxSingleCableCoreCount < requiredCores) {
    return `Sai loại/số lõi cáp cho mạch 3 pha${neutralNote}: cần cáp đa lõi tối thiểu ${requiredCores}C. Cáp ${stats.maxSingleCableCoreCount}C (${phaseCableText}) không đủ số lõi.`;
  }
  if (stats.totalPhaseCores > 0 && stats.totalPhaseCores < requiredCores) {
    return `Thiếu số lượng dây/lõi cáp pha cho mạch 3 pha${neutralNote}: cần tối thiểu ${requiredCores} dây, hiện tại chỉ ghi nhận ${stats.totalPhaseCores} lõi/sợi cáp (${phaseCableText}).`;
  }
  return null;
}

// ----------------------------------------------------------------------------
// 5. (TUỲ CHỌN) SO SÁNH TIẾT DIỆN CÁP THỰC TẾ VỚI TIẾT DIỆN YÊU CẦU
// ----------------------------------------------------------------------------

/**
 * Kiểm tra tiết diện cáp thực tế (cableText) có đạt yêu cầu tối thiểu (reqText) không.
 * Không bắt buộc khớp loại vỏ/chất liệu, chỉ so sánh trị số tiết diện lớn nhất.
 * VD: isCableSectionOK("3x1C-6 Cu/XLPE/PVC", "3x1C-4") -> true (6 >= 4)
 */
export function isCableSectionOK(cableText: string, reqText: string): boolean {
  if (!cableText || !reqText) return false;

  const cleanCable = cableText.trim().toUpperCase();
  const cleanReq = reqText.trim().toUpperCase();

  if (cleanCable.includes(cleanReq)) return true;

  const stripMaterials = (s: string) =>
    s
      .replace(/\b(CU|AL|XLPE|PVC|CXV|CV|CXL|LSZH|LSO|FR|RUBBER|MM2|SQMM|CABLE|CAP|DAY|DÂY)\b/g, '')
      .replace(/[\/\s]/g, ' ')
      .trim();

  const cMatClean = stripMaterials(cleanCable);
  const rMatClean = stripMaterials(cleanReq);
  if (cMatClean.includes(rMatClean) || rMatClean.includes(cMatClean)) return true;

  const getSectionValues = (str: string): number[] => {
    const s = str.replace(/,/g, '.');
    const matches = s.match(/(?:[xX*()\s+]|^)(\d+(?:\.\d+)?)/g);
    if (!matches) return [];
    const nums: number[] = [];
    for (const raw of matches) {
      const cleanNum = raw.replace(/[^0-9.]/g, '');
      const val = parseFloat(cleanNum);
      if (!isNaN(val) && val > 0) nums.push(val);
    }
    return nums;
  };

  const reqNums = getSectionValues(rMatClean);
  const cableNums = getSectionValues(cMatClean);
  if (reqNums.length === 0) return true;
  if (cableNums.length === 0) return false;

  return Math.max(...cableNums) >= Math.max(...reqNums);
}

// ----------------------------------------------------------------------------
// 6. VÍ DỤ SỬ DỤNG (chạy thử: `npx tsx cable-section-algorithm.ts` hoặc
//    xoá phần này nếu chỉ cần import các hàm ở trên vào dự án khác)
// ----------------------------------------------------------------------------

function demo() {
  const samples = [
    '3x1C-4 Cu/XLPE/PVC',
    '1x4C-2.5',
    '2x(3x1C-6) Cu/PVC',
    '4x4',
    '2x4',
    '4C-10 Cu/XLPE/PVC + 1C-4',
    '',
  ];

  for (const s of samples) {
    const dict = parseTextToDict(s);
    const stats = getCableCoreStats(s);
    console.log('----------------------------------------');
    console.log('Input          :', JSON.stringify(s));
    console.log('Parsed entries :', [...dict.values()]);
    console.log('Core stats     :', stats);
  }

  console.log('----------------------------------------');
  console.log('isCableSectionOK("3x1C-6", "3x1C-4") =', isCableSectionOK('3x1C-6', '3x1C-4'));
  console.log('isCableSectionOK("3x1C-4", "3x1C-6") =', isCableSectionOK('3x1C-4', '3x1C-6'));

  console.log('----------------------------------------');
  console.log('=== Kiểm tra số dây cho mạch/lộ 3 pha ===');

  // Case bug thực tế: tủ DB2-3F-COM — 3 pha (3P), tải KHÔNG cân bằng
  // (R=1.87, Y=1.61, B=1.87 kVA), cáp lộ vào chỉ 2x(1C-4.0) -> thiếu dây N.
  {
    const phaseText = '2x(1C-4.0) Cu/XLPE/PVC';
    const stats = getCableCoreStats(phaseText);
    const balanced = isBalancedThreePhaseLoad(1.87, 1.61, 1.87);
    const msg = checkThreePhaseCoreCount(phaseText, stats, balanced);
    console.log('DB2-3F-COM (không cân bằng, chỉ 2 sợi) ->', msg ?? 'OK (không nên xảy ra!)');
  }

  // Case đúng: cùng tải không cân bằng nhưng đã đủ 4 sợi 1 lõi (L1,L2,L3,N)
  {
    const phaseText = '4x1C-240mm2 Cu/XLPE/PVC';
    const stats = getCableCoreStats(phaseText);
    const balanced = isBalancedThreePhaseLoad(30.4, 26.96, 26.96);
    const msg = checkThreePhaseCoreCount(phaseText, stats, balanced);
    console.log('MSB-01 (không cân bằng, đủ 4 sợi) ->', msg ?? 'OK');
  }

  // Case đúng: tải 3 pha cân bằng (động cơ), chỉ cần 3 sợi/lõi
  {
    const phaseText = '3C-2.5 Cu/XLPE/PVC';
    const stats = getCableCoreStats(phaseText);
    const balanced = isBalancedThreePhaseLoad(22.0, 22.0, 22.0);
    const msg = checkThreePhaseCoreCount(phaseText, stats, balanced);
    console.log('Động cơ 3 pha cân bằng (3C) ->', msg ?? 'OK');
  }
}

// Chỉ chạy demo khi file này được thực thi trực tiếp (không chạy khi import)
if (typeof require !== 'undefined' && require.main === module) {
  demo();
}
