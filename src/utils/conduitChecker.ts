/**
 * Conduit Size Checker — bám sát VBA KiemTraOng3 / KIEMTRAONGMOD
 *
 * 1) Parse text cáp pha + PE → dict key = "td|vo|core"
 * 2) Tra Ø ngoài cáp từ bảng DKC (outerDias)
 * 3) Tính tổng diện tích cáp
 * 4) Ø ống tối thiểu = √(4 × A_cáp / tileMax / π)
 * 5) Chọn ống từ DKO (cột Ø trong), tính % lấp đầy
 * 6) soLuongOng >= 2: chia bó pha 3/4, phân phối, chọn ống từng ống
 */

import { CableOuterDiaRow, CableSheathType, ConduitSpec } from '../types';
import {
  DEFAULT_CABLE_OUTER_DIAS,
  collectSheathColumns,
  detectSheathType,
  lookupOuterDia,
  normalizeSheathKey,
  formatSheathColumnName,
} from './outerDiaTable';
import { DEFAULT_CONDUITS, findConduit, STANDARD_CONDUITS } from './conduitTable';
import { CABLE_SHEATH_TYPES } from '../types';

export type { ConduitSpec };
export { STANDARD_CONDUITS, DEFAULT_CONDUITS };

const PI = Math.PI;
const EPS = 0.0001;

/** Key VBA: "tiết_diện|vỏ|số_lõi" */
export type CableDictKey = string;

export interface CableDictEntry {
  sectionMM2: number;
  sheathRaw: string;
  sheathType?: CableSheathType;
  coreCount: number;
  count: number;
}

export interface ParsedCableItem {
  count: number;
  coreCount: number;
  sectionMM2: number;
  outerDiaMM: number;
  unitAreaMM2: number;
  totalAreaMM2: number;
  sheathType?: CableSheathType;
}

export interface ConduitCheckResult {
  isConduitUsed: boolean;
  specifiedLabel?: string;
  specifiedCount?: number;
  actualFillRatio?: number;
  maxFillRatioAllowed: number;
  reqConduitLabel?: string;
  /** Chuỗi đề xuất kiểu VBA: "4x4  @D20 (28.5%)" */
  recommendation?: string;
  /** Diện tích tổng cáp (mm²) */
  totalCableAreaMM2?: number;
  /** Ø trong ống tối thiểu (1 ống) */
  minInnerDiaMM?: number;
  isOverfilled: boolean;
  /** Ống đã chọn nhỏ hơn đề xuất (theo Ø trong) */
  isUndersized?: boolean;
  /** Không parse / không tra được bảng */
  hasError?: boolean;
  message?: string;
}

/**
 * Tra đường kính ngoài cáp (mm).
 * strict=true: không fallback công thức (giống VBA báo lỗi).
 */
export function getCableOuterDiameter(
  coreCount: number,
  sectionMM2: number,
  sheathType?: CableSheathType,
  odTable: CableOuterDiaRow[] = DEFAULT_CABLE_OUTER_DIAS,
  strict = false
): number | undefined {
  const fromTable = lookupOuterDia(odTable, coreCount, sectionMM2, sheathType, strict);
  if (fromTable != null && fromTable > 0) return fromTable;
  if (strict) return undefined;

  if (coreCount === 1) return 1.8 * Math.sqrt(sectionMM2) + 1.2;
  return 2.2 * Math.sqrt(coreCount * sectionMM2) + 3.0;
}

export function getCableArea(
  coreCount: number,
  sectionMM2: number,
  sheathType?: CableSheathType,
  odTable?: CableOuterDiaRow[]
): number {
  const outerDia = getCableOuterDiameter(coreCount, sectionMM2, sheathType, odTable) ?? 0;
  return PI * Math.pow(outerDia / 2, 2);
}

/** Chuẩn hóa tên vỏ về CableSheathType (hỗ trợ cả chuỗi VBA lowercase + cột tùy chỉnh) */
export function resolveSheathType(
  raw: string,
  fallbackText?: string,
  availableSheaths?: string[]
): CableSheathType | undefined {
  const pool = availableSheaths?.length
    ? availableSheaths
    : [...CABLE_SHEATH_TYPES];

  const cleaned = raw
    .replace(/\bmm2\b/gi, '')
    .replace(/\bsqmm\b/gi, '')
    .trim();

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

function makeDictKey(sectionMM2: number, sheathRaw: string, coreCount: number): CableDictKey {
  return `${sectionMM2}|${sheathRaw.toLowerCase()}|${coreCount}`;
}

function addToDict(dict: Map<CableDictKey, CableDictEntry>, entry: Omit<CableDictEntry, 'count'> & { count: number }) {
  const key = makeDictKey(entry.sectionMM2, entry.sheathRaw, entry.coreCount);
  const existing = dict.get(key);
  if (existing) {
    existing.count += entry.count;
  } else {
    dict.set(key, { ...entry });
  }
}

/**
 * Parse text cáp → dictionary (giống VBA ParseTextToDict).
 * Hỗ trợ material sau kích thước (VBA) và material trước (Excel thường gặp).
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

    // Bỏ mm2 để regex sạch hơn
    textPart = textPart.replace(/\s*(?:mm2|sqmm)\b/gi, '');
    const partFallbackSheath = detectSheathType(rawPart, availableSheaths);

    // Nhóm 1 VBA: (cum)?(n)x1C-size material
    const re1 = /(\d+)?x?\(?(\d+)\s*x\s*1\s*[cC]\s*[-_x\u00d7\s]\s*(\d+(?:[.,]\d+)?)(?:\))?\s*([\w/\-]*)/gi;
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

    // Nhóm 2 VBA: kC-size (bỏ qua 1C vì nhóm 1 đã xử lý)
    const re2 =
      /(\d+)?\s*[x]?\s*\(?\s*(?:(\d+)\s*[x]\s*)?(\d+)\s*[cC]\s*[-_x\u00d7]?\s*(\d+(?:[.,]\d+)?)(?:\))?\s*([\w/\-]*)/gi;
    const matched2: string[] = [];
    while ((m = re2.exec(tmp)) !== null) {
      const outerCnt = m[1] ? parseInt(m[1], 10) : 1;
      const innerCnt = m[2] ? parseInt(m[2], 10) : 1;
      const soCore2 = parseInt(m[3], 10);
      const td2 = parseFloat(m[4].replace(',', '.'));
      let voBoc2 = (m[5] || '').trim();
      if (!voBoc2) voBoc2 = partFallbackSheath?.toLowerCase() || '';
      if (td2 <= 0 || soCore2 <= 0) continue;
      if (soCore2 === 1) continue;

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

    // Nhóm 3: Nxsize không có chữ C
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

    // 1C đơn lẻ bị bỏ sót (vd "1c-4 cu/pvc")
    if (matched1.length === 0) {
      // Ho tro ca "1c-4", "(1C-1.5)", "2x(1Cx1.5)" - phai giu he so nhan truoc dau ngoac
      const re1b =
        /(?:(\d+)\s*[x\u00d7]\s*)?\(?\s*1\s*[cC]\s*[-_x\u00d7]\s*(\d+(?:[.,]\d+)?)\s*\)?\s*([\w/\-]*)/gi;
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

/** Merge hai dict (pha + PE) */
export function mergeCableDicts(
  phase: Map<CableDictKey, CableDictEntry>,
  pe: Map<CableDictKey, CableDictEntry>
): Map<CableDictKey, CableDictEntry> {
  const all = new Map<CableDictKey, CableDictEntry>();
  for (const [k, v] of phase) {
    all.set(k, { ...v });
  }
  for (const [k, v] of pe) {
    const existing = all.get(k);
    if (existing) existing.count += v.count;
    else all.set(k, { ...v });
  }
  return all;
}

/**
 * PickConduit2Col VBA: chọn ống đầu tiên có Ø trong >= requiredDiaMM
 */
export function pickConduitByInnerDia(
  conduits: ConduitSpec[],
  requiredDiaMM: number,
  eps = EPS
): { label: string; innerDiaMM: number; conduit: ConduitSpec } | null {
  const sorted = [...conduits].sort((a, b) => {
    if (a.material !== b.material) return a.material === 'PVC' ? -1 : 1;
    return a.innerDiaMM - b.innerDiaMM;
  });

  for (const c of sorted) {
    if (c.innerDiaMM > 0 && c.innerDiaMM + eps >= requiredDiaMM) {
      return { label: c.label || `DN${Math.round(c.innerDiaMM)}`, innerDiaMM: c.innerDiaMM, conduit: c };
    }
  }
  return null;
}

function formatSection(td: number): string {
  return Number.isInteger(td) || Math.abs(td - Math.round(td)) < EPS ? String(Math.round(td)) : String(td);
}

function formatCableDetail(count: number, coreCount: number, sectionMM2: number): string {
  const tdStr = formatSection(sectionMM2);
  if (coreCount > 1) return `${count}x${coreCount}C-${tdStr}`;
  return `${count}x${tdStr}`;
}

interface AreaCalcResult {
  totalAreaMM2: number;
  areaByKey: Map<CableDictKey, number>; // unit area per key
  detail: string;
  error?: string;
}

/**
 * Tính diện tích từng loại cáp từ dict + bảng OD (strict như VBA).
 */
function calcAreasFromDict(
  dictAll: Map<CableDictKey, CableDictEntry>,
  odTable: CableOuterDiaRow[]
): AreaCalcResult {
  if (dictAll.size === 0) {
    return { totalAreaMM2: 0, areaByKey: new Map(), detail: '', error: '#Cannot recognize cable from input' };
  }

  let tongDienTich = 0;
  const areaByKey = new Map<CableDictKey, number>();
  const detailParts: string[] = [];

  for (const [key, entry] of dictAll) {
    if (!entry.sheathType) {
      return {
        totalAreaMM2: 0,
        areaByKey,
        detail: '',
        error: `#Material not found: ${entry.sheathRaw}`,
      };
    }

    const dkCap = lookupOuterDia(odTable, entry.coreCount, entry.sectionMM2, entry.sheathType, true);
    if (dkCap == null || dkCap <= 0) {
      return {
        totalAreaMM2: 0,
        areaByKey,
        detail: '',
        error: `#Cable diameter not found: ${entry.sectionMM2} mm^2 for ${entry.coreCount}C - ${entry.sheathType}`,
      };
    }

    const s = PI * Math.pow(dkCap / 2, 2);
    areaByKey.set(key, s);
    tongDienTich += s * entry.count;
    detailParts.push(formatCableDetail(entry.count, entry.coreCount, entry.sectionMM2));
  }

  return {
    totalAreaMM2: tongDienTich,
    areaByKey,
    detail: detailParts.join(' + '),
  };
}

/** Ø ống tối thiểu từ diện tích cáp và % fill (công thức VBA) */
export function requiredInnerDiaMM(cableAreaMM2: number, tileMaxPercent: number): number {
  const dtToiDa = cableAreaMM2 / (tileMaxPercent / 100);
  return Math.sqrt((4 * dtToiDa) / PI);
}

interface PipeItem {
  key: CableDictKey;
  area: number;
  isPhase: boolean;
  count: number;
}

/**
 * Mở rộng dict thành items để chia nhiều ống (VBA expand_items).
 */
function expandItemsForMultiPipe(
  dictAll: Map<CableDictKey, CableDictEntry>,
  dictPhase: Map<CableDictKey, CableDictEntry>,
  areaByKey: Map<CableDictKey, number>
): PipeItem[] {
  const items: PipeItem[] = [];

  for (const [key, entry] of dictAll) {
    const Sitem = areaByKey.get(key) ?? 0;
    const phaseCount = dictPhase.get(key)?.count ?? 0;
    const peCount = entry.count - phaseCount;

    let groupSize = 1;
    if (phaseCount > 0) {
      if (phaseCount % 4 === 0) groupSize = 4;
      else if (phaseCount % 3 === 0) groupSize = 3;
      else groupSize = 1;
    }

    if (phaseCount > 0) {
      const nBundles = groupSize > 1 ? Math.floor(phaseCount / groupSize) : 0;
      const rest = phaseCount - nBundles * groupSize;
      for (let i = 0; i < nBundles; i++) {
        items.push({ key, area: Sitem * groupSize, isPhase: true, count: groupSize });
      }
      for (let i = 0; i < rest; i++) {
        items.push({ key, area: Sitem, isPhase: true, count: 1 });
      }
    }

    for (let i = 0; i < peCount; i++) {
      items.push({ key, area: Sitem, isPhase: false, count: 1 });
    }
  }

  // Sort giảm dần diện tích
  items.sort((a, b) => b.area - a.area);
  return items;
}

/**
 * Hàm chính tương đương VBA KiemTraOng3.
 */
export function kiemTraOng3(
  phaseText: string,
  peText: string,
  odTable: CableOuterDiaRow[] = DEFAULT_CABLE_OUTER_DIAS,
  conduits: ConduitSpec[] = DEFAULT_CONDUITS,
  tileMax = 35,
  soLuongOng = 1
): {
  ok: boolean;
  recommendation: string;
  totalCableAreaMM2: number;
  minInnerDiaMM?: number;
  reqLabel?: string;
  fillPercent?: number;
  pipeCount: number;
  error?: string;
} {
  const sheathPool = collectSheathColumns(odTable);
  const dictPhase = parseTextToDict(phaseText, sheathPool);
  const dictPE = parseTextToDict(peText, sheathPool);
  const dictAll = mergeCableDicts(dictPhase, dictPE);

  const areas = calcAreasFromDict(dictAll, odTable);
  if (areas.error) {
    return { ok: false, recommendation: areas.error, totalCableAreaMM2: 0, pipeCount: soLuongOng, error: areas.error };
  }

  const soOng = Math.max(1, soLuongOng);

  // ===== 1 ống =====
  if (soOng === 1) {
    const dkOngToiThieu = requiredInnerDiaMM(areas.totalAreaMM2, tileMax);
    const picked = pickConduitByInnerDia(conduits, dkOngToiThieu);
    if (!picked) {
      const err = `#No suitable conduit found, min required approx. ${dkOngToiThieu.toFixed(2)} mm`;
      return {
        ok: false,
        recommendation: err,
        totalCableAreaMM2: areas.totalAreaMM2,
        minInnerDiaMM: dkOngToiThieu,
        pipeCount: 1,
        error: err,
      };
    }

    const dienTichOng1 = PI * Math.pow(picked.innerDiaMM / 2, 2);
    const tileThucTe = Math.round((areas.totalAreaMM2 / dienTichOng1) * 1000) / 10; // 1 decimal like VBA Round(...,1)

    const recommendation = `${areas.detail}  @${picked.label} (${tileThucTe.toFixed(1)}%)`;
    return {
      ok: true,
      recommendation,
      totalCableAreaMM2: areas.totalAreaMM2,
      minInnerDiaMM: dkOngToiThieu,
      reqLabel: picked.label,
      fillPercent: tileThucTe,
      pipeCount: 1,
    };
  }

  // ===== Nhiều ống =====
  const items = expandItemsForMultiPipe(dictAll, dictPhase, areas.areaByKey);
  if (items.length === 0) {
    const err = '#No cable after expand';
    return { ok: false, recommendation: err, totalCableAreaMM2: areas.totalAreaMM2, pipeCount: soOng, error: err };
  }

  const usedArea = new Array(soOng).fill(0);
  const pipeHasPhase = new Array(soOng).fill(false);
  const pipeDict: Map<CableDictKey, number>[] = Array.from({ length: soOng }, () => new Map());

  for (const item of items) {
    let bestPipe = 0;

    if (item.isPhase) {
      for (let p = 0; p < soOng; p++) {
        if (p === 0 || usedArea[p] < usedArea[bestPipe]) bestPipe = p;
      }
    } else {
      // PE: ưu tiên ống đã có pha
      let found = false;
      for (let pass = 1; pass <= 2 && !found; pass++) {
        bestPipe = -1;
        for (let p = 0; p < soOng; p++) {
          if (pass === 1 && !pipeHasPhase[p]) continue;
          if (bestPipe < 0 || usedArea[p] < usedArea[bestPipe]) bestPipe = p;
        }
        if (bestPipe >= 0) found = true;
      }
      if (bestPipe < 0) bestPipe = 0;
    }

    usedArea[bestPipe] += item.area;
    if (item.isPhase) pipeHasPhase[bestPipe] = true;
    pipeDict[bestPipe].set(item.key, (pipeDict[bestPipe].get(item.key) || 0) + item.count);
  }

  const pipeLabel: string[] = [];
  const pipeFill: number[] = [];

  for (let p = 0; p < soOng; p++) {
    if (usedArea[p] > 0) {
      const dkMin = requiredInnerDiaMM(usedArea[p], tileMax);
      const picked = pickConduitByInnerDia(conduits, dkMin);
      if (!picked) {
        const err = `#No suitable conduit for pipe ${p + 1}, min required approx. ${dkMin.toFixed(2)} mm`;
        return {
          ok: false,
          recommendation: err,
          totalCableAreaMM2: areas.totalAreaMM2,
          pipeCount: soOng,
          error: err,
        };
      }
      const dienTichOngPipe = PI * Math.pow(picked.innerDiaMM / 2, 2);
      pipeLabel[p] = picked.label;
      pipeFill[p] = (usedArea[p] / dienTichOngPipe) * 100;
    } else {
      pipeLabel[p] = '';
      pipeFill[p] = 0;
    }
  }

  // Build: D25@30.1%[4x4] + D20@22.0%[1x4]
  const parts: string[] = [];
  let maxLabel = pipeLabel[0] || '';
  let maxInner = 0;

  for (let p = 0; p < soOng; p++) {
    if (pipeDict[p].size === 0) continue;
    const itemStrs: string[] = [];
    for (const [key, cnt] of pipeDict[p]) {
      const entry = dictAll.get(key);
      if (!entry) continue;
      itemStrs.push(formatCableDetail(cnt, entry.coreCount, entry.sectionMM2));
    }
    parts.push(`${pipeLabel[p]}@${pipeFill[p].toFixed(1)}%[${itemStrs.join(' + ')}]`);

    const c = findConduit(conduits, pipeLabel[p]);
    if (c && c.innerDiaMM >= maxInner) {
      maxInner = c.innerDiaMM;
      maxLabel = pipeLabel[p];
    }
  }

  const recommendation = parts.join(' + ');
  const worstFill = Math.max(...pipeFill.filter((f) => f > 0), 0);

  return {
    ok: true,
    recommendation,
    totalCableAreaMM2: areas.totalAreaMM2,
    reqLabel: soOng > 1 ? `${soOng}x${maxLabel}` : maxLabel,
    fillPercent: Math.round(worstFill * 10) / 10,
    pipeCount: soOng,
  };
}

/** Tương thích API cũ — parse thành ParsedCableItem[] */
export function parseCableString(
  text: string,
  odTable: CableOuterDiaRow[] = DEFAULT_CABLE_OUTER_DIAS
): ParsedCableItem[] {
  const dict = parseTextToDict(text, collectSheathColumns(odTable));
  const items: ParsedCableItem[] = [];
  for (const entry of dict.values()) {
    const od = getCableOuterDiameter(entry.coreCount, entry.sectionMM2, entry.sheathType, odTable);
    const outerDiaMM = od ?? 0;
    const unitArea = PI * Math.pow(outerDiaMM / 2, 2);
    items.push({
      count: entry.count,
      coreCount: entry.coreCount,
      sectionMM2: entry.sectionMM2,
      outerDiaMM,
      unitAreaMM2: unitArea,
      totalAreaMM2: entry.count * unitArea,
      sheathType: entry.sheathType,
    });
  }
  return items;
}

export function calculateTotalCableArea(
  phaseText: string,
  peText: string,
  odTable: CableOuterDiaRow[] = DEFAULT_CABLE_OUTER_DIAS
): { totalAreaMM2: number; items: ParsedCableItem[] } {
  const phaseItems = parseCableString(phaseText, odTable);
  const peItems = parseCableString(peText, odTable);
  const items = [...phaseItems, ...peItems];
  return { totalAreaMM2: items.reduce((s, i) => s + i.totalAreaMM2, 0), items };
}

export function pickSuitableConduit(
  totalCableAreaMM2: number,
  maxFillRatio: number = 0.35,
  pipeCount: number = 1,
  conduits: ConduitSpec[] = DEFAULT_CONDUITS
): { conduit: ConduitSpec; fillRatio: number } | null {
  if (totalCableAreaMM2 <= 0) return null;
  const areaPerPipe = totalCableAreaMM2 / Math.max(1, pipeCount);
  const dkMin = requiredInnerDiaMM(areaPerPipe, maxFillRatio * 100);
  const picked = pickConduitByInnerDia(conduits, dkMin);
  if (!picked) return null;
  const totalConduitArea = picked.conduit.areaMM2 * Math.max(1, pipeCount);
  return {
    conduit: picked.conduit,
    fillRatio: (totalCableAreaMM2 / totalConduitArea) * 100,
  };
}

export function parseInstalledConduit(
  installMethod?: string,
  conduits: ConduitSpec[] = DEFAULT_CONDUITS
): { count: number; label: string; conduitSpec?: ConduitSpec } | null {
  if (!installMethod) return null;

  const text = installMethod.toUpperCase().trim();

  if (text.includes('TRAY') || text.includes('MÁNG') || text.includes('LADDER') || text.includes('TRUNKING')) {
    if (!text.includes('CONDUIT') && !text.includes('ỐNG')) {
      return null;
    }
  }

  const dMatch = text.match(/(?:D|DN)\s*(\d{2,3})/i);
  if (!dMatch) {
    if (text.includes('CONDUIT') || text.includes('ỐNG')) {
      return { count: 1, label: 'D20', conduitSpec: findConduit(conduits, 'D20') };
    }
    return null;
  }

  const diaNum = parseInt(dMatch[1], 10);
  const label = `D${diaNum}`;
  const preferPvc = text.includes('PVC') || !text.includes('HDPE');
  const conduitSpec =
    findConduit(conduits, label, preferPvc ? 'PVC' : 'HDPE') ||
    findConduit(conduits, diaNum) ||
    findConduit(conduits, label);

  let count = 1;
  const countMatch =
    text.match(/(\d+)\s*(?:CÁI\s*)?(?:ỐNG|CONDUIT)/i) ||
    text.match(/(?:IN|TRONG)\s*(\d+)\s*(?:ỐNG|CONDUIT)/i);
  if (countMatch && parseInt(countMatch[1], 10) > 0) {
    count = parseInt(countMatch[1], 10);
  }

  return { count, label, conduitSpec };
}

/**
 * Review rule: chạy KiemTraOng3 rồi so với ống đã chọn ở cột N (installMethod).
 */
export function evaluateConduitRule(
  phaseText: string,
  peText: string,
  installMethod?: string,
  maxFillRatioAllowed: number = 35,
  odTable: CableOuterDiaRow[] = DEFAULT_CABLE_OUTER_DIAS,
  conduits: ConduitSpec[] = DEFAULT_CONDUITS
): ConduitCheckResult | null {
  if (!phaseText && !peText) return null;

  const installed = parseInstalledConduit(installMethod, conduits);
  const pipeCount = installed?.count ?? 1;

  // Cable tray / không dùng ống → không check fill
  if (installMethod) {
    const t = installMethod.toUpperCase();
    if (
      (t.includes('TRAY') || t.includes('MÁNG') || t.includes('LADDER') || t.includes('TRUNKING')) &&
      !t.includes('CONDUIT') &&
      !t.includes('ỐNG')
    ) {
      return {
        isConduitUsed: false,
        maxFillRatioAllowed,
        isOverfilled: false,
      };
    }
  }

  const calc = kiemTraOng3(phaseText, peText, odTable, conduits, maxFillRatioAllowed, pipeCount);

  if (calc.error) {
    // Chỉ báo khi có vẻ dùng ống hoặc có text cáp rõ
    if (!installed && !phaseText) return null;
    return {
      isConduitUsed: !!installed,
      specifiedLabel: installed?.label,
      specifiedCount: installed?.count,
      maxFillRatioAllowed,
      recommendation: calc.recommendation,
      totalCableAreaMM2: calc.totalCableAreaMM2,
      isOverfilled: false,
      hasError: true,
      message: `Conduit check: ${calc.error}`,
    };
  }

  const reqConduitLabel = calc.reqLabel;
  const recommendation = calc.recommendation;

  if (!installed) {
    // Có cáp nhưng chưa ghi ống — cảnh báo đề xuất (warning phía reviewer)
    return {
      isConduitUsed: false,
      maxFillRatioAllowed,
      reqConduitLabel,
      recommendation,
      totalCableAreaMM2: calc.totalCableAreaMM2,
      minInnerDiaMM: calc.minInnerDiaMM,
      isOverfilled: false,
      message: `Missing conduit (Limit ${maxFillRatioAllowed}%->${reqConduitLabel || '?'})`,
    };
  }

  const conduitSpec = installed.conduitSpec || findConduit(conduits, installed.label);
  if (!conduitSpec) {
    return {
      isConduitUsed: true,
      specifiedLabel: installed.label,
      specifiedCount: installed.count,
      maxFillRatioAllowed,
      reqConduitLabel,
      recommendation,
      isOverfilled: false,
      hasError: true,
      message: `Unknown conduit ${installed.label} (Limit ${maxFillRatioAllowed}%->${reqConduitLabel || '?'})`,
    };
  }

  // So fill thực tế với ống đã chọn
  const totalConduitArea = conduitSpec.areaMM2 * installed.count;
  const actualFillRatio =
    totalConduitArea > 0 ? (calc.totalCableAreaMM2 / totalConduitArea) * 100 : 999;
  const isOverfilled = actualFillRatio > maxFillRatioAllowed + 0.1;

  // So kích thước: ống chọn nhỏ hơn ống đề xuất (theo Ø trong)
  const reqPicked = reqConduitLabel
    ? findConduit(conduits, String(reqConduitLabel).replace(/^\d+x/i, ''))
    : null;
  const isUndersized =
    !!reqPicked && conduitSpec.innerDiaMM + EPS < reqPicked.innerDiaMM;

  let message: string | undefined;
  if (isOverfilled || isUndersized) {
    // Thông báo ngắn: Conduit fill: 54,2% (Limit 35%->D32)
    const fillTxt = actualFillRatio.toFixed(1).replace('.', ',');
    const suggest = reqConduitLabel || recommendation || '?';
    message = `Conduit fill: ${fillTxt}% (Limit ${maxFillRatioAllowed}%->${suggest})`;
  }

  return {
    isConduitUsed: true,
    specifiedLabel: installed.label,
    specifiedCount: installed.count,
    actualFillRatio,
    maxFillRatioAllowed,
    reqConduitLabel,
    recommendation,
    totalCableAreaMM2: calc.totalCableAreaMM2,
    minInnerDiaMM: calc.minInnerDiaMM,
    isOverfilled,
    isUndersized,
    message,
  };
}
