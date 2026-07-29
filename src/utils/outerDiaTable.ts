/**
 * Bảng tra đường kính tổng ngoài cáp (OD) theo tiết diện / số lõi / quy cách vỏ
 * Nguồn: vùng cột P–U sheet Spec. Cable
 */

import {
  CableOuterDiaRow,
  CableSheathType,
  CABLE_SHEATH_TYPES,
} from '../types';

/** Làm tròn OD về 1 chữ số thập phân */
function r1(n: number): number {
  return Math.round(n * 10) / 10;
}

function row(
  coreCount: number,
  sectionMM2: number,
  od: Partial<Record<string, number>>
): CableOuterDiaRow {
  const odBySheath: Partial<Record<string, number>> = {};
  for (const [k, v] of Object.entries(od)) {
    if (v !== undefined && v !== null && !Number.isNaN(v) && v > 0) {
      odBySheath[k] = r1(v);
    }
  }
  return { coreCount, sectionMM2, odBySheath };
}

/**
 * Bảng OD mặc định (TCVN / catalogue) — khớp Spec. Cable.xlsx
 */
export const DEFAULT_CABLE_OUTER_DIAS: CableOuterDiaRow[] = [
  // —— 1C ——
  row(1, 1.5, { 'CU/PVC': 3.2, 'CU/XLPE/PVC': 5.3, 'CU/MICA/XLPE/FR-PVC': 6.4, 'CU/MICA/XLPE/LSZH': 6.4 }),
  row(1, 2.5, { 'CU/PVC': 3.6, 'CU/XLPE/PVC': 5.7, 'CU/MICA/XLPE/FR-PVC': 6.9, 'CU/MICA/XLPE/LSZH': 6.9 }),
  row(1, 4, { 'CU/PVC': 4.6, 'CU/XLPE/PVC': 6.3, 'CU/MICA/XLPE/FR-PVC': 7.4, 'CU/MICA/XLPE/LSZH': 7.4 }),
  row(1, 6, { 'CU/PVC': 5.1, 'CU/XLPE/PVC': 6.8, 'CU/MICA/XLPE/FR-PVC': 8.0, 'CU/MICA/XLPE/LSZH': 8.0 }),
  row(1, 10, { 'CU/PVC': 6.1, 'CU/XLPE/PVC': 7.5, 'CU/MICA/XLPE/FR-PVC': 8.6, 'CU/MICA/XLPE/LSZH': 8.6 }),
  row(1, 16, { 'CU/PVC': 6.7, 'CU/XLPE/PVC': 8.4, 'CU/MICA/XLPE/FR-PVC': 9.5, 'CU/MICA/XLPE/LSZH': 9.5 }),
  row(1, 25, { 'CU/PVC': 8.2, 'CU/XLPE/PVC': 9.9, 'CU/MICA/XLPE/FR-PVC': 11.0, 'CU/MICA/XLPE/LSZH': 11.0 }),
  row(1, 35, { 'CU/PVC': 9.3, 'CU/XLPE/PVC': 11.0, 'CU/MICA/XLPE/FR-PVC': 12.1, 'CU/MICA/XLPE/LSZH': 12.1 }),
  row(1, 50, { 'CU/PVC': 10.8, 'CU/XLPE/PVC': 12.3, 'CU/MICA/XLPE/FR-PVC': 13.4, 'CU/MICA/XLPE/LSZH': 13.4 }),
  row(1, 70, { 'CU/PVC': 12.5, 'CU/XLPE/PVC': 14.2, 'CU/MICA/XLPE/FR-PVC': 15.3, 'CU/MICA/XLPE/LSZH': 15.3 }),
  row(1, 95, { 'CU/PVC': 14.5, 'CU/XLPE/PVC': 16.0, 'CU/MICA/XLPE/FR-PVC': 17.1, 'CU/MICA/XLPE/LSZH': 17.1 }),
  row(1, 120, { 'CU/PVC': 15.9, 'CU/XLPE/PVC': 17.6, 'CU/MICA/XLPE/FR-PVC': 18.7, 'CU/MICA/XLPE/LSZH': 18.7 }),
  row(1, 150, { 'CU/PVC': 17.7, 'CU/XLPE/PVC': 19.6, 'CU/MICA/XLPE/FR-PVC': 20.7, 'CU/MICA/XLPE/LSZH': 20.7 }),
  row(1, 185, { 'CU/PVC': 19.7, 'CU/XLPE/PVC': 21.6, 'CU/MICA/XLPE/FR-PVC': 22.7, 'CU/MICA/XLPE/LSZH': 22.7 }),
  row(1, 240, { 'CU/PVC': 22.4, 'CU/XLPE/PVC': 24.3, 'CU/MICA/XLPE/FR-PVC': 25.4, 'CU/MICA/XLPE/LSZH': 25.4 }),
  row(1, 300, { 'CU/PVC': 25.2, 'CU/XLPE/PVC': 27.0, 'CU/MICA/XLPE/FR-PVC': 28.2, 'CU/MICA/XLPE/LSZH': 28.2 }),
  row(1, 400, { 'CU/PVC': 28.4, 'CU/XLPE/PVC': 30.4, 'CU/MICA/XLPE/FR-PVC': 31.5, 'CU/MICA/XLPE/LSZH': 31.5 }),
  row(1, 500, { 'CU/XLPE/PVC': 34.0, 'CU/MICA/XLPE/FR-PVC': 35.0, 'CU/MICA/XLPE/LSZH': 35.0 }),
  row(1, 630, { 'CU/XLPE/PVC': 39.0, 'CU/MICA/XLPE/FR-PVC': 40.0, 'CU/MICA/XLPE/LSZH': 40.0 }),

  // —— 2C ——
  row(2, 1.5, { 'CU/XLPE/PVC': 10.2, 'CU/MICA/XLPE/FR-PVC': 12.4, 'CU/MICA/XLPE/LSZH': 12.4, 'CU/PVC/PVC': 10.6 }),
  row(2, 2.5, { 'CU/XLPE/PVC': 11.1, 'CU/MICA/XLPE/FR-PVC': 13.3, 'CU/MICA/XLPE/LSZH': 13.3, 'CU/PVC/PVC': 11.5 }),
  row(2, 4, { 'CU/XLPE/PVC': 12.1, 'CU/MICA/XLPE/FR-PVC': 14.4, 'CU/MICA/XLPE/LSZH': 14.4, 'CU/PVC/PVC': 13.3 }),
  row(2, 6, { 'CU/XLPE/PVC': 13.3, 'CU/MICA/XLPE/FR-PVC': 15.5, 'CU/MICA/XLPE/LSZH': 15.5, 'CU/PVC/PVC': 14.5 }),
  row(2, 10, { 'CU/XLPE/PVC': 13.7, 'CU/MICA/XLPE/FR-PVC': 16.0, 'CU/MICA/XLPE/LSZH': 16.0, 'CU/PVC/PVC': 14.9 }),
  row(2, 16, { 'CU/XLPE/PVC': 15.5, 'CU/MICA/XLPE/FR-PVC': 17.8, 'CU/MICA/XLPE/LSZH': 17.8, 'CU/PVC/PVC': 16.7 }),
  row(2, 25, { 'CU/XLPE/PVC': 18.6, 'CU/MICA/XLPE/FR-PVC': 20.9, 'CU/MICA/XLPE/LSZH': 20.9, 'CU/PVC/PVC': 19.8 }),
  row(2, 35, { 'CU/XLPE/PVC': 20.7, 'CU/MICA/XLPE/FR-PVC': 23.0, 'CU/MICA/XLPE/LSZH': 23.0, 'CU/PVC/PVC': 21.9 }),
  row(2, 50, { 'CU/XLPE/PVC': 23.4, 'CU/MICA/XLPE/FR-PVC': 25.7, 'CU/MICA/XLPE/LSZH': 25.7 }),
  row(2, 70, { 'CU/XLPE/PVC': 27.2, 'CU/MICA/XLPE/FR-PVC': 29.5, 'CU/MICA/XLPE/LSZH': 29.5 }),
  row(2, 95, { 'CU/XLPE/PVC': 30.8, 'CU/MICA/XLPE/FR-PVC': 33.0, 'CU/MICA/XLPE/LSZH': 33.0 }),
  row(2, 120, { 'CU/XLPE/PVC': 34.2, 'CU/MICA/XLPE/FR-PVC': 36.4, 'CU/MICA/XLPE/LSZH': 36.4 }),
  row(2, 150, { 'CU/XLPE/PVC': 38.0, 'CU/MICA/XLPE/FR-PVC': 40.3, 'CU/MICA/XLPE/LSZH': 40.3 }),
  row(2, 185, { 'CU/XLPE/PVC': 42.1, 'CU/MICA/XLPE/FR-PVC': 44.4, 'CU/MICA/XLPE/LSZH': 44.4 }),
  row(2, 240, { 'CU/XLPE/PVC': 48.0, 'CU/MICA/XLPE/FR-PVC': 50.2, 'CU/MICA/XLPE/LSZH': 50.2 }),
  row(2, 300, { 'CU/XLPE/PVC': 53.5, 'CU/MICA/XLPE/FR-PVC': 55.7, 'CU/MICA/XLPE/LSZH': 55.7 }),
  row(2, 400, { 'CU/XLPE/PVC': 60.2, 'CU/MICA/XLPE/FR-PVC': 62.5, 'CU/MICA/XLPE/LSZH': 62.5 }),

  // —— 3C ——
  row(3, 1.5, { 'CU/XLPE/PVC': 10.6, 'CU/MICA/XLPE/FR-PVC': 13.1, 'CU/MICA/XLPE/LSZH': 13.1, 'CU/PVC/PVC': 11.1 }),
  row(3, 2.5, { 'CU/XLPE/PVC': 11.6, 'CU/MICA/XLPE/FR-PVC': 14.0, 'CU/MICA/XLPE/LSZH': 14.0, 'CU/PVC/PVC': 12.0 }),
  row(3, 4, { 'CU/XLPE/PVC': 12.8, 'CU/MICA/XLPE/FR-PVC': 15.2, 'CU/MICA/XLPE/LSZH': 15.2, 'CU/PVC/PVC': 14.1 }),
  row(3, 6, { 'CU/XLPE/PVC': 14.0, 'CU/MICA/XLPE/FR-PVC': 16.4, 'CU/MICA/XLPE/LSZH': 16.4, 'CU/PVC/PVC': 15.3 }),
  row(3, 10, { 'CU/XLPE/PVC': 14.6, 'CU/MICA/XLPE/FR-PVC': 17.0, 'CU/MICA/XLPE/LSZH': 17.0, 'CU/PVC/PVC': 15.9 }),
  row(3, 16, { 'CU/XLPE/PVC': 16.5, 'CU/MICA/XLPE/FR-PVC': 18.9, 'CU/MICA/XLPE/LSZH': 18.9, 'CU/PVC/PVC': 17.8 }),
  row(3, 25, { 'CU/XLPE/PVC': 19.9, 'CU/MICA/XLPE/FR-PVC': 22.3, 'CU/MICA/XLPE/LSZH': 22.3, 'CU/PVC/PVC': 21.2 }),
  row(3, 35, { 'CU/XLPE/PVC': 22.1, 'CU/MICA/XLPE/FR-PVC': 24.5, 'CU/MICA/XLPE/LSZH': 24.5, 'CU/PVC/PVC': 23.4 }),
  row(3, 50, { 'CU/XLPE/PVC': 25.0, 'CU/MICA/XLPE/FR-PVC': 27.5, 'CU/MICA/XLPE/LSZH': 27.5 }),
  row(3, 70, { 'CU/XLPE/PVC': 29.3, 'CU/MICA/XLPE/FR-PVC': 31.7, 'CU/MICA/XLPE/LSZH': 31.7 }),
  row(3, 95, { 'CU/XLPE/PVC': 33.0, 'CU/MICA/XLPE/FR-PVC': 35.4, 'CU/MICA/XLPE/LSZH': 35.4 }),
  row(3, 120, { 'CU/XLPE/PVC': 36.6, 'CU/MICA/XLPE/FR-PVC': 39.0, 'CU/MICA/XLPE/LSZH': 39.0 }),
  row(3, 150, { 'CU/XLPE/PVC': 40.9, 'CU/MICA/XLPE/FR-PVC': 43.3, 'CU/MICA/XLPE/LSZH': 43.3 }),
  row(3, 185, { 'CU/XLPE/PVC': 45.7, 'CU/MICA/XLPE/FR-PVC': 48.2, 'CU/MICA/XLPE/LSZH': 48.2 }),
  row(3, 240, { 'CU/XLPE/PVC': 51.6, 'CU/MICA/XLPE/FR-PVC': 54.0, 'CU/MICA/XLPE/LSZH': 54.0 }),
  row(3, 300, { 'CU/XLPE/PVC': 57.5, 'CU/MICA/XLPE/FR-PVC': 59.9, 'CU/MICA/XLPE/LSZH': 59.9 }),
  row(3, 400, { 'CU/XLPE/PVC': 64.9, 'CU/MICA/XLPE/FR-PVC': 67.8, 'CU/MICA/XLPE/LSZH': 67.8 }),

  // —— 4C ——
  row(4, 1.5, { 'CU/XLPE/PVC': 11.4, 'CU/MICA/XLPE/FR-PVC': 14.1, 'CU/MICA/XLPE/LSZH': 14.1, 'CU/PVC/PVC': 11.9 }),
  row(4, 2.5, { 'CU/XLPE/PVC': 12.5, 'CU/MICA/XLPE/FR-PVC': 15.2, 'CU/MICA/XLPE/LSZH': 15.2, 'CU/PVC/PVC': 13.0 }),
  row(4, 4, { 'CU/XLPE/PVC': 13.8, 'CU/MICA/XLPE/FR-PVC': 16.5, 'CU/MICA/XLPE/LSZH': 16.5, 'CU/PVC/PVC': 15.3 }),
  row(4, 6, { 'CU/XLPE/PVC': 15.2, 'CU/MICA/XLPE/FR-PVC': 17.9, 'CU/MICA/XLPE/LSZH': 17.9, 'CU/PVC/PVC': 16.6 }),
  row(4, 10, { 'CU/XLPE/PVC': 15.9, 'CU/MICA/XLPE/FR-PVC': 18.6, 'CU/MICA/XLPE/LSZH': 18.6, 'CU/PVC/PVC': 17.4 }),
  row(4, 16, { 'CU/XLPE/PVC': 18.1, 'CU/MICA/XLPE/FR-PVC': 20.8, 'CU/MICA/XLPE/LSZH': 20.8, 'CU/PVC/PVC': 19.5 }),
  row(4, 25, { 'CU/XLPE/PVC': 21.8, 'CU/MICA/XLPE/FR-PVC': 24.5, 'CU/MICA/XLPE/LSZH': 24.5, 'CU/PVC/PVC': 23.3 }),
  row(4, 35, { 'CU/XLPE/PVC': 24.4, 'CU/MICA/XLPE/FR-PVC': 27.1, 'CU/MICA/XLPE/LSZH': 27.1, 'CU/PVC/PVC': 25.8 }),
  row(4, 50, { 'CU/XLPE/PVC': 27.8, 'CU/MICA/XLPE/FR-PVC': 30.5, 'CU/MICA/XLPE/LSZH': 30.5 }),
  row(4, 70, { 'CU/XLPE/PVC': 32.6, 'CU/MICA/XLPE/FR-PVC': 35.3, 'CU/MICA/XLPE/LSZH': 35.3 }),
  row(4, 95, { 'CU/XLPE/PVC': 36.7, 'CU/MICA/XLPE/FR-PVC': 39.4, 'CU/MICA/XLPE/LSZH': 39.4 }),
  row(4, 120, { 'CU/XLPE/PVC': 40.9, 'CU/MICA/XLPE/FR-PVC': 43.6, 'CU/MICA/XLPE/LSZH': 43.6 }),
  row(4, 150, { 'CU/XLPE/PVC': 45.9, 'CU/MICA/XLPE/FR-PVC': 48.6, 'CU/MICA/XLPE/LSZH': 48.6 }),
  row(4, 185, { 'CU/XLPE/PVC': 51.0, 'CU/MICA/XLPE/FR-PVC': 53.7, 'CU/MICA/XLPE/LSZH': 53.7 }),
  row(4, 240, { 'CU/XLPE/PVC': 57.5, 'CU/MICA/XLPE/FR-PVC': 60.2, 'CU/MICA/XLPE/LSZH': 60.2 }),
  row(4, 300, { 'CU/XLPE/PVC': 64.1, 'CU/MICA/XLPE/FR-PVC': 67.2, 'CU/MICA/XLPE/LSZH': 67.2 }),
  row(4, 400, { 'CU/XLPE/PVC': 72.8, 'CU/MICA/XLPE/FR-PVC': 75.5, 'CU/MICA/XLPE/LSZH': 75.5 }),
];

/** Chuẩn hóa tên quy cách vỏ để so khớp */
export function normalizeSheathKey(raw: string): string {
  return raw
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/MICA/g, 'MICA');
}

/** Chuẩn hóa tên cột vỏ khi người dùng / Excel thêm mới */
export function formatSheathColumnName(raw: string): string {
  return raw.trim().replace(/\s+/g, '').toUpperCase();
}

/**
 * Danh sách cột vỏ hiển thị: mặc định + khóa có trong bảng + cột extras (UI vừa thêm)
 */
export function collectSheathColumns(
  table: CableOuterDiaRow[],
  extras: string[] = []
): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  const add = (s: string) => {
    const key = formatSheathColumnName(s);
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(key);
  };
  for (const s of CABLE_SHEATH_TYPES) add(s);
  for (const r of table) {
    for (const k of Object.keys(r.odBySheath || {})) add(k);
  }
  for (const s of extras) add(s);
  return out;
}

/**
 * Nhận diện quy cách vỏ từ chuỗi mô tả cáp (VD: "Cu/XLPE/PVC 3C-2.5mm2")
 * Ưu tiên khớp loại cụ thể hơn trước; có thể bổ sung pool cột tùy chỉnh.
 */
export function detectSheathType(
  cableText: string,
  availableSheaths?: string[]
): CableSheathType | undefined {
  if (!cableText) return undefined;
  const t = normalizeSheathKey(cableText);

  // Khớp cột tùy chỉnh / đã biết theo độ dài tên giảm dần (tránh XLPE ăn mất MICA/XLPE)
  const pool = (availableSheaths?.length ? availableSheaths : [...CABLE_SHEATH_TYPES])
    .slice()
    .sort((a, b) => normalizeSheathKey(b).length - normalizeSheathKey(a).length);
  for (const s of pool) {
    const ns = normalizeSheathKey(s);
    if (ns.length >= 3 && t.includes(ns)) return formatSheathColumnName(s);
  }

  if (t.includes('MICA') && (t.includes('LSZH') || t.includes('LSOH'))) {
    return 'CU/MICA/XLPE/LSZH';
  }
  if (t.includes('MICA') && (t.includes('FR-PVC') || t.includes('FRPVC') || t.includes('FR/'))) {
    return 'CU/MICA/XLPE/FR-PVC';
  }
  if (t.includes('MICA')) {
    // Mica mặc định coi như FR-PVC nếu không ghi LSZH
    return 'CU/MICA/XLPE/FR-PVC';
  }
  if (t.includes('XLPE')) {
    return 'CU/XLPE/PVC';
  }
  if (t.includes('PVC/PVC') || t.includes('PVCPVC')) {
    return 'CU/PVC/PVC';
  }
  if (t.includes('PVC') || t.includes('CXV') || t.includes('/CV')) {
    return 'CU/PVC';
  }
  return undefined;
}

/** Thứ tự fallback khi thiếu OD đúng loại vỏ */
const SHEATH_FALLBACK: CableSheathType[] = [
  'CU/XLPE/PVC',
  'CU/PVC/PVC',
  'CU/PVC',
  'CU/MICA/XLPE/FR-PVC',
  'CU/MICA/XLPE/LSZH',
];

/**
 * Tra OD (mm) từ bảng. Trả về undefined nếu không có trong bảng.
 * @param strict nếu true và có sheathType: không fallback sang loại vỏ khác (giống VBA)
 */
export function lookupOuterDia(
  table: CableOuterDiaRow[],
  coreCount: number,
  sectionMM2: number,
  sheathType?: CableSheathType,
  strict = false
): number | undefined {
  const found = table.find(
    (r) => r.coreCount === coreCount && Math.abs(r.sectionMM2 - sectionMM2) < 0.05
  );
  if (!found) return undefined;

  if (sheathType) {
    const key = formatSheathColumnName(sheathType);
    if (found.odBySheath[key] != null) {
      return found.odBySheath[key];
    }
    // Thử khớp không phân biệt hoa thường với khóa trong hàng
    for (const [k, v] of Object.entries(found.odBySheath)) {
      if (normalizeSheathKey(k) === normalizeSheathKey(sheathType) && v != null) return v;
    }
    if (strict) return undefined;
  }

  // Fallback theo thứ tự ưu tiên
  for (const key of SHEATH_FALLBACK) {
    const v = found.odBySheath[key];
    if (v != null && v > 0) return v;
  }

  // Lấy giá trị đầu tiên còn lại
  for (const key of Object.keys(found.odBySheath)) {
    const v = found.odBySheath[key];
    if (v != null && v > 0) return v;
  }
  return undefined;
}

/** Nhóm bảng theo số lõi (để render UI) */
export function groupOuterDiaByCore(table: CableOuterDiaRow[]): Map<number, CableOuterDiaRow[]> {
  const map = new Map<number, CableOuterDiaRow[]>();
  for (const r of table) {
    const list = map.get(r.coreCount) || [];
    list.push(r);
    map.set(r.coreCount, list);
  }
  for (const [, list] of map) {
    list.sort((a, b) => a.sectionMM2 - b.sectionMM2);
  }
  return map;
}

/** Map tiêu đề cột Excel → tên quy cách vỏ (gồm cột tùy chỉnh) */
export function mapHeaderToSheathType(header: string): CableSheathType | undefined {
  const raw = String(header || '').trim();
  if (!raw) return undefined;
  const n = normalizeSheathKey(raw);

  // Bỏ header không phải cột quy cách vỏ (tiết diện / ống / nhãn phụ)
  if (
    !n ||
    n.includes('TIẾT') ||
    n.includes('TIET') ||
    n.includes('SECTION') ||
    n === 'MM2' ||
    n === 'STT' ||
    n === 'TEN' ||
    n.includes('ONG') ||
    n.includes('ỐNG') ||
    n.includes('DUONGKINH') ||
    n.includes('ĐƯỜNGKÍNH') ||
    n.includes('HDPE')
  ) {
    return undefined;
  }

  if (n.includes('MICA') && (n.includes('LSZH') || n.includes('LSOH'))) return 'CU/MICA/XLPE/LSZH';
  if (n.includes('MICA')) return 'CU/MICA/XLPE/FR-PVC';
  if (n.includes('XLPE')) return 'CU/XLPE/PVC';
  if (n.includes('PVC/PVC')) return 'CU/PVC/PVC';
  if (n === 'CU/PVC' || n === 'PVC') return 'CU/PVC';

  // Cột tùy chỉnh: chỉ nhận tên có dấu hiệu quy cách cáp
  if (n.includes('CU') || n.includes('XLPE') || n.includes('MICA') || n.includes('LSZH') || n.includes('LSOH')) {
    return formatSheathColumnName(raw);
  }
  return undefined;
}
