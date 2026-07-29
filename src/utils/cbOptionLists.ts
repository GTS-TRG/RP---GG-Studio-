/**
 * Danh sách tùy chọn Loại CB / Số cực / Isc — nguồn Spec. Cable
 * Dùng làm dropdown trên Panel Schedule (giống CB_Rating / In).
 */

export interface SpecListItem {
  /** Text hiển thị / ghi Excel */
  label: string;
  excelSheet?: string;
  excelRow?: number;
  excelCol?: number;
}

/** Loại CB mặc định */
export const DEFAULT_CB_TYPES: SpecListItem[] = [
  'MCB',
  'MCCB',
  'ACB',
  'RCCB',
  'RCBO',
  'FUSE',
].map((label) => ({ label }));

/** Số cực mặc định */
export const DEFAULT_POLE_OPTIONS: SpecListItem[] = [
  '1P',
  '1P+N',
  '2P',
  '3P',
  '3P+N',
  '4P',
].map((label) => ({ label }));

/** Isc (kA) mặc định */
export const DEFAULT_ISC_OPTIONS: SpecListItem[] = [
  '6kA',
  '10kA',
  '15kA',
  '25kA',
  '35kA',
  '50kA',
  '65kA',
  '70kA',
  '85kA',
  '100kA',
].map((label) => ({ label }));

/** Chuẩn hóa loại CB: "rcbo (30ma)" → "RCBO(30mA)" — giữ đơn vị mA dễ đọc */
export function formatCbTypeLabel(raw: string | number): string {
  const s = String(raw ?? '').trim();
  if (!s) return '';
  let out = s.replace(/\s+/g, ' ');
  // Gọn khoảng trắng quanh ngoặc: "RCBO ( 30mA )" → "RCBO(30mA)"
  out = out.replace(/\s*\(\s*/g, '(').replace(/\s*\)\s*/g, ')');
  // Viết hoa phần tên loại (MCB/MCCB/RCBO...)
  out = out.replace(/^([A-Za-z]+)/, (m) => m.toUpperCase());
  // Chuẩn hóa dòng rò: (30MA)/(30 ma) → (30mA)
  out = out.replace(/\((\d+(?:[.,]\d+)?)\s*(?:MA|ma|mA)?\)/g, '($1mA)');
  return out;
}

/**
 * Chuẩn hóa số cực: "3p" → "3P", "1PN" / "1P N" → "1P+N"
 */
export function formatPoleLabel(raw: string | number): string {
  let s = String(raw ?? '').trim().toUpperCase().replace(/\s+/g, '');
  if (!s) return '';
  // 1PN / 3PN → 1P+N / 3P+N
  s = s.replace(/(\d+)PN$/i, '$1P+N');
  s = s.replace(/\+\+/g, '+');
  return s;
}

/** Chuẩn hóa Isc: "65" / "65 kA" / "65KA" → "65kA" */
export function formatIscLabel(raw: string | number): string {
  if (typeof raw === 'number' && raw > 0) {
    const n = Number.isInteger(raw) ? String(raw) : String(raw).replace('.', ',');
    return `${n}kA`;
  }
  const s = String(raw ?? '').trim();
  if (!s) return '';
  const m = s.match(/(\d+(?:[.,]\d+)?)/);
  if (!m) return s.replace(/\s+/g, '');
  const num = m[1].replace('.', ',');
  return `${num}kA`;
}

/** Gộp unique theo nhãn đã chuẩn hóa — giữ item đầu tiên */
export function uniqueSpecList(
  items: SpecListItem[],
  normalize: (raw: string) => string = (x) => x.trim()
): SpecListItem[] {
  const seen = new Set<string>();
  const out: SpecListItem[] = [];
  for (const item of items) {
    const label = normalize(item.label);
    if (!label) continue;
    const key = label.toUpperCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ ...item, label });
  }
  return out;
}

/**
 * Thêm giá trị hiện tại vào đầu list nếu chưa có
 * (tránh mất giá trị khi Excel có text ngoài danh sách Spec)
 */
export function withCurrentOption(
  options: string[],
  current: string | number | undefined | null,
  normalize: (raw: string | number) => string
): string[] {
  const cur = normalize(current ?? '');
  if (!cur) return options;
  if (options.some((o) => normalize(o) === cur)) return options;
  return [cur, ...options];
}
