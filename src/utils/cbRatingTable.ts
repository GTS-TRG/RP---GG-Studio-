/**
 * Danh sách CB Rating (In) — nguồn Spec. Cable cột E (CB_Rating)
 * Dùng làm dropdown cột In trên Panel Schedule.
 */

export interface CbRatingItem {
  /** Text hiển thị / ghi Excel, VD: "16A" */
  label: string;
  /** Giá trị số ampe */
  amp: number;
  excelSheet?: string;
  excelRow?: number;
  excelCol?: number;
}

/** Danh sách mặc định khớp Spec. Cable — cột CB_Rating */
export const DEFAULT_CB_RATINGS: CbRatingItem[] = [
  '10A', '16A', '20A', '25A', '32A', '40A', '50A', '63A', '80A', '100A',
  '125A', '150A', '160A', '200A', '225A', '250A', '300A', '320A', '350A', '400A',
  '450A', '500A', '630A', '800A', '1000A', '1250A', '1500A', '1600A', '2000A',
  '2500A', '3200A', '4000A',
].map((label) => ({
  label,
  amp: parseInt(label, 10),
}));

/** Chuẩn hóa nhãn In: "16" / "16 A" / "16a" → "16A" */
export function formatCbRatingLabel(raw: string | number): string {
  if (typeof raw === 'number' && raw > 0) return `${Math.round(raw)}A`;
  const s = String(raw || '').trim();
  if (!s) return '';
  const m = s.match(/(\d+(?:[.,]\d+)?)/);
  if (!m) return s.toUpperCase().replace(/\s+/g, '');
  const n = Math.round(Number(m[1].replace(',', '.')));
  if (!n || Number.isNaN(n)) return s.toUpperCase().replace(/\s+/g, '');
  return `${n}A`;
}

export function sortCbRatings(list: CbRatingItem[]): CbRatingItem[] {
  return [...list].sort((a, b) => a.amp - b.amp);
}

/** Gộp danh sách — giữ label unique theo amp */
export function mergeCbRatings(...lists: CbRatingItem[][]): CbRatingItem[] {
  const map = new Map<number, CbRatingItem>();
  for (const list of lists) {
    for (const item of list) {
      if (item.amp <= 0) continue;
      const label = formatCbRatingLabel(item.label || item.amp);
      map.set(item.amp, {
        ...item,
        label,
        amp: item.amp,
      });
    }
  }
  return sortCbRatings(Array.from(map.values()));
}
