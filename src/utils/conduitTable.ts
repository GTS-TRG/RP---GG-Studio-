/**
 * Bảng quy cách ống luồn dây (đường kính trong) — nguồn Spec. Cable cột AC–AF (PVC)
 * Dùng làm cơ sở tính % lấp đầy và lựa chọn ống.
 */

import { ConduitSpec } from '../types';

/** Tính diện tích mặt cắt trong ống từ đường kính trong */
export function conduitInnerArea(innerDiaMM: number): number {
  return Math.PI * Math.pow(innerDiaMM / 2, 2);
}

/** Tạo dòng ống với area tự tính từ innerDia */
export function makeConduitSpec(
  partial: Omit<ConduitSpec, 'areaMM2'> & { areaMM2?: number }
): ConduitSpec {
  return {
    ...partial,
    areaMM2: partial.areaMM2 ?? conduitInnerArea(partial.innerDiaMM),
  };
}

/**
 * Bảng mặc định ống PVC — khớp Spec. Cable.xlsx (ỐNG LUỒN DÂY PVC)
 */
export const DEFAULT_CONDUITS: ConduitSpec[] = [
  makeConduitSpec({ label: 'D16', material: 'PVC', outerDiaMM: 16, wallThicknessMM: 1.2, innerDiaMM: 13.6 }),
  makeConduitSpec({ label: 'D20', material: 'PVC', outerDiaMM: 20, wallThicknessMM: 1.4, innerDiaMM: 17.2 }),
  makeConduitSpec({ label: 'D25', material: 'PVC', outerDiaMM: 25, wallThicknessMM: 1.5, innerDiaMM: 22.0 }),
  makeConduitSpec({ label: 'D32', material: 'PVC', outerDiaMM: 32, wallThicknessMM: 1.86, innerDiaMM: 28.3 }),
  makeConduitSpec({ label: 'D40', material: 'PVC', outerDiaMM: 40, wallThicknessMM: 2.1, innerDiaMM: 35.8 }),
  makeConduitSpec({ label: 'D50', material: 'PVC', outerDiaMM: 50, wallThicknessMM: 2.4, innerDiaMM: 45.2 }),
  makeConduitSpec({ label: 'D63', material: 'PVC', outerDiaMM: 63, wallThicknessMM: 3.0, innerDiaMM: 57.0 }),
];

/** Alias tương thích tên cũ */
export const STANDARD_CONDUITS = DEFAULT_CONDUITS;

/** Tìm ống theo nhãn Dxx hoặc đường kính ngoài */
export function findConduit(
  table: ConduitSpec[],
  labelOrOuter: string | number,
  material?: ConduitSpec['material']
): ConduitSpec | undefined {
  const list = material ? table.filter((c) => c.material === material) : table;
  if (typeof labelOrOuter === 'number') {
    return list.find((c) => c.outerDiaMM === labelOrOuter || c.label === `D${labelOrOuter}`);
  }
  const upper = labelOrOuter.toUpperCase().replace(/\s+/g, '');
  return list.find((c) => c.label.toUpperCase() === upper || `D${c.outerDiaMM}` === upper);
}

/** Sắp xếp ống tăng dần theo đường kính trong / diện tích */
export function sortConduits(table: ConduitSpec[]): ConduitSpec[] {
  return [...table].sort(
    (a, b) => a.material.localeCompare(b.material) || a.innerDiaMM - b.innerDiaMM
  );
}
