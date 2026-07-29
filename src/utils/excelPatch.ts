/**
 * Patch từng ô Excel — chỉ ghi các vị trí đã sửa trên web app,
 * không rebuild lại toàn bộ sheet từ AOA.
 */

import * as XLSX from 'xlsx';

export interface CellPatch {
  sheetName: string;
  /** Hàng 0-based (SheetJS) */
  row: number;
  /** Cột 0-based (SheetJS) */
  col: number;
  value: string | number;
}

export function patchKey(p: Pick<CellPatch, 'sheetName' | 'row' | 'col'>): string {
  return `${p.sheetName}|${p.row}|${p.col}`;
}

/** Gộp patch: cùng ô thì giữ giá trị mới nhất */
export function upsertPatch(map: Map<string, CellPatch>, patch: CellPatch): Map<string, CellPatch> {
  const next = new Map(map);
  next.set(patchKey(patch), patch);
  return next;
}

/**
 * Áp danh sách patch vào workbook gốc (in-place trên object Sheets).
 * Chỉ đụng đúng ô trong patch — giữ nguyên sheet / ô khác.
 */
export function applyCellPatches(workbook: XLSX.WorkBook, patches: CellPatch[]): number {
  let applied = 0;
  for (const p of patches) {
    const ws = workbook.Sheets[p.sheetName];
    if (!ws) continue;

    const addr = XLSX.utils.encode_cell({ r: p.row, c: p.col });

    // Ô sạch — tránh copy comment / error / hyperlink làm XLSX.write lỗi
    const nextCell: XLSX.CellObject =
      typeof p.value === 'number' && !Number.isNaN(p.value)
        ? { t: 'n', v: p.value }
        : { t: 's', v: String(p.value ?? '') };

    ws[addr] = nextCell;
    applied += 1;

    // Mở rộng !ref nếu ô nằm ngoài range hiện tại
    if (ws['!ref']) {
      const range = XLSX.utils.decode_range(ws['!ref']);
      if (p.row < range.s.r) range.s.r = p.row;
      if (p.col < range.s.c) range.s.c = p.col;
      if (p.row > range.e.r) range.e.r = p.row;
      if (p.col > range.e.c) range.e.c = p.col;
      ws['!ref'] = XLSX.utils.encode_range(range);
    }
  }
  return applied;
}

/**
 * Clone nông Sheets đủ để write an toàn (tránh mutate fail giữa chừng).
 * SheetJS WorkBook không có deepClone chuẩn — dùng write→read làm bản sạch.
 */
export function cloneWorkbookForWrite(workbook: XLSX.WorkBook): XLSX.WorkBook {
  try {
    const raw = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    return XLSX.read(raw, { type: 'array', cellFormula: true });
  } catch {
    // Nếu clone fail, ghi trực tiếp workbook gốc
    return workbook;
  }
}

/** Serialize workbook đã patch → ArrayBuffer */
export function workbookToArrayBuffer(workbook: XLSX.WorkBook): ArrayBuffer {
  // Không bật cellStyles — bản community SheetJS hay lỗi khi write styles
  const out = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
  }) as ArrayBuffer | Uint8Array;

  // type:'array' trả về ArrayBuffer — không phải Uint8Array.
  // Nếu gọi thẳng copy.set(ArrayBuffer) sẽ KHÔNG chép được byte nào
  // và tạo ra file toàn số 0 (Excel không mở được).
  const view = out instanceof Uint8Array ? out : new Uint8Array(out);

  // Copy sang buffer mới (tránh SharedArrayBuffer / offset lệch)
  const copy = new Uint8Array(view.byteLength);
  copy.set(view);
  return copy.buffer;
}

/** Tải file về máy (fallback khi không có File System Access) */
export function downloadWorkbookBuffer(buffer: ArrayBuffer, fileName: string): void {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName || 'UPDATED.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}

/** Map cột circuit panel (chữ → 0-based) */
export const PANEL_COL = {
  lineName: 0, // A
  description: 1, // B
  cbType: 7, // H
  poleVal: 8, // I
  cbText: 9, // J
  cbIsc: 10, // K
  phaseCableText: 11, // L
  peCableText: 12, // M
  installMethod: 13, // N
} as const;

export type PanelEditableField = keyof typeof PANEL_COL;
