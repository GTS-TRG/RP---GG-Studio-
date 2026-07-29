/**
 * Bảng tra dùng chung — nạp từ public/data/Spec. Cable.xlsx
 *
 * Sửa file Excel đó rồi bấm "Tải lại bảng tra" (hoặc F5) để app cập nhật.
 * Review tủ điện luôn dùng bảng tra này, không lấy Spec từ file Excel đang kiểm tra.
 * Nếu thiếu file / parse lỗi -> dùng bảng mặc định trong code.
 */
import { CableOuterDiaRow, CableSpecRow, ConduitSpec } from '../types';
import { DEFAULT_CABLE_SPECS } from './specTable';
import { DEFAULT_CABLE_OUTER_DIAS } from './outerDiaTable';
import { DEFAULT_CONDUITS } from './conduitTable';
import { CbRatingItem, DEFAULT_CB_RATINGS } from './cbRatingTable';
import {
  DEFAULT_CB_TYPES,
  DEFAULT_ISC_OPTIONS,
  DEFAULT_POLE_OPTIONS,
  SpecListItem,
} from './cbOptionLists';

/** Đường dẫn file Spec. Cable.xlsx trong public/ (Vite serve từ gốc site) */
export const SPEC_CABLE_URL = 'data/Spec. Cable.xlsx';

export interface LookupTables {
  cableSpecs: CableSpecRow[];
  outerDias: CableOuterDiaRow[];
  conduits: ConduitSpec[];
  cbRatings: CbRatingItem[];
  cbTypes: SpecListItem[];
  poleOptions: SpecListItem[];
  iscOptions: SpecListItem[];
}

export interface LoadResult {
  tables: LookupTables;
  /** true = lấy từ Spec. Cable.xlsx; false = bảng mặc định trong code */
  fromFile: boolean;
  updatedAt?: string;
  /** Sheet đã đọc trong file Spec */
  sheetName?: string;
  /** Lý do phải quay về bảng mặc định / cảnh báo parse */
  error?: string;
}

/** Bảng mặc định biên dịch sẵn trong code — luôn dùng được */
export function builtInTables(): LookupTables {
  return {
    cableSpecs: [...DEFAULT_CABLE_SPECS],
    outerDias: [...DEFAULT_CABLE_OUTER_DIAS],
    conduits: [...DEFAULT_CONDUITS],
    cbRatings: [...DEFAULT_CB_RATINGS],
    cbTypes: [...DEFAULT_CB_TYPES],
    poleOptions: [...DEFAULT_POLE_OPTIONS],
    iscOptions: [...DEFAULT_ISC_OPTIONS],
  };
}

/**
 * Nạp bảng tra từ Spec. Cable.xlsx.
 * cacheBust=true để lấy bản mới nhất khi bấm "Tải lại bảng tra".
 */
export async function loadLookupTables(cacheBust = false): Promise<LoadResult> {
  const url = cacheBust
    ? `${SPEC_CABLE_URL}?t=${Date.now()}`
    : SPEC_CABLE_URL;

  try {
    const res = await fetch(url, { cache: cacheBust ? 'reload' : 'default' });
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }
    const buffer = await res.arrayBuffer();
    // Dynamic import tránh vòng phụ thuộc lookupTables <-> excelParser
    const { parseLookupTablesFromSpecBuffer } = await import('./excelParser');
    const { tables, sheetName, warnings } = parseLookupTablesFromSpecBuffer(buffer);
    if (warnings.length) {
      console.warn('[Spec. Cable.xlsx]', warnings.join(' | '));
    }
    return {
      tables,
      fromFile: true,
      sheetName: sheetName ?? undefined,
      updatedAt: new Date().toISOString().slice(0, 10),
      error: warnings.length ? warnings.join(' | ') : undefined,
    };
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    console.warn('[Spec. Cable.xlsx] Không nạp được file, dùng bảng mặc định:', detail);
    return { tables: builtInTables(), fromFile: false, error: detail };
  }
}

/** Áp bảng tra vào các state setter của App */
export function applyLookupTablesToState(
  tables: LookupTables,
  setters: {
    setSpecs: (v: CableSpecRow[]) => void;
    setOuterDias: (v: CableOuterDiaRow[]) => void;
    setConduits: (v: ConduitSpec[]) => void;
    setCbRatings: (v: CbRatingItem[]) => void;
    setCbTypes: (v: SpecListItem[]) => void;
    setPoleOptions: (v: SpecListItem[]) => void;
    setIscOptions: (v: SpecListItem[]) => void;
  }
): void {
  setters.setSpecs([...tables.cableSpecs]);
  setters.setOuterDias([...tables.outerDias]);
  setters.setConduits([...tables.conduits]);
  setters.setCbRatings([...tables.cbRatings]);
  setters.setCbTypes([...tables.cbTypes]);
  setters.setPoleOptions([...tables.poleOptions]);
  setters.setIscOptions([...tables.iscOptions]);
}
