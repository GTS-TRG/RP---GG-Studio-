import * as XLSX from 'xlsx';
import { DEFAULT_CABLE_SPECS } from './specTable';
import { DEFAULT_CABLE_OUTER_DIAS, mapHeaderToSheathType } from './outerDiaTable';
import {
  CableOuterDiaRow,
  CableSheathType,
  CableSpecRow,
  ConduitSpec,
  PanelSheetData,
  ProjectConfig,
  RawCircuitRow,
} from '../types';
import { DEFAULT_CONDUITS, makeConduitSpec, sortConduits } from './conduitTable';
import {
  CbRatingItem,
  DEFAULT_CB_RATINGS,
  formatCbRatingLabel,
  sortCbRatings,
} from './cbRatingTable';
import {
  DEFAULT_CB_TYPES,
  DEFAULT_ISC_OPTIONS,
  DEFAULT_POLE_OPTIONS,
  formatCbTypeLabel,
  formatIscLabel,
  formatPoleLabel,
  SpecListItem,
  uniqueSpecList,
} from './cbOptionLists';
import { extractNumber, extractNumberFromCell, extractTextFromCell, isExcelErrorCell, getExcelErrorLabel, isSummaryRow, normalizePoleValue, cleanCableSectionOnly } from './helpers';
import { detectPanelLayout } from './panelLayout';
import type { LookupTables } from './lookupTables';

export interface ParsedWorkbook {
  sheets: PanelSheetData[];
  specs: CableSpecRow[];
  outerDias: CableOuterDiaRow[];
  conduits: ConduitSpec[];
  /** Danh sách In (CB_Rating) từ Spec. Cable cột E */
  cbRatings: CbRatingItem[];
  cbTypes: SpecListItem[];
  poleOptions: SpecListItem[];
  iscOptions: SpecListItem[];
  specSheetName: string | null;
  rawWorkbook: XLSX.WorkBook;
  /** Số sheet bị bỏ qua vì đang ẩn (Hidden / VeryHidden) */
  skippedHiddenSheets?: string[];
}

/**
 * SheetJS: Workbook.Sheets[i].Hidden
 * 0 / undefined = hiện, 1 = Hidden, 2 = VeryHidden
 */
function isSheetHidden(workbook: XLSX.WorkBook, sheetName: string): boolean {
  const metaList = workbook.Workbook?.Sheets;
  if (!metaList || !Array.isArray(metaList)) return false;

  const idx = workbook.SheetNames.indexOf(sheetName);
  if (idx < 0) return false;

  // Khớp theo index (thứ tự SheetNames ≈ Workbook.Sheets), fallback theo name
  const meta =
    (metaList[idx] && (metaList[idx].name === sheetName || !metaList[idx].name)
      ? metaList[idx]
      : undefined) ?? metaList.find((s) => s && s.name === sheetName);

  if (!meta) return false;
  return meta.Hidden === 1 || meta.Hidden === 2;
}

/**
 * Parses an Excel binary buffer into panel schedule structured data.
 * Bảng tra (CB/OD/ống/options) luôn lấy từ fallbackTables (app),
 * KHÔNG lấy từ sheet Spec. Cable trong file đang kiểm tra.
 */
export function parseExcelWorkbook(
  arrayBuffer: ArrayBuffer,
  config: ProjectConfig,
  /** Bảng tra của app (từ public/data/Spec. Cable.xlsx). Thiếu -> bảng trong code. */
  fallbackTables?: LookupTables
): ParsedWorkbook {
  // Uint8Array ổn định hơn ArrayBuffer thuần với SheetJS trên một số trình duyệt
  const data = new Uint8Array(arrayBuffer);
  if (data.byteLength < 64) {
    throw new Error('Excel file buffer is empty or too small.');
  }

  const workbook = XLSX.read(data, {
    type: 'array',
    cellFormula: true,
  });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('Excel workbook has no sheets.');
  }

  // Chỉ để nhận diện / bỏ qua sheet Spec trong file kiểm tra — không dùng số liệu
  const specSheetName = findSpecSheetName(workbook);

  // Review luôn dùng bảng tra app
  const tables = fallbackTables ?? {
    cableSpecs: DEFAULT_CABLE_SPECS,
    outerDias: DEFAULT_CABLE_OUTER_DIAS,
    conduits: DEFAULT_CONDUITS,
    cbRatings: DEFAULT_CB_RATINGS,
    cbTypes: DEFAULT_CB_TYPES,
    poleOptions: DEFAULT_POLE_OPTIONS,
    iscOptions: DEFAULT_ISC_OPTIONS,
  };

  const specs = [...tables.cableSpecs];
  const outerDias = [...tables.outerDias];
  const conduits = [...tables.conduits];
  const cbRatings = [...tables.cbRatings];
  const cbTypes = [...tables.cbTypes];
  const poleOptions = [...tables.poleOptions];
  const iscOptions = [...tables.iscOptions];

  const sheetsData: PanelSheetData[] = [];
  const skippedHiddenSheets: string[] = [];

  for (const sheetName of workbook.SheetNames) {
    // Bỏ qua sheet đang ẩn — không đưa vào thẩm tra
    if (isSheetHidden(workbook, sheetName)) {
      skippedHiddenSheets.push(sheetName);
      continue;
    }

    const upperName = sheetName.trim().toUpperCase();
    if (
      upperName === 'REPORT' ||
      upperName === 'REVIEW_REPORT' ||
      upperName === 'SINGLE LINE' ||
      (specSheetName && sheetName === specSheetName)
    ) {
      continue;
    }

    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const parsedSheet = parsePanelSheet(sheetName, ws, config);
    if (parsedSheet.circuits.length > 0) {
      sheetsData.push(parsedSheet);
    }
  }

  return {
    sheets: sheetsData,
    specs,
    outerDias,
    conduits,
    cbRatings,
    cbTypes,
    poleOptions,
    iscOptions,
    // Không gắn nguồn Spec từ file kiểm tra — app dùng Spec. Cable.xlsx riêng
    specSheetName: null,
    rawWorkbook: workbook,
    skippedHiddenSheets,
  };
}

/** Tìm tên sheet Spec. Cable trong workbook (ưu tiên sheet đang hiện) */
function findSpecSheetName(workbook: XLSX.WorkBook): string | null {
  for (const name of workbook.SheetNames) {
    if (isSheetHidden(workbook, name)) continue;
    const norm = name.replace(/\s+/g, '').toUpperCase();
    if (norm === 'SPECCABLE' || norm === 'SPEC.CABLE' || norm.includes('SPEC')) {
      return name;
    }
  }
  for (const name of workbook.SheetNames) {
    const norm = name.replace(/\s+/g, '').toUpperCase();
    if (norm === 'SPECCABLE' || norm === 'SPEC.CABLE' || norm.includes('SPEC')) {
      return name;
    }
  }
  return null;
}

/**
 * Đọc toàn bộ bảng tra từ workbook Spec. Cable.xlsx (file nguồn của app).
 * Hỗ trợ:
 *  - Format mới (nhiều sheet): CB_SPEC / CB_OPTIONS / OUTER_DIA / CONDUIT
 *  - Format cũ (1 sheet Spec. Cable với nhiều vùng cột)
 */
export function parseLookupTablesFromSpecBuffer(arrayBuffer: ArrayBuffer): {
  tables: LookupTables;
  sheetName: string | null;
  warnings: string[];
} {
  const data = new Uint8Array(arrayBuffer);
  if (data.byteLength < 64) {
    throw new Error('Spec. Cable.xlsx buffer is empty or too small.');
  }

  const workbook = XLSX.read(data, { type: 'array', cellFormula: true });
  if (!workbook.SheetNames?.length) {
    throw new Error('Spec. Cable.xlsx has no sheets.');
  }

  const warnings: string[] = [];
  const fallback = {
    cableSpecs: DEFAULT_CABLE_SPECS,
    outerDias: DEFAULT_CABLE_OUTER_DIAS,
    conduits: DEFAULT_CONDUITS,
    cbRatings: DEFAULT_CB_RATINGS,
    cbTypes: DEFAULT_CB_TYPES,
    poleOptions: DEFAULT_POLE_OPTIONS,
    iscOptions: DEFAULT_ISC_OPTIONS,
  };

  const namesUpper = workbook.SheetNames.map((n) => n.trim().toUpperCase());
  const isMulti =
    namesUpper.includes('CB_SPEC') ||
    namesUpper.includes('OUTER_DIA') ||
    namesUpper.includes('CONDUIT');

  let cableSpecs: CableSpecRow[] = [];
  let outerDias: CableOuterDiaRow[] = [];
  let conduits: ConduitSpec[] = [];
  let cbRatings: CbRatingItem[] = [];
  let cbTypes: SpecListItem[] = [];
  let poleOptions: SpecListItem[] = [];
  let iscOptions: SpecListItem[] = [];
  let sheetName: string | null = null;

  if (isMulti) {
    sheetName = 'CB_SPEC+OUTER_DIA+CONDUIT';
    const cbWs = findSheet(workbook, 'CB_SPEC');
    if (cbWs) cableSpecs = parseCbSpecSheet(cbWs, 'CB_SPEC');

    const odWs = findSheet(workbook, 'OUTER_DIA');
    if (odWs) outerDias = parseOuterDiaFlatSheet(odWs, 'OUTER_DIA');

    const conduitWs = findSheet(workbook, 'CONDUIT');
    if (conduitWs) conduits = parseConduitFlatSheet(conduitWs, 'CONDUIT');

    const optWs = findSheet(workbook, 'CB_OPTIONS');
    if (optWs) {
      cbRatings = parseCbRatingColumn(optWs, 'CB_OPTIONS');
      cbTypes = parseSpecListColumn(optWs, 'CB_OPTIONS', matchCbTypeHeader, formatCbTypeLabel);
      poleOptions = parseSpecListColumn(optWs, 'CB_OPTIONS', matchPoleHeader, formatPoleLabel);
      iscOptions = parseSpecListColumn(optWs, 'CB_OPTIONS', matchIscHeader, formatIscLabel);
    }
  } else {
    // Format cũ: một sheet Spec. Cable
    sheetName = findSpecSheetName(workbook) ?? workbook.SheetNames[0] ?? null;
    if (!sheetName || !workbook.Sheets[sheetName]) {
      throw new Error('Spec. Cable.xlsx: cannot find Spec sheet.');
    }
    const ws = workbook.Sheets[sheetName];
    cableSpecs = parseSpecSheet(ws, sheetName);
    outerDias = parseOuterDiaRegion(ws, sheetName);
    conduits = parseConduitRegion(ws, sheetName);
    cbRatings = parseCbRatingColumn(ws, sheetName);
    cbTypes = parseSpecListColumn(ws, sheetName, matchCbTypeHeader, formatCbTypeLabel);
    poleOptions = parseSpecListColumn(ws, sheetName, matchPoleHeader, formatPoleLabel);
    iscOptions = parseSpecListColumn(ws, sheetName, matchIscHeader, formatIscLabel);
  }

  if (cableSpecs.length === 0) {
    warnings.push('cableSpecs trống — dùng mặc định.');
    cableSpecs = [...fallback.cableSpecs];
  }
  if (outerDias.length === 0) {
    warnings.push('outerDias trống — dùng mặc định.');
    outerDias = [...fallback.outerDias];
  }
  if (conduits.length === 0) {
    warnings.push('conduits trống — dùng mặc định.');
    conduits = [...fallback.conduits];
  }
  if (cbRatings.length === 0) {
    warnings.push('cbRatings trống — dùng mặc định.');
    cbRatings = [...fallback.cbRatings];
  }
  if (cbTypes.length === 0) {
    warnings.push('cbTypes trống — dùng mặc định.');
    cbTypes = [...fallback.cbTypes];
  }
  if (poleOptions.length === 0) {
    warnings.push('poleOptions trống — dùng mặc định.');
    poleOptions = [...fallback.poleOptions];
  }
  if (iscOptions.length === 0) {
    warnings.push('iscOptions trống — dùng mặc định.');
    iscOptions = [...fallback.iscOptions];
  }

  return {
    sheetName,
    warnings,
    tables: {
      cableSpecs,
      outerDias,
      conduits,
      cbRatings,
      cbTypes,
      poleOptions,
      iscOptions,
    },
  };
}

/** Tìm sheet theo tên (không phân biệt hoa thường / khoảng trắng) */
function findSheet(workbook: XLSX.WorkBook, want: string): XLSX.WorkSheet | null {
  const norm = want.replace(/\s+/g, '').toUpperCase();
  const name = workbook.SheetNames.find((n) => n.replace(/\s+/g, '').toUpperCase() === norm);
  return name ? workbook.Sheets[name] : null;
}

/**
 * Sheet CB_SPEC: cbAmp | phaseMM2 | peMM2 (header hàng 1)
 */
function parseCbSpecSheet(ws: XLSX.WorkSheet, sheetName: string): CableSpecRow[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:C200');
  // Bỏ qua hàng header nếu cột A không phải số
  const specs: CableSpecRow[] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cellA = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const cellB = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    const cellC = ws[XLSX.utils.encode_cell({ r, c: 2 })];
    const cbVal = cellA ? extractNumber(cellA.v) : 0;
    if (cbVal <= 0) continue;
    specs.push({
      cbAmp: cbVal,
      phaseText: cellB?.v != null ? cleanCableSectionOnly(String(cellB.v)) : '',
      peText: cellC?.v != null ? cleanCableSectionOnly(String(cellC.v)) : '',
      excelRow: r,
      excelSheet: sheetName,
    });
  }
  return specs.sort((a, b) => a.cbAmp - b.cbAmp);
}

/**
 * Sheet OUTER_DIA: coreCount | sectionMM2 | CU/PVC | CU/XLPE/PVC | ...
 */
function parseOuterDiaFlatSheet(ws: XLSX.WorkSheet, sheetName: string): CableOuterDiaRow[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z200');
  // Header row 0
  const headerRow = range.s.r;
  const sheathCols: { col: number; type: string }[] = [];
  for (let c = range.s.c + 2; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    if (!cell || cell.v == null) continue;
    const mapped = mapHeaderToSheathType(String(cell.v));
    if (mapped) sheathCols.push({ col: c, type: mapped });
  }

  const rows: CableOuterDiaRow[] = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const coreCell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const secCell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    const coreCount = coreCell ? extractNumber(coreCell.v) : 0;
    const sectionMM2 = secCell ? extractNumber(secCell.v) : 0;
    if (coreCount <= 0 || sectionMM2 <= 0) continue;

    const odBySheath: Partial<Record<string, number>> = {};
    const excelSheathCols: Record<string, number> = {};
    let hasAny = false;
    for (const { col, type } of sheathCols) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
      if (!cell || cell.v == null || cell.v === '') continue;
      const od = extractNumber(cell.v);
      if (od > 0) {
        odBySheath[type] = Math.round(od * 10) / 10;
        excelSheathCols[type] = col;
        hasAny = true;
      }
    }
    if (!hasAny) continue;

    rows.push({
      coreCount,
      sectionMM2,
      odBySheath,
      excelSheet: sheetName,
      excelRow: r,
      excelSectionCol: 1,
      excelSheathCols,
    });
  }
  return rows;
}

/**
 * Sheet CONDUIT: label | material | outerDiaMM | wallThicknessMM | innerDiaMM | note
 */
function parseConduitFlatSheet(ws: XLSX.WorkSheet, sheetName: string): ConduitSpec[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:F100');
  const headerRow = range.s.r;
  const header: Record<string, number> = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    if (!cell || cell.v == null) continue;
    const key = String(cell.v).trim().toLowerCase().replace(/\s+/g, '');
    header[key] = c;
  }

  const colLabel = header['label'] ?? 0;
  const colMat = header['material'] ?? 1;
  const colOuter = header['outerdiamm'] ?? header['outerdia'] ?? 2;
  const colThick = header['wallthicknessmm'] ?? header['wallthickness'] ?? 3;
  const colInner = header['innerdiamm'] ?? header['innerdia'] ?? 4;
  const colNote = header['note'] ?? 5;

  const rows: ConduitSpec[] = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const outerCell = ws[XLSX.utils.encode_cell({ r, c: colOuter })];
    const innerCell = ws[XLSX.utils.encode_cell({ r, c: colInner })];
    if (!outerCell || outerCell.v == null || !innerCell || innerCell.v == null) continue;
    const outerDiaMM = extractNumber(outerCell.v);
    const innerDiaMM = extractNumber(innerCell.v);
    if (outerDiaMM <= 0 || innerDiaMM <= 0) continue;

    const labelCell = ws[XLSX.utils.encode_cell({ r, c: colLabel })];
    const matCell = ws[XLSX.utils.encode_cell({ r, c: colMat })];
    const thickCell = ws[XLSX.utils.encode_cell({ r, c: colThick })];
    const noteCell = ws[XLSX.utils.encode_cell({ r, c: colNote })];

    const matRaw = matCell?.v != null ? String(matCell.v).toUpperCase() : 'PVC';
    const material: ConduitSpec['material'] = matRaw.includes('HDPE') ? 'HDPE' : 'PVC';
    const label =
      labelCell?.v != null && String(labelCell.v).trim()
        ? String(labelCell.v).trim()
        : `D${Math.round(outerDiaMM)}`;
    const wallThicknessMM =
      thickCell?.v != null && thickCell.v !== '' ? extractNumber(thickCell.v) : undefined;

    rows.push(
      makeConduitSpec({
        label,
        material,
        outerDiaMM,
        wallThicknessMM: wallThicknessMM && wallThicknessMM > 0 ? wallThicknessMM : undefined,
        innerDiaMM,
        note: noteCell?.v != null ? String(noteCell.v) : undefined,
        excelSheet: sheetName,
        excelRow: r,
        excelOuterCol: colOuter,
        excelThickCol: colThick,
        excelInnerCol: colInner,
      })
    );
  }
  return sortConduits(rows);
}

/**
 * Parses the Spec. Cable worksheet
 */
function parseSpecSheet(ws: XLSX.WorkSheet, sheetName: string): CableSpecRow[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:C100');
  const specs: CableSpecRow[] = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const cellA = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const cellB = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    const cellC = ws[XLSX.utils.encode_cell({ r, c: 2 })];

    const cbVal = cellA ? extractNumber(cellA.v) : 0;
    const phVal = cellB && cellB.v ? cleanCableSectionOnly(String(cellB.v)) : '';
    const peVal = cellC && cellC.v ? cleanCableSectionOnly(String(cellC.v)) : '';

    if (cbVal > 0) {
      specs.push({
        cbAmp: cbVal,
        phaseText: phVal,
        peText: peVal,
        excelRow: r,
        excelSheet: sheetName,
      });
    }
  }

  return specs.sort((a, b) => a.cbAmp - b.cbAmp);
}

/**
 * Parse cột CB_Rating (thường cột E) trên Spec. Cable
 * Header: CB_Rating / CB RATING / IN RATING
 */
function parseCbRatingColumn(ws: XLSX.WorkSheet, sheetName: string): CbRatingItem[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z100');
  let headerRow = -1;
  let ratingCol = -1;

  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15); r++) {
    for (let c = 0; c <= Math.min(range.e.c, 20); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const text = String(cell.v).trim().toUpperCase().replace(/\s+/g, '_');
      if (
        text === 'CB_RATING' ||
        text === 'CBRATING' ||
        text.includes('CB_RATING') ||
        (text.includes('CB') && text.includes('RATING')) ||
        text === 'IN_RATING'
      ) {
        headerRow = r;
        ratingCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }

  // Fallback: cột E (index 4) nếu hàng 1 có dạng số/A
  if (headerRow < 0) {
    ratingCol = 4;
    headerRow = 0;
  }

  const items: CbRatingItem[] = [];
  const seen = new Set<number>();

  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: ratingCol })];
    if (!cell || cell.v == null || cell.v === '') continue;

    // Bỏ qua nếu ô là tiêu đề lặp
    const raw = String(cell.v).trim();
    if (/rating|tiết|tiet|section/i.test(raw) && !/\d/.test(raw)) continue;

    const amp = extractNumber(cell.v);
    if (amp <= 0 || seen.has(amp)) continue;
    seen.add(amp);

    items.push({
      label: formatCbRatingLabel(raw.includes('A') || raw.includes('a') ? raw : amp),
      amp: Math.round(amp),
      excelSheet: sheetName,
      excelRow: r,
      excelCol: ratingCol,
    });
  }

  return sortCbRatings(items);
}

function matchCbTypeHeader(norm: string): boolean {
  return (
    norm === 'CB_TYPE' ||
    norm === 'CBTYPE' ||
    norm === 'TYPE_CB' ||
    norm.includes('LOAI_CB') ||
    norm.includes('LOẠI_CB') ||
    (norm.includes('CB') && norm.includes('TYPE') && !norm.includes('RATING'))
  );
}

function matchPoleHeader(norm: string): boolean {
  return (
    norm === 'POLE' ||
    norm === 'POLES' ||
    norm === 'CB_POLE' ||
    norm === 'CB_POLES' ||
    norm.includes('SO_CUC') ||
    norm.includes('SỐ_CỰC') ||
    norm.includes('SOCUC')
  );
}

function matchIscHeader(norm: string): boolean {
  return (
    norm === 'ISC' ||
    norm === 'ISC_KA' ||
    norm === 'CB_ISC' ||
    norm.includes('ISC') ||
    norm.includes('BREAKING') ||
    norm.includes('DONG_CAT') ||
    norm.includes('DÒNG_CẮT')
  );
}

/**
 * Parse cột danh sách tùy chọn trên Spec. Cable theo header matcher
 */
function parseSpecListColumn(
  ws: XLSX.WorkSheet,
  sheetName: string,
  matchHeader: (normHeader: string) => boolean,
  normalizeLabel: (raw: string) => string
): SpecListItem[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z100');
  let headerRow = -1;
  let col = -1;

  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 20); r++) {
    for (let c = 0; c <= Math.min(range.e.c, 25); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const text = String(cell.v).trim().toUpperCase().replace(/\s+/g, '_');
      if (matchHeader(text)) {
        headerRow = r;
        col = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }

  if (headerRow < 0 || col < 0) return [];

  const items: SpecListItem[] = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
    if (!cell || cell.v == null || cell.v === '') continue;
    const raw = String(cell.v).trim();
    if (!raw) continue;
    // Bỏ dòng header lặp / tiêu đề
    if (matchHeader(raw.toUpperCase().replace(/\s+/g, '_'))) continue;

    const label = normalizeLabel(raw);
    if (!label) continue;
    items.push({
      label,
      excelSheet: sheetName,
      excelRow: r,
      excelCol: col,
    });
  }

  return uniqueSpecList(items, normalizeLabel);
}

/**
 * Parse vùng bảng đường kính ngoài cáp (cột P–U trở đi)
 * Header: TIẾT DIỆN | CU/PVC | CU/XLPE/PVC | ...
 * Nhóm số lõi: hàng có P = "1C" / "2C" / "3C" / "4C"...
 */
function parseOuterDiaRegion(ws: XLSX.WorkSheet, sheetName: string): CableOuterDiaRow[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z200');
  const rows: CableOuterDiaRow[] = [];

  // Tìm hàng header chứa "TIẾT DIỆN" trong cột P trở đi (c >= 15)
  let headerRow = -1;
  let sectionCol = -1;
  const sheathCols: { col: number; type: CableSheathType }[] = [];

  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 30); r++) {
    for (let c = 14; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const text = String(cell.v).trim().toUpperCase();
      if (text.includes('TIẾT DIỆN') || text.includes('TIET DIEN') || text === 'SECTION') {
        headerRow = r;
        sectionCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }

  if (headerRow < 0 || sectionCol < 0) {
    return [];
  }

  // Đọc header quy cách vỏ từ các cột bên phải cột tiết diện (hỗ trợ cột tùy chỉnh)
  for (let c = sectionCol + 1; c <= Math.min(sectionCol + 20, range.e.c); c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    if (!cell || cell.v == null) continue;
    const mapped = mapHeaderToSheathType(String(cell.v));
    if (mapped) {
      sheathCols.push({ col: c, type: mapped });
    }
  }

  if (sheathCols.length === 0) {
    return [];
  }

  const sheathColMap: Record<string, number> = {};
  for (const { col, type } of sheathCols) {
    sheathColMap[type] = col;
  }

  let currentCore = 0;

  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const cellP = ws[XLSX.utils.encode_cell({ r, c: sectionCol })];
    if (!cellP || cellP.v == null || cellP.v === '') continue;

    const raw = String(cellP.v).trim();
    const coreMatch = raw.toUpperCase().match(/^(\d+)\s*C$/i);
    if (coreMatch) {
      currentCore = parseInt(coreMatch[1], 10);
      continue;
    }

    const section = extractNumber(cellP.v);
    if (section <= 0 || currentCore <= 0) continue;

    const odBySheath: Partial<Record<string, number>> = {};
    let hasAny = false;
    for (const { col, type } of sheathCols) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
      if (!cell || cell.v == null || cell.v === '') continue;
      const od = extractNumber(cell.v);
      if (od > 0) {
        odBySheath[type] = Math.round(od * 10) / 10;
        hasAny = true;
      }
    }

    if (hasAny) {
      rows.push({
        coreCount: currentCore,
        sectionMM2: section,
        odBySheath,
        excelSheet: sheetName,
        excelRow: r,
        excelSectionCol: sectionCol,
        excelSheathCols: { ...sheathColMap },
      });
    }
  }

  return rows;
}

/**
 * Parse bảng ống luồn PVC (Spec. Cable cột AC–AF)
 * Header: STT | Đường kính ngoài | Độ dày ống | Đường kính trong
 * Label = D{outerDia}
 */
function parseConduitRegion(ws: XLSX.WorkSheet, sheetName: string): ConduitSpec[] {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:AZ100');
  let headerRow = -1;
  let outerCol = -1;
  let thickCol = -1;
  let innerCol = -1;

  const normHeader = (s: string) =>
    s
      .toUpperCase()
      .replace(/\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 40); r++) {
    let foundOuter = -1;
    let foundInner = -1;
    let foundThick = -1;
    for (let c = 20; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const text = normHeader(String(cell.v));
      if (text.includes('ĐƯỜNG KÍNH NGOÀI') || text.includes('DUONG KINH NGOAI') || text.includes('OUTER')) {
        foundOuter = c;
      }
      if (text.includes('ĐỘ DÀY') || text.includes('DO DAY') || text.includes('THICK')) {
        foundThick = c;
      }
      if (
        (text.includes('ĐƯỜNG KÍNH TRONG') || text.includes('DUONG KINH TRONG') || text.includes('INNER')) &&
        !text.includes('NGOÀI') &&
        !text.includes('NGOAI')
      ) {
        foundInner = c;
      }
    }
    // Chỉ nhận header thuộc khối ống PVC (có cả ngoài + trong gần nhau)
    if (foundOuter >= 0 && foundInner >= 0 && foundInner - foundOuter <= 3) {
      headerRow = r;
      outerCol = foundOuter;
      thickCol = foundThick;
      innerCol = foundInner;
      break;
    }
  }

  if (headerRow < 0 || outerCol < 0 || innerCol < 0) {
    return [];
  }

  // Xác định material từ hàng tiêu đề phía trên (PVC / HDPE)
  let material: ConduitSpec['material'] = 'PVC';
  for (let r = Math.max(range.s.r, headerRow - 3); r <= headerRow; r++) {
    for (let c = Math.max(0, outerCol - 2); c <= outerCol + 4; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const t = String(cell.v).toUpperCase();
      if (t.includes('HDPE')) material = 'HDPE';
      if (t.includes('PVC')) material = 'PVC';
    }
  }

  const rows: ConduitSpec[] = [];
  for (let r = headerRow + 1; r <= Math.min(headerRow + 30, range.e.r); r++) {
    const outerCell = ws[XLSX.utils.encode_cell({ r, c: outerCol })];
    const innerCell = ws[XLSX.utils.encode_cell({ r, c: innerCol })];
    if (!outerCell || outerCell.v == null || outerCell.v === '') continue;
    if (!innerCell || innerCell.v == null || innerCell.v === '') continue;

    const outerDia = extractNumber(outerCell.v);
    const innerDia = extractNumber(innerCell.v);
    if (outerDia <= 0 || innerDia <= 0) continue;

    let wallThickness: number | undefined;
    if (thickCol >= 0) {
      const thickCell = ws[XLSX.utils.encode_cell({ r, c: thickCol })];
      if (thickCell && thickCell.v != null && thickCell.v !== '') {
        const tw = extractNumber(thickCell.v);
        if (tw > 0) wallThickness = Math.round(tw * 100) / 100;
      }
    }

    rows.push(
      makeConduitSpec({
        label: `D${Math.round(outerDia)}`,
        material,
        outerDiaMM: Math.round(outerDia * 10) / 10,
        wallThicknessMM: wallThickness,
        innerDiaMM: Math.round(innerDia * 10) / 10,
        excelSheet: sheetName,
        excelRow: r,
        excelOuterCol: outerCol,
        excelThickCol: thickCol >= 0 ? thickCol : undefined,
        excelInnerCol: innerCol,
      })
    );
  }

  return sortConduits(rows);
}

/**
 * Nhận diện tủ MSB (Main Switchboard):
 * - Chỉ dựa vào TÊN SHEET có chứa "MSB".
 * Tủ DB/LP/EM/... (tên sheet không chứa MSB) -> tủ thường, không áp dụng Rule Isc >= 65kA
 */
function detectIsMsbPanel(sheetName: string): boolean {
  const nameUp = sheetName.toUpperCase().replace(/\s+/g, '');
  return nameUp.includes('MSB');
}

/**
 * Parses an individual Panel Schedule worksheet
 */
function parsePanelSheet(sheetName: string, ws: XLSX.WorkSheet, config: ProjectConfig): PanelSheetData {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z200');
  const isMSB = detectIsMsbPanel(sheetName);

  // Do bo cuc theo tieu de; khong dò được thì dùng bố cục cũ + config.startRow
  const layout = detectPanelLayout(ws, config.startRow);
  const cols = layout.cols;
  const startRowIdx = layout.dataStartRow;

  // Cot dinh danh dong: uu tien Ten mach, khong co thi dung Mo ta
  const keyCol = cols.lineName ?? cols.description ?? 0;

  // Find end row (summary row) - xet ca cot ten mach lan cot mo ta,
  // vi nhan dong tong co the nam o mot trong hai
  let endRowIndex = range.e.r;
  for (let r = startRowIdx; r <= range.e.r; r++) {
    const probeCols = [keyCol, cols.description].filter(
      (c): c is number => c !== undefined
    );
    const isEnd = probeCols.some((c) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return cell && cell.v && isSummaryRow(String(cell.v));
    });
    if (isEnd) {
      endRowIndex = r - 1;
      break;
    }
  }

  const circuits: RawCircuitRow[] = [];

  /** Doc o theo cot da do; cot khong ton tai -> undefined */
  const cellAt = (r: number, c: number | undefined) =>
    c === undefined ? undefined : ws[XLSX.utils.encode_cell({ r, c })];

  for (let r = startRowIdx; r <= endRowIndex; r++) {
    const keyCell = cellAt(r, keyCol);
    const keyText = keyCell && keyCell.v != null ? String(keyCell.v).trim() : '';

    if (!keyText || isSummaryRow(keyText)) {
      continue;
    }

    // Bo qua dong tieu de phu / ghi chu: phai co it nhat 1 thong so dien
    const hasElectricalData = [
      cols.cbType,
      cols.cbText,
      cols.poleVal,
      cols.iCalc,
      cols.phaseCableText,
    ].some((c) => {
      if (c === undefined) return false;
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return cell && cell.v != null && String(cell.v).trim() !== '';
    });
    if (!hasElectricalData) {
      continue;
    }

    const cellName = cellAt(r, cols.lineName);
    const cellDesc = cellAt(r, cols.description);
    const cellR = cellAt(r, cols.rLoad);
    const cellY = cellAt(r, cols.yLoad);
    const cellB = cellAt(r, cols.bLoad);
    const cellI = cellAt(r, cols.iCalc);
    const cellType = cellAt(r, cols.cbType);
    const cellPole = cellAt(r, cols.poleVal);
    const cellIn = cellAt(r, cols.cbText);
    const cellIsc = cellAt(r, cols.cbIsc);
    const cellPhase = cellAt(r, cols.phaseCableText);
    const cellPe = cellAt(r, cols.peCableText);
    const cellInstall = cellAt(r, cols.installMethod);
    const cellZ = ws[XLSX.utils.encode_cell({ r, c: 25 })]; // Col Z - Bypass (quy uoc app)

    const lineName = cellName && cellName.v != null ? String(cellName.v).trim() : '';
    const description = extractTextFromCell(cellDesc);
    const rLoad = extractNumberFromCell(cellR);
    const yLoad = extractNumberFromCell(cellY);
    const bLoad = extractNumberFromCell(cellB);
    const iCalc = extractNumberFromCell(cellI);
    const cbType = extractTextFromCell(cellType).toUpperCase();
    const poleVal = cellPole && !isExcelErrorCell(cellPole) ? normalizePoleValue(cellPole.v) : '';
    const cbText = extractTextFromCell(cellIn);
    const cbAmp = extractNumber(cbText);
    const cbIsc = extractTextFromCell(cellIsc);
    const iscAmp = extractNumber(cbIsc);
    const phaseCableText = extractTextFromCell(cellPhase);
    const peCableText = extractTextFromCell(cellPe);
    const installMethod = extractTextFromCell(cellInstall);
    const bypassText = extractTextFromCell(cellZ).toUpperCase();

    // Phat hien o loi Excel (#REF!, #N/A...) tren cac cot da do
    const errorFields: string[] = [];
    const checkErr = (cell: XLSX.CellObject | undefined, col: number | undefined) => {
      if (col === undefined || !cell) return;
      if (isExcelErrorCell(cell)) {
        errorFields.push(`${XLSX.utils.encode_col(col)}:${getExcelErrorLabel(cell)}`);
      }
    };
    checkErr(cellR, cols.rLoad);
    checkErr(cellY, cols.yLoad);
    checkErr(cellB, cols.bLoad);
    checkErr(cellI, cols.iCalc);
    checkErr(cellType, cols.cbType);
    checkErr(cellPole, cols.poleVal);
    checkErr(cellIn, cols.cbText);
    checkErr(cellIsc, cols.cbIsc);
    checkErr(cellPhase, cols.phaseCableText);
    checkErr(cellPe, cols.peCableText);
    checkErr(cellInstall, cols.installMethod);
    const hasExcelError = errorFields.length > 0;

    // Check cell comment tren o dinh danh
    let commentText = '';
    if (keyCell && (keyCell as any).c && Array.isArray((keyCell as any).c)) {
      commentText = (keyCell as any).c.map((c: any) => c.t || '').join(' ').toUpperCase();
    }

    // Sheet khong do duoc bo cuc (vd sheet tong hop/CSDL) rat de bi doc nham thanh mach.
    // Voi cac sheet do, doi hoi dong phai co CB that su: dung loai CB hoac co dong dinh muc.
    if (!layout.detected) {
      const looksLikeCb = /^(MCB|MCCB|ACB|RCCB|RCBO|RCD|ELCB|MPCB|FUSE)/.test(cbType);
      if (!looksLikeCb && cbAmp <= 0) {
        continue;
      }
    }

    const isBypassed =
      commentText.includes('OK') ||
      commentText.includes('IGNORE') ||
      bypassText.includes('OK') ||
      bypassText.includes('IGNORE');

    // Formulas if available
    const formulaLoad = cellR?.f ? `=${cellR.f}` : cellY?.f ? `=${cellY.f}` : cellB?.f ? `=${cellB.f}` : undefined;
    const formulaCB = cellIn?.f ? `=${cellIn.f}` : undefined;

    circuits.push({
      rowIndex: r + 1, // 1-indexed for row presentation
      lineName,
      description,
      rLoad,
      yLoad,
      bLoad,
      iCalc,
      cbType,
      poleVal,
      cbText,
      cbAmp,
      cbIsc,
      iscAmp,
      phaseCableText,
      peCableText,
      installMethod,
      formulaLoad,
      formulaCB,
      commentText,
      bypassText,
      isBypassed,
      hasExcelError,
      excelErrorFields: hasExcelError ? errorFields : undefined,
    });
  }

  return {
    sheetName,
    isMSB,
    startRow: startRowIdx + 1,
    endRow: endRowIndex + 1,
    circuits,
    cols,
    layoutDetected: layout.detected,
  };
}
