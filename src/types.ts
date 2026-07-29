/**
 * Types and interfaces for the Electrical Panel Schedule Reviewer application
 */
import type { PanelCols } from './utils/panelLayout';

export interface CableSpecRow {
  cbAmp: number;
  phaseText: string;
  peText: string;
  /** Tọa độ Excel để ghi ngược (0-based row); thiếu = không patch được */
  excelRow?: number;
  excelSheet?: string;
}

/** Các quy cách vỏ cáp mặc định trong bảng tra đường kính ngoài */
export const CABLE_SHEATH_TYPES = [
  'CU/PVC',
  'CU/XLPE/PVC',
  'CU/MICA/XLPE/FR-PVC',
  'CU/MICA/XLPE/LSZH',
  'CU/PVC/PVC',
] as const;

/** Tên quy cách vỏ — gồm mặc định + cột tùy chỉnh người dùng thêm */
export type CableSheathType = string;

/**
 * Một dòng trong bảng đường kính tổng ngoài cáp
 * (theo tiết diện + số lõi + quy cách vỏ)
 */
export interface CableOuterDiaRow {
  coreCount: number; // 1, 2, 3, 4, 5...
  sectionMM2: number; // 1.5, 2.5, 4...
  /** OD (mm) theo từng loại vỏ — thiếu khóa = không có dữ liệu */
  odBySheath: Partial<Record<string, number>>;
  /** Tọa độ Excel để ghi ngược */
  excelSheet?: string;
  excelRow?: number;
  excelSectionCol?: number;
  excelSheathCols?: Record<string, number>;
}

/**
 * Quy cách ống luồn dây (PVC / HDPE) — dùng để tính % lấp đầy và chọn ống
 */
export interface ConduitSpec {
  label: string; // e.g. "D16", "D20"
  material: 'PVC' | 'HDPE';
  outerDiaMM: number; // Đường kính ngoài danh định (mm)
  wallThicknessMM?: number; // Độ dày ống (mm)
  innerDiaMM: number; // Đường kính trong (mm) — cơ sở tính diện tích hữu dụng
  areaMM2: number; // Diện tích trong (mm²) = π*(inner/2)^2
  note?: string;
  /** Tọa độ Excel để ghi ngược */
  excelSheet?: string;
  excelRow?: number;
  excelOuterCol?: number;
  excelThickCol?: number;
  excelInnerCol?: number;
}

export interface RawCircuitRow {
  rowIndex: number;
  lineName: string;       // Col A
  description: string;    // Col B
  rLoad: number;          // Col D
  yLoad: number;          // Col E
  bLoad: number;          // Col F
  iCalc: number;          // Col G
  cbType: string;         // Col H
  poleVal: string;        // Col I
  cbText: string;         // Col J
  cbAmp: number;
  cbIsc: string;          // Col K
  iscAmp: number;
  phaseCableText: string; // Col L
  peCableText: string;    // Col M
  installMethod?: string; // Col N (c: 13) - GIẢI PHÁP LẮP ĐẶT / INSTALLATION METHOD
  formulaLoad?: string;
  formulaCB?: string;
  commentText?: string;
  bypassText?: string;    // Col Z (26)
  isBypassed: boolean;
  /** Ô công thức lỗi Excel (#REF!, #N/A, ...) trên dòng mạch */
  hasExcelError?: boolean;
  excelErrorFields?: string[]; // e.g. ['G','J','L']
}

export interface PanelSheetData {
  sheetName: string;
  isMSB: boolean;
  startRow: number;
  endRow: number;
  circuits: RawCircuitRow[];
  /** Vi tri cot thuc te da do duoc tren sheet nay (0-based) */
  cols: PanelCols;
  /** true = do duoc theo tieu de; false = dung bo cuc mac dinh */
  layoutDetected: boolean;
}

export interface ReviewIssue {
  id: string;
  sheetName: string;
  lineName: string;
  shortDesc: string;
  lineNameWithDesc: string;
  description: string;
  isWarning: boolean;
  ruleCode: 'RULE_MSB_ISC_AMP' | 'RULE_ISC_TYPE' | 'RULE_SPARE' | 'RULE_POLE_PHASE' | 'RULE_RATING_LOAD' | 'RULE_MULTICORE_LAYERS' | 'RULE_PHASE_CABLE' | 'RULE_PE_CABLE' | 'RULE_CONDUIT_SIZE' | 'RULE_LINK_MISMATCH' | 'RULE_MISSING_CB' | 'RULE_EXCEL_REF';
  status: 'UNRESOLVED' | 'OK' | 'IGNORE' | 'APPROVED';
  remarks: string;
  rowIndex: number;
  isResolved: boolean;
}

export interface ProjectConfig {
  minIscMsb: number;      // Default 65 kA
  minAmpMsb: number;      // Default 32 A
  maxMcbIsc: number;      // Default 15 kA
  minMccbIsc: number;     // Default 18 kA
  startRow: number;        // Default 13
  safetyFactorNormal: number; // Default 1.25
  safetyFactorFire: number;   // Default 1.50
  /** Ti le lap day ong luon toi da cho phep (%) - Default 35 */
  conduitFillPercent: number;
}

export interface ReviewSummary {
  fileName: string;
  totalPanels: number;
  totalCircuits: number;
  errorCount: number;
  warningCount: number;
  resolvedCount: number;
  specCount: number;
}
