/**
 * Sinh lại public/data/Spec. Cable.xlsx dạng gọn (nhiều sheet rõ ràng)
 * Chạy: npx tsx scripts/rebuild-spec-cable.mts
 */
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { parseLookupTablesFromSpecBuffer } from '../src/utils/excelParser';
import { extractNumber } from '../src/utils/helpers';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outPath = path.resolve(__dirname, '../public/data/Spec. Cable.xlsx');

// Đọc file hiện tại (hoặc bản cũ) làm nguồn số liệu
const oldBuf = fs.readFileSync(outPath);
const oldAb = oldBuf.buffer.slice(oldBuf.byteOffset, oldBuf.byteOffset + oldBuf.byteLength);
const parsed = parseLookupTablesFromSpecBuffer(oldAb);
const wbOld = XLSX.read(oldBuf, { type: 'buffer' });
const wsOld = wbOld.Sheets[wbOld.SheetNames.find((n) => /spec/i.test(n)) || wbOld.SheetNames[0]];

const cbSpec: { cbAmp: number; phaseMM2: string; peMM2: string }[] = [];
for (let r = 2; r <= 50; r++) {
  const a = wsOld[XLSX.utils.encode_cell({ r, c: 0 })];
  const b = wsOld[XLSX.utils.encode_cell({ r, c: 1 })];
  const c = wsOld[XLSX.utils.encode_cell({ r, c: 2 })];
  if (!a || a.v == null) continue;
  const amp = extractNumber(a.v);
  if (amp <= 0) continue;
  cbSpec.push({
    cbAmp: amp,
    phaseMM2: b?.v != null ? String(b.v).trim() : '',
    peMM2: c?.v != null ? String(c.v).trim() : '',
  });
}

function colList(col: number, startRow: number, endRow: number): string[] {
  const out: string[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const cell = wsOld[XLSX.utils.encode_cell({ r, c: col })];
    if (!cell || cell.v == null || cell.v === '') continue;
    out.push(String(cell.v).trim());
  }
  return out;
}

const cbRatings = colList(4, 2, 40);
const cbPoles = colList(6, 2, 20);
const cbIsc = colList(8, 2, 20);
const cbTypes = colList(10, 2, 20);

const jackets = [
  'CU/PVC',
  'CU/XLPE/PVC',
  'CU/MICA/XLPE/FR-PVC',
  'CU/MICA/XLPE/LSZH',
  'CU/PVC/PVC',
] as const;

type OdFlat = {
  coreCount: number;
  sectionMM2: number;
} & Partial<Record<(typeof jackets)[number], number>>;

const odRows: OdFlat[] = [];
for (const row of parsed.tables.outerDias) {
  const o: OdFlat = { coreCount: row.coreCount, sectionMM2: row.sectionMM2 };
  let has = false;
  for (const j of jackets) {
    const v = row.odBySheath[j];
    if (v != null && v > 0) {
      o[j] = v;
      has = true;
    }
  }
  if (has) odRows.push(o);
}

const conduits: {
  label: string;
  material: string;
  outerDiaMM: number;
  wallThicknessMM: number;
  innerDiaMM: number;
  note: string;
}[] = [];

for (let r = 3; r <= 12; r++) {
  const outer = wsOld[XLSX.utils.encode_cell({ r, c: 29 })];
  const thick = wsOld[XLSX.utils.encode_cell({ r, c: 30 })];
  const inner = wsOld[XLSX.utils.encode_cell({ r, c: 31 })];
  if (!outer || outer.v == null || !inner || inner.v == null) continue;
  const outerDiaMM = extractNumber(outer.v);
  const innerDiaMM = extractNumber(inner.v);
  if (outerDiaMM <= 0 || innerDiaMM <= 0) continue;
  const wallThicknessMM =
    thick?.v != null ? extractNumber(thick.v) : Math.round((outerDiaMM - innerDiaMM) * 50) / 100;
  conduits.push({
    label: `D${Math.round(outerDiaMM)}`,
    material: 'PVC',
    outerDiaMM,
    wallThicknessMM,
    innerDiaMM,
    note: '',
  });
}

const wb = XLSX.utils.book_new();

const readme = XLSX.utils.aoa_to_sheet([
  ['Spec. Cable — lookup tables for VoltReview'],
  [''],
  ['Edit the sheets below, save this file, then click "Tải lại bảng tra" in the web app.'],
  [''],
  ['Sheet', 'Purpose'],
  ['CB_SPEC', 'CB rated current (A) → required phase / PE section (mm²)'],
  ['CB_OPTIONS', 'Lists for CB_Rating, CB_Pole, CB_Isc, CB_Type'],
  ['OUTER_DIA', 'Cable outer diameter (mm) by coreCount × section × jacket'],
  ['CONDUIT', 'Conduit sizes for fill checks'],
  [''],
  ['Removed (unused by app)', '_OnPhase, Dxx→innerDia mini-table, HDPE block, empty OD stub rows'],
]);
XLSX.utils.book_append_sheet(wb, readme, 'README');

XLSX.utils.book_append_sheet(
  wb,
  XLSX.utils.json_to_sheet(cbSpec.map((r) => ({ cbAmp: r.cbAmp, phaseMM2: r.phaseMM2, peMM2: r.peMM2 }))),
  'CB_SPEC'
);

const maxOpt = Math.max(cbRatings.length, cbPoles.length, cbIsc.length, cbTypes.length, 1);
const optAoA: (string | number)[][] = [['CB_Rating', 'CB_Pole', 'CB_Isc', 'CB_Type']];
for (let i = 0; i < maxOpt; i++) {
  optAoA.push([cbRatings[i] ?? '', cbPoles[i] ?? '', cbIsc[i] ?? '', cbTypes[i] ?? '']);
}
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(optAoA), 'CB_OPTIONS');

const odAoA: (string | number)[][] = [
  ['coreCount', 'sectionMM2', ...jackets],
];
for (const r of odRows) {
  odAoA.push([
    r.coreCount,
    r.sectionMM2,
    r['CU/PVC'] ?? '',
    r['CU/XLPE/PVC'] ?? '',
    r['CU/MICA/XLPE/FR-PVC'] ?? '',
    r['CU/MICA/XLPE/LSZH'] ?? '',
    r['CU/PVC/PVC'] ?? '',
  ]);
}
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(odAoA), 'OUTER_DIA');

XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(conduits), 'CONDUIT');

XLSX.writeFile(wb, outPath);
console.log(`Wrote ${outPath}`);
console.log({
  cbSpec: cbSpec.length,
  ratings: cbRatings.length,
  poles: cbPoles.length,
  isc: cbIsc.length,
  types: cbTypes.length,
  od: odRows.length,
  conduits: conduits.length,
});
