/**
 * Patch ô Excel kiểu “phẫu thuật” trên file OOXML gốc (ZIP + XML).
 * Không dùng SheetJS.write rebuild — tránh làm hỏng workbook (macro, style, formula liên sheet…).
 */

import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import {
  applyCellPatches,
  CellPatch,
  cloneWorkbookForWrite,
  workbookToArrayBuffer,
} from './excelPatch';

function escapeXml(text: string): string {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function encodeCellA1(row0: number, col0: number): string {
  return XLSX.utils.encode_cell({ r: row0, c: col0 });
}

function isZipXlsx(buffer: ArrayBuffer): boolean {
  if (!buffer || buffer.byteLength < 4) return false;
  const u8 = new Uint8Array(buffer);
  // PK\x03\x04
  return u8[0] === 0x50 && u8[1] === 0x4b && u8[2] === 0x03 && u8[3] === 0x04;
}

/**
 * Kiểm tra buffer có phải workbook Excel hợp lệ (đủ sheet, đọc được bằng SheetJS)
 */
export function validateExcelBuffer(
  buffer: ArrayBuffer,
  expectedSheetNames?: string[]
): { ok: true; sheetNames: string[] } | { ok: false; reason: string } {
  if (!isZipXlsx(buffer)) {
    return { ok: false, reason: 'Output is not a valid ZIP/XLSX package (bad file header).' };
  }
  if (buffer.byteLength < 64) {
    return { ok: false, reason: 'Output file is empty or too small.' };
  }
  try {
    const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellFormula: true });
    const names = wb.SheetNames || [];
    if (names.length === 0) {
      return { ok: false, reason: 'Output workbook has no sheets.' };
    }
    if (expectedSheetNames && expectedSheetNames.length > 0) {
      const missing = expectedSheetNames.filter((n) => !names.includes(n));
      if (missing.length > 0) {
        return {
          ok: false,
          reason: `Output is missing sheets: ${missing.join(', ')}. Save aborted to protect your file.`,
        };
      }
    }
    return { ok: true, sheetNames: names };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, reason: `Output failed Excel parse check: ${msg}` };
  }
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

/** Map tên sheet → đường dẫn worksheet trong ZIP (vd xl/worksheets/sheet1.xml) */
async function buildSheetPathMap(zip: JSZip): Promise<Map<string, string>> {
  const map = new Map<string, string>();

  const wbFile =
    zip.file('xl/workbook.xml') ||
    zip.file(/xl\/workbook\.xml$/i)?.[0] ||
    null;
  if (!wbFile) {
    throw new Error('Invalid workbook: missing xl/workbook.xml');
  }

  const relsFile =
    zip.file('xl/_rels/workbook.xml.rels') ||
    zip.file(/xl\/_rels\/workbook\.xml\.rels$/i)?.[0] ||
    null;
  if (!relsFile) {
    throw new Error('Invalid workbook: missing xl/_rels/workbook.xml.rels');
  }

  const wbXml = await wbFile.async('string');
  const relsXml = await relsFile.async('string');

  const ridToTarget = new Map<string, string>();
  const relRe =
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/gi;
  const relReAlt =
    /<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="([^"]+)"[^>]*\/?>/gi;

  let m: RegExpExecArray | null;
  while ((m = relRe.exec(relsXml))) {
    ridToTarget.set(m[1], m[2].replace(/\\/g, '/'));
  }
  while ((m = relReAlt.exec(relsXml))) {
    ridToTarget.set(m[2], m[1].replace(/\\/g, '/'));
  }

  // <sheet name="..." r:id="rId1" /> hoặc id="..."
  const sheetRe =
    /<sheet\b[^>]*\bname="([^"]+)"[^>]*\b(?:r:)?id="([^"]+)"[^>]*\/?>/gi;
  const sheetReAlt =
    /<sheet\b[^>]*\b(?:r:)?id="([^"]+)"[^>]*\bname="([^"]+)"[^>]*\/?>/gi;

  const addSheet = (nameRaw: string, rid: string) => {
    const name = decodeXmlEntities(nameRaw);
    const target = ridToTarget.get(rid);
    if (!target) return;
    // Target thường là "worksheets/sheet1.xml" (relative to xl/)
    let path = target;
    if (!path.startsWith('xl/')) {
      path = `xl/${path.replace(/^\//, '')}`;
    }
    // Chuẩn hóa ./ 
    path = path.replace(/\/\.\//g, '/');
    map.set(name, path);
  };

  while ((m = sheetRe.exec(wbXml))) {
    addSheet(m[1], m[2]);
  }
  while ((m = sheetReAlt.exec(wbXml))) {
    addSheet(m[2], m[1]);
  }

  if (map.size === 0) {
    throw new Error('Could not map worksheet names inside the Excel package.');
  }
  return map;
}

function buildCellXml(addr: string, value: string | number): string {
  if (typeof value === 'number' && !Number.isNaN(value)) {
    return `<c r="${addr}"><v>${value}</v></c>`;
  }
  const text = escapeXml(String(value ?? ''));
  // inlineStr: không đụng sharedStrings.xml → an toàn hơn
  return `<c r="${addr}" t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`;
}

/**
 * Thay hoặc chèn một ô <c r="A1">...</c> trong XML worksheet.
 * Giữ nguyên style (thuộc tính s=) nếu ô cũ có.
 */
function upsertCellInSheetXml(sheetXml: string, addr: string, value: string | number): string {
  const cellRe = new RegExp(`<c\\b[^>]*\\br="${addr}"[^>]*(?:/>|>[\\s\\S]*?</c>)`, 'i');

  let styleAttr = '';
  const existing = sheetXml.match(cellRe);
  if (existing) {
    const styleMatch = existing[0].match(/\bs="([^"]+)"/);
    if (styleMatch) styleAttr = ` s="${styleMatch[1]}"`;
  }

  let newCell = buildCellXml(addr, value);
  // Chèn lại style id nếu có
  if (styleAttr) {
    newCell = newCell.replace('<c ', `<c${styleAttr} `);
  }

  if (existing) {
    return sheetXml.replace(cellRe, newCell);
  }

  // Ô mới: chèn vào trước </sheetData>
  if (/<\/sheetData>/i.test(sheetXml)) {
    // Tìm row chứa ô này
    const rowNum = XLSX.utils.decode_cell(addr).r + 1; // 1-based
    const rowRe = new RegExp(`<row\\b[^>]*\\br="${rowNum}"[^>]*>`, 'i');
    if (rowRe.test(sheetXml)) {
      // Chèn trước </row> của hàng đó
      const rowBlockRe = new RegExp(
        `(<row\\b[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`,
        'i'
      );
      return sheetXml.replace(rowBlockRe, `$1$2${newCell}$3`);
    }
    // Không có row → tạo row mới trước </sheetData>
    const newRow = `<row r="${rowNum}">${newCell}</row>`;
    return sheetXml.replace(/<\/sheetData>/i, `${newRow}</sheetData>`);
  }

  throw new Error(`Worksheet XML has no sheetData; cannot patch cell ${addr}.`);
}

/**
 * Áp patches lên buffer .xlsx/.xlsm gốc — giữ nguyên phần còn lại của package.
 */
export async function patchXlsxBuffer(
  originalBuffer: ArrayBuffer,
  patches: CellPatch[]
): Promise<ArrayBuffer> {
  if (!isZipXlsx(originalBuffer)) {
    throw new Error(
      'Surgical save only supports .xlsx / .xlsm. Re-open the file as .xlsx, or export via Download (legacy rebuild).'
    );
  }
  if (!patches.length) {
    return originalBuffer.slice(0);
  }

  const zip = await JSZip.loadAsync(originalBuffer, { checkCRC32: true });
  const sheetMap = await buildSheetPathMap(zip);

  // Gom patch theo sheet
  const bySheet = new Map<string, CellPatch[]>();
  for (const p of patches) {
    const list = bySheet.get(p.sheetName) || [];
    list.push(p);
    bySheet.set(p.sheetName, list);
  }

  for (const [sheetName, sheetPatches] of bySheet) {
    const path = sheetMap.get(sheetName);
    if (!path) {
      throw new Error(`Sheet "${sheetName}" was not found inside the Excel package.`);
    }

    const normalize = (p: string) => p.replace(/\\/g, '/').replace(/^\//, '');
    const wanted = normalize(path).toLowerCase();
    const actualPath =
      Object.keys(zip.files).find((k) => !zip.files[k].dir && normalize(k).toLowerCase() === wanted) ||
      path;

    const sheetFile = zip.file(actualPath);
    if (!sheetFile) {
      throw new Error(`Worksheet file missing in package: ${path}`);
    }

    let xml = await sheetFile.async('string');
    for (const p of sheetPatches) {
      const addr = encodeCellA1(p.row, p.col);
      xml = upsertCellInSheetXml(xml, addr, p.value);
    }
    zip.file(actualPath, xml);
  }

  // Xóa calcChain nếu có — tránh Excel báo lỗi công thức cũ trỏ ô đã đổi thành giá trị
  const calcNames = Object.keys(zip.files).filter((n) => /xl\/calcChain\.xml$/i.test(n));
  for (const n of calcNames) {
    zip.remove(n);
  }

  // Gỡ Override calcChain khỏi [Content_Types].xml nếu còn sót
  const ctFile = zip.file('[Content_Types].xml');
  if (ctFile && calcNames.length > 0) {
    let ct = await ctFile.async('string');
    ct = ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/gi, '');
    zip.file('[Content_Types].xml', ct);
  }

  const out = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  return out;
}

/**
 * Xây buffer lưu an toàn:
 * 1) Ưu tiên patch ZIP trên bytes gốc
 * 2) Validate parse được + đủ sheet
 * 3) Fallback SheetJS chỉ khi không có buffer gốc (demo) — vẫn validate
 */
export async function buildSafeSaveBuffer(opts: {
  originalBuffer: ArrayBuffer | null;
  workbook: XLSX.WorkBook;
  patches: CellPatch[];
}): Promise<{ buffer: ArrayBuffer; method: 'zip-patch' | 'sheetjs-rebuild' }> {
  const expectedSheets = opts.workbook.SheetNames || [];

  if (opts.originalBuffer && isZipXlsx(opts.originalBuffer)) {
    const buffer = await patchXlsxBuffer(opts.originalBuffer, opts.patches);
    const check = validateExcelBuffer(buffer, expectedSheets);
    if (check.ok === false) {
      throw new Error(`Safe save validation failed (zip-patch):\n${check.reason}`);
    }
    // Chốt chặn cuối chống truncate. Ngưỡng phải rộng: JSZip ghi lại bằng DEFLATE,
    // trong khi file gốc có thể lưu gần như không nén — khi đó output nhỏ hơn nhiều
    // lần vẫn hoàn toàn hợp lệ (đã qua validateExcelBuffer: parse được + đủ sheet).
    const MIN_RATIO = 0.1;
    if (
      opts.originalBuffer.byteLength > 10_000 &&
      buffer.byteLength < opts.originalBuffer.byteLength * MIN_RATIO
    ) {
      throw new Error(
        `Safe save validation failed: output shrank too much (${opts.originalBuffer.byteLength} → ${buffer.byteLength} bytes). Save aborted.`
      );
    }
    return { buffer, method: 'zip-patch' };
  }

  // Fallback: demo / file không phải OOXML
  const working = cloneWorkbookForWrite(opts.workbook);
  applyCellPatches(working, opts.patches);
  const buffer = workbookToArrayBuffer(working);
  const check = validateExcelBuffer(buffer);
  if (check.ok === false) {
    throw new Error(`Safe save validation failed (rebuild):\n${check.reason}`);
  }
  return { buffer, method: 'sheetjs-rebuild' };
}
