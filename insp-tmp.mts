import * as XLSX from 'xlsx';
import fs from 'fs';
const wb = XLSX.read(fs.readFileSync(process.argv[2]), { type: 'buffer' });
console.log('SHEETS:', JSON.stringify(wb.SheetNames));
for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  const rg = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  console.log(`\n=== ${name}  ref=${ws['!ref']} ===`);
  for (let r = rg.s.r; r <= Math.min(rg.e.r, 18); r++) {
    const cells: string[] = [];
    for (let c = rg.s.c; c <= Math.min(rg.e.c, 40); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null || cell.v === '') continue;
      let v = String(cell.v).replace(/\s+/g, ' ').trim();
      if (v.length > 22) v = v.slice(0, 22) + '~';
      cells.push(`${XLSX.utils.encode_col(c)}=${v}`);
    }
    if (cells.length) console.log(`R${r + 1}: ${cells.join(' | ')}`);
  }
}
