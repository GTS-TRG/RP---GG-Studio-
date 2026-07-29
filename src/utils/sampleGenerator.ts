import * as XLSX from 'xlsx';

/**
 * Creates a sample electrical panel calculation Excel file (.xlsx)
 * with realistic Vietnamese panel schedules and intentional audit cases.
 */
export function createSampleWorkbook(): Uint8Array {
  const wb = XLSX.utils.book_new();

  // 1. SPEC CABLE SHEET
  const specData = [
    ['CB Ampere (In)', 'Phase Cable (Mm2)', 'PE Cable (Mm2)'],
    [6, '1.5', '1.5'],
    [10, '1.5', '1.5'],
    [16, '2.5', '2.5'],
    [20, '4', '4'],
    [25, '4', '4'],
    [32, '6', '6'],
    [40, '10', '10'],
    [50, '10', '10'],
    [63, '16', '16'],
    [80, '25', '16'],
    [100, '35', '16'],
    [125, '50', '25'],
    [160, '70', '35'],
    [200, '95', '50'],
    [250, '120', '70'],
    [315, '185', '95'],
    [400, '240', '120'],
  ];

  const specWs = XLSX.utils.aoa_to_sheet(specData);

  // Vùng cột P–U: bảng đường kính ngoài theo tiết diện / số lõi / quy cách vỏ
  const odBlock: (string | number | null)[][] = [
    ['TIẾT DIỆN', 'CU/PVC', 'CU/XLPE/PVC', 'Cu/Mica/XLPE/FR-PVC', 'Cu/Mica/XLPE/LSZH', 'CU/PVC/PVC'],
    ['1C', null, null, null, null, null],
    [1.5, 3.2, 5.3, 6.4, 6.4, null],
    [2.5, 3.6, 5.7, 6.9, 6.9, null],
    [4, 4.6, 6.3, 7.4, 7.4, null],
    [6, 5.1, 6.8, 8, 8, null],
    [10, 6.1, 7.5, 8.6, 8.6, null],
    [16, 6.7, 8.4, 9.5, 9.5, null],
    ['2C', null, null, null, null, null],
    [1.5, null, 10.2, 12.4, 12.4, 10.6],
    [2.5, null, 11.1, 13.3, 13.3, 11.5],
    [4, null, 12.1, 14.4, 14.4, 13.3],
    [6, null, 13.3, 15.5, 15.5, 14.5],
    ['3C', null, null, null, null, null],
    [1.5, null, 10.6, 13.1, 13.1, 11.1],
    [2.5, null, 11.6, 14, 14, 12],
    [4, null, 12.8, 15.2, 15.2, 14.1],
    [6, null, 14, 16.4, 16.4, 15.3],
    ['4C', null, null, null, null, null],
    [1.5, null, 11.4, 14.1, 14.1, 11.9],
    [2.5, null, 12.5, 15.2, 15.2, 13],
    [4, null, 13.8, 16.5, 16.5, 15.3],
    [6, null, 15.2, 17.9, 17.9, 16.6],
  ];
  XLSX.utils.sheet_add_aoa(specWs, odBlock, { origin: 'P1' });

  // Vùng cột AC–AF: bảng ống luồn PVC (đường kính trong)
  const conduitBlock: (string | number | null)[][] = [
    ['ỐNG LUỒN DÂY PVC', null, null, null],
    ['STT', 'Đường kính ngoài\n(mm)', 'Độ dày ống\n(mm)', 'Đường kính trong\n(mm)'],
    [1, 16, 1.2, 13.6],
    [2, 20, 1.4, 17.2],
    [3, 25, 1.5, 22],
    [4, 32, 1.86, 28.3],
    [5, 40, 2.1, 35.8],
    [6, 50, 2.4, 45.2],
    [7, 63, 3, 57],
  ];
  XLSX.utils.sheet_add_aoa(specWs, conduitBlock, { origin: 'AC1' });

  // Cột E–H: danh sách dropdown CB (In / Loại / Số cực / Isc)
  const cbOptionBlock: (string | number)[][] = [
    ['CB_Rating', 'CB_Type', 'Pole', 'Isc_kA'],
    ['10A', 'MCB', '1P', '6kA'],
    ['16A', 'MCCB', '1P+N', '10kA'],
    ['20A', 'ACB', '2P', '15kA'],
    ['25A', 'RCCB', '3P', '25kA'],
    ['32A', 'RCBO', '3P+N', '35kA'],
    ['40A', 'FUSE', '4P', '50kA'],
    ['50A', '', '', '65kA'],
    ['63A', '', '', '70kA'],
    ['80A', '', '', '85kA'],
    ['100A', '', '', '100kA'],
    ['125A', '', '', ''],
    ['160A', '', '', ''],
    ['200A', '', '', ''],
    ['250A', '', '', ''],
    ['315A', '', '', ''],
    ['400A', '', '', ''],
  ];
  XLSX.utils.sheet_add_aoa(specWs, cbOptionBlock, { origin: 'E1' });

  XLSX.utils.book_append_sheet(wb, specWs, 'Spec. Cable');

  // Helper to construct a Panel Sheet header (Rows 1 to 12)
  function createPanelSheetRows(panelTitle: string): (string | number)[][] {
    return [
      ['DỰ ÁN / PROJECT:', 'TÒA NHÀ VĂN PHÒNG KHÁCH SẠN HIGH-TECH'],
      ['HẠNG MỤC / ITEM:', 'BẢNG TÍNH TẢI VÀ THIẾT BỊ BẢO VỆ TỦ ĐIỆN'],
      ['TÊN TỦ ĐIỆN / PANEL:', panelTitle],
      [''],
      ['MÃ MẠCH', 'MÔ TẢ PHỤ TẢI', '', 'CÔNG SUẤT VÀ DÒNG ĐIỆN TÍNH TOÁN', '', '', '', 'THIẾT BỊ BẢO VỆ (CB)', '', '', '', 'DÂY DẪN ĐIỆN (CABLE)', ''],
      ['Line Name', 'Circuit Description', '', 'Phase R (kW)', 'Phase Y (kW)', 'Phase B (kW)', 'I calc (A)', 'Type', 'Pole', 'Rating (In)', 'Isc (kA)', 'Phase Cable Spec', 'PE Cable Spec'],
      ['(A)', '(B)', '(C)', '(D)', '(E)', '(F)', '(G)', '(H)', '(I)', '(J)', '(K)', '(L)', '(M)'],
      [''],
      [''],
      [''],
      [''],
      [''],
    ];
  }

  // 2. MSB SHEET (Tủ tổng MSB-01)
  const msbRows = createPanelSheetRows('TỦ ĐIỆN TỔNG MSB-01');
  msbRows.push(
    ['CP-01', 'CẤP NGUỒN CHO TỦ DB-L01', '', 5.2, 5.2, 5.2, 25.5, 'MCCB', '3P', '40A', '65kA', 'Cu/XLPE/PVC 4x1C-10mm2', 'Cu/PVC 1x10mm2', 'IN CABLE TRAY'],
    ['L1', 'CHIẾU SÁNG SẢNH CHÍNH TẦNG 1', '', 1.8, 0, 0, 8.2, 'MCB', '1P', '16A', '10kA', 'Cu/XLPE/PVC 2x2.5mm2', 'Cu/PVC 1x2.5mm2', 'IN 02 CONDUIT PVC D16'], // Error: MSB MCB Isc 10kA < 65kA min required for MSB!
    ['S1', 'Ổ CẮM TẦNG TRỄN', '', 2.5, 0, 0, 11.4, 'MCB', '1P', '20A', '65kA', 'Cu/XLPE/PVC 2x4mm2', 'Cu/PVC 1x4mm2', 'IN 02 CONDUIT PVC D20'], // Error: Socket S must use 1P+N!
    ['P1', 'CẤP NGUỒN MÁY BƠM NƯỚC CẤP', '', 6.5, 6.5, 6.5, 32.0, 'MCCB', '3P', '50A', '10kA', 'Cu/XLPE/PVC 4x1C-10mm2', 'Cu/PVC 1x10mm2', 'IN CONDUIT PVC D25'], // Error: MSB MCCB Isc 10kA < 65kA & MCB Rating < 32A checks!
    ['P2 (PCCC)', 'CẤP NGUỒN BƠM CHỮA CHÁY (S/D)', '', 22.0, 22.0, 22.0, 105.0, 'MCCB', '3P', '160A', '65kA', 'Cu/XLPE/PVC 4x1C-35mm2', 'Cu/PVC 1x16mm2', 'IN CABLE TRAY'], // Fire pump (1.5 safety factor)
    ['SP-01', 'DỰ PHÒNG CHỜ MỞ RỘNG (SPARE)', '', 0, 0, 0, 0, 'MCCB', '3P', '63A', '65kA', 'Cu/XLPE/PVC 4x1C-16mm2', 'Cu/PVC 1x16mm2', 'IN CABLE TRAY'],
    ['CÔNG SUẤT TÍNH TOÁN (TOTAL CONNECTED LOAD):', '', '', '', '', '', '', '', '', '', '', '', '', '']
  );

  const msbWs = XLSX.utils.aoa_to_sheet(msbRows);
  XLSX.utils.book_append_sheet(wb, msbWs, 'MSB-01');

  // 3. DB-L01 SHEET (Tủ tầng DB-L01)
  const dblRows = createPanelSheetRows('TỦ ĐIỆN PHÂN PHỐI CHIẾU SÁNG DB-L01');
  dblRows.push(
    ['L1', 'CHIẾU SÁNG HÀNH LANG TẦNG 1', '', 1.2, 0, 0, 5.5, 'MCB', '1P', '16A', '10kA', 'Cu/XLPE/PVC 2x2.5mm2', 'Cu/PVC 1x2.5mm2', 'IN 02 CONDUIT PVC D16'],
    ['L2', 'CHIẾU SÁNG PHÒNG HỌP TẦNG 1', '', 1.5, 0, 0, 6.8, 'MCB', '1P', '16A', '10kA', 'Cu/XLPE/PVC 2x2.5mm2', 'Cu/PVC 1x2.5mm2', 'IN 02 CONDUIT PVC D16'],
    ['S1', 'Ổ CẮM LÀM VIỆC KHU KHÁCH', '', 2.8, 0, 0, 12.7, 'MCB', '1P+N', '20A', '10kA', 'Cu/XLPE/PVC 2x4mm2', 'Cu/PVC 1x4mm2', 'IN 02 CONDUIT PVC D20'],
    ['S2', 'Ổ CẮM PHÒNG TEABREAK', '', 3.5, 0, 0, 15.9, 'MCB', '1P+N', '32A', '10kA', 'Cu/XLPE/PVC 1x2.5mm2', 'Cu/PVC 1x2.5mm2', 'IN 02 CONDUIT PVC D20'], // Error: Oversized CB (32A vs 20A req) & Wrong phase cable (2.5mm2 < 6mm2 required for 32A)!
    ['P1', 'CẤP NGUỒN CHO QUẠT TẠO ÁP HÀNH LANG (PCCC)', '', 3.0, 3.0, 3.0, 14.8, 'MCB', '3P', '32A', '10kA', 'Cu/XLPE/PVC 4C-6mm2', 'Cu/PVC 1x6mm2', 'IN CONDUIT PVC D25'],
    ['SP1', 'DỰ PHÒNG TỦ DB-L01', '', 0, 0, 0, 0, 'MCB', '', '', '', '', '', ''], // Warning: Spare missing CB parameters
    ['DÒNG ĐIỆN TÍNH TOÁN TỔNG (CURRENT A)', '', '', '', '', '', '', '', '', '', '', '', '', '']
  );

  const dblWs = XLSX.utils.aoa_to_sheet(dblRows);
  XLSX.utils.book_append_sheet(wb, dblWs, 'DB-L01');

  // Convert to ArrayBuffer
  const arrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(arrayBuffer);
}
