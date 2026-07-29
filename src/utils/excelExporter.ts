import * as XLSX from 'xlsx';
import { ReviewIssue } from '../types';

/**
 * Exports the review report issues into an Excel file (.xlsx)
 * matching the VBA REPORT worksheet structure with audit trail preservation
 */
export function exportReportToExcel(issues: ReviewIssue[], originalFileName: string): void {
  const wb = XLSX.utils.book_new();

  const reportRows: (string | number)[][] = [];

  // Title Row (Merged A1:F1 in Excel)
  reportRows.push(['CRITICAL ERRORS & WARNINGS (AUDIT TRAIL)', '', '', '', '', '']);
  // Header Row (A2:F2)
  reportRows.push(['#', 'Sheet (Panel)', 'Line (Circuit)', 'Description', 'Status', 'Remarks']);

  // Unresolved Issues first
  const unresolved = issues.filter(i => i.status === 'UNRESOLVED' || !i.isResolved);
  const resolved = issues.filter(i => i.status !== 'UNRESOLVED' && i.isResolved);

  let count = 1;

  for (const issue of unresolved) {
    reportRows.push([
      count++,
      issue.sheetName,
      issue.lineNameWithDesc,
      issue.description,
      issue.status,
      issue.remarks || '',
    ]);
  }

  for (const issue of resolved) {
    reportRows.push([
      count++,
      issue.sheetName,
      issue.lineNameWithDesc,
      issue.description,
      issue.status,
      issue.remarks || '',
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(reportRows);

  // Column Widths
  ws['!cols'] = [
    { wch: 6 },  // #
    { wch: 20 }, // Sheet
    { wch: 35 }, // Line
    { wch: 55 }, // Description
    { wch: 15 }, // Status
    { wch: 35 }, // Remarks
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'REPORT');

  const baseName = originalFileName ? originalFileName.replace(/\.[^/.]+$/, '') : 'Panel_Schedule';
  const exportName = `${baseName}_REVIEW_REPORT.xlsx`;

  XLSX.writeFile(wb, exportName);
}
