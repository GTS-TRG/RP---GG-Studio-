import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  CableOuterDiaRow,
  CableSpecRow,
  ConduitSpec,
  PanelSheetData,
  ProjectConfig,
  ReviewIssue,
  ReviewSummary,
} from './types';
import { DEFAULT_CABLE_SPECS } from './utils/specTable';
import { DEFAULT_CABLE_OUTER_DIAS } from './utils/outerDiaTable';
import { DEFAULT_CONDUITS } from './utils/conduitTable';
import { DEFAULT_CB_RATINGS, formatCbRatingLabel } from './utils/cbRatingTable';
import type { CbRatingItem } from './utils/cbRatingTable';
import {
  DEFAULT_CB_TYPES,
  DEFAULT_ISC_OPTIONS,
  DEFAULT_POLE_OPTIONS,
  SpecListItem,
  formatCbTypeLabel,
} from './utils/cbOptionLists';
import { DEFAULT_CONFIG, reviewWorkbookPanels } from './utils/panelReviewer';
import { parseExcelWorkbook } from './utils/excelParser';
import { createSampleWorkbook } from './utils/sampleGenerator';
import { exportReportToExcel } from './utils/excelExporter';
import {
  applyCellPatches,
  CellPatch,
  downloadWorkbookBuffer,
  PANEL_COL,
  PanelEditableField,
  upsertPatch,
} from './utils/excelPatch';
import { buildSafeSaveBuffer } from './utils/excelZipPatch';
import {
  ExcelFileHandle,
  openExcelWithHandle,
  readFileFromHandle,
  saveBufferAsNewFile,
  supportsFileSystemAccess,
  writeBufferToHandle,
} from './utils/fileSystemAccess';
import { extractNumber, normalizePoleValue } from './utils/helpers';
import {
  LookupTables,
  builtInTables,
  loadLookupTables,
  applyLookupTablesToState,
} from './utils/lookupTables';

import { Header } from './components/Header';
import { UploadZone } from './components/UploadZone';
import { SummaryCards } from './components/SummaryCards';
import { ReportTable } from './components/ReportTable';
import { PanelViewer } from './components/PanelViewer';
import { SpecEditor } from './components/SpecEditor';
import { ConfigModal } from './components/ConfigModal';
import { RulesGuideModal } from './components/RulesGuideModal';

import { FileSpreadsheet, Cpu, Table, Sparkles } from 'lucide-react';

export default function App() {
  const [config, setConfig] = useState<ProjectConfig>(DEFAULT_CONFIG);
  const [specs, setSpecs] = useState<CableSpecRow[]>(DEFAULT_CABLE_SPECS);
  const [outerDias, setOuterDias] = useState<CableOuterDiaRow[]>(DEFAULT_CABLE_OUTER_DIAS);
  const [conduits, setConduits] = useState<ConduitSpec[]>(DEFAULT_CONDUITS);
  const [cbRatings, setCbRatings] = useState<CbRatingItem[]>(DEFAULT_CB_RATINGS);
  const [cbTypes, setCbTypes] = useState<SpecListItem[]>(DEFAULT_CB_TYPES);
  const [poleOptions, setPoleOptions] = useState<SpecListItem[]>(DEFAULT_POLE_OPTIONS);
  const [iscOptions, setIscOptions] = useState<SpecListItem[]>(DEFAULT_ISC_OPTIONS);
  const [panels, setPanels] = useState<PanelSheetData[]>([]);
  const [specSheetName, setSpecSheetName] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [issues, setIssues] = useState<ReviewIssue[]>([]);
  const [userAuditMap, setUserAuditMap] = useState<Map<string, { status: ReviewIssue['status']; remarks: string }>>(new Map());

  const [activeTab, setActiveTab] = useState<'REPORT' | 'PANELS' | 'SPECS'>('REPORT');
  const [isConfigOpen, setIsConfigOpen] = useState(false);
  const [isRulesOpen, setIsRulesOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  /** Workbook gốc + bytes file gốc (để patch ZIP an toàn) + handle FSA */
  const [rawWorkbook, setRawWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [originalFileBuffer, setOriginalFileBuffer] = useState<ArrayBuffer | null>(null);
  const [fileHandle, setFileHandle] = useState<ExcelFileHandle | null>(null);
  const [cellPatches, setCellPatches] = useState<Map<string, CellPatch>>(new Map());

  /** Bang tra nap tu public/data/Spec. Cable.xlsx — dung de review (khong lay Spec trong file kiem tra) */
  const [lookup, setLookup] = useState<LookupTables>(() => builtInTables());
  const [lookupInfo, setLookupInfo] = useState<{ fromFile: boolean; updatedAt?: string }>({
    fromFile: false,
  });
  const lookupRef = useRef<LookupTables>(lookup);
  lookupRef.current = lookup;

  /** Tu dong theo doi file tren dia (poll lastModified qua handle) */
  const [autoWatch, setAutoWatch] = useState(true);
  /** File tren dia da doi nhung chua nap duoc (dang co sua chua luu tren web) */
  const [diskChanged, setDiskChanged] = useState(false);
  const [lastSyncAt, setLastSyncAt] = useState<Date | null>(null);
  /** Dau van tay file da nap: mtime + size */
  const fileStampRef = useRef<{ mtime: number; size: number } | null>(null);
  /** Chan cac lan doc chong nhau */
  const busyRef = useRef(false);

  /** Đồng bộ state bảng tra + ref từ kết quả load Spec. Cable.xlsx */
  const applyTables = (tables: LookupTables) => {
    lookupRef.current = tables;
    setLookup(tables);
    applyLookupTablesToState(tables, {
      setSpecs,
      setOuterDias,
      setConduits,
      setCbRatings,
      setCbTypes,
      setPoleOptions,
      setIscOptions,
    });
    setSpecSheetName(null);
  };

  // Nạp bảng tra từ Spec. Cable.xlsx trước, rồi mới mở file demo
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const res = await loadLookupTables();
      if (cancelled) return;
      applyTables(res.tables);
      setLookupInfo({ fromFile: res.fromFile, updatedAt: res.updatedAt });
      loadDemoSample();
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  /** Nạp lại bảng tra từ Spec. Cable.xlsx và áp dụng ngay cho review */
  const handleReloadLookupTables = async () => {
    const res = await loadLookupTables(true);
    applyTables(res.tables);
    setLookupInfo({ fromFile: res.fromFile, updatedAt: res.updatedAt });

    if (!res.fromFile) {
      alert(
        'Unable to load Spec. Cable.xlsx from public/data/.\n\n' +
          `Detail: ${res.error ?? 'unknown'}\n\nUsing built-in default tables.`
      );
      return;
    }

    alert(
      `Lookup tables reloaded from Spec. Cable.xlsx.\n\n` +
        `CB: ${res.tables.cableSpecs.length} rows | Cable OD: ${res.tables.outerDias.length}` +
        ` | Conduit: ${res.tables.conduits.length}` +
        (res.updatedAt ? `\nUpdated: ${res.updatedAt}` : '')
    );
  };

  const registerPatch = (patch: CellPatch) => {
    setCellPatches((prev) => upsertPatch(prev, patch));
  };

  const handleProcessArrayBuffer = (
    buffer: ArrayBuffer,
    name: string,
    handle: ExcelFileHandle | null = null,
    fileSizeBytes?: number,
    keepActiveTab = false
  ) => {
    setIsLoading(true);
    try {
      if (!buffer || buffer.byteLength < 64) {
        throw new Error(
          'File buffer is empty. The workbook may be corrupted or was truncated by a failed save. Restore a backup and open again.'
        );
      }

      const tables = lookupRef.current;
      const parsed = parseExcelWorkbook(buffer, config, tables);
      const sizeKb = Math.round((fileSizeBytes ?? buffer.byteLength) / 1024);

      if (parsed.sheets.length === 0) {
        const allNames = parsed.rawWorkbook?.SheetNames || [];
        const sheetNames = allNames.join(', ') || '(none)';
        const hidden = parsed.skippedHiddenSheets?.length
          ? `\nHidden sheets skipped: ${parsed.skippedHiddenSheets.join(', ')}`
          : '';

        const looksLikeBlankStub =
          allNames.length === 1 && /^sheet1$/i.test(allNames[0] || '');

        throw new Error(
          `Cannot load "${name}" — no panel schedule found.\n\n` +
            `File size on disk: ${sizeKb} KB\n` +
            `Sheets on disk: ${sheetNames}${hidden}\n\n` +
            (looksLikeBlankStub
              ? `The DISK file only has a blank "Sheet1".\n` +
                `If Excel currently shows 4 panel sheets, that is often an unsaved/restored view — ` +
                `the file on disk is still the old damaged copy.\n\n` +
                `Fix:\n` +
                `1) In Excel: File → Save As → e.g. Test2_OK.xlsx\n` +
                `2) Close Excel\n` +
                `3) Open workbook → select Test2_OK.xlsx (the new file)\n`
              : `No usable panel rows were detected in the visible sheets.\n` +
                `Check that circuit data starts at the configured start row, and sheets are not all hidden.\n`)
        );
      }

      // Chỉ nạp sheet tủ điện; bảng tra giữ từ Spec. Cable.xlsx của app
      setPanels(parsed.sheets);
      setSpecs([...tables.cableSpecs]);
      setOuterDias([...tables.outerDias]);
      setConduits([...tables.conduits]);
      setCbRatings([...tables.cbRatings]);
      setCbTypes([...tables.cbTypes]);
      setPoleOptions([...tables.poleOptions]);
      setIscOptions([...tables.iscOptions]);
      setSpecSheetName(null);
      setFileName(name);
      setRawWorkbook(parsed.rawWorkbook);
      // Giữ bản copy bytes gốc — nguồn sự thật khi Save (không rebuild bằng SheetJS)
      setOriginalFileBuffer(buffer.slice(0));
      setFileHandle(handle);
      setCellPatches(new Map());
      if (!keepActiveTab) setActiveTab('PANELS');

      if (parsed.skippedHiddenSheets && parsed.skippedHiddenSheets.length > 0) {
        console.info(
          `Skipped ${parsed.skippedHiddenSheets.length} hidden sheet(s):`,
          parsed.skippedHiddenSheets.join(', ')
        );
      }

      const rawIssues = reviewWorkbookPanels(
        parsed.sheets,
        tables.cableSpecs,
        config,
        tables.outerDias,
        tables.conduits
      );
      setIssues(rawIssues);
    } catch (err) {
      console.error('Error parsing Excel file:', err);
      if (handle) {
        setFileHandle(null);
      }
      const detail = err instanceof Error ? err.message : String(err);
      alert(detail || 'Unable to read the Excel file. Please check that the file is a valid .xlsx or .xlsm workbook.');
    } finally {
      setIsLoading(false);
    }
  };

  /** Đọc lại file từ đĩa. silent = true khi chạy tự động (không alert) */
  const reloadFromDisk = async (silent = false): Promise<boolean> => {
    if (!fileHandle || busyRef.current) return false;

    busyRef.current = true;
    if (!silent) setIsLoading(true);
    try {
      const { file, buffer } = await readFileFromHandle(fileHandle);
      // Giữ nguyên tab đang xem để review liền mạch
      handleProcessArrayBuffer(buffer, file.name, fileHandle, file.size, true);
      fileStampRef.current = { mtime: file.lastModified, size: file.size };
      setDiskChanged(false);
      setLastSyncAt(new Date());
      return true;
    } catch (err) {
      // Khi Excel đang ghi file, đọc có thể lỗi — chế độ tự động sẽ thử lại ở nhịp sau
      if (!silent) {
        console.error(err);
        const detail = err instanceof Error ? err.message : String(err);
        alert(`Reload thất bại.\n\n${detail}`);
      }
      return false;
    } finally {
      busyRef.current = false;
      if (!silent) setIsLoading(false);
    }
  };

  /** Reload thủ công — từ nút bấm */
  const handleReloadFromDisk = async () => {
    if (!fileHandle) {
      alert(
        'Reload cần mở file bằng nút "Open workbook" trước.\n\n' +
          'Khi đó web app giữ được liên kết tới file trên đĩa, ' +
          'nên mỗi lần bạn lưu file trong Excel là kiểm tra lại ngay.'
      );
      return;
    }

    if (cellPatches.size > 0) {
      const ok = window.confirm(
        `Đang có ${cellPatches.size} ô sửa trên web chưa lưu.\n\n` +
          'Reload sẽ đọc lại file từ đĩa và BỎ các thay đổi này.\n\nTiếp tục Reload?'
      );
      if (!ok) return;
    }

    await reloadFromDisk(false);
  };

  const handleFileUpload = (file: File) => {
    setFileHandle(null);
    fileStampRef.current = null;
    setDiskChanged(false);
    const reader = new FileReader();
    reader.onload = (e) => {
      if (e.target?.result) {
        handleProcessArrayBuffer(e.target.result as ArrayBuffer, file.name, null, file.size);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  /** Mở file bằng File System Access API */
  const handleOpenWithFileSystem = async () => {
    if (!supportsFileSystemAccess()) {
      alert('File System Access is not supported in this browser. Please use Chrome or Edge, or choose a file via the upload button.');
      return;
    }
    try {
      const opened = await openExcelWithHandle();
      if (!opened) return;
      handleProcessArrayBuffer(opened.buffer, opened.file.name, opened.handle, opened.file.size);
      fileStampRef.current = { mtime: opened.file.lastModified, size: opened.file.size };
      setDiskChanged(false);
      setLastSyncAt(new Date());
    } catch (err) {
      console.error(err);
      const detail = err instanceof Error ? err.message : String(err);
      alert(detail || 'Unable to open the Excel file.');
    }
  };

  const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer => {
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  };

  const loadDemoSample = () => {
    setFileHandle(null);
    fileStampRef.current = null;
    setDiskChanged(false);
    const sampleBuffer = createSampleWorkbook();
    handleProcessArrayBuffer(toArrayBuffer(sampleBuffer), 'BANG_TINH_TAI_DIEN_DEMO.xlsx', null);
  };

  const handleDownloadSample = () => {
    const sampleBuffer = createSampleWorkbook();
    const blob = new Blob([sampleBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'BANG_TINH_TAI_DIEN_MAU_DEMO.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };

  /** Cập nhật mạch trên UI + đăng ký patch ô Excel tương ứng */
  const handleUpdateCircuit = (
    sheetName: string,
    rowIndex: number,
    field: PanelEditableField,
    value: string
  ) => {
    setPanels((prev) =>
      prev.map((panel) => {
        if (panel.sheetName !== sheetName) return panel;
        return {
          ...panel,
          circuits: panel.circuits.map((c) => {
            if (c.rowIndex !== rowIndex) return c;
            const next = { ...c };
            if (field === 'cbType') next.cbType = formatCbTypeLabel(value);
            else if (field === 'poleVal') next.poleVal = normalizePoleValue(value);
            else if (field === 'cbText') {
              const label = formatCbRatingLabel(value);
              next.cbText = label;
              next.cbAmp = extractNumber(label);
            } else if (field === 'cbIsc') {
              next.cbIsc = value;
              next.iscAmp = extractNumber(value);
            } else if (field === 'phaseCableText') next.phaseCableText = value;
            else if (field === 'peCableText') next.peCableText = value;
            else if (field === 'installMethod') next.installMethod = value;
            else if (field === 'lineName') next.lineName = value;
            else if (field === 'description') next.description = value;
            return next;
          }),
        };
      })
    );

    // Cot ghi nguoc phai lay theo bo cuc DA DO cua chinh sheet do,
    // vi moi form Excel dat cac cot o vi tri khac nhau.
    const targetPanel = panels.find((p) => p.sheetName === sheetName);
    const targetCol = targetPanel?.cols?.[field] ?? PANEL_COL[field];

    registerPatch({
      sheetName,
      row: rowIndex - 1, // Excel 1-based → SheetJS 0-based
      col: targetCol,
      value:
        field === 'poleVal'
          ? normalizePoleValue(value)
          : field === 'cbType'
          ? formatCbTypeLabel(value)
          : field === 'cbText'
          ? formatCbRatingLabel(value)
          : value,
    });
  };

  /** Lưu các ô đã sửa — Save As bản sao (patch ZIP trên bytes gốc, không rebuild SheetJS) */
  const handleSaveToExcel = async () => {
    if (!rawWorkbook) {
      alert('No workbook loaded.');
      return;
    }
    if (cellPatches.size === 0) {
      alert('No cell changes to save.');
      return;
    }

    setIsSaving(true);
    try {
      const patches: CellPatch[] = Array.from(cellPatches.values());
      const { buffer, method } = await buildSafeSaveBuffer({
        originalBuffer: originalFileBuffer,
        workbook: rawWorkbook,
        patches,
      });

      const base = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'Panel_Schedule';
      const suggested = `${base}_UPDATED.xlsx`;

      const savedAs = await saveBufferAsNewFile(buffer, suggested);
      if (savedAs) {
        // Đồng bộ memory + bytes gốc = bản vừa lưu (lần Save sau patch tiếp trên bản này)
        applyCellPatches(rawWorkbook, patches);
        setOriginalFileBuffer(buffer.slice(0));
        setCellPatches(new Map());
        alert(
          `Saved ${patches.length} cell change(s) to:\n${savedAs}\n\n` +
            `Method: ${method === 'zip-patch' ? 'safe ZIP cell patch (original structure kept)' : 'workbook rebuild'}\n` +
            `Tip: Keep using this new copy. Do not overwrite a file that is open in Excel.`
        );
        return;
      }

      downloadWorkbookBuffer(buffer, suggested);
      applyCellPatches(rawWorkbook, patches);
      setOriginalFileBuffer(buffer.slice(0));
      setCellPatches(new Map());
      alert(
        `Saved ${patches.length} cell change(s) as download: ${suggested}\n\nOriginal file was not modified.`
      );
    } catch (err) {
      console.error(err);
      const detail = err instanceof Error ? err.message : String(err);
      if (detail === 'Save cancelled.') return;
      alert(
        `Failed to save Excel file.\n\n${detail}\n\n` +
          `Your original file on disk was NOT modified by this failed save.`
      );
    } finally {
      setIsSaving(false);
    }
  };

  /** Ghi đè file gốc — chỉ qua ZIP patch + validate; vẫn khuyến nghị Save as copy */
  const handleOverwriteOriginal = async () => {
    if (!rawWorkbook || !fileHandle) {
      alert('Open a file with "Open workbook" first to enable overwrite.');
      return;
    }
    if (cellPatches.size === 0) {
      alert('No cell changes to save.');
      return;
    }
    if (!originalFileBuffer) {
      alert('Original file bytes are missing. Use Save as copy instead.');
      return;
    }

    const ok = window.confirm(
      'Overwrite the ORIGINAL Excel file?\n\n' +
        'Strongly recommended: use "Save as copy" instead.\n\n' +
        'Only continue if:\n' +
        '1) The file is CLOSED in Excel\n' +
        '2) You have a backup\n\n' +
        'Continue overwrite?'
    );
    if (!ok) return;

    setIsSaving(true);
    try {
      const patches: CellPatch[] = Array.from(cellPatches.values());
      const { buffer, method } = await buildSafeSaveBuffer({
        originalBuffer: originalFileBuffer,
        workbook: rawWorkbook,
        patches,
      });

      if (method !== 'zip-patch') {
        throw new Error(
          'Refuse to overwrite original without ZIP patch mode. Use Save as copy instead.'
        );
      }

      await writeBufferToHandle(fileHandle, buffer);
      applyCellPatches(rawWorkbook, patches);
      setOriginalFileBuffer(buffer.slice(0));
      setCellPatches(new Map());
      alert(`Overwrote original file with ${patches.length} cell change(s) (safe ZIP patch).`);
    } catch (err) {
      console.error(err);
      const detail = err instanceof Error ? err.message : String(err);
      alert(
        `Overwrite failed.\n\n${detail}\n\n` +
          `If the file no longer opens, restore from backup / OneDrive version history.\n` +
          `Prefer "Save as copy" next time.`
      );
    } finally {
      setIsSaving(false);
    }
  };

  /**
   * Theo doi file tren dia: dinh ky doc metadata (lastModified + size) qua handle.
   * Trinh duyet khong co API watch file, nen day la cach nhe nhat -
   * getFile() chi lay metadata, khong doc noi dung.
   */
  useEffect(() => {
    if (!fileHandle || !autoWatch) return;

    let cancelled = false;
    const POLL_MS = 1500;

    const tick = async () => {
      if (cancelled || busyRef.current) return;
      try {
        const file = await fileHandle.getFile();
        const stamp = fileStampRef.current;
        const changed =
          !stamp || file.lastModified !== stamp.mtime || file.size !== stamp.size;
        if (!changed || cancelled) return;

        // Excel ghi file theo nhieu buoc -> doi 1 nhip cho file on dinh
        await new Promise((r) => setTimeout(r, 400));
        if (cancelled) return;
        const again = await fileHandle.getFile();
        if (again.lastModified !== file.lastModified || again.size !== file.size) {
          return; // van dang ghi, de nhip sau xu ly
        }

        // Co sua chua luu tren web -> khong tu ghi de, chi bao hieu
        if (cellPatches.size > 0) {
          setDiskChanged(true);
          return;
        }

        await reloadFromDisk(true);
      } catch {
        // File dang bi khoa / dang ghi -> bo qua, thu lai o nhip sau
      }
    };

    const id = window.setInterval(tick, POLL_MS);
    return () => {
      cancelled = true;
      window.clearInterval(id);
    };
  }, [fileHandle, autoWatch, cellPatches.size, config]);

  // Re-run review whenever panels, specs, outerDias, conduits, or config change
  useEffect(() => {
    if (panels.length > 0) {
      const rawIssues = reviewWorkbookPanels(panels, specs, config, outerDias, conduits);

      const merged = rawIssues.map((iss) => {
        const key = `${iss.sheetName}|${iss.lineName}|${iss.description}`;
        if (userAuditMap.has(key)) {
          const override = userAuditMap.get(key)!;
          return {
            ...iss,
            status: override.status,
            remarks: override.remarks,
            isResolved: override.status !== 'UNRESOLVED',
          };
        }
        return iss;
      });

      setIssues(merged);
    }
  }, [panels, specs, config, outerDias, conduits]);

  const handleUpdateIssueStatus = (issueId: string, status: ReviewIssue['status'], remarks: string) => {
    const nextMap = new Map(userAuditMap);
    setIssues((prev) =>
      prev.map((i) => {
        if (i.id === issueId) {
          const updated = {
            ...i,
            status,
            remarks,
            isResolved: status !== 'UNRESOLVED',
          };
          const key = `${i.sheetName}|${i.lineName}|${i.description}`;
          nextMap.set(key, { status, remarks });
          return updated;
        }
        return i;
      })
    );
    setUserAuditMap(nextMap);
  };

  const handleBatchUpdateStatus = (status: ReviewIssue['status']) => {
    const nextMap = new Map(userAuditMap);
    setIssues((prev) =>
      prev.map((i) => {
        const key = `${i.sheetName}|${i.lineName}|${i.description}`;
        nextMap.set(key, { status, remarks: i.remarks });
        return {
          ...i,
          status,
          isResolved: status !== 'UNRESOLVED',
        };
      })
    );
    setUserAuditMap(nextMap);
  };

  const handleExportReport = () => {
    if (issues.length === 0) {
      alert('No review findings available to export.');
      return;
    }
    exportReportToExcel(issues, fileName);
  };

  const summary: ReviewSummary = useMemo(() => {
    const totalCircuits = panels.reduce((acc, p) => acc + p.circuits.length, 0);
    const errorCount = issues.filter((i) => !i.isWarning).length;
    const warningCount = issues.filter((i) => i.isWarning).length;
    const resolvedCount = issues.filter((i) => i.status !== 'UNRESOLVED').length;

    return {
      fileName,
      totalPanels: panels.length,
      totalCircuits,
      errorCount,
      warningCount,
      resolvedCount,
      specCount: specs.length,
    };
  }, [panels, issues, specs, fileName]);

  const panelNames = useMemo(() => panels.map((p) => p.sheetName), [panels]);
  const pendingPatchCount = cellPatches.size;
  const canSaveOriginal = !!fileHandle;

  return (
    <div className="min-h-screen bg-[#EEF2F6] text-[#1A2332] flex flex-col font-sans antialiased">
      {/* Top Header */}
      <Header
        onOpenConfig={() => setIsConfigOpen(true)}
        onOpenRules={() => setIsRulesOpen(true)}
        onDownloadSample={handleDownloadSample}
        onExportReport={handleExportReport}
        onSaveToExcel={handleSaveToExcel}
        hasIssues={issues.length > 0}
        fileName={fileName}
        pendingPatchCount={pendingPatchCount}
        canSaveOriginal={canSaveOriginal}
        isSaving={isSaving}
      />

      {/* Main: task strip tren + vung lam viec full width */}
      <main className="flex-1 w-full min-h-0">
        <div className="h-full w-full px-2 sm:px-3 lg:px-4 py-3 flex flex-col gap-3">
          {/* Top task strip — file & thống kê thẩm tra */}
          <div className="w-full flex flex-col xl:flex-row gap-3 items-stretch">
            <div className="w-full xl:w-[420px] 2xl:w-[560px] shrink-0">
              <UploadZone
                onFileUpload={handleFileUpload}
                onOpenWithFileSystem={handleOpenWithFileSystem}
                onLoadDemo={loadDemoSample}
                fileName={fileName}
                totalPanels={summary.totalPanels}
                totalCircuits={summary.totalCircuits}
                specSheetName={specSheetName}
                isLoading={isLoading}
                variant="banner"
                canWriteOriginal={canSaveOriginal}
                pendingPatchCount={pendingPatchCount}
                onSaveToExcel={handleSaveToExcel}
                onOverwriteOriginal={handleOverwriteOriginal}
                isSaving={isSaving}
                onReload={handleReloadFromDisk}
                canReload={canSaveOriginal}
                autoWatch={autoWatch}
                onToggleAutoWatch={setAutoWatch}
                diskChanged={diskChanged}
                lastSyncAt={lastSyncAt}
              />
            </div>

            {panels.length > 0 && (
              <div className="flex-1 min-w-0">
                <SummaryCards summary={summary} variant="grid" />
              </div>
            )}
          </div>

          {/* Workspace — tabs + tables (full width) */}
          <section className="flex-1 min-w-0 w-full space-y-4">
            {panels.length > 0 && (
              <div className="flex items-center justify-between border-b border-[#D5DEE8] pb-2">
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => setActiveTab('REPORT')}
                    className={`px-4 py-2 text-xs font-bold rounded-full border transition-all flex items-center space-x-2 ${
                      activeTab === 'REPORT'
                        ? 'bg-[#2F6F4E] text-white border-[#2F6F4E] shadow-sm'
                        : 'bg-[#F5F8FB] text-[#5A6A7A] hover:bg-[#E8EEF4] hover:text-[#1A2332] border-[#D5DEE8]'
                    }`}
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                    <span>Review Report ({issues.length})</span>
                  </button>

                  <button
                    onClick={() => setActiveTab('PANELS')}
                    className={`px-4 py-2 text-xs font-bold rounded-full border transition-all flex items-center space-x-2 ${
                      activeTab === 'PANELS'
                        ? 'bg-[#2F6F4E] text-white border-[#2F6F4E] shadow-sm'
                        : 'bg-[#F5F8FB] text-[#5A6A7A] hover:bg-[#E8EEF4] hover:text-[#1A2332] border-[#D5DEE8]'
                    }`}
                  >
                    <Cpu className="w-4 h-4" />
                    <span>Panel Schedules ({panels.length})</span>
                  </button>

                  <button
                    onClick={() => setActiveTab('SPECS')}
                    className={`px-4 py-2 text-xs font-bold rounded-full border transition-all flex items-center space-x-2 ${
                      activeTab === 'SPECS'
                        ? 'bg-[#2F6F4E] text-white border-[#2F6F4E] shadow-sm'
                        : 'bg-[#F5F8FB] text-[#5A6A7A] hover:bg-[#E8EEF4] hover:text-[#1A2332] border-[#D5DEE8]'
                    }`}
                  >
                    <Table className="w-4 h-4" />
                    <span>
                      Cable Specs ({specs.length} / OD {outerDias.length} / Pipe {conduits.length} / In{' '}
                      {cbRatings.length} / CB {cbTypes.length})
                    </span>
                  </button>
                </div>

                <div className="hidden xl:flex items-center space-x-1.5 text-xs text-[#6B7C8C] shrink-0 ml-2">
                  <Sparkles className="w-3.5 h-3.5 text-[#2D8A55]" />
                  <span>
                    {pendingPatchCount > 0
                      ? `${pendingPatchCount} unsaved cell change(s)`
                      : 'Professional VBA-equivalent panel review mode'}
                  </span>
                </div>
              </div>
            )}

            {panels.length === 0 && !isLoading && (
              <div className="bg-[#FFFFFF] border border-dashed border-[#C5D0DC] rounded-2xl p-10 text-center text-sm text-[#5A6A7A]">
                Nạp file Excel từ task panel bên trái để bắt đầu kiểm tra.
              </div>
            )}

            {panels.length > 0 && activeTab === 'REPORT' && (
              <ReportTable
                issues={issues}
                panelNames={panelNames}
                onUpdateIssueStatus={handleUpdateIssueStatus}
                onBatchUpdateStatus={handleBatchUpdateStatus}
              />
            )}

            {panels.length > 0 && activeTab === 'PANELS' && (
              <PanelViewer
                panels={panels}
                issues={issues}
                cbRatings={cbRatings}
                cbTypes={cbTypes}
                poleOptions={poleOptions}
                iscOptions={iscOptions}
                onUpdateCircuit={handleUpdateCircuit}
              />
            )}

            {panels.length > 0 && activeTab === 'SPECS' && (
              <SpecEditor
                specs={specs}
                onUpdateSpecs={(newSpecs) => setSpecs(newSpecs)}
                outerDias={outerDias}
                onUpdateOuterDias={(rows) => setOuterDias(rows)}
                conduits={conduits}
                onUpdateConduits={(rows) => setConduits(rows)}
                cbRatings={cbRatings}
                onUpdateCbRatings={(rows) => setCbRatings(rows)}
                cbTypes={cbTypes}
                onUpdateCbTypes={(rows) => setCbTypes(rows)}
                poleOptions={poleOptions}
                onUpdatePoleOptions={(rows) => setPoleOptions(rows)}
                iscOptions={iscOptions}
                onUpdateIscOptions={(rows) => setIscOptions(rows)}
                specSheetName={specSheetName}
                onRegisterPatch={registerPatch}
                conduitFillPercent={config.conduitFillPercent}
                onUpdateConduitFillPercent={(pct) =>
                  setConfig((prev) => ({ ...prev, conduitFillPercent: pct }))
                }
                onReloadLookupTables={handleReloadLookupTables}
                lookupFromFile={lookupInfo.fromFile}
                lookupUpdatedAt={lookupInfo.updatedAt}
              />
            )}
          </section>
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-[#FFFFFF] border-t border-[#D5DEE8] py-4 text-center text-xs text-[#6B7C8C] mt-auto">
        <div className="w-full px-3 sm:px-4 lg:px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>
            © Electrical Panel Reviewer — Web app for electrical load schedule review (Vietnamese &amp; IEC standards).
          </span>
          <span className="text-[#2D8A55] font-medium">
            Auto checks: Rule 13 (Isc ≥ 65kA &amp; In ≥ 32A — MSB only), distribution Isc limits, CB selection, cable &amp; conduit.
          </span>
        </div>
      </footer>

      {/* Modals */}
      <ConfigModal
        isOpen={isConfigOpen}
        onClose={() => setIsConfigOpen(false)}
        config={config}
        onSaveConfig={(newCfg) => setConfig(newCfg)}
      />

      <RulesGuideModal
        isOpen={isRulesOpen}
        onClose={() => setIsRulesOpen(false)}
      />
    </div>
  );
}
