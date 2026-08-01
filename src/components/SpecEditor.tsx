import React, { useMemo, useState } from 'react';
import {
  CableOuterDiaRow,
  CableSheathType,
  CableSpecRow,
  CABLE_SHEATH_TYPES,
  ConduitSpec,
} from '../types';
import { Table, Plus, Trash2, RotateCcw, Ruler, ChevronDown, ChevronRight, Circle, Columns3, Zap } from 'lucide-react';
import { DEFAULT_CABLE_SPECS } from '../utils/specTable';
import {
  DEFAULT_CABLE_OUTER_DIAS,
  collectSheathColumns,
  formatSheathColumnName,
  groupOuterDiaByCore,
  normalizeSheathKey,
} from '../utils/outerDiaTable';
import { DEFAULT_CONDUITS, makeConduitSpec, sortConduits } from '../utils/conduitTable';
import { CellPatch } from '../utils/excelPatch';
import {
  CbRatingItem,
  DEFAULT_CB_RATINGS,
  formatCbRatingLabel,
  sortCbRatings,
} from '../utils/cbRatingTable';
import {
  DEFAULT_CB_TYPES,
  DEFAULT_ISC_OPTIONS,
  DEFAULT_POLE_OPTIONS,
  formatCbTypeLabel,
  formatIscLabel,
  formatPoleLabel,
  SpecListItem,
  uniqueSpecList,
} from '../utils/cbOptionLists';

interface SpecEditorProps {
  specs: CableSpecRow[];
  onUpdateSpecs: (newSpecs: CableSpecRow[]) => void;
  outerDias: CableOuterDiaRow[];
  onUpdateOuterDias: (rows: CableOuterDiaRow[]) => void;
  conduits: ConduitSpec[];
  onUpdateConduits: (rows: ConduitSpec[]) => void;
  cbRatings: CbRatingItem[];
  onUpdateCbRatings: (rows: CbRatingItem[]) => void;
  cbTypes: SpecListItem[];
  onUpdateCbTypes: (rows: SpecListItem[]) => void;
  poleOptions: SpecListItem[];
  onUpdatePoleOptions: (rows: SpecListItem[]) => void;
  iscOptions: SpecListItem[];
  onUpdateIscOptions: (rows: SpecListItem[]) => void;
  specSheetName?: string | null;
  /** Đăng ký patch ô Excel khi sửa Spec / OD / ống */
  onRegisterPatch?: (patch: CellPatch) => void;
  /** Tỉ lệ lấp đầy ống luồn tối đa cho phép (%) */
  conduitFillPercent: number;
  onUpdateConduitFillPercent: (pct: number) => void;
  /** Nạp lại bảng tra từ public/data/Spec. Cable.xlsx */
  onReloadLookupTables?: () => void;
  /** Bảng tra hiện đang lấy từ Spec. Cable.xlsx hay từ code */
  lookupFromFile?: boolean;
  lookupUpdatedAt?: string;
}

/** Nhãn hiển thị cột vỏ (mặc định + tùy chỉnh) */
function sheathLabel(s: string): string {
  const defaults: Record<string, string> = {
    'CU/PVC': 'CU/PVC',
    'CU/XLPE/PVC': 'CU/XLPE/PVC',
    'CU/MICA/XLPE/FR-PVC': 'Cu/Mica/XLPE/FR-PVC',
    'CU/MICA/XLPE/LSZH': 'Cu/Mica/XLPE/LSZH',
    'CU/PVC/PVC': 'CU/PVC/PVC',
  };
  return defaults[s] || s;
}

function isBuiltinSheath(s: string): boolean {
  return (CABLE_SHEATH_TYPES as readonly string[]).some(
    (b) => normalizeSheathKey(b) === normalizeSheathKey(s)
  );
}

export const SpecEditor: React.FC<SpecEditorProps> = ({
  specs,
  onUpdateSpecs,
  outerDias,
  onUpdateOuterDias,
  conduits,
  onUpdateConduits,
  cbRatings,
  onUpdateCbRatings,
  cbTypes,
  onUpdateCbTypes,
  poleOptions,
  onUpdatePoleOptions,
  iscOptions,
  onUpdateIscOptions,
  specSheetName,
  onRegisterPatch,
  conduitFillPercent,
  onUpdateConduitFillPercent,
  onReloadLookupTables,
  lookupFromFile = false,
  lookupUpdatedAt,
}) => {
  const [editingSpecs, setEditingSpecs] = useState<CableSpecRow[]>([...specs]);
  const [editingOuterDias, setEditingOuterDias] = useState<CableOuterDiaRow[]>([...outerDias]);
  const [editingConduits, setEditingConduits] = useState<ConduitSpec[]>([...conduits]);
  const [editingCbRatings, setEditingCbRatings] = useState<CbRatingItem[]>([...cbRatings]);
  const [editingCbTypes, setEditingCbTypes] = useState<SpecListItem[]>([...cbTypes]);
  const [editingPoles, setEditingPoles] = useState<SpecListItem[]>([...poleOptions]);
  const [editingIsc, setEditingIsc] = useState<SpecListItem[]>([...iscOptions]);
  const [newCbAmp, setNewCbAmp] = useState<number>(0);
  const [newPhaseText, setNewPhaseText] = useState<string>('');
  const [newPeText, setNewPeText] = useState<string>('');
  const [newOdCore, setNewOdCore] = useState<number>(1);
  const [newOdSection, setNewOdSection] = useState<number>(0);
  /** Cột vỏ vừa thêm (chưa có dữ liệu OD) — giữ để hiện trên bảng */
  const [extraSheathColumns, setExtraSheathColumns] = useState<string[]>([]);
  const [newSheathName, setNewSheathName] = useState('');
  const [newConduitOuter, setNewConduitOuter] = useState<number>(0);
  const [newConduitInner, setNewConduitInner] = useState<number>(0);
  const [newConduitThick, setNewConduitThick] = useState<number>(0);
  const [newCbRatingAmp, setNewCbRatingAmp] = useState<number>(0);
  const [newCbType, setNewCbType] = useState('');
  const [newPole, setNewPole] = useState('');
  const [newIsc, setNewIsc] = useState('');
  // Mặc định thu gọn để tiết kiệm diện tích; mở khi cần chỉnh sửa
  const [specExpanded, setSpecExpanded] = useState(false);
  const [odExpanded, setOdExpanded] = useState(false);
  const [conduitExpanded, setConduitExpanded] = useState(false);
  const [cbRatingExpanded, setCbRatingExpanded] = useState(false);
  const [cbOptsExpanded, setCbOptsExpanded] = useState(false);

  React.useEffect(() => {
    setEditingSpecs([...specs]);
  }, [specs]);

  React.useEffect(() => {
    setEditingOuterDias([...outerDias]);
  }, [outerDias]);

  React.useEffect(() => {
    setEditingConduits([...conduits]);
  }, [conduits]);

  React.useEffect(() => {
    setEditingCbRatings([...cbRatings]);
  }, [cbRatings]);

  React.useEffect(() => {
    setEditingCbTypes([...cbTypes]);
  }, [cbTypes]);

  React.useEffect(() => {
    setEditingPoles([...poleOptions]);
  }, [poleOptions]);

  React.useEffect(() => {
    setEditingIsc([...iscOptions]);
  }, [iscOptions]);

  const outerDiaGroups = useMemo(() => groupOuterDiaByCore(editingOuterDias), [editingOuterDias]);
  const sheathColumns = useMemo(
    () => collectSheathColumns(editingOuterDias, extraSheathColumns),
    [editingOuterDias, extraSheathColumns]
  );
  const usableFillPct = conduitFillPercent > 0 ? conduitFillPercent : 35;

  const handleFieldChange = (index: number, field: keyof CableSpecRow, value: string | number) => {
    const updated = [...editingSpecs];
    const row = updated[index];
    updated[index] = {
      ...row,
      [field]: field === 'cbAmp' ? Number(value) : String(value),
    };
    const sorted = [...updated].sort((a, b) => a.cbAmp - b.cbAmp);
    setEditingSpecs(sorted);
    onUpdateSpecs(sorted);

    // Ghi ngược cột A/B/C nếu có tọa độ Excel
    if (onRegisterPatch && row.excelSheet != null && row.excelRow != null) {
      const col = field === 'cbAmp' ? 0 : field === 'phaseText' ? 1 : field === 'peText' ? 2 : -1;
      if (col >= 0) {
        onRegisterPatch({
          sheetName: row.excelSheet,
          row: row.excelRow,
          col,
          value: field === 'cbAmp' ? Number(value) : String(value),
        });
      }
    }
  };

  const handleAddRow = (e: React.FormEvent) => {
    e.preventDefault();
    if (newCbAmp <= 0) return;

    const updated = [
      ...editingSpecs,
      { cbAmp: newCbAmp, phaseText: newPhaseText || '1.5', peText: newPeText || '1.5' },
    ].sort((a, b) => a.cbAmp - b.cbAmp);

    setEditingSpecs(updated);
    onUpdateSpecs(updated);
    setNewCbAmp(0);
    setNewPhaseText('');
    setNewPeText('');
  };

  const handleDeleteRow = (index: number) => {
    const updated = editingSpecs.filter((_, i) => i !== index);
    setEditingSpecs(updated);
    onUpdateSpecs(updated);
  };

  const handleResetToDefault = () => {
    setEditingSpecs([...DEFAULT_CABLE_SPECS]);
    onUpdateSpecs([...DEFAULT_CABLE_SPECS]);
  };

  const commitOuterDias = (updated: CableOuterDiaRow[]) => {
    const sorted = [...updated].sort(
      (a, b) => a.coreCount - b.coreCount || a.sectionMM2 - b.sectionMM2
    );
    setEditingOuterDias(sorted);
    onUpdateOuterDias(sorted);
  };

  const handleOdChange = (
    coreCount: number,
    sectionMM2: number,
    sheath: CableSheathType,
    rawValue: string
  ) => {
    const updated = editingOuterDias.map((r) => ({ ...r, odBySheath: { ...r.odBySheath } }));
    const idx = updated.findIndex(
      (r) => r.coreCount === coreCount && Math.abs(r.sectionMM2 - sectionMM2) < 0.05
    );
    if (idx < 0) return;

    const trimmed = rawValue.trim();
    let numVal: number | '' = '';
    if (trimmed === '') {
      delete updated[idx].odBySheath[sheath];
    } else {
      const num = Number(trimmed.replace(',', '.'));
      if (!Number.isNaN(num) && num > 0) {
        updated[idx].odBySheath[sheath] = Math.round(num * 10) / 10;
        numVal = updated[idx].odBySheath[sheath]!;
      }
    }
    commitOuterDias(updated);

    const src = updated[idx];
    const col = src.excelSheathCols?.[sheath];
    if (onRegisterPatch && src.excelSheet != null && src.excelRow != null && col != null && numVal !== '') {
      onRegisterPatch({
        sheetName: src.excelSheet,
        row: src.excelRow,
        col,
        value: numVal,
      });
    }
  };

  const handleAddOdRow = (e: React.FormEvent) => {
    e.preventDefault();
    if (newOdCore <= 0 || newOdSection <= 0) return;
    const exists = editingOuterDias.some(
      (r) => r.coreCount === newOdCore && Math.abs(r.sectionMM2 - newOdSection) < 0.05
    );
    if (exists) return;

    commitOuterDias([
      ...editingOuterDias,
      { coreCount: newOdCore, sectionMM2: newOdSection, odBySheath: {} },
    ]);
    setNewOdSection(0);
  };

  /** Thêm cột quy cách vỏ cáp mới */
  const handleAddSheathColumn = (e: React.FormEvent) => {
    e.preventDefault();
    const name = formatSheathColumnName(newSheathName);
    if (!name) return;
    if (sheathColumns.some((s) => normalizeSheathKey(s) === normalizeSheathKey(name))) {
      alert('This sheath column already exists.');
      return;
    }
    setExtraSheathColumns((prev) => [...prev, name]);
    setNewSheathName('');
  };

  /** Xóa cột quy cách vỏ (và dữ liệu OD tương ứng) */
  const handleDeleteSheathColumn = (sheath: string) => {
    const updated = editingOuterDias.map((r) => {
      const od = { ...r.odBySheath };
      for (const k of Object.keys(od)) {
        if (normalizeSheathKey(k) === normalizeSheathKey(sheath)) {
          delete od[k];
        }
      }
      return { ...r, odBySheath: od };
    });
    commitOuterDias(updated);
    setExtraSheathColumns((prev) =>
      prev.filter((s) => normalizeSheathKey(s) !== normalizeSheathKey(sheath))
    );
  };

  const handleDeleteOdRow = (coreCount: number, sectionMM2: number) => {
    commitOuterDias(
      editingOuterDias.filter(
        (r) => !(r.coreCount === coreCount && Math.abs(r.sectionMM2 - sectionMM2) < 0.05)
      )
    );
  };

  const handleResetOuterDias = () => {
    commitOuterDias([...DEFAULT_CABLE_OUTER_DIAS]);
    setExtraSheathColumns([]);
  };

  const commitConduits = (updated: ConduitSpec[]) => {
    const sorted = sortConduits(updated.map((c) => makeConduitSpec(c)));
    setEditingConduits(sorted);
    onUpdateConduits(sorted);
  };

  const handleConduitFieldChange = (
    index: number,
    field: 'outerDiaMM' | 'wallThicknessMM' | 'innerDiaMM',
    rawValue: string
  ) => {
    const updated = editingConduits.map((c) => ({ ...c }));
    const num = Number(rawValue.replace(',', '.'));
    if (Number.isNaN(num) || num < 0) return;

    if (field === 'outerDiaMM') {
      updated[index].outerDiaMM = num;
      updated[index].label = `D${Math.round(num)}`;
    } else if (field === 'wallThicknessMM') {
      updated[index].wallThicknessMM = num;
    } else {
      updated[index].innerDiaMM = num;
    }
    commitConduits(updated);

    const src = updated[index];
    if (onRegisterPatch && src.excelSheet != null && src.excelRow != null) {
      const col =
        field === 'outerDiaMM'
          ? src.excelOuterCol
          : field === 'wallThicknessMM'
          ? src.excelThickCol
          : src.excelInnerCol;
      if (col != null) {
        onRegisterPatch({
          sheetName: src.excelSheet,
          row: src.excelRow,
          col,
          value: num,
        });
      }
    }
  };

  const handleAddConduit = (e: React.FormEvent) => {
    e.preventDefault();
    if (newConduitOuter <= 0 || newConduitInner <= 0) return;
    const label = `D${Math.round(newConduitOuter)}`;
    if (editingConduits.some((c) => c.label === label && c.material === 'PVC')) return;

    commitConduits([
      ...editingConduits,
      makeConduitSpec({
        label,
        material: 'PVC',
        outerDiaMM: newConduitOuter,
        wallThicknessMM: newConduitThick > 0 ? newConduitThick : undefined,
        innerDiaMM: newConduitInner,
      }),
    ]);
    setNewConduitOuter(0);
    setNewConduitInner(0);
    setNewConduitThick(0);
  };

  const handleDeleteConduit = (index: number) => {
    commitConduits(editingConduits.filter((_, i) => i !== index));
  };

  const handleResetConduits = () => {
    commitConduits([...DEFAULT_CONDUITS]);
  };

  const commitCbRatings = (updated: CbRatingItem[]) => {
    const sorted = sortCbRatings(updated);
    setEditingCbRatings(sorted);
    onUpdateCbRatings(sorted);
  };

  const handleAddCbRating = (e: React.FormEvent) => {
    e.preventDefault();
    if (newCbRatingAmp <= 0) return;
    const label = formatCbRatingLabel(newCbRatingAmp);
    if (editingCbRatings.some((r) => r.amp === newCbRatingAmp)) return;
    commitCbRatings([...editingCbRatings, { label, amp: newCbRatingAmp }]);
    setNewCbRatingAmp(0);
  };

  const handleCbRatingLabelChange = (index: number, raw: string) => {
    const label = formatCbRatingLabel(raw);
    const amp = Number(label.replace(/A/i, '')) || 0;
    if (amp <= 0) return;
    const updated = editingCbRatings.map((r, i) =>
      i === index ? { ...r, label, amp } : r
    );
    const row = editingCbRatings[index];
    commitCbRatings(updated);
    if (onRegisterPatch && row?.excelSheet != null && row.excelRow != null && row.excelCol != null) {
      onRegisterPatch({
        sheetName: row.excelSheet,
        row: row.excelRow,
        col: row.excelCol,
        value: label,
      });
    }
  };

  const handleDeleteCbRating = (index: number) => {
    commitCbRatings(editingCbRatings.filter((_, i) => i !== index));
  };

  const handleResetCbRatings = () => {
    commitCbRatings([...DEFAULT_CB_RATINGS]);
  };

  const commitList = (
    kind: 'type' | 'pole' | 'isc',
    updated: SpecListItem[]
  ) => {
    if (kind === 'type') {
      const next = uniqueSpecList(updated, formatCbTypeLabel);
      setEditingCbTypes(next);
      onUpdateCbTypes(next);
    } else if (kind === 'pole') {
      const next = uniqueSpecList(updated, formatPoleLabel);
      setEditingPoles(next);
      onUpdatePoleOptions(next);
    } else {
      const next = uniqueSpecList(updated, formatIscLabel);
      setEditingIsc(next);
      onUpdateIscOptions(next);
    }
  };

  const handleAddListItem = (kind: 'type' | 'pole' | 'isc', raw: string) => {
    const normalize =
      kind === 'type' ? formatCbTypeLabel : kind === 'pole' ? formatPoleLabel : formatIscLabel;
    const label = normalize(raw);
    if (!label) return;
    const list =
      kind === 'type' ? editingCbTypes : kind === 'pole' ? editingPoles : editingIsc;
    if (list.some((x) => normalize(x.label) === label)) return;
    commitList(kind, [...list, { label }]);
  };

  const handleDeleteListItem = (kind: 'type' | 'pole' | 'isc', index: number) => {
    const list =
      kind === 'type' ? editingCbTypes : kind === 'pole' ? editingPoles : editingIsc;
    commitList(
      kind,
      list.filter((_, i) => i !== index)
    );
  };

  const handleResetCbOpts = () => {
    commitList('type', [...DEFAULT_CB_TYPES]);
    commitList('pole', [...DEFAULT_POLE_OPTIONS]);
    commitList('isc', [...DEFAULT_ISC_OPTIONS]);
  };

  return (
    <div className="space-y-4">
      {/* Nguồn bảng tra + thao tác nạp lại */}
      {onReloadLookupTables && (
        <div className="bg-[#F5F8FB] border border-[#D5DEE8] rounded-2xl px-4 py-3 flex flex-wrap items-center justify-between gap-3">
          <div className="min-w-0 text-xs text-[#5A6A7A]">
            <span className="font-bold text-[#1A2332] uppercase tracking-wider text-[11px]">
              Lookup source
            </span>
            <div className="mt-0.5 leading-snug">
              {lookupFromFile ? (
                <>
                  Using{' '}
                  <strong className="text-[#1B7A45] font-mono">public/data/Spec. Cable.xlsx</strong>
                  {lookupUpdatedAt ? ` (loaded ${lookupUpdatedAt})` : ''}
                  <span className="block text-[11px] text-[#5A6A7A] mt-0.5">
                    Sheets: CB_SPEC, CB_OPTIONS, OUTER_DIA, CONDUIT. Edit the Excel file, then reload.
                    Panel review always uses this app table — Spec sheets inside the audited workbook are ignored.
                  </span>
                </>
              ) : (
                <span className="text-[#B45309]">
                  Using built-in defaults (could not load Spec. Cable.xlsx)
                </span>
              )}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              onClick={onReloadLookupTables}
              title="Reload public/data/Spec. Cable.xlsx and apply immediately"
              className="px-3.5 py-1.5 text-xs font-semibold bg-[#2F6F4E] hover:bg-[#255A3F] text-white rounded-full transition-colors flex items-center space-x-1.5"
            >
              <RotateCcw className="w-3.5 h-3.5" />
                <span>Tải lại bảng tra</span>
              </button>
          </div>
        </div>
      )}

      {/* ========== BẢNG 1: CB → tiết diện ========== */}
      <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm">
        <div
          className={`p-4 bg-[#F5F8FB] flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
            specExpanded ? 'border-b border-[#D5DEE8]' : ''
          }`}
        >
          <button
            type="button"
            onClick={() => setSpecExpanded((v) => !v)}
            className="flex items-center space-x-2 text-left min-w-0 flex-1 group"
            aria-expanded={specExpanded}
          >
            <span className="p-1.5 rounded-lg text-[#1B7A45] bg-[#E6F4EC] border border-[#A8D4B8] group-hover:bg-[#D1FAE5] transition-colors shrink-0">
              {specExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
            </span>
            <div className="p-2 bg-[#E6F4EC] text-[#1B7A45] rounded-xl border border-[#A8D4B8] shrink-0">
              <Table className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider flex flex-wrap items-center gap-2">
                <span>1. Cable Cross-Section by CB Rating</span>
                <span className="font-mono font-semibold normal-case tracking-normal text-[#2D8A55] bg-[#F0F4F8] border border-[#C5D0DC] rounded-full px-2 py-0.5">
                  {editingSpecs.length} rows
                </span>
                <span className="text-[10px] font-semibold normal-case tracking-normal text-[#6B7C8C]">
                  {specExpanded ? 'Collapse' : 'Expand'}
                </span>
              </h2>
              <p className="text-xs text-[#5A6A7A] truncate">
                {specSheetName
                  ? `Loaded from Excel sheet [${specSheetName}] (columns A–C)`
                  : 'Using default TCVN / IEC lookup table'}
              </p>
            </div>
          </button>

          {specExpanded && (
            <button
              onClick={handleResetToDefault}
              className="px-3.5 py-1.5 text-xs font-semibold bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] rounded-full border border-[#C5D0DC] transition-colors flex items-center space-x-1.5 shrink-0 self-start sm:self-auto"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#2D8A55]" />
              <span>Reset Default</span>
            </button>
          )}
        </div>

        {specExpanded && (
          <>
            <form
              onSubmit={handleAddRow}
              className="p-4 bg-[#F5F8FB]/60 border-b border-[#D5DEE8] flex flex-wrap items-center gap-3 text-xs"
            >
              <span className="font-semibold text-[#1A2332]">Add CB rule:</span>
              <input
                type="number"
                placeholder="CB In (A)"
                value={newCbAmp || ''}
                onChange={(e) => setNewCbAmp(Number(e.target.value))}
                className="w-28 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
              />
              <input
                type="text"
                placeholder="Phase section (e.g. 2.5)"
                value={newPhaseText}
                onChange={(e) => setNewPhaseText(e.target.value)}
                className="w-48 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
              />
              <input
                type="text"
                placeholder="PE section (e.g. 2.5)"
                value={newPeText}
                onChange={(e) => setNewPeText(e.target.value)}
                className="w-36 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#1B7A45] font-mono"
              />
              <button
                type="submit"
                className="px-4 py-1.5 bg-[#2F6F4E] hover:bg-[#255A3F] text-white font-semibold rounded-full transition-colors flex items-center space-x-1"
              >
                <Plus className="w-4 h-4" />
                <span>Add row</span>
              </button>
            </form>

            <div className="overflow-x-auto max-h-[420px]">
              <table className="w-full text-left text-xs text-[#1A2332] border-collapse">
                <thead className="bg-[#E8EEF4] text-[#1B7A45] uppercase font-bold border-b border-[#D5DEE8] text-[11px] sticky top-0 z-10">
                  <tr>
                    <th className="py-2.5 px-4 w-12 text-center font-bold">#</th>
                    <th className="py-2.5 px-4 w-36 text-right font-bold">CB Rating In (A)</th>
                    <th className="py-2.5 px-4 font-bold">Phase Cable (mm²)</th>
                    <th className="py-2.5 px-4 font-bold">PE Cable (mm²)</th>
                    <th className="py-2.5 px-4 w-20 text-center font-bold">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#D5DEE8] font-mono text-[11px]">
                  {editingSpecs.map((spec, idx) => (
                    <tr key={idx} className="hover:bg-[#F5F8FB] transition-colors">
                      <td className="py-2 px-4 text-center text-[#5A6A7A]">{idx + 1}</td>
                      <td className="py-2 px-4 text-right font-bold text-[#1B7A45]">
                        <input
                          type="number"
                          value={spec.cbAmp}
                          onChange={(e) => handleFieldChange(idx, 'cbAmp', e.target.value)}
                          className="w-24 bg-[#F0F4F8] border border-[#C5D0DC] text-right text-[#1B7A45] font-bold rounded px-2 py-1 focus:outline-none focus:border-[#1B7A45]"
                        />
                        <span className="ml-1 text-[#5A6A7A]">A</span>
                      </td>
                      <td className="py-2 px-4">
                        <input
                          type="text"
                          value={spec.phaseText}
                          onChange={(e) => handleFieldChange(idx, 'phaseText', e.target.value)}
                          className="w-full max-w-xs bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded px-2 py-1 focus:outline-none focus:border-[#1B7A45]"
                        />
                      </td>
                      <td className="py-2 px-4">
                        <input
                          type="text"
                          value={spec.peText}
                          onChange={(e) => handleFieldChange(idx, 'peText', e.target.value)}
                          className="w-full max-w-xs bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded px-2 py-1 focus:outline-none focus:border-[#1B7A45]"
                        />
                      </td>
                      <td className="py-2 px-4 text-center">
                        <button
                          onClick={() => handleDeleteRow(idx)}
                          title="Delete this row"
                          className="p-1 text-[#5A6A7A] hover:text-[#DC2626] hover:bg-[#FEE2E2] rounded transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      {/* ========== BẢNG 2: OD theo tiết diện / lõi / vỏ ========== */}
      <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm">
        <div
          className={`p-4 bg-[#F5F8FB] flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
            odExpanded ? 'border-b border-[#D5DEE8]' : ''
          }`}
        >
          <button
            type="button"
            onClick={() => setOdExpanded((v) => !v)}
            className="flex items-center space-x-2 text-left min-w-0 flex-1 group"
            aria-expanded={odExpanded}
          >
            <span className="p-1.5 rounded-lg text-[#0369A1] bg-[#E0F2FE] border border-[#7DD3FC] group-hover:bg-[#BAE6FD] transition-colors shrink-0">
              {odExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
            </span>
            <div className="p-2 bg-[#E0F2FE] text-[#0369A1] rounded-xl border border-[#7DD3FC] shrink-0">
              <Ruler className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider flex flex-wrap items-center gap-2">
                <span>2. Cable Outer Diameter (OD) by Section &amp; Sheath</span>
                <span className="font-mono font-semibold normal-case tracking-normal text-[#0369A1] bg-[#F0F4F8] border border-[#C5D0DC] rounded-full px-2 py-0.5">
                  {editingOuterDias.length} rows
                </span>
                <span className="text-[10px] font-semibold normal-case tracking-normal text-[#6B7C8C]">
                  {odExpanded ? 'Collapse' : 'Expand'}
                </span>
              </h2>
              <p className="text-xs text-[#5A6A7A] truncate">
                {specSheetName
                  ? `Loaded from Excel sheet [${specSheetName}] (columns P–U)`
                  : 'Using default catalogue OD table — used for conduit fill check'}
              </p>
            </div>
          </button>

          {odExpanded && (
            <button
              onClick={handleResetOuterDias}
              className="px-3.5 py-1.5 text-xs font-semibold bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] rounded-full border border-[#C5D0DC] transition-colors flex items-center space-x-1.5 shrink-0 self-start sm:self-auto"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#0369A1]" />
              <span>Reset Default OD</span>
            </button>
          )}
        </div>

        {odExpanded && (
          <>
            <form
              onSubmit={handleAddOdRow}
              className="p-4 bg-[#F5F8FB]/60 border-b border-[#D5DEE8] flex flex-wrap items-center gap-3 text-xs"
            >
              <span className="font-semibold text-[#1A2332]">Add OD row:</span>
              <select
                value={newOdCore}
                onChange={(e) => setNewOdCore(Number(e.target.value))}
                className="bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#0369A1] font-mono"
              >
                {[1, 2, 3, 4, 5].map((c) => (
                  <option key={c} value={c}>
                    {c}C
                  </option>
                ))}
              </select>
              <input
                type="number"
                step="0.1"
                placeholder="Section (mm²)"
                value={newOdSection || ''}
                onChange={(e) => setNewOdSection(Number(e.target.value))}
                className="w-32 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#0369A1] font-mono"
              />
              <button
                type="submit"
                className="px-4 py-1.5 bg-[#0284C7] hover:bg-[#0369A1] text-white font-semibold rounded-full transition-colors flex items-center space-x-1"
              >
                <Plus className="w-4 h-4" />
                <span>Add row</span>
              </button>
            </form>

            {/* Thêm cột quy cách vỏ cáp mới */}
            <form
              onSubmit={handleAddSheathColumn}
              className="px-4 py-3 bg-[#ECFEFF]/70 border-b border-[#D5DEE8] flex flex-wrap items-center gap-3 text-xs"
            >
              <span className="font-semibold text-[#1A2332] flex items-center gap-1.5">
                <Columns3 className="w-3.5 h-3.5 text-[#0369A1]" />
                Add sheath column:
              </span>
              <input
                type="text"
                placeholder="e.g. CU/XLPE/LSZH"
                value={newSheathName}
                onChange={(e) => setNewSheathName(e.target.value)}
                className="w-48 sm:w-56 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#0369A1] font-mono"
              />
              <button
                type="submit"
                className="px-4 py-1.5 bg-[#0E7490] hover:bg-[#155E75] text-white font-semibold rounded-full transition-colors flex items-center space-x-1"
              >
                <Plus className="w-4 h-4" />
                <span>Add column</span>
              </button>
              <span className="text-[10px] text-[#6B7C8C]">
                New cable sheath type → new OD column
              </span>
            </form>

            <div className="overflow-x-auto max-h-[560px]">
              {[...outerDiaGroups.keys()]
                .sort((a, b) => a - b)
                .map((core) => {
                  const rows = outerDiaGroups.get(core) || [];
                  return (
                    <div key={core}>
                      <div className="sticky top-0 z-20 bg-[#ECFEFF] border-y border-[#A5F3FC] px-4 py-1.5 text-[11px] font-bold text-[#0E7490] uppercase tracking-wider">
                        {core}C — {core === 1 ? 'Single core' : `${core}-core`} cables
                      </div>
                      <table className="w-full text-left text-xs text-[#1A2332] border-collapse">
                        <thead className="bg-[#E8EEF4] text-[#0369A1] uppercase font-bold border-b border-[#D5DEE8] text-[10px]">
                          <tr>
                            <th className="py-2 px-3 w-24 text-right font-bold">Section mm²</th>
                            {sheathColumns.map((s) => (
                              <th key={s} className="py-2 px-2 font-bold whitespace-nowrap">
                                <div className="inline-flex items-center gap-1">
                                  <span>{sheathLabel(s)}</span>
                                  {!isBuiltinSheath(s) && (
                                    <button
                                      type="button"
                                      title="Remove this sheath column"
                                      onClick={() => handleDeleteSheathColumn(s)}
                                      className="p-0.5 text-[#5A6A7A] hover:text-[#DC2626] hover:bg-[#FEE2E2] rounded transition-colors"
                                    >
                                      <Trash2 className="w-3 h-3" />
                                    </button>
                                  )}
                                </div>
                              </th>
                            ))}
                            <th className="py-2 px-3 w-16 text-center font-bold">Action</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-[#D5DEE8] font-mono text-[11px]">
                          {rows.map((r) => (
                            <tr
                              key={`${r.coreCount}-${r.sectionMM2}`}
                              className="hover:bg-[#F5F8FB] transition-colors"
                            >
                              <td className="py-1.5 px-3 text-right font-bold text-[#0369A1]">
                                {r.sectionMM2}
                              </td>
                              {sheathColumns.map((s) => (
                                <td key={s} className="py-1.5 px-2">
                                  <input
                                    type="text"
                                    inputMode="decimal"
                                    value={r.odBySheath[s] ?? ''}
                                    placeholder="—"
                                    onChange={(e) =>
                                      handleOdChange(r.coreCount, r.sectionMM2, s, e.target.value)
                                    }
                                    className="w-16 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded px-1.5 py-1 text-center focus:outline-none focus:border-[#0369A1]"
                                  />
                                </td>
                              ))}
                              <td className="py-1.5 px-3 text-center">
                                <button
                                  onClick={() => handleDeleteOdRow(r.coreCount, r.sectionMM2)}
                                  title="Delete this OD row"
                                  className="p-1 text-[#5A6A7A] hover:text-[#DC2626] hover:bg-[#FEE2E2] rounded transition-colors"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  );
                })}
              {editingOuterDias.length === 0 && (
                <div className="p-8 text-center text-xs text-[#6B7C8C]">
                  No outer diameter rows. Click &quot;Reset Default OD&quot; or add a row.
                </div>
              )}
            </div>
          </>
        )}
      </div>

      {/* ========== BẢNG 3: Ống luồn — đường kính trong ========== */}
      <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm">
        <div
          className={`p-4 bg-[#F5F8FB] flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
            conduitExpanded ? 'border-b border-[#D5DEE8]' : ''
          }`}
        >
          <button
            type="button"
            onClick={() => setConduitExpanded((v) => !v)}
            className="flex items-center space-x-2 text-left min-w-0 flex-1 group"
            aria-expanded={conduitExpanded}
          >
            <span className="p-1.5 rounded-lg text-[#B45309] bg-[#FEF3C7] border border-[#F59E0B] group-hover:bg-[#B45309] transition-colors shrink-0">
              {conduitExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
            </span>
            <div className="p-2 bg-[#FEF3C7] text-[#B45309] rounded-xl border border-[#F59E0B] shrink-0">
              <Circle className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider flex flex-wrap items-center gap-2">
                <span>3. Conduit Inner Diameter (PVC Pipe Sizes)</span>
                <span className="font-mono font-semibold normal-case tracking-normal text-[#B45309] bg-[#F0F4F8] border border-[#C5D0DC] rounded-full px-2 py-0.5">
                  {editingConduits.length} sizes
                </span>
                <span className="text-[10px] font-semibold normal-case tracking-normal text-[#6B7C8C]">
                  {conduitExpanded ? 'Collapse' : 'Expand'}
                </span>
              </h2>
              <p className="text-xs text-[#5A6A7A] truncate">
                {specSheetName
                  ? `Loaded from Excel sheet [${specSheetName}] (PVC conduit columns) — fill limit ${usableFillPct}%`
                  : `Default PVC conduit ID table — basis for conduit selection (max fill ${usableFillPct}%)`}
              </p>
            </div>
          </button>

          {conduitExpanded && (
            <div className="flex flex-wrap items-center gap-2 shrink-0 self-start sm:self-auto">
              <label
                className="flex items-center gap-2 px-3 py-1.5 bg-[#FFFBEB] border border-[#F59E0B] rounded-full text-xs"
                title="Tỉ lệ luồn dây tối đa trong ống. Đổi giá trị này sẽ kiểm tra lại toàn bộ các mạch."
              >
                <span className="font-semibold text-[#1A2332] whitespace-nowrap">
                  Tỉ lệ lấp đầy tối đa
                </span>
                <input
                  type="number"
                  min={1}
                  max={100}
                  step={1}
                  value={conduitFillPercent}
                  onChange={(e) => {
                    const v = Number(e.target.value);
                    if (!Number.isFinite(v)) return;
                    onUpdateConduitFillPercent(Math.min(100, Math.max(1, v)));
                  }}
                  className="w-16 bg-white border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2 py-1 text-right font-mono font-bold focus:outline-none focus:border-[#B45309]"
                />
                <span className="font-bold text-[#B45309]">%</span>
              </label>

              <button
                onClick={handleResetConduits}
                className="px-3.5 py-1.5 text-xs font-semibold bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] rounded-full border border-[#C5D0DC] transition-colors flex items-center space-x-1.5"
              >
                <RotateCcw className="w-3.5 h-3.5 text-[#B45309]" />
                <span>Reset Default Pipes</span>
              </button>
            </div>
          )}
        </div>

        {conduitExpanded && (
          <>
            <form
              onSubmit={handleAddConduit}
              className="p-4 bg-[#F5F8FB]/60 border-b border-[#D5DEE8] flex flex-wrap items-center gap-3 text-xs"
            >
              <span className="font-semibold text-[#1A2332]">Add conduit:</span>
              <input
                type="number"
                step="0.1"
                placeholder="Outer Ø (mm)"
                value={newConduitOuter || ''}
                onChange={(e) => setNewConduitOuter(Number(e.target.value))}
                className="w-28 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#B45309] font-mono"
              />
              <input
                type="number"
                step="0.01"
                placeholder="Wall (mm)"
                value={newConduitThick || ''}
                onChange={(e) => setNewConduitThick(Number(e.target.value))}
                className="w-24 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#B45309] font-mono"
              />
              <input
                type="number"
                step="0.1"
                placeholder="Inner Ø (mm)"
                value={newConduitInner || ''}
                onChange={(e) => setNewConduitInner(Number(e.target.value))}
                className="w-28 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#B45309] font-mono"
              />
              <button
                type="submit"
                className="px-4 py-1.5 bg-[#B45309] hover:bg-[#92400E] text-white font-semibold rounded-full transition-colors flex items-center space-x-1"
              >
                <Plus className="w-4 h-4" />
                <span>Add row</span>
              </button>
            </form>

            <div className="overflow-x-auto max-h-[420px]">
              <table className="w-full text-left text-xs text-[#1A2332] border-collapse">
                <thead className="bg-[#E8EEF4] text-[#B45309] uppercase font-bold border-b border-[#D5DEE8] text-[11px] sticky top-0 z-10">
                  <tr>
                    <th className="py-2.5 px-4 w-12 text-center font-bold">#</th>
                    <th className="py-2.5 px-4 font-bold">Label</th>
                    <th className="py-2.5 px-4 font-bold">Material</th>
                    <th className="py-2.5 px-4 text-right font-bold">Outer Ø (mm)</th>
                    <th className="py-2.5 px-4 text-right font-bold">Wall (mm)</th>
                    <th className="py-2.5 px-4 text-right font-bold">Inner Ø (mm)</th>
                    <th className="py-2.5 px-4 text-right font-bold">Inner Area (mm²)</th>
                    <th className="py-2.5 px-4 text-right font-bold">Usable @{usableFillPct}%</th>
                    <th className="py-2.5 px-4 w-20 text-center font-bold">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#D5DEE8] font-mono text-[11px]">
                  {editingConduits.map((c, idx) => (
                    <tr key={`${c.label}-${c.material}-${idx}`} className="hover:bg-[#F5F8FB] transition-colors">
                      <td className="py-2 px-4 text-center text-[#5A6A7A]">{idx + 1}</td>
                      <td className="py-2 px-4 font-bold text-[#B45309]">{c.label}</td>
                      <td className="py-2 px-4 text-[#5A6A7A]">{c.material}</td>
                      <td className="py-2 px-4 text-right">
                        <input
                          type="number"
                          step="0.1"
                          value={c.outerDiaMM}
                          onChange={(e) => handleConduitFieldChange(idx, 'outerDiaMM', e.target.value)}
                          className="w-20 bg-[#F0F4F8] border border-[#C5D0DC] text-right text-[#1A2332] rounded px-2 py-1 focus:outline-none focus:border-[#B45309]"
                        />
                      </td>
                      <td className="py-2 px-4 text-right">
                        <input
                          type="number"
                          step="0.01"
                          value={c.wallThicknessMM ?? ''}
                          placeholder="—"
                          onChange={(e) => handleConduitFieldChange(idx, 'wallThicknessMM', e.target.value)}
                          className="w-20 bg-[#F0F4F8] border border-[#C5D0DC] text-right text-[#1A2332] rounded px-2 py-1 focus:outline-none focus:border-[#B45309]"
                        />
                      </td>
                      <td className="py-2 px-4 text-right">
                        <input
                          type="number"
                          step="0.1"
                          value={c.innerDiaMM}
                          onChange={(e) => handleConduitFieldChange(idx, 'innerDiaMM', e.target.value)}
                          className="w-20 bg-[#F0F4F8] border border-[#C5D0DC] text-right text-[#B45309] font-bold rounded px-2 py-1 focus:outline-none focus:border-[#B45309]"
                        />
                      </td>
                      <td className="py-2 px-4 text-right text-[#5A6A7A]">{c.areaMM2.toFixed(1)}</td>
                      <td className="py-2 px-4 text-right text-[#2D8A55]">
                        {(c.areaMM2 * (usableFillPct / 100)).toFixed(1)}
                      </td>
                      <td className="py-2 px-4 text-center">
                        <button
                          onClick={() => handleDeleteConduit(idx)}
                          title="Delete this conduit size"
                          className="p-1 text-[#5A6A7A] hover:text-[#DC2626] hover:bg-[#FEE2E2] rounded transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {editingConduits.length === 0 && (
                <div className="p-8 text-center text-xs text-[#6B7C8C]">
                  No conduit sizes. Click &quot;Reset Default Pipes&quot; or add a row.
                </div>
              )}
            </div>
          </>
        )}
      </div>

      {/* ========== BẢNG 4: CB Rating (In) — dropdown Panel Schedule ========== */}
      <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm">
        <div
          className={`p-4 bg-[#F5F8FB] flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
            cbRatingExpanded ? 'border-b border-[#D5DEE8]' : ''
          }`}
        >
          <button
            type="button"
            onClick={() => setCbRatingExpanded((v) => !v)}
            className="flex items-center space-x-2 text-left min-w-0 flex-1 group"
            aria-expanded={cbRatingExpanded}
          >
            <span className="p-1.5 rounded-lg text-[#7C3AED] bg-[#EDE9FE] border border-[#C4B5FD] group-hover:bg-[#DDD6FE] transition-colors shrink-0">
              {cbRatingExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
            </span>
            <div className="p-2 bg-[#EDE9FE] text-[#7C3AED] rounded-xl border border-[#C4B5FD] shrink-0">
              <Zap className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider flex flex-wrap items-center gap-2">
                <span>4. CB Rating (In) — Panel dropdown list</span>
                <span className="font-mono font-semibold normal-case tracking-normal text-[#7C3AED] bg-[#F0F4F8] border border-[#C5D0DC] rounded-full px-2 py-0.5">
                  {editingCbRatings.length} ratings
                </span>
                <span className="text-[10px] font-semibold normal-case tracking-normal text-[#6B7C8C]">
                  {cbRatingExpanded ? 'Collapse' : 'Expand'}
                </span>
              </h2>
              <p className="text-xs text-[#5A6A7A] truncate">
                {specSheetName
                  ? `Loaded from [${specSheetName}] column CB_Rating (E) — used for In (A) dropdown`
                  : 'Default CB ratings — used for In (A) dropdown on Panel Schedules'}
              </p>
            </div>
          </button>

          {cbRatingExpanded && (
            <button
              onClick={handleResetCbRatings}
              className="px-3.5 py-1.5 text-xs font-semibold bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] rounded-full border border-[#C5D0DC] transition-colors flex items-center space-x-1.5 shrink-0 self-start sm:self-auto"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#7C3AED]" />
              <span>Reset Default Ratings</span>
            </button>
          )}
        </div>

        {cbRatingExpanded && (
          <>
            <form
              onSubmit={handleAddCbRating}
              className="p-4 bg-[#F5F8FB]/60 border-b border-[#D5DEE8] flex flex-wrap items-center gap-3 text-xs"
            >
              <span className="font-semibold text-[#1A2332]">Add CB rating:</span>
              <input
                type="number"
                step="1"
                placeholder="Amp (e.g. 16)"
                value={newCbRatingAmp || ''}
                onChange={(e) => setNewCbRatingAmp(Number(e.target.value))}
                className="w-32 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#7C3AED] font-mono"
              />
              <button
                type="submit"
                className="px-4 py-1.5 bg-[#7C3AED] hover:bg-[#6D28D9] text-white font-semibold rounded-full transition-colors flex items-center space-x-1"
              >
                <Plus className="w-4 h-4" />
                <span>Add rating</span>
              </button>
            </form>

            <div className="overflow-x-auto max-h-[420px]">
              <table className="w-full text-left text-xs text-[#1A2332] border-collapse">
                <thead className="bg-[#EDE9FE] text-[#7C3AED] uppercase font-bold border-b border-[#D5DEE8] text-[11px] sticky top-0 z-10">
                  <tr>
                    <th className="py-2.5 px-4 w-12 text-center font-bold">#</th>
                    <th className="py-2.5 px-4 font-bold">CB_Rating (In)</th>
                    <th className="py-2.5 px-4 text-right font-bold">Amp</th>
                    <th className="py-2.5 px-4 w-20 text-center font-bold">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#D5DEE8] font-mono text-[11px]">
                  {editingCbRatings.map((r, idx) => (
                    <tr key={`${r.amp}-${idx}`} className="hover:bg-[#F5F8FB] transition-colors">
                      <td className="py-2 px-4 text-center text-[#5A6A7A]">{idx + 1}</td>
                      <td className="py-2 px-4">
                        <input
                          type="text"
                          value={r.label}
                          onChange={(e) => handleCbRatingLabelChange(idx, e.target.value)}
                          className="w-28 bg-[#F0F4F8] border border-[#C5D0DC] text-[#1A2332] rounded px-2 py-1 focus:outline-none focus:border-[#7C3AED] font-bold text-[#7C3AED]"
                        />
                      </td>
                      <td className="py-2 px-4 text-right text-[#5A6A7A]">{r.amp}</td>
                      <td className="py-2 px-4 text-center">
                        <button
                          type="button"
                          onClick={() => handleDeleteCbRating(idx)}
                          title="Delete rating"
                          className="p-1 text-[#5A6A7A] hover:text-[#DC2626] hover:bg-[#FEE2E2] rounded transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {editingCbRatings.length === 0 && (
                <div className="p-8 text-center text-xs text-[#6B7C8C]">
                  No CB ratings. Click &quot;Reset Default Ratings&quot; or add a rating.
                </div>
              )}
            </div>
          </>
        )}
      </div>

      {/* ========== BẢNG 5: Loại CB / Số cực / Isc ========== */}
      <div className="bg-[#FFFFFF] rounded-3xl border border-[#D5DEE8] overflow-hidden shadow-sm">
        <div
          className={`p-4 bg-[#F5F8FB] flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
            cbOptsExpanded ? 'border-b border-[#D5DEE8]' : ''
          }`}
        >
          <button
            type="button"
            onClick={() => setCbOptsExpanded((v) => !v)}
            className="flex items-center space-x-2 text-left min-w-0 flex-1 group"
            aria-expanded={cbOptsExpanded}
          >
            <span className="p-1.5 rounded-lg text-[#0F766E] bg-[#CCFBF1] border border-[#5EEAD4] group-hover:bg-[#99F6E4] transition-colors shrink-0">
              {cbOptsExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
            </span>
            <div className="p-2 bg-[#CCFBF1] text-[#0F766E] rounded-xl border border-[#5EEAD4] shrink-0">
              <Table className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h2 className="text-xs font-bold text-[#1A2332] uppercase tracking-wider flex flex-wrap items-center gap-2">
                <span>5. CB Type / Pole / Isc lists</span>
                <span className="font-mono font-semibold normal-case tracking-normal text-[#0F766E] bg-[#F0F4F8] border border-[#C5D0DC] rounded-full px-2 py-0.5">
                  {editingCbTypes.length} / {editingPoles.length} / {editingIsc.length}
                </span>
                <span className="text-[10px] font-semibold normal-case tracking-normal text-[#6B7C8C]">
                  {cbOptsExpanded ? 'Collapse' : 'Expand'}
                </span>
              </h2>
              <p className="text-xs text-[#5A6A7A] truncate">
                {specSheetName
                  ? `From [${specSheetName}] — dropdowns for Loại CB, Số cực, Isc`
                  : 'Default lists — dropdowns for Loại CB, Số cực, Isc on Panel Schedules'}
              </p>
            </div>
          </button>

          {cbOptsExpanded && (
            <button
              onClick={handleResetCbOpts}
              className="px-3.5 py-1.5 text-xs font-semibold bg-[#F0F4F8] hover:bg-[#DCE4EC] text-[#1A2332] rounded-full border border-[#C5D0DC] transition-colors flex items-center space-x-1.5 shrink-0 self-start sm:self-auto"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#0F766E]" />
              <span>Reset Defaults</span>
            </button>
          )}
        </div>

        {cbOptsExpanded && (
          <div className="p-4 grid grid-cols-1 md:grid-cols-3 gap-4">
            {(
              [
                {
                  kind: 'type' as const,
                  title: 'Loại CB',
                  list: editingCbTypes,
                  value: newCbType,
                  setValue: setNewCbType,
                  placeholder: 'e.g. MCCB',
                },
                {
                  kind: 'pole' as const,
                  title: 'Số cực',
                  list: editingPoles,
                  value: newPole,
                  setValue: setNewPole,
                  placeholder: 'e.g. 3P',
                },
                {
                  kind: 'isc' as const,
                  title: 'Isc (kA)',
                  list: editingIsc,
                  value: newIsc,
                  setValue: setNewIsc,
                  placeholder: 'e.g. 65kA',
                },
              ] as const
            ).map((col) => (
              <div key={col.kind} className="border border-[#D5DEE8] rounded-2xl overflow-hidden bg-[#F8FAFC]">
                <div className="px-3 py-2 bg-[#CCFBF1] border-b border-[#99F6E4] text-[11px] font-bold uppercase text-[#0F766E]">
                  {col.title} ({col.list.length})
                </div>
                <form
                  className="p-2 flex gap-1.5 border-b border-[#D5DEE8]"
                  onSubmit={(e) => {
                    e.preventDefault();
                    handleAddListItem(col.kind, col.value);
                    col.setValue('');
                  }}
                >
                  <input
                    type="text"
                    value={col.value}
                    placeholder={col.placeholder}
                    onChange={(e) => col.setValue(e.target.value)}
                    className="flex-1 bg-[#FFFFFF] border border-[#C5D0DC] rounded-lg px-2 py-1 text-xs font-mono focus:outline-none focus:border-[#0F766E]"
                  />
                  <button
                    type="submit"
                    className="px-2.5 py-1 bg-[#0F766E] hover:bg-[#115E59] text-white text-xs font-semibold rounded-lg"
                  >
                    <Plus className="w-3.5 h-3.5" />
                  </button>
                </form>
                <div className="max-h-[280px] overflow-y-auto divide-y divide-[#E2E8F0]">
                  {col.list.map((item, idx) => (
                    <div
                      key={`${col.kind}-${item.label}-${idx}`}
                      className="flex items-center justify-between px-3 py-1.5 text-xs font-mono hover:bg-[#FFFFFF]"
                    >
                      <span className="font-semibold text-[#1A2332]">{item.label}</span>
                      <button
                        type="button"
                        onClick={() => handleDeleteListItem(col.kind, idx)}
                        className="p-1 text-[#5A6A7A] hover:text-[#DC2626] rounded"
                        title="Delete"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                  {col.list.length === 0 && (
                    <div className="p-4 text-center text-[11px] text-[#6B7C8C]">Empty list</div>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
