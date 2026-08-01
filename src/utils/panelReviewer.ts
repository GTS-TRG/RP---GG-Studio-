import { CableOuterDiaRow, CableSpecRow, ConduitSpec, PanelSheetData, ProjectConfig, RawCircuitRow, ReviewIssue } from '../types';
import { evaluateConduitRule, parseTextToDict } from './conduitChecker';
import {
  getPrefix,
  getSheetNameFromFormula,
  getShortDesc,
  isBalancedThreePhaseLoad,
  isFireCircuit,
  isLowSpeedFan,
  isThreePhaseByLoads,
  isValidSinglePhasePole,
  isValidThreePhasePole,
} from './helpers';
import { getRequiredCBFromSpec, getSpecIndex } from './specTable';
import { DEFAULT_CABLE_OUTER_DIAS } from './outerDiaTable';
import { DEFAULT_CONDUITS } from './conduitTable';

export interface CableCoreStats {
  totalPhaseCores: number;
  singleCoreCableCount: number;
  multiCoreCableCount: number;
  maxSingleCableCoreCount: number;
}

/**
 * Parses phase cable specification text to determine the number of cores and cable count
 */
export function getCableCoreStats(phaseCableText: string): CableCoreStats {
  const dict = parseTextToDict(phaseCableText);
  let totalPhaseCores = 0;
  let singleCoreCableCount = 0;
  let multiCoreCableCount = 0;
  let maxSingleCableCoreCount = 0;

  for (const entry of dict.values()) {
    const totalCoresForEntry = entry.count * entry.coreCount;
    totalPhaseCores += totalCoresForEntry;
    if (entry.coreCount === 1) {
      singleCoreCableCount += entry.count;
    } else {
      multiCoreCableCount += entry.count;
    }
    if (entry.coreCount > maxSingleCableCoreCount) {
      maxSingleCableCoreCount = entry.coreCount;
    }
  }

  // Fallback regex pattern matching if parseTextToDict returned nothing but string exists
  if (totalPhaseCores === 0 && phaseCableText && phaseCableText.trim()) {
    const textUpper = phaseCableText.toUpperCase();
    const match1C = textUpper.match(/(\d+)\s*X?\s*1\s*C/);
    if (match1C) {
      const cnt = parseInt(match1C[1], 10);
      if (cnt > 0) {
        singleCoreCableCount = cnt;
        totalPhaseCores = cnt;
        maxSingleCableCoreCount = 1;
      }
    } else {
      const matchKC = textUpper.match(/(\d+)\s*C/);
      if (matchKC) {
        const cores = parseInt(matchKC[1], 10);
        if (cores > 0) {
          multiCoreCableCount = 1;
          totalPhaseCores = cores;
          maxSingleCableCoreCount = cores;
        }
      }
    }
  }

  return {
    totalPhaseCores,
    singleCoreCableCount,
    multiCoreCableCount,
    maxSingleCableCoreCount,
  };
}

/**
 * Kiểm tra số lượng dây/lõi cáp pha cho mạch 3 pha (không áp dụng cho Sao/Tam giác).
 * - Tải cân bằng (isBalanced=true, vd động cơ 3 pha thuần): chỉ cần 3 dây (L1,L2,L3).
 * - Tải không cân bằng (isBalanced=false, vd tủ tổng gộp nhiều tải 1 pha rải trên 3 pha,
 *   lộ vào tủ phân phối...): dòng trung tính khác 0, bắt buộc có dây N -> cần tối thiểu 4 dây.
 * Trả về thông báo lỗi (string) nếu sai, null nếu đạt.
 */
export function checkThreePhaseCoreCount(
  phaseCableText: string,
  stats: CableCoreStats,
  isBalanced: boolean
): string | null {
  const requiredCores = isBalanced ? 3 : 4;
  const neutralNote = isBalanced
    ? ''
    : ' (tải giữa các pha R/Y/B không cân bằng → dòng trung tính khác 0, bắt buộc phải có dây trung tính N)';

  if (stats.singleCoreCableCount > 0 && stats.multiCoreCableCount === 0 && stats.singleCoreCableCount < requiredCores) {
    return `Sai số lượng cáp 1 lõi cho mạch 3 pha${neutralNote}: cần tối thiểu ${requiredCores} sợi cáp 1 lõi (${requiredCores}x1C) hoặc cáp đa lõi (${requiredCores}C trở lên). Hiện tại chỉ có ${stats.singleCoreCableCount} sợi 1C (${phaseCableText}).`;
  }
  if (stats.multiCoreCableCount > 0 && stats.singleCoreCableCount === 0 && stats.maxSingleCableCoreCount < requiredCores) {
    return `Sai loại/số lõi cáp cho mạch 3 pha${neutralNote}: cần cáp đa lõi tối thiểu ${requiredCores}C. Cáp ${stats.maxSingleCableCoreCount}C (${phaseCableText}) không đủ số lõi.`;
  }
  if (stats.totalPhaseCores > 0 && stats.totalPhaseCores < requiredCores) {
    return `Thiếu số lượng dây/lõi cáp pha cho mạch 3 pha${neutralNote}: cần tối thiểu ${requiredCores} dây, hiện tại chỉ ghi nhận ${stats.totalPhaseCores} lõi/sợi cáp (${phaseCableText}).`;
  }
  return null;
}

export const DEFAULT_CONFIG: ProjectConfig = {
  minIscMsb: 65,          // kA
  minAmpMsb: 32,          // A
  maxMcbIsc: 15,          // kA
  minMccbIsc: 18,         // kA
  startRow: 13,
  safetyFactorNormal: 1.25,
  safetyFactorFire: 1.50,
  conduitFillPercent: 35,   // %
};

/**
 * Reviews a list of panel schedule sheets against technical electrical standards
 */
export function reviewWorkbookPanels(
  panels: PanelSheetData[],
  specs: CableSpecRow[],
  config: ProjectConfig = DEFAULT_CONFIG,
  outerDias: CableOuterDiaRow[] = DEFAULT_CABLE_OUTER_DIAS,
  conduits: ConduitSpec[] = DEFAULT_CONDUITS
): ReviewIssue[] {
  const issues: ReviewIssue[] = [];
  let issueCounter = 1;

  for (const panel of panels) {
    for (const circuit of panel.circuits) {
      const rowIssues = reviewOneCircuit(
        panel.sheetName,
        panel.isMSB,
        circuit,
        specs,
        config,
        outerDias,
        conduits
      );

      for (const item of rowIssues) {
        const shortDesc = getShortDesc(circuit.description);
        const lineNameWithDesc = shortDesc ? `${circuit.lineName} (${shortDesc})` : circuit.lineName;

        issues.push({
          id: `ISSUE-${issueCounter++}`,
          sheetName: panel.sheetName,
          lineName: circuit.lineName,
          shortDesc,
          lineNameWithDesc,
          description: item.text,
          isWarning: item.isWarning,
          ruleCode: item.ruleCode,
          status: 'UNRESOLVED',
          remarks: '',
          rowIndex: circuit.rowIndex,
          isResolved: false,
        });
      }
    }

    const incomingIssues = reviewIncomingCable(panel);
    for (const item of incomingIssues) {
      issues.push({
        id: `ISSUE-${issueCounter++}`,
        sheetName: panel.sheetName,
        lineName: 'INCOMING',
        shortDesc: panel.footer?.incoming?.source || 'Lộ vào / Incoming',
        lineNameWithDesc: `INCOMING (${panel.footer?.incoming?.source || 'Lộ vào / Incoming'})`,
        description: item.text,
        isWarning: item.isWarning,
        ruleCode: item.ruleCode,
        status: 'UNRESOLVED',
        remarks: '',
        rowIndex: panel.footer?.incoming?.rowIndex ?? panel.endRow,
        isResolved: false,
      });
    }
  }

  return issues;
}

interface InternalIssue {
  text: string;
  isWarning: boolean;
  ruleCode: ReviewIssue['ruleCode'];
}

/**
 * Thẩm tra cáp lộ vào (Incoming) — CB tổng + cáp nguồn cấp cho tủ.
 * Trước đây khối này chỉ đọc để hiển thị, không tham gia thẩm tra nên các lỗi như
 * "tủ 3 pha nhưng lộ vào chỉ có 2 sợi cáp pha" (thiếu dây trung tính N) không bị phát hiện.
 */
function reviewIncomingCable(panel: PanelSheetData): InternalIssue[] {
  const issues: InternalIssue[] = [];
  const inc = panel.footer?.incoming;
  if (!inc || !inc.phaseCableText) return issues;

  const totals = panel.footer?.calcPower ?? panel.footer?.ratedPower;
  const rVal = totals?.r ?? 0;
  const yVal = totals?.y ?? 0;
  const bVal = totals?.b ?? 0;

  const poleIsThreePhase = isValidThreePhasePole(inc.poleVal);
  const poleIsSinglePhase = isValidSinglePhasePole(inc.poleVal);
  // Ưu tiên số cực CB (3P/4P) vì luôn có sẵn; chỉ suy từ tải R/Y/B khi cột số cực trống/không hợp lệ.
  const isThreePhase = poleIsThreePhase || (!poleIsSinglePhase && isThreePhaseByLoads(rVal, yVal, bVal));

  const stats = getCableCoreStats(inc.phaseCableText);

  if (isThreePhase) {
    const isBalanced = isBalancedThreePhaseLoad(rVal, yVal, bVal);
    const coreCountMsg = checkThreePhaseCoreCount(inc.phaseCableText, stats, isBalanced);
    if (coreCountMsg) {
      issues.push({
        text: `[Lộ vào/Incoming] ${coreCountMsg}`,
        isWarning: false,
        ruleCode: 'RULE_PHASE_CABLE',
      });
    }
  } else if (poleIsSinglePhase) {
    if (stats.singleCoreCableCount > 0 && stats.multiCoreCableCount === 0 && stats.singleCoreCableCount < 2) {
      issues.push({
        text: `[Lộ vào/Incoming] Sai số lượng cáp cho lộ vào 1 pha: cần tối thiểu 2 sợi cáp 1 lõi (2x1C) cho dây pha L và dây trung tính N. Hiện tại chỉ có ${stats.singleCoreCableCount} sợi 1C (${inc.phaseCableText}).`,
        isWarning: false,
        ruleCode: 'RULE_PHASE_CABLE',
      });
    } else if (stats.multiCoreCableCount > 0 && stats.singleCoreCableCount === 0 && stats.maxSingleCableCoreCount < 2) {
      issues.push({
        text: `[Lộ vào/Incoming] Sai số lõi cáp cho lộ vào 1 pha: cần tối thiểu cáp 2C cho L và N. Cáp hiện tại (${inc.phaseCableText}) chỉ có ${stats.maxSingleCableCoreCount} lõi.`,
        isWarning: false,
        ruleCode: 'RULE_PHASE_CABLE',
      });
    }
  }

  return issues;
}

/**
 * Helper to check if cable cross-section meets or exceeds requirement
 * without enforcing specific cable type / material strings (Cu/PVC, Cu/XLPE/PVC, etc.)
 */
export function isCableSectionOK(cableText: string, reqText: string): boolean {
  if (!cableText || !reqText) return false;

  const cleanCable = cableText.trim().toUpperCase();
  const cleanReq = reqText.trim().toUpperCase();

  // 1. Direct string match
  if (cleanCable.includes(cleanReq)) return true;

  // Function to remove common conductor/insulation brand & material words
  const stripMaterials = (s: string) => {
    return s
      .replace(/\b(CU|AL|XLPE|PVC|CXV|CV|CXL|LSZH|LSO|FR|RUBBER|MM2|SQMM|CABLE|CAP|DAY|DÂY)\b/g, '')
      .replace(/[\/\s]/g, ' ')
      .trim();
  };

  const cMatClean = stripMaterials(cleanCable);
  const rMatClean = stripMaterials(cleanReq);

  if (cMatClean.includes(rMatClean) || rMatClean.includes(cMatClean)) return true;

  // 2. Extract numeric cross-section values (e.g. 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240, 300)
  const getSectionValues = (str: string): number[] => {
    const s = str.replace(/,/g, '.');
    const matches = s.match(/(?:[xX*()\s+]|^)(\d+(?:\.\d+)?)/g);
    if (!matches) return [];

    const nums: number[] = [];
    for (const raw of matches) {
      const cleanNum = raw.replace(/[^0-9.]/g, '');
      const val = parseFloat(cleanNum);
      if (!isNaN(val) && val > 0) {
        nums.push(val);
      }
    }
    return nums;
  };

  const reqNums = getSectionValues(rMatClean);
  const cableNums = getSectionValues(cMatClean);

  if (reqNums.length === 0) return true; // No numerical requirement found
  if (cableNums.length === 0) return false;

  const reqSection = Math.max(...reqNums);
  const cableSection = Math.max(...cableNums);

  return cableSection >= reqSection;
}

/**
 * Reviews a single circuit row according to the VBA logic
 */
function reviewOneCircuit(
  sheetName: string,
  isMSB: boolean,
  row: RawCircuitRow,
  specs: CableSpecRow[],
  config: ProjectConfig,
  outerDias: CableOuterDiaRow[],
  conduits: ConduitSpec[]
): InternalIssue[] {
  const issues: InternalIssue[] = [];

  // Skip bypassed rows
  if (row.isBypassed) {
    return issues;
  }

  const prefix = getPrefix(row.lineName);

  // Skip CP prefix circuits completely as in VBA
  if (prefix === 'CP') {
    return issues;
  }

  const descLower = row.description.toLowerCase();
  const lineLower = row.lineName.toLowerCase();

  const isSpare =
    prefix === 'SP' ||
    lineLower.startsWith('sp') ||
    descLower.includes('spare') ||
    descLower.includes('du phong') ||
    descLower.includes('dự phòng');

  const isControlCircuit =
    prefix === 'CT' ||
    lineLower.startsWith('ct') ||
    descLower.includes('dieu khien') ||
    descLower.includes('điều khiển') ||
    descLower.includes('mạch đk') ||
    descLower.includes('mach dk') ||
    descLower.includes('control');

  const isStarDelta = descLower.includes('(s/d)');

  // =======================================================
  // 0. EXCEL FORMULA ERROR (#REF! = SheetJS code 23, etc.)
  // =======================================================
  if (row.hasExcelError && row.excelErrorFields && row.excelErrorFields.length > 0) {
    issues.push({
      text: 'Excel formula error',
      isWarning: false,
      ruleCode: 'RULE_EXCEL_REF',
    });
  }

  // =======================================================
  // 1. CHECK ISC AND IN RATING
  // - Isc ≥ minIscMsb (mặc định 65kA) + In ≥ minAmpMsb: CHỈ tủ MSB
  // - Tủ khác MSB: KHÔNG yêu cầu ≥ 65kA; chỉ áp ngưỡng MCB/MCCB phân phối
  // =======================================================
  const iscAmp = row.iscAmp;
  const cbAmp = row.cbAmp;
  const cbType = row.cbType;

  if (isMSB) {
    // RULE 13 — chỉ Main Switchboard (MSB)
    if (cbType === 'MCB' || cbType === 'MCCB') {
      if (iscAmp > 0 && iscAmp < config.minIscMsb) {
        issues.push({
          text: `MSB ${cbType} Isc (${row.cbIsc}) < ${config.minIscMsb}kA (minimum required for MSB: ${config.minIscMsb}kA).`,
          isWarning: isSpare,
          ruleCode: 'RULE_MSB_ISC_AMP',
        });
      }
      if (cbAmp > 0 && cbAmp < config.minAmpMsb) {
        issues.push({
          text: `MSB ${cbType} In (${cbAmp}A) < ${config.minAmpMsb}A (minimum required for MSB: ${config.minAmpMsb}A).`,
          isWarning: isSpare,
          ruleCode: 'RULE_MSB_ISC_AMP',
        });
      }
    }
  } else if (iscAmp > 0) {
    // Tủ phân phối / chiếu sáng — không áp dụng tiêu chí 65kA
    if (cbType === 'MCB' && iscAmp > config.maxMcbIsc) {
      issues.push({
        text: `MCB Isc (${row.cbIsc}) > ${config.maxMcbIsc}kA (max allowed for distribution MCB: ${config.maxMcbIsc}kA).`,
        isWarning: isSpare,
        ruleCode: 'RULE_ISC_TYPE',
      });
    } else if (cbType === 'MCCB' && iscAmp < config.minMccbIsc) {
      issues.push({
        text: `MCCB Isc (${row.cbIsc}) < ${config.minMccbIsc}kA (minimum required for distribution MCCB: ${config.minMccbIsc}kA).`,
        isWarning: isSpare,
        ruleCode: 'RULE_ISC_TYPE',
      });
    }
  }

  // =======================================================
  // 2. SPARE CIRCUIT LOGIC
  // =======================================================
  if (isSpare) {
    if (!cbType || !row.poleVal || cbAmp <= 0 || !row.cbIsc) {
      issues.push({
        text: 'Thiếu thông số CB (Loại CB, Số cực Pole, Dòng định mức In hoặc Dòng cắt Isc) cho mạch dự phòng (Spare).',
        isWarning: true,
        ruleCode: 'RULE_SPARE',
      });
    }
    return issues; // Spare checks stop here
  }

  // =======================================================
  // 3. CHECK PHASE AND CB POLES
  // =======================================================
  const isThreePhase = isThreePhaseByLoads(row.rLoad, row.yLoad, row.bLoad);
  if (isThreePhase) {
    if (!isValidThreePhasePole(row.poleVal)) {
      issues.push({
        text: 'Mạch 3 pha nhưng chọn CB sai số cực (bắt buộc phải là 3P hoặc 4P).',
        isWarning: false,
        ruleCode: 'RULE_POLE_PHASE',
      });
    }
  } else {
    if (!isValidSinglePhasePole(row.poleVal)) {
      issues.push({
        text: 'Mạch 1 pha -> Số cực: 1P, 1P+N, hoặc 2P',
        isWarning: false,
        ruleCode: 'RULE_POLE_PHASE',
      });
    }
    if (prefix === 'S' && row.poleVal !== '1P+N') {
      issues.push({
        text: 'Mạch ổ cắm (tiền tố S) bắt buộc phải dùng CB 1P+N bảo vệ dây nguội.',
        isWarning: false,
        ruleCode: 'RULE_POLE_PHASE',
      });
    }
  }

  // =======================================================
  // 4. CHECK CB RATING VS CALCULATED LOAD
  // =======================================================
  if (row.iCalc > 0) {
    const isFire = isFireCircuit(row.description, row.lineName);
    const safetyFactor = isFire ? config.safetyFactorFire : config.safetyFactorNormal;
    const calculatedLoad = row.iCalc * safetyFactor;
    const reqCB = getRequiredCBFromSpec(calculatedLoad, specs);

    if (cbAmp <= 0) {
      issues.push({
        text: 'Thiếu hoặc sai định dạng dòng định mức CB (In).',
        isWarning: false,
        ruleCode: 'RULE_MISSING_CB',
      });
    } else if (reqCB > 0 && cbAmp < reqCB) {
      if (calculatedLoad - cbAmp > 2) {
        issues.push({
          text: `CB được chọn (${cbAmp}A) nhỏ hơn yêu cầu tính toán (${reqCB}A). Tải tính toán an toàn: ${calculatedLoad.toFixed(1)}A.`,
          isWarning: false,
          ruleCode: 'RULE_RATING_LOAD',
        });
      }
    } else if (reqCB > 0 && cbAmp > reqCB) {
      const reqIdx = getSpecIndex(reqCB, specs);
      const selIdx = getSpecIndex(cbAmp, specs);
      if (selIdx > reqIdx + 1) {
        issues.push({
          text: `Chọn CB quá cỡ (Đã chọn: ${cbAmp}A, Tính toán yêu cầu: ${reqCB}A).`,
          isWarning: true,
          ruleCode: 'RULE_RATING_LOAD',
        });
      }
    }
  }

  // =======================================================
  // 5. CHECK CABLE SPECIFICATIONS (Skip for CT control circuits and Spare)
  // =======================================================
  if (cbAmp > 0 && !isControlCircuit) {
    const specIdx = getSpecIndex(cbAmp, specs);
    let reqPhaseText = '';
    let reqPEText = '';
    let reqPhaseText_Reduced = '';
    let reqPhaseText_Reduced2 = '';
    let reqPEText_Reduced = '';
    let reqPEText_Reduced2 = '';

    if (specIdx >= 0 && specs[specIdx]) {
      reqPhaseText = specs[specIdx].phaseText;
      reqPEText = specs[specIdx].peText;

      if (isStarDelta) {
        if (specIdx > 0 && specs[specIdx - 1]) {
          reqPhaseText_Reduced = specs[specIdx - 1].phaseText;
          reqPEText_Reduced = specs[specIdx - 1].peText;
        }
        if (specIdx > 1 && specs[specIdx - 2]) {
          reqPhaseText_Reduced2 = specs[specIdx - 2].phaseText;
          reqPEText_Reduced2 = specs[specIdx - 2].peText;
        }
      }
    }

    // --- Phase Cable Check ---
    if (!row.phaseCableText) {
      issues.push({
        text: 'Thiếu tiết diện dây pha (Phase cable).',
        isWarning: false,
        ruleCode: 'RULE_PHASE_CABLE',
      });
    } else {
      // 1. Check phase cable cross-section text matching
      if (reqPhaseText) {
        let phaseOK = isCableSectionOK(row.phaseCableText, reqPhaseText);
        if (!phaseOK && isStarDelta) {
          if (reqPhaseText_Reduced && isCableSectionOK(row.phaseCableText, reqPhaseText_Reduced)) {
            phaseOK = true;
          }
          if (!phaseOK && reqPhaseText_Reduced2 && isCableSectionOK(row.phaseCableText, reqPhaseText_Reduced2)) {
            phaseOK = true;
          }
        }

        if (!phaseOK) {
          const detail = isStarDelta && reqPhaseText_Reduced
            ? `Sai tiết diện dây pha (Yêu cầu tiết diện: ${reqPhaseText}, hoặc ${reqPhaseText_Reduced} cho khởi động Sao/Tam giác)`
            : `Sai tiết diện dây pha (Yêu cầu tiết diện: ${reqPhaseText})`;
          issues.push({
            text: detail,
            isWarning: isSpare,
            ruleCode: 'RULE_PHASE_CABLE',
          });
        }
      }

      // 2. Check phase cable core quantity / conductor count matching phase requirements
      const stats = getCableCoreStats(row.phaseCableText);

      if (isThreePhase) {
        if (isStarDelta) {
          if (stats.totalPhaseCores > 0 && stats.totalPhaseCores < 6) {
            issues.push({
              text: `Sai số lượng dây pha cho động cơ Sao/Tam giác (S/D): Cần ít nhất 6 dây pha (2 bộ 3x1C hoặc 2x3C), hiện tại chỉ ghi nhận ${stats.totalPhaseCores} lõi/sợi cáp (${row.phaseCableText}).`,
              isWarning: isSpare,
              ruleCode: 'RULE_PHASE_CABLE',
            });
          }
        } else {
          // Mạch 3 pha thường: tải cân bằng chỉ cần 3 dây (L1,L2,L3); tải không cân bằng
          // (vd gộp nhiều tải 1 pha rải trên 3 pha) bắt buộc có dây N -> cần tối thiểu 4 dây.
          const isBalanced = isBalancedThreePhaseLoad(row.rLoad, row.yLoad, row.bLoad);
          const coreCountMsg = checkThreePhaseCoreCount(row.phaseCableText, stats, isBalanced);
          if (coreCountMsg) {
            issues.push({
              text: coreCountMsg,
              isWarning: isSpare,
              ruleCode: 'RULE_PHASE_CABLE',
            });
          }
        }
      } else if (isValidSinglePhasePole(row.poleVal)) {
        // 1-phase circuit requires at least 2 conductors (L + N)
        if (stats.singleCoreCableCount > 0 && stats.multiCoreCableCount === 0 && stats.singleCoreCableCount < 2) {
          issues.push({
            text: `Sai số lượng cáp cho mạch 1 pha: Mạch 1 pha cần tối thiểu 2 sợi cáp 1 lõi (2x1C) cho dây pha L và dây trung tính N. Hiện tại chỉ có ${stats.singleCoreCableCount} sợi 1C (${row.phaseCableText}).`,
            isWarning: isSpare,
            ruleCode: 'RULE_PHASE_CABLE',
          });
        } else if (stats.multiCoreCableCount > 0 && stats.singleCoreCableCount === 0 && stats.maxSingleCableCoreCount < 2) {
          issues.push({
            text: `Sai số lõi cáp cho mạch 1 pha: Cần tối thiểu cáp 2C (hoặc 2x1C) cho L và N. Cáp hiện tại (${row.phaseCableText}) chỉ có ${stats.maxSingleCableCoreCount} lõi.`,
            isWarning: isSpare,
            ruleCode: 'RULE_PHASE_CABLE',
          });
        }
      }
    }

    // --- PE Cable Check ---
    if (!isLowSpeedFan(row.description)) {
      if (!row.peCableText) {
        let hasIntegratedPE = false;
        if (!isThreePhase && row.phaseCableText.toUpperCase().includes('3C')) hasIntegratedPE = true;
        if (isThreePhase && row.phaseCableText.toUpperCase().includes('5C')) hasIntegratedPE = true;

        if (!hasIntegratedPE) {
          issues.push({
            text: 'Thiếu tiết diện dây nối đất PE (PE cable).',
            isWarning: false,
            ruleCode: 'RULE_PE_CABLE',
          });
        }
      } else if (reqPEText) {
        let peOK = isCableSectionOK(row.peCableText, reqPEText);
        if (!peOK && isStarDelta) {
          if (reqPEText_Reduced && isCableSectionOK(row.peCableText, reqPEText_Reduced)) {
            peOK = true;
          }
          if (!peOK && reqPEText_Reduced2 && isCableSectionOK(row.peCableText, reqPEText_Reduced2)) {
            peOK = true;
          }
        }

        if (!peOK) {
          const detail = isStarDelta && reqPEText_Reduced
            ? `Sai tiết diện dây PE (Yêu cầu tiết diện: ${reqPEText}, hoặc ${reqPEText_Reduced} cho Sao/Tam giác)`
            : `Sai tiết diện dây PE (Yêu cầu tiết diện: ${reqPEText})`;
          issues.push({
            text: detail,
            isWarning: isSpare,
            ruleCode: 'RULE_PE_CABLE',
          });
        }
      }
    }
  }

  // =======================================================
  // 6. CHECK CONDUIT FILL RATIO AND SIZE (VBA KiemTraOng3 / KIEMTRAONGMOD)
  // =======================================================
  if (!isControlCircuit && (row.phaseCableText || row.peCableText)) {
    const conduitResult = evaluateConduitRule(
      row.phaseCableText,
      row.peCableText || '',
      row.installMethod,
      config.conduitFillPercent > 0 ? config.conduitFillPercent : 35,
      outerDias,
      conduits
    );

    if (conduitResult?.message) {
      // Overfill / undersized = error; thiếu ống / lỗi tra bảng = warning
      const soft =
        conduitResult.hasError ||
        (!conduitResult.isConduitUsed && !!conduitResult.recommendation);
      issues.push({
        text: conduitResult.message,
        isWarning: isSpare || soft,
        ruleCode: 'RULE_CONDUIT_SIZE',
      });
    }
  }

  // =======================================================
  // 7. CROSS-LINK FORMULA MISMATCH CHECK
  // =======================================================
  const sourceSheetLoad = getSheetNameFromFormula(row.formulaLoad);
  const sourceSheetCB = getSheetNameFromFormula(row.formulaCB);

  if (sourceSheetLoad && sourceSheetCB && sourceSheetLoad !== sourceSheetCB) {
    issues.push({
      text: `Cảnh báo liên kết công thức: Công suất lấy từ tủ [${sourceSheetLoad}] nhưng chọn CB lại lấy từ tủ [${sourceSheetCB}].`,
      isWarning: true,
      ruleCode: 'RULE_LINK_MISMATCH',
    });
  }

  return issues;
}
