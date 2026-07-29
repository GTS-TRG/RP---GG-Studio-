import { ReviewIssue } from '../types';

/**
 * Cột dữ liệu mà một lỗi thẩm tra thuộc về.
 * 'row' = lỗi mức dòng, không gắn với ô cụ thể (lỗi công thức, lệch liên kết sheet).
 */
export type IssueField =
  | 'cbType'
  | 'poleVal'
  | 'cbText'
  | 'cbIsc'
  | 'phaseCableText'
  | 'peCableText'
  | 'installMethod'
  | 'row';

/**
 * Quy lỗi về đúng ô thông số để hiển thị icon ngay tại đó.
 * Rule 13 (MSB) sinh 2 loại thông báo trên cùng ruleCode nên phải tách theo nội dung.
 */
export function fieldForIssue(issue: ReviewIssue): IssueField {
  switch (issue.ruleCode) {
    case 'RULE_MSB_ISC_AMP':
      // "MSB MCB Isc (...) < 65kA" -> cột Isc; "MSB MCB In (...) < 32A" -> cột In
      return /\bIsc\b/i.test(issue.description) ? 'cbIsc' : 'cbText';
    case 'RULE_ISC_TYPE':
      return 'cbIsc';
    case 'RULE_SPARE':
      return 'cbType';
    case 'RULE_POLE_PHASE':
      return 'poleVal';
    case 'RULE_RATING_LOAD':
    case 'RULE_MISSING_CB':
      return 'cbText';
    case 'RULE_PHASE_CABLE':
    case 'RULE_MULTICORE_LAYERS':
      return 'phaseCableText';
    case 'RULE_PE_CABLE':
      return 'peCableText';
    case 'RULE_CONDUIT_SIZE':
      return 'installMethod';
    case 'RULE_EXCEL_REF':
    case 'RULE_LINK_MISMATCH':
    default:
      return 'row';
  }
}

/** Gom lỗi của một dòng theo từng ô */
export function groupIssuesByField(issues: ReviewIssue[]): Map<IssueField, ReviewIssue[]> {
  const map = new Map<IssueField, ReviewIssue[]>();
  for (const iss of issues) {
    const field = fieldForIssue(iss);
    const list = map.get(field) || [];
    list.push(iss);
    map.set(field, list);
  }
  return map;
}
