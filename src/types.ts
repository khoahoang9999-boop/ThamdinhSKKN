/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface TeacherInfo {
  teacherName: string;
  birthYear: string;
  role: string;
  schoolName: string;
  stage: 'Mầm non' | 'Tiểu học' | 'THCS' | 'Khác';
  subject: string;
}

export interface InitiativeInput {
  teacher: TeacherInfo;
  initiativeTitle: string;
  appliedDate: string;
  initiativeText: string;
}

export interface CriterionResult {
  score: number;
  levelName: string; // e.g. Rất tốt, Tốt, Khá, v.v.
  pros?: string[];    // Positive points (legacy)
  cons?: string[];    // Weaknesses / critiques (legacy)
  analysis: string[]; // Detailed analysis
  comparison?: string[]; // Added: So sánh với giải pháp thông thường
  hinhThuc?: string[];   // Added: Tính hình thức
}

export interface EvaluationResult {
  id: string;
  teacher: TeacherInfo;
  initiativeTitle: string;
  appliedDate: string;
  initiativeText: string;
  evaluatedAt: string;
  
  // V2 format
  tinhCapThiet?: {
    analysis: string[];
    levelName: string;
  };
  tinhMoi?: CriterionResult;
  tinhKhoaHoc?: CriterionResult;
  minhChung?: CriterionResult;
  uuDiem?: string[];
  hanChe?: string[];

  // Common/Legacy
  hieuQua: CriterionResult;
  phamVi: CriterionResult;
  
  improvements?: string[]; // (legacy)
  summary: string;       // Dynamic summary / IV. Đánh giá chung
  
  totalScore: number;
  classification: string;
  pronoun?: 'thay_co' | 'tac_gia';
  evaluationMode?: 'full' | 'comment_only';
  isCouncilAppraisal?: boolean;
  plagiarismResult?: PlagiarismResult | null;
}

export interface DetailedSource {
  document_title: string;
  author: string;
  exact_url: string;
  matched_snippet: string;
}

export interface PlagiarismSource {
  id: string;
  name: string;
  percent: number;
  wordsCount: number;
  url: string;
  color: 'red' | 'yellow' | 'blue' | 'purple' | 'orange';
  is_matched?: boolean;
  match_percent?: number;
  detailed_source?: DetailedSource;
}

export interface PlagiarismSegment {
  text: string;
  isDuplicate: boolean;
  sourceId: string | null;
  type: 'red' | 'yellow' | 'none';
}

export interface SpellingError {
  errorText: string;
  correction: string;
  context: string;
  reason: string;
}

export interface PlagiarismResult {
  id: string;
  title: string;
  totalDuplicatePercent: number;
  aiGeneratedPercent?: number;
  warningLevel: 'An toàn (Thấp)' | 'Vi phạm nhẹ' | 'Vi phạm trung bình' | 'Vi phạm nghiêm trọng' | string;
  extractedText: string;
  sources: PlagiarismSource[];
  segments: PlagiarismSegment[];
  aiSegments?: string[];
  spellingErrors?: SpellingError[];
  checkedAt: string;
}

