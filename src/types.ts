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
  phone?: string;
  email?: string;
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
  analysis: string[]; // Detailed analysis / Nhận xét chi tiết
  pros?: string[];    // Đánh giá / Điểm mới thể hiện ở (gạch đầu dòng)
  cons?: string[];    // Hạn chế / Tồn tại
}

export interface CouncilMember {
  id: string;
  name: string;
  unit: string;
  role: string;
}

export interface CouncilEvaluationResult {
  suCanThiet: CriterionResult;
  tinhMoi: CriterionResult;
  giaiPhap: CriterionResult;
  hieuQua: CriterionResult;
  khaNangApDung: CriterionResult;
  summary: string;
  totalScore: number;
  classification: string;
  evaluatedAt: string;
}

export interface EvaluationResult {
  id: string;
  teacher: TeacherInfo;
  initiativeTitle: string;
  appliedDate: string;
  initiativeText: string;
  evaluatedAt: string;
  
  // 5 New criteria structures
  suCanThiet: CriterionResult;    // 1. Sự cần thiết (tối đa 10đ)
  tinhMoi: CriterionResult;       // 2. Tính mới, tính sáng tạo (tối đa 20đ)
  giaiPhap: CriterionResult;      // 3. Nội dung và giải pháp (tối đa 30đ)
  hieuQua: CriterionResult;       // 4. Hiệu quả áp dụng (tối đa 30đ)
  khaNangApDung: CriterionResult;  // 5. Khả năng áp dụng, phạm vi ảnh hưởng (tối đa 10đ)

  summary: string;       // Kết luận chung / III. KẾT LUẬN
  totalScore: number;
  classification: string;
  pronoun?: 'thay_co' | 'tac_gia';
  evaluationMode?: 'full' | 'comment_only';
  isCouncilAppraisal?: boolean;
  plagiarismResult?: PlagiarismResult | null;
  uuDiem?: string[];
  hanChe?: string[];
  improvements?: string[];
  councilName?: string;
  member1Name?: string;
  member1Unit?: string;
  member1Role?: string;
  member2Name?: string;
  member2Unit?: string;
  member2Role?: string;
  member3Name?: string;
  member3Unit?: string;
  member3Role?: string;
  councilMembers?: CouncilMember[];
  councilResult?: CouncilEvaluationResult;
}

export interface DetailedSource {
  document_title: string;
  author: string;
  exact_url: string;
  alternative_urls?: string[];
  matched_snippet: string;
  search_keywords?: string;
  website_name?: string;
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

