/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useRef, MouseEvent, ChangeEvent, DragEvent } from 'react';
import { 
  CheckCircle2, 
  AlertCircle, 
  XCircle, 
  Info, 
  FileText, 
  Printer, 
  BookOpen, 
  Award, 
  User, 
  Calendar, 
  Building, 
  Clock, 
  Phone,
  Mail, 
  Sparkles, 
  History, 
  ChevronRight, 
  ArrowUpRight, 
  GraduationCap, 
  RefreshCw, 
  Bookmark, 
  Trash2,
  Pencil,
  FileDown,
  FileUp,
  ExternalLink,
  Zap,
  Key,
  Copy,
  Search,
  Eye,
  EyeOff,
  Plus,
  Minus,
  ArrowRight,
  Check,
  Loader2,
  Settings,
  LogOut,
  Link,
  HelpCircle,
  Download
} from 'lucide-react';
import { useReactToPrint } from 'react-to-print';
import { PlagiarismResult, PlagiarismSource, PlagiarismSegment, EvaluationResult, TeacherInfo, CouncilMember, CouncilEvaluationResult } from './types';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { asBlob } from 'html-docx-js-typescript';
import { HighlightedText } from './components/HighlightedText';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { extractTextFromPDFFile } from './utils/pdfExtract';
import LoginScreen from './components/LoginScreen';
import { auth, db } from './lib/firebase';
import { useAuthState } from 'react-firebase-hooks/auth';
import { signOut } from 'firebase/auth';
import { doc, getDoc, setDoc } from 'firebase/firestore';

// Setup PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

// Helpers
const renderList = (lines?: string[], defaultText = ".......................................................................................................................................................................................") => {
  if (!lines || lines.length === 0) return defaultText;
  return lines.map((line, idx) => {
    let text = line.trim();
    if (text.startsWith('-') || text.startsWith('•') || text.startsWith('+')) {
      text = text.substring(1).trim();
    }
    return <p key={idx} className="mb-1" style={{ textAlign: 'justify' }}>- {text}</p>;
  });
};

export default function App() {
  // Application states
  const [teacher, setTeacher] = useState<TeacherInfo>({
    teacherName: '',
    birthYear: '',
    role: '',
    schoolName: '',
    stage: 'Tiểu học',
    subject: '',
    phone: '',
    email: ''
  });
  
  const [initiativeTitle, setInitiativeTitle] = useState('');
  const [appliedDate, setAppliedDate] = useState('2026-06-11');
  const [initiativeText, setInitiativeText] = useState('');

  // Council appraisal committee member states
  const [councilName, setCouncilName] = useState('CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN');
  const [member1Name, setMember1Name] = useState('');
  const [member1Unit, setMember1Unit] = useState('');
  const [member1Role, setMember1Role] = useState('');
  const [member2Name, setMember2Name] = useState('');
  const [member2Unit, setMember2Unit] = useState('');
  const [member2Role, setMember2Role] = useState('');
  const [member3Name, setMember3Name] = useState('');
  const [member3Unit, setMember3Unit] = useState('');
  const [member3Role, setMember3Role] = useState('');
  
  const [councilMembers, setCouncilMembers] = useState<CouncilMember[]>([
    { id: '1', name: '', unit: '', role: '' },
    { id: '2', name: '', unit: '', role: '' },
    { id: '3', name: '', unit: '', role: '' },
  ]);
  
  const [apiKeys, setApiKeys] = useState<string[]>(['', '', '']);
  const [visibleKeys, setVisibleKeys] = useState<Record<number, boolean>>({});
  
  const [selectedModelAppraisal, setSelectedModelAppraisal] = useState('gemini-3.5-flash');
  const [selectedModelPlag, setSelectedModelPlag] = useState('gemini-3.1-flash-lite');
  const [selectedModelExtract, setSelectedModelExtract] = useState('gemini-3.1-flash-lite');

  const [isLoading, setIsLoading] = useState(false);
  const [loadingMessageIndex, setLoadingMessageIndex] = useState(0);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  
  const [user, loadingAuth] = useAuthState(auth);
  
  // Evaluation History
  const [history, setHistory] = useState<EvaluationResult[]>([]);
  const [currentResult, setCurrentResult] = useState<EvaluationResult | null>(null);
  
  // Re-evaluation / Re-grading panel states
  const [showReAppraisalPanel, setShowReAppraisalPanel] = useState(false);
  const [desiredScore, setDesiredScore] = useState<number | ''>('');
  const [reAppraisalNotes, setReAppraisalNotes] = useState('');
  
  // Active Tab for Legal Handbook / Rubric in LHS
  const [lhsTab, setLhsTab] = useState<'rubric' | 'legal' | 'xml_guide'>('rubric');
  
  // Print Mode Modal
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [printTarget, setPrintTarget] = useState<'department' | 'council'>('department');
  const [activeAppraisalView, setActiveAppraisalView] = useState<'department' | 'council'>('department');
  const printRef = useRef<HTMLDivElement>(null);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);
  const [isGeneratingWord, setIsGeneratingWord] = useState(false);

  const getRadioCheck = (score: number | undefined, maxScore: number) => {
    if (score === undefined) return 3;
    const ratio = score / maxScore;
    if (ratio >= 0.85) return 0;
    if (ratio >= 0.6) return 1;
    if (ratio >= 0.4) return 2;
    return 3;
  };

  const handleDownloadPDF = useReactToPrint({
    contentRef: printRef,
    documentTitle: currentResult ? `PhieuGiamDinh_${currentResult.teacher.teacherName || 'TacGia'}` : 'Phieu_Danh_Gia_SKKN',
    onBeforePrint: () => {
      setIsGeneratingPDF(true);
      return Promise.resolve();
    },
    onAfterPrint: () => {
      setIsGeneratingPDF(false);
    },
    onPrintError: () => {
      setIsGeneratingPDF(false);
    }
  });

  const handleDownloadWord = async () => {
    setIsGeneratingWord(true);
    try {
      if (!currentResult) return;
      
      const r = currentResult;
      const isCouncilTarget = printTarget === 'council' && r.councilResult;
      const srcEval = isCouncilTarget ? r.councilResult! : r;

      const evalDate = new Date(srcEval.evaluatedAt || r.evaluatedAt || new Date());
      const dateStr = `Hàm Yên, ngày ${evalDate.getDate()} tháng ${evalDate.getMonth() + 1} năm ${evalDate.getFullYear()}`;

      // Custom Appraisal Council details
      const rCouncilName = r.councilName || councilName || 'CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN';
      const rCouncilMembers: CouncilMember[] = r.councilMembers || councilMembers || [];
      const rMember1Name = r.member1Name || member1Name || '';
      const rMember1Unit = r.member1Unit || member1Unit || '';
      const rMember1Role = r.member1Role || member1Role || '';
      const rMember2Name = r.member2Name || member2Name || '';
      const rMember2Unit = r.member2Unit || member2Unit || '';
      const rMember2Role = r.member2Role || member2Role || '';
      const rMember3Name = r.member3Name || member3Name || '';
      const rMember3Unit = r.member3Unit || member3Unit || '';
      const rMember3Role = r.member3Role || member3Role || '';

      const hasCustomAppraisalMembers = rCouncilMembers.length > 0 
        ? rCouncilMembers.some(m => m.name.trim() !== '') 
        : Boolean(rMember1Name.trim() || rMember2Name.trim() || rMember3Name.trim());

      // Helper to generate list lines for Word
      const getWordListString = (lines?: string[]) => {
        if (!lines || lines.length === 0) {
          return `<p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;">.......................................................................................................................................................................................</p>`;
        }
        return lines.map(line => {
          let text = line.trim();
          if (text.startsWith('-') || text.startsWith('•') || text.startsWith('+')) {
            text = text.substring(1).trim();
          }
          return `<p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;">- ${text}</p>`;
        }).join('');
      };

      const suCanThietStr = getWordListString(srcEval.suCanThiet?.analysis);
      const tinhMoiStr = getWordListString(srcEval.tinhMoi?.analysis);
      const tinhMoiConsStr = srcEval.tinhMoi?.cons && srcEval.tinhMoi.cons.length > 0 
        ? `<p style="margin-top: 6pt; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;"><strong>Hạn chế:</strong></p>${getWordListString(srcEval.tinhMoi.cons)}` 
        : "";

      const giaiPhapStr = getWordListString(srcEval.giaiPhap?.analysis);
      const giaiPhapProsStr = srcEval.giaiPhap?.pros && srcEval.giaiPhap.pros.length > 0
        ? `<p style="margin-top: 6pt; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;"><strong>Nhận xét chung:</strong></p>${getWordListString(srcEval.giaiPhap.pros)}`
        : "";
      const giaiPhapConsStr = srcEval.giaiPhap?.cons && srcEval.giaiPhap.cons.length > 0
        ? `<p style="margin-top: 6pt; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;"><strong>Hạn chế:</strong></p>${getWordListString(srcEval.giaiPhap.cons)}`
        : "";

      const hieuQuaStr = getWordListString(srcEval.hieuQua?.analysis);
      const hieuQuaConsStr = srcEval.hieuQua?.cons && srcEval.hieuQua.cons.length > 0
        ? `<p style="margin-top: 6pt; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;"><strong>Hạn chế:</strong></p>${getWordListString(srcEval.hieuQua.cons)}`
        : "";

      const khaNangApDungStr = getWordListString(srcEval.khaNangApDung?.analysis);
      const khaNangApDungConsStr = srcEval.khaNangApDung?.cons && srcEval.khaNangApDung.cons.length > 0
        ? `<p style="margin-top: 6pt; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.45;"><strong>Hạn chế:</strong></p>${getWordListString(srcEval.khaNangApDung.cons)}`
        : "";

      const docTitle = isCouncilTarget
        ? `<div style="text-align: center; margin-bottom: 18pt; margin-top: 12pt;">
            <p class="bold uppercase" style="font-size: 15pt; text-align: center; margin: 0; font-family: 'Times New Roman', Times, serif; font-weight: bold;">PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN</p>
            <p class="bold uppercase" style="font-size: 14pt; text-align: center; margin: 4pt 0 0 0; font-family: 'Times New Roman', Times, serif; font-weight: bold;">HỘI ĐỒNG THẨM ĐỊNH SÁNG KIẾN</p>
          </div>`
        : hasCustomAppraisalMembers ? `
          <div style="text-align: center; margin-bottom: 18pt; margin-top: 12pt;">
            <p class="bold uppercase" style="font-size: 15pt; text-align: center; margin: 0; font-family: 'Times New Roman', Times, serif; font-weight: bold;">PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN</p>
            <p class="bold uppercase" style="font-size: 14pt; text-align: center; margin: 4pt 0 0 0; font-family: 'Times New Roman', Times, serif; font-weight: bold;">${rCouncilName}</p>
          </div>
          ` : `
          <div style="text-align: center; margin-bottom: 24pt; margin-top: 12pt;">
            <p class="bold uppercase" style="font-size: 15pt; text-align: center; margin: 0; font-family: 'Times New Roman', Times, serif; font-weight: bold;">PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN</p>
          </div>
          `;

      const signerTitle = isCouncilTarget
        ? "TM. HỘI ĐỒNG THẨM ĐỊNH SÁNG KIẾN"
        : "NGƯỜI NHẬN XÉT, ĐÁNH GIÁ";

      const msWordObj = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>Phiếu Nhận Xét Đánh Giá Sáng Kiến</title>
        <style>
          @page WordSection1 {
            size: 210mm 297mm;
            margin-top: 20mm;
            margin-bottom: 20mm;
            margin-left: 30mm;
            margin-right: 15mm;
          }
          div.WordSection1 { page: WordSection1; }
          body {
            font-family: "Times New Roman", Times, serif;
            font-size: 14pt;
            line-height: 1.45;
            color: #000000;
          }
          p, div {
            margin-top: 0;
            margin-bottom: 6pt;
            text-align: justify;
            font-family: "Times New Roman", Times, serif;
            font-size: 14pt;
            line-height: 1.45;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 12pt;
            font-family: "Times New Roman", Times, serif;
            font-size: 14pt;
          }
          table.border-table {
            border: 1px solid #000000;
          }
          table.border-table th, table.border-table td {
            border: 1px solid #000000;
            padding: 6pt;
          }
          .text-center { text-align: center; }
          .text-right { text-align: right; }
          .bold { font-weight: bold; }
          .italic { font-style: italic; }
          .underline { text-decoration: underline; }
          .uppercase { text-transform: uppercase; }
        </style>
      </head>
      <body>
        <div class="WordSection1">
          <!-- TOP BANNER (Quốc hiệu Tiêu ngữ) -->
          <table style="width: 100%; border: none; margin-bottom: 18pt;">
            <tr style="border: none;">
              <td style="width: 100%; text-align: center; border: none; padding: 0;">
                <p class="bold" style="font-size: 13pt; text-align: center; margin-bottom: 2pt; text-transform: uppercase; font-family: 'Times New Roman', Times, serif;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</p>
                <p class="bold" style="font-size: 14pt; text-align: center; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif;">Độc lập - Tự do - Hạnh phúc</p>
                <p style="text-align: center; margin-top: 0; margin-bottom: 12pt;">
                  <span style="border-bottom: 1.5px solid #000000; width: 150px; display: inline-block; height: 1px;"></span>
                </p>
              </td>
            </tr>
            <tr style="border: none;">
              <td style="width: 100%; text-align: right; border: none; padding: 0;">
                <p class="italic" style="font-size: 14pt; text-align: right; margin-top: 6pt; margin-bottom: 0; font-family: 'Times New Roman', Times, serif;">${dateStr}</p>
              </td>
            </tr>
          </table>

          <!-- DOCUMENT TITLE -->
          ${docTitle}

          <!-- TEACHER INFORMATION -->
          ${hasCustomAppraisalMembers ? `
          <div style="margin-bottom: 18pt; line-height: 1.45;">
            ${rCouncilMembers.length > 0 ? rCouncilMembers.map((m, idx) => `
              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>${idx + 1}. Họ và tên Thành viên ${idx + 1}:</strong> ${m.name || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Đơn vị công tác: ${m.unit || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Chức vụ: ${m.role || '...................................................'}</p>
            `).join('') + `
              <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>${rCouncilMembers.length + 1}. Tên giải pháp đề nghị công nhận sáng kiến:</strong> ${r.initiativeTitle ? `“${r.initiativeTitle}”` : '...................................................'}</p>
            ` : `
              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>1. Họ và tên Thành viên 1:</strong> ${rMember1Name || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Đơn vị công tác: ${rMember1Unit || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Chức vụ: ${rMember1Role || '...................................................'}</p>

              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>2. Họ và tên Thành viên 2:</strong> ${rMember2Name || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Đơn vị công tác: ${rMember2Unit || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Chức vụ: ${rMember2Role || '...................................................'}</p>

              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>3. Họ và tên Thành viên 3:</strong> ${rMember3Name || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Đơn vị công tác: ${rMember3Unit || '...................................................'}</p>
              <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt; padding-left: 20px;">- Chức vụ: ${rMember3Role || '...................................................'}</p>

              <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>4. Tên giải pháp đề nghị công nhận sáng kiến:</strong> ${r.initiativeTitle ? `“${r.initiativeTitle}”` : '...................................................'}</p>
            `}

            <p style="margin-top: 12pt; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">- Họ và tên Tác giả: ${r.teacher.teacherName || '...................................................'}</p>
            <p style="margin-top: 0; margin-bottom: 4pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">- Chức vụ, đơn vị công tác: ${r.teacher.role || '............................'}, ${r.teacher.schoolName || '............................................'}</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">- Điện thoại: ${r.teacher.phone || '......................................'}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Email:</strong> ${r.teacher.email || '..........................'}</p>
          </div>
          ` : `
          <div style="margin-bottom: 18pt; line-height: 1.45;">
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Họ và tên tác giả:</strong> ${r.teacher.teacherName || '...................................................'}</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Chức vụ:</strong> ${r.teacher.role || '...................................................'}</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Đơn vị công tác:</strong> ${r.teacher.schoolName || '...................................................'}</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Tên sáng kiến:</strong> ${r.initiativeTitle ? `“${r.initiativeTitle}”` : '...................................................'}</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Lĩnh vực áp dụng:</strong> ${r.teacher.subject || 'Giáo dục mầm non'}</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Điện thoại:</strong> ${r.teacher.phone || '......................................'}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Email:</strong> ${r.teacher.email || '..........................'}</p>
          </div>
          `}

          <!-- I. NHẬN XÉT, ĐÁNH GIÁ CHI TIẾT -->
          <p class="bold uppercase" style="font-size: 14pt; font-weight: bold; margin-top: 18pt; margin-bottom: 12pt; font-family: 'Times New Roman', Times, serif;">I. NHẬN XÉT, ĐÁNH GIÁ CHI TIẾT</p>

          <!-- 1. Sự cần thiết -->
          <div style="margin-bottom: 12pt;">
            <p class="bold" style="margin-top: 0; margin-bottom: 4pt; text-transform: uppercase; font-family: 'Times New Roman', Times, serif; font-weight: bold;">1. Về sự cần thiết của sáng kiến</p>
            <div style="margin-bottom: 4pt;">
              ${suCanThietStr}
            </div>
            <p style="margin-top: 4pt; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Điểm: ${srcEval.suCanThiet?.score || 0}/10 điểm</strong></p>
          </div>

          <!-- 2. Tính mới -->
          <div style="margin-bottom: 12pt;">
            <p class="bold" style="margin-top: 0; margin-bottom: 4pt; text-transform: uppercase; font-family: 'Times New Roman', Times, serif; font-weight: bold;">2. Về tính mới, tính sáng tạo của sáng kiến</p>
            <div style="margin-bottom: 4pt;">
              ${tinhMoiStr}
              ${tinhMoiConsStr}
            </div>
            <p style="margin-top: 4pt; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Điểm: ${srcEval.tinhMoi?.score || 0}/20 điểm</strong></p>
          </div>

          <!-- 3. Nội dung và giải pháp -->
          <div style="margin-bottom: 12pt;">
            <p class="bold" style="margin-top: 0; margin-bottom: 4pt; text-transform: uppercase; font-family: 'Times New Roman', Times, serif; font-weight: bold;">3. Về nội dung và chất lượng các giải pháp</p>
            <div style="margin-bottom: 4pt;">
              ${giaiPhapStr}
              ${giaiPhapProsStr}
              ${giaiPhapConsStr}
            </div>
            <p style="margin-top: 4pt; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Điểm: ${srcEval.giaiPhap?.score || 0}/30 điểm</strong></p>
          </div>

          <!-- 4. Hiệu quả áp dụng -->
          <div style="margin-bottom: 12pt;">
            <p class="bold" style="margin-top: 0; margin-bottom: 4pt; text-transform: uppercase; font-family: 'Times New Roman', Times, serif; font-weight: bold;">4. Về hiệu quả áp dụng</p>
            <div style="margin-bottom: 4pt;">
              ${hieuQuaStr}
              ${hieuQuaConsStr}
            </div>
            <p style="margin-top: 4pt; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Điểm: ${srcEval.hieuQua?.score || 0}/30 điểm</strong></p>
          </div>

          <!-- 5. Khả năng áp dụng -->
          <div style="margin-bottom: 12pt;">
            <p class="bold" style="margin-top: 0; margin-bottom: 4pt; text-transform: uppercase; font-family: 'Times New Roman', Times, serif; font-weight: bold;">5. Về khả năng áp dụng và phạm vi ảnh hưởng</p>
            <div style="margin-bottom: 4pt;">
              ${khaNangApDungStr}
              ${khaNangApDungConsStr}
            </div>
            <p style="margin-top: 4pt; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;"><strong>Điểm: ${srcEval.khaNangApDung?.score || 0}/10 điểm</strong></p>
          </div>

          <!-- II. TỔNG HỢP KẾT QUẢ CHẤM ĐIỂM -->
          <p class="bold uppercase" style="font-size: 14pt; font-weight: bold; margin-top: 18pt; margin-bottom: 12pt; font-family: 'Times New Roman', Times, serif;">II. TỔNG HỢP KẾT QUẢ CHẤM ĐIỂM</p>
          <table border="1" cellpadding="6" cellspacing="0" style="width: 100%; border-collapse: collapse; border: 1px solid #000000; margin-bottom: 18pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">
            <thead>
              <tr style="background-color: #f2f2f2;">
                <th style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; font-weight: bold; text-align: center;">Nội dung đánh giá</th>
                <th style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; font-weight: bold; text-align: center; width: 120px;">Điểm tối đa</th>
                <th style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; font-weight: bold; text-align: center; width: 120px;">Điểm chấm</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: left;">Sự cần thiết của sáng kiến</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center;">10</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center; font-weight: bold;">${srcEval.suCanThiet?.score || 0}</td>
              </tr>
              <tr>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: left;">Tính mới, tính sáng tạo</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center;">20</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center; font-weight: bold;">${srcEval.tinhMoi?.score || 0}</td>
              </tr>
              <tr>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: left;">Nội dung và giải pháp</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center;">30</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center; font-weight: bold;">${srcEval.giaiPhap?.score || 0}</td>
              </tr>
              <tr>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: left;">Hiệu quả áp dụng</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center;">30</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center; font-weight: bold;">${srcEval.hieuQua?.score || 0}</td>
              </tr>
              <tr>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: left;">Khả năng áp dụng, phạm vi ảnh hưởng</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center;">10</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; text-align: center; font-weight: bold;">${srcEval.khaNangApDung?.score || 0}</td>
              </tr>
              <tr style="font-weight: bold;">
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; font-weight: bold; text-align: center;">Tổng cộng</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; font-weight: bold; text-align: center;">100</td>
                <td style="border: 1px solid #000000; padding: 6pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.15; font-weight: bold; text-align: center; color: #900000;">${srcEval.totalScore}</td>
              </tr>
            </tbody>
          </table>

          <!-- III. KẾT LUẬN -->
          <p class="bold uppercase" style="font-size: 14pt; font-weight: bold; margin-top: 18pt; margin-bottom: 12pt; font-family: 'Times New Roman', Times, serif;">III. KẾT LUẬN</p>
          <div style="line-height: 1.45; margin-bottom: 18pt;">
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">Sáng kiến: ${r.initiativeTitle ? `“${r.initiativeTitle}”` : '...................................................'} là sáng kiến có giá trị thực tiễn, đáp ứng yêu cầu đổi mới giáo dục mầm non, góp phần nâng cao chất lượng chuẩn bị cho trẻ trước khi vào lớp 1.</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">Hệ thống giải pháp được xây dựng tương đối đồng bộ, phù hợp với điều kiện thực tế của đơn vị, có khả năng áp dụng rộng rãi tại các cơ sở giáo dục mầm non có điều kiện tương đồng.</p>
            <p style="margin-top: 0; margin-bottom: 6pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">Tuy nhiên, để nâng cao giá trị khoa học và sức thuyết phục của sáng kiến, tác giả cần bổ sung thêm số liệu định lượng, các bảng đối chứng trước và sau tác động, đồng thời tăng cường minh chứng về phạm vi ảnh hưởng và khả năng nhân rộng.</p>
            <p class="bold uppercase" style="margin-top: 12pt; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif; font-weight: bold;">Xếp loại đề nghị: ${srcEval.classification || 'LOẠI KHÁ'}</p>
            <p class="bold" style="margin-top: 0; margin-bottom: 12pt; font-family: 'Times New Roman', Times, serif;">Tổng điểm: ${srcEval.totalScore}/100 điểm</p>
          </div>

          <!-- SIGNATURES -->
          <table style="width: 100%; border: none; margin-top: 24pt; font-family: 'Times New Roman', Times, serif; font-size: 14pt;">
            <tr style="border: none;">
              <td style="width: 45%; border: none;"></td>
              <td style="width: 55%; text-align: center; vertical-align: top; border: none; padding: 0;">
                <p class="italic" style="font-size: 14pt; text-align: center; margin-bottom: 4pt; font-family: 'Times New Roman', Times, serif;">${dateStr}</p>
                <p class="bold uppercase" style="font-size: 14pt; text-align: center; margin-bottom: 2pt; font-family: 'Times New Roman', Times, serif; font-weight: bold;">${signerTitle}</p>
                <p class="italic" style="font-size: 14pt; text-align: center; margin-bottom: 72pt; font-family: 'Times New Roman', Times, serif;">(Ký, ghi rõ họ và tên)</p>
                <p class="bold" style="font-size: 14pt; text-align: center; margin-top: 0; font-family: 'Times New Roman', Times, serif;">${isCouncilTarget ? 'ỦY VIÊN THƯ KÝ HỘI ĐỒNG' : (reviewerName || '........................................')}</p>
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>`;

      try {
        const docxBlob = await asBlob(msWordObj);
        saveAs(docxBlob as Blob, currentResult ? `PhieuGiamDinh_${currentResult.teacher.teacherName || 'TacGia'}.docx` : 'Phieu_Danh_Gia_SKKN.docx');
      } catch (err) {
        console.error('Lỗi khi convert docx bằng html-docx-js-typescript, thực hiện fallback sang doc thô:', err);
        const blob = new Blob(['\ufeff', msWordObj], {
          type: 'application/msword;charset=utf-8'
        });
        saveAs(blob, currentResult ? `PhieuGiamDinh_${currentResult.teacher.teacherName || 'TacGia'}.doc` : 'Phieu_Danh_Gia_SKKN.doc');
      }
    } catch (e) {
      console.error(e);
    } finally {
      setIsGeneratingWord(false);
    }
  };

  // Configuration settings for appraisal
  const [pronoun, setPronoun] = useState<'thay_co' | 'tac_gia'>('thay_co');
  const [evaluationMode, setEvaluationMode] = useState<'full' | 'comment_only'>('full');
  const [isCouncilAppraisal, setIsCouncilAppraisal] = useState(false);
  const [sortExcelBySchool, setSortExcelBySchool] = useState(false);
  const [excelIncludeRemarks, setExcelIncludeRemarks] = useState(true);
  const [reviewerName, setReviewerName] = useState('');

  // Plagiarism & Duplicate Checker states
  const [mainTab, setMainTab] = useState<'guide' | 'info' | 'plagiarism' | 'appraisal' | 'settings'>('info');
  const [activeTab, setActiveTab] = useState<'evaluate' | 'plagiarism'>('evaluate');
  const [plagText, setPlagText] = useState('');
  const [plagFileBase64, setPlagFileBase64] = useState('');
  const [plagFileUrl, setPlagFileUrl] = useState<string | null>(null);
  const [plagFileName, setPlagFileName] = useState('');
  const [isPlagLoading, setIsPlagLoading] = useState(false);
  const [plagError, setPlagError] = useState<string | null>(null);
  const [plagResult, setPlagResult] = useState<PlagiarismResult | null>(null);
  const [hoveredSourceId, setHoveredSourceId] = useState<string | null>(null);
  const [activeSourceId, setActiveSourceId] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isEditingPlag, setIsEditingPlag] = useState(false);

  const updatePlagiarismDataInResult = (newPlag: PlagiarismResult) => {
    setPlagResult(newPlag);
    
    if (currentResult) {
      setCurrentResult(prev => {
        if (!prev) return prev;
        
        let baseClass = 'Không đạt';
        const total = prev.totalScore || 0;
        if (total >= 90) {
          baseClass = 'Xuất sắc';
        } else if (total >= 80) {
          baseClass = 'Tốt';
        } else if (total >= 50) {
          baseClass = 'Khá';
        } else {
          baseClass = 'Không đạt';
        }
        
        let finalClass = baseClass;
        const levelLower = newPlag.warningLevel.toLowerCase();
        if (levelLower.includes('không đạt') || levelLower.includes('nghiêm trọng') || levelLower.includes('không đạt') || newPlag.totalDuplicatePercent > 30) {
          finalClass = 'Không đạt (Mức độ vi phạm đạo văn quá quy định)';
        } else if (newPlag.aiGeneratedPercent !== undefined && newPlag.aiGeneratedPercent > 10) {
          finalClass = 'Không đạt (Vượt quá giới hạn nội dung AI sinh)';
        }
        
        const updatedRes = {
          ...prev,
          classification: finalClass,
          plagiarismResult: newPlag
        };
        
        // Cập nhật lịch sử đồng bộ
        setHistory(prevHist => {
          const updatedHist = prevHist.map(h => h.id === prev.id ? updatedRes : h);
          syncToFirestore({ history: updatedHist });
          return updatedHist;
        });
        
        return updatedRes;
      });
    }
  };


  const [isExtractingInfo, setIsExtractingInfo] = useState(false);
  const [extractionStep, setExtractionStep] = useState(0);

  // Custom loader messages to fit Ham Yen, Tuyen Quang context
  const loaderMessages = [
    "Đang phân tích cấu trúc Báo cáo Sáng kiến...",
    "Đang đối chiếu với Quy chế hoạt động Hội đồng ban hành kèm Quyết định số 270/QĐ-HĐSK...",
    "Đang thẩm định Tiêu chí 1: Hiệu quả áp dụng (Hiệu quả kinh tế & lợi ích xã hội)...",
    "Đang thẩm định Tiêu chí 2: Phạm vi ảnh hưởng (Các minh chứng và số liệu đối chứng trước/sau)...",
    "Đang áp dụng logic chấm điểm và tự động phân loại xếp hạng...",
    "Đang rà soát và tinh chỉnh văn phong sư phạm của Hội đồng...",
    "Đang xuất kết quả báo thẩm định chi tiết..."
  ];

  const plagiarismLoaderMessages = [
    "Đang làm sạch và tiền xử lý văn bản báo cáo...",
    "Đang kết nối hệ thống dữ liệu học thuật trên internet...",
    "Đang quét sâu và đối chiếu với thư viện Sáng kiến tham khảo...",
    "Đang nhận diện các dấu hiệu cấu trúc văn bản do trí tuệ nhân tạo (AI) tạo ra...",
    "Đang kiểm tra lỗi chính tả và tính tiêu chuẩn của văn phong...",
    "Đang tổng hợp điểm tương đồng và trích xuất đường dẫn gốc...",
    "Đang hoàn thiện phiếu kết quả phân tích Đạo văn..."
  ];

  // Rotate loader messages during await
  useEffect(() => {
    let interval: any;
    if (isLoading) {
      interval = setInterval(() => {
        setLoadingMessageIndex((prev) => {
          const arr = mainTab === 'plagiarism' ? plagiarismLoaderMessages : loaderMessages;
          if (prev < arr.length - 1) {
             return prev + 1;
          }
          return prev; // Stop at the last message to avoid losing timeline progress
        });
      }, 5000);
    } else {
      setLoadingMessageIndex(0);
    }
    return () => clearInterval(interval);
  }, [isLoading, mainTab]);

  // Sync with Firestore
  const syncToFirestore = async (dataToSync: any) => {
    if (!user) return;
    try {
      await setDoc(doc(db, 'users', user.uid), dataToSync, { merge: true });
    } catch (error) {
      console.error("Error saving to Firestore:", error);
    }
  };

  // Load user data on auth change
  useEffect(() => {
    const loadUserData = async () => {
      if (user) {
        try {
          const docRef = doc(doc(db, 'users', user.uid).firestore, 'users', user.uid);
          const docSnap = await getDoc(docRef);
          if (docSnap.exists()) {
            const data = docSnap.data();
            
            if (data.history && Array.isArray(data.history)) {
              setHistory(data.history);
              if (data.history.length > 0) {
                const first = data.history[0];
                setCurrentResult(first);
                setPlagResult(first.plagiarismResult || null);
                if (first.teacher) setTeacher(first.teacher);
                if (first.initiativeTitle) setInitiativeTitle(first.initiativeTitle);
                if (first.appliedDate) setAppliedDate(first.appliedDate);
                if (first.initiativeText) setInitiativeText(first.initiativeText);
                setCouncilName(first.councilName || 'CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN');
                setMember1Name(first.member1Name || '');
                setMember1Unit(first.member1Unit || '');
                setMember1Role(first.member1Role || '');
                setMember2Name(first.member2Name || '');
                setMember2Unit(first.member2Unit || '');
                setMember2Role(first.member2Role || '');
                setMember3Name(first.member3Name || '');
                setMember3Unit(first.member3Unit || '');
                setMember3Role(first.member3Role || '');
                
                const loadedMembers: CouncilMember[] = first.councilMembers || [
                  { id: '1', name: first.member1Name || '', unit: first.member1Unit || '', role: first.member1Role || '' },
                  { id: '2', name: first.member2Name || '', unit: first.member2Unit || '', role: first.member2Role || '' },
                  { id: '3', name: first.member3Name || '', unit: first.member3Unit || '', role: first.member3Role || '' },
                ];
                setCouncilMembers(loadedMembers);
              }
            }
            if (data.apiKeys && Array.isArray(data.apiKeys)) {
              setApiKeys(data.apiKeys);
            }
            if (data.reviewerName) {
              setReviewerName(data.reviewerName);
            }
            if (data.councilName !== undefined) setCouncilName(data.councilName);
            if (data.member1Name !== undefined) setMember1Name(data.member1Name);
            if (data.member1Unit !== undefined) setMember1Unit(data.member1Unit);
            if (data.member1Role !== undefined) setMember1Role(data.member1Role);
            if (data.member2Name !== undefined) setMember2Name(data.member2Name);
            if (data.member2Unit !== undefined) setMember2Unit(data.member2Unit);
            if (data.member2Role !== undefined) setMember2Role(data.member2Role);
            if (data.member3Name !== undefined) setMember3Name(data.member3Name);
            if (data.member3Unit !== undefined) setMember3Unit(data.member3Unit);
            if (data.member3Role !== undefined) setMember3Role(data.member3Role);

            if (data.councilMembers && Array.isArray(data.councilMembers)) {
              setCouncilMembers(data.councilMembers);
            } else if (data.member1Name !== undefined || data.member2Name !== undefined || data.member3Name !== undefined) {
              setCouncilMembers([
                { id: '1', name: data.member1Name || '', unit: data.member1Unit || '', role: data.member1Role || '' },
                { id: '2', name: data.member2Name || '', unit: data.member2Unit || '', role: data.member2Role || '' },
                { id: '3', name: data.member3Name || '', unit: data.member3Unit || '', role: data.member3Role || '' },
              ]);
            }
          }
        } catch (error) {
          console.error("Error loading user data from Firestore:", error);
        }
      } else {
        // Reset state on logout
        setHistory([]);
        setCurrentResult(null);
        setPlagResult(null);
        setApiKeys(['']);
        setReviewerName('');
        setCouncilName('CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN');
        setMember1Name('');
        setMember1Unit('');
        setMember1Role('');
        setMember2Name('');
        setMember2Unit('');
        setMember2Role('');
        setMember3Name('');
        setMember3Unit('');
        setMember3Role('');
        setCouncilMembers([
          { id: '1', name: '', unit: '', role: '' },
          { id: '2', name: '', unit: '', role: '' },
          { id: '3', name: '', unit: '', role: '' },
        ]);
      }
    };
    loadUserData();
  }, [user]);

  const handleKeyChange = (index: number, value: string) => {
    const newKeys = [...apiKeys];
    newKeys[index] = value;
    setApiKeys(newKeys);
    syncToFirestore({ apiKeys: newKeys });
  };

  const handleAddKey = () => {
    const newKeys = [...apiKeys, ''];
    setApiKeys(newKeys);
    syncToFirestore({ apiKeys: newKeys });
  };

  const handleRemoveKey = (index: number) => {
    if (apiKeys.length <= 1) return;
    const newKeys = apiKeys.filter((_, i) => i !== index);
    setApiKeys(newKeys);
    syncToFirestore({ apiKeys: newKeys });
  };

  const toggleKeyVisibility = (index: number) => {
    setVisibleKeys(prev => ({ ...prev, [index]: !prev[index] }));
  };

  // Save history
  const saveToHistory = (newResult: EvaluationResult) => {
    const updated = [newResult, ...history.filter(h => h.id !== newResult.id)];
    setHistory(updated);
    syncToFirestore({ history: updated });
  };

  const deleteHistoryItem = (id: string, e: MouseEvent) => {
    e.stopPropagation();
    const updated = history.filter(h => h.id !== id);
    setHistory(updated);
    syncToFirestore({ history: updated });
    if (currentResult?.id === id) {
      const nextRes = updated.length > 0 ? updated[0] : null;
      setCurrentResult(nextRes);
      setPlagResult(nextRes ? nextRes.plagiarismResult || null : null);
      if (nextRes) {
        if (nextRes.teacher) setTeacher(nextRes.teacher);
        if (nextRes.initiativeTitle) setInitiativeTitle(nextRes.initiativeTitle);
        if (nextRes.appliedDate) setAppliedDate(nextRes.appliedDate);
        if (nextRes.initiativeText) setInitiativeText(nextRes.initiativeText);
      } else {
        setInitiativeTitle('');
        setInitiativeText('');
      }
    }
  };

  const handleExportExcel = async () => {
    if (history.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Danh sách SKKN');

    // Mở rộng các ô tiêu đề
    worksheet.mergeCells('A1:J1');
    const title1 = worksheet.getCell('A1');
    title1.value = 'DANH SÁCH';
    title1.font = { name: 'Times New Roman', size: 14, bold: true };
    title1.alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.mergeCells('A2:J2');
    const title2 = worksheet.getCell('A2');
    const currentYear = new Date().getFullYear();
    title2.value = `ĐỀ NGHỊ CÔNG NHẬN SÁNG KIẾN CÓ HIỆU QUẢ ÁP DỤNG, PHẠM VI ẢNH HƯỞNG CẤP XÃ, NĂM ${currentYear}`;
    title2.font = { name: 'Times New Roman', size: 13, bold: true };
    title2.alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.mergeCells('A3:J3');
    const title3 = worksheet.getCell('A3');
    title3.value = `(Kèm theo Quyết định số       /QĐ-UBND ngày     tháng     năm ${currentYear} của Ủy ban nhân dân xã Hàm Yên)`;
    title3.font = { name: 'Times New Roman', size: 13, italic: true };
    title3.alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.addRow([]); // Dòng trống

    // Header bảng (Dòng 5 và 6)
    worksheet.mergeCells('A5:A6');
    worksheet.mergeCells('B5:B6');
    worksheet.mergeCells('C5:C6');
    worksheet.mergeCells('D5:D6');
    worksheet.mergeCells('E5:E6');
    worksheet.mergeCells('F5:F6');
    worksheet.mergeCells('G5:G6');
    worksheet.mergeCells('H5:I5');
    worksheet.mergeCells('J5:J6');

    const headerNames = [
      { cell: 'A5', label: 'STT' },
      { cell: 'B5', label: 'Tên sáng kiến kinh nghiệm' },
      { cell: 'C5', label: 'Họ và tên tác giả,\nnhóm tác giả' },
      { cell: 'D5', label: 'Chức vụ' },
      { cell: 'E5', label: 'Đơn vị công tác' },
      { cell: 'F5', label: 'Ý kiến thẩm định của\nphòng Văn hóa - Xã hội' },
      { cell: 'G5', label: 'Ý kiến của hội\nđồng thẩm định' },
      { cell: 'H5', label: 'Kết quả đánh giá của\nhội đồng thẩm định' },
      { cell: 'H6', label: 'Đạt' },
      { cell: 'I6', label: 'Không đạt' },
      { cell: 'J5', label: 'Ghi chú' }
    ];

    headerNames.forEach(h => {
      const cell = worksheet.getCell(h.cell);
      cell.value = h.label;
      cell.font = { name: 'Times New Roman', size: 12, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });

    // Apply borders to header cells (row 5 and 6, columns A to J)
    for (let c = 1; c <= 10; c++) {
      ['5', '6'].forEach(r => {
        const cell = worksheet.getCell(`${String.fromCharCode(64 + c)}${r}`);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    }

    // Chi tiết bảng
    let sortedHistory = [...history];
    if (sortExcelBySchool) {
      sortedHistory.sort((a, b) => (a.teacher.schoolName || "").localeCompare(b.teacher.schoolName || ""));
    }

    sortedHistory.forEach((h, idx) => {
      let isPass = "";
      let isFail = "";
      let totalScore = 0;
      
      totalScore = h.totalScore || 0;

      if (h.classification) {
        if (h.classification.includes('Không đạt')) {
          isFail = 'Không đạt';
        } else {
          isPass = 'Đạt';
        }
      } else {
        if (totalScore >= 50) {
          isPass = "Đạt";
        } else {
          isFail = "Không đạt";
        }
      }

      // Trích xuất ít nhất 2 lời nhận xét chi tiết để ghi vào biểu Excel
      let remarks: string[] = [];
      if (h.uuDiem && h.uuDiem.length > 0) {
        remarks = h.uuDiem.map(u => u.trim().replace(/^-\s*/, '')).filter(u => u.length > 0);
      }
      
      // Nếu chưa đủ 2 nhận xét, lấy thêm từ phần phân tích chi tiết của các tiêu chí
      if (remarks.length < 2) {
        const potentialSources = [
          ...(h.suCanThiet?.analysis || []),
          ...(h.tinhMoi?.analysis || []),
          ...(h.giaiPhap?.analysis || []),
          ...(h.hieuQua?.analysis || []),
          ...(h.khaNangApDung?.analysis || [])
        ];
        potentialSources.forEach(s => {
          if (remarks.length < 2 && s && s.trim().length > 10) {
            remarks.push(s.trim());
          }
        });
      }
      
      // Nếu vẫn chưa đủ, bổ sung các nhận xét chuyên môn tiêu chuẩn
      if (remarks.length < 2) {
        const defaultComments = [
          "Giải pháp khoa học, thiết thực và có tính khả thi cao tại đơn vị.",
          "Nội dung trình bày rõ ràng, bám sát các tiêu chuẩn kỹ thuật chuyên môn.",
          "Đóng góp tích cực vào việc đổi mới phương pháp và nâng cao chất lượng công tác."
        ];
        while (remarks.length < 2) {
          remarks.push(defaultComments[remarks.length]);
        }
      }

      const finalRemarks = remarks.slice(0, 2).map((r, i) => `${i + 1}. ${r}`).join('\n');
      const vhxhComment = excelIncludeRemarks 
        ? `Tổng điểm: ${totalScore}/100.\nĐánh giá sơ bộ: ${isPass || isFail}.\nNhận xét:\n${finalRemarks}`
        : `Tổng điểm: ${totalScore}/100.\nĐánh giá sơ bộ: ${isPass || isFail}.`;
      
      const isCouncil = Boolean(isCouncilAppraisal || h.isCouncilAppraisal);
      let councilOpinion = "";
      let councilPass = "";
      let councilFail = "";

      if (isCouncil) {
        councilOpinion = excelIncludeRemarks
          ? `Đồng ý với đánh giá sơ bộ.\nTổng điểm: ${totalScore}/100.\nNhận xét:\n${finalRemarks}`
          : `Đồng ý với đánh giá sơ bộ.\nTổng điểm: ${totalScore}/100.`;
        if (isPass === "Đạt") councilPass = "X";
        if (isFail === "Không đạt") councilFail = "X";
      }

      const row = worksheet.addRow([
        (idx + 1).toString(),
        h.initiativeTitle || "Báo cáo sáng kiến kinh nghiệm (Chưa xác định)",
        h.teacher.teacherName || "Đang cập nhật...",
        h.teacher.role || "Giáo viên",
        h.teacher.schoolName || "",
        vhxhComment, // Ý kiến thẩm định của phòng Văn hóa - Xã hội
        councilOpinion, // Ý kiến của hội đồng thẩm định
        councilPass, // Đạt
        councilFail, // Không đạt
        ""  // Ghi chú
      ]);
      
      row.eachCell((cell, colNumber) => {
        cell.font = { name: 'Times New Roman', size: 12 };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        // Căn giữa cho STT, Đạt, Không đạt, Ghi chú
        if (colNumber === 1 || colNumber >= 8) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else {
          cell.alignment = { vertical: 'middle', wrapText: true };
        }
      });
    });

    // Điều chỉnh độ rộng cột
    worksheet.getColumn(1).width = 6;
    worksheet.getColumn(2).width = 40;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 25;
    worksheet.getColumn(6).width = 25;
    worksheet.getColumn(7).width = 25;
    worksheet.getColumn(8).width = 10;
    worksheet.getColumn(9).width = 12;
    worksheet.getColumn(10).width = 12;

    // Xuất file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, "Danh sách SKKN kem theo.xlsx");
  };

  const getFriendlyErrorMessage = (errorMsg: string) => {
    try {
      if (typeof errorMsg !== 'string') return errorMsg;
      if (errorMsg.includes('has been suspended') || errorMsg.includes('PERMISSION_DENIED')) {
        return "API Key của nền tảng đã bị Google khóa hạn mức. Bạn vui lòng tự tạo API Key cá nhân (miễn phí) tại https://aistudio.google.com/app/apikey và dán vào Kho lưu trữ API Key phía dưới bài nhé.";
      }
      if (errorMsg.includes('API key not valid') || errorMsg.includes('INVALID_ARGUMENT')) {
        return "API Key bạn cung cấp không hợp lệ. Vui lòng kiểm tra lại cấu hình kho API Key.";
      }
      if (errorMsg.includes('Quota') || errorMsg.includes("429") || errorMsg.includes("exhausted")) {
         return "API Key hiện tại đã kiệt sức (Quota Exceeded). Vui lòng đổi lịch trình hoặc dán thêm API Key từ mọt tài khoản gmail khác để sử dụng tiếp.";
      }
      if (errorMsg.includes('503') || errorMsg.includes("high demand") || errorMsg.includes("UNAVAILABLE")) {
         return "Hệ thống AI hiện đang quá tải do có quá nhiều người sử dụng trên toàn cầu. Vui lòng thử đổi sang mô hình khác (ví dụ: Gemini 2.5 Flash / Gemini 3.5 Flash) hoặc thử lại sau vài giây.";
      }
      if (errorMsg.includes('Failed to fetch') || errorMsg.includes('fetch failed')) {
         return "Không thể kết nối tới máy chủ AI. Vui lòng kiểm tra lại kết nối mạng của bạn.";
      }
      return `${errorMsg}`;
    } catch {
      return errorMsg;
    }
  };

  // Internal appraisal calling worker
  const runAppraisalEvaluation = async (
    textToEval: string,
    titleToEval: string,
    teacherToEval: TeacherInfo,
    associatedPlagResult?: PlagiarismResult | null,
    requestedScore?: number | null,
    reEvaluationNotes?: string | null
  ) => {
    try {
      const response = await fetch('/api/evaluate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          apiKeys,
          model: selectedModelAppraisal,
          teacher: {
            ...teacherToEval,
            teacherName: teacherToEval.teacherName?.normalize('NFC') || '',
            schoolName: teacherToEval.schoolName?.normalize('NFC') || '',
            role: teacherToEval.role?.normalize('NFC') || '',
            subject: teacherToEval.subject?.normalize('NFC') || ''
          },
          initiativeTitle: titleToEval.normalize('NFC'),
          appliedDate,
          initiativeText: textToEval.normalize('NFC'),
          fileBase64: plagFileBase64,
          fileName: plagFileName ? plagFileName.normalize('NFC') : '',
          pronoun,
          evaluationMode,
          requestedScore,
          reEvaluationNotes
        })
      });

      if (!response.ok) {
        if (response.status === 413) {
          throw new Error('File tải lên quá lớn (Vượt giới hạn 4.5MB của Vercel). Vui lòng nén file PDF hoặc chuyển sang chế độ Copy/Paste văn bản.');
        }
        let errData;
        try { errData = await response.json(); } catch (e) { throw new Error(`Lỗi máy chủ (${response.status}): ${response.statusText}`); }
        throw new Error(getFriendlyErrorMessage(errData?.details || errData?.error || 'Lỗi hệ thống'));
      }

      const evalData = await response.json();
      
      // Parse new detailed scores
      const calcScore = (obj: any, fallback: number) => Number(obj?.score) || fallback;
      
      const sScore = calcScore(evalData.suCanThiet, 8);
      const mScore = calcScore(evalData.tinhMoi, 16);
      const gScore = calcScore(evalData.giaiPhap, 24);
      const hScore = calcScore(evalData.hieuQua, 24);
      const kScore = calcScore(evalData.khaNangApDung, 8);
      const total = sScore + mScore + gScore + hScore + kScore;
      
      let finalClass = 'Không đạt';
      if (total >= 90) {
        finalClass = 'Xuất sắc';
      } else if (total >= 80) {
        finalClass = 'Tốt';
      } else if (total >= 50) {
        finalClass = 'Khá';
      } else {
        finalClass = 'Không đạt';
      }

      // Check regulations
      if (associatedPlagResult) {
         if (associatedPlagResult.warningLevel.includes('Không đạt') || associatedPlagResult.warningLevel.includes('nghiêm trọng')) {
             finalClass = 'Không đạt (Mức độ vi phạm đạo văn quá quy định)';
         }
         if (associatedPlagResult.aiGeneratedPercent !== undefined && associatedPlagResult.aiGeneratedPercent > 10) {
             finalClass = 'Không đạt (Vượt quá giới hạn nội dung AI sinh)';
         }
      }

      const generatedResult: EvaluationResult = {
        id: 'skkn_' + Date.now(),
        teacher: { ...teacherToEval },
        initiativeTitle: titleToEval,
        appliedDate,
        initiativeText: textToEval,
        evaluatedAt: new Date().toISOString(),
        suCanThiet: {
          score: sScore,
          levelName: evalData.suCanThiet?.levelName || '',
          analysis: evalData.suCanThiet?.analysis || [],
          pros: evalData.suCanThiet?.pros || []
        },
        tinhMoi: {
          score: mScore,
          levelName: evalData.tinhMoi?.levelName || '',
          analysis: evalData.tinhMoi?.analysis || [],
          pros: evalData.tinhMoi?.pros || [],
          cons: evalData.tinhMoi?.cons || []
        },
        giaiPhap: {
          score: gScore,
          levelName: evalData.giaiPhap?.levelName || '',
          analysis: evalData.giaiPhap?.analysis || [],
          pros: evalData.giaiPhap?.pros || [],
          cons: evalData.giaiPhap?.cons || []
        },
        hieuQua: {
          score: hScore,
          levelName: evalData.hieuQua?.levelName || '',
          analysis: evalData.hieuQua?.analysis || [],
          cons: evalData.hieuQua?.cons || []
        },
        khaNangApDung: {
          score: kScore,
          levelName: evalData.khaNangApDung?.levelName || '',
          analysis: evalData.khaNangApDung?.analysis || [],
          cons: evalData.khaNangApDung?.cons || []
        },
        uuDiem: evalData.uuDiem || [],
        hanChe: evalData.hanChe || [],
        improvements: evalData.improvements || [],
        summary: evalData.summary || 'Sáng kiến kinh nghiệm.',
        totalScore: total,
        classification: finalClass,
        pronoun,
        evaluationMode,
        isCouncilAppraisal,
        plagiarismResult: associatedPlagResult || null,
        councilName,
        member1Name,
        member1Unit,
        member1Role,
        member2Name,
        member2Unit,
        member2Role,
        member3Name,
        member3Unit,
        member3Role,
        councilMembers
      };

      setCurrentResult(generatedResult);
      saveToHistory(generatedResult);
      setMainTab('appraisal');
    } catch (err: any) {
      console.error(err);
      setErrorMsg(`Không thể hoàn thành việc thẩm định ở Hội đồng: ${getFriendlyErrorMessage(err.message)}. Vui lòng thử lại.`);
    }
  };



  const handleStartPlagiarismCheckOnly = async () => {
    if (!plagFileBase64 && (!initiativeText || initiativeText.trim().length < 30)) {
      setErrorMsg("Vui lòng tải lên một tệp Sáng kiến (.PDF/.TXT) hoặc dán dầy đủ nội dung để kiểm tra.");
      return;
    }

    setIsLoading(true);
    setErrorMsg(null);
    setPlagError(null);
    setPlagResult(null);
    setLoadingMessageIndex(0); // Bắt đầu kết nối
    setMainTab('plagiarism');

    try {
      const response = await fetch('/api/plagiarism-check', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          apiKeys,
          model: selectedModelPlag,
          text: (plagFileBase64 ? '' : initiativeText).normalize('NFC'),
          fileBase64: plagFileBase64,
          fileName: plagFileName ? plagFileName.normalize('NFC') : ''
        })
      });

      if (!response.ok) {
        if (response.status === 413) {
          throw new Error('File tải lên quá lớn (Vượt giới hạn 4.5MB của Vercel). Vui lòng nén file PDF hoặc chuyển sang chế độ Copy/Paste văn bản.');
        }
        let errData;
        try { errData = await response.json(); } catch (e) { throw new Error(`Lỗi máy chủ (${response.status}): ${response.statusText}`); }
        throw new Error(getFriendlyErrorMessage(errData?.details || errData?.error || 'Lỗi hệ thống rà soát trùng lặp.'));
      }

      const data = await response.json();
      
      const instantiatedPlag: PlagiarismResult = {
        id: 'plag_' + Date.now(),
        title: plagFileName || initiativeTitle || 'Văn bản quét tự dán',
        totalDuplicatePercent: data.totalDuplicatePercent,
        warningLevel: data.warningLevel,
        extractedText: data.extractedText,
        sources: data.sources || [],
        segments: data.segments || [],
        aiGeneratedPercent: data.aiGeneratedPercent,
        aiSegments: data.aiSegments || [],
        spellingErrors: data.spellingErrors || [],
        checkedAt: new Date().toISOString()
      };
      
      updatePlagiarismDataInResult(instantiatedPlag);
      
      let finalTitle = initiativeTitle;
      if (data.extractedTitle && !initiativeTitle) {
        finalTitle = data.extractedTitle;
        setInitiativeTitle(finalTitle);
      }
      
      if (data.extractedText && (!initiativeText || initiativeText.includes("📂") || initiativeText.includes("📁") || initiativeText.length < 50)) {
        setInitiativeText(data.extractedText);
      }
      
    } catch (err: any) {
      console.error(err);
      setErrorMsg(`Quá trình kiểm tra trùng lặp thất bại: ${getFriendlyErrorMessage(err.message)}. Vui lòng thử lại.`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleStartAppraisalOnly = async () => {
    if (!plagFileBase64 && (!initiativeText || initiativeText.trim().length < 30)) {
      setErrorMsg("Vui lòng tải lên một tệp Sáng kiến (.PDF/.TXT) hoặc dán dầy đủ nội dung để thẩm định.");
      return;
    }

    setMainTab('appraisal');
    setIsLoading(true);
    setLoadingMessageIndex(0);
    setErrorMsg(null);
    try {
      await runAppraisalEvaluation(
        initiativeText || (plagResult?.extractedText || "Sáng kiến kinh nghiệm"), 
        initiativeTitle || "Sáng kiến kinh nghiệm", 
        teacher, 
        plagResult
      );
    } catch (e: any) {
      setErrorMsg(`Quá trình phân tích thẩm định thất bại: ${e.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleTriggerReEvaluation = async () => {
    if (!currentResult) return;
    setIsLoading(true);
    setLoadingMessageIndex(0);
    setErrorMsg(null);
    try {
      const textToUse = initiativeText || currentResult.initiativeText || (plagResult?.extractedText || "Sáng kiến kinh nghiệm");
      const titleToUse = initiativeTitle || currentResult.initiativeTitle || "Sáng kiến kinh nghiệm";
      const teacherToUse = teacher || currentResult.teacher;

      const isCouncil = activeAppraisalView === 'council';

      const response = await fetch('/api/evaluate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          apiKeys,
          model: selectedModelAppraisal,
          teacher: {
            ...teacherToUse,
            teacherName: teacherToUse.teacherName?.normalize('NFC') || '',
            schoolName: teacherToUse.schoolName?.normalize('NFC') || '',
            role: teacherToUse.role?.normalize('NFC') || '',
            subject: teacherToUse.subject?.normalize('NFC') || ''
          },
          initiativeTitle: titleToUse.normalize('NFC'),
          appliedDate,
          initiativeText: textToUse.normalize('NFC'),
          fileBase64: plagFileBase64,
          fileName: plagFileName ? plagFileName.normalize('NFC') : '',
          pronoun,
          evaluationMode,
          requestedScore: desiredScore === '' ? null : desiredScore,
          reEvaluationNotes: reAppraisalNotes === '' ? null : reAppraisalNotes,
          evaluateTarget: isCouncil ? 'council' : 'department'
        })
      });

      if (!response.ok) {
        let errData;
        try { errData = await response.json(); } catch (e) { throw new Error(`Lỗi máy chủ (${response.status}): ${response.statusText}`); }
        throw new Error(getFriendlyErrorMessage(errData?.details || errData?.error || 'Lỗi hệ thống'));
      }

      const evalData = await response.json();
      
      const calcScore = (obj: any, fallback: number) => Number(obj?.score) || fallback;
      const sScore = calcScore(evalData.suCanThiet, 8);
      const mScore = calcScore(evalData.tinhMoi, 16);
      const gScore = calcScore(evalData.giaiPhap, 24);
      const hScore = calcScore(evalData.hieuQua, 24);
      const kScore = calcScore(evalData.khaNangApDung, 8);
      const total = sScore + mScore + gScore + hScore + kScore;
      
      let finalClass = 'Không đạt';
      if (total >= 90) {
        finalClass = 'Xuất sắc';
      } else if (total >= 80) {
        finalClass = 'Tốt';
      } else if (total >= 50) {
        finalClass = 'Khá';
      } else {
        finalClass = 'Không đạt';
      }

      if (isCouncil) {
        const councilEval: CouncilEvaluationResult = {
          suCanThiet: {
            score: sScore,
            levelName: evalData.suCanThiet?.levelName || '',
            analysis: evalData.suCanThiet?.analysis || [],
            pros: evalData.suCanThiet?.pros || []
          },
          tinhMoi: {
            score: mScore,
            levelName: evalData.tinhMoi?.levelName || '',
            analysis: evalData.tinhMoi?.analysis || [],
            pros: evalData.tinhMoi?.pros || [],
            cons: evalData.tinhMoi?.cons || []
          },
          giaiPhap: {
            score: gScore,
            levelName: evalData.giaiPhap?.levelName || '',
            analysis: evalData.giaiPhap?.analysis || [],
            pros: evalData.giaiPhap?.pros || [],
            cons: evalData.giaiPhap?.cons || []
          },
          hieuQua: {
            score: hScore,
            levelName: evalData.hieuQua?.levelName || '',
            analysis: evalData.hieuQua?.analysis || [],
            cons: evalData.hieuQua?.cons || []
          },
          khaNangApDung: {
            score: kScore,
            levelName: evalData.khaNangApDung?.levelName || '',
            analysis: evalData.khaNangApDung?.analysis || [],
            cons: evalData.khaNangApDung?.cons || []
          },
          summary: evalData.summary || 'Sáng kiến kinh nghiệm.',
          totalScore: total,
          classification: finalClass,
          evaluatedAt: new Date().toISOString()
        };

        const updatedResult: EvaluationResult = {
          ...currentResult,
          councilResult: councilEval
        };

        setCurrentResult(updatedResult);
        saveToHistory(updatedResult);
      } else {
        const generatedResult: EvaluationResult = {
          ...currentResult,
          suCanThiet: {
            score: sScore,
            levelName: evalData.suCanThiet?.levelName || '',
            analysis: evalData.suCanThiet?.analysis || [],
            pros: evalData.suCanThiet?.pros || []
          },
          tinhMoi: {
            score: mScore,
            levelName: evalData.tinhMoi?.levelName || '',
            analysis: evalData.tinhMoi?.analysis || [],
            pros: evalData.tinhMoi?.pros || [],
            cons: evalData.tinhMoi?.cons || []
          },
          giaiPhap: {
            score: gScore,
            levelName: evalData.giaiPhap?.levelName || '',
            analysis: evalData.giaiPhap?.analysis || [],
            pros: evalData.giaiPhap?.pros || [],
            cons: evalData.giaiPhap?.cons || []
          },
          hieuQua: {
            score: hScore,
            levelName: evalData.hieuQua?.levelName || '',
            analysis: evalData.hieuQua?.analysis || [],
            cons: evalData.hieuQua?.cons || []
          },
          khaNangApDung: {
            score: kScore,
            levelName: evalData.khaNangApDung?.levelName || '',
            analysis: evalData.khaNangApDung?.analysis || [],
            cons: evalData.khaNangApDung?.cons || []
          },
          uuDiem: evalData.uuDiem || [],
          hanChe: evalData.hanChe || [],
          improvements: evalData.improvements || [],
          summary: evalData.summary || 'Sáng kiến kinh nghiệm.',
          totalScore: total,
          classification: finalClass
        };

        setCurrentResult(generatedResult);
        saveToHistory(generatedResult);
      }
      
      setShowReAppraisalPanel(false);
      setDesiredScore('');
      setReAppraisalNotes('');
    } catch (e: any) {
      setErrorMsg(`Chấm điểm lại thất bại: ${e.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const runCouncilAppraisal = async () => {
    if (!currentResult) return;
    setIsLoading(true);
    setLoadingMessageIndex(0);
    setErrorMsg(null);
    try {
      const response = await fetch('/api/evaluate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          apiKeys,
          model: selectedModelAppraisal,
          teacher: {
            ...currentResult.teacher,
            teacherName: currentResult.teacher.teacherName?.normalize('NFC') || '',
            schoolName: currentResult.teacher.schoolName?.normalize('NFC') || '',
            role: currentResult.teacher.role?.normalize('NFC') || '',
            subject: currentResult.teacher.subject?.normalize('NFC') || ''
          },
          initiativeTitle: currentResult.initiativeTitle.normalize('NFC'),
          appliedDate: currentResult.appliedDate || appliedDate,
          initiativeText: currentResult.initiativeText.normalize('NFC'),
          fileBase64: plagFileBase64,
          fileName: plagFileName ? plagFileName.normalize('NFC') : '',
          pronoun,
          evaluationMode,
          evaluateTarget: 'council'
        })
      });

      if (!response.ok) {
        let errData;
        try { errData = await response.json(); } catch (e) { throw new Error(`Lỗi máy chủ (${response.status}): ${response.statusText}`); }
        throw new Error(getFriendlyErrorMessage(errData?.details || errData?.error || 'Lỗi hệ thống'));
      }

      const evalData = await response.json();
      
      const calcScore = (obj: any, fallback: number) => Number(obj?.score) || fallback;
      const sScore = calcScore(evalData.suCanThiet, 8);
      const mScore = calcScore(evalData.tinhMoi, 16);
      const gScore = calcScore(evalData.giaiPhap, 24);
      const hScore = calcScore(evalData.hieuQua, 24);
      const kScore = calcScore(evalData.khaNangApDung, 8);
      const total = sScore + mScore + gScore + hScore + kScore;
      
      let finalClass = 'Không đạt';
      if (total >= 90) {
        finalClass = 'Xuất sắc';
      } else if (total >= 80) {
        finalClass = 'Tốt';
      } else if (total >= 50) {
        finalClass = 'Khá';
      } else {
        finalClass = 'Không đạt';
      }

      const councilEval: CouncilEvaluationResult = {
        suCanThiet: {
          score: sScore,
          levelName: evalData.suCanThiet?.levelName || '',
          analysis: evalData.suCanThiet?.analysis || [],
          pros: evalData.suCanThiet?.pros || []
        },
        tinhMoi: {
          score: mScore,
          levelName: evalData.tinhMoi?.levelName || '',
          analysis: evalData.tinhMoi?.analysis || [],
          pros: evalData.tinhMoi?.pros || [],
          cons: evalData.tinhMoi?.cons || []
        },
        giaiPhap: {
          score: gScore,
          levelName: evalData.giaiPhap?.levelName || '',
          analysis: evalData.giaiPhap?.analysis || [],
          pros: evalData.giaiPhap?.pros || [],
          cons: evalData.giaiPhap?.cons || []
        },
        hieuQua: {
          score: hScore,
          levelName: evalData.hieuQua?.levelName || '',
          analysis: evalData.hieuQua?.analysis || [],
          cons: evalData.hieuQua?.cons || []
        },
        khaNangApDung: {
          score: kScore,
          levelName: evalData.khaNangApDung?.levelName || '',
          analysis: evalData.khaNangApDung?.analysis || [],
          cons: evalData.khaNangApDung?.cons || []
        },
        summary: evalData.summary || 'Sáng kiến kinh nghiệm.',
        totalScore: total,
        classification: finalClass,
        evaluatedAt: new Date().toISOString()
      };

      const updatedResult: EvaluationResult = {
        ...currentResult,
        councilResult: councilEval
      };

      setCurrentResult(updatedResult);
      saveToHistory(updatedResult);
      setActiveAppraisalView('council');
    } catch (e: any) {
      setErrorMsg(`Thẩm định ở Hội đồng thất bại: ${e.message}`);
    } finally {
      setIsLoading(false);
    }
  };


  // Override bypass method if duplication exceeds limit but user wants to proceed
  const handleBypassPlagiarismAndEvaluate = async () => {
    if (!plagResult) return;
    setIsLoading(true);
    setErrorMsg(null);
    try {
      const textToUse = initiativeText || plagResult.extractedText;
      const finalTitle = initiativeTitle || "Sáng kiến kinh nghiệm trích xuất đặc cách";
      await runAppraisalEvaluation(textToUse, finalTitle, teacher, plagResult);
    } catch (e: any) {
      setErrorMsg(`Lỗi thẩm định đặc cách: ${getFriendlyErrorMessage(e.message)}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    processUploadedFile(file);
  };

  const processUploadedFile = async (file: File) => {
    setPlagFileName(file.name);
    setPlagError(null);
    setErrorMsg(null);
    setPlagFileUrl(URL.createObjectURL(file));

    setIsExtractingInfo(true);
    setExtractionStep(1); // Mức 1: Bắt đầu đọc file

    let base64Data = '';
    let textData = '';
    
    try {
      if (file.name.endsWith('.pdf')) {
        setInitiativeText(`📁 Đang đọc và định dạng văn bản từ PDF ("${file.name}")...`);
        textData = await extractTextFromPDFFile(file);
        if (!textData || textData.trim() === '') {
          throw new Error('Không tìm thấy chữ trong PDF (Có thể đây là PDF dạng ảnh/scan). Vui lòng dùng file PDF chứa văn bản để hệ thống có thể phân tích.');
        }
        setInitiativeText('🔄 Đang định dạng văn bản bằng AI (Markdown)...');
        try {
          const formatRes = await fetch('/api/format-markdown', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ text: textData, apiKeys, model: selectedModelPlag })
          });
          const formatData = await formatRes.json();
          if (formatData.formattedText) {
              textData = formatData.formattedText;
          }
        } catch(e) {
          console.error('Format Markdown Error:', e);
        }
        setInitiativeText(textData);
        
        // Fallback for large files to avoid Vercel 4.5MB limit: only pass text to server
        if (file.size > 3.5 * 1024 * 1024) { // > 3.5MB
          setPlagFileBase64(''); // Đánh dấu không có base64
        } else {
          // Standard base64 parsing for smaller files
          base64Data = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target?.result as string);
            reader.onerror = reject;
            reader.readAsDataURL(file);
          });
          setPlagFileBase64(base64Data);
        }
      } else if (file.name.endsWith('.txt')) {
        textData = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => resolve(e.target?.result as string);
          reader.onerror = reject;
          reader.readAsText(file);
        });
        setInitiativeText(textData);
        setPlagFileBase64('');
      } else {
        setErrorMsg("Hệ thống hiện hỗ trợ tải lên trực tiếp tệp tin .PDF hoặc .TXT. Vui lòng tải lên đúng định dạng.");
        setIsExtractingInfo(false);
        return;
      }
    } catch (e: any) {
      setErrorMsg(`Lỗi khi đọc file: ${e.message}`);
      setIsExtractingInfo(false);
      return;
    }

    if (!base64Data && !textData) {
      setIsExtractingInfo(false);
      return;
    }

    setExtractionStep(2); // Mức 2: Gửi API để AI trích xuất thông tin

    // Call extract-info
    try {
      const response = await fetch('/api/extract-info', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          apiKeys,
          model: selectedModelExtract,
          fileName: file.name,
          fileBase64: base64Data,
          text: textData
        })
      });

      if (response.ok) {
        const data = await response.json();
        
        const cleanNull = (val: any) => {
          if (typeof val === 'string' && (val.toLowerCase() === 'null' || val.toLowerCase() === '(null)')) {
            return '...';
          }
          return val;
        };

        if (data.extractedTitle && cleanNull(data.extractedTitle) !== '...') {
           setInitiativeTitle(cleanNull(data.extractedTitle));
        }
        
        setTeacher(prev => ({
          ...prev,
          teacherName: cleanNull(data.extractedAuthor) || prev.teacherName,
          schoolName: cleanNull(data.extractedSchool) || prev.schoolName,
          stage: cleanNull(data.extractedStage) || prev.stage,
          subject: cleanNull(data.extractedSubject) || prev.subject,
          role: cleanNull(data.extractedRole) || prev.role,
          birthYear: cleanNull(data.extractedBirthYear) || prev.birthYear,
          phone: cleanNull(data.extractedPhone) || prev.phone || '',
          email: cleanNull(data.extractedEmail) || prev.email || ''
        }));
        
        setExtractionStep(3); // Mức 3: Hoàn thành

        if (file.name.endsWith('.pdf') && file.size <= 3.5 * 1024 * 1024) {
           setInitiativeText(`📁 Đã nhận diện được tập tin PDF: "${file.name}"\n(Đã tự động điền thông tin Giáo viên. Đang chuẩn bị chuyển tab...)`);
        }
        
        // Auto navigate to the next tab based on user intent
        setTimeout(() => {
          setIsExtractingInfo(false);
          setMainTab('plagiarism');
        }, 1500);
      } else {
        if (response.status === 413) {
           setErrorMsg(`Trích xuất tự động thất bại: File tải lên quá lớn (Vượt giới hạn 4.5MB của Vercel). Vui lòng nén file PDF hoặc chuyển sang chế độ Copy/Paste.`);
        } else {
           let errorData;
           try { errorData = await response.json(); } catch (e) { errorData = { error: `Lỗi máy chủ (${response.status})` }; }
           const errDetails = getFriendlyErrorMessage(errorData?.details || errorData?.error || '');
           if (typeof errDetails === 'string' && errDetails.includes('Quota')) {
              setErrorMsg(`Trích xuất tự động thất bại: Key đã hết lượt gọi (Quota Exceeded). Vui lòng điền thêm API Key dự phòng 😢`);
           } else {
              setErrorMsg(`Trích xuất thông tin tự động thất bại: ${errDetails}`);
           }
        }
        setIsExtractingInfo(false);
      }
    } catch (err: any) {
      console.warn("Failed to extract info auto:", err);
      setErrorMsg(`Trích xuất tự động thất bại: ${getFriendlyErrorMessage(err.message || 'Lỗi mạng')}`);
      setIsExtractingInfo(false);
    }
  };

  const handleDragOver = (e: DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      processUploadedFile(file);
    }
  };

  const getClassificationBadge = (classification: 'Xuất sắc' | 'Tốt' | 'Khá' | 'Không đạt') => {
    switch (classification) {
      case 'Xuất sắc':
        return (
          <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold uppercase tracking-wider bg-[#5a5a40] text-white">
            <Award className="w-4 h-4 text-[#eaeada]" /> Xuất sắc
          </span>
        );
      case 'Tốt':
        return (
          <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold uppercase tracking-wider bg-[#70705a] text-white">
            <CheckCircle2 className="w-4 h-4 text-[#eaeada]" /> Tốt
          </span>
        );
      case 'Khá':
        return (
          <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold uppercase tracking-wider bg-[#8b8b68] text-white">
            <CheckCircle2 className="w-4 h-4 text-[#eaeada]" /> Khá
          </span>
        );
      default:
        return (
          <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold uppercase tracking-wider bg-[#7a7a5c] text-white">
            <XCircle className="w-4 h-4 text-[#eaeada]" /> Không đạt
          </span>
        );
    }
  };

  if (loadingAuth) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <Loader2 className="w-8 h-8 animate-spin text-blue-600" />
      </div>
    );
  }

  if (!user) {
    return <LoginScreen onLogin={() => {}} />;
  }

  return (
    <div id="app-workspace" className="min-h-screen bg-natural-bg text-natural-text font-sans flex flex-col selection:bg-natural-accent border-[12px] border-natural-border md:border-[16px] print-force-static">
      
      {/* 🇻🇳 Red-Gold Prestige Vietnamese Administrative Header changed to elegant Natural Tones */}
      <header id="header-bar" className="bg-natural-primary text-white shadow-md border-b-4 border-natural-secondary no-print">
        <div className="max-w-[1600px] mx-auto px-6 py-5 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center border border-white/30 shrink-0">
              <GraduationCap className="w-6 h-6 text-white" />
            </div>
            <div className="text-center md:text-left">
              <span className="text-xs uppercase tracking-wider font-bold text-natural-border block mb-0.5">
                ỦY BAN NHÂN DÂN XÃ HÀM YÊN — HỘI ĐỒNG SÁNG KIẾN
              </span>
              <h1 className="text-xl md:text-2xl font-bold tracking-tight">
                Hệ thống Trợ lý Thẩm định Báo cáo Sáng kiến kinh nghiệm
              </h1>
            </div>
          </div>
          
          <div className="flex flex-col md:items-end text-center md:text-right gap-1">
            <div className="flex items-center gap-3">
              <div className="bg-white/10 backdrop-blur-sm border border-white/20 px-3.5 py-1.5 rounded-lg text-xs font-mono flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse inline-block"></span>
                Thẩm định online: <span className="text-natural-border font-bold">Quyết định 270/QĐ-HĐSK</span>
              </div>
              <button
                onClick={() => signOut(auth)}
                className="bg-red-500/20 hover:bg-red-500/30 text-white border border-red-500/50 px-3.5 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-colors"
                title="Đăng xuất"
              >
                <LogOut className="w-3.5 h-3.5" />
                <span className="hidden sm:inline">Đăng xuất</span>
              </button>
            </div>
            <p className="text-[11px] text-white/80">
              Cơ quan Thường trực: Sở Khoa học và Công nghệ tỉnh Tuyên Quang
            </p>
          </div>
        </div>
      </header>

      {/* 🧭 Simplified Unified Workflow Status Strip */}
      <div className="bg-natural-primary/5 border-b border-natural-border px-6 py-3 shrink-0 flex justify-between items-center no-print">
        <div className="text-xs text-natural-primary font-bold flex items-center gap-2">
          <Sparkles className="w-4 h-4 text-natural-secondary animate-pulse" /> QUY TRÌNH KIỂM ĐỊNH TỰ ĐỘNG KHÔNG THAO TÁC PHỨC TẠP
        </div>
        <div className="text-[11px] text-natural-muted font-semibold flex items-center gap-1.5 bg-white/60 px-3 py-1 rounded-lg border border-natural-border/60">
          Chế độ: <strong>Thẩm định Tổng thể Quyết định 270</strong>
        </div>
      </div>

      {/* Main Container */}
      <main className="flex-1 max-w-[1600px] w-full mx-auto px-4 lg:px-8 py-6 md:py-8 flex flex-col gap-6 md:gap-8 no-print">

        <div className="flex border-b border-natural-border gap-2 overflow-x-auto custom-scrollbar">
           <button 
             onClick={() => setMainTab('guide')}
             className={`px-4 py-3 font-bold uppercase tracking-wider text-[13px] border-b-2 transition whitespace-nowrap flex items-center gap-2 ${mainTab === 'guide' ? 'border-natural-primary text-natural-primary bg-natural-primary/5' : 'border-transparent text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
           >
             <HelpCircle className="w-4 h-4" /> Hướng dẫn
           </button>
           <button 
             onClick={() => setMainTab('info')}
             className={`px-4 py-3 font-bold uppercase tracking-wider text-[13px] border-b-2 transition whitespace-nowrap flex items-center gap-2 ${mainTab === 'info' ? 'border-natural-primary text-natural-primary bg-natural-primary/5' : 'border-transparent text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
           >
             <FileText className="w-4 h-4" /> 1. Hồ sơ Sáng kiến
           </button>
           <button 
             onClick={() => setMainTab('plagiarism')}
             className={`px-4 py-3 font-bold uppercase tracking-wider text-[13px] border-b-2 transition whitespace-nowrap flex items-center gap-2 ${mainTab === 'plagiarism' ? 'border-natural-primary text-natural-primary bg-natural-primary/5' : 'border-transparent text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
           >
             <Search className="w-4 h-4" /> 2. Đạo văn, Chính tả & Sử dụng AI {plagResult ? `(${plagResult.totalDuplicatePercent}%)` : ''}
           </button>
           <button 
             onClick={() => setMainTab('appraisal')}
             className={`px-4 py-3 font-bold uppercase tracking-wider text-[13px] border-b-2 transition whitespace-nowrap flex items-center gap-2 ${mainTab === 'appraisal' ? 'border-natural-primary text-natural-primary bg-natural-primary/5' : 'border-transparent text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
           >
             <Award className="w-4 h-4" /> 3. Thẩm định & Xuất phiếu {currentResult ? `(${currentResult.totalScore}đ)` : ''}
           </button>
           <button 
             onClick={() => setMainTab('settings')}
             className={`px-4 py-3 font-bold uppercase tracking-wider text-[13px] border-b-2 transition whitespace-nowrap flex items-center gap-2 ml-auto ${mainTab === 'settings' ? 'border-natural-primary text-natural-primary bg-natural-primary/5' : 'border-transparent text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
           >
             <Settings className="w-4 h-4" /> Cài đặt API & Hệ thống
           </button>
        </div>
        
        {/* Global Error warning bar */}
        {errorMsg && (
          <div className="bg-red-50 border-l-4 border-red-500 p-4 rounded-xl text-sm text-red-800 flex items-start gap-3 shadow-sm mx-auto w-full max-w-4xl cursor-pointer" onClick={() => setErrorMsg(null)}>
            <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
            <div>{errorMsg}</div>
          </div>
        )}

        {mainTab === 'guide' && (
        <section id="guide-section" className="w-full max-w-4xl mx-auto flex flex-col gap-6">
          <div className="bg-white rounded-2xl border border-natural-border shadow-sm overflow-hidden">
            <div className="bg-natural-accent border-b border-natural-border px-5 py-3.5 flex items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <HelpCircle className="w-5 h-5 text-natural-primary" />
                <h2 className="font-bold text-natural-primary text-sm md:text-base">Hướng dẫn sử dụng Hệ thống Thẩm định SKKN</h2>
              </div>
              <button
                onClick={async () => {
                  const htmlContent = `
                    <!DOCTYPE html>
                    <html>
                    <head>
                      <meta charset='utf-8'>
                      <title>Hướng dẫn sử dụng Hệ thống Thẩm định SKKN</title>
                      <style>
                        body { font-family: 'Times New Roman', Times, serif; font-size: 14pt; line-height: 1.5; }
                        h2 { font-size: 18pt; text-align: center; color: #000; font-weight: bold; margin-bottom: 20pt; }
                        h3 { font-size: 14pt; color: #000; font-weight: bold; margin-top: 16pt; margin-bottom: 8pt; }
                        p.step-desc { margin-left: 20pt; margin-top: 0; margin-bottom: 4pt; }
                        .highlight { color: red; font-weight: bold; }
                        .step { font-weight: bold; }
                      </style>
                    </head>
                    <body>
                      <h2>HƯỚNG DẪN SỬ DỤNG HỆ THỐNG THẨM ĐỊNH SKKN</h2>
                      
                      <h3>Bước 1: Cấu hình API Key (Bắt buộc)</h3>
                      <p class="step-desc"><span class="step">B1:</span> Chuyển sang thẻ <span class="highlight">"Cài đặt API & Hệ thống"</span> ở thanh menu.</p>
                      <p class="step-desc"><span class="step">B2:</span> <span class="highlight">Nhập Key API</span> vào ô tương ứng. Nếu chưa nhập, AI sẽ <span class="highlight">không thể hoạt động</span> (báo lỗi).</p>
                      <p class="step-desc"><span class="step">B3:</span> Nhấn <span class="highlight">"Lưu cấu hình"</span>.</p>

                      <h3>Bước 2: Tải lên Sáng kiến & Nhập thông tin giám khảo</h3>
                      <p class="step-desc"><span class="step">B1:</span> Tại thẻ <strong>"1. Hồ sơ Sáng kiến"</strong>, <span class="highlight">tải file Sáng kiến lên</span> (hỗ trợ .PDF, .TXT) hoặc dán nội dung.</p>
                      <p class="step-desc"><span class="step">B2:</span> <span class="highlight">Chờ AI tự động trích xuất</span> và điền đầy đủ các thông tin tác giả.</p>
                      <p class="step-desc"><span class="step">B3:</span> <span class="highlight">Nhập Tên Giám khảo / Người nhận xét</span> ở bên cột thông tin.</p>
                      <p class="step-desc"><span class="step">B4:</span> Nhấn <span class="highlight">"Tiếp tục: 2. Đạo văn, Chính tả & Sử dụng AI"</span> để chuyển sang bước 2.</p>

                      <h3>Bước 3: Kiểm tra Đạo văn & Lỗi kỹ thuật</h3>
                      <p class="step-desc"><span class="step">B1:</span> Chuyển sang thẻ <span class="highlight">"2. Đạo văn, Chính tả & Sử dụng AI"</span>.</p>
                      <p class="step-desc"><span class="step">B2:</span> Nhấn nút <span class="highlight">"Bắt đầu quét"</span> để hệ thống phân tích.</p>
                      <p class="step-desc"><span class="step">B3:</span> <span class="highlight">Xem kết quả đánh giá:</span> Tỷ lệ đạo văn, Dấu hiệu AI sinh, và Lỗi chính tả.</p>
                      <p class="step-desc"><span class="step">B4:</span> Nhấn <span class="highlight">"Tiếp tục: 3. Thẩm định (Chấm điểm)"</span> để sang bước 3.</p>

                      <h3>Bước 4: Thẩm định (Chấm điểm) & Sinh nhận xét</h3>
                      <p class="step-desc"><span class="step">B1:</span> Chuyển sang thẻ <span class="highlight">"3. Thẩm định & Xuất phiếu"</span>.</p>
                      <p class="step-desc"><span class="step">B2:</span> Nhấn nút <span class="highlight">"Bắt đầu Thẩm định (Chấm điểm)"</span>.</p>
                      <p class="step-desc"><span class="step">B3:</span> <span class="highlight">Chờ hệ thống AI đóng vai trò Hội đồng</span> đọc toàn bộ nội dung, chấm điểm chi tiết và viết nhận xét ưu/nhược điểm.</p>

                      <h3>Bước 5: Chỉnh sửa & Xuất phiếu kết quả</h3>
                      <p class="step-desc"><span class="step">B1:</span> <span class="highlight">Kiểm tra và tự do chỉnh sửa</span> mọi thông tin trên phiếu (nhận xét, điểm số, thông tin).</p>
                      <p class="step-desc"><span class="step">B2:</span> Kéo xuống cuối trang, chọn <span class="highlight">"In Phiếu Thẩm định (PDF)"</span> để tải bản báo cáo về.</p>
                      <p class="step-desc"><span class="step">B3:</span> Hoặc chọn <span class="highlight">"Tải Danh sách Excel"</span> / <span class="highlight">"Lưu file JSON"</span> để lưu trữ kết quả đánh giá.</p>
                    </body>
                    </html>
                  `;
                  try {
                    const docxBlob = await asBlob(htmlContent);
                    saveAs(docxBlob as Blob, 'Huong_dan_su_dung_tham_dinh_SKKN.docx');
                  } catch (err) {
                    console.error('Lỗi khi tải file docx:', err);
                    // Fallback to older doc export if something fails
                    const blob = new Blob(['\ufeff' + htmlContent], { type: 'application/msword' });
                    saveAs(blob, 'Huong_dan_su_dung_tham_dinh_SKKN.docx');
                  }
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-white text-natural-primary border border-natural-primary/30 rounded-lg shadow-sm hover:bg-natural-primary hover:text-white transition-colors"
              >
                <Download className="w-4 h-4" /> Tải về (.docx)
              </button>
            </div>
            <div className="p-6 flex flex-col gap-6 text-sm text-natural-text leading-relaxed">
              <div className="flex flex-col gap-2">
                <h3 className="font-bold text-natural-primary text-base flex items-center gap-2"><div className="w-6 h-6 rounded-full bg-natural-primary text-white flex items-center justify-center text-xs">1</div> Bước 1: Cấu hình API Key (Bắt buộc)</h3>
                <ul className="list-none pl-8 flex flex-col gap-1">
                  <li>B1: Chuyển sang thẻ <strong className="text-red-600">"Cài đặt API & Hệ thống"</strong> ở thanh menu.</li>
                  <li>B2: <strong className="text-red-600">Nhập Key API</strong> vào ô tương ứng. Nếu chưa nhập, AI sẽ <strong className="text-red-600">không thể hoạt động</strong> (báo lỗi).</li>
                  <li>B3: Nhấn <strong className="text-red-600">"Lưu cấu hình"</strong>.</li>
                </ul>
              </div>

              <div className="flex flex-col gap-2">
                <h3 className="font-bold text-natural-primary text-base flex items-center gap-2"><div className="w-6 h-6 rounded-full bg-natural-primary text-white flex items-center justify-center text-xs">2</div> Bước 2: Tải lên Sáng kiến & Nhập thông tin giám khảo</h3>
                <ul className="list-none pl-8 flex flex-col gap-1">
                  <li>B1: Tại thẻ <strong>"1. Hồ sơ Sáng kiến"</strong>, <strong className="text-red-600">tải file Sáng kiến lên</strong> (hỗ trợ .PDF, .TXT) hoặc dán nội dung.</li>
                  <li>B2: <strong className="text-red-600">Chờ AI tự động trích xuất</strong> và điền đầy đủ các thông tin tác giả.</li>
                  <li>B3: <strong className="text-red-600">Nhập Tên Giám khảo / Người nhận xét</strong> ở bên cột thông tin.</li>
                  <li>B4: Nhấn <strong className="text-red-600">"Tiếp tục: 2. Đạo văn, Chính tả & Sử dụng AI"</strong> để chuyển sang bước 2.</li>
                </ul>
              </div>

              <div className="flex flex-col gap-2">
                <h3 className="font-bold text-natural-primary text-base flex items-center gap-2"><div className="w-6 h-6 rounded-full bg-natural-primary text-white flex items-center justify-center text-xs">3</div> Bước 3: Kiểm tra Đạo văn & Lỗi kỹ thuật</h3>
                <ul className="list-none pl-8 flex flex-col gap-1">
                  <li>B1: Chuyển sang thẻ <strong className="text-red-600">"2. Đạo văn, Chính tả & Sử dụng AI"</strong>.</li>
                  <li>B2: Nhấn nút <strong className="text-red-600">"Bắt đầu quét"</strong> để hệ thống phân tích.</li>
                  <li>B3: <strong className="text-red-600">Xem kết quả đánh giá:</strong> Tỷ lệ đạo văn, Dấu hiệu AI sinh, và Lỗi chính tả.</li>
                  <li>B4: Nhấn <strong className="text-red-600">"Tiếp tục: 3. Thẩm định (Chấm điểm)"</strong> để sang bước 3.</li>
                </ul>
              </div>

              <div className="flex flex-col gap-2">
                <h3 className="font-bold text-natural-primary text-base flex items-center gap-2"><div className="w-6 h-6 rounded-full bg-natural-primary text-white flex items-center justify-center text-xs">4</div> Bước 4: Thẩm định (Chấm điểm) & Sinh nhận xét</h3>
                <ul className="list-none pl-8 flex flex-col gap-1">
                  <li>B1: Chuyển sang thẻ <strong className="text-red-600">"3. Thẩm định & Xuất phiếu"</strong>.</li>
                  <li>B2: Nhấn nút <strong className="text-red-600">"Bắt đầu Thẩm định (Chấm điểm)"</strong>.</li>
                  <li>B3: <strong className="text-red-600">Chờ hệ thống AI đóng vai trò Hội đồng</strong> đọc toàn bộ nội dung, chấm điểm chi tiết và viết nhận xét ưu/nhược điểm.</li>
                </ul>
              </div>

              <div className="flex flex-col gap-2">
                <h3 className="font-bold text-natural-primary text-base flex items-center gap-2"><div className="w-6 h-6 rounded-full bg-natural-primary text-white flex items-center justify-center text-xs">5</div> Bước 5: Chỉnh sửa & Xuất phiếu kết quả</h3>
                <ul className="list-none pl-8 flex flex-col gap-1">
                  <li>B1: <strong className="text-red-600">Kiểm tra và tự do chỉnh sửa</strong> mọi thông tin trên phiếu (nhận xét, điểm số, thông tin).</li>
                  <li>B2: Kéo xuống cuối trang, chọn <strong className="text-red-600">"In Phiếu Thẩm định (PDF)"</strong> để tải bản báo cáo về.</li>
                  <li>B3: Hoặc chọn <strong className="text-red-600">"Tải Danh sách Excel"</strong> / <strong className="text-red-600">"Lưu file JSON"</strong> để lưu trữ kết quả đánh giá.</li>
                </ul>
              </div>
            </div>
          </div>
        </section>
        )}

        {mainTab === 'info' && (
        <section id="input-section" className="w-full max-w-full mx-auto flex flex-col gap-6">
          {/* 📋 LẮP ĐẶT THÔNG TIN VÀ NỘI DUNG SÁNG KIẾN */}

          {/* Form Thông tin Tác giả & Sáng kiến */}
          <div className="bg-white rounded-2xl border border-natural-border shadow-sm overflow-hidden">
            <div className="bg-natural-accent border-b border-natural-border px-5 py-3.5 flex items-center justify-between">
              <h2 className="font-bold text-natural-primary text-sm md:text-base flex items-center gap-2">
                <FileText className="w-5 h-5 text-natural-primary" /> Nhập Hồ sơ & Nội dung Sáng kiến kinh nghiệm
              </h2>
              <span className="text-xs text-natural-muted font-mono bg-white/60 px-2 py-0.5 rounded border border-natural-border">2025 - 2026</span>
            </div>
            
            <div className="p-5 md:p-6 flex flex-col gap-5">
              
              {/* Main Content Area */}
              <div>
                <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-2 gap-2">
                  <label className="text-xs font-bold text-natural-text uppercase flex items-center">
                    <span>Nội dung Sáng kiến hoặc Báo cáo</span>
                    <span className="text-[11px] text-natural-primary italic uppercase font-bold ml-2">PDF / TXT</span>
                  </label>
                  <div className="flex items-center gap-2">
                    <label className="text-[10px] font-bold text-natural-muted uppercase tracking-wider">Mô hình AI trích xuất:</label>
                    <select
                      value={selectedModelExtract}
                      onChange={(e) => setSelectedModelExtract(e.target.value)}
                      className="text-xs font-semibold bg-white border border-amber-300 text-amber-900 rounded-lg px-2 py-1 outline-none w-[160px] shadow-sm"
                    >
                      <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro Preview (Mạnh mẽ nhất)</option>
                      <option value="gemini-pro">Gemini Pro Latest</option>
                      <option value="gemini-3.5-flash">Gemini 3.5 Flash (Cân bằng, Rất ổn định)</option>
                      <option value="gemini-3.0-flash-preview">Gemini 3 Flash Preview (Thử nghiệm)</option>
                      <option value="gemini-flash">Gemini Flash Latest</option>
                      <option value="gemini-2.5-flash">Gemini 2.5 Flash (Thế hệ trước)</option>
                      <option value="gemini-3.1-flash-lite">Gemini 3.1 Flash Lite (Nhanh nhất, Khuyên dùng)</option>
                    </select>
                  </div>
                </div>
                
                {/* File Upload Area */}
                <div 
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  className={`relative border-2 border-dashed rounded-xl p-4 flex flex-col items-center justify-center text-center transition-all ${
                    isDragging 
                      ? 'border-natural-primary bg-natural-primary/5' 
                      : 'border-natural-border/80 hover:border-natural-primary/50 bg-[#faf9f5]'
                  }`}
                  style={{ minHeight: '120px' }}
                >
                  <input
                    type="file"
                    accept=".pdf,.txt"
                    onChange={handleFileChange}
                    className="absolute inset-0 w-full h-full cursor-pointer opacity-0"
                  />
                  <div className="w-10 h-10 rounded-full bg-white flex items-center justify-center shadow-sm border border-natural-border/40 text-natural-muted mb-2">
                    <FileUp className="w-5 h-5 text-natural-primary" />
                  </div>
                  
                  {plagFileName ? (
                    <div className="space-y-1">
                      <p className="text-xs font-bold text-natural-primary">{plagFileName}</p>
                      <p className="text-[10px] text-emerald-600 font-bold uppercase tracking-wider bg-emerald-50 px-2 py-0.5 rounded inline-block">Đã đính kèm tệp PDF/TXT</p>
                    </div>
                  ) : (
                    <div className="space-y-1">
                      <p className="text-xs font-bold text-natural-text">Kéo thả hoặc Tải lên (.PDF / .TXT)</p>
                      <p className="text-[10px] text-natural-muted">Vui lòng tải lên tài liệu để phân tích và đánh giá.</p>
                    </div>
                  )}
                </div>

                {/* Extraction Timeline */}
                {isExtractingInfo && (
                  <div className="mt-4 p-5 border border-natural-border rounded-xl bg-white shadow-sm font-sans relative overflow-hidden">
                    <div className="absolute top-0 left-0 h-1 bg-[#8b8b68] transition-all duration-500" style={{ width: `${(extractionStep / 3) * 100}%` }}></div>
                    
                    <p className="text-xs font-bold text-natural-text uppercase mb-4 flex items-center gap-2">
                      <Loader2 className="w-4 h-4 text-natural-primary animate-spin" />
                      Đang phân tích tài liệu tải lên...
                    </p>
                    
                    <div className="flex flex-col gap-3 relative ml-1">
                      <div className="absolute left-2.5 top-2.5 bottom-2.5 w-px bg-natural-border"></div>
                      
                      {/* Step 1 */}
                      <div className={`flex items-center gap-3 relative z-10 transition-all duration-300 ${extractionStep >= 1 ? 'opacity-100 translate-y-0' : 'opacity-40 translate-y-1'}`}>
                        <div className={`w-5 h-5 rounded-full flex items-center justify-center shrink-0 border ${extractionStep >= 1 ? 'border-[#8b8b68] bg-[#f8f7f3]' : 'border-natural-border bg-white'}`}>
                          {extractionStep > 1 ? <Check className="w-3 h-3 text-[#8b8b68]" /> : (extractionStep === 1 ? <Loader2 className="w-3 h-3 text-[#8b8b68] animate-spin" /> : <div className="w-1.5 h-1.5 rounded-full bg-natural-border"></div>)}
                        </div>
                        <div className="flex flex-col">
                          <span className={`text-sm ${extractionStep >= 1 ? 'font-semibold text-natural-primary' : 'text-natural-muted'}`}>Bóc tách nội dung tập tin PDF/TXT</span>
                        </div>
                      </div>
                      
                      {/* Step 2 */}
                      <div className={`flex items-center gap-3 relative z-10 transition-all duration-300 ${extractionStep >= 2 ? 'opacity-100 translate-y-0' : 'opacity-40 translate-y-1'}`}>
                        <div className={`w-5 h-5 rounded-full flex items-center justify-center shrink-0 border ${extractionStep >= 2 ? 'border-[#8b8b68] bg-[#f8f7f3]' : 'border-natural-border bg-white'}`}>
                          {extractionStep > 2 ? <Check className="w-3 h-3 text-[#8b8b68]" /> : (extractionStep === 2 ? <Loader2 className="w-3 h-3 text-[#8b8b68] animate-spin" /> : <div className="w-1.5 h-1.5 rounded-full bg-natural-border"></div>)}
                        </div>
                        <div className="flex flex-col">
                          <span className={`text-sm ${extractionStep >= 2 ? 'font-semibold text-natural-primary' : 'text-natural-muted'}`}>AI Phân tích tác giả, chức vụ, trường lớp...</span>
                        </div>
                      </div>
                      
                      {/* Step 3 */}
                      <div className={`flex items-center gap-3 relative z-10 transition-all duration-300 ${extractionStep >= 3 ? 'opacity-100 translate-y-0' : 'opacity-40 translate-y-1'}`}>
                        <div className={`w-5 h-5 rounded-full flex items-center justify-center shrink-0 border ${extractionStep >= 3 ? 'border-emerald-600 bg-emerald-50' : 'border-natural-border bg-white'}`}>
                          {extractionStep >= 3 ? <Check className="w-3 h-3 text-emerald-600" /> : <div className="w-1.5 h-1.5 rounded-full bg-natural-border"></div>}
                        </div>
                        <div className="flex flex-col">
                          <span className={`text-sm ${extractionStep >= 3 ? 'font-semibold text-emerald-700' : 'text-natural-muted'}`}>Hoàn thiện dữ liệu! Tự động chuyển tab sau 2s...</span>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

              {/* Submit Buttons */}
              <div className="flex flex-col gap-3 mt-4 mb-2">
                <div className="flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setMainTab('plagiarism')}
                    className="flex-1 min-w-[200px] border border-[#8b8b68] bg-[#fdfdfa] hover:bg-[#8b8b68] hover:text-white text-[#8b8b68] text-[12px] font-bold uppercase tracking-wider py-3 rounded-xl transition flex justify-center items-center gap-1.5 cursor-pointer shadow-sm"
                  >
                     Tiếp tục: 2. Đạo văn, Chính tả & Sử dụng AI <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
                
                <div className="flex justify-end mt-2">
                  <button
                    type="button"
                    onClick={() => {
                      setTeacher({ teacherName: '', birthYear: '', role: '', schoolName: '', stage: 'Tiểu học', subject: '', phone: '', email: '' });
                      setInitiativeTitle('');
                      setInitiativeText('');
                      setErrorMsg(null);
                      setPlagFileBase64('');
                      setPlagFileName('');
                      setPlagResult(null);
                      setCurrentResult(null);
                    }}
                    className="px-4 py-2 border border-natural-border text-natural-muted hover:text-natural-primary hover:bg-natural-accent text-xs font-bold uppercase tracking-wider rounded-lg transition cursor-pointer"
                  >
                    Xóa trắng làm lại
                  </button>
                </div>
              </div>

              </div>

              {/* Tác giả Grid */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5 flex items-center gap-1.5">
                    <User className="w-3.5 h-3.5 text-natural-muted" /> Họ và tên Giáo viên
                  </label>
                  <input
                    type="text"
                    value={teacher.teacherName}
                    onChange={(e) => setTeacher({...teacher, teacherName: e.target.value})}
                    placeholder="Ví dụ: Nguyễn Thị Mai"
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5 flex items-center gap-1.5">
                    <Calendar className="w-3.5 h-3.5 text-natural-muted" /> Năm sinh
                  </label>
                  <input
                    type="text"
                    value={teacher.birthYear}
                    onChange={(e) => setTeacher({...teacher, birthYear: e.target.value})}
                    placeholder="Ví dụ: 1991"
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5 flex items-center gap-1.5">
                    <Building className="w-3.5 h-3.5 text-natural-muted" /> Đơn vị công tác
                  </label>
                  <input
                    type="text"
                    value={teacher.schoolName}
                    onChange={(e) => setTeacher({...teacher, schoolName: e.target.value})}
                    placeholder="Ví dụ: Trường Tiểu học Hàm Yên"
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                      Cấp học
                    </label>
                    <select
                      value={teacher.stage}
                      onChange={(e: any) => setTeacher({...teacher, stage: e.target.value})}
                      className="w-full text-sm border border-natural-border rounded-xl px-2 py-2 focus:border-natural-primary focus:outline-none bg-[#fffefc]"
                    >
                      <option value="Mầm non">Mầm non</option>
                      <option value="Tiểu học">Tiểu học</option>
                      <option value="THCS">THCS</option>
                      <option value="Khác">Khác</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                      Bộ môn / Lĩnh vực
                    </label>
                    <input
                      type="text"
                      value={teacher.subject}
                      onChange={(e) => setTeacher({...teacher, subject: e.target.value})}
                      placeholder="Ví dụ: Tiếng Việt"
                      className="w-full text-sm border border-natural-border rounded-xl px-2.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5 flex items-center gap-1.5">
                    <Phone className="w-3.5 h-3.5 text-natural-muted" /> Số điện thoại tác giả
                  </label>
                  <input
                    type="text"
                    value={teacher.phone || ''}
                    onChange={(e) => setTeacher({...teacher, phone: e.target.value})}
                    placeholder="Ví dụ: 0912345678"
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5 flex items-center gap-1.5">
                    <Mail className="w-3.5 h-3.5 text-natural-muted" /> Email liên hệ
                  </label>
                  <input
                    type="email"
                    value={teacher.email || ''}
                    onChange={(e) => setTeacher({...teacher, email: e.target.value})}
                    placeholder="Ví dụ: nguyenmai@gmail.com"
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>
              </div>

              {/* Title of Initiative */}
              <div>
                <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                  Tên Sáng kiến kinh nghiệm (Chính xác theo trang bìa Báo cáo)
                </label>
                <input
                  type="text"
                  value={initiativeTitle}
                  onChange={(e) => setInitiativeTitle(e.target.value)}
                  placeholder="Ví dụ: Một số biện pháp nâng cao hiệu quả giáo dục..."
                  className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2.5 font-medium tracking-tight text-natural-text focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                />
              </div>

              {/* Basic Meta Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                    Chức danh chuyên môn
                  </label>
                  <input
                    type="text"
                    value={teacher.role}
                    onChange={(e) => setTeacher({...teacher, role: e.target.value})}
                    placeholder="Ví dụ: Giáo viên chủ nhiệm lớp 5A"
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                    Tên Giám khảo / Người nhận xét
                  </label>
                  <input
                    type="text"
                    value={reviewerName}
                    onChange={(e) => {
                      setReviewerName(e.target.value);
                      syncToFirestore({ reviewerName: e.target.value });
                    }}
                    placeholder="Tên Giám khảo..."
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                    Thời điểm áp dụng đầu tiên
                  </label>
                  <input
                    type="date"
                    value={appliedDate}
                    onChange={(e) => setAppliedDate(e.target.value)}
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc]"
                  />
                </div>
              </div>

              {/* Hội đồng Thẩm định (Tùy chọn) */}
              <div className="border-t border-natural-border/60 pt-5 mt-4">
                <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                  <div className="flex items-center gap-2">
                    <div className="h-4 w-1 bg-natural-primary rounded-full"></div>
                    <h3 className="text-xs font-bold text-natural-primary uppercase tracking-wider">Thành viên Tổ Thẩm định / Hội đồng Thẩm định (Tùy chọn)</h3>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const newId = 'member_' + Date.now();
                      const updated = [...councilMembers, { id: newId, name: '', unit: '', role: '' }];
                      setCouncilMembers(updated);
                      syncToFirestore({ councilMembers: updated });
                    }}
                    className="flex items-center gap-1.5 text-[11px] font-bold text-natural-primary hover:text-natural-secondary bg-[#faf9f5] border border-natural-border px-3 py-1.5 rounded-xl transition cursor-pointer hover:bg-natural-accent"
                  >
                    <Plus className="w-3.5 h-3.5" /> Thêm thành viên
                  </button>
                </div>
                
                <div className="mb-4">
                  <label className="block text-xs font-bold text-natural-text uppercase mb-1.5">
                    Hội đồng / Tổ thẩm định nào (Ví dụ: CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN)
                  </label>
                  <input
                    type="text"
                    value={councilName}
                    onChange={(e) => {
                      setCouncilName(e.target.value);
                      syncToFirestore({ councilName: e.target.value });
                    }}
                    placeholder="Nhập tên hội đồng hoặc tổ thẩm định..."
                    className="w-full text-sm border border-natural-border rounded-xl px-3.5 py-2.5 focus:border-natural-primary focus:ring-1 focus:ring-natural-primary focus:outline-none bg-[#fffefc] text-natural-text font-medium"
                  />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {councilMembers.map((m, index) => (
                    <div key={m.id || index} className="bg-[#faf9f5]/50 p-4 rounded-xl border border-natural-border/60 space-y-3 shadow-sm relative">
                      <div className="flex items-center justify-between border-b border-natural-border/60 pb-1.5">
                        <div className="text-[11px] font-bold text-natural-primary uppercase">Thành viên {index + 1}</div>
                        {councilMembers.length > 1 && (
                          <button
                            type="button"
                            onClick={() => {
                              const updated = councilMembers.filter(member => member.id !== m.id);
                              setCouncilMembers(updated);
                              syncToFirestore({ councilMembers: updated });
                            }}
                            className="text-red-500 hover:text-red-700 transition opacity-80 hover:opacity-100 cursor-pointer p-1 rounded hover:bg-red-50"
                            title="Xóa thành viên này"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-natural-text uppercase mb-1">Họ và tên</label>
                        <input
                          type="text"
                          value={m.name}
                          onChange={(e) => {
                            const updated = councilMembers.map(member => member.id === m.id ? { ...member, name: e.target.value } : member);
                            setCouncilMembers(updated);
                            syncToFirestore({ councilMembers: updated });
                          }}
                          placeholder="Ví dụ: Nguyễn Văn A"
                          className="w-full text-xs border border-natural-border rounded-lg px-2.5 py-1.5 focus:border-natural-primary focus:outline-none bg-white text-natural-text"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-natural-text uppercase mb-1">Đơn vị công tác</label>
                        <input
                          type="text"
                          value={m.unit}
                          onChange={(e) => {
                            const updated = councilMembers.map(member => member.id === m.id ? { ...member, unit: e.target.value } : member);
                            setCouncilMembers(updated);
                            syncToFirestore({ councilMembers: updated });
                          }}
                          placeholder="Ví dụ: Trường Tiểu học Phù Lưu"
                          className="w-full text-xs border border-natural-border rounded-lg px-2.5 py-1.5 focus:border-natural-primary focus:outline-none bg-white text-natural-text"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-natural-text uppercase mb-1">Chức vụ</label>
                        <input
                          type="text"
                          value={m.role}
                          onChange={(e) => {
                            const updated = councilMembers.map(member => member.id === m.id ? { ...member, role: e.target.value } : member);
                            setCouncilMembers(updated);
                            syncToFirestore({ councilMembers: updated });
                          }}
                          placeholder="Ví dụ: Tổ trưởng chuyên môn"
                          className="w-full text-xs border border-natural-border rounded-lg px-2.5 py-1.5 focus:border-natural-primary focus:outline-none bg-white text-natural-text"
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* 📖 CẨM NANG HƯỚNG DẪN HỘI ĐỒNG (Trái - Dưới) */}
          <div className="bg-white rounded-2xl border border-natural-border shadow-sm overflow-hidden">
            <div className="border-b border-natural-border flex text-xs md:text-sm bg-natural-accent">
              <button
                onClick={() => setLhsTab('rubric')}
                className={`flex-1 px-4 py-3 font-semibold transition border-r border-natural-border flex items-center justify-center gap-1.5 ${lhsTab === 'rubric' ? 'bg-white text-natural-primary border-t-2 border-t-natural-primary' : 'text-natural-muted hover:text-natural-text bg-natural-accent/50 hover:bg-natural-accent'}`}
              >
                <Award className="w-4 h-4 text-natural-secondary" /> Thang điểm Tuyên Quang
              </button>
              <button
                onClick={() => setLhsTab('legal')}
                className={`flex-1 px-4 py-3 font-semibold transition border-r border-natural-border flex items-center justify-center gap-1.5 ${lhsTab === 'legal' ? 'bg-white text-natural-primary border-t-2 border-t-natural-primary' : 'text-natural-muted hover:text-natural-text bg-natural-accent/50 hover:bg-natural-accent'}`}
              >
                <BookOpen className="w-4 h-4 text-natural-secondary" /> Căn cứ Quyết định 270
              </button>
              <button
                onClick={() => setLhsTab('xml_guide')}
                className={`flex-1 px-4 py-3 font-semibold transition flex items-center justify-center gap-1.5 ${lhsTab === 'xml_guide' ? 'bg-white text-natural-primary border-t-2 border-t-natural-primary' : 'text-natural-muted hover:text-natural-text bg-natural-accent/50 hover:bg-natural-accent'}`}
              >
                <Clock className="w-4 h-4 text-natural-secondary" /> Quy trình nộp Hàm Yên
              </button>
            </div>

            <div className="p-5 text-sm leading-relaxed text-natural-text">
              
              {/* Thang điểm chi tiết */}
              {lhsTab === 'rubric' && (
                <div className="flex flex-col gap-4">
                  <div className="flex items-start gap-2 text-natural-text">
                    <Info className="w-4 h-4 text-natural-secondary shrink-0 mt-0.5" />
                    <p className="text-xs italic font-medium">Bản trích lục chính thức Tiêu chuẩn chấm điểm từ Ủy ban nhân dân và Hội đồng Sáng kiến tỉnh Tuyên Quang:</p>
                  </div>

                  <div className="border border-natural-border rounded-xl overflow-hidden bg-[#fffefb]">
                    <table className="min-w-full text-xs">
                      <thead className="bg-[#fcfbf7] text-natural-text border-b border-natural-border">
                        <tr>
                          <th className="px-3 py-2 text-left font-bold">Tiêu chí 1: Hiệu quả áp dụng (Max 50đ)</th>
                          <th className="px-3 py-2 text-right font-bold w-20">Điểm số</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-natural-border/60 text-natural-text/95">
                        <tr><td className="px-3 py-2">Mức xuất sắc: Hiệu quả kinh tế - xã hội xuất sắc, có minh chứng rõ ràng</td><td className="px-3 py-2 text-right font-bold text-natural-primary">40-50đ</td></tr>
                        <tr><td className="px-3 py-2">Mức tốt: Có hiệu quả áp dụng ở mức tốt, nâng cao năng lực học sinh</td><td className="px-3 py-2 text-right font-bold text-natural-primary">30-39đ</td></tr>
                        <tr><td className="px-3 py-2">Mức khá: Có hiệu quả tuy nhiên chưa rõ rệt hoặc thiếu tính đột phá</td><td className="px-3 py-2 text-right font-bold text-natural-primary">20-29đ</td></tr>
                      </tbody>
                    </table>
                  </div>

                  <div className="border border-natural-border rounded-xl overflow-hidden bg-[#fffefb]">
                    <table className="min-w-full text-xs">
                      <thead className="bg-[#fcfbf7] text-natural-text border-b border-natural-border">
                        <tr>
                          <th className="px-3 py-2 text-left font-bold">Tiêu chí 2: Phạm vi ảnh hưởng cấp Xã (Max 50đ)</th>
                          <th className="px-3 py-2 text-right font-bold w-20">Điểm số</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-natural-border/60 text-natural-text/95">
                        <tr><td className="px-3 py-2">Mức xuất sắc: Có tiềm năng nhân rộng toàn cụm trường/cấp xã</td><td className="px-3 py-2 text-right font-bold text-natural-primary">40-50đ</td></tr>
                        <tr><td className="px-3 py-2">Mức tốt: Đã áp dụng hiệu quả rõ rệt trên quy mô toàn trường</td><td className="px-3 py-2 text-right font-bold text-natural-primary">30-39đ</td></tr>
                        <tr><td className="px-3 py-2">Mức khá: Tác dụng trong phạm vi lớp học/nội bộ tổ chuyên môn</td><td className="px-3 py-2 text-right font-bold text-natural-primary">20-29đ</td></tr>
                      </tbody>
                    </table>
                  </div>

                  <div className="bg-natural-accent p-3.5 rounded-xl border border-natural-border">
                    <h4 className="text-xs font-bold text-natural-primary uppercase mb-1.5">Công thức phân loại kết quả:</h4>
                    <ul className="list-disc pl-4 text-xs space-y-1 text-natural-text/90">
                      <li><strong className="text-natural-primary">Xuất sắc:</strong> Tổng điểm <span className="font-semibold text-natural-text">≥ 80</span>, không tiêu chí nào dưới 40.</li>
                      <li><strong className="text-natural-secondary">Tốt:</strong> Tổng điểm <span className="font-semibold text-natural-text">≥ 60</span>, không tiêu chí nào dưới 30.</li>
                      <li><strong className="text-natural-muted">Khá:</strong> Tổng điểm <span className="font-semibold text-natural-text">≥ 40</span>, không tiêu chí nào dưới 20.</li>
                    </ul>
                  </div>
                </div>
              )}

              {/* Căn cứ Quyết định 270 */}
              {lhsTab === 'legal' && (
                <div className="space-y-3 text-xs text-natural-text/90">
                  <div className="bg-natural-accent p-3 rounded-xl border border-natural-border text-natural-text font-medium">
                    Quyết định số 270/QĐ-HĐSK ngày 22 tháng 12 năm 2025 của Hội đồng Sáng kiến tỉnh Tuyên Quang. ký bởi Phó Chủ tịch UBND tỉnh Vương Ngoc Hà.
                  </div>
                  <p><strong>Điều 1. Phạm vi và đối tượng:</strong> Áp dụng đối với các cơ quan, tổ chức, cá nhân tham gia vào hoạt động sáng kiến giáo dục công nhận để xét tặng danh hiệu thi đua lớp chiến sĩ thi đua cấp cơ sở hoặc cấp tỉnh.</p>
                  <p><strong>Điều 7. Tiêu chuẩn cơ bản về tính mới:</strong></p>
                  <ul className="list-disc pl-4 space-y-1">
                    <li>Không trùng lặp với nội dung của bất cứ giải pháp đã công bố hoặc nộp trước đó.</li>
                    <li>Chưa bị bộc lộ công khai đến mức có thể sao chép ngay khi chưa xin phê duyệt đề án.</li>
                    <li>Sáng kiến giáo dục phải có cấu trúc giải pháp rõ rệt không lạm dụng lý thuyết bài học khô khan.</li>
                  </ul>
                  <p><strong>Điều 8. Minh chứng hiệu lực:</strong> Hồ sơ bắt buộc phải đính kèm số liệu thể nghiệm cụ thể trong tối thiểu một học kỳ hoặc năm học học tập để chứng minh năng lực tiến bộ của học sinh.</p>
                </div>
              )}

              {/* Quy trình nộp hồ sơ */}
              {lhsTab === 'xml_guide' && (
                <div className="space-y-3 text-xs text-natural-text/90">
                  <p className="font-bold text-natural-text">Căn cứ Công văn số 777/UBND-VHXH ngày 25/5/2026 của UBND xã Hàm Yên:</p>
                  <p>Mỗi giáo viên khi hoàn thiện và nộp báo cáo lên phải đóng tập hồ sơ nộp về Phòng Văn hóa - Xã hội xã chậm nhất đến hết ngày <strong className="text-natural-primary font-bold">10/6/2026</strong>.</p>
                  
                  <h4 className="font-bold text-natural-primary">Danh mục hồ sơ quy định:</h4>
                  <ol className="list-decimal pl-4 space-y-1">
                    <li>Đơn yêu cầu công nhận sáng kiến (Theo Mẫu số 4 kèm theo Quy chế).</li>
                    <li>Báo cáo hiệu quả áp dụng, phạm vi ảnh hưởng (Mẫu số 5).</li>
                    <li>Tóm tắt ngắn gọn sáng kiến (Mẫu số 6).</li>
                    <li>Giấy xác nhận hiệu quả áp dụng thực tế từ Hiệu trưởng nhà trường (Mẫu số 7).</li>
                  </ol>
                  <p className="text-natural-muted italic">Hội đồng sáng kiến xã Hàm Yên chịu trách nhiệm tổng hợp, chấm chéo và gửi lên Hội đồng cấp Tỉnh xem xét công nhận.</p>
                </div>
              )}

            </div>
          </div>

        </section>
        )}

        {/* 📊 BẢNG KẾT QUẢ THẨM ĐỊNH CHI TIẾT */}
        {(mainTab === 'plagiarism' || mainTab === 'appraisal') && (
        <section id="results-section" className="w-full mx-auto flex flex-col gap-6">
          
          {/* Lịch sử lưu trữ cục bộ */}
          {history.length > 0 && (
            <div className="bg-white p-4 rounded-2xl border border-natural-border shadow-sm flex flex-col gap-2">
              <h3 className="text-xs font-bold text-natural-muted uppercase tracking-wider flex items-center justify-between">
                <span className="flex items-center gap-1.5"><History className="w-4 h-4 text-natural-muted" /> Danh sách sáng kiến đã thẩm định ({history.length})</span>
                <div className="flex items-center gap-2 flex-wrap justify-end">
                  <label className="flex items-center gap-1.5 cursor-pointer text-[10px] text-natural-text border border-natural-border px-2 py-0.5 border-dashed rounded shrink-0 transition hover:bg-natural-accent">
                    <input 
                      type="checkbox" 
                      checked={isCouncilAppraisal}
                      onChange={(e) => setIsCouncilAppraisal(e.target.checked)}
                      className="w-3 h-3 rounded text-blue-600 focus:ring-blue-500"
                    />
                    Kết quả HĐTĐ (Excel)
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer text-[10px] text-natural-text border border-natural-border px-2 py-0.5 border-dashed rounded shrink-0 transition hover:bg-natural-accent">
                    <input 
                      type="checkbox" 
                      checked={excelIncludeRemarks}
                      onChange={(e) => setExcelIncludeRemarks(e.target.checked)}
                      className="w-3 h-3 rounded text-purple-600 focus:ring-purple-500"
                    />
                    Có lời nhận xét
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer text-[10px] text-natural-text border border-natural-border px-2 py-0.5 border-dashed rounded shrink-0 transition hover:bg-natural-accent">
                    <input 
                      type="checkbox"
                      checked={sortExcelBySchool}
                      onChange={(e) => setSortExcelBySchool(e.target.checked)}
                      className="w-3 h-3 rounded text-emerald-600 focus:ring-emerald-500"
                    />
                    Sắp xếp theo trường
                  </label>
                  <button 
                    onClick={handleExportExcel}
                    className="text-[10px] text-emerald-700 font-bold bg-emerald-50 px-2.5 py-0.5 rounded border border-emerald-200 hover:bg-emerald-100 flex items-center gap-1 cursor-pointer transition"
                  >
                    <FileDown className="w-3.5 h-3.5" /> Xuất Excel
                  </button>
                  <span className="text-[10px] text-natural-primary font-mono bg-natural-accent px-2 py-0.5 rounded border border-natural-border">Đã lưu offline</span>
                </div>
              </h3>
              <div className="flex flex-col gap-1.5 max-h-48 overflow-y-auto pr-1">
                {history.map((h) => {
                  const isCurrent = h.id === currentResult?.id;
                  return (
                    <div
                      key={h.id}
                      onClick={() => {
                        setCurrentResult(h);
                        if (h.teacher) setTeacher(h.teacher);
                        if (h.initiativeTitle) setInitiativeTitle(h.initiativeTitle);
                        if (h.appliedDate) setAppliedDate(h.appliedDate);
                        if (h.initiativeText) setInitiativeText(h.initiativeText);
                        setPlagResult(h.plagiarismResult || null);
                        setCouncilName(h.councilName || 'CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN');
                        setMember1Name(h.member1Name || '');
                        setMember1Unit(h.member1Unit || '');
                        setMember1Role(h.member1Role || '');
                        setMember2Name(h.member2Name || '');
                        setMember2Unit(h.member2Unit || '');
                        setMember2Role(h.member2Role || '');
                        setMember3Name(h.member3Name || '');
                        setMember3Unit(h.member3Unit || '');
                        setMember3Role(h.member3Role || '');

                        const loadedMembers: CouncilMember[] = h.councilMembers || [
                          { id: '1', name: h.member1Name || '', unit: h.member1Unit || '', role: h.member1Role || '' },
                          { id: '2', name: h.member2Name || '', unit: h.member2Unit || '', role: h.member2Role || '' },
                          { id: '3', name: h.member3Name || '', unit: h.member3Unit || '', role: h.member3Role || '' },
                        ];
                        setCouncilMembers(loadedMembers);
                      }}
                      className={`p-2.5 rounded-xl border text-xs flex items-center justify-between cursor-pointer transition ${isCurrent ? 'bg-[#f4f3e6] border-natural-secondary' : 'bg-natural-accent/30 border-natural-border hover:bg-natural-accent'}`}
                    >
                      <div className="flex flex-col gap-0.5 max-w-[80%]">
                        <span className={`font-semibold text-natural-text line-clamp-1 ${isCurrent ? 'text-natural-primary font-black' : ''}`}>
                          {h.initiativeTitle}
                        </span>
                        <span className="text-natural-muted text-[10px]">
                          {h.teacher.teacherName} — {h.teacher.schoolName}
                        </span>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${h.classification === 'Xuất sắc' ? 'bg-natural-primary text-white' : h.classification === 'Khá' ? 'bg-natural-secondary text-white' : 'bg-natural-muted text-white'}`}>
                          {h.totalScore}đ
                        </span>
                        <button
                          onClick={(e) => deleteHistoryItem(h.id, e)}
                          className="text-natural-muted hover:text-red-600 p-1 rounded hover:bg-white/50"
                          title="Xóa bản ghi"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* BẢNG KẾT QUẢ ĐÁNH GIÁ CHÍNH */}
          <div className="bg-white rounded-2xl border border-natural-border shadow-sm relative overflow-hidden flex-1 flex flex-col min-h-[400px]">
            
            {/* Loading state overlays */}
            {isLoading && (
              <div className="absolute inset-0 bg-white/95 z-20 flex flex-col items-center p-8 bg-gradient-to-b from-white to-[#fdfcf9] overflow-y-auto">
                <div className="text-center mt-6 mb-8 w-full">
                  <div className="relative mb-5 mx-auto w-16 h-16">
                    {/* Glowing, looping outer ring spinner */}
                    <div className="w-16 h-16 border-4 border-natural-accent border-t-natural-primary rounded-full animate-spin"></div>
                    <Sparkles className="w-6 h-6 text-natural-secondary absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 pulse-heart animate-pulse" />
                  </div>
                  <h4 className="font-bold text-natural-primary uppercase tracking-widest text-[13px] mb-2">{mainTab === 'appraisal' ? 'Đang tiến hành mô phỏng Hội đồng Thẩm định' : 'Đang quét và kiểm định Sơ bộ...'}</h4>
                  <p className="text-xs text-natural-muted">Vui lòng không đóng trình duyệt trong quá trình này (khoảng 30 giây)</p>
                </div>

                {(mainTab === 'appraisal' || mainTab === 'plagiarism') && (
                  <div className="w-full max-w-lg mx-auto bg-white p-6 rounded-2xl border border-natural-border shadow-sm">
                     <div className="flex flex-col gap-4">
                        {(mainTab === 'plagiarism' ? plagiarismLoaderMessages : loaderMessages).map((msg, index) => {
                           const isCompleted = index < loadingMessageIndex;
                           const isCurrent = index === loadingMessageIndex;
                           const isPending = index > loadingMessageIndex;
                           
                           return (
                             <div key={index} className={`flex items-start gap-4 p-2 rounded-lg transition-all duration-300 ${isCurrent ? 'bg-natural-accent/50 scale-[1.02] -mx-2 px-4 border border-natural-border shadow-sm' : 'opacity-70'}`}>
                               <div className="mt-0.5 shrink-0">
                                 {isCompleted ? (
                                   <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center">
                                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                                   </div>
                                 ) : isCurrent ? (
                                   <div className="w-5 h-5 rounded-full bg-blue-50 border border-blue-200 flex items-center justify-center">
                                      <Loader2 className="w-3.5 h-3.5 text-blue-600 animate-spin" />
                                   </div>
                                 ) : (
                                   <div className="w-5 h-5 rounded-full border border-natural-border bg-white flex items-center justify-center">
                                      <div className="w-1.5 h-1.5 rounded-full bg-natural-border"></div>
                                   </div>
                                 )}
                               </div>
                               <div className="flex-1 text-left">
                                 <p className={`text-[13px] ${isCurrent ? 'font-bold text-natural-primary' : isCompleted ? 'font-semibold text-natural-text' : 'text-natural-muted'}`}>
                                   {msg}
                                 </p>
                                 {isCurrent && (
                                   <div className="w-full bg-natural-accent h-1.5 rounded-full overflow-hidden mt-2 border border-natural-border">
                                     <div className="bg-natural-primary h-full animate-[shimmer_1.5s_infinite]" style={{ width: '80%', backgroundImage: 'linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent)' }}></div>
                                   </div>
                                 )}
                               </div>
                             </div>
                           );
                        })}
                     </div>
                  </div>
                )}
                
                {(mainTab !== 'appraisal' && mainTab !== 'plagiarism') && (
                  <div className="flex flex-col items-center mt-10">
                    <p className="text-sm text-natural-primary font-bold px-4 max-w-sm text-center">
                      {loaderMessages[loadingMessageIndex]}
                    </p>
                    <div className="w-48 bg-natural-accent h-1.5 rounded-full overflow-hidden mt-6">
                      <div className="bg-natural-primary h-full animate-[shimmer_1.5s_infinite]" style={{ width: '80%', backgroundImage: 'linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent)' }}></div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Empty state: No evaluation or plagiarism results yet */}
            {!currentResult && !plagResult && !isLoading && mainTab !== 'plagiarism' && (
              <div className="flex-1 flex flex-col items-center justify-center p-8 text-center bg-gradient-to-b from-white to-[#fdfcf9]">
                <div className="bg-natural-accent text-natural-secondary p-4 rounded-full mb-4">
                  <FileText className="w-12 h-12" />
                </div>
                <h3 className="font-bold text-natural-primary text-base mb-1.5">Trực phòng Hội đồng Sáng kiến</h3>
                <p className="text-sm text-natural-muted max-w-sm leading-relaxed">
                  Vui lòng tải tệp Sáng kiến (.PDF/.TXT) ở bước 1 để trợ lý AI bắt đầu tự động quét trùng lặp internet & chấm điểm thi đua.
                </p>
                
                <div className="mt-8 border-t border-natural-border pt-6 w-full text-left">
                  <span className="text-[10px] font-bold text-natural-muted tracking-wider uppercase block mb-3 text-center">Các bước quy trình kiểm định tự động</span>
                  <div className="grid grid-cols-2 gap-3">
                    <div className="bg-natural-accent/30 p-3 rounded-xl border border-natural-border text-xs">
                      <strong className="text-natural-primary block mb-1">1. Quét trùng lặp học thuật</strong>
                      <span className="text-natural-muted text-[11px]">Đối chiếu trực tiếp trên mạng internet đảm bảo tính mới ≥ 75%.</span>
                    </div>
                    <div className="bg-natural-accent/30 p-3 rounded-xl border border-natural-border text-xs">
                      <strong className="text-natural-primary block mb-1">2. Thẩm định cho điểm</strong>
                      <span className="text-natural-muted text-[11px]">Tự động chấm điểm 2 tiêu chí lớn theo Quyết định 270/QĐ-HĐSK.</span>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Real evaluated results display (Unified layout) */}
            {!isLoading && (
              <div className="flex-1 flex flex-col">

                {/* Empty Fallbacks for specific tabs */}
                {mainTab === 'appraisal' && !currentResult && (
                  <div className="flex-1 p-10 flex flex-col items-center justify-center gap-4">
                    <div className="text-sm font-medium text-natural-muted text-center italic">
                      Chưa có kết quả Thẩm định điểm cho tệp này.
                    </div>
                    
                    <div className="flex flex-col items-center gap-2 mb-2">
                       <label className="text-[10px] font-bold text-natural-muted uppercase tracking-wider">Mô hình AI sử dụng:</label>
                       <select
                          value={selectedModelAppraisal}
                          onChange={(e) => setSelectedModelAppraisal(e.target.value)}
                          className="text-xs font-semibold bg-white border border-amber-300 text-amber-900 rounded-lg px-3 py-1.5 outline-none w-[220px] shadow-sm"
                        >
                          <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro Preview (Mạnh mẽ nhất)</option>
                          <option value="gemini-pro">Gemini Pro Latest</option>
                          <option value="gemini-3.5-flash">Gemini 3.5 Flash (Cân bằng, Rất ổn định)</option>
                          <option value="gemini-3.0-flash-preview">Gemini 3 Flash Preview (Thử nghiệm)</option>
                          <option value="gemini-flash">Gemini Flash Latest</option>
                          <option value="gemini-2.5-flash">Gemini 2.5 Flash (Thế hệ trước)</option>
                          <option value="gemini-3.1-flash-lite">Gemini 3.1 Flash Lite (Nhanh nhất, Khuyên dùng)</option>
                        </select>
                    </div>

                    <button
                      type="button"
                      onClick={handleStartAppraisalOnly}
                      disabled={isLoading || (!plagFileBase64 && initiativeText.length < 30)}
                      className="min-w-[200px] border border-purple-300 bg-purple-50 hover:bg-purple-100 text-purple-800 text-xs font-bold uppercase tracking-wider py-3 px-6 rounded-xl transition flex justify-center items-center gap-1.5 cursor-pointer disabled:opacity-50 shadow-sm"
                    >
                      <Sparkles className="w-4 h-4" /> Bắt đầu Thẩm định (Chấm điểm)
                    </button>
                  </div>
                )}
                {mainTab === 'plagiarism' && (
                  <div className="flex-1 flex flex-col md:flex-row min-h-[750px] h-[calc(100vh-200px)]">
                    {/* Left Pane: PDF / Document Viewer */}
                    <div className="w-full md:w-[55%] border-r border-natural-border bg-gray-50 flex flex-col">
                      <div className="p-3 border-b border-natural-border bg-white shrink-0 flex items-center gap-2">
                        <FileText className="w-4 h-4 text-natural-primary" />
                        <span className="text-xs font-bold text-natural-text uppercase">Bản xem trước Sáng kiến</span>
                      </div>
                      <div className="flex-1 overflow-auto p-0 relative">
                        {initiativeText ? (
                          <div className="absolute inset-0 overflow-auto">
                            <HighlightedText text={initiativeText} result={plagResult} />
                          </div>
                        ) : plagFileBase64 ? (
                          <div className="absolute inset-0 flex items-center justify-center text-sm text-natural-muted italic p-4 text-center">
                            Đang xử lý văn bản...
                          </div>
                        ) : (
                          <div className="absolute inset-0 flex items-center justify-center text-sm text-natural-muted italic p-4 text-center">
                            Chưa có tệp nào được tải lên.
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Right Pane: AI & Spelling Results / Action */}
                    <div className="w-full md:w-[45%] flex flex-col bg-white">
                      <div className="p-3 border-b border-natural-border bg-[#faf9f5] shrink-0 flex items-center justify-between gap-2">
                         <span className="text-xs font-bold text-natural-text uppercase shrink-0">Phân tích Nội dung</span>
                         <div className="flex items-center gap-1.5">
                           <select
                              value={selectedModelPlag}
                              onChange={(e) => setSelectedModelPlag(e.target.value)}
                              className="text-[10px] font-semibold bg-white border border-amber-300 text-amber-900 rounded px-1.5 py-1 outline-none min-w-[120px] shadow-sm max-w-[140px] truncate"
                            >
                              <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro Preview (Mạnh mẽ nhất)</option>
                              <option value="gemini-pro">Gemini Pro Latest</option>
                              <option value="gemini-3.5-flash">Gemini 3.5 Flash (Cân bằng, Rất ổn định)</option>
                              <option value="gemini-3.0-flash-preview">Gemini 3 Flash Preview (Thử nghiệm)</option>
                              <option value="gemini-flash">Gemini Flash Latest</option>
                              <option value="gemini-2.5-flash">Gemini 2.5 Flash (Thế hệ trước)</option>
                              <option value="gemini-3.1-flash-lite">Gemini 3.1 Flash Lite (Nhanh nhất, Khuyên dùng)</option>
                            </select>
                            {plagResult && (
                              <button
                                onClick={() => handleStartPlagiarismCheckOnly()}
                                disabled={isLoading}
                                className="text-[10px] font-bold bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 px-2 py-1 rounded shadow-sm flex items-center gap-1 whitespace-nowrap disabled:opacity-50"
                              >
                                <RefreshCw className="w-3 h-3" /> Quét lại
                              </button>
                            )}
                         </div>
                      </div>
                      
                      {!plagResult ? (
                        <div className="flex-1 p-8 flex flex-col items-center justify-center gap-4">
                          <div className="text-sm font-medium text-natural-muted text-center italic max-w-[280px]">
                            Chưa quét dữ liệu. Vui lòng bấm dể nhận báo cáo AI, Chính tả, và Đạo văn từ nội dung bên trái.
                          </div>
                          <button
                            type="button"
                            onClick={() => handleStartPlagiarismCheckOnly()}
                            disabled={isLoading || (!plagFileBase64 && initiativeText.length < 30)}
                            className="w-full max-w-[240px] border border-blue-300 bg-blue-50 hover:bg-blue-100 text-blue-800 text-xs font-bold uppercase tracking-wider py-3 rounded-xl transition flex justify-center items-center gap-1.5 cursor-pointer disabled:opacity-50"
                          >
                            <Search className="w-4 h-4" /> Bắt đầu quét
                          </button>
                        </div>
                      ) : (
                        <div className="flex-1 flex flex-col p-4 gap-4 overflow-y-auto custom-scrollbar">
                          <div className="flex flex-col gap-3 bg-natural-accent/20 p-4 rounded-xl border border-natural-border font-sans shrink-0">
                            <div className="flex justify-between items-center border-b border-natural-border pb-1.5">
                              <h4 className="text-xs font-bold text-natural-primary uppercase">Tổng hợp quét Đạo văn, AI & Chính tả</h4>
                              <button 
                                type="button"
                                onClick={() => setIsEditingPlag(!isEditingPlag)} 
                                className="px-2.5 py-1 text-[10px] font-bold uppercase tracking-wider rounded bg-white hover:bg-natural-accent border border-natural-border text-natural-primary transition cursor-pointer flex items-center gap-1.5 shadow-sm"
                              >
                                <Pencil className="w-3 h-3 text-natural-muted" /> {isEditingPlag ? "Đóng chỉnh sửa" : "Chỉnh sửa"}
                              </button>
                            </div>
                            
                            {isEditingPlag && (
                              <div className="bg-white/90 p-3 rounded-lg border border-natural-border/80 flex flex-col gap-3 text-xs mb-1 shadow-sm">
                                <div className="font-bold text-[10px] text-natural-primary uppercase tracking-wider">Điều chỉnh kết quả rà soát & chính tả</div>
                                <div className="grid grid-cols-2 gap-3">
                                  <div>
                                    <label className="block text-[10px] font-bold text-natural-muted uppercase mb-1">Tỷ lệ Trùng lặp (%)</label>
                                    <input 
                                      type="number" 
                                      min={0} max={100} 
                                      value={plagResult.totalDuplicatePercent}
                                      onChange={(e) => {
                                        const val = Math.min(100, Math.max(0, parseInt(e.target.value) || 0));
                                        let level = plagResult.warningLevel;
                                        if (val <= 15) level = 'An toàn (Thấp)';
                                        else if (val <= 25) level = 'Vi phạm nhẹ';
                                        else if (val <= 30) level = 'Vi phạm trung bình';
                                        else level = 'Vi phạm nghiêm trọng (Vượt quy định)';
                                        
                                        updatePlagiarismDataInResult({
                                          ...plagResult,
                                          totalDuplicatePercent: val,
                                          warningLevel: level
                                        });
                                      }}
                                      className="w-full text-xs font-semibold border border-natural-border rounded px-2 py-1 focus:outline-none focus:border-natural-primary"
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[10px] font-bold text-natural-muted uppercase mb-1">Tỷ lệ viết bằng AI (%)</label>
                                    <input 
                                      type="number" 
                                      min={0} max={100} 
                                      value={plagResult.aiGeneratedPercent || 0}
                                      onChange={(e) => {
                                        const val = Math.min(100, Math.max(0, parseInt(e.target.value) || 0));
                                        updatePlagiarismDataInResult({
                                          ...plagResult,
                                          aiGeneratedPercent: val
                                        });
                                      }}
                                      className="w-full text-xs font-semibold border border-natural-border rounded px-2 py-1 focus:outline-none focus:border-natural-primary"
                                    />
                                  </div>
                                </div>
                                <div className="grid grid-cols-1 gap-2">
                                  <div>
                                    <label className="block text-[10px] font-bold text-natural-muted uppercase mb-1">Cảnh báo / Kết luận đạo văn</label>
                                    <select
                                      value={plagResult.warningLevel}
                                      onChange={(e) => {
                                        updatePlagiarismDataInResult({
                                          ...plagResult,
                                          warningLevel: e.target.value
                                        });
                                      }}
                                      className="w-full text-xs font-semibold border border-natural-border rounded px-2 py-1 bg-white focus:outline-none focus:border-natural-primary"
                                    >
                                      <option value="An toàn (Thấp)">An toàn (Thấp)</option>
                                      <option value="Vi phạm nhẹ">Vi phạm nhẹ</option>
                                      <option value="Vi phạm trung bình">Vi phạm trung bình</option>
                                      <option value="Vi phạm nghiêm trọng">Vi phạm nghiêm trọng</option>
                                      <option value="Không đạt (Đạo văn quá quy định)">Không đạt (Đạo văn quá quy định)</option>
                                    </select>
                                  </div>
                                </div>
                                
                                <div className="border-t border-natural-border/60 pt-2">
                                  <div className="font-bold text-[10px] text-natural-primary uppercase tracking-wider mb-1.5">Thêm lỗi chính tả thủ công</div>
                                  <div className="grid grid-cols-2 gap-2 mb-2">
                                    <input 
                                      type="text" 
                                      id="new-error-text"
                                      placeholder="Từ viết sai" 
                                      className="text-xs border border-natural-border rounded px-2 py-1 focus:outline-none"
                                    />
                                    <input 
                                      type="text" 
                                      id="new-error-correction"
                                      placeholder="Sửa lại thành" 
                                      className="text-xs border border-natural-border rounded px-2 py-1 focus:outline-none"
                                    />
                                  </div>
                                  <input 
                                    type="text" 
                                    id="new-error-reason"
                                    placeholder="Lý do / Nguyên tắc chính tả" 
                                    className="w-full text-xs border border-natural-border rounded px-2 py-1 mb-2 focus:outline-none"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const errTextEl = document.getElementById('new-error-text') as HTMLInputElement;
                                      const corrEl = document.getElementById('new-error-correction') as HTMLInputElement;
                                      const reasonEl = document.getElementById('new-error-reason') as HTMLInputElement;
                                      if (errTextEl && corrEl) {
                                        const errText = errTextEl.value.trim();
                                        const correction = corrEl.value.trim();
                                        const reason = reasonEl ? reasonEl.value.trim() : '';
                                        if (errText && correction) {
                                          const newErr = {
                                            errorText: errText,
                                            correction: correction,
                                            context: `Trong văn bản có từ "${errText}"`,
                                            reason: reason || 'Lỗi chính tả thông thường'
                                          };
                                          const updatedErrors = [...(plagResult.spellingErrors || []), newErr];
                                          updatePlagiarismDataInResult({
                                            ...plagResult,
                                            spellingErrors: updatedErrors
                                          });
                                          errTextEl.value = '';
                                          corrEl.value = '';
                                          if (reasonEl) reasonEl.value = '';
                                        }
                                      }
                                    }}
                                    className="w-full bg-natural-primary hover:bg-[#434330] text-white font-bold py-1 px-3 rounded text-[10px] uppercase transition cursor-pointer"
                                  >
                                    Thêm lỗi chính tả
                                  </button>
                                </div>
                              </div>
                            )}

                            <div className="grid grid-cols-3 gap-3">
                              {/* Đạo văn Circle */}
                              <div className="flex flex-col items-center bg-white p-2.5 rounded-xl border border-natural-border/60 shadow-sm">
                                <div className="relative w-14 h-14 flex items-center justify-center">
                                  <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                                    <circle cx="50" cy="50" r="40" stroke="#e5e7eb" strokeWidth="8" fill="transparent" />
                                    <circle 
                                      cx="50" cy="50" r="40" 
                                      stroke={plagResult.totalDuplicatePercent > 25 ? "#dc2626" : "#8b8b68"} 
                                      strokeWidth="8" fill="transparent" 
                                      strokeDasharray={`${2 * Math.PI * 40}`}
                                      strokeDashoffset={`${2 * Math.PI * 40 * (1 - plagResult.totalDuplicatePercent / 100)}`}
                                      strokeLinecap="round"
                                    />
                                  </svg>
                                  <div className="absolute flex flex-col items-center">
                                    <span className={`text-[12px] font-black font-mono ${plagResult.totalDuplicatePercent > 25 ? 'text-red-600' : 'text-natural-primary'}`}>{plagResult.totalDuplicatePercent}%</span>
                                  </div>
                                </div>
                                <span className="text-[10px] font-bold mt-2 text-natural-text text-center leading-tight">Đạo văn<br/>(Trùng lặp)</span>
                              </div>

                              {/* AI Circle */}
                              <div className="flex flex-col items-center bg-white p-2.5 rounded-xl border border-natural-border/60 shadow-sm">
                                <div className="relative w-14 h-14 flex items-center justify-center">
                                  <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                                    <circle cx="50" cy="50" r="40" stroke="#f3e8ff" strokeWidth="8" fill="transparent" />
                                    <circle 
                                      cx="50" cy="50" r="40" 
                                      stroke="#9333ea" 
                                      strokeWidth="8" fill="transparent" 
                                      strokeDasharray={`${2 * Math.PI * 40}`}
                                      strokeDashoffset={`${2 * Math.PI * 40 * (1 - (plagResult.aiGeneratedPercent || 0) / 100)}`}
                                      strokeLinecap="round"
                                    />
                                  </svg>
                                  <div className="absolute flex flex-col items-center">
                                    <span className="text-[12px] font-black text-purple-700 font-mono">{plagResult.aiGeneratedPercent || 0}%</span>
                                  </div>
                                </div>
                                <span className="text-[10px] font-bold mt-2 text-purple-800 text-center leading-tight">Sử dụng<br/>AI viết</span>
                              </div>

                              {/* Chính tả Circle */}
                              <div className="flex flex-col items-center bg-white p-2.5 rounded-xl border border-natural-border/60 shadow-sm">
                                <div className="relative w-14 h-14 flex items-center justify-center">
                                  <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                                    <circle cx="50" cy="50" r="40" stroke="#ffedd5" strokeWidth="8" fill="transparent" />
                                    <circle 
                                      cx="50" cy="50" r="40" 
                                      stroke="#ea580c" 
                                      strokeWidth="8" fill="transparent" 
                                      strokeDasharray={`${2 * Math.PI * 40}`}
                                      strokeDashoffset={`${2 * Math.PI * 40 * (1 - Math.min(100, ((plagResult.spellingErrors?.length || 0) / Math.max(1, (plagResult.extractedText?.split(/\s+/).length || 1))) * 100 * 5) / 100)}`}
                                      strokeLinecap="round"
                                    />
                                  </svg>
                                  <div className="absolute flex flex-col items-center">
                                    <span className="text-[12px] font-black text-orange-600 font-mono">
                                      {plagResult.spellingErrors?.length ? ((plagResult.spellingErrors.length / Math.max(1, plagResult.extractedText?.split(/\s+/).length || 1)) * 100).toFixed(1) : 0}%
                                    </span>
                                  </div>
                                </div>
                                <span className="text-[10px] font-bold mt-2 text-orange-700 text-center leading-tight">Lỗi chính tả<br/>({plagResult.spellingErrors?.length || 0} lỗi)</span>
                              </div>
                            </div>
                            
                            <div className="mt-1 bg-white p-3 rounded-lg border border-natural-border shadow-sm">
                              <p className="text-xs font-semibold text-natural-text flex items-center justify-between">
                                Kết luận: 
                                <span className={`font-bold px-2 py-1 rounded-md text-white ${plagResult.totalDuplicatePercent > 25 || (plagResult.aiGeneratedPercent || 0) > 20 ? 'bg-red-600' : 'bg-emerald-600'}`}>
                                  {plagResult.totalDuplicatePercent > 25 || (plagResult.aiGeneratedPercent || 0) > 20 ? 'Không đạt yêu cầu' : 'Đạt yêu cầu (An toàn)'}
                                </span>
                              </p>
                              <p className="text-[10px] text-natural-muted mt-2 border-t border-natural-border pt-2 italic">
                                {plagResult.warningLevel} - Hệ thống phân tích văn bản qua kho dữ liệu khổng lồ và nhận diện ngôn ngữ AI.
                              </p>
                            </div>
                            
                            {plagResult.aiSegments && plagResult.aiSegments.length > 0 && (
                              <div className="mt-3 bg-purple-50 p-2 rounded-lg border border-purple-100">
                                <span className="text-[10px] font-bold text-purple-800 uppercase tracking-wider block mb-1">
                                  Dấu hiệu AI sinh (Nghi ngờ):
                                </span>
                                <ul className="list-disc pl-3 space-y-1">
                                  {plagResult.aiSegments.slice(0, 2).map((aiText, idx) => (
                                    <li key={idx} className="text-[9px] text-purple-700 italic line-clamp-3">
                                      "{aiText}"
                                    </li>
                                  ))}
                                  {plagResult.aiSegments.length > 2 && (
                                    <li className="text-[9px] text-purple-600 font-medium list-none italic mt-1">
                                      + {plagResult.aiSegments.length - 2} đoạn khác...
                                    </li>
                                  )}
                                </ul>
                              </div>
                            )}
                          </div>

                          {(plagResult.spellingErrors && plagResult.spellingErrors.length > 0) && (
                            <div className="space-y-2">
                              <span className="text-[11px] font-bold text-natural-muted uppercase tracking-wider block">
                                Lỗi chính tả đề xuất ({plagResult.spellingErrors.length}):
                              </span>
                              <div className="grid grid-cols-1 gap-2">
                                {plagResult.spellingErrors.map((err, i) => (
                                  <div key={i} className="text-xs p-3 rounded-lg bg-red-50/50 border border-red-100 flex flex-col gap-1.5 relative group">
                                    <div className="flex justify-between items-start gap-2">
                                      <div className="flex flex-wrap gap-1 items-center">
                                        <span className="font-semibold text-red-700 line-through bg-red-100/50 px-1 rounded">{err.errorText}</span> 
                                        <ArrowRight className="w-3 h-3 text-natural-muted mx-0.5" /> 
                                        <span className="font-semibold text-emerald-700 bg-emerald-100/50 px-1 rounded">{err.correction}</span>
                                      </div>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const updatedErrors = plagResult.spellingErrors?.filter((_, index) => index !== i) || [];
                                          updatePlagiarismDataInResult({
                                            ...plagResult,
                                            spellingErrors: updatedErrors
                                          });
                                        }}
                                        className="p-1 hover:bg-red-100 text-red-600 rounded transition cursor-pointer"
                                        title="Xóa lỗi này"
                                      >
                                        <Trash2 className="w-3.5 h-3.5" />
                                      </button>
                                    </div>
                                    <div className="text-[10px] italic text-neutral-500 border-l-2 border-red-200 pl-2">"...{err.context}..."</div>
                                    <div className="text-[10px] text-neutral-600 font-medium">{err.reason}</div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}

                          <div className="space-y-2 mt-4">
                            <span className="text-[11px] font-bold text-natural-muted uppercase tracking-wider block">
                              Nguồn vi phạm nổi bật:
                            </span>
                            {plagResult.sources.length === 0 ? (
                              <p className="text-xs text-emerald-800 font-bold tracking-normal py-3 text-center italic bg-emerald-50 rounded-xl border border-emerald-100 px-3">
                                ✅ Trí tuệ độc lập (0% trùng lặp).
                              </p>
                            ) : (
                              <div className="flex flex-col gap-2">
                                {plagResult.sources.map((src, idx) => {
                                  const isHighlighted = hoveredSourceId === src.id || activeSourceId === src.id;
                                  const hasDetailed = !!src.detailed_source;
                                  const title = hasDetailed ? src.detailed_source?.document_title : src.name;
                                  
                                  return (
                                    <div
                                      key={src.id}
                                      id={`source-card-${src.id}`}
                                      onMouseEnter={() => setHoveredSourceId(src.id)}
                                      onMouseLeave={() => setHoveredSourceId(null)}
                                      onClick={() => {
                                        setActiveSourceId(src.id);
                                        const el = document.getElementById(`duplicate-${src.id}`);
                                        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                      }}
                                      className={`p-3 rounded-lg border transition-all duration-300 flex flex-col gap-2 cursor-pointer ${
                                        isHighlighted
                                          ? 'bg-natural-accent border-natural-primary shadow-sm scale-[1.02] ring-1 ring-natural-secondary/30'
                                          : 'bg-[#faf9f5] border-natural-border hover:bg-natural-accent/40'
                                      }`}
                                    >
                                      <div className="flex items-start justify-between gap-2">
                                        <div className="flex items-start gap-2 max-w-[70%]">
                                          <span className="inline-flex items-center justify-center bg-red-500 text-white rounded-sm text-[11px] font-bold px-1.5 py-0.5 mr-1 shrink-0">
                                            {idx + 1}
                                          </span>
                                          <div className="space-y-0.5 w-full overflow-hidden">
                                            <p className="text-[11px] font-bold text-[#4d4d38] leading-tight line-clamp-2" title={title}>
                                              {src.match_percent ?? src.percent}% - {title || src.name}
                                            </p>
                                            <p className="text-[9px] text-natural-muted">
                                              {hasDetailed && src.detailed_source && src.detailed_source.author && src.detailed_source.author !== 'Ẩn danh' && !src.detailed_source.author.includes('Không xác định') && (
                                                <span className="mr-2">Tác giả: {src.detailed_source.author}</span>
                                              )}
                                              {hasDetailed && src.detailed_source?.website_name && (
                                                <span>Nguồn: {src.detailed_source.website_name}</span>
                                              )}
                                            </p>
                                            {(() => {
                                              const exactUrlOk = hasDetailed && src.detailed_source?.exact_url && (!src.detailed_source.exact_url.includes('sangkienkinhnghiem.net') && !src.detailed_source.exact_url.includes('giaoan.link'));
                                              const searchUrl = src.detailed_source?.search_keywords ? `https://www.google.com/search?q=${encodeURIComponent(src.detailed_source.search_keywords)}` : null;
                                              let linkToShow = exactUrlOk ? src.detailed_source!.exact_url : searchUrl;
                                              if (exactUrlOk && src.detailed_source?.matched_snippet) {
                                                linkToShow += `#:~:text=${encodeURIComponent(src.detailed_source.matched_snippet)}`;
                                              }
                                              
                                              const alternativeUrls = src.detailed_source?.alternative_urls || [];
                                              
                                              if (!linkToShow && alternativeUrls.length === 0) {
                                                return (
                                                  <div className="flex items-center gap-1.5 mt-1">
                                                    <span className="text-[10px] font-medium text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200 truncate w-[160px] sm:w-[250px]" title="Nguồn minh chứng từ Kho tri thức Số ngành Giáo dục">
                                                      Nguồn minh chứng từ Kho tri thức Số ngành Giáo dục
                                                    </span>
                                                  </div>
                                                );
                                              }
                                              
                                              const allLinks = [];
                                              if (linkToShow) {
                                                allLinks.push({ url: linkToShow, label: 'Nguồn chính' });
                                              }
                                              alternativeUrls.forEach((url, i) => {
                                                if (url && typeof url === 'string' && url.trim().startsWith('http')) {
                                                  allLinks.push({ url: url.trim(), label: `Minh chứng ${i + 1}` });
                                                }
                                              });

                                              return (
                                                <div className="flex flex-col gap-1 mt-1">
                                                  {allLinks.map((item, i) => (
                                                    <div key={i} className="flex items-center gap-1.5">
                                                      <Link className="w-3 h-3 text-blue-500 shrink-0" />
                                                      <span className="text-[9px] font-bold text-blue-800 shrink-0">{item.label}:</span>
                                                      <a
                                                        href={item.url}
                                                        target="_blank"
                                                        rel="noreferrer"
                                                        className="text-[9px] text-blue-600 hover:text-blue-800 hover:underline truncate block w-[120px] sm:w-[160px]"
                                                        title={item.url}
                                                      >
                                                        {item.url}
                                                      </a>
                                                    </div>
                                                  ))}
                                                </div>
                                              );
                                            })()}
                                          </div>
                                        </div>
                                        {(() => {
                                          const exactUrlOk = hasDetailed && src.detailed_source?.exact_url && (!src.detailed_source.exact_url.includes('sangkienkinhnghiem.net') && !src.detailed_source.exact_url.includes('giaoan.link'));
                                          const searchUrl = src.detailed_source?.search_keywords ? `https://www.google.com/search?q=${encodeURIComponent(src.detailed_source.search_keywords)}` : null;
                                          let linkToShow = exactUrlOk ? src.detailed_source!.exact_url : searchUrl;
                                          if (exactUrlOk && src.detailed_source?.matched_snippet) {
                                            linkToShow += `#:~:text=${encodeURIComponent(src.detailed_source.matched_snippet)}`;
                                          }
                                          if (!linkToShow) return null;
                                          return (
                                          <a
                                            href={linkToShow}
                                            target="_blank"
                                            rel="noreferrer"
                                            className="shrink-0 text-[10px] font-bold text-natural-primary uppercase bg-white border border-natural-border rounded px-2 py-1 flex items-center gap-1 hover:bg-natural-primary hover:text-white transition"
                                          >
                                            <ExternalLink className="w-3 h-3" /> {exactUrlOk ? 'NGUỒN' : 'TÌM KIẾM'}
                                          </a>
                                          );
                                        })()}
                                      </div>
                                      {hasDetailed && src.detailed_source?.matched_snippet && (
                                        <div className="mt-1 pl-6">
                                          <p className="text-[10px] text-red-700 bg-red-50/50 p-1.5 border border-red-100 rounded italic line-clamp-3">
                                            "{src.detailed_source.matched_snippet}"
                                          </p>
                                        </div>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                          
                          {plagResult.spellingErrors && plagResult.spellingErrors.length > 0 && (
                            <div className="space-y-2 mt-4">
                              <span className="text-[11px] font-bold text-natural-muted uppercase tracking-wider block">
                                Lỗi chính tả & Cấu trúc cần khắc phục:
                              </span>
                              <div className="flex flex-col gap-2 max-h-[300px] overflow-y-auto custom-scrollbar pr-2">
                                {plagResult.spellingErrors.map((err, idx) => (
                                  <div
                                    key={idx}
                                    onClick={() => {
                                      const el = document.getElementById(`spelling-${idx + 1}`);
                                      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                    }}
                                    className="p-3 rounded-lg border border-red-100 bg-red-50 hover:bg-red-100 transition-colors flex flex-col gap-1.5 cursor-pointer"
                                  >
                                    <div className="flex justify-between items-start">
                                      <div className="text-[11px] font-semibold text-red-900">
                                        <span className="inline-flex items-center justify-center w-4 h-4 bg-red-600 text-white rounded-full text-[9px] mr-1.5 font-bold">{idx + 1}</span>
                                        Từ chấm lỗi: <span className="line-through opacity-80">{err.errorText}</span>
                                      </div>
                                      <span className="text-[10px] font-bold text-red-700 uppercase bg-white border border-red-200 rounded px-1.5 py-0.5">Sửa thành: {err.correction}</span>
                                    </div>
                                    <p className="text-[10px] text-red-800/80 italic mt-0.5 leading-tight">{err.reason}</p>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                          
                          <div className="mt-2 pt-3 border-t border-natural-border flex justify-end shrink-0">
                            <button
                              type="button"
                              onClick={() => setMainTab('appraisal')}
                              className="w-full py-2.5 border border-[#8b8b68] bg-[#fdfdfa] hover:bg-[#8b8b68] hover:text-white text-[#8b8b68] text-xs font-bold uppercase tracking-wider rounded-lg transition flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                            >
                               Tiếp tục: 3. Thẩm định (Chấm điểm) <ArrowRight className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}
                {/* End merged tab */}

                {mainTab === 'appraisal' && currentResult && (() => {
                  const isCouncilView = activeAppraisalView === 'council';
                  const hasCouncilResult = !!currentResult.councilResult;
                  const displayedEval = isCouncilView && hasCouncilResult ? currentResult.councilResult! : currentResult;

                  const uuDiemList = ('uuDiem' in displayedEval ? (displayedEval as any).uuDiem : currentResult.uuDiem) || [];
                  const hanCheList = ('hanChe' in displayedEval ? (displayedEval as any).hanChe : currentResult.hanChe) || [];
                  const improvementsList = ('improvements' in displayedEval ? (displayedEval as any).improvements : currentResult.improvements) || [];

                  return (
                    <div className="flex-1 flex flex-col overflow-hidden">
                      {/* Segmented control / Tab switcher */}
                      <div className="p-4 bg-[#faf9f5] border-b border-natural-border shrink-0">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                          <div className="flex bg-[#f3f1e8] p-1 rounded-xl gap-1 max-w-sm w-full border border-natural-border/30">
                            <button
                              type="button"
                              onClick={() => setActiveAppraisalView('department')}
                              className={`flex-1 py-1.5 px-3 text-xs font-bold uppercase tracking-wider rounded-lg transition-all flex items-center justify-center gap-1.5 ${!isCouncilView ? 'bg-natural-primary text-white shadow-sm' : 'text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
                            >
                              Sơ bộ (Phòng VH-XH)
                            </button>
                            <button
                              type="button"
                              onClick={() => setActiveAppraisalView('council')}
                              className={`flex-1 py-1.5 px-3 text-xs font-bold uppercase tracking-wider rounded-lg transition-all flex items-center justify-center gap-1.5 ${isCouncilView ? 'bg-amber-700 text-white shadow-sm' : 'text-natural-muted hover:text-natural-text hover:bg-natural-accent/50'}`}
                            >
                              Hội đồng Thẩm định {hasCouncilResult && <span className="ml-1 px-1.5 py-0.5 text-[9px] bg-white text-amber-700 rounded-full font-extrabold">Đã chấm</span>}
                            </button>
                          </div>
                          
                          {/* Print & Action Buttons */}
                          {(!isCouncilView || hasCouncilResult) && (
                            <div className="flex items-center gap-2">
                              <button
                                type="button"
                                onClick={() => setShowPrintModal(true)}
                                className="px-3 py-1.5 bg-natural-primary hover:bg-[#434330] border border-natural-primary hover:border-[#434330] text-[#eaeada] hover:text-white rounded-lg text-xs font-bold uppercase tracking-wider shadow transition flex items-center gap-1.5 cursor-pointer"
                              >
                                <Printer className="w-4 h-4" /> Thống kê & In phiếu
                              </button>
                              
                              <button
                                type="button"
                                onClick={() => {
                                  setShowReAppraisalPanel(!showReAppraisalPanel);
                                  if (!showReAppraisalPanel) {
                                    setDesiredScore(displayedEval.totalScore || '');
                                  }
                                }}
                                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-[#f3f1e8] hover:bg-[#8b8b68] text-[#8b8b68] hover:text-white rounded-lg transition border border-[#8b8b68]/30 shadow-sm cursor-pointer"
                              >
                                <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
                                {showReAppraisalPanel 
                                  ? (isCouncilView ? "Đóng bảng thẩm định" : "Đóng bảng chấm lại") 
                                  : (isCouncilView ? "Thẩm định lại" : "Chấm điểm lại")
                                }
                              </button>
                            </div>
                          )}
                        </div>

                        {/* Special banner for Council view */}
                        {isCouncilView && hasCouncilResult && (
                          <div className="mt-3 p-2 bg-amber-50 border border-amber-200 rounded-lg text-[11px] text-amber-900 font-semibold flex items-center gap-1.5">
                            <Award className="w-4 h-4 text-amber-700" /> Kết quả chấm điểm và đánh giá độc lập của Hội đồng Thẩm định Sáng kiến xã Hàm Yên.
                          </div>
                        )}
                      </div>

                      {/* Main evaluation section body */}
                      {isCouncilView && !hasCouncilResult ? (
                        <div className="flex-1 overflow-y-auto p-8 flex flex-col items-center justify-center text-center gap-5">
                          <div className="w-16 h-16 bg-amber-50 rounded-full flex items-center justify-center border border-amber-200">
                            <Award className="w-8 h-8 text-amber-700 animate-pulse" />
                          </div>
                          <div className="max-w-md space-y-2">
                            <h3 className="text-sm font-bold text-natural-text uppercase tracking-wider">Hội đồng Thẩm định Sáng kiến</h3>
                            <p className="text-xs text-natural-muted leading-relaxed">
                              Sáng kiến này hiện chưa có kết quả chấm điểm độc lập từ Hội đồng Thẩm định xã Hàm Yên.
                              Nhận xét của Hội đồng sẽ được tạo mới hoàn toàn với góc nhìn tập thể khách quan, phản biện sâu sắc, và bám sát các tiêu chuẩn của Nghị định số 13/2012/NĐ-CP.
                            </p>
                          </div>

                          <div className="flex flex-col items-center gap-2 max-w-xs w-full bg-amber-50/50 p-4 rounded-xl border border-amber-200/50">
                            <label className="text-[10px] font-extrabold text-amber-900 uppercase tracking-wider flex items-center gap-1">
                              <Sparkles className="w-3.5 h-3.5 text-amber-600 animate-pulse" /> Mô hình AI sử dụng:
                            </label>
                            <select
                              value={selectedModelAppraisal}
                              onChange={(e) => setSelectedModelAppraisal(e.target.value)}
                              className="text-xs font-semibold bg-white border border-amber-300 text-amber-900 rounded-lg px-3 py-2 outline-none w-full shadow-sm focus:border-amber-500 cursor-pointer text-center"
                            >
                              <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro Preview (Mạnh mẽ nhất)</option>
                              <option value="gemini-pro">Gemini Pro Latest</option>
                              <option value="gemini-3.5-flash">Gemini 3.5 Flash (Cân bằng, Rất ổn định)</option>
                              <option value="gemini-3.0-flash-preview">Gemini 3 Flash Preview (Thử nghiệm)</option>
                              <option value="gemini-flash">Gemini Flash Latest</option>
                              <option value="gemini-2.5-flash">Gemini 2.5 Flash (Thế hệ trước)</option>
                              <option value="gemini-3.1-flash-lite">Gemini 3.1 Flash Lite (Nhanh nhất, Khuyên dùng)</option>
                            </select>
                          </div>

                          <button
                            type="button"
                            disabled={isLoading}
                            onClick={runCouncilAppraisal}
                            className="px-5 py-2.5 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold uppercase tracking-widest rounded-xl transition shadow hover:shadow-md flex items-center gap-2 cursor-pointer disabled:opacity-50"
                          >
                            {isLoading ? (
                              <>
                                <RefreshCw className="w-4 h-4 animate-spin" /> Đang thẩm định...
                              </>
                            ) : (
                              <>
                                <Sparkles className="w-4 h-4" /> Kích hoạt Đánh giá Hội đồng
                              </>
                            )}
                          </button>
                        </div>
                      ) : (
                        <>
                          {/* Re-appraisal panel inside the viewport */}
                          {showReAppraisalPanel && (
                            <div className="p-5 border-b border-natural-border bg-amber-50/20 shrink-0">
                              <div className="bg-amber-50/60 border border-amber-200 p-4 rounded-xl flex flex-col gap-3 text-left shadow-inner">
                                <div className="flex flex-col gap-1">
                                  <span className="text-xs font-bold text-amber-950 flex items-center gap-1">
                                    <Sparkles className="w-4 h-4 text-amber-600" /> {isCouncilView ? "Tính năng Thẩm Định Lại Hội Đồng (AI Re-evaluation)" : "Tính năng Chấm Điểm Lại (AI Re-evaluation)"}
                                  </span>
                                  <p className="text-[11px] text-amber-800 leading-relaxed">
                                    Chọn mô hình, chế độ thẩm định, điểm số mong muốn và các lưu ý riêng để trợ lý AI thực hiện thẩm định và lập luận nhận xét chính xác nhất.
                                  </p>
                                </div>

                                {/* Model and Mode selection row */}
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1.5">
                                    <label className="text-[10px] font-bold text-amber-900 uppercase">Mô hình AI sử dụng:</label>
                                    <select
                                      value={selectedModelAppraisal}
                                      onChange={(e) => setSelectedModelAppraisal(e.target.value)}
                                      className="text-xs font-semibold bg-white border border-amber-200 text-amber-900 rounded-lg p-2 outline-none focus:border-amber-400 shadow-sm"
                                    >
                                      <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro Preview (Mạnh mẽ nhất)</option>
                                      <option value="gemini-pro">Gemini Pro Latest</option>
                                      <option value="gemini-3.5-flash">Gemini 3.5 Flash (Cân bằng, Rất ổn định)</option>
                                      <option value="gemini-3.0-flash-preview">Gemini 3 Flash Preview (Thử nghiệm)</option>
                                      <option value="gemini-flash">Gemini Flash Latest</option>
                                      <option value="gemini-2.5-flash">Gemini 2.5 Flash (Thế hệ trước)</option>
                                      <option value="gemini-3.1-flash-lite">Gemini 3.1 Flash Lite (Nhanh nhất, Khuyên dùng)</option>
                                    </select>
                                  </div>
                                  <div className="flex flex-col gap-1.5">
                                    <label className="text-[10px] font-bold text-amber-900 uppercase">Chế độ phân tích AI:</label>
                                    <select
                                      value={evaluationMode}
                                      onChange={(e: any) => setEvaluationMode(e.target.value)}
                                      className="text-xs font-semibold bg-white border border-amber-200 text-amber-900 rounded-lg p-2 outline-none focus:border-amber-400 shadow-sm"
                                    >
                                      <option value="full">Chấm điểm & Xếp loại đầy đủ</option>
                                      <option value="comment_only">Chỉ Góp ý chuyên môn</option>
                                    </select>
                                  </div>
                                </div>

                                {/* Desired score and custom notes row */}
                                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                  {evaluationMode === 'full' ? (
                                    <>
                                      <div className="flex flex-col gap-1.5 sm:col-span-1">
                                        <label className="text-[10px] font-bold text-amber-900 uppercase">Điểm số mong muốn (0-100):</label>
                                        <input
                                          type="number"
                                          min="0"
                                          max="100"
                                          value={desiredScore}
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            if (val === '') {
                                              setDesiredScore('');
                                            } else {
                                              const num = Math.min(100, Math.max(0, parseInt(val) || 0));
                                              setDesiredScore(num);
                                            }
                                          }}
                                          placeholder="Ví dụ: 88"
                                          className="text-xs bg-white border border-amber-200 text-amber-900 rounded-lg p-2 outline-none focus:border-amber-400 font-bold shadow-sm"
                                        />
                                      </div>
                                      <div className="flex flex-col gap-1.5 sm:col-span-2">
                                        <label className="text-[10px] font-bold text-amber-900 uppercase">Yêu cầu chỉnh sửa / Lưu ý riêng cho AI:</label>
                                        <input
                                          type="text"
                                          value={reAppraisalNotes}
                                          onChange={(e) => setReAppraisalNotes(e.target.value)}
                                          placeholder="Ví dụ: Nâng điểm vì tính ứng dụng cao / Tập trung phân tích biện pháp..."
                                          className="text-xs bg-white border border-amber-200 text-amber-900 rounded-lg p-2 outline-none focus:border-amber-400 shadow-sm"
                                        />
                                      </div>
                                    </>
                                  ) : (
                                    <div className="flex flex-col gap-1.5 sm:col-span-3">
                                      <label className="text-[10px] font-bold text-amber-900 uppercase">Yêu cầu chỉnh sửa / Lưu ý riêng cho AI:</label>
                                      <input
                                        type="text"
                                        value={reAppraisalNotes}
                                        onChange={(e) => setReAppraisalNotes(e.target.value)}
                                        placeholder="Ví dụ: Tập trung làm rõ các điểm cần khắc phục và nâng cao chất lượng sư phạm..."
                                        className="text-xs bg-white border border-amber-200 text-amber-900 rounded-lg p-2 outline-none focus:border-amber-400 shadow-sm"
                                      />
                                    </div>
                                  )}
                                </div>

                                <div className="flex justify-end gap-2 mt-1">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setShowReAppraisalPanel(false);
                                      setDesiredScore('');
                                      setReAppraisalNotes('');
                                    }}
                                    className="px-3 py-1.5 bg-white border border-amber-200 hover:bg-amber-100 text-amber-900 font-bold text-xs rounded-lg transition cursor-pointer"
                                  >
                                    Hủy bỏ
                                  </button>
                                  <button
                                    type="button"
                                    onClick={handleTriggerReEvaluation}
                                    disabled={isLoading}
                                    className="px-4 py-1.5 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded-lg transition shadow-sm flex items-center gap-1 cursor-pointer disabled:opacity-50"
                                  >
                                    <Sparkles className="w-3.5 h-3.5 animate-pulse" /> {isCouncilView ? "Bắt đầu thẩm định lại" : "Bắt đầu chấm lại"}
                                  </button>
                                </div>
                              </div>
                            </div>
                          )}

                          {/* Score breakdown sliders/bars */}
                          {currentResult.evaluationMode !== 'comment_only' && (
                            <div className="p-5 border-b border-natural-border bg-[#faf9f5] flex flex-col gap-4 shrink-0">
                              <h4 className="text-[11px] font-bold text-natural-muted uppercase tracking-wider mb-2">
                                Điểm số thành phần chấm chi tiết:
                              </h4>
                              
                              <div className="space-y-3">
                                {/* Sự cần thiết */}
                                {displayedEval.suCanThiet && (
                                  <div>
                                    <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                                      <span>Sự cần thiết của sáng kiến</span>
                                      <span className="text-natural-primary font-bold">{displayedEval.suCanThiet?.score || 0} / 10đ</span>
                                    </div>
                                    <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                                      <div className="bg-amber-600 h-full rounded-full" style={{ width: `${((displayedEval.suCanThiet?.score || 0) / 10) * 100}%` }}></div>
                                    </div>
                                  </div>
                                )}

                                {/* Tính mới */}
                                {displayedEval.tinhMoi && (
                                  <div>
                                    <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                                      <span>Tính mới, sáng tạo</span>
                                      <span className="text-natural-primary font-bold">{displayedEval.tinhMoi?.score || 0} / 20đ</span>
                                    </div>
                                    <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                                      <div className="bg-blue-600 h-full rounded-full" style={{ width: `${((displayedEval.tinhMoi?.score || 0) / 20) * 100}%` }}></div>
                                    </div>
                                  </div>
                                )}

                                {/* Nội dung và giải pháp */}
                                {displayedEval.giaiPhap && (
                                  <div>
                                    <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                                      <span>Nội dung và chất lượng giải pháp</span>
                                      <span className="text-natural-primary font-bold">{displayedEval.giaiPhap?.score || 0} / 30đ</span>
                                    </div>
                                    <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                                      <div className="bg-indigo-600 h-full rounded-full" style={{ width: `${((displayedEval.giaiPhap?.score || 0) / 30) * 100}%` }}></div>
                                    </div>
                                  </div>
                                )}
                                
                                {/* Hiệu quả áp dụng */}
                                {displayedEval.hieuQua && (
                                  <div>
                                    <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                                      <span>Hiệu quả áp dụng</span>
                                      <span className="text-natural-primary font-bold">{displayedEval.hieuQua?.score || 0} / 30đ</span>
                                    </div>
                                    <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                                      <div className="bg-emerald-600 h-full rounded-full" style={{ width: `${((displayedEval.hieuQua?.score || 0) / 30) * 100}%` }}></div>
                                    </div>
                                  </div>
                                )}

                                {/* Khả năng áp dụng */}
                                {displayedEval.khaNangApDung && (
                                  <div>
                                    <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                                      <span>Khả năng áp dụng, nhân rộng</span>
                                      <span className="text-natural-primary font-bold">{displayedEval.khaNangApDung?.score || 0} / 10đ</span>
                                    </div>
                                    <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                                      <div className="bg-orange-500 h-full rounded-full" style={{ width: `${((displayedEval.khaNangApDung?.score || 0) / 10) * 100}%` }}></div>
                                    </div>
                                  </div>
                                )}
                              </div>

                              {/* Total score box under the sliders */}
                              <div className="mt-4 pt-4 border-t border-natural-border/60 flex items-center justify-between">
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-bold text-natural-muted uppercase tracking-wider">Tổng điểm đánh giá:</span>
                                  <span className={`px-2.5 py-1 text-xs font-bold rounded-lg ${
                                    isCouncilView ? 'bg-amber-100 text-amber-800 border border-amber-200' : 'bg-natural-primary/10 text-natural-primary border border-natural-primary/20'
                                  }`}>
                                    {displayedEval.classification || 'Khá'}
                                  </span>
                                </div>
                                <div className="flex items-baseline gap-1">
                                  <span className={`text-2xl font-black ${isCouncilView ? 'text-amber-700' : 'text-red-700'}`}>
                                    {displayedEval.totalScore}
                                  </span>
                                  <span className="text-xs font-bold text-natural-muted">/ 100đ</span>
                                </div>
                              </div>
                            </div>
                          )}

                          {/* Structured Reviews */}
                          <div className="p-5 space-y-6 flex-1 overflow-y-auto max-h-[500px]">
                            
                            {plagResult && !isCouncilView && (
                              <div className="space-y-2 pt-3 border border-red-200 bg-red-50 p-4 rounded-xl mb-4">
                                <h4 className="text-[12px] font-bold text-red-800 uppercase tracking-widest flex items-center gap-1.5">
                                  <Clock className="w-4 h-4" /> KẾT QUẢ KIỂM ĐỊNH KỸ THUẬT (ĐẠO VĂN & AI)
                                </h4>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs mt-2 border-t border-red-100 pt-3">
                                   <div>
                                      <strong className="text-red-900 block mb-1">1. Tỷ lệ trùng lặp Internet: </strong>
                                      <span className="font-bold text-red-700 text-lg">{plagResult.totalDuplicatePercent}%</span>
                                      <span className="text-natural-muted italic ml-2">Cảnh báo: {plagResult.warningLevel}</span>
                                      <div className="text-[11px] text-red-600 mt-1 font-medium">(Quy định tỷ lệ trùng lặp không quá 30%)</div>
                                   </div>
                                   <div>
                                      <strong className="text-purple-900 block mb-1">2. Tỷ lệ suy đoán do AI sinh: </strong>
                                      <span className="font-bold text-purple-700 text-lg">{plagResult.aiGeneratedPercent !== undefined ? `${plagResult.aiGeneratedPercent}%` : 'Chưa phát hiện dấu hiệu'}</span>
                                   </div>
                                   <div className="col-span-1 sm:col-span-2">
                                      <strong className="text-amber-900 block mb-1">3. Lỗi chính tả & Ngữ pháp: </strong>
                                      <span className="font-bold text-amber-700">{plagResult.spellingErrors?.length || 0} lỗi phát hiện</span>
                                   </div>
                                </div>
                              </div>
                            )}

                            <div className="text-center font-bold text-natural-primary uppercase text-sm mb-4">
                              {isCouncilView ? 'PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN CỦA HỘI ĐỒNG THẨM ĐỊNH' : 'PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN'}
                            </div>

                            <div className="space-y-4">
                              <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">I. Nhận xét về nội dung sáng kiến</h3>
                              
                              {displayedEval.suCanThiet && (
                                <div className="space-y-1">
                                  <h4 className="font-semibold text-xs text-natural-primary">1. Sự cần thiết của sáng kiến</h4>
                                  <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                                    {displayedEval.suCanThiet.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                                    {displayedEval.suCanThiet.pros && displayedEval.suCanThiet.pros.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-amber-900">Đánh giá chung:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.suCanThiet.pros.map((p, i) => <li key={i}>{p}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {displayedEval.suCanThiet.levelName}</p>
                                  </div>
                                </div>
                              )}

                              {displayedEval.tinhMoi && (
                                <div className="space-y-1">
                                  <h4 className="font-semibold text-xs text-natural-primary">2. Tính mới, tính sáng tạo</h4>
                                  <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                                    {displayedEval.tinhMoi.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                                    {displayedEval.tinhMoi.pros && displayedEval.tinhMoi.pros.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-blue-900">Điểm mới thể hiện ở:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.tinhMoi.pros.map((p, i) => <li key={i}>{p}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    {displayedEval.tinhMoi.cons && displayedEval.tinhMoi.cons.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-red-900">Hạn chế:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.tinhMoi.cons.map((c, i) => <li key={i}>{c}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {displayedEval.tinhMoi.levelName}</p>
                                  </div>
                                </div>
                              )}

                              {displayedEval.giaiPhap && (
                                <div className="space-y-1">
                                  <h4 className="font-semibold text-xs text-natural-primary">3. Nội dung và chất lượng các giải pháp</h4>
                                  <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                                    {displayedEval.giaiPhap.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                                    {displayedEval.giaiPhap.pros && displayedEval.giaiPhap.pros.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-indigo-900">Nhận xét chung về hệ thống giải pháp:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.giaiPhap.pros.map((p, i) => <li key={i}>{p}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    {displayedEval.giaiPhap.cons && displayedEval.giaiPhap.cons.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-red-900">Hạn chế:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.giaiPhap.cons.map((c, i) => <li key={i}>{c}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {displayedEval.giaiPhap.levelName}</p>
                                  </div>
                                </div>
                              )}

                              {displayedEval.hieuQua && (
                                <div className="space-y-1">
                                  <h4 className="font-semibold text-xs text-natural-primary">4. Hiệu quả áp dụng</h4>
                                  <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                                    {displayedEval.hieuQua.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                                    {displayedEval.hieuQua.cons && displayedEval.hieuQua.cons.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-red-900">Hạn chế về hiệu quả:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.hieuQua.cons.map((c, i) => <li key={i}>{c}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    {displayedEval.hieuQua.levelName && <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {displayedEval.hieuQua.levelName}</p>}
                                  </div>
                                </div>
                              )}

                              {displayedEval.khaNangApDung && (
                                <div className="space-y-1">
                                  <h4 className="font-semibold text-xs text-natural-primary">5. Khả năng áp dụng và phạm vi ảnh hưởng</h4>
                                  <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                                    {displayedEval.khaNangApDung.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                                    {displayedEval.khaNangApDung.cons && displayedEval.khaNangApDung.cons.length > 0 && (
                                      <div className="mt-1">
                                        <strong className="text-[11px] text-red-900">Hạn chế về nhân rộng:</strong>
                                        <ul className="list-disc pl-4 text-[11px] text-natural-muted">
                                          {displayedEval.khaNangApDung.cons.map((c, i) => <li key={i}>{c}</li>)}
                                        </ul>
                                      </div>
                                    )}
                                    {displayedEval.khaNangApDung.levelName && <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {displayedEval.khaNangApDung.levelName}</p>}
                                  </div>
                                </div>
                              )}
                            </div>

                            {uuDiemList.length > 0 && (
                              <div className="space-y-2">
                                <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">II. Ưu điểm</h3>
                                <ul className="list-disc pl-5 text-xs text-natural-text space-y-1">
                                  {uuDiemList.map((u: string, i: number) => <li key={i}>{u}</li>)}
                                </ul>
                              </div>
                            )}

                            {hanCheList.length > 0 && (
                              <div className="space-y-2">
                                <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">III. Tồn tại, hạn chế</h3>
                                <ul className="list-disc pl-5 text-xs text-natural-text space-y-1">
                                  {hanCheList.map((u: string, i: number) => <li key={i}>{u}</li>)}
                                </ul>
                              </div>
                            )}

                            <div className="space-y-2">
                              <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">IV. Đánh giá chung</h3>
                              <p className="text-xs text-natural-text font-medium leading-relaxed italic border-l-3 border-natural-primary pl-3 bg-natural-accent p-2 rounded-r">
                                "{displayedEval.summary}"
                              </p>
                            </div>

                            {/* Display legacy improvements if available */}
                            {improvementsList.length > 0 && (
                              <div className="space-y-2 pt-3 border-t border-natural-border bg-natural-accent/60 p-4 rounded-xl outline outline-1 outline-natural-border">
                                <h4 className="text-[11px] font-bold text-natural-primary uppercase tracking-widest flex items-center gap-1.5">
                                  <Bookmark className="w-3 h-3 text-natural-primary" /> Đề xuất cải thiện:
                                </h4>
                                <ol className="list-decimal pl-5 text-xs text-natural-text space-y-1 px-1">
                                  {improvementsList.map((imp: string, idx: number) => (
                                    <li key={idx} className="leading-relaxed">{imp}</li>
                                  ))}
                                </ol>
                              </div>
                            )}

                          </div>

                          {/* Footer Action buttons */}
                          <div className="p-4 border-t border-natural-border bg-[#faf9f5] flex justify-between items-center gap-3 shrink-0">
                            <span className="text-[10px] text-natural-muted flex items-center gap-1">
                              <Clock className="w-3 h-3" /> Đã thẩm định: {new Date(displayedEval.evaluatedAt || currentResult.evaluatedAt).toLocaleTimeString('vi-VN')}
                            </span>
                            
                            <div className="flex items-center gap-2">
                              <button
                                type="button"
                                onClick={() => {
                                  setShowReAppraisalPanel(!showReAppraisalPanel);
                                  if (!showReAppraisalPanel) {
                                    setDesiredScore(displayedEval.totalScore || '');
                                  }
                                }}
                                className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold bg-[#f3f1e8] hover:bg-[#8b8b68] text-[#8b8b68] hover:text-white rounded-xl transition border border-[#8b8b68]/30 shadow-sm cursor-pointer"
                              >
                                <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
                                {showReAppraisalPanel 
                                  ? (isCouncilView ? "Đóng bảng thẩm định" : "Đóng bảng chấm lại") 
                                  : (isCouncilView ? "Thẩm định lại" : "Chấm điểm lại")
                                }
                              </button>

                              <button
                                onClick={() => setShowPrintModal(true)}
                                className="px-4 py-2 bg-natural-primary hover:bg-[#434330] border border-natural-primary hover:border-[#434330] text-[#eaeada] hover:text-white rounded-xl text-xs font-bold uppercase tracking-wider shadow hover:shadow-md transition flex items-center gap-1.5 cursor-pointer"
                              >
                                <Printer className="w-4.5 h-4.5" /> Thống kê & In phiếu
                              </button>
                            </div>
                          </div>
                        </>
                      )}
                    </div>
                  );
                })()}



                {/* (Removed old Plagiarism Analysis block) */}
            </div>
          )}
        </div>
      </section>
      )}

      {mainTab === 'settings' && (
      <section id="settings-section" className="w-full max-w-4xl mx-auto flex flex-col gap-6">
        <div className="bg-white rounded-2xl border border-natural-border shadow-sm p-6 sm:p-8">
          <div className="flex flex-col gap-2 mb-6">
            <h2 className="text-lg font-extrabold text-natural-text uppercase tracking-wider flex items-center gap-2">
              <Settings className="w-5 h-5 text-natural-primary" /> Cấu hình Hệ thống & API
            </h2>
            <p className="text-sm text-natural-muted">Thiết lập khóa API và các chế độ phân tích cho hệ thống Thẩm định Sáng kiến.</p>
          </div>

          <div className="bg-amber-50 rounded-xl border border-amber-200/60 p-5 flex flex-col gap-4">
            <div className="flex flex-col gap-2">
              <h4 className="text-[13px] font-bold text-amber-900 uppercase tracking-wider flex items-center justify-between">
                <span className="flex items-center gap-1.5"><Key className="w-5 h-5" /> Provider: Mô hình Trí tuệ Nhân tạo</span>
              </h4>
              <p className="text-[12px] text-amber-800 leading-relaxed bg-amber-100/50 p-3 rounded-lg border border-amber-200/50 mt-1">
                Lưu ý để không bị lỗi Quota Error: <strong>Bạn nên lấy API Key từ các tài khoản Google (Gmail) khác nhau hoàn toàn!</strong> Việc chọn model cho từng tác vụ đã được di chuyển sang các tab chức năng tương ứng.
              </p>
            </div>

            <div className="flex justify-between items-end mt-2 mb-1">
              <h4 className="text-xs font-bold text-amber-800 uppercase tracking-wider flex items-center gap-1.5">
                Danh sách API Keys
              </h4>
              <button
                onClick={handleAddKey}
                className="text-[11px] font-bold text-amber-800 bg-amber-200 hover:bg-amber-300 px-3 py-1.5 rounded-lg border border-amber-300 flex items-center gap-1.5 transition-colors shadow-sm"
              >
                <Plus className="w-3.5 h-3.5" /> Thêm API Key mới
              </button>
            </div>

            <div className="flex flex-col gap-3">
              {apiKeys.map((keyStr, index) => (
                <div key={index} className="flex-1 flex flex-col sm:flex-row items-center relative gap-2 sm:gap-3">
                  <div className="relative flex-1 w-full">
                    <input
                      type={visibleKeys[index] ? "text" : "password"}
                      placeholder={`Nhập API Key ${index + 1}...`}
                      value={keyStr}
                      onChange={(e) => handleKeyChange(index, e.target.value)}
                      className="w-full text-sm font-mono border border-amber-300 rounded-lg pl-3 pr-10 py-2.5 focus:outline-none focus:ring-2 focus:ring-amber-500 bg-white"
                    />
                    <button
                      type="button"
                      onClick={() => toggleKeyVisibility(index)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-amber-600 hover:text-amber-900 focus:outline-none"
                      tabIndex={-1}
                    >
                      {visibleKeys[index] ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                    </button>
                  </div>
                  {apiKeys.length > 1 && (
                    <button
                      type="button"
                      onClick={() => handleRemoveKey(index)}
                      className="p-2.5 text-red-600 bg-red-50 border border-red-100 hover:bg-red-100 hover:border-red-200 rounded-lg focus:outline-none transition-colors shrink-0 flex items-center gap-1.5 font-semibold text-[11px]"
                      title="Xóa Key này"
                    >
                      <Trash2 className="w-4 h-4" /> <span className="hidden sm:inline">Xóa</span>
                    </button>
                  )}
                </div>
              ))}
            </div>
            
            {/* Tùy chọn cấu hình thẩm định */}
            <div className="mt-8 pt-6 border-t border-amber-200/60">
              <h4 className="text-xs font-bold text-natural-primary uppercase tracking-wider mb-4 flex items-center gap-1.5">
                <Sparkles className="w-4.5 h-4.5 text-natural-secondary" /> Cấu hình Thẩm định Báo cáo
              </h4>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div className="bg-white p-4 rounded-xl border border-natural-border shadow-sm flex flex-col gap-2">
                  <label className="block text-xs font-bold text-natural-text uppercase">
                    Đại từ xưng hô trên Phiếu
                  </label>
                  <p className="text-[11px] text-natural-muted mb-2">Cách gọi tác giả trong nhận xét.</p>
                  <select
                    value={pronoun}
                    onChange={(e: any) => setPronoun(e.target.value)}
                    className="w-full text-sm border border-natural-border rounded-lg px-3 py-2 bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-natural-primary focus:border-transparent transition-all"
                  >
                    <option value="thay_co">"Thầy/Cô" (Giao tiếp thường ngày)</option>
                    <option value="tac_gia">"Tác giả" (Trang trọng, công văn)</option>
                  </select>
                </div>

                <div className="bg-white p-4 rounded-xl border border-natural-border shadow-sm flex flex-col gap-2">
                  <label className="block text-xs font-bold text-natural-text uppercase">
                    Chế độ phân tích AI
                  </label>
                  <p className="text-[11px] text-natural-muted mb-2">Kiểu đầu ra khi xuất phiếu thẩm định.</p>
                  <select
                    value={evaluationMode}
                    onChange={(e: any) => setEvaluationMode(e.target.value)}
                    className="w-full text-sm border border-natural-border rounded-lg px-3 py-2 bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-natural-primary focus:border-transparent transition-all"
                  >
                    <option value="full">Chấm điểm & Xếp loại đầy đủ</option>
                    <option value="comment_only">Chỉ Góp ý chuyên môn</option>
                  </select>
                </div>
              </div>
            </div>

          </div>
        </div>
      </section>
      )}
    </main>

      {/* FOOTER */}
      <footer className="bg-[#2c2c1e] text-natural-border/70 py-6 border-t border-natural-border/10 text-center text-xs mt-auto no-print">
        <div className="max-w-[1600px] mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>© 2026 Hệ thống hỗ trợ Sáng kiến cải tiến. Bảo lưu mọi quyền.</p>
          <div className="flex items-center gap-4 text-natural-border/40">
            <span>Phiên bản 2.5 (Tự động hóa AI)</span>
          </div>
        </div>
      </footer>

      {/* 🏛️ PRINT MODAL (Chuẩn mẫu 100% không tự sáng tạo) */}
      {showPrintModal && currentResult && (() => {
        const prCouncilName = currentResult.councilName || councilName || 'CỦA THÀNH VIÊN TỔ THẨM ĐỊNH SÁNG KIẾN';
        const prCouncilMembers: CouncilMember[] = currentResult.councilMembers || councilMembers || [];
        const prMember1Name = currentResult.member1Name || member1Name || '';
        const prMember1Unit = currentResult.member1Unit || member1Unit || '';
        const prMember1Role = currentResult.member1Role || member1Role || '';
        const prMember2Name = currentResult.member2Name || member2Name || '';
        const prMember2Unit = currentResult.member2Unit || member2Unit || '';
        const prMember2Role = currentResult.member2Role || member2Role || '';
        const prMember3Name = currentResult.member3Name || member3Name || '';
        const prMember3Unit = currentResult.member3Unit || member3Unit || '';
        const prMember3Role = currentResult.member3Role || member3Role || '';

        const hasPrCustomAppraisalMembers = prCouncilMembers.length > 0
          ? prCouncilMembers.some(m => m.name.trim() !== '')
          : Boolean(prMember1Name.trim() || prMember2Name.trim() || prMember3Name.trim());

        const isCouncilTarget = printTarget === 'council' && currentResult.councilResult;
        const evalToUse = isCouncilTarget ? currentResult.councilResult! : currentResult;

        return (
          <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex items-start justify-center overflow-y-auto p-4 md:p-8 print-force-static">
            <div className="bg-white text-black w-full max-w-4xl rounded-2xl shadow-2xl relative flex flex-col my-4 print-force-static">
            
            {/* Modal header options */}
            <div className="bg-natural-accent border-b border-natural-border px-6 py-4 flex flex-col md:flex-row items-center justify-between no-print rounded-t-2xl gap-3">
              <div className="flex flex-col md:flex-row items-start md:items-center gap-4">
                <h3 className="font-bold text-natural-primary flex items-center gap-1.5 text-sm md:text-base">
                  <Printer className="w-5 h-5 text-natural-primary" /> Phiếu Nhận Xét, Đánh Giá Sáng Kiến
                </h3>
                {currentResult.councilResult && (
                  <div className="flex items-center gap-1 bg-natural-muted/10 p-1 rounded-xl border border-natural-border">
                    <button
                      onClick={() => setPrintTarget('department')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                        printTarget === 'department'
                          ? 'bg-natural-primary text-white shadow-sm'
                          : 'text-natural-muted hover:text-natural-primary'
                      }`}
                    >
                      Sơ bộ (Phòng VH-XH)
                    </button>
                    <button
                      onClick={() => setPrintTarget('council')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                        printTarget === 'council'
                          ? 'bg-natural-primary text-white shadow-sm'
                          : 'text-natural-muted hover:text-natural-primary'
                      }`}
                    >
                      Hội đồng Thẩm định
                    </button>
                  </div>
                )}
              </div>
              <div className="flex items-center gap-2 flex-wrap justify-center">
                <button
                  onClick={() => handleDownloadWord()}
                  disabled={isGeneratingWord}
                  className="px-4 py-2 bg-[#2b579a] hover:bg-[#1a3f78] text-white rounded-xl text-xs font-bold uppercase tracking-wider shadow flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                  title="Lưu file định dạng Word (có thể chỉnh sửa)"
                >
                  <div className="flex font-serif italic w-4 h-4">W</div> {isGeneratingWord ? 'Đang tạo...' : 'Lưu Word'}
                </button>
                <button
                  onClick={() => handleDownloadPDF()}
                  disabled={isGeneratingPDF}
                  className="px-4 py-2 bg-[#d73229] hover:bg-[#b0211a] text-white rounded-xl text-xs font-bold uppercase tracking-wider shadow flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                  title="Đây là chức năng in/tải trên mọi môi trường"
                >
                  <Printer className="w-4 h-4" /> {isGeneratingPDF ? 'Đang kích hoạt...' : 'In / Lưu PDF'}
                </button>
                <button
                  onClick={() => setShowPrintModal(false)}
                  className="px-3.5 py-2 border border-natural-border hover:bg-white rounded-xl text-xs font-bold uppercase tracking-wider cursor-pointer text-natural-muted"
                >
                  Đóng
                </button>
              </div>
            </div>

            {/* Printable Area content (Strict administrative Vietnam layout - Matches PDF format exactly) */}
            <div ref={printRef} className="p-8 md:p-12 overflow-y-auto flex-1 font-serif leading-relaxed text-black max-w-[210mm] mx-auto bg-white print-force-static text-justify content-justify" style={{ fontFamily: '"Times New Roman", Times, serif', fontSize: '14pt' }}>
              
              {/* Top Banner */}
              <div className="header-banner" style={{ marginBottom: '1.5rem' }}>
                <div style={{ textAlign: 'center' }}>
                  <div style={{ fontWeight: 'bold', fontSize: '13pt' }}>CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                  <div style={{ fontWeight: 'bold', fontSize: '13pt' }}>Độc lập - Tự do - Hạnh phúc</div>
                  <div style={{ borderBottom: '1.5px solid black', width: '160px', margin: '4px auto 0' }}></div>
                </div>
                <div style={{ textAlign: 'right', fontStyle: 'italic', fontSize: '14pt', marginTop: '12px', fontWeight: 'normal' }}>
                  Hàm Yên, ngày {new Date(evalToUse.evaluatedAt || currentResult.evaluatedAt).getDate()} tháng {new Date(evalToUse.evaluatedAt || currentResult.evaluatedAt).getMonth() + 1} năm {new Date(evalToUse.evaluatedAt || currentResult.evaluatedAt).getFullYear()}
                </div>
              </div>

              {/* Title doc */}
              <div style={{ lineHeight: '150%' }}>&nbsp;</div>
              <div style={{ textAlign: 'center', marginBottom: '18pt' }}>
                {isCouncilTarget ? (
                  <>
                    <div className="font-bold uppercase tracking-tight" style={{ fontSize: '15pt', textAlign: 'center', fontWeight: 'bold' }}>
                      PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN
                    </div>
                    <div className="font-bold uppercase tracking-tight" style={{ fontSize: '14pt', textAlign: 'center', fontWeight: 'bold', marginTop: '4px' }}>
                      HỘI ĐỒNG THẨM ĐỊNH SÁNG KIẾN
                    </div>
                  </>
                ) : hasPrCustomAppraisalMembers ? (
                  <>
                    <div className="font-bold uppercase tracking-tight" style={{ fontSize: '15pt', textAlign: 'center', fontWeight: 'bold' }}>
                      PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN
                    </div>
                    <div className="font-bold uppercase tracking-tight" style={{ fontSize: '14pt', textAlign: 'center', fontWeight: 'bold', marginTop: '4px' }}>
                      {prCouncilName}
                    </div>
                  </>
                ) : (
                  <div className="font-bold uppercase tracking-tight" style={{ fontSize: '15pt', textAlign: 'center', fontWeight: 'bold' }}>
                    PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN
                  </div>
                )}
              </div>

              {/* Thông tin tác giả trực tiếp theo mẫu */}
              {hasPrCustomAppraisalMembers && !isCouncilTarget ? (
                <div style={{ marginBottom: '18pt', lineHeight: '1.6' }} className="space-y-2">
                  {prCouncilMembers.length > 0 ? (
                    <>
                      {prCouncilMembers.map((m, idx) => (
                        <div key={m.id || idx} className="space-y-1">
                          <div><strong>{idx + 1}. Họ và tên Thành viên {idx + 1}:</strong> {m.name || '...................................................'}</div>
                          <div className="pl-6">- Đơn vị công tác: {m.unit || '...................................................'}</div>
                          <div className="pl-6">- Chức vụ: {m.role || '...................................................'}</div>
                        </div>
                      ))}
                      <div className="pt-2"><strong>{prCouncilMembers.length + 1}. Tên giải pháp đề nghị công nhận sáng kiến:</strong> {currentResult.initiativeTitle ? `“${currentResult.initiativeTitle}”` : '...................................................'}</div>
                    </>
                  ) : (
                    <>
                      <div><strong>1. Họ và tên Thành viên 1:</strong> {prMember1Name || '...................................................'}</div>
                      <div className="pl-6">- Đơn vị công tác: {prMember1Unit || '...................................................'}</div>
                      <div className="pl-6">- Chức vụ: {prMember1Role || '...................................................'}</div>

                      <div><strong>2. Họ và tên Thành viên 2:</strong> {prMember2Name || '...................................................'}</div>
                      <div className="pl-6">- Đơn vị công tác: {prMember2Unit || '...................................................'}</div>
                      <div className="pl-6">- Chức vụ: {prMember2Role || '...................................................'}</div>

                      <div><strong>3. Họ và tên Thành viên 3:</strong> {prMember3Name || '...................................................'}</div>
                      <div className="pl-6">- Đơn vị công tác: {prMember3Unit || '...................................................'}</div>
                      <div className="pl-6">- Chức vụ: {prMember3Role || '...................................................'}</div>

                      <div><strong>4. Tên giải pháp đề nghị công nhận sáng kiến:</strong> {currentResult.initiativeTitle ? `“${currentResult.initiativeTitle}”` : '...................................................'}</div>
                    </>
                  )}

                  <div className="pt-3">- Họ và tên Tác giả: {currentResult.teacher.teacherName || '...................................................'}</div>
                  <div>- Chức vụ, đơn vị công tác: {currentResult.teacher.role || '............................'}, {currentResult.teacher.schoolName || '............................................'}</div>
                  <div>- Điện thoại: {currentResult.teacher.phone || '......................................'}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Email:</strong> {currentResult.teacher.email || '..........................'}</div>
                </div>
              ) : (
                <div style={{ marginBottom: '18pt', lineHeight: '1.6' }} className="space-y-1">
                  <div><strong>Họ và tên tác giả:</strong> {currentResult.teacher.teacherName || '...................................................'}</div>
                  <div><strong>Chức vụ:</strong> {currentResult.teacher.role || '...................................................'}</div>
                  <div><strong>Đơn vị công tác:</strong> {currentResult.teacher.schoolName || '...................................................'}</div>
                  <div><strong>Tên sáng kiến:</strong> {currentResult.initiativeTitle ? `“${currentResult.initiativeTitle}”` : '...................................................'}</div>
                  <div><strong>Lĩnh vực áp dụng:</strong> {currentResult.teacher.subject || 'Giáo dục mầm non'}</div>
                  <div><strong>Điện thoại:</strong> {currentResult.teacher.phone || '......................................'}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Email:</strong> {currentResult.teacher.email || '..........................'}</div>
                </div>
              )}

              {/* I. NHẬN XÉT, ĐÁNH GIÁ CHI TIẾT */}
              <div>
                <h4 className="font-bold mb-4 mt-6" style={{ fontSize: '14pt', fontWeight: 'bold' }}>I. NHẬN XÉT, ĐÁNH GIÁ CHI TIẾT</h4>

                {/* 1. Sự cần thiết */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">1. Về sự cần thiết của sáng kiến</div>
                  <div style={{ textAlign: 'justify' }} className="mb-3 leading-relaxed">
                    {evalToUse.suCanThiet?.analysis && evalToUse.suCanThiet.analysis.length > 0 
                      ? renderList(evalToUse.suCanThiet.analysis)
                      : 'Sáng kiến được lựa chọn nghiên cứu xuất phát từ yêu cầu thực tiễn của công tác giáo dục trẻ. Việc nghiên cứu và áp dụng sáng kiến là cần thiết, phù hợp với yêu cầu đổi mới giáo dục hiện nay, góp phần nâng cao chất lượng chăm sóc, giáo dục.'}
                  </div>
                  <div className="font-bold italic mb-1">Đánh giá:</div>
                  <div className="space-y-0.5 mb-2 ml-4">
                    <div>- Xác định đúng vấn đề trọng tâm.</div>
                    <div>- Phù hợp chủ trương đổi mới giáo dục.</div>
                    <div>- Có giá trị thực tiễn cao.</div>
                  </div>
                  <div className="mb-4"><strong>Điểm: {evalToUse.suCanThiet?.score || 0}/10 điểm</strong></div>
                </div>

                {/* 2. Tính mới */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">2. Về tính mới, tính sáng tạo của sáng kiến</div>
                  <div style={{ textAlign: 'justify' }} className="mb-2 leading-relaxed">
                    {evalToUse.tinhMoi?.analysis && evalToUse.tinhMoi.analysis.length > 0 
                      ? renderList(evalToUse.tinhMoi.analysis)
                      : 'Sáng kiến đã xây dựng được hệ thống các biện pháp chỉ đạo tương đối đồng bộ nhằm nâng cao chất lượng giáo dục tại đơn vị.'}
                  </div>
                  {evalToUse.tinhMoi?.cons && evalToUse.tinhMoi.cons.length > 0 && (
                    <div className="mb-2 text-justify">
                      <span className="font-bold italic">Hạn chế: </span>
                      {renderList(evalToUse.tinhMoi.cons, "")}
                    </div>
                  )}
                  <div className="font-bold italic mb-1">Đánh giá:</div>
                  <div className="space-y-0.5 mb-2 ml-4">
                    <div>- Có tính mới trong phạm vi cơ sở.</div>
                    <div>- Có sự sáng tạo trong tổ chức thực hiện.</div>
                    <div>- Tính đổi mới khá rõ nét.</div>
                  </div>
                  <div className="mb-4"><strong>Điểm: {evalToUse.tinhMoi?.score || 0}/20 điểm</strong></div>
                </div>

                {/* 3. Nội dung và giải pháp */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">3. Về nội dung và chất lượng các giải pháp</div>
                  <div style={{ textAlign: 'justify' }} className="mb-2 leading-relaxed">
                    {evalToUse.giaiPhap?.analysis && evalToUse.giaiPhap.analysis.length > 0 
                      ? renderList(evalToUse.giaiPhap.analysis)
                      : 'Tác giả đã xây dựng hệ thống giải pháp có tính logic, đồng bộ và phù hợp với mục tiêu đề ra.'}
                  </div>
                  {evalToUse.giaiPhap?.pros && evalToUse.giaiPhap.pros.length > 0 && (
                    <div className="mb-2 text-justify">
                      <span className="font-bold italic">Nhận xét chung: </span>
                      {renderList(evalToUse.giaiPhap.pros, "")}
                    </div>
                  )}
                  {evalToUse.giaiPhap?.cons && evalToUse.giaiPhap.cons.length > 0 && (
                    <div className="mb-2 text-justify">
                      <span className="font-bold italic">Hạn chế: </span>
                      {renderList(evalToUse.giaiPhap.cons, "")}
                    </div>
                  )}
                  <div className="mb-4"><strong>Điểm: {evalToUse.giaiPhap?.score || 0}/30 điểm</strong></div>
                </div>

                {/* 4. Hiệu quả áp dụng */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">4. Về hiệu quả áp dụng</div>
                  <div style={{ textAlign: 'justify' }} className="mb-2 leading-relaxed">
                    {evalToUse.hieuQua?.analysis && evalToUse.hieuQua.analysis.length > 0 
                      ? renderList(evalToUse.hieuQua.analysis)
                      : 'Sau quá trình triển khai, sáng kiến đã góp phần nâng cao năng lực chuyên môn của giáo viên, rèn luyện kỹ năng và sự tự tin cho học sinh.'}
                  </div>
                  {evalToUse.hieuQua?.cons && evalToUse.hieuQua.cons.length > 0 && (
                    <div className="mb-2 text-justify">
                      <span className="font-bold italic">Hạn chế: </span>
                      {renderList(evalToUse.hieuQua.cons, "")}
                    </div>
                  )}
                  <div className="mb-4"><strong>Điểm: {evalToUse.hieuQua?.score || 0}/30 điểm</strong></div>
                </div>

                {/* 5. Khả năng áp dụng */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">5. Về khả năng áp dụng và phạm vi ảnh hưởng</div>
                  <div style={{ textAlign: 'justify' }} className="mb-2 leading-relaxed">
                    {evalToUse.khaNangApDung?.analysis && evalToUse.khaNangApDung.analysis.length > 0 
                      ? renderList(evalToUse.khaNangApDung.analysis)
                      : 'Sáng kiến đã được triển khai hiệu quả tại nhà trường. Các biện pháp phù hợp với điều kiện thực tế, dễ thực hiện và có khả năng áp dụng tại các đơn vị tương đồng.'}
                  </div>
                  {evalToUse.khaNangApDung?.cons && evalToUse.khaNangApDung.cons.length > 0 && (
                    <div className="mb-2 text-justify">
                      <span className="font-bold italic">Hạn chế: </span>
                      {renderList(evalToUse.khaNangApDung.cons, "")}
                    </div>
                  )}
                  <div className="mb-4"><strong>Điểm: {evalToUse.khaNangApDung?.score || 0}/10 điểm</strong></div>
                </div>
              </div>

              {/* Table Scores */}
              <h4 className="font-bold mb-4 mt-6" style={{ fontSize: '14pt', fontWeight: 'bold' }}>II. TỔNG HỢP KẾT QUẢ CHẤM ĐIỂM</h4>
              <div className="mb-6">
                <table className="w-full border-collapse border border-black text-xs text-left" style={{ fontSize: '14pt' }}>
                  <thead>
                    <tr className="bg-neutral-50 text-center font-bold">
                      <th className="border border-black p-2.5" style={{ padding: '8px', border: '1px solid black' }}>Nội dung đánh giá</th>
                      <th className="border border-black p-2.5 w-32 text-center" style={{ textAlign: 'center', width: '120px', padding: '8px', border: '1px solid black' }}>Điểm tối đa</th>
                      <th className="border border-black p-2.5 w-32 text-center" style={{ textAlign: 'center', width: '120px', padding: '8px', border: '1px solid black' }}>Điểm chấm</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="border border-black p-2.5" style={{ padding: '8px', border: '1px solid black' }}>Sự cần thiết của sáng kiến</td>
                      <td className="border border-black p-2.5 text-center" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>10</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>{evalToUse.suCanThiet?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5" style={{ padding: '8px', border: '1px solid black' }}>Tính mới, tính sáng tạo</td>
                      <td className="border border-black p-2.5 text-center" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>20</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>{evalToUse.tinhMoi?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5" style={{ padding: '8px', border: '1px solid black' }}>Nội dung và giải pháp</td>
                      <td className="border border-black p-2.5 text-center" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>30</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>{evalToUse.giaiPhap?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5" style={{ padding: '8px', border: '1px solid black' }}>Hiệu quả áp dụng</td>
                      <td className="border border-black p-2.5 text-center" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>30</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>{evalToUse.hieuQua?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5" style={{ padding: '8px', border: '1px solid black' }}>Khả năng áp dụng, phạm vi ảnh hưởng</td>
                      <td className="border border-black p-2.5 text-center" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>10</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', padding: '8px', border: '1px solid black' }}>{evalToUse.khaNangApDung?.score || 0}</td>
                    </tr>
                    <tr className="font-bold" style={{ fontWeight: 'bold' }}>
                      <td className="border border-black p-2.5 font-bold text-center" style={{ textAlign: 'center', fontWeight: 'bold', padding: '8px', border: '1px solid black' }}><strong>Tổng cộng</strong></td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', fontWeight: 'bold', padding: '8px', border: '1px solid black' }}><strong>100</strong></td>
                      <td className="border border-black p-2.5 text-center font-bold text-red-900" style={{ textAlign: 'center', fontWeight: 'bold', padding: '8px', border: '1px solid black' }}><strong>{evalToUse.totalScore}</strong></td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* III. KẾT LUẬN */}
              <h4 className="font-bold mb-4 mt-6" style={{ fontSize: '14pt', fontWeight: 'bold' }}>III. KẾT LUẬN</h4>
              <div style={{ textAlign: 'justify', lineHeight: '1.6' }} className="space-y-3 mb-6">
                <div>
                  Sáng kiến: {currentResult.initiativeTitle ? `“${currentResult.initiativeTitle}”` : '...................................................'} là sáng kiến có giá trị thực tiễn, đáp ứng yêu cầu đổi mới giáo dục mầm non, góp phần nâng cao chất lượng chuẩn bị cho trẻ trước khi vào lớp 1.
                </div>
                <div>
                  Hệ thống giải pháp được xây dựng tương đối đồng bộ, phù hợp với điều kiện thực tế của đơn vị, có khả năng áp dụng rộng rãi tại các cơ sở giáo dục mầm non có điều kiện tương đồng.
                </div>
                <div>
                  Tuy nhiên, để nâng cao giá trị khoa học và sức thuyết phục của sáng kiến, tác giả cần bổ sung thêm số liệu định lượng, các bảng đối chứng trước và sau tác động, đồng thời tăng cường minh chứng về phạm vi ảnh hưởng và khả năng nhân rộng.
                </div>
                <div className="pt-2 font-bold uppercase" style={{ fontWeight: 'bold' }}>
                  Xếp loại đề nghị: {evalToUse.classification || 'LOẠI KHÁ'}
                </div>
                <div className="font-bold" style={{ fontWeight: 'bold' }}>
                  Tổng điểm: {evalToUse.totalScore}/100 điểm
                </div>
              </div>

              {/* Signatures */}
              <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '2.5rem', border: 'none' }} border={0}>
                <tbody>
                  <tr style={{ border: 'none' }}>
                    <td style={{ width: '45%', border: 'none' }}></td>
                    <td style={{ width: '55%', textAlign: 'center', verticalAlign: 'top', fontSize: '14pt', border: 'none' }}>
                      <div style={{ fontStyle: 'italic', fontSize: '14pt', marginBottom: '4px' }}>
                        Hàm Yên, ngày {new Date(evalToUse.evaluatedAt || currentResult.evaluatedAt).getDate()} tháng {new Date(evalToUse.evaluatedAt || currentResult.evaluatedAt).getMonth() + 1} năm {new Date(evalToUse.evaluatedAt || currentResult.evaluatedAt).getFullYear()}
                      </div>
                      {isCouncilTarget ? (
                        <>
                          <strong style={{ display: 'block', fontSize: '14pt', fontWeight: 'bold' }}>TM. HỘI ĐỒNG THẨM ĐỊNH SÁNG KIẾN</strong>
                          <div style={{ fontSize: '14pt', fontStyle: 'italic', marginBottom: '80px' }}>(Ký, ghi rõ họ và tên)</div>
                          <div style={{ fontWeight: 'bold' }}>ỦY VIÊN THƯ KÝ HỘI ĐỒNG</div>
                        </>
                      ) : (
                        <>
                          <strong style={{ display: 'block', fontSize: '14pt', fontWeight: 'bold' }}>NGƯỜI NHẬN XÉT, ĐÁNH GIÁ</strong>
                          <div style={{ fontSize: '14pt', fontStyle: 'italic', marginBottom: '80px' }}>(Ký, ghi rõ họ và tên)</div>
                          <div style={{ fontWeight: 'bold' }}>{reviewerName || '........................................'}</div>
                        </>
                      )}
                    </td>
                  </tr>
                </tbody>
              </table>

            </div>

            {/* Print Modal close block */}
            <div className="bg-natural-accent border-t border-natural-border p-4 flex justify-end gap-2 text-xs font-semibold no-print rounded-b-2xl">
              <button
                onClick={() => handleDownloadWord()}
                disabled={isGeneratingWord}
                className="px-5 py-2 bg-[#2b579a] hover:bg-[#1a3f78] text-white rounded-xl uppercase tracking-wider transition flex items-center gap-1 cursor-pointer disabled:opacity-50"
              >
                <div className="flex font-serif font-bold italic w-4 h-4 mr-1">W</div> {isGeneratingWord ? 'Đang tạo...' : 'Lưu dạng Word (DOC)'}
              </button>
              <button
                onClick={() => handleDownloadPDF()}
                disabled={isGeneratingPDF}
                className="px-5 py-2 bg-[#d73229] hover:bg-[#b0211a] text-white rounded-xl uppercase tracking-wider transition flex items-center gap-1 cursor-pointer disabled:opacity-50"
              >
                <Printer className="w-4 h-4" /> {isGeneratingPDF ? 'Đang kích hoạt...' : 'In / Lưu PDF'}
              </button>
              <button
                onClick={() => setShowPrintModal(false)}
                className="px-4 py-2 border border-natural-border hover:bg-white rounded-xl text-natural-primary uppercase tracking-wider transition cursor-pointer"
              >
                Đóng
              </button>
            </div>

          </div>
        </div>
      );
    })()}

    </div>
  );
}
