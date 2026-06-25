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
  Sparkles, 
  History, 
  ChevronRight, 
  ArrowUpRight, 
  GraduationCap, 
  RefreshCw, 
  Bookmark, 
  Trash2,
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
import { PlagiarismResult, PlagiarismSource, PlagiarismSegment, EvaluationResult, TeacherInfo } from './types';
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
    subject: ''
  });
  
  const [initiativeTitle, setInitiativeTitle] = useState('');
  const [appliedDate, setAppliedDate] = useState('2026-06-11');
  const [initiativeText, setInitiativeText] = useState('');
  
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
  
  // Active Tab for Legal Handbook / Rubric in LHS
  const [lhsTab, setLhsTab] = useState<'rubric' | 'legal' | 'xml_guide'>('rubric');
  
  // Print Mode Modal
  const [showPrintModal, setShowPrintModal] = useState(false);
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
      if (!printRef.current) return;
      
      let htmlString = printRef.current.innerHTML;

      // Replace spans with block class to divs so Word natively breaks lines
      htmlString = htmlString.replace(/<span class="block([^"]*)">([\s\S]*?)<\/span>/g, '<div class="$1">$2</div>');

      // Convert HTML structure to Blob with ms-word mime type (Works well for older Office reading)
      const msWordObj = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head><meta charset='utf-8'><title>Export</title>
      <style>
        @page WordSection1 {
          size: 8.5in 11.0in;
          margin: 1.0in 1.0in 1.0in 1.0in;
        }
        div.WordSection1 { page: WordSection1; }
        body { font-family: "Times New Roman", Times, serif; font-size: 14pt; padding: 0; margin: 0; line-height: 1.5; }
        p, div { margin-top: 0; mso-line-height-rule: auto; line-height: 150%; }
        p { margin-bottom: 6pt; }
        .space-y-1 > div, .space-y-1 > p { margin-top: 0; margin-bottom: 6pt; }
        .space-y-2 > div, .space-y-2 > p { margin-top: 0; margin-bottom: 6pt; }
        .space-y-4 > div, .space-y-4 > p { margin-top: 0; margin-bottom: 6pt; }
        table { border-collapse: collapse; width: 100%; border: 1px solid black; mso-line-height-rule: exactly; line-height: 100%; }
        .header-banner div { line-height: 1.0 !important; mso-line-height-rule: exactly !important; margin: 0 !important; }
        .content-justify, .content-justify p, .content-justify div { text-align: justify; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        .font-bold { font-weight: bold; }
        .text-center { text-align: center; }
        .uppercase { text-transform: uppercase; }
        .mb-1 { margin-bottom: 6pt; }
        .mb-2 { margin-bottom: 6pt; }
        .mb-4 { margin-bottom: 12pt; }
        .mb-6 { margin-bottom: 18pt; }
        .mb-8 { margin-bottom: 24pt; }
        .mt-1 { margin-top: 6pt; }
        .mt-2 { margin-top: 6pt; }
        .mt-4 { margin-top: 12pt; }
        .mt-6 { margin-top: 18pt; }
        .my-6 { margin-top: 18pt; margin-bottom: 18pt; }
        .block { display: block; }
        h3 { font-size: 16pt; margin: 0; padding: 0;}
        h2 { font-size: 17pt; margin: 0; padding: 0;}
        h4 { font-size: 14pt; margin: 0; padding: 0;}
        .underline { text-decoration: underline; }
        .italic { font-style: italic; }
        .text-right { text-align: right; }
        .text-justify { text-align: justify; line-height: 150%; }
        .leading-relaxed { line-height: 150%; }
        .pl-5 { padding-left: 1.5rem; }
        .ml-4 { margin-left: 1rem; }
        .ml-2 { margin-left: 0.5rem; }
        .ml-3 { margin-left: 0.75rem; }
        .p-4 { padding: 1rem; }
      </style>
      </head><body><div class="WordSection1">${htmlString}</div></body></html>`;

      const blob = new Blob(['\ufeff', msWordObj], {
        type: 'application/msword;charset=utf-8'
      });
      saveAs(blob, currentResult ? `PhieuGiamDinh_${currentResult.teacher.teacherName || 'TacGia'}.doc` : 'Phieu_Danh_Gia_SKKN.doc');
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
          const docRef = doc(db, 'users', user.uid);
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
              }
            }
            if (data.apiKeys && Array.isArray(data.apiKeys)) {
              setApiKeys(data.apiKeys);
            }
            if (data.reviewerName) {
              setReviewerName(data.reviewerName);
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
      
      const mScore = h.tinhMoi?.score || 0;
      const kScore = h.tinhKhoaHoc?.score || 0;
      const hScore = h.hieuQua?.score || 0;
      const pScore = h.phamVi?.score || 0;
      const cScore = h.minhChung?.score || 0;
      
      if (h.hieuQua && h.phamVi) {
        if (h.tinhMoi) {
          // New format
          totalScore = mScore + kScore + hScore + pScore + cScore;
        } else {
          // Legacy format
          totalScore = hScore + pScore;
        }
      }

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

      const vhxhComment = `Tổng điểm: ${totalScore}/100.\nĐánh giá sơ bộ: ${isPass || isFail}`;
      
      const isCouncil = Boolean(isCouncilAppraisal || h.isCouncilAppraisal);
      let councilOpinion = "";
      let councilPass = "";
      let councilFail = "";

      if (isCouncil) {
        councilOpinion = `Đồng ý với đánh giá sơ bộ.\nTổng điểm: ${totalScore}/100.`;
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
  const runAppraisalEvaluation = async (textToEval: string, titleToEval: string, teacherToEval: TeacherInfo, associatedPlagResult?: PlagiarismResult | null) => {
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
          evaluationMode
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
      
      const mScore = calcScore(evalData.tinhMoi, 8);
      const kScore = calcScore(evalData.tinhKhoaHoc, 12);
      const cScore = calcScore(evalData.minhChung, 12);
      const hScore = calcScore(evalData.hieuQua, 32);
      const pScore = calcScore(evalData.phamVi, 15);
      const total = mScore + kScore + cScore + hScore + pScore;
      
      let finalClass = 'Không đạt';
      // Mức xếp loại mang tính tương đối (theo user example >= 85 là Khá, có thể >= 90 là Tốt, >= 95 Xuất sắc)
      // Dựa theo thông tư 18, thường thì Khá từ 65-79, Tốt từ 80-89, Xuất sắc 90-100 (tùy địa phương) 
      // nhưng cứ xếp Khá là mốc 60-79, Tốt 80-89, Xuất sắc >= 90
      if (total >= 90 && mScore >= 8 && kScore >= 12 && hScore >= 36 && pScore >= 18) {
        finalClass = 'Xuất sắc';
      } else if (total >= 80 && mScore >= 7 && kScore >= 10 && hScore >= 32 && pScore >= 15) {
        finalClass = 'Tốt';
      } else if (total >= 50 && mScore >= 5 && kScore >= 8 && hScore >= 20 && pScore >= 10) {
        finalClass = 'Khá';
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
        tinhCapThiet: {
          levelName: evalData.tinhCapThiet?.levelName || '',
          analysis: evalData.tinhCapThiet?.analysis || []
        },
        tinhMoi: {
          score: mScore,
          levelName: evalData.tinhMoi?.levelName || '',
          analysis: evalData.tinhMoi?.analysis || [],
          pros: evalData.tinhMoi?.pros || [],
          cons: evalData.tinhMoi?.cons || []
        },
        tinhKhoaHoc: {
          score: kScore,
          levelName: evalData.tinhKhoaHoc?.levelName || '',
          analysis: evalData.tinhKhoaHoc?.analysis || [],
          pros: evalData.tinhKhoaHoc?.pros || [],
          cons: evalData.tinhKhoaHoc?.cons || []
        },
        minhChung: {
          score: cScore,
          levelName: evalData.minhChung?.levelName || '',
          analysis: evalData.minhChung?.analysis || [],
          pros: evalData.minhChung?.pros || [],
          cons: evalData.minhChung?.cons || []
        },
        hieuQua: {
          score: hScore,
          levelName: evalData.hieuQua?.levelName || '',
          pros: evalData.hieuQua?.pros || [],
          cons: evalData.hieuQua?.cons || [],
          analysis: evalData.hieuQua?.analysis || []
        },
        phamVi: {
          score: pScore,
          levelName: evalData.phamVi?.levelName || '',
          pros: evalData.phamVi?.pros || [],
          cons: evalData.phamVi?.cons || [],
          analysis: evalData.phamVi?.analysis || []
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
        plagiarismResult: associatedPlagResult || null
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
        checkedAt: new Date().toISOString()
      };
      
      setPlagResult(instantiatedPlag);
      
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
          birthYear: cleanNull(data.extractedBirthYear) || prev.birthYear
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
                      setTeacher({ teacherName: '', birthYear: '', role: '', schoolName: '', stage: 'Tiểu học', subject: '' });
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
                            <h4 className="text-xs font-bold text-natural-primary uppercase border-b border-natural-border pb-1.5">Tổng hợp quét Đạo văn, AI & Chính tả</h4>
                            
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
                                  <div key={i} className="text-xs p-3 rounded-lg bg-red-50/50 border border-red-100 flex flex-col gap-1.5">
                                    <div className="flex flex-wrap gap-1 items-center">
                                      <span className="font-semibold text-red-700 line-through bg-red-100/50 px-1 rounded">{err.errorText}</span> 
                                      <ArrowRight className="w-3 h-3 text-natural-muted mx-0.5" /> 
                                      <span className="font-semibold text-emerald-700 bg-emerald-100/50 px-1 rounded">{err.correction}</span>
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

                {mainTab === 'appraisal' && currentResult && (
                  <div className="flex-1 flex flex-col overflow-hidden">
                    {/* Score breakdown sliders/bars */}
                    {currentResult.evaluationMode !== 'comment_only' && (
                      <div className="p-5 border-b border-natural-border bg-[#faf9f5] flex flex-col gap-4 shrink-0">
                        <h4 className="text-[11px] font-bold text-natural-muted uppercase tracking-wider mb-2">
                          Điểm số thành phần chấm chi tiết:
                        </h4>
                        
                        <div className="space-y-3">
                          {/* Tính mới */}
                          <div>
                            <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                              <span>Tính mới, sáng tạo</span>
                              <span className="text-natural-primary font-bold">{currentResult.tinhMoi?.score || 0} / 10đ</span>
                            </div>
                            <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                              <div className="bg-blue-600 h-full rounded-full" style={{ width: `${((currentResult.tinhMoi?.score || 0) / 10) * 100}%` }}></div>
                            </div>
                          </div>

                          {/* Tính khoa học */}
                          <div>
                            <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                              <span>Tính khoa học, hình thức</span>
                              <span className="text-natural-primary font-bold">{currentResult.tinhKhoaHoc?.score || 0} / 15đ</span>
                            </div>
                            <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                              <div className="bg-indigo-600 h-full rounded-full" style={{ width: `${((currentResult.tinhKhoaHoc?.score || 0) / 15) * 100}%` }}></div>
                            </div>
                          </div>
                          
                          {/* Minh chứng, số liệu */}
                          {currentResult.minhChung && (
                            <div>
                              <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                                <span>Minh chứng, số liệu</span>
                                <span className="text-natural-primary font-bold">{currentResult.minhChung?.score || 0} / 15đ</span>
                              </div>
                              <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                                <div className="bg-orange-500 h-full rounded-full" style={{ width: `${((currentResult.minhChung?.score || 0) / 15) * 100}%` }}></div>
                              </div>
                            </div>
                          )}
                          
                          {/* Hiệu quả áp dụng */}
                          <div>
                            <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                              <span>Hiệu quả áp dụng</span>
                              <span className="text-natural-primary font-bold">{currentResult.hieuQua.score} / 40đ</span>
                            </div>
                            <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                              <div className="bg-amber-600 h-full rounded-full" style={{ width: `${(currentResult.hieuQua.score / 40) * 100}%` }}></div>
                            </div>
                          </div>

                          {/* Khả năng áp dụng, nhân rộng */}
                          <div>
                            <div className="flex justify-between text-[11px] font-semibold text-natural-text mb-1">
                              <span>Khả năng áp dụng, nhân rộng</span>
                              <span className="text-natural-primary font-bold">{currentResult.phamVi.score} / 20đ</span>
                            </div>
                            <div className="w-full bg-natural-accent h-2 rounded-full overflow-hidden flex border border-natural-border/30">
                              <div className="bg-emerald-600 h-full rounded-full" style={{ width: `${(currentResult.phamVi.score / 20) * 100}%` }}></div>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}

                    {/* Structured Reviews */}
                    <div className="p-5 space-y-6 flex-1 overflow-y-auto max-h-[500px]">
                      
                      {plagResult && (
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

                      <div className="text-center font-bold text-natural-primary uppercase text-sm mb-4">PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN</div>

                      <div className="space-y-4">
                        <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">I. Nhận xét về nội dung sáng kiến</h3>
                        
                        {currentResult.tinhCapThiet && (
                          <div className="space-y-1">
                            <h4 className="font-semibold text-xs text-natural-primary">1. Tính cấp thiết</h4>
                            <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                              {currentResult.tinhCapThiet.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                              <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {currentResult.tinhCapThiet.levelName}</p>
                            </div>
                          </div>
                        )}

                        {currentResult.tinhMoi && (
                          <div className="space-y-1">
                            <h4 className="font-semibold text-xs text-natural-primary">2. Tính mới, sáng tạo</h4>
                            <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                              {currentResult.tinhMoi.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                              <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {currentResult.tinhMoi.levelName}</p>
                            </div>
                          </div>
                        )}

                        <div className="space-y-1 pt-1">
                          <h4 className="font-semibold text-xs text-natural-primary">3. Tính hiệu quả áp dụng</h4>
                          <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                            {currentResult.hieuQua.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                            {currentResult.hieuQua.levelName && <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {currentResult.hieuQua.levelName}</p>}
                          </div>
                        </div>

                        <div className="space-y-1 pt-1">
                          <h4 className="font-semibold text-xs text-natural-primary">4. Khả năng áp dụng, nhân rộng</h4>
                          <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                            {currentResult.phamVi.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                            {currentResult.phamVi.levelName && <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {currentResult.phamVi.levelName}</p>}
                          </div>
                        </div>

                        {currentResult.tinhKhoaHoc && (
                          <div className="space-y-1 pt-1">
                            <h4 className="font-semibold text-xs text-natural-primary">5. Tính khoa học, hình thức</h4>
                            <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                              {currentResult.tinhKhoaHoc.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                              <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {currentResult.tinhKhoaHoc.levelName}</p>
                            </div>
                          </div>
                        )}

                        {currentResult.minhChung && (
                          <div className="space-y-1 pt-1">
                            <h4 className="font-semibold text-xs text-natural-primary">6. Minh chứng, số liệu</h4>
                            <div className="text-xs text-natural-text leading-relaxed pl-3 space-y-1">
                              {currentResult.minhChung.analysis.map((line, idx) => <p key={idx}>{line}</p>)}
                              <p className="italic font-medium text-natural-secondary pt-1">Đánh giá: {currentResult.minhChung.levelName}</p>
                            </div>
                          </div>
                        )}
                      </div>

                      {currentResult.uuDiem && currentResult.uuDiem.length > 0 && (
                        <div className="space-y-2">
                          <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">II. Ưu điểm</h3>
                          <ul className="list-disc pl-5 text-xs text-natural-text space-y-1">
                            {currentResult.uuDiem.map((u, i) => <li key={i}>{u}</li>)}
                          </ul>
                        </div>
                      )}

                      {currentResult.hanChe && currentResult.hanChe.length > 0 && (
                        <div className="space-y-2">
                          <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">III. Tồn tại, hạn chế</h3>
                          <ul className="list-disc pl-5 text-xs text-natural-text space-y-1">
                            {currentResult.hanChe.map((u, i) => <li key={i}>{u}</li>)}
                          </ul>
                        </div>
                      )}

                      <div className="space-y-2">
                        <h3 className="font-bold text-xs uppercase text-natural-text border-b border-natural-border pb-1">IV. Đánh giá chung</h3>
                        <p className="text-xs text-natural-text font-medium leading-relaxed italic border-l-3 border-natural-primary pl-3 bg-natural-accent p-2 rounded-r">
                          "{currentResult.summary}"
                        </p>
                      </div>

                      {/* Display legacy improvements if available */}
                      {currentResult.improvements && currentResult.improvements.length > 0 && (
                        <div className="space-y-2 pt-3 border-t border-natural-border bg-natural-accent/60 p-4 rounded-xl outline outline-1 outline-natural-border">
                          <h4 className="text-[11px] font-bold text-natural-primary uppercase tracking-widest flex items-center gap-1.5">
                            <Bookmark className="w-3 h-3 text-natural-primary" /> Đề xuất cải thiện:
                          </h4>
                          <ol className="list-decimal pl-5 text-xs text-natural-text space-y-1 px-1">
                            {currentResult.improvements.map((imp, idx) => (
                              <li key={idx} className="leading-relaxed">{imp}</li>
                            ))}
                          </ol>
                        </div>
                      )}

                    </div>

                    {/* Footer Action buttons */}
                    <div className="p-4 border-t border-natural-border bg-[#faf9f5] flex justify-between gap-3 shrink-0">
                      <span className="text-[10px] text-natural-muted flex items-center gap-1">
                        <Clock className="w-3 h-3" /> Đã thẩm định: {new Date(currentResult.evaluatedAt).toLocaleTimeString('vi-VN')}
                      </span>
                      
                      <button
                        onClick={() => setShowPrintModal(true)}
                        className="px-4 py-2 bg-natural-primary hover:bg-[#434330] border border-natural-primary hover:border-[#434330] text-[#eaeada] hover:text-white rounded-xl text-xs font-bold uppercase tracking-wider shadow hover:shadow-md transition flex items-center gap-1.5 cursor-pointer"
                      >
                        <Printer className="w-4.5 h-4.5" /> Thống kê & In phiếu
                      </button>
                    </div>
                  </div>
                )}



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
          <p>© 2026 Hội đồng Sáng kiến xã Hàm Yên, tỉnh Tuyên Quang. Bảo lưu mọi quyền.</p>
          <div className="flex items-center gap-4 text-natural-border/40">
            <span>Phiên bản 2.5 (Tự động hóa AI)</span>
            <span>•</span>
            <span>Quy chế Quyết định số 270/QĐ-HĐSK</span>
          </div>
        </div>
      </footer>

      {/* 🏛️ PRINT MODAL (Replicating the official paper-work format directly) */}
      {showPrintModal && currentResult && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex items-start justify-center overflow-y-auto p-4 md:p-8 print-force-static">
          <div className="bg-white text-black w-full max-w-4xl rounded-2xl shadow-2xl relative flex flex-col my-4 print-force-static">
            
            {/* Modal header options */}
            <div className="bg-natural-accent border-b border-natural-border px-6 py-4 flex flex-col md:flex-row items-center justify-between no-print rounded-t-2xl gap-3">
              <h3 className="font-bold text-natural-primary flex items-center gap-1.5 text-sm md:text-base">
                <Printer className="w-5 h-5 text-natural-primary" /> Phiếu Chấm Điểm & Đánh Giá
              </h3>
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

            {/* Printable Area content (Strict administrative Vietnam layout) */}
            <div ref={printRef} className="p-8 md:p-12 overflow-y-auto flex-1 font-serif leading-relaxed text-black max-w-[210mm] mx-auto bg-white print-force-static text-justify content-justify" style={{ fontFamily: '"Times New Roman", Times, serif', fontSize: '14pt' }}>
              
              {/* Top Banner */}
              <table className="header-banner" style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '2rem', border: 'none' }} border={0}>
                <tbody>
                  <tr style={{ border: 'none' }}>
                    <td style={{ width: '45%', textAlign: 'center', verticalAlign: 'top', padding: 0, border: 'none' }}>
                      <div style={{ fontWeight: 'bold', fontSize: '13pt' }}>UBND XÃ HÀM YÊN</div>
                      <div style={{ fontWeight: 'bold', fontSize: '13pt' }}>HỘI ĐỒNG SÁNG KIẾN</div>
                      <div style={{ borderBottom: '1px solid black', width: '120px', margin: '4px auto 0' }}></div>
                      <div style={{ fontStyle: 'italic', fontSize: '13pt', marginTop: '4px', fontWeight: 'normal' }}>
                        (Thành lập theo QĐ số 615 /QĐ-UBND<br/>ngày 04/08/2025)
                      </div>
                    </td>
                    <td style={{ width: '55%', textAlign: 'center', verticalAlign: 'top', padding: 0, border: 'none' }}>
                      <div style={{ fontWeight: 'bold', fontSize: '13pt' }}>CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                      <div style={{ fontWeight: 'bold', fontSize: '13pt' }}>Độc lập - Tự do - Hạnh phúc</div>
                      <div style={{ borderBottom: '1.5px solid black', width: '160px', margin: '4px auto 0' }}></div>
                      <div style={{ fontStyle: 'italic', fontSize: '14pt', marginTop: '12px', fontWeight: 'normal' }}>
                        Hàm Yên, ngày {new Date(currentResult.evaluatedAt).getDate()} tháng {new Date(currentResult.evaluatedAt).getMonth() + 1} năm {new Date(currentResult.evaluatedAt).getFullYear()}
                      </div>
                    </td>
                  </tr>
                </tbody>
              </table>

              {/* Title doc */}
              <div style={{ lineHeight: '150%' }}>&nbsp;</div>
              <div style={{ textAlign: 'center', marginBottom: '12pt' }}>
                <div className="font-bold uppercase tracking-tight" style={{ fontSize: '14pt', textAlign: 'center' }}>
                  PHIẾU ĐÁNH GIÁ
                </div>
                <div className="font-bold uppercase tracking-wider" style={{ fontSize: '14pt', textAlign: 'center' }}>
                  HỒ SƠ ĐỀ NGHỊ CÔNG NHẬN SÁNG KIẾN
                </div>
              </div>

              {/* I. THÔNG TIN CHUNG */}
              <div>
                <h4 className="font-bold mb-2">I. THÔNG TIN CHUNG</h4>
                <div className="space-y-1 mb-6">
                  <div><strong>Họ và tên tác giả:</strong> {currentResult.teacher.teacherName || '...................................................'}</div>
                  <div><strong>Chức vụ, đơn vị công tác:</strong> {currentResult.teacher.role ? `${currentResult.teacher.role}, ` : ''}{currentResult.teacher.schoolName || '...................................................'}</div>
                  <div><strong>Tên sáng kiến:</strong> {currentResult.initiativeTitle}</div>
                  <div><strong>Lĩnh vực áp dụng:</strong> {currentResult.teacher.subject || '...................................................'}</div>
                  <div className="mt-1">
                    <strong>Phạm vi đề nghị công nhận:</strong>&nbsp;&nbsp;&nbsp;&nbsp;
                    <span>☑ Cấp cơ sở</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <span>□ Cấp tỉnh</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <span>□ Toàn quốc</span>
                  </div>
                  <div className="mt-2"><strong>Người nhận xét:</strong> {reviewerName || '...................................................'}</div>
                  <div><strong>Ngày nhận xét, đánh giá:</strong> Ngày {new Date(currentResult.evaluatedAt).getDate()} tháng {new Date(currentResult.evaluatedAt).getMonth() + 1} năm {new Date(currentResult.evaluatedAt).getFullYear()}</div>
                </div>

                <h4 className="font-bold mb-4 mt-6">II. NHẬN XÉT, ĐÁNH GIÁ THEO CÁC TIÊU CHÍ</h4>

                {/* 1. Điểm Mới */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">1. ĐÁNH GIÁ TÍNH MỚI, TÍNH SÁNG TẠO (Tối đa 10 điểm)</div>
                  <div className="italic mb-2 font-medium">1.1. Nội dung nhận xét</div>
                  <div className="mb-2 italic">a) Mức độ mới của giải pháp</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{getRadioCheck(currentResult.tinhMoi?.score, 10) === 0 ? '☑' : '□'} Hoàn toàn mới</div>
                    <div>{getRadioCheck(currentResult.tinhMoi?.score, 10) === 1 ? '☑' : '□'} Cải tiến từ giải pháp đã có</div>
                    <div>{getRadioCheck(currentResult.tinhMoi?.score, 10) === 2 ? '☑' : '□'} Kết hợp các giải pháp cũ theo hướng mới</div>
                    <div>{getRadioCheck(currentResult.tinhMoi?.score, 10) === 3 ? '☑' : '□'} Chưa thể hiện rõ tính mới</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[60px] leading-relaxed">
                    {currentResult.tinhMoi?.analysis && currentResult.tinhMoi.analysis.length > 0 
                      ? renderList(currentResult.tinhMoi.analysis)
                      : (currentResult.tinhMoi && currentResult.tinhMoi.score >= 8.5 
                          ? 'Sáng kiến có tính mới đột phá, hoàn toàn khác biệt so với các giải pháp trước đây.' 
                          : currentResult.tinhMoi && currentResult.tinhMoi.score >= 7
                          ? 'Sáng kiến có sự cải tiến và sáng tạo dựa trên các phương pháp cũ.'
                          : 'Sáng kiến có sự kết hợp các giải pháp nhưng tính mới chưa thật sự nổi bật.')}
                  </div>

                  <div className="mb-2 italic">b) Điểm mới nổi bật của sáng kiến</div>
                  <div className="mb-2">Giải pháp đổi mới ở nội dung nào:</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.tinhMoi && currentResult.tinhMoi.score >= 5 ? '☑' : '□'} Phương pháp thực hiện</div>
                    <div>{currentResult.tinhMoi && currentResult.tinhMoi.score >= 7 ? '☑' : '□'} Quy trình thực hiện</div>
                    <div>{currentResult.tinhMoi && currentResult.tinhMoi.score >= 6 ? '☑' : '□'} Cách tổ chức</div>
                    <div>□ Công cụ hỗ trợ</div>
                    <div>□ Phương pháp quản lý</div>
                    <div>{currentResult.tinhMoi && currentResult.tinhMoi.score >= 8.5 ? '☑' : '□'} Ứng dụng công nghệ</div>
                    <div>□ Khác: ...................................................................</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét cụ thể:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {currentResult.tinhMoi?.pros && currentResult.tinhMoi.pros.length > 0 
                      ? renderList(currentResult.tinhMoi.pros)
                      : (currentResult.tinhMoi && currentResult.tinhMoi.score >= 7 
                          ? 'Sáng kiến có tính mới rõ rệt trong việc đổi mới phương pháp thực hiện và cách tổ chức.' 
                          : 'Sáng kiến có cải tiến từ giải pháp cũ nhưng chưa có điểm mới thực sự đột phá.')}
                  </div>

                  <div className="mb-2 italic">c) So sánh với giải pháp thông thường trước đây</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {currentResult.tinhMoi?.comparison && currentResult.tinhMoi.comparison.length > 0 
                      ? renderList(currentResult.tinhMoi.comparison)
                      : (currentResult.tinhMoi && currentResult.tinhMoi.score >= 7 
                          ? 'Giải pháp có nhiều điểm ưu việt, khắc phục được các hạn chế của phương pháp giáo dục truyền thống.' 
                          : 'Chưa thể hiện rõ sự vượt trội so với các phương pháp hiện hành.')}
                  </div>

                  <div className="mb-2 italic">d) Hạn chế về tính mới (nếu có)</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                     {currentResult.tinhMoi?.cons && currentResult.tinhMoi.cons.length > 0
                        ? renderList(currentResult.tinhMoi.cons)
                        : (currentResult.tinhMoi && currentResult.tinhMoi.score >= 8
                            ? ''
                            : 'Giải pháp mang tính kế thừa, thiếu các yếu tố đột phá để làm nên một mô hình hoàn toàn mới.')}
                  </div>

                  <div className="italic mb-1 font-medium">1.2. Điểm chấm</div>
                  <div className="mb-4">Điểm đề xuất: <strong>{currentResult.tinhMoi?.score || '.............'}</strong> / 10 điểm</div>
                </div>

                {/* 2. Hiệu quả */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">2. ĐÁNH GIÁ TÍNH HIỆU QUẢ (Tối đa 40 điểm)</div>
                  <div className="italic mb-2 font-medium">2.1. Hiệu quả thực tiễn</div>
                  <div className="mb-2 italic">a) Hiệu quả chuyên môn</div>
                  <div className="mb-1">Nâng cao chất lượng công việc:</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{getRadioCheck(currentResult.hieuQua?.score, 40) === 0 ? '☑' : '□'} Rõ rệt</div>
                    <div>{getRadioCheck(currentResult.hieuQua?.score, 40) === 1 ? '☑' : '□'} Khá</div>
                    <div>{getRadioCheck(currentResult.hieuQua?.score, 40) === 2 ? '☑' : '□'} Trung bình</div>
                    <div>{getRadioCheck(currentResult.hieuQua?.score, 40) === 3 ? '☑' : '□'} Chưa rõ</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[60px] leading-relaxed">
                    {currentResult.hieuQua?.analysis && currentResult.hieuQua.analysis.length > 0 
                      ? renderList(currentResult.hieuQua.analysis)
                      : (currentResult.hieuQua && currentResult.hieuQua.score >= 32 
                          ? 'Hiệu quả chuyên môn được nâng cao rõ rệt, cải thiện đáng kể chất lượng công việc.' 
                          : currentResult.hieuQua && currentResult.hieuQua.score >= 25
                          ? 'Đã mang lại hiệu quả chuyên môn khá tốt trong phạm vi áp dụng.'
                          : 'Hiệu quả chuyên môn ở mức trung bình, cần thêm thời gian để kiểm chứng.')}
                  </div>

                  <div className="mb-2 italic">b) Hiệu quả kinh tế (nếu có)</div>
                  <div className="mb-1">Tiết kiệm:</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 25 ? '☑' : '□'} Thời gian</div>
                    <div>□ Kinh phí</div>
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 28 ? '☑' : '□'} Nhân lực</div>
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 30 ? '☑' : '□'} Hồ sơ, thủ tục</div>
                    <div>□ Khác: ...................................................................</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {currentResult.hieuQua && currentResult.hieuQua.score >= 25 ? 'Sáng kiến góp phần tiết kiệm thời gian, nhân lực đáng kể khi áp dụng vào thực tế giảng dạy và quản lý.' : 'Chưa có minh chứng rõ ràng về hiệu quả kinh tế hoặc tiết kiệm chi phí.'}
                  </div>

                  <div className="mb-2 italic">c) Hiệu quả xã hội</div>
                  <div className="mb-1">Tác động tích cực đến:</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 20 ? '☑' : '□'} Người học</div>
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 24 ? '☑' : '□'} Cơ quan, đơn vị</div>
                    <div>□ Nhân dân</div>
                    <div>□ Phụ huynh</div>
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 32 ? '☑' : '□'} Cộng đồng</div>
                    <div>{currentResult.hieuQua && currentResult.hieuQua.score >= 35 ? '☑' : '□'} Chuyển đổi số</div>
                    <div>□ Cải cách hành chính</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {currentResult.hieuQua?.pros && currentResult.hieuQua.pros.length > 0 
                      ? renderList(currentResult.hieuQua.pros)
                      : (currentResult.hieuQua && currentResult.hieuQua.score >= 25 
                          ? 'Sáng kiến mang lại hiệu quả xã hội tích cực, đặc biệt đối với người học và cơ quan đơn vị.' 
                          : 'Hiệu quả xã hội ở mức khá, cần có thêm thời gian để đánh giá tác động sâu rộng hơn.')}
                  </div>

                  <div className="mb-2 italic">d) Minh chứng hiệu quả</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 10 ? '☑' : '□'} Có số liệu đối chiếu trước – sau</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 8 ? '☑' : '□'} Có bảng biểu minh chứng</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 12 ? '☑' : '□'} Có hình ảnh/video minh chứng</div>
                    <div>□ Có xác nhận của đơn vị</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score < 8 ? '☑' : '□'} Chưa đầy đủ minh chứng</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {currentResult.minhChung && currentResult.minhChung.score >= 10 ? 'Minh chứng được cung cấp tương đối đầy đủ, có bảng biểu và số liệu minh hoạ rõ ràng.' : 'Chưa cung cấp đầy đủ minh chứng thực tế, hình ảnh hoặc số liệu để tăng tính thuyết phục.'}
                  </div>

                  <div className="mb-2 italic">e) Tồn tại, hạn chế</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {(currentResult.hanChe && currentResult.hanChe.length > 0) || (currentResult.hieuQua?.cons && currentResult.hieuQua.cons.length > 0)
                      ? renderList(currentResult.hanChe || currentResult.hieuQua?.cons)
                      : (currentResult.hieuQua && currentResult.hieuQua.score >= 32
                          ? ''
                          : 'Cần có thêm những thống kê chi tiết để đánh giá định lượng được hiệu quả một cách chính xác nhất.')}
                  </div>

                  <div className="italic mb-1 font-medium">2.2. Điểm chấm</div>
                  <div className="mb-4">Điểm đề xuất: <strong>{currentResult.hieuQua?.score || '.............'}</strong> / 40 điểm</div>
                </div>

                {/* 3. Khả năng áp dụng */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">3. ĐÁNH GIÁ KHẢ NĂNG ÁP DỤNG, NHÂN RỘNG (Tối đa 20 điểm)</div>
                  <div className="italic mb-2 font-medium">3.1. Khả năng áp dụng</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 14 ? '☑' : '□'} Dễ áp dụng</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 17 ? '☑' : '□'} Có thể triển khai diện rộng</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 12 ? '☑' : '□'} Phù hợp thực tiễn cơ sở</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 15 ? '☑' : '□'} Ít kinh phí</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 16 ? '☑' : '□'} Dễ thực hiện</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score < 14 ? '☑' : '□'} Khó triển khai diện rộng</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[60px] leading-relaxed">
                    {currentResult.phamVi?.analysis && currentResult.phamVi.analysis.length > 0 
                      ? renderList(currentResult.phamVi.analysis)
                      : (currentResult.phamVi && currentResult.phamVi.score >= 15 
                          ? 'Sáng kiến dễ áp dụng, triển khai nhanh và tiết kiệm chi phí, phù hợp với thực tiễn.' 
                          : 'Khả năng triển khai diện rộng còn hạn chế, đòi hỏi điều kiện cơ sở vật chất nhất định.')}
                  </div>

                  <div className="italic mb-2 font-medium">3.2. Khả năng nhân rộng</div>
                  <div className="mb-1">Có thể áp dụng:</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 10 ? '☑' : '□'} Trong tổ chuyên môn</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 14 ? '☑' : '□'} Trong cơ quan, đơn vị</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 17 ? '☑' : '□'} Toàn ngành</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 18 ? '☑' : '□'} Liên ngành</div>
                    <div>{currentResult.phamVi && currentResult.phamVi.score >= 19 ? '☑' : '□'} Phạm vi cấp tỉnh</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                     {currentResult.phamVi?.pros && currentResult.phamVi.pros.length > 0 
                      ? renderList(currentResult.phamVi.pros) 
                      : (currentResult.phamVi && currentResult.phamVi.score >= 15 
                          ? 'Sáng kiến có khả năng áp dụng và nhân rộng cao trong toàn ngành giáo dục hoặc cấp tỉnh.' 
                          : 'Khả năng nhân rộng chủ yếu ở cấp cơ sở hoặc trong tổ chuyên môn.')}
                  </div>

                  <div className="italic mb-2 font-medium">3.3. Hạn chế trong nhân rộng</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                     {currentResult.phamVi?.cons && currentResult.phamVi.cons.length > 0
                        ? renderList(currentResult.phamVi.cons)
                        : (currentResult.phamVi && currentResult.phamVi.score >= 15
                            ? ''
                            : 'Cần sự đồng bộ về cơ sở vật chất và nhân lực để triển khai ở quy mô lớn hơn.')}
                  </div>

                  <div className="italic mb-1 font-medium">3.4. Điểm chấm</div>
                  <div className="mb-4">Điểm đề xuất: <strong>{currentResult.phamVi?.score || '.............'}</strong> / 20 điểm</div>
                </div>

                {/* 4. Hình thức trình bày */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">4. ĐÁNH GIÁ TÍNH KHOA HỌC, HÌNH THỨC TRÌNH BÀY (Tối đa 15 điểm)</div>
                  <div className="italic mb-2 font-medium">4.1. Hình thức trình bày</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 8 ? '☑' : '□'} Đúng thể thức</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 10 ? '☑' : '□'} Bố cục logic</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 12 ? '☑' : '□'} Diễn đạt rõ ràng</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 14 ? '☑' : '□'} Có phụ lục minh chứng</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 13 ? '☑' : '□'} Có số liệu thống kê</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 14 ? '☑' : '□'} Trình bày khoa học</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                    {currentResult.tinhKhoaHoc?.hinhThuc && currentResult.tinhKhoaHoc.hinhThuc.length > 0 
                      ? renderList(currentResult.tinhKhoaHoc.hinhThuc)
                      : (currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 12 
                          ? 'Báo cáo sáng kiến được trình bày đúng thể thức, bố cục logic, diễn đạt mạch lạc, rõ ràng.' 
                          : 'Hình thức trình bày đáp ứng yêu cầu cơ bản, tuy nhiên cần chú ý thêm về tính logic của bố cục.')}
                  </div>

                  <div className="italic mb-2 font-medium">4.2. Tính khoa học</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 8 ? '☑' : '□'} Có cơ sở lý luận</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 10 ? '☑' : '□'} Có cơ sở thực tiễn</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 11 ? '☑' : '□'} Có phương pháp nghiên cứu</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 12 ? '☑' : '□'} Có quy trình thực hiện</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 13 ? '☑' : '□'} Có đánh giá kết quả</div>
                    <div>{currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score < 8 ? '☑' : '□'} Chưa thể hiện rõ tính khoa học</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[60px] leading-relaxed">
                     {currentResult.tinhKhoaHoc?.analysis && currentResult.tinhKhoaHoc.analysis.length > 0 
                      ? renderList(currentResult.tinhKhoaHoc.analysis)
                      : (currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 12 
                          ? 'Sáng kiến có đầy đủ cơ sở lý luận, thực tiễn và phương pháp nghiên cứu chặt chẽ.' 
                          : 'Đã xây dựng được quy trình nhưng chưa nêu bật được cơ sở khoa học một cách thuyết phục nhất.')}
                  </div>

                  <div className="italic mb-2 font-medium">4.3. Tồn tại, hạn chế</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                     {currentResult.tinhKhoaHoc?.cons && currentResult.tinhKhoaHoc.cons.length > 0
                        ? renderList(currentResult.tinhKhoaHoc.cons)
                        : (currentResult.tinhKhoaHoc && currentResult.tinhKhoaHoc.score >= 12 
                            ? ''
                            : 'Cần trau chuốt thêm về phần đánh giá kết quả và các số liệu minh họa.')}
                  </div>

                  <div className="italic mb-1 font-medium">4.4. Điểm chấm</div>
                  <div className="mb-4">Điểm đề xuất: <strong>{currentResult.tinhKhoaHoc?.score || '.............'}</strong> / 15 điểm</div>
                </div>

                {/* 5. Minh chứng */}
                <div className="mb-6">
                  <div className="font-bold mb-2 uppercase">5. ĐÁNH GIÁ MINH CHỨNG, SỐ LIỆU (Tối đa 15 điểm)</div>
                  <div className="italic mb-2 font-medium">Nội dung đánh giá</div>
                  <div className="space-y-1 mb-2 ml-2">
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 10 ? '☑' : '□'} Có minh chứng đầy đủ</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 12 ? '☑' : '□'} Có số liệu đối chứng</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 13 ? '☑' : '□'} Có khảo sát đầu vào – đầu ra</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 14 ? '☑' : '□'} Có xác nhận thực tế</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score >= 11 ? '☑' : '□'} Có tính khách quan</div>
                    <div>{currentResult.minhChung && currentResult.minhChung.score < 10 ? '☑' : '□'} Minh chứng chưa đầy đủ</div>
                  </div>
                  <div className="mb-2 italic">Nhận xét:</div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[60px] leading-relaxed">
                     {currentResult.minhChung?.analysis && currentResult.minhChung.analysis.length > 0 
                      ? renderList(currentResult.minhChung.analysis)
                      : (currentResult.minhChung && currentResult.minhChung.score >= 12 
                          ? 'Minh chứng được sắp xếp hệ thống, có tính đối chứng rõ ràng và đảm bảo khách quan.' 
                          : 'Các minh chứng đã được đính kèm nhưng cần bổ sung thêm số liệu đánh giá đầu ra chi tiết hơn.')}
                  </div>
                  <div style={{ textAlign: 'justify' }} className="mb-4 min-h-[40px] leading-relaxed">
                     {renderList(currentResult.minhChung?.cons, "")}
                  </div>

                  <div className="italic mb-1 font-medium">Điểm chấm</div>
                  <div className="mb-4">Điểm đề xuất: <strong>{currentResult.minhChung?.score || '.............'}</strong> / 15 điểm</div>
                </div>

              {/* Table Scores */}
              <h4 className="font-bold mb-2 mt-4">III. TỔNG HỢP ĐIỂM</h4>
              <div className="mb-6">
                <table className="w-full border-collapse border border-black text-xs text-left" style={{ fontSize: '14pt' }}>
                  <thead>
                    <tr className="bg-neutral-50 text-center font-bold">
                      <th className="border border-black p-2.5 w-12 text-center" style={{ textAlign: 'center' }}>STT</th>
                      <th className="border border-black p-2.5" style={{ textAlign: 'center' }}>Nội dung đánh giá</th>
                      <th className="border border-black p-2.5 w-24 text-center" style={{ textAlign: 'center' }}>Điểm tối đa</th>
                      <th className="border border-black p-2.5 w-24 text-center" style={{ textAlign: 'center' }}>Điểm chấm</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>1</td>
                      <td className="border border-black p-2.5">Tính mới, sáng tạo</td>
                      <td className="border border-black p-2.5 text-center text-neutral-600" style={{ textAlign: 'center' }}>10đ</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>{currentResult.tinhMoi?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>2</td>
                      <td className="border border-black p-2.5">Tính hiệu quả</td>
                      <td className="border border-black p-2.5 text-center text-neutral-600" style={{ textAlign: 'center' }}>40đ</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>{currentResult.hieuQua?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>3</td>
                      <td className="border border-black p-2.5">Khả năng áp dụng, nhân rộng</td>
                      <td className="border border-black p-2.5 text-center text-neutral-600" style={{ textAlign: 'center' }}>20đ</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>{currentResult.phamVi?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>4</td>
                      <td className="border border-black p-2.5">Tính khoa học, hình thức trình bày</td>
                      <td className="border border-black p-2.5 text-center text-neutral-600" style={{ textAlign: 'center' }}>15đ</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>{currentResult.tinhKhoaHoc?.score || 0}</td>
                    </tr>
                    <tr>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>5</td>
                      <td className="border border-black p-2.5">Minh chứng, số liệu</td>
                      <td className="border border-black p-2.5 text-center text-neutral-600" style={{ textAlign: 'center' }}>15đ</td>
                      <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center' }}>{currentResult.minhChung?.score || 0}</td>
                    </tr>
                    {currentResult.evaluationMode !== 'comment_only' && (
                      <tr className="bg-neutral-50/50 font-bold" style={{ fontWeight: 'bold' }}>
                        <td className="border border-black p-2.5 text-center font-bold" colSpan={2} style={{ textAlign: 'center', fontWeight: 'bold' }}><strong>TỔNG CỘNG</strong></td>
                        <td className="border border-black p-2.5 text-center font-bold" style={{ textAlign: 'center', fontWeight: 'bold' }}><strong>100đ</strong></td>
                        <td className="border border-black p-2.5 text-center font-bold text-base underline text-red-900" style={{ textAlign: 'center', fontWeight: 'bold' }}><strong>{currentResult.totalScore}đ</strong></td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* AI Checking Data */}
              <div className="mb-6 p-4 border border-black text-sm" style={{ fontSize: '14pt' }}>
                <h4 className="font-bold underline mb-3 uppercase tracking-tight text-center">PHIẾU KIỂM ĐỊNH KỸ THUẬT: ĐẠO VĂN, NỘI DUNG AI CAO & LỖI CHÍNH TẢ</h4>
                <div className="space-y-2">
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <strong>1. Tỷ lệ trùng lặp Internet (Đạo văn): </strong>
                      <span className="font-bold text-red-700">{plagResult ? plagResult.totalDuplicatePercent : 'Không có thông tin'}%</span>
                      <div className="text-xs text-neutral-600 italic mt-0.5">Mức cảnh báo đề xuất: {plagResult ? plagResult.warningLevel : 'Chưa phân tích'}</div>
                      <div className="text-xs text-red-600 font-bold mt-1">Lưu ý: Quy định tỷ lệ trùng lặp không quá 30%.</div>
                      {plagResult && plagResult.sources && plagResult.sources.length > 0 && (
                        <div className="mt-2 text-[12px] text-gray-800">
                           <strong className="block mb-1">Nguồn vi phạm nổi bật:</strong>
                           <ul className="list-disc pl-4 space-y-1">
                              {plagResult.sources.slice(0, 3).map((src, idx) => {
                                const exactUrlOk = src.detailed_source?.exact_url && !src.detailed_source.exact_url.includes('sangkienkinhnghiem.net') && !src.detailed_source.exact_url.includes('giaoan.link');
                                const searchUrl = src.detailed_source?.search_keywords ? `https://www.google.com/search?q=${encodeURIComponent(src.detailed_source.search_keywords)}` : null;
                                let finalUrl = exactUrlOk ? src.detailed_source!.exact_url : searchUrl;
                                if (exactUrlOk && src.detailed_source?.matched_snippet) {
                                  finalUrl += `#:~:text=${encodeURIComponent(src.detailed_source.matched_snippet)}`;
                                }
                                
                                const alternativeUrls = src.detailed_source?.alternative_urls || [];
                                const allLinksForPrint = [];
                                if (finalUrl) allLinksForPrint.push({ url: finalUrl, label: 'Nguồn chính' });
                                alternativeUrls.forEach((url, i) => {
                                  if (url && typeof url === 'string' && url.trim().startsWith('http')) {
                                    allLinksForPrint.push({ url: url.trim(), label: `Minh chứng ${i + 1}` });
                                  }
                                });

                                return (
                                <li key={idx}>
                                  <span className="font-semibold">{src.detailed_source?.document_title || src.name}</span> ({src.match_percent ?? src.percent}%)
                                  {src.detailed_source?.author && !src.detailed_source.author.includes('Không xác định') && ` - Tác giả: ${src.detailed_source.author}`}
                                  {src.detailed_source?.website_name && ` - Nguồn: ${src.detailed_source.website_name}`}
                                  {allLinksForPrint.length > 0 ? (
                                    <div className="mt-1 space-y-1">
                                      {allLinksForPrint.map((lnk, i) => (
                                        <div key={i}>
                                          <span className="font-semibold text-blue-800">{lnk.label}:</span> <a href={lnk.url} target="_blank" rel="noreferrer" className="text-blue-600 underline break-all">{lnk.url}</a>
                                        </div>
                                      ))}
                                    </div>
                                  ) : (
                                    <><br/><span className="text-emerald-700 italic border border-emerald-200 bg-emerald-50 px-1 rounded inline-block mt-0.5">Nguồn: Kho tri thức Số ngành Giáo dục</span></>
                                  )}
                                  {src.detailed_source?.matched_snippet && <div className="italic text-gray-600 mt-0.5 bg-gray-50 p-1 border border-gray-200 rounded">"... {src.detailed_source.matched_snippet} ..."</div>}
                                </li>
                              )})}
                           </ul>
                        </div>
                      )}
                    </div>
                    <div>
                      <strong>2. Tỷ lệ văn bản do AI sinh (Nghi ngờ): </strong>
                      <span className="font-bold text-purple-700">{plagResult ? `${plagResult.aiGeneratedPercent || 0}%` : 'Chưa phân tích'}</span>
                      {plagResult?.aiSegments && plagResult.aiSegments.length > 0 && (
                        <div className="mt-2 text-[12px] text-purple-900 bg-purple-50 p-2 border border-purple-100 rounded">
                           <strong className="block mb-1">Dấu hiệu vi phạm:</strong>
                           <ul className="list-disc pl-4 space-y-1">
                              {plagResult.aiSegments.slice(0, 3).map((aiText, idx) => (
                                <li key={idx} className="italic line-clamp-2">"... {aiText} ..."</li>
                              ))}
                              {plagResult.aiSegments.length > 3 && (
                                <li className="font-medium list-none italic mt-1 text-purple-700">+ {plagResult.aiSegments.length - 3} đoạn khác...</li>
                              )}
                           </ul>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="mt-4">
                    <strong>3. Báo cáo lỗi chính tả / Ngữ pháp cơ bản: </strong>
                    <span className="font-bold">{plagResult?.spellingErrors ? plagResult.spellingErrors.length : 0} lỗi phát hiện</span>
                  </div>
                  {plagResult?.spellingErrors && plagResult.spellingErrors.length > 0 && (
                     <div className="mt-2 text-xs border border-gray-300 p-2 max-h-40 overflow-y-hidden line-clamp-3 italic text-neutral-600 bg-gray-50">
                        {plagResult.spellingErrors.slice(0, 3).map((err, i) => (
                           <div key={i}>- Từ sai: <span className="line-through">{err.errorText}</span> → Sửa: {err.correction} (Lý do: {err.reason})</div>
                        ))}
                        {plagResult.spellingErrors.length > 3 && <div>... và {plagResult.spellingErrors.length - 3} lỗi khác.</div>}
                     </div>
                  )}
                  <div className="text-[11px] text-gray-500 italic text-center mt-3 border-t pt-2 border-gray-200">
                    Kết quả phân tích này được trích xuất tự động qua trí tuệ nhân tạo.
                  </div>
                </div>
              </div>

              {/* Classification administrative result */}
              <h4 className="font-bold mb-2">IV. XẾP LOẠI ĐỀ NGHỊ</h4>
              <div className="bg-neutral-50 p-4 border border-black rounded text-sm mb-6 space-y-2">
                <div>
                  Căn cứ kết quả chấm điểm, đề nghị xếp loại: 
                  <span className="font-bold underline uppercase ml-2 select-all text-red-900 border border-red-900 px-3 py-1 ml-3">
                    {currentResult.classification}
                  </span>
                </div>
              </div>

              {/* Improvements */}
              <div className="mb-8">
                <h4 className="font-bold mb-2">V. NHẬN XÉT CHUNG VÀ KIẾN NGHỊ</h4>
                <div className="space-y-4">
                  <div>
                    <strong className="block mb-1">1. Ưu điểm nổi bật:</strong>
                    {currentResult.uuDiem && currentResult.uuDiem.length > 0 ? (
                      <div className="space-y-1" style={{ textAlign: 'justify' }}>
                        {currentResult.uuDiem.map((u, idx) => (
                          <div key={idx} className="ml-4">- {u.replace(/\.$/, '')}</div>
                        ))}
                      </div>
                    ) : (
                      <div className="italic text-gray-700" style={{ textAlign: 'justify' }}>Sáng kiến trình bày rõ ràng, có tính ứng dụng thực tiễn, giải quyết được một số khó khăn trong công tác chuyên môn.</div>
                    )}
                  </div>
                  <div>
                    <strong className="block mb-1">2. Tồn tại, hạn chế:</strong>
                    {(currentResult.hanChe && currentResult.hanChe.length > 0) || (currentResult.improvements && currentResult.improvements.length > 0) ? (
                      <div className="space-y-1" style={{ textAlign: 'justify' }}>
                        {(currentResult.hanChe || currentResult.improvements || []).map((imp, idx) => (
                          <div key={idx} className="ml-4">- {imp.replace(/\.$/, '')}</div>
                        ))}
                      </div>
                    ) : (
                      <div className="italic text-gray-700" style={{ textAlign: 'justify' }}>Cần tiếp tục theo dõi và bổ sung minh chứng định lượng trong quá trình triển khai thực tế.</div>
                    )}
                  </div>
                  <div className="pt-2">
                    <strong className="block mb-2">3. Kiến nghị, đề xuất:</strong>
                    <div className="flex flex-col gap-2 ml-4 mb-3">
                      <div className="block">{currentResult.classification.includes('Không đạt') ? '□' : '☑'} Đề nghị công nhận sáng kiến</div>
                      <div className="block">{'□'} Đề nghị chỉnh sửa, hoàn thiện</div>
                      <div className="block">{currentResult.classification.includes('Không đạt') ? '☑' : '□'} Đề nghị không công nhận</div>
                    </div>
                    <div className="block mt-2">
                      <strong className="mr-1">Ý kiến khác:</strong>&nbsp;
                      <span className="italic text-gray-700 text-sm">Không có.</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Signatures */}
              <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '3rem', border: 'none' }} border={0}>
                <tbody>
                  <tr style={{ border: 'none' }}>
                    <td style={{ width: '50%', border: 'none' }}></td>
                    <td style={{ width: '50%', textAlign: 'center', verticalAlign: 'top', fontSize: '14pt', border: 'none' }}>
                      <strong style={{ display: 'block' }}>NGƯỜI NHẬN XÉT, ĐÁNH GIÁ</strong>
                      <div style={{ fontSize: '14pt', fontStyle: 'italic', marginBottom: '80px' }}>(Ký và ghi rõ họ tên)</div>
                      <div style={{ fontWeight: 'bold' }}>{reviewerName || '........................................'}</div>
                    </td>
                  </tr>
                </tbody>
              </table>

            </div>

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
      )}

    </div>
  );
}
