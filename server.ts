/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from 'express';
import path from 'path';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';
import * as pdfParseModule from 'pdf-parse';

const pdfParse = (pdfParseModule as any).default || pdfParseModule;

// Load environment variables manually
dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));

// Lazy initializer for Gemini Client
let aiClient: GoogleGenAI | null = null;
function getGeminiClient(): GoogleGenAI {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      throw new Error('GEMINI_API_KEY environment variable is required');
    }
    aiClient = new GoogleGenAI({
      apiKey: key,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        }
      }
    });
  }
  return aiClient;
}

const executeWithKeys = async <T>(apiKeys: string[] | undefined, task: (ai: GoogleGenAI) => Promise<T>): Promise<T> => {
  const fallbackKey = process.env.GEMINI_API_KEY;
  const keysToTry = Array.isArray(apiKeys) ? apiKeys.filter(k => typeof k === 'string' && k.trim() !== '') : [];
  if (fallbackKey && !keysToTry.includes(fallbackKey)) {
    keysToTry.push(fallbackKey);
  }

  if (keysToTry.length === 0) {
    throw new Error('Chưa cung cấp API Key dự phòng và GEMINI_API_KEY bị thiếu trên server');
  }

  let lastError;
  for (const key of keysToTry) {
     try {
       const ai = new GoogleGenAI({
         apiKey: key.trim(),
         httpOptions: {
           headers: { 'User-Agent': 'aistudio-build' }
         }
       });
       return await task(ai);
     } catch (e: any) {
        lastError = e;
        const msg = (e.message || '').toLowerCase();
        console.error(`Error with key ending in ${key.slice(-4)}:`, e);
        
        console.log(`Key ending in ${key.slice(-4)} failed (${msg.substring(0, 50)}...). Trying next key if available...`);
        lastError = new Error(`Tất cả API Key cung cấp đều gặp lỗi hoặc hết hạn mức. Lỗi cuối: ${e.message}`);
        continue; 
     }
  }
  
  throw lastError;
}

// 1. API: Health Check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

/// 2. API: Evaluate SKKN
app.post('/api/evaluate', async (req, res) => {
  const { apiKeys, model, teacher, initiativeTitle, appliedDate, initiativeText, fileBase64, fileName, pronoun = 'thay_co', evaluationMode = 'full' } = req.body;
  const targetModel = model || 'gemini-3.1-flash-lite';

  if (!initiativeTitle && !initiativeText && !fileBase64) {
    res.status(400).json({ error: 'Nội dung sáng kiến không được để trống!' });
    return;
  }

  try {
    let pronounInstruction = "";
    let pronounLabel = "Thầy/Cô";
    if (pronoun === 'thay_co') {
      pronounLabel = "Thầy/Cô";
      pronounInstruction = `- BẮT BUỘC tất cả các câu nhận xét chi tiết, điểm mạnh, hạn chế, phân tích và đề xuất cải thiện dành cho tác giả đều phải bắt đầu bằng đại từ xưng hô là "Thầy/Cô" (Ví dụ: "Thầy/Cô đã phân tích rất rõ...", "Thầy/Cô cần bổ sung thêm...", "Thầy/Cô nên thiết kế..."). Tuyệt đối không dùng từ xưng hô khác ở đầu câu.`;
    } else {
      pronounLabel = "Tác giả";
      pronounInstruction = `- BẮT BUỘC tất cả các câu nhận xét chi tiết, điểm mạnh, hạn chế, phân tích và đề xuất cải thiện dành cho tác giả đều phải bắt đầu bằng đại từ xưng hô là "Tác giả" (Ví dụ: "Tác giả đã thể hiện rõ...", "Tác giả cần bổ sung thêm...", "Tác giả nên thiết kế..."). Tuyệt đối không dùng từ xưng hô khác ở đầu câu.`;
    }

    const systemInstruction = `Ý THỨC VAI TRÒ & NGỮ CẢNH CẤP XÃ:
Bạn là Giám khảo thuộc Hội đồng Sáng kiến xã Hàm Yên, tỉnh Tuyên Quang. Nhiệm vụ của bạn là thẩm định và viết lời nhận xét chuyên môn CHI TIẾT, KHOA HỌC (PHIẾU NHẬN XÉT, ĐÁNH GIÁ SÁNG KIẾN) cho Báo cáo Sáng kiến kinh nghiệm (SKKN) giáo dục cấp cơ sở. ĐẶC BIỆT LƯU Ý, TRONG ĐÁNH GIÁ PHẢI BÁM SÁT TIÊU CHUẨN CÔNG NHẬN SÁNG KIẾN THEO NGHỊ ĐỊNH SỐ 13/2012/NĐ-CP.

QUY ĐỊNH BAREM ĐÁNH GIÁ (Thang điểm 100):
Phân tích và nhận xét văn bản dựa trên các tiêu chí để đưa ra điểm số phù hợp (khắt khe, hiếm khi cho điểm tuyệt đối), phải tham chiếu các điều kiện của Nghị định 13/2012/NĐ-CP:
1. Tính cấp thiết (Không chấm điểm, chỉ nhận xét): Sự cần thiết của đề tài. Nhận diện vấn đề thực tiễn.
2. Tính mới, sáng tạo (10 điểm - Tương ứng Khoản 1 Điều 4, Nghị định 13/2012/NĐ-CP): Điểm mới, sự khác biệt. Không trùng lặp, chưa bị bộc lộ công khai, chưa từng được áp dụng bởi người khác. Trừ điểm nếu chỉ lập lại các giải pháp cũ hoặc chỉ là thực hiện chức năng nhiệm vụ thông thường. (Mức điểm phổ biến: 7-9)
3. Tính khoa học, hình thức (15 điểm): Trình bày logic, đúng thể thức, có cơ sở lý luận và thực tiễn, diễn đạt rõ ràng. (Mức điểm phổ biến: 10-13)
4. Minh chứng, số liệu (15 điểm): Đánh giá minh chứng rõ ràng, sự đầy đủ của khảo sát đầu vào - đầu ra, số liệu đối chứng, và tính khách quan. Trừ điểm nếu chỉ nhận xét định tính mà không có số liệu cụ thể. (Mức điểm phổ biến: 10-13)
5. Tính hiệu quả áp dụng (40 điểm - Tương ứng Khoản 2 Điều 4, Nghị định 13/2012/NĐ-CP): Hiệu quả thực tiễn chuyên môn, hiệu quả kinh tế hoặc xã hội. Trừ nghiêm khắc nếu hiệu quả không rõ rệt. (Mức điểm phổ biến: 32-36)
6. Khả năng áp dụng, nhân rộng (20 điểm): Phạm vi lan tỏa, tính khả thi áp dụng thực tiễn trong ngành học, liên ngành, hoặc quy mô lớn hơn (trường, tỉnh). (Mức điểm phổ biến: 15-18)

QUY ĐỊNH BẮT BUỘC VỀ VĂN PHONG NHẬN XÉT:
- BẮT BUỘC: ĐỂ TRÁNH DÀI DÒNG KHI XUẤT RA WORD, các mảng nhận xét (analysis, pros, cons, comparison, hinhThuc...) BẮT BUỘC PHẢI TÁCH Ý NHỎ. Tức là mỗi ý tưởng / mỗi câu nhận xét nên là 1 phần tử riêng biệt trong Array. KHÔNG viết một đoạn văn dài gộp nhiều câu và nhiều dấu chấm (.) vào 1 phần tử Array duy nhất. Hãy chia nhỏ ra để xuống dòng hợp lý.
- KHÔNG tự ý thêm ký tự "- " ở đầu nội dung của mỗi phần tử, chỉ cần ngắt ý và trả về mảng chuỗi.
- Văn phong đánh giá phải mang tính sư phạm, pháp lý (có dẫn chứng tiêu chuẩn NĐ 13/2012/NĐ-CP khi phù hợp), lập luận khách quan, khoa học, chỉ ra rõ ràng cả ƯU ĐIỂM và HẠN CHẾ.
- Khắt khe hơn trong đánh giá, không chấm điểm quá cao (hướng tới mức 82-88 điểm thay vì 90+ điểm với các SKKN bình thường). Với các SKKN thiếu bảng biểu số liệu định lượng, tổng điểm chỉ nên dao động từ 85 - 87 điểm.
- KHÔNG nói chung chung, phải phân tích sâu vào nội dung và biện pháp cụ thể của sáng kiến.
- Đảm bảo có mục nhận xét hạn chế và đề xuất khắc phục cụ thể.
- Việc tính tổng điểm tự tính dựa trên tổng các điểm thành phần.
${pronounInstruction}

Hãy phân tích toàn văn Báo cáo Sáng kiến và trả về cấu trúc JSON đúng theo Schema gồm các tiêu chí trên.`;

    const modelPrompt = `Dưới đây là thông tin và báo cáo sáng kiến cần thẩm định:
Họ tên giáo viên: ${teacher?.teacherName || 'Chưa rõ'}
Năm sinh: ${teacher?.birthYear || 'Chưa rõ'}
Chức vụ: ${teacher?.role || 'Chưa rõ'}
Đơn vị công tác: ${teacher?.schoolName || 'Chưa rõ'}
Cấp học: ${teacher?.stage || 'Chưa rõ'}
Bộ môn: ${teacher?.subject || 'Chưa rõ'}

Tiêu đề sáng kiến: "${initiativeTitle}"
Ngày áp dụng: ${appliedDate || 'Chưa rõ'}

CHẾ ĐỘ REVIEW ĐƯỢC CHỌN:
- Ngôi xưng hô: ${pronounLabel}
- Chế độ đánh giá: ${evaluationMode === 'comment_only' ? 'Chỉ nhận xét chuyên môn (không đặt nặng tính phân loại điểm số)' : 'Thẩm định chấm điểm và xếp loại đầy đủ'}

NỘI DUNG SÁNG KIẾN ĐỂ THAM KHẢO (Nếu có file đính kèm, AI phải ƯU TIÊN đọc nội dung trực tiếp qua file đính kèm để trích xuất cả thông tin, hình ảnh, bảng biểu cho thật chuẩn xác):
"""
${initiativeText}
"""`;

    const contents = [];
    contents.push({ text: modelPrompt });
    
    if (fileBase64) {
      const base64Data = fileBase64.replace(/^data:[^;]+;base64,/, '');
      contents.push({
        inlineData: {
          data: base64Data,
          mimeType: 'application/pdf'
        }
      });
    }

    const parsedResult = await executeWithKeys(apiKeys, async (ai) => {
      const response = await ai.models.generateContent({
        model: targetModel,
        contents: contents,
        config: {
          systemInstruction,
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              tinhCapThiet: {
                type: Type.OBJECT,
                properties: {
                  levelName: { type: Type.STRING, description: "Đánh giá ngắn (Ví dụ: Đạt yêu cầu, có tính thực tiễn cao.)" },
                  analysis: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Nhận xét chi tiết về tính cấp thiết." }
                },
                required: ['levelName', 'analysis']
              },
              tinhMoi: {
                type: Type.OBJECT,
                properties: {
                  score: { type: Type.INTEGER, description: "Điểm (0-10)" },
                  levelName: { type: Type.STRING, description: "Đánh giá ngắn (Ví dụ: Có tính cải tiến và vận dụng...)" },
                  analysis: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Nhận xét chi tiết về tính mới." },
                  pros: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Ưu điểm về tính mới, tính sáng tạo" },
                  cons: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Hạn chế về tính mới, tính sáng tạo" },
                  comparison: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Nhận xét so sánh với các giải pháp thông thường trước đây" }
                },
                required: ['score', 'levelName', 'analysis']
              },
              tinhKhoaHoc: {
                type: Type.OBJECT,
                properties: {
                  score: { type: Type.INTEGER, description: "Điểm (0-15)" },
                  levelName: { type: Type.STRING, description: "Đánh giá ngắn" },
                  analysis: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Nhận xét chi tiết tính khoa học." },
                  hinhThuc: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Nhận xét chi tiết về hình thức trình bày (Đúng thể thức, bố cục, diễn đạt...)" },
                  pros: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Ưu điểm về khoa học và hình thức" },
                  cons: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Hạn chế về khoa học và hình thức" }
                },
                required: ['score', 'levelName', 'analysis']
              },
              minhChung: {
                type: Type.OBJECT,
                properties: {
                  score: { type: Type.INTEGER, description: "Điểm (0-15)" },
                  levelName: { type: Type.STRING, description: "Đánh giá ngắn" },
                  analysis: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Nhận xét chi tiết về minh chứng, số liệu." },
                  pros: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Ưu điểm về minh chứng, số liệu" },
                  cons: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Hạn chế về minh chứng, số liệu" }
                },
                required: ['score', 'levelName', 'analysis']
              },
              hieuQua: {
                type: Type.OBJECT,
                properties: {
                  score: { type: Type.INTEGER, description: "Điểm (0-40)" },
                  levelName: { type: Type.STRING, description: "Đánh giá ngắn" },
                  analysis: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Phân tích cực kỳ chi tiết các kết quả đạt được." },
                  pros: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Ưu điểm về hiệu quả áp dụng" },
                  cons: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Hạn chế về hiệu quả áp dụng" }
                },
                required: ['score', 'levelName', 'analysis']
              },
              phamVi: {
                type: Type.OBJECT,
                properties: {
                  score: { type: Type.INTEGER, description: "Điểm (0-20)" },
                  levelName: { type: Type.STRING, description: "Đánh giá ngắn" },
                  analysis: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Phân tích khả năng nhân rộng." },
                  pros: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Ưu điểm về phạm vi ảnh hưởng" },
                  cons: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Hạn chế về phạm vi ảnh hưởng" }
                },
                required: ['score', 'levelName', 'analysis']
              },
              uuDiem: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
                description: "Các ý ưu điểm tổng thể (II. Ưu điểm)"
              },
              hanChe: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
                description: "Các ý tồn tại, hạn chế (III. Tồn tại, hạn chế)"
              },
              summary: {
                type: Type.STRING,
                description: "Đánh giá chung (IV. Đánh giá chung)"
              }
            },
            required: ['tinhCapThiet', 'tinhMoi', 'tinhKhoaHoc', 'minhChung', 'hieuQua', 'phamVi', 'uuDiem', 'hanChe', 'summary']
          }
        }
      });

      let resultText = response.text;
      if (!resultText) {
        throw new Error('Mô hình Gemini không phản hồi văn bản hợp lệ.');
      }
      resultText = resultText.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');
      return JSON.parse(resultText.normalize('NFC'));
    });

    res.json(parsedResult);
  } catch (error: any) {
    console.error('Error during evaluation:', error);
    res.status(500).json({ 
      error: 'Quá trình thẩm định gặp lỗi trên hệ thống trí tuệ nhân tạo!', 
      details: error.message 
    });
  }
});

// 3. API: Extract Info (Fast)
app.post('/api/extract-info', async (req, res) => {
  const { apiKeys, model, fileBase64, fileName, text } = req.body;
  const targetModel = model || 'gemini-3.1-flash-lite';
  if (!fileBase64 && !text) {
    res.status(400).json({ error: 'Missing file/text for extraction' });
    return;
  }

  try {
    const systemInstruction = `Bạn là Trợ lý số trích xuất thông tin hành chính.
Nhiệm vụ: Trích xuất các thông tin cơ bản ở phần đầu báo cáo sáng kiến (thường ở trang bìa hoặc các trang đầu tiên). Trả về dưới dạng JSON.
Lưu ý quan trọng: Tên tác giả (người viết), chức vụ, đơn vị công tác và tên sáng kiến thường có thể nằm bên trong các Bảng (Table). Hãy đọc thật kỹ nội dung các cột "Họ và tên", "Chức danh", "Nơi công tác" để lấy chính xác thông tin.`;

    const contents = [];
    if (fileBase64) {
      const base64Data = fileBase64.replace(/^data:[^;]+;base64,/, '');
      contents.push({
        inlineData: { data: base64Data, mimeType: 'application/pdf' }
      });
      contents.push({ text: `Trích xuất thông tin hành chính từ trang bìa của file PDF "${fileName || 'skkn.pdf'}".` });
    } else {
      contents.push({ text: `Trích xuất thông tin hành chính từ văn bản bên dưới:\n\n"""\n${text.substring(0, 10000)}\n"""` });
    }

    const parsedResult = await executeWithKeys(apiKeys, async (ai) => {
      const response = await ai.models.generateContent({
        model: targetModel,
        contents: contents,
        config: {
          systemInstruction,
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              extractedTitle: { type: Type.STRING, description: "Tiêu đề của báo cáo sáng kiến kinh nghiệm." },
              extractedAuthor: { type: Type.STRING, description: "Họ tên tác giả giáo viên." },
              extractedSchool: { type: Type.STRING, description: "Tên đơn vị trường học công tác." },
              extractedStage: { type: Type.STRING, description: "Cấp học (ví dụ: Mầm non, Tiểu học, THCS, THPT). Suy luận từ tên trường nếu có." },
              extractedSubject: { type: Type.STRING, description: "Tên hệ học, môn học hoặc chuyên ngành." },
              extractedRole: { type: Type.STRING, description: "Chức vụ (vd: Giáo viên, Hiệu phó) hoặc nhiệm vụ giảng dạy." },
              extractedBirthYear: { type: Type.STRING, description: "Năm sinh tác giả (nếu có)." }
            }
          }
        }
      });

      let resultText = response.text;
      if (!resultText) {
        throw new Error('No valid text returned');
      }

      resultText = resultText.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');
      return JSON.parse(resultText.normalize('NFC'));
    });

    console.log("Extracted info successfully:", parsedResult);
    res.json(parsedResult);
  } catch (error: any) {
    console.error('Error extracting info:', error);
    res.status(500).json({ error: 'Lỗi trích xuất thông tin', details: error.message });
  }
});

// 4. API: Plagiarism and Duplicate Check
app.post('/api/plagiarism-check', async (req, res) => {
  const { apiKeys, model, text, fileBase64, fileName } = req.body;
  const targetModel = model || 'gemini-3.1-flash-lite';

  if (!text && !fileBase64) {
    res.status(400).json({ error: 'Vui lòng cung cấp nội dung văn bản hoặc tải lên file PDF sáng kiến!' });
    return;
  }

  try {
    const systemInstruction = `Bạn là Trợ lý số Quét và Kiểm tra Trùng lặp Sáng kiến Kinh nghiệm thuộc Hội đồng Sáng kiến xã Hàm Yên, Tuyên Quang.
Nhiệm vụ của bạn là:
1. Rà soát trùng lặp, đạo văn.
2. Kiểm tra lỗi chính tả trong văn bản (bỏ qua các lỗi đặc thù theo chuyên môn, thuật ngữ chuyên ngành).
3. Đánh giá tỷ lệ nội dung do AI tạo ra (AI-generated content).

HƯỚNG DẪN HOẠT ĐỘNG:
1. Tách văn bản thành chuỗi phân đoạn liên tục.
2. Rà soát từng phân đoạn trùng lặp, bỏ qua lý luận chung. Đánh dấu isDuplicate: true nếu có trùng lặp.
3. CHỈ ĐÁNH DẤU LÀ ĐẠO VĂN (isDuplicate: true) NẾU PHÁT HIỆN MỘT ĐOẠN VĂN DÀI (Ít nhất 30 từ liên tục hoặc 2-3 câu) ĐƯỢC SAO CHÉP Y HỆT từ một tài liệu cụ thể. BỎ QUA hoàn toàn các câu ngắn (dưới 1 dòng), các định nghĩa, lý luận chung chung, các câu tiêu đề, mục tiêu giáo dục thông thường, tên đề tài, các cụm từ chuyên môn sư phạm quen thuộc. TUYỆT ĐỐI KHÔNG BẮT LỖI TRÙNG LẶP CHO CÁC ĐOẠN NGẮN mang tính chất giới thiệu hoặc mục đích chung. Phải khắt khe trong việc đánh giá đạo văn, tránh bắt nhầm các câu phổ biến.
4. TÌM RA THÔNG TIN CỤ THỂ của các tài liệu gốc (Tên học liệu, phần nội dung của tài liệu gốc bị Đạo văn dán vào \`matched_snippet\`). Chỉ trích dẫn \`matched_snippet\` nếu nó thực sự là một đoạn dài và sao chép y hệt.
5. Kiểm tra lỗi chính tả: Ghi nhận các lỗi ngữ pháp, lỗi đánh máy rõ ràng. KHÔNG tính các từ ngữ chuyên môn (ví dụ trong Tin học, Toán học...).
6. Phát hiện AI CHUYÊN SÂU: Quét và phân tích kỹ lưỡng từng câu văn để phát hiện các dấu hiệu sinh tự động (AI-generated). Nhận diện các đoạn văn có cấu trúc khuôn mẫu, sáo rỗng, lặp ý, sử dụng từ ngữ đao to búa lớn đặc trưng của AI (ChatGPT, Gemini, Claude...). Ước lượng tỷ lệ % nội dung được viết bởi AI một cách khắt khe. Nếu đoạn nào bị nghi ngờ cao do AI viết, trích xuất chính xác vào mảng \`aiSegments\`.
7. Đánh giá Mức cảnh báo (warningLevel): Đánh giá mức độ vi phạm dựa trên tỷ lệ trùng lặp. Lưu ý quy định: Tỷ lệ trùng lặp không được vượt quá 30%. Nếu tổng tỷ lệ trùng lặp > 30%, ghi "Vi phạm quy định (trên 30%)". Nếu <= 30%, ghi "Đạt yêu cầu (dưới 30%)".

Bắt buộc trả về đúng định dạng JSON có cấu trúc sau:
{
  "totalDuplicatePercent": 24.5,
  "aiGeneratedPercent": 15.0,
  "warningLevel": "Vi phạm nhẹ",
  "extractedTitle": "...",
  "extractedAuthor": "...",
  "extractedSchool": "...",
  "extractedStage": "...",
  "extractedSubject": "...",
  "extractedRole": "...",
  "extractedBirthYear": "...",
  "sources": [{
    "id": "src1", 
    "name": "Giáo án Hoạt động trải nghiệm...", 
    "percent": 12.5, 
    "wordsCount": 100, 
    "url": "https://...", 
    "color": "red",
    "is_matched": true,
    "match_percent": 12.5,
    "detailed_source": {
      "document_title": "Tên sách giáo khoa / Tên sáng kiến gốc / Tên bài báo gốc...",
      "author": "NẾU TÌM ĐƯỢC TÊN TÁC GIẢ CHÍNH XÁC thì ghi. NẾU KHÔNG THÌ GHI LÀ 'Không xác định', tuyệt đối không được tự bịa ra tin tác giả.",
      "exact_url": "BẮT BUỘC TRẢ VỀ ĐƯỜNG DẪN GOOGLE SEARCH TÌM KIẾM CHÍNH XÁC MỘT ĐOẠN VĂN DÀI TRONG NGOẶC KÉP. Mẫu: https://www.google.com/search?q=\\"Trích một đoạn văn dài ít nhất 15 từ từ nội dung gửi vào để minh chứng trùng lặp\\".",
      "matched_snippet": "Bắt buộc trích ĐÚNG NGUYÊN VĂN một đoạn dài ít nhất 15 từ từ bài viết được kiểm tra mà bạn cho là đạo văn. Phải giống y hệt từng chữ để người dùng đối chiếu."
    }
  }],
  "segments": [{"text": "...", "isDuplicate": true, "sourceId": "src1", "type": "red"}],
  "spellingErrors": [{
    "errorText": "sáng kiến", 
    "correction": "sáng tạo", 
    "context": "Đây là một sáng kiến hay...", 
    "reason": "Lỗi đánh máy"
  }],
  "aiSegments": [
    "Việc áp dụng các phương pháp giáo dục hiện đại...",
    "Bên cạnh đó, việc đổi mới phương pháp giảng dạy..."
  ]
}
Lưu ý: Chỉ trả về JSON, KHÔNG BỔ SUNG BẤT KỲ VĂN BẢN NÀO KHÁC. Đảm bảo bóc tách đoạn văn nghi ngờ AI tạo vào mảng aiSegments nếu aiGeneratedPercent > 0.`;

    let finalExtractedText = text || '';
    const contents = [];
    if (fileBase64) {
      const base64Data = fileBase64.replace(/^data:[^;]+;base64,/, '');
      try {
        const buffer = Buffer.from(base64Data, 'base64');
        const pdfData = await pdfParse(buffer);
        finalExtractedText = pdfData.text;
      } catch (e) {
        console.error('Error parsing PDF:', e);
      }
      
      contents.push({
        inlineData: {
          data: base64Data,
          mimeType: 'application/pdf'
        }
      });
      contents.push({
        text: `Văn bản được trích xuất từ PDF:\n\n"""\n${finalExtractedText}\n"""\n\nHãy phân tích văn bản sáng kiến kinh nghiệm này, đối chiếu trùng lặp chi tiết từng phân đoạn và trả về kết quả quét internet.`
      });
    } else {
      contents.push({
        text: `Hãy phân tích văn bản sáng kiến kinh nghiệm sau đây, đối chiếu trùng lặp chi tiết từng phân đoạn và trả về kết quả quét internet:
        
"""
${text}
"""`
      });
    }

    const parsedResult = await executeWithKeys(apiKeys, async (ai) => {
      const response = await ai.models.generateContent({
        model: targetModel,
        contents: contents,
        config: {
          systemInstruction: systemInstruction + `\n\nBắt buộc trả về đúng định dạng JSON có cấu trúc sau:
{
  "totalDuplicatePercent": 24.5,
  "warningLevel": "Vi phạm nhẹ",
  "extractedTitle": "...",
  "extractedAuthor": "...",
  "extractedSchool": "...",
  "extractedStage": "...",
  "extractedSubject": "...",
  "extractedRole": "...",
  "extractedBirthYear": "...",
  "sources": [{
    "id": "src1", 
    "name": "Giáo án Hoạt động trải nghiệm hướng nghiệp 8...", 
    "percent": 12.5, 
    "wordsCount": 100, 
    "url": "...", 
    "color": "red",
    "is_matched": true,
    "match_percent": 12.5,
    "detailed_source": {
      "document_title": "Giáo án Hoạt động trải nghiệm hướng nghiệp 8...",
      "author": "NẾU TÌM ĐƯỢC TÊN TÁC GIẢ CHÍNH XÁC thì ghi. NẾU KHÔNG THÌ GHI LÀ 'Không xác định', tuyệt đối không được tự bịa ra tin tác giả.",
      "exact_url": "BẮT BUỘC TRẢ VỀ ĐƯỜNG DẪN GOOGLE SEARCH TÌM KIẾM CHÍNH XÁC MỘT ĐOẠN VĂN DÀI TRONG NGOẶC KÉP. Mẫu: https://www.google.com/search?q=\\"Trích một đoạn văn dài ít nhất 15 từ từ nội dung gửi vào để minh chứng trùng lặp\\".",
      "matched_snippet": "Bắt buộc trích ĐÚNG NGUYÊN VĂN một đoạn dài ít nhất 15 từ từ bài viết được kiểm tra mà bạn cho là đạo văn. Phải giống y hệt từng chữ để người dùng đối chiếu."
    }
  }],
  "segments": [{"text": "...", "isDuplicate": true, "sourceId": "src1", "type": "red"}],
  "spellingErrors": [{
    "errorText": "sáng kiến", 
    "correction": "sáng tạo", 
    "context": "Đây là một sáng kiến hay...", 
    "reason": "Lỗi đánh máy"
  }],
  "aiSegments": [
    "Việc áp dụng các phương pháp giáo dục hiện đại...",
    "Bên cạnh đó, việc đổi mới phương pháp giảng dạy..."
  ]
}
Lưu ý: Chỉ trả về JSON, KHÔNG BỔ SUNG BẤT KỲ VĂN BẢN NÀO KHÁC. Đảm bảo bóc tách đoạn văn nghi ngờ AI tạo vào mảng aiSegments nếu aiGeneratedPercent > 0.`,
          responseMimeType: 'application/json'
        }
      });

      let resultText = response.text;
      if (!resultText) {
        throw new Error('Hệ thống AI không phản hồi kết quả rà soát trùng lặp hợp lệ.');
      }

      resultText = resultText.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');
      const parsedData = JSON.parse(resultText.normalize('NFC'));
      if (finalExtractedText && finalExtractedText.length > 50) {
        parsedData.extractedText = finalExtractedText;
      }
      return parsedData;
    });

    res.json(parsedResult);
  } catch (error: any) {
    console.error('Error during plagiarism check:', error);
    res.status(500).json({ 
      error: 'Hệ thống rà soát trùng lặp gặp lỗi kỹ thuật!', 
      details: error.message 
    });
  }
});


// Setup Vite Dev server integration
async function startServer() {
  if (process.env.NODE_ENV !== 'production' && !process.env.VERCEL) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else if (!process.env.VERCEL) {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  if (!process.env.VERCEL) {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`[Hàm Yên SDK] Server is running on http://0.0.0.0:${PORT}`);
    });
  }
}

if (!process.env.VERCEL) {
  startServer();
}

export default app;
