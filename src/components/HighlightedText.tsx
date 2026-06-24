import React, { useMemo } from 'react';
import { PlagiarismResult } from '../types';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeRaw from 'rehype-raw';

interface Props {
  text: string;
  result: PlagiarismResult | null;
}

export const HighlightedText: React.FC<Props> = ({ text, result }) => {
  const finalMarkdown = useMemo(() => {
    if (!result) return text;

    // We will build a list of parts, and iteratively split them.
    type PartType = 'normal' | 'spelling' | 'ai' | 'duplicate';
    let parts: { text: string; type: PartType; meta?: string; sourceId?: string; sourceIndex?: number }[] = [{ text, type: 'normal' }];

    // Helper to split parts by a substring
    const applySplit = (
      targetText: string, 
      type: PartType, 
      meta?: string, 
      sourceId?: string, 
      sourceIndex?: number, 
      isCaseInsensitive = false
    ) => {
      if (!targetText || targetText.length < 5) return; // avoid too short matches
      
      const newParts: typeof parts = [];
      parts.forEach(part => {
        if (part.type !== 'normal') {
          newParts.push(part);
          return;
        }

        const words = targetText.trim().split(/\s+/).filter(Boolean);
        if (words.length === 0) return;
        
        const escapeRegExp = (string: string) => string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        // Allow any whitespace sequence between words to make matching robust against newlines/tabs
        const regexStr = '(' + words.map(escapeRegExp).join('\\s+') + ')';
        const regex = new RegExp(regexStr, isCaseInsensitive ? 'gi' : 'g');
        
        const splitArr = part.text.split(regex);
        
        splitArr.forEach((s, index) => {
          if (s.length === 0) return;
          if (index % 2 === 1) {
            newParts.push({ text: s, type, meta, sourceId, sourceIndex });
          } else {
            newParts.push({ text: s, type: 'normal' });
          }
        });
      });
      parts = newParts;
    };

    // 1. Highlight duplicate segments
    if (result.segments && result.segments.length > 0) {
      result.segments.forEach(seg => {
        if (seg.isDuplicate && seg.text && seg.text.length >= 10) {
           // Find index of the source to show a number
           const sIdx = result.sources?.findIndex(s => s.id === seg.sourceId);
           applySplit(seg.text, 'duplicate', undefined, seg.sourceId, sIdx !== undefined && sIdx !== -1 ? sIdx + 1 : undefined);
        }
      });
    }

    // 2. Highlight AI segments
    if (result.aiSegments && result.aiSegments.length > 0) {
      result.aiSegments.forEach(aiSeg => {
         applySplit(aiSeg, 'ai');
      });
    }

    // 3. Highlight spelling errors
    if (result.spellingErrors && result.spellingErrors.length > 0) {
      result.spellingErrors.forEach((err, idx) => {
         if (!err.errorText || err.errorText.length < 2) return;
         applySplit(err.errorText, 'spelling', err.correction, undefined, idx + 1, true);
      });
    }

    // Convert parts into markdown string with injected HTML
    return parts.map(part => {
      // Escape HTML entities in the normal markdown text to prevent syntax conflicts? No, we shouldn't escape normal text, because we want it to parse as markdown.
      // But we shouldn't have rogue HTML in `part.text`. Assuming standard markdown.
      if (part.type === 'normal') {
        return part.text;
      }
      
      const escapeHtml = (unsafe: string) => {
        return unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
      };

      const safeText = escapeHtml(part.text);

      if (part.type === 'spelling') {
        const idxBadge = part.sourceIndex ? `<span class="inline-flex items-center justify-center h-4 w-4 bg-orange-500 text-white rounded-full text-[9px] mr-1 font-bold align-middle mx-1">${part.sourceIndex}</span>` : '';
        const metaText = escapeHtml(part.meta || '');
        return `<span id="spelling-${part.sourceIndex}" class="relative group inline-block bg-orange-100/80 text-orange-900 border-b-2 border-orange-400 font-medium px-1 rounded-sm cursor-help transition-colors hover:bg-orange-200">${idxBadge}${safeText}<span class="absolute bottom-full left-1/2 -translate-x-1/2 mb-1 hidden group-hover:block whitespace-nowrap bg-gray-900 text-white text-[11px] py-1 px-2 rounded font-bold z-20 shadow-lg">Sửa thành: <span class="text-emerald-400">${metaText}</span><svg class="absolute text-gray-900 h-2 w-full left-0 top-full" x="0px" y="0px" viewBox="0 0 255 255" xmlSpace="preserve"><polygon class="fill-current" points="0,0 127.5,127.5 255,0"/></svg></span></span>`;
      }
      if (part.type === 'ai') {
        return `<span class="inline bg-purple-100/60 text-purple-900 border-b border-purple-400 px-1 rounded-sm" title="Đoạn văn này có dấu hiệu sinh bởi AI">${safeText}</span>`;
      }
      if (part.type === 'duplicate') {
        const badge = part.sourceIndex ? `<span class="inline-flex items-center justify-center bg-red-600 text-white rounded-md text-[10px] font-bold px-1.5 py-0.5 mr-1 shadow-sm leading-none">[ ${part.sourceIndex} ]</span>` : '';
        return `<span id="duplicate-${part.sourceId}" class="inline bg-red-100/60 text-red-900 px-1 rounded-sm border-b border-red-200" title="Đoạn văn này có dấu hiệu trùng lặp">${badge}${safeText}</span>`;
      }
      return safeText;
    }).join('');
  }, [text, result]);

  if (!result) {
    return (
      <div className="preview-box p-4 h-full">
        <ReactMarkdown remarkPlugins={[remarkGfm]} rehypePlugins={[rehypeRaw]}>
          {text}
        </ReactMarkdown>
      </div>
    );
  }

  return (
    <div className="bg-white relative h-full flex flex-col">
      <div className="mb-4 sticky top-0 bg-white/90 backdrop-blur-sm p-3 border-b border-natural-border z-10 flex gap-4 text-xs font-semibold shadow-sm overflow-x-auto shrink-0">
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 bg-red-100 border border-red-300 inline-block"></span> Trùng lặp/Đạo văn</span>
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-full bg-purple-100 border border-purple-300 inline-block"></span> Khả năng do AI tạo</span>
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-full bg-orange-100 border border-orange-300 inline-block"></span> Lỗi chính tả</span>
      </div>
      
      <div className="preview-box flex-1 w-full mx-auto max-w-4xl p-6">
        <ReactMarkdown remarkPlugins={[remarkGfm]} rehypePlugins={[rehypeRaw]}>
          {finalMarkdown}
        </ReactMarkdown>
      </div>
    </div>
  );
};


