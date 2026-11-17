/** Đọc tệp đính kèm thành chữ (TXT/CSV/MD/JSON/HTML, PDF, DOCX) hoặc ảnh (base64) để gửi kèm câu hỏi.
 *  PDF quét (không có lớp chữ) được giữ nguyên tệp (kind "file") để gửi cho model tự đọc. */

import { formatPages } from "./pdfPages";

export interface Attachment {
  name: string;
  /** text = chữ ghép vào câu hỏi · image = ảnh gửi kèm · file = tệp gửi nguyên vẹn cho model đọc (PDF quét) */
  kind: "text" | "image" | "file";
  text?: string;
  mime?: string;
  /** base64: ảnh, tệp gửi nguyên vẹn; với .docx là bản gốc của tệp để so sánh phiên bản trong Word */
  data?: string;
  note?: string;
  /** Số trang (PDF) */
  pages?: number;
  /** PDF lẫn trang quét: các trang không có lớp chữ */
  scanPages?: number[];
}

export const PDF_MIME = "application/pdf";

/** Cỡ tối đa của một PDF gửi nguyên tệp. Anthropic nhận tối đa 32 MB cho cả yêu cầu (base64 làm tệp to thêm một phần ba). */
export const MAX_PDF_FILE_BYTES = 15 * 1024 * 1024;
/** Tổng cỡ (ký tự base64) của mọi tệp gửi nguyên vẹn trong một hội thoại. Tệp nằm lại trong lịch sử và được gửi lại ở mọi lượt
 *  (lịch sử chỉ thêm), nên quá trần này thì yêu cầu vượt giới hạn 32 MB của Anthropic và mọi lượt sau đều lỗi. */
export const MAX_FILES_TOTAL_B64 = 24 * 1024 * 1024;
/** Tổng số ký tự base64 của các tệp gửi nguyên vẹn đang nằm trong một lịch sử hội thoại */
export function filesBase64Chars(history: { role: string; files?: { data: string }[] }[]): number {
  return history.reduce((n, m) => n + (m.role === "user" && m.files ? m.files.reduce((k, f) => k + f.data.length, 0) : 0), 0);
}
/** Cỡ tối đa của tệp .docx được giữ bản gốc để so sánh phiên bản */
export const MAX_DOCX_KEEP_BYTES = 20 * 1024 * 1024;
export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export const MAX_ATTACH_CHARS = 120000;

export function clipText(t: string, max = MAX_ATTACH_CHARS): { text: string; note?: string } {
  if (t.length <= max) return { text: t };
  return { text: t.slice(0, max), note: `Tệp dài, đã lấy ${max.toLocaleString("vi-VN")}/${t.length.toLocaleString("vi-VN")} ký tự đầu.` };
}

/** Lấy chữ từ document.xml của DOCX (giữ ngắt đoạn, tab, ô bảng). */
export function docxXmlToText(xml: string): string {
  return xml
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br[^>]*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n")
    .replace(/<\/w:tc>/g, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function toBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

async function pdfText(buf: ArrayBuffer): Promise<{ text: string; pages: number; blank: number[] }> {
  const pdfjs: any = await import("pdfjs-dist");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdf.worker.min.js", document.baseURI).toString();
  }
  // pdf.js chiếm lấy vùng nhớ được đưa vào: đưa một bản sao để còn giữ được tệp gốc (gửi nguyên tệp khi là PDF quét)
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf.slice(0)) }).promise;
  const pages: string[] = [];
  /** Các trang không có lớp chữ (ảnh quét lẫn trong tệp) */
  const blank: number[] = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const tc = await page.getTextContent();
    let line = "";
    const out: string[] = [];
    for (const it of tc.items as any[]) {
      line += it.str;
      if (it.hasEOL) {
        out.push(line);
        line = "";
      }
    }
    if (line) out.push(line);
    if (!out.join("").trim()) blank.push(i);
    pages.push(`[Trang ${i}]\n${out.join("\n")}`);
    if (pages.join("\n").length > MAX_ATTACH_CHARS) break;
  }
  return { text: pages.join("\n\n"), pages: doc.numPages, blank };
}

export async function readAttachment(file: File): Promise<Attachment> {
  const name = file.name;
  const lower = name.toLowerCase();
  if (/^image\/(png|jpe?g|gif|webp)$/.test(file.type)) {
    if (file.size > 5 * 1024 * 1024) throw new Error(`Ảnh "${name}" quá lớn (tối đa 5 MB).`);
    return { name, kind: "image", mime: file.type, data: toBase64(await file.arrayBuffer()) };
  }
  if (lower.endsWith(".pdf")) {
    const buf = await file.arrayBuffer();
    const { text: t, pages, blank } = await pdfText(buf);
    if (!t.replace(/\[Trang \d+\]/g, "").trim()) {
      // PDF quét: không có lớp chữ → giữ nguyên tệp để model tự đọc (Anthropic, Gemini… ; nguồn khác sẽ được báo rõ khi gửi)
      if (file.size > MAX_PDF_FILE_BYTES) {
        throw new Error(`PDF "${name}" là bản quét (không có lớp chữ) và nặng ${(file.size / 1048576).toFixed(1)} MB, quá mức gửi nguyên tệp (tối đa ${MAX_PDF_FILE_BYTES / 1048576} MB). Hãy tách tệp hoặc giảm độ phân giải.`);
      }
      return { name, kind: "file", mime: "application/pdf", data: toBase64(buf), pages, note: `PDF quét, ${pages} trang, gửi nguyên tệp` };
    }
    const clipped = clipText(t);
    // Tệp lẫn trang quét: nói rõ trang nào chưa đọc được, không bỏ im lặng
    // Giữ bản gốc (trong bộ nhớ của phiên) để AI đọc các trang quét hoặc phần bị cắt bằng công cụ read_attachment_pages
    const keep = (blank.length || clipped.note) && file.size <= MAX_PDF_FILE_BYTES ? { mime: PDF_MIME, data: toBase64(buf), scanPages: blank } : {};
    const scanNote = blank.length
      ? `${blank.length} trang không có lớp chữ (ảnh quét: trang ${formatPages(blank.slice(0, 60))}${blank.length > 60 ? "…" : ""}) nên CHƯA có trong phần chữ dưới đây` +
        (keep.data ? "" : "; tệp quá lớn để giữ lại, cần đọc thì tách các trang đó thành tệp PDF riêng rồi đính kèm")
      : "";
    return { name, kind: "text", pages, text: clipped.text, ...keep, ...(clipped.note || scanNote ? { note: [clipped.note, scanNote].filter(Boolean).join(" ") } : {}) };
  }
  if (lower.endsWith(".docx")) {
    const JSZip = (await import("jszip")).default;
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file("word/document.xml")?.async("string");
    if (!xml) throw new Error(`Không đọc được "${name}".`);
    // Giữ bản gốc của tệp: trong Word, AI so tệp này với văn bản đang mở bằng công cụ so sánh phiên bản
    const keep = file.size <= MAX_DOCX_KEEP_BYTES ? { mime: DOCX_MIME, data: toBase64(buf) } : {};
    return { name, kind: "text", ...keep, ...clipText(docxXmlToText(xml)) };
  }
  if (/\.(txt|md|markdown|csv|tsv|json|xml|html?|ya?ml|log)$/i.test(lower) || file.type.startsWith("text/")) {
    let t = await file.text();
    if (/\.html?$/i.test(lower)) t = t.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, "").replace(/<[^>]+>/g, " ");
    return { name, kind: "text", ...clipText(t) };
  }
  throw new Error(`Chưa hỗ trợ loại tệp "${name}". Dùng PDF, DOCX, TXT, CSV, MD, JSON hoặc ảnh.`);
}

/** Ghép chữ của các tệp vào câu hỏi. opts.compare = đang ở Word, có công cụ so sánh phiên bản: ghi kèm cách gọi cho tệp .docx */
export function withAttachments(question: string, atts: Attachment[], opts: { compare?: boolean; pageTool?: boolean; byPages?: string[] } = {}): string {
  const texts = atts.filter((a) => a.kind === "text");
  const files = atts.filter((a) => a.kind === "file");
  if (!texts.length && !files.length) return question;
  /** Tệp PDF quét không gửi nguyên tệp ở lượt này: AI đọc theo trang bằng công cụ */
  const byPages = new Set(opts.byPages || []);
  const viaTool = (a: Attachment) =>
    `\n--- Tệp đính kèm: ${a.name} (PDF quét, ${a.pages ?? "?"} trang) — nội dung CHƯA được gửi kèm. Đọc bằng công cụ read_attachment_pages với file = "${a.name}" và pages (ví dụ "1-4"; mỗi lần tối đa 4 trang quét); đọc đủ các trang cần cho câu hỏi rồi mới trả lời ---`;
  return [
    question,
    ...files.filter((a) => byPages.has(a.name)).map(viaTool),
    ...files.filter((a) => !byPages.has(a.name)).map((a) => `\n--- Tệp đính kèm: ${a.name} (${a.note || "gửi nguyên tệp"}) — tệp được gửi kèm tin nhắn này để bạn đọc trực tiếp, kể cả chữ trong ảnh quét ---`),
    ...texts.map((a) => {
      const hint =
        opts.compare && a.mime === DOCX_MIME && a.data
          ? `\n(Muốn so tệp này với văn bản đang mở theo từng điều khoản: gọi word_compare_versions với file = "${a.name}"; đừng tự so bằng mắt.)`
          : "";
      const pageHint =
        opts.pageTool && a.mime === PDF_MIME && a.data
          ? `\n(Tệp có ${a.pages ?? "?"} trang. Phần chưa có ở đây — trang quét, phần bị cắt — đọc bằng công cụ read_attachment_pages với file = "${a.name}" và pages.)`
          : "";
      return `\n--- Tệp đính kèm: ${a.name}${a.note ? ` (${a.note})` : ""} ---${hint}${pageHint}\n${a.text}\n--- Hết tệp ${a.name} ---`;
    }),
  ].join("\n");
}
