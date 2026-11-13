/** Đọc tệp đính kèm thành chữ (TXT/CSV/MD/JSON/HTML, PDF, DOCX) hoặc ảnh (base64) để gửi kèm câu hỏi. */

export interface Attachment {
  name: string;
  kind: "text" | "image";
  text?: string;
  mime?: string;
  data?: string;
  note?: string;
}

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

async function pdfText(buf: ArrayBuffer): Promise<string> {
  const pdfjs: any = await import("pdfjs-dist");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdf.worker.min.js", document.baseURI).toString();
  }
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf) }).promise;
  const pages: string[] = [];
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
    pages.push(`[Trang ${i}]\n${out.join("\n")}`);
    if (pages.join("\n").length > MAX_ATTACH_CHARS) break;
  }
  return pages.join("\n\n");
}

export async function readAttachment(file: File): Promise<Attachment> {
  const name = file.name;
  const lower = name.toLowerCase();
  if (/^image\/(png|jpe?g|gif|webp)$/.test(file.type)) {
    if (file.size > 5 * 1024 * 1024) throw new Error(`Ảnh "${name}" quá lớn (tối đa 5 MB).`);
    return { name, kind: "image", mime: file.type, data: toBase64(await file.arrayBuffer()) };
  }
  if (lower.endsWith(".pdf")) {
    const t = await pdfText(await file.arrayBuffer());
    if (!t.replace(/\[Trang \d+\]/g, "").trim()) throw new Error(`PDF "${name}" không có lớp chữ (có thể là ảnh scan).`);
    return { name, kind: "text", ...clipText(t) };
  }
  if (lower.endsWith(".docx")) {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const xml = await zip.file("word/document.xml")?.async("string");
    if (!xml) throw new Error(`Không đọc được "${name}".`);
    return { name, kind: "text", ...clipText(docxXmlToText(xml)) };
  }
  if (/\.(txt|md|markdown|csv|tsv|json|xml|html?|ya?ml|log)$/i.test(lower) || file.type.startsWith("text/")) {
    let t = await file.text();
    if (/\.html?$/i.test(lower)) t = t.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, "").replace(/<[^>]+>/g, " ");
    return { name, kind: "text", ...clipText(t) };
  }
  throw new Error(`Chưa hỗ trợ loại tệp "${name}". Dùng PDF, DOCX, TXT, CSV, MD, JSON hoặc ảnh.`);
}

/** Ghép chữ của các tệp vào câu hỏi */
export function withAttachments(question: string, atts: Attachment[]): string {
  const texts = atts.filter((a) => a.kind === "text");
  if (!texts.length) return question;
  return [
    question,
    ...texts.map((a) => `\n--- Tệp đính kèm: ${a.name}${a.note ? ` (${a.note})` : ""} ---\n${a.text}\n--- Hết tệp ${a.name} ---`),
  ].join("\n");
}
