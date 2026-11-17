/** Mở tệp PDF đính kèm bằng pdf.js và đọc từng trang: chữ của lớp chữ, hoặc ảnh JPEG của trang quét (cần DOM: canvas).
 *  Tài liệu được giữ mở theo tên tệp trong phiên để lần gọi sau không phải phân tích lại. */
import { isScanPage } from "./pdfPages";

/** Ảnh trang: khoảng 150 dpi (72 điểm một inch), cạnh dài không quá 1.600 điểm ảnh — đủ đọc chữ in cỡ 12, khoảng 1.800 token một trang */
const DPI_SCALE = 150 / 72;
const MAX_SIDE = 1600;
const JPEG_QUALITY = 0.82;

export interface PdfPage {
  page: number;
  /** Số ký tự lớp chữ của trang */
  chars: number;
  scan: boolean;
  text: string;
}

interface OpenDoc {
  key: string;
  doc: any;
  /** Số ký tự lớp chữ từng trang, đọc dần */
  text: Map<number, string>;
}

const open = new Map<string, OpenDoc>();
const MAX_OPEN = 3;

async function pdfjsLib(): Promise<any> {
  const pdfjs: any = await import("pdfjs-dist");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdf.worker.min.js", document.baseURI).toString();
  return pdfjs;
}

function fromBase64(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function openPdf(name: string, data: string): Promise<OpenDoc> {
  // Khóa theo tên và cỡ: đính kèm lại một tệp cùng tên nhưng nội dung khác thì mở lại
  const key = `${name}|${data.length}`;
  const have = open.get(name);
  if (have && have.key === key) return have;
  if (have) {
    try {
      await have.doc.destroy();
    } catch {
      /* bỏ qua */
    }
    open.delete(name);
  }
  const pdfjs = await pdfjsLib();
  const doc = await pdfjs.getDocument({ data: fromBase64(data) }).promise;
  const od: OpenDoc = { key, doc, text: new Map() };
  open.set(name, od);
  while (open.size > MAX_OPEN) {
    const first = open.keys().next().value as string;
    try {
      await open.get(first)!.doc.destroy();
    } catch {
      /* bỏ qua */
    }
    open.delete(first);
  }
  return od;
}

async function pageText(od: OpenDoc, n: number): Promise<string> {
  const have = od.text.get(n);
  if (have !== undefined) return have;
  const page = await od.doc.getPage(n);
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
  const t = out.join("\n").trim();
  od.text.set(n, t);
  return t;
}

export async function pdfPageCount(name: string, data: string): Promise<number> {
  return (await openPdf(name, data)).doc.numPages;
}

/** Chữ và phân loại của các trang được nêu */
export async function readPdfPages(name: string, data: string, pages: number[]): Promise<PdfPage[]> {
  const od = await openPdf(name, data);
  const out: PdfPage[] = [];
  for (const n of pages) {
    const text = await pageText(od, n);
    const chars = text.replace(/\s/g, "").length;
    out.push({ page: n, chars, scan: isScanPage(chars), text });
  }
  return out;
}

/** Dựng ảnh JPEG (base64) của một trang */
export async function renderPdfPage(name: string, data: string, n: number): Promise<{ mime: string; data: string; width: number; height: number }> {
  const od = await openPdf(name, data);
  const page = await od.doc.getPage(n);
  const base = page.getViewport({ scale: 1 });
  const scale = Math.min(DPI_SCALE, MAX_SIDE / Math.max(base.width, base.height));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(viewport.width);
  canvas.height = Math.round(viewport.height);
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Không dựng được ảnh trang (không có canvas).");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  const url = canvas.toDataURL("image/jpeg", JPEG_QUALITY);
  const w = canvas.width;
  const h = canvas.height;
  canvas.width = 0;
  canvas.height = 0;
  return { mime: "image/jpeg", data: url.slice(url.indexOf(",") + 1), width: w, height: h };
}
