/** Đoạn và số điều của tệp .docx đính kèm, đọc thẳng từ gói OOXML (Đợt 12, phần D — `tasks/dot/dot-12/dot-12.md`). Hàm thuần.
 *  Hợp đồng thường đánh số bằng danh sách tự động của Word (kiểu đoạn Heading1 → "Điều %1.", Heading2 → "%1.%2."…): chữ thuần của tệp
 *  KHÔNG có các số này, nên AI không tra được "Điều 6.4". Ở đây dựng lại nhãn số như Word hiện (numbering.xml + numPr của đoạn hoặc của
 *  kiểu đoạn, có kế thừa basedOn) rồi đưa vào `buildOutline` dùng chung với Word đang mở. Chữ đã xóa có theo dõi (<w:del>) không tính. */
import { OutlineItem, ParaRow, findClause } from "./outline";

export interface DocxParts {
  document: string;
  styles?: string;
  numbering?: string;
}

interface Lvl {
  start: number;
  fmt: string;
  text: string;
}

const attr = (tag: string, name: string): string | undefined => {
  const m = new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag);
  return m ? m[1] : undefined;
};
const decodeXml = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function parseNumbering(xml = "") {
  const abstracts = new Map<string, Lvl[]>();
  for (const m of xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="([^"]+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
    const lvls: Lvl[] = [];
    for (const l of m[2].matchAll(/<w:lvl\b[^>]*w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g)) {
      const body = l[2];
      lvls[Number(l[1])] = {
        start: Number(attr(/<w:start\b[^>]*>/.exec(body)?.[0] || "", "val") ?? 1),
        fmt: attr(/<w:numFmt\b[^>]*>/.exec(body)?.[0] || "", "val") || "decimal",
        text: decodeXml(attr(/<w:lvlText\b[^>]*>/.exec(body)?.[0] || "", "val") ?? ""),
      };
    }
    abstracts.set(m[1], lvls);
  }
  const nums = new Map<string, string>();
  for (const m of xml.matchAll(/<w:num\b[^>]*w:numId="([^"]+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
    const a = attr(/<w:abstractNumId\b[^>]*>/.exec(m[2])?.[0] || "", "val");
    if (a !== undefined) nums.set(m[1], a);
  }
  return { abstracts, nums };
}

interface StyleInfo {
  basedOn?: string;
  numId?: string;
  ilvl?: number;
  outline?: number;
}

function numPrOf(pPr: string): { numId?: string; ilvl?: number } {
  const np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr)?.[1];
  if (!np) return {};
  const numId = attr(/<w:numId\b[^>]*>/.exec(np)?.[0] || "", "val");
  const il = attr(/<w:ilvl\b[^>]*>/.exec(np)?.[0] || "", "val");
  return { numId, ilvl: il === undefined ? undefined : Number(il) };
}

function parseStyles(xml = ""): Map<string, StyleInfo> {
  const out = new Map<string, StyleInfo>();
  for (const m of xml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    const body = m[2];
    const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(body)?.[1] || "";
    const ol = attr(/<w:outlineLvl\b[^>]*>/.exec(pPr)?.[0] || "", "val");
    out.set(m[1], {
      basedOn: attr(/<w:basedOn\b[^>]*>/.exec(body)?.[0] || "", "val"),
      ...numPrOf(pPr),
      ...(ol !== undefined ? { outline: Number(ol) } : {}),
    });
  }
  return out;
}

/** Thuộc tính của kiểu đoạn sau khi tính kế thừa basedOn (tối đa 10 tầng) */
function resolveStyle(styles: Map<string, StyleInfo>, id?: string): StyleInfo {
  const chain: StyleInfo[] = [];
  for (let k = 0, cur = id; cur && k < 10; k++) {
    const s = styles.get(cur);
    if (!s) break;
    chain.push(s);
    cur = s.basedOn;
  }
  const r: StyleInfo = {};
  for (const s of chain.reverse()) {
    if (s.numId !== undefined) r.numId = s.numId;
    if (s.ilvl !== undefined) r.ilvl = s.ilvl;
    if (s.outline !== undefined) r.outline = s.outline;
  }
  return r;
}

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
function roman(n: number): string {
  let s = "";
  for (const [v, r] of ROMAN) while (n >= v) (s += r), (n -= v);
  return s;
}
function letters(n: number): string {
  // Word: a … z, aa … zz (lặp chữ)
  const c = String.fromCharCode(97 + ((n - 1) % 26));
  return c.repeat(Math.floor((n - 1) / 26) + 1);
}
export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case "lowerLetter":
      return letters(n);
    case "upperLetter":
      return letters(n).toUpperCase();
    case "lowerRoman":
      return roman(n);
    case "upperRoman":
      return roman(n).toUpperCase();
    case "decimalZero":
      return n < 10 ? "0" + n : String(n);
    default:
      return String(n);
  }
}

/** Các đoạn của thân tệp .docx kèm nhãn số dựng lại như Word hiện */
export function docxParagraphs(parts: DocxParts): ParaRow[] {
  const { abstracts, nums } = parseNumbering(parts.numbering);
  const styles = parseStyles(parts.styles);
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(parts.document)?.[1] ?? parts.document;
  // Vùng bảng (để đánh dấu đoạn trong ô bảng)
  const tables: [number, number][] = [];
  {
    const re = /<w:tbl>|<\/w:tbl>/g;
    const stack: number[] = [];
    for (let m: RegExpExecArray | null; (m = re.exec(body)); ) {
      if (m[0] === "<w:tbl>") stack.push(m.index);
      else if (stack.length) {
        const s = stack.pop()!;
        if (!stack.length) tables.push([s, m.index]);
      }
    }
  }
  const inTable = (pos: number) => tables.some(([a, b]) => pos > a && pos < b);
  /** Bộ đếm theo danh sách trừu tượng: các numId cùng một abstractNum đếm chung (như Word) */
  const counters = new Map<string, (number | undefined)[]>();
  const rows: ParaRow[] = [];
  for (const m of body.matchAll(/<w:p\/>|<w:p\b(?!P|r)[^>]*>([\s\S]*?)<\/w:p>/g)) {
    const inner = m[1] || "";
    const pPr = /^\s*<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(inner)?.[1] || "";
    const styleId = attr(/<w:pStyle\b[^>]*>/.exec(pPr)?.[0] || "", "val");
    const st = resolveStyle(styles, styleId);
    const direct = numPrOf(pPr);
    const numId = direct.numId ?? st.numId;
    const ilvl = direct.ilvl ?? st.ilvl ?? 0;
    const olDirect = attr(/<w:outlineLvl\b[^>]*>/.exec(pPr)?.[0] || "", "val");
    const outline = olDirect !== undefined ? Number(olDirect) : st.outline;
    const live = inner.replace(/<w:del\b[\s\S]*?<\/w:del>/g, "").replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, "");
    // Chữ theo đúng thứ tự: <w:t> (không tính <w:delText>), dấu tab, ngắt dòng
    const text = decodeXml(
      [...live.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\b[^>]*\/>/g)].map((x) => (x[1] !== undefined ? x[1] : x[0].startsWith("<w:tab") ? "\t" : " ")).join(""),
    );
    let listString: string | undefined;
    if (numId && numId !== "0") {
      const abs = nums.get(numId);
      const lvls = abs !== undefined ? abstracts.get(abs) : undefined;
      const lv = lvls?.[ilvl];
      if (abs !== undefined && lvls && lv && lv.fmt !== "bullet" && lv.fmt !== "none") {
        const c = counters.get(abs) || [];
        c[ilvl] = (c[ilvl] ?? (lv.start - 1)) + 1;
        for (let k = ilvl + 1; k < c.length; k++) c[k] = undefined;
        counters.set(abs, c);
        listString = lv.text.replace(/%(\d)/g, (_, d) => {
          const k = Number(d) - 1;
          const L = lvls[k];
          return formatNumber(c[k] ?? L?.start ?? 1, L?.fmt || "decimal");
        });
      }
    }
    rows.push({
      i: rows.length,
      text,
      ...(styleId ? { style: styleId } : {}),
      ...(outline !== undefined && outline >= 0 && outline <= 8 ? { outlineLevel: outline + 1 } : {}),
      ...(listString ? { listString, listLevel: ilvl, listId: Number(numId) } : {}),
      ...(inTable(m.index ?? 0) ? { inTable: true } : {}),
    });
  }
  return rows;
}

/** Chữ hiện của một đoạn kèm nhãn số (như người đọc thấy trong Word) */
export function paraDisplay(r: ParaRow): string {
  return r.listString ? `${r.listString} ${r.text}`.trim() : r.text;
}


/** Tệp .docx đính kèm từ cỡ này (ký tự) thì không chép nguyên vào tin nhắn mà gửi mục lục + đọc theo điều (≈ 6.000 token chữ tiếng Việt) */
export const DOCX_INLINE_MAX_CHARS = 21600;
/** Mỗi lần đọc theo điều trả tối đa chừng này ký tự */
export const DOCX_READ_MAX_CHARS = 20000;

/** Mục lục gọn để gửi kèm tin nhắn, mỗi dòng "Điều 5 [¶40–52] Tiêu đề": các mục cấp 0–1 (Điều, khoản cấp một); dài quá maxChars thì
 *  chỉ kê cấp 0 (Điều, Phụ lục) để luôn thấy đủ cả văn bản (Chạy thật 22/11: mục lục cắt ở Điều 20, AI phải đọc dò từng dải đoạn) */
export function docxOutlineText(items: OutlineItem[], maxChars = 6000): string {
  const line = (it: OutlineItem, w: number) => `${it.num} [¶${it.i}–${it.end}] ${it.title.slice(0, w)}`;
  const two = items.filter((it) => it.level <= 1 && it.num).map((it) => line(it, it.level ? 50 : 70));
  if (two.join("\n").length <= maxChars) return two.join("\n");
  const top = items.filter((it) => it.level === 0 && it.num).map((it) => line(it, 70));
  const out: string[] = [];
  let n = 0;
  for (const l of top) {
    if (n + l.length + 1 > maxChars) {
      out.push("… (mục lục còn tiếp — đọc theo dải đoạn để xem phần sau)");
      break;
    }
    out.push(l);
    n += l.length + 1;
  }
  return out.join("\n") + (out.length ? "\n(Mục lục chỉ kê cấp Điều; đọc một Điều để thấy các khoản của nó.)" : "");
}

/** Dải đoạn "120-160" hoặc "87" → [đầu, cuối] trong giới hạn số đoạn */
export function parseParaRange(spec: unknown, total: number): [number, number] | null {
  const m = /^\s*¶?(\d+)\s*(?:[-–]\s*¶?(\d+))?\s*$/.exec(String(spec ?? ""));
  if (!m) return null;
  const a = Number(m[1]);
  const b = m[2] === undefined ? a : Number(m[2]);
  if (a > b || a >= total) return null;
  return [a, Math.min(b, total - 1)];
}

export interface DocxSection {
  /** Đoạn đã trả (dải) */
  from: number;
  to: number;
  text: string;
  /** Đoạn kế tiếp chưa trả (khi phần đọc bị giới hạn ký tự) */
  next?: number;
  clause?: string;
}

/** Chữ của một điều ("Điều 6.4", "18.2(b)") hoặc một dải đoạn, mỗi đoạn một dòng có số ¶, tối đa maxChars ký tự */
export function readDocxSection(rows: ParaRow[], items: OutlineItem[], q: { clause?: string; paragraphs?: string }, maxChars = DOCX_READ_MAX_CHARS): DocxSection {
  let from: number;
  let to: number;
  let clause: string | undefined;
  if (q.clause && String(q.clause).trim()) {
    const hit = findClause(items, String(q.clause));
    if (!hit.length) throw new Error(`Không thấy "${q.clause}" trong mục lục của tệp. Xem mục lục ở tin nhắn đính kèm, hoặc đọc theo dải đoạn (paragraphs).`);
    from = hit[0].i;
    to = hit[0].end;
    clause = hit[0].num;
  } else {
    const r = parseParaRange(q.paragraphs, rows.length);
    if (!r) throw new Error(`Cần clause (ví dụ "Điều 6.4") hoặc paragraphs (ví dụ "120-160", tệp có ${rows.length} đoạn, đánh số từ 0).`);
    [from, to] = r;
  }
  const out: string[] = [];
  let used = 0;
  let k = from;
  for (; k <= to; k++) {
    const line = `[¶${k}] ${paraDisplay(rows[k])}`;
    if (used + line.length > maxChars && out.length) break;
    out.push(line.length > maxChars ? line.slice(0, maxChars) : line);
    used += line.length + 1;
  }
  return { from, to: k - 1, text: out.join("\n"), ...(k <= to ? { next: k } : {}), ...(clause ? { clause } : {}) };
}

/** Đọc gói .docx (ArrayBuffer hoặc base64) → các đoạn kèm nhãn số */
export async function docxRowsFrom(data: ArrayBuffer | string): Promise<ParaRow[]> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(data as any, typeof data === "string" ? { base64: true } : undefined);
  const get = async (n: string) => (await zip.file(n)?.async("string")) || undefined;
  const document = await get("word/document.xml");
  if (!document) throw new Error("Tệp .docx không có word/document.xml.");
  return docxParagraphs({ document, styles: await get("word/styles.xml"), numbering: await get("word/numbering.xml") });
}
