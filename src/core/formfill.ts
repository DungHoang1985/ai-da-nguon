/** Điền mẫu trong Word (Đợt 8): phần thuần, không phụ thuộc Office — kiểm thử được bằng npm test.
 *  - Chỗ trống gõ tay trong đoạn ("……", "____", "[●]", "ngày   tháng   năm"): vị trí, nhãn đứng trước, chuỗi cần tìm trong Word
 *    và đó là lần xuất hiện thứ mấy (Word tìm theo chuỗi; tìm một đoạn chấm ngắn hơn dãy thật chỉ thay một phần — đã đo 18/11/2026).
 *  - Ô điền mẫu: danh sách mục và định dạng ngày đọc từ OOXML của ô (API đọc danh sách mục không có trên mọi bản Word).
 *  - Giá trị điền: ô đánh dấu, ô chọn, ô ngày. */
import { BLANKS, SIGN_LINE } from "./consistency";

export interface Blank {
  /** Vị trí trong chữ của đoạn */
  at: number;
  len: number;
  /** Chuỗi chỗ trống đúng như trong đoạn (để tìm trong Word) */
  str: string;
  /** Lần xuất hiện thứ mấy (từ 0) của `str` trong đoạn, đếm không chồng nhau từ đầu đoạn */
  occurrence: number;
  /** Chữ đứng trước chỗ trống trong đoạn (từ sau chỗ trống trước đó), đã gọn — để AI biết chỗ này điền gì */
  label: string;
  /** Chữ đứng ngay sau (tới chỗ trống kế hoặc 40 ký tự) */
  after: string;
}

const tidy = (s: string) => s.replace(/[\u0001\u0002]/g, "").replace(/ /g, " ");

/** Chỗ trống trong một đoạn (đã gộp các chỗ chồng nhau), theo thứ tự trong đoạn. Dòng ký tên ngắn thì không có chỗ trống nào. */
export function findBlanks(rawText: string): Blank[] {
  const text = tidy(rawText);
  if (!text.trim()) return [];
  if (text.length <= 100 && SIGN_LINE.test(text)) return [];
  const hits: { at: number; len: number }[] = [];
  for (const re of BLANKS) {
    re.lastIndex = 0;
    for (const m of text.matchAll(re)) hits.push({ at: m.index as number, len: m[0].length });
  }
  hits.sort((a, b) => a.at - b.at || b.len - a.len);
  const merged: { at: number; len: number }[] = [];
  for (const h of hits) {
    const last = merged[merged.length - 1];
    if (last && h.at < last.at + last.len) last.len = Math.max(last.len, h.at + h.len - last.at);
    else merged.push({ ...h });
  }
  return merged.map((h, k) => {
    // "ngày   tháng   năm": chỗ trống là các khoảng trắng giữa các chữ, giữ cả cụm để thay một lần
    const str = text.slice(h.at, h.at + h.len);
    const prevEnd = k > 0 ? merged[k - 1].at + merged[k - 1].len : 0;
    const nextAt = k + 1 < merged.length ? merged[k + 1].at : text.length;
    const label = text
      .slice(prevEnd, h.at)
      .replace(/\s+/g, " ")
      .trim()
      .slice(-80);
    const after = text
      .slice(h.at + h.len, Math.min(nextAt, h.at + h.len + 40))
      .replace(/\s+/g, " ")
      .trim();
    return { at: h.at, len: h.len, str, occurrence: occurrenceOf(text, str, h.at), label, after };
  });
}

/** Đếm số lần `str` xuất hiện (không chồng nhau, từ đầu) trước vị trí `at` — đúng thứ tự kết quả tìm của Word */
export function occurrenceOf(text: string, str: string, at: number): number {
  if (!str) return 0;
  let n = 0;
  let from = 0;
  for (;;) {
    const k = text.indexOf(str, from);
    if (k < 0 || k >= at) return n;
    n++;
    from = k + str.length;
  }
}

/** Số lần `str` xuất hiện trong cả đoạn (không chồng nhau) */
export function countOccurrences(text: string, str: string): number {
  return occurrenceOf(tidy(text), str, Number.MAX_SAFE_INTEGER);
}

/** Danh sách mục của ô chọn / ô chọn hoặc nhập, đọc từ OOXML của ô (w:listItem) */
export function listItemsFromOoxml(xml: string): string[] {
  const sdt = /<w:sdt>[\s\S]*?<\/w:sdtPr>/.exec(String(xml || ""))?.[0] || "";
  return [...sdt.matchAll(/<w:listItem\b([^>]*)\/?>/g)]
    .map((m) => /w:displayText="([^"]*)"/.exec(m[1])?.[1] ?? /w:value="([^"]*)"/.exec(m[1])?.[1] ?? "")
    .map(unxml)
    .filter((x) => x !== "");
}

/** Định dạng ngày của ô ngày (w:dateFormat), ví dụ "dd/MM/yyyy"; không có thì null */
export function dateFormatFromOoxml(xml: string): string | null {
  const m = /<w:date\b[\s\S]*?<w:dateFormat w:val="([^"]*)"/.exec(String(xml || ""));
  return m ? unxml(m[1]) : null;
}

const unxml = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
const xmlEsc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Ô chọn: OOXML mới của ô với mục `item` được chọn (bỏ trạng thái "đang hiện chữ gợi ý", thay chữ trong w:sdtContent).
 *  null nếu không dựng được (không thấy ô hoặc không có chữ trong ô). */
export function selectDropdownOoxml(xml: string, item: string): string | null {
  const src = String(xml || "");
  const m = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(src);
  if (!m) return null;
  const sdt = m[0];
  const content = /<w:sdtContent>([\s\S]*)<\/w:sdtContent>/.exec(sdt);
  if (!content) return null;
  // Giữ run đầu tiên (định dạng), thay chữ của nó, bỏ các run sau
  const firstRun = /<w:r\b[\s\S]*?<\/w:r>/.exec(content[1]);
  if (!firstRun) return null;
  const run = /<w:t\b[^>]*>[\s\S]*?<\/w:t>|<w:t\b[^>]*\/>/.test(firstRun[0])
    ? firstRun[0].replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>|<w:t\b[^>]*\/>/, `<w:t xml:space="preserve">${xmlEsc(item)}</w:t>`)
    : firstRun[0].replace(/<\/w:r>$/, `<w:t xml:space="preserve">${xmlEsc(item)}</w:t></w:r>`);
  // Bỏ kiểu chữ gợi ý (PlaceholderText) để chữ đã chọn không hiện màu xám
  const cleanRun = run.replace(/<w:rStyle w:val="PlaceholderText"\/>/g, "");
  const inner = content[1].replace(/<w:r\b[\s\S]*<\/w:r>/, cleanRun);
  const sdtNew = sdt.replace(/<w:showingPlcHdr\/>/g, "").replace(/<w:sdtContent>[\s\S]*<\/w:sdtContent>/, `<w:sdtContent>${inner}</w:sdtContent>`);
  return src.slice(0, m.index) + sdtNew + src.slice((m.index as number) + sdt.length);
}

/** Giá trị cho ô đánh dấu: true / false, hoặc null nếu không hiểu */
export function checkboxValue(v: unknown): boolean | null {
  if (typeof v === "boolean") return v;
  const s = String(v ?? "")
    .trim()
    .toLowerCase();
  if (/^(true|1|x|có|co|yes|đánh dấu|danh dau|chọn|chon|☒|☑|✓|✔)$/.test(s)) return true;
  if (/^(false|0|không|khong|no|bỏ|bo|bỏ chọn|☐)$/.test(s)) return false;
  return null;
}

/** Ô chọn: tìm mục khớp giá trị (không phân biệt hoa thường, bỏ dấu cách thừa); null nếu không có */
export function matchListItem(items: string[], v: unknown): string | null {
  const want = String(v ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
  if (!want) return null;
  return items.find((x) => x.replace(/\s+/g, " ").trim().toLowerCase() === want) ?? null;
}

/** Đọc một ngày người dùng / AI đưa: "18/11/2026", "18-11-2026", "2026-11-18", "ngày 18 tháng 11 năm 2026". null nếu không chắc. */
export function parseDate(v: unknown): { d: number; m: number; y: number } | null {
  const s = String(v ?? "").trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/.exec(s);
  let d: number, mo: number, y: number;
  if (m) [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  else if ((m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/.exec(s))) [d, mo, y] = [Number(m[1]), Number(m[2]), Number(m[3])];
  else if ((m = /^ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})$/i.exec(s))) [d, mo, y] = [Number(m[1]), Number(m[2]), Number(m[3])];
  else return null;
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
  return { d, m: mo, y };
}

/** Viết ngày theo định dạng của ô (dd, d, MM, M, yyyy, yy); định dạng có chữ tháng (MMM…) thì dùng dd/MM/yyyy */
export function formatDate(x: { d: number; m: number; y: number }, fmt: string | null): string {
  const f = fmt && !/MMM|ddd/.test(fmt) ? fmt : "dd/MM/yyyy";
  const two = (n: number) => String(n).padStart(2, "0");
  return f.replace(/yyyy|yy|MM|M|dd|d/g, (t) =>
    t === "yyyy" ? String(x.y) : t === "yy" ? two(x.y % 100) : t === "MM" ? two(x.m) : t === "M" ? String(x.m) : t === "dd" ? two(x.d) : String(x.d)
  );
}
