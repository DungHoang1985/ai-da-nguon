/** Điền mẫu trong Word (Đợt 8): phần thuần, không phụ thuộc Office — kiểm thử được bằng npm test.
 *  - Chỗ trống gõ tay trong đoạn ("……", "____", "[●]", "ngày   tháng   năm"): vị trí, nhãn đứng trước, chuỗi cần tìm trong Word
 *    và đó là lần xuất hiện thứ mấy (Word tìm theo chuỗi; tìm một đoạn chấm ngắn hơn dãy thật chỉ thay một phần — đã đo 18/11/2026).
 *  - Ô điền mẫu: danh sách mục và định dạng ngày đọc từ OOXML của ô (API đọc danh sách mục không có trên mọi bản Word).
 *  - Giá trị điền: ô đánh dấu, ô chọn, ô ngày. */
import { BLANKS, SIGN_LINE, notABlank } from "./consistency";

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

/** Chữ để dò mẫu: dấu cách không ngắt thành dấu cách thường, CÙNG độ dài với chữ gốc, nên vị trí dò dùng thẳng trên chữ gốc;
 *  chuỗi đưa cho Word tìm lấy từ chữ gốc, giữ dấu cách không ngắt (soát mã Đợt 8, #10). */
const tidy = (s: string) => s.replace(/ /g, " ");

/** Chỗ trống trong một đoạn (đã gộp các chỗ chồng nhau), theo thứ tự trong đoạn. Dòng ký tên ngắn thì không có chỗ trống nào. */
export function findBlanks(rawText: string): Blank[] {
  const raw = String(rawText ?? "");
  const text = tidy(raw);
  if (!text.trim()) return [];
  if (text.length <= 100 && SIGN_LINE.test(text)) return [];
  const hits: { at: number; len: number }[] = [];
  for (const re of BLANKS) {
    re.lastIndex = 0;
    for (const m of text.matchAll(re)) if (!notABlank(text, m.index as number, m[0].length)) hits.push({ at: m.index as number, len: m[0].length });
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
    const str = raw.slice(h.at, h.at + h.len);
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
    return { at: h.at, len: h.len, str, occurrence: occurrenceOf(raw, str, h.at), label, after };
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
  return occurrenceOf(String(text ?? ""), str, Number.MAX_SAFE_INTEGER);
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

/** Vị trí ô điền mẫu ngoài cùng (w:sdt, tính cả ô lồng bên trong) trong một chuỗi OOXML */
function outerSdt(src: string): { start: number; end: number } | null {
  const open = /<w:sdt(?=[\s>])/g;
  const first = open.exec(src);
  if (!first) return null;
  const tag = /<w:sdt(?=[\s>])|<\/w:sdt>/g;
  tag.lastIndex = first.index;
  let depth = 0;
  for (let m = tag.exec(src); m; m = tag.exec(src)) {
    depth += m[0] === "</w:sdt>" ? -1 : 1;
    if (depth === 0) return { start: first.index, end: m.index + m[0].length };
  }
  return null;
}

/** Bỏ dấu vết sửa định dạng có theo dõi (w:rPrChange, w:pPrChange): chúng chứa w:rPr / w:pPr lồng bên trong, làm phép tách định dạng
 *  bằng biểu thức không tham lam cắt sai và sinh OOXML hở thẻ (soát mã Đợt 8, #5). Nội dung mới không cần dấu vết đó. */
const dropFormatChanges = (xml: string) => xml.replace(/<w:rPrChange\b[\s\S]*?<\/w:rPrChange>/g, "").replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, "");

/** Định dạng chữ (w:rPr) của run đầu tiên trong một đoạn XML. Ô đang hiện chữ gợi ý thì bỏ kiểu chữ của run (kiểu chữ gợi ý màu xám,
 *  tên kiểu khác nhau theo ngôn ngữ của Word — soát mã Đợt 8, #13); còn lại chỉ bỏ kiểu "PlaceholderText". */
function firstRunProps(xml: string, placeholder: boolean): string {
  const r = /<w:r(?=[\s>])[^>]*>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)?/.exec(xml);
  let props = r?.[1] || "";
  props = placeholder ? props.replace(/<w:rStyle\b[^>]*\/>/g, "") : props.replace(/<w:rStyle w:val="PlaceholderText"\/>/g, "");
  return props.replace(/<w:rPr>\s*<\/w:rPr>/, "");
}

/** Bỏ ký tự điều khiển làm hỏng XML (giữ tab và xuống dòng) — soát mã Đợt 8, #16 */
const safeText = (s: string) => String(s ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, " ");

/** Một run chữ; tab thành w:tab */
const runOf = (props: string, text: string) =>
  safeText(text)
    .split("\t")
    .map((part, k) => `<w:r>${props}${k ? "<w:tab/>" : ""}<w:t xml:space="preserve">${xmlEsc(part)}</w:t></w:r>`)
    .join("");

/** Đặt phông chữ của run (ký hiệu ô đánh dấu phải dùng phông của trạng thái đánh dấu — soát mã Đợt 8, #11) */
const withFont = (props: string, font: string | undefined) => {
  if (!font) return props;
  const f = `<w:rFonts w:ascii="${xmlEsc(font)}" w:eastAsia="${xmlEsc(font)}" w:hAnsi="${xmlEsc(font)}" w:cs="${xmlEsc(font)}"/>`;
  if (!props) return `<w:rPr>${f}</w:rPr>`;
  return /<w:rFonts\b[^>]*\/>/.test(props) ? props.replace(/<w:rFonts\b[^>]*\/>/, f) : props.replace("<w:rPr>", `<w:rPr>${f}`);
};

/** Lý do KHÔNG nên ghi đè cả ô này (null = ghi được). Đọc từ OOXML của ô.
 *  - ô gắn dữ liệu (w:dataBinding): Word lấy lại giá trị từ phần dữ liệu, kết quả có thể không như đã báo (#14);
 *  - ô chứa bảng, hình, đối tượng: chữ mới sẽ xóa mất bảng / hình (#6);
 *  - ô bao cả ô bảng hoặc hàng bảng: không đặt đoạn văn vào đó được (#17). */
export function controlRefusal(xml: string): string | null {
  const src = String(xml || "");
  const pos = outerSdt(src);
  if (!pos) return "không đọc được cấu trúc ô";
  const sdt = src.slice(pos.start, pos.end);
  if (/<w:dataBinding\b/.test(sdt)) return "ô gắn với phần dữ liệu của văn bản (data binding): add-in không ghi đè, hãy điền tay";
  const content = /<w:sdtContent>([\s\S]*)<\/w:sdtContent>/.exec(sdt)?.[1] || "";
  if (/<w:tc(?=[\s>])|<w:tr(?=[\s>])/.test(content)) return "ô bao cả ô bảng hoặc hàng bảng: add-in không ghi đè, hãy điền tay";
  if (/<w:tbl(?=[\s>])|<w:drawing\b|<w:pict\b|<w:object\b|<m:oMath/.test(content)) return "ô đang chứa bảng, hình hoặc công thức: ghi chữ vào sẽ xóa mất, hãy sửa tay";
  return null;
}

/** OOXML mới của ô điền mẫu với giá trị mới — ghi cả ô bằng insertOoxml thay cho ghi chữ vào ô.
 *  Vì sao (đo 18/11/2026, `tham-do-8-sau-hoan-tac.js`): ghi chữ vào ô rồi Hoàn tác (từ chối sửa đổi) để lại trong ô 5 dấu cách thay
 *  chữ gợi ý, và từ đó Word từ chối mọi lệnh ghi vào ô (GeneralException); ô đánh dấu thì lệch (w14:checked = 1 mà hiện ☐).
 *  Thay cả ô thì Hoàn tác trả lại nguyên ô cũ và điền lại được (ô chọn đã chạy theo cách này).
 *  - text: chữ mới (ô chữ, ô chọn, ô chọn hoặc nhập, ô ngày); nhiều dòng → ô trong dòng xuống dòng bằng w:br, ô chứa đoạn thì mỗi dòng một đoạn;
 *  - checked: ô đánh dấu (đặt w14:checked, ký hiệu và phông theo checkedState / uncheckedState của ô);
 *  - date: ô ngày, ghi thêm w:fullDate để bộ chọn ngày của Word hiện đúng ngày.
 *  null nếu không dựng được (gọi controlRefusal trước để có lý do). */
export function fillControlOoxml(xml: string, v: { text?: string; checked?: boolean; date?: { d: number; m: number; y: number } }): string | null {
  const src = String(xml || "");
  const pos = outerSdt(src);
  if (!pos) return null;
  const sdt = src.slice(pos.start, pos.end);
  if (controlRefusal(src)) return null;
  // Thuộc tính ô có thể tự đóng <w:sdtPr/> (ô không có thuộc tính nào)
  const prM = /<w:sdtPr>[\s\S]*?<\/w:sdtPr>|<w:sdtPr\s*\/>/.exec(sdt);
  // Nội dung ô có thể rỗng dạng tự đóng <w:sdtContent/> (ô trống sau khi từ chối sửa đổi)
  const ctM = /<w:sdtContent>([\s\S]*)<\/w:sdtContent>|<w:sdtContent\s*\/>/.exec(sdt);
  if (!prM || !ctM) return null;
  const placeholder = /<w:showingPlcHdr\/>/.test(prM[0]);
  let pr = prM[0].replace(/<w:showingPlcHdr\/>/g, "");
  const content = dropFormatChanges(ctM[1] ?? "");
  // Ô chứa đoạn (w:p bên trong, kể cả <w:p/>) hay nằm trong một đoạn; nội dung rỗng thì xét ô có đang nằm trong w:p không
  const opened = (src.slice(0, pos.start).match(/<w:p(?=[\s>])/g) || []).length - (src.slice(0, pos.start).match(/<\/w:p>/g) || []).length;
  const block = content ? /<w:p(?=[\s>/])/.test(content) : opened <= 0;
  const props = firstRunProps(content, placeholder);
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(content)?.[0] || "";
  let inner: string;
  if (v.checked !== undefined) {
    if (!/<w14:checkbox>/.test(pr)) return null;
    pr = /<w14:checked\b/.test(pr)
      ? pr.replace(/<w14:checked w14:val="[^"]*"\/>/, `<w14:checked w14:val="${v.checked ? 1 : 0}"/>`)
      : pr.replace(/<w14:checkbox>/, `<w14:checkbox><w14:checked w14:val="${v.checked ? 1 : 0}"/>`);
    const st = new RegExp(`<w14:${v.checked ? "checkedState" : "uncheckedState"}\\b[^>]*/>`).exec(pr)?.[0] || "";
    const code = /w14:val="([0-9A-Fa-f]+)"/.exec(st)?.[1];
    const font = /w14:font="([^"]*)"/.exec(st)?.[1];
    const glyph = code ? String.fromCodePoint(parseInt(code, 16)) : v.checked ? "☒" : "☐";
    inner = runOf(withFont(props, font), glyph);
    if (block) inner = `<w:p>${pPr}${inner}</w:p>`;
  } else {
    const lines = String(v.text ?? "").split(/\r\n|\r|\n/);
    if (block) {
      inner = lines.map((l) => `<w:p>${pPr}${runOf(props, l)}</w:p>`).join("");
    } else {
      inner = lines.map((l, k) => (k ? `<w:r>${props}<w:br/></w:r>` : "") + runOf(props, l)).join("");
    }
    if (v.date) {
      const iso = `${v.date.y}-${String(v.date.m).padStart(2, "0")}-${String(v.date.d).padStart(2, "0")}T00:00:00Z`;
      pr = pr.replace(/<w:date\b([^>]*?)(\/?)>/, (_m, attrs: string, slash: string) => `<w:date${attrs.replace(/\s*w:fullDate="[^"]*"/, "")} w:fullDate="${iso}"${slash}>`);
    }
  }
  const sdtNew =
    sdt.slice(0, prM.index) +
    pr +
    sdt.slice(prM.index + prM[0].length).replace(/<w:sdtContent>[\s\S]*<\/w:sdtContent>|<w:sdtContent\s*\/>/, () => `<w:sdtContent>${inner}</w:sdtContent>`);
  return src.slice(0, pos.start) + sdtNew + src.slice(pos.end);
}

/** Ô chọn: OOXML mới của ô với mục `item` được chọn (giữ lại cho kiểm thử cũ; dùng fillControlOoxml). */
export function selectDropdownOoxml(xml: string, item: string): string | null {
  return fillControlOoxml(xml, { text: item });
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
  // Định dạng có chữ tháng / thứ, chữ trong ngoặc, giờ phút giây: viết dd/MM/yyyy cho chắc (soát mã Đợt 8, #15)
  const f = fmt && !/MMM|ddd|['"]|[Hhms]/.test(fmt) ? fmt : "dd/MM/yyyy";
  const two = (n: number) => String(n).padStart(2, "0");
  return f.replace(/yyyy|yy|MM|M|dd|d/g, (t) =>
    t === "yyyy" ? String(x.y) : t === "yy" ? two(x.y % 100) : t === "MM" ? two(x.m) : t === "M" ? String(x.m) : t === "dd" ? two(x.d) : String(x.d)
  );
}
