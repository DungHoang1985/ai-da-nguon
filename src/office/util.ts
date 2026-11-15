/** Hàm thuần (không cần Office) dùng cho các công cụ — có kiểm thử đơn vị. */

export function colToLetters(col: number): string {
  // col: 0-based
  let s = "";
  let n = col + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Tách "Sheet 1!A1:B2" / "'My Sheet'!A1" thành { sheet, address } */
export function splitAddress(full: string, sheet?: string): { sheet?: string; address: string } {
  const s = String(full || "").trim();
  const m = s.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
  if (m) return { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), address: m[3].replace(/\$/g, "") };
  return { sheet: sheet || undefined, address: s.replace(/\$/g, "") };
}

/** Ô trên-trái của địa chỉ, vd "B2:D9" → { col: 1, row: 1 } */
export function topLeft(address: string): { col: number; row: number } {
  const a = splitAddress(address).address.split(":")[0];
  const m = a.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return { col: 0, row: 0 };
  return { col: lettersToCol(m[1]), row: parseInt(m[2], 10) - 1 };
}

export function cellAddress(col: number, row: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

/** Chuẩn hóa mảng 2 chiều: đủ hàng, mỗi hàng cùng số cột. */
export function normalize2D(values: any): any[][] {
  let v = values;
  if (typeof v === "string") {
    try {
      v = JSON.parse(v);
    } catch {
      v = [[v]];
    }
  }
  if (!Array.isArray(v)) v = [[v]];
  if (v.length && !Array.isArray(v[0])) v = [v];
  const width = Math.max(1, ...v.map((r: any[]) => (Array.isArray(r) ? r.length : 1)));
  return v.map((r: any) => {
    const row = Array.isArray(r) ? r : [r];
    const out = row.map((x: any) => (x === null || x === undefined ? "" : typeof x === "object" ? JSON.stringify(x) : x));
    while (out.length < width) out.push("");
    return out;
  });
}

export function sanitizeSheetName(name: string): string {
  const s = String(name || "Sheet")
    .replace(/[\\/?*[\]:]/g, " ")
    .replace(/^'+|'+$/g, "")
    .trim();
  return (s || "Sheet").slice(0, 31);
}

/**
 * Tìm phần khác nhau duy nhất giữa câu cũ và câu mới (theo ranh giới từ),
 * để chỉ sửa đúng chỗ đó → Track Changes gọn.
 */
export function diffMiddle(oldText: string, newText: string): { prefix: string; removed: string; added: string; suffix: string } {
  if (oldText === newText) return { prefix: oldText, removed: "", added: "", suffix: "" };
  let p = 0;
  const maxP = Math.min(oldText.length, newText.length);
  while (p < maxP && oldText[p] === newText[p]) p++;
  let s = 0;
  const maxS = Math.min(oldText.length - p, newText.length - p);
  while (s < maxS && oldText[oldText.length - 1 - s] === newText[newText.length - 1 - s]) s++;
  const ns = (c: string | undefined) => c !== undefined && /\S/.test(c);
  // Lùi về ranh giới từ (chỉ khi điểm cắt nằm giữa một từ) để thay trọn từ
  while (p > 0 && ns(oldText[p - 1]) && (ns(oldText[p]) || ns(newText[p]))) p--;
  const midS = (k: number) => {
    const bo = oldText.length - k - 1;
    const bn = newText.length - k - 1;
    return ns(oldText[oldText.length - k]) && ((bo >= p && ns(oldText[bo])) || (bn >= p && ns(newText[bn])));
  };
  while (s > 0 && midS(s)) s--;
  return {
    prefix: oldText.slice(0, p),
    removed: oldText.slice(p, oldText.length - s),
    added: newText.slice(p, newText.length - s),
    suffix: oldText.slice(oldText.length - s),
  };
}

/** Word coi ^ là ký tự đặc biệt trong tìm kiếm */
/**
 * Số chỗ tối đa được thay trong một lần gọi. Model hay truyền 0 hoặc -1 với ý "tất cả":
 * giá trị không phải số nguyên dương thì dùng mặc định, không được ra số âm (thay 0 chỗ mà vẫn báo "đã thay").
 */
export function replaceLimit(v: any, def = 50): number {
  const n = Math.floor(Number(v));
  return Number.isFinite(n) && n >= 1 ? n : def;
}

export function escapeWordSearch(s: string): string {
  return s.replace(/\^/g, "^^");
}

export type LineStyle = "Normal" | "Heading1" | "Heading2" | "Heading3" | "ListBullet" | "ListNumber";

/** Đọc ký hiệu đầu dòng kiểu Markdown để chọn kiểu đoạn trong Word. */
export function parseMarkdownLine(line: string): { text: string; style: LineStyle } {
  let m = line.match(/^\s*(#{1,3})\s+(.*)$/);
  if (m) return { text: m[2], style: (`Heading${m[1].length}` as LineStyle) };
  m = line.match(/^\s*[-*•]\s+(.*)$/);
  if (m) return { text: m[1], style: "ListBullet" };
  m = line.match(/^\s*\d+[.)]\s+(.*)$/);
  if (m) return { text: m[1], style: "ListNumber" };
  return { text: line.replace(/\*\*(.*?)\*\*/g, "$1"), style: "Normal" };
}

export function truncate(s: any, n: number): string {
  const t = String(s ?? "");
  return t.length > n ? t.slice(0, n) + "…" : t;
}

/** Ngân sách độ dài JSON cho kết quả công cụ đọc: thấp hơn mức vòng lặp tác tử cắt (30.000 ký tự) để kết quả
 *  không bao giờ bị cắt giữa chừng mà AI không biết (lỗi Word PH3 của rà soát 15/11/2026). */
export const RESULT_JSON_BUDGET = 24000;

/** Lấy lần lượt phần tử cho tới khi JSON vượt ngân sách. Phần tử đầu tiên quá dài thì rút gọn bằng shrink (nếu có). */
export function takeWithinBudget<T>(items: T[], budget: number, shrink?: (x: T, room: number) => T): { picked: T[]; truncated: boolean } {
  const picked: T[] = [];
  let used = 2;
  for (let k = 0; k < items.length; k++) {
    let x = items[k];
    let len = JSON.stringify(x).length + 1;
    if (used + len > budget) {
      if (!picked.length && shrink) {
        x = shrink(x, Math.max(100, budget - used - 200));
        len = JSON.stringify(x).length + 1;
        picked.push(x);
        return { picked, truncated: items.length > 1 };
      }
      return { picked, truncated: true };
    }
    picked.push(x);
    used += len;
  }
  return { picked, truncated: false };
}

/** Kết quả word_read_paragraphs: tính theo độ dài JSON thật; next_start và truncated đặt lên đầu. */
export function packParagraphs(all: { text: string; style: string }[], start: number, end: number, budget = RESULT_JSON_BUDGET) {
  const s0 = Math.max(0, Math.floor(Number(start)) || 0);
  const e0 = Math.min(all.length, Math.max(s0, Math.floor(Number(end)) || s0 + 200));
  const slice = all.slice(s0, e0).map((p, k) => ({ i: s0 + k, style: String(p.style ?? ""), text: String(p.text ?? "") }));
  const { picked, truncated } = takeWithinBudget(slice, budget - 400, (x, room) => ({
    ...x,
    text: x.text.slice(0, room) + "… [đoạn quá dài, đã cắt phần sau]",
  }));
  const last = picked.length ? picked[picked.length - 1].i : s0 - 1;
  const more = last + 1 < all.length;
  return {
    ok: true,
    total_paragraphs: all.length,
    ...(more ? { next_start: last + 1 } : {}),
    ...(truncated ? { truncated: true, note: `Kết quả dài nên dừng ở đoạn ${last}. Gọi lại với start = next_start để đọc tiếp.` } : {}),
    read_from: s0,
    read_to: last,
    paragraphs: picked,
  };
}

/** Số đoạn của nhận xét / sửa đổi: tìm theo định danh đoạn (uniqueLocalId, WordApi 1.6) — đúng cả khi có đoạn trùng nội dung;
 *  Word cũ không có định danh thì so chữ như trước. */
export function paragraphIndex(anchorId: string | undefined, anchorText: string, ids: string[] | null, texts: string[]): number {
  if (ids && anchorId) {
    const i = ids.indexOf(anchorId);
    if (i >= 0) return i;
  }
  return texts.indexOf(anchorText);
}

/** Trạng thái văn bản đưa vào lời nhắc hệ thống: vùng chọn, chế độ theo dõi, nhận xét lên đầu; danh sách tiêu đề cắt theo phần tử
 *  (không cắt chuỗi JSON) cho vừa ngân sách. */
export function overviewForPrompt(ov: any, budget = 6000): string {
  if (!ov || typeof ov !== "object") return JSON.stringify(ov ?? null);
  const { selection, track_changes, comments, paragraphs, characters, headings, first_paragraphs, ...rest } = ov;
  const allHeadings: any[] = Array.isArray(headings) ? headings : [];
  let sel = selection;
  let first: any[] = Array.isArray(first_paragraphs) ? first_paragraphs : [];
  let hs = allHeadings;
  const build = () =>
    JSON.stringify({
      selection: sel,
      track_changes,
      comments,
      paragraphs,
      characters,
      ...rest,
      ...(hs.length < allHeadings.length ? { headings_truncated: true, headings_total: allHeadings.length } : {}),
      headings: hs,
      first_paragraphs: first,
    });
  let s = build();
  while (s.length > budget && hs.length) {
    hs = hs.slice(0, Math.floor(hs.length * 0.8));
    s = build();
  }
  while (s.length > budget && first.length) {
    first = first.slice(0, first.length - 1);
    s = build();
  }
  if (s.length > budget && typeof sel === "string") {
    sel = truncate(sel, Math.max(200, sel.length - (s.length - budget) - 20));
    s = build();
  }
  return s;
}
