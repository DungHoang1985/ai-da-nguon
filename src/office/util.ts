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

/** Tách "A1:B2,'Quý 1, 2026'!C3:D4" thành từng vùng: dấu phẩy nằm trong tên sheet đặt trong nháy đơn không phải dấu ngăn. */
export function splitAreas(address: any): string[] {
  const s = String(address ?? "");
  const out: string[] = [];
  let cur = "";
  let quoted = false;
  for (const ch of s) {
    if (ch === "'") quoted = !quoted;
    if (ch === "," && !quoted) {
      if (cur.trim()) out.push(cur.trim());
      cur = "";
    } else cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
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

export interface TableFillPlan {
  /** Ô có sẵn sẽ được ghi: hàng, cột, giá trị mới, giá trị cũ */
  writes: { row: number; col: number; value: string; old: string }[];
  /** Hàng mới thêm vào cuối bảng */
  extraRows: string[][];
  /** Các ô đang có dữ liệu khác với giá trị sắp ghi */
  overwrites: { row: number; col: number; value: string; old: string }[];
}

/** Tính trước việc điền bảng Word: ô nào được ghi, hàng nào thêm mới, ô nào đang có dữ liệu sẽ bị ghi đè.
 *  current = giá trị hiện tại của bảng; start_row mặc định 1 (bỏ hàng tiêu đề), start_col mặc định 0. */
export function planTableFill(current: string[][], values: any, startRow?: any, startCol?: any): TableFillPlan {
  const vals = normalize2D(values).map((r) => r.map((v) => String(v)));
  const rowCount = current.length;
  const cols = current[0]?.length || 0;
  const r0 = startRow === undefined || startRow === null ? 1 : Math.max(0, Number(startRow) || 0);
  const c0 = Math.max(0, Number(startCol) || 0);
  const plan: TableFillPlan = { writes: [], extraRows: [], overwrites: [] };
  vals.forEach((row, i) => {
    const r = r0 + i;
    if (r < rowCount) {
      row.forEach((v, j) => {
        if (c0 + j >= cols) return;
        const old = String(current[r]?.[c0 + j] ?? "");
        const w = { row: r, col: c0 + j, value: v, old };
        plan.writes.push(w);
        if (old.trim() && old.trim() !== v.trim()) plan.overwrites.push(w);
      });
    } else {
      const full = Array(cols).fill("");
      row.forEach((v, j) => c0 + j < cols && (full[c0 + j] = v));
      plan.extraRows.push(full);
    }
  });
  return plan;
}

export function escapeWordSearch(s: string): string {
  return s.replace(/\^/g, "^^");
}

export type LineStyle = "Normal" | "Heading1" | "Heading2" | "Heading3" | "ListBullet" | "ListNumber";

/** Đọc ký hiệu đầu dòng kiểu Markdown để chọn kiểu đoạn trong Word.
 *  numbered = false: dòng bắt đầu bằng "số." được giữ nguyên chữ (số điều khoản gõ tay như "5. Thanh toán", "2026. Năm tài chính"). */
export function parseMarkdownLine(line: string, numbered = true): { text: string; style: LineStyle } {
  let m = line.match(/^\s*(#{1,3})\s+(.*)$/);
  if (m) return { text: m[2].replace(/\*\*(.*?)\*\*/g, "$1"), style: (`Heading${m[1].length}` as LineStyle) };
  m = line.match(/^\s*[-*•]\s+(.*)$/);
  if (m) return { text: m[1], style: "ListBullet" };
  m = numbered ? line.match(/^\s*\d+[.)]\s+(.*)$/) : null;
  if (m) return { text: m[1], style: "ListNumber" };
  return { text: line.replace(/\*\*(.*?)\*\*/g, "$1"), style: "Normal" };
}

/** Các dòng có phải một danh sách đánh số Markdown thật không: có ít nhất hai dòng "số." và dòng đầu là 1.
 *  Một dòng lẻ "5. Thanh toán" là số điều khoản gõ tay, không được đổi thành danh sách đánh lại từ 1. */
export function isMarkdownNumberedList(lines: string[]): boolean {
  const nums = lines.map((l) => l.match(/^\s*(\d+)[.)]\s+\S/)).filter(Boolean) as RegExpMatchArray[];
  return nums.length >= 2 && Number(nums[0][1]) === 1;
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
export function packParagraphs(
  all: { text: string; style: string; num?: string; changed?: boolean; deleted?: boolean; notes?: string[] }[],
  start: number,
  end: number,
  budget = RESULT_JSON_BUDGET
) {
  const s0 = Math.max(0, Math.floor(Number(start)) || 0);
  const e0 = Math.min(all.length, Math.max(s0, Math.floor(Number(end)) || s0 + 200));
  // num = số điều khoản (kể cả số do Word tự đánh, vốn không nằm trong chữ của đoạn); changed = đoạn đang có sửa đổi chờ duyệt;
  // deleted = đoạn đã bị xóa có theo dõi
  const slice = all.slice(s0, e0).map((p, k) => ({
    i: s0 + k,
    ...(p.num ? { num: p.num } : {}),
    style: String(p.style ?? ""),
    text: String(p.text ?? ""),
    ...(p.changed ? { changed: true } : {}),
    ...(p.deleted ? { deleted: true } : {}),
    // notes = chú thích cuối trang / cuối văn bản neo vào đoạn này
    ...(p.notes?.length ? { notes: p.notes } : {}),
  }));
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

const foldAccents = (s: string) =>
  s
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");

/** Tìm một cụm chữ trong danh sách đoạn (bản hiện tại). Trả mỗi đoạn khớp một dòng: vị trí lần khớp đầu, số lần khớp.
 *  ignoreAccents: bỏ dấu tiếng Việt và hoa thường ở cả hai phía (người dùng gõ không dấu). */
export function searchRows(
  rows: { i: number; text: string; deleted?: boolean }[],
  query: string,
  opts: { matchCase?: boolean; wholeWord?: boolean; ignoreAccents?: boolean } = {}
): { total: number; hits: { i: number; at: number; len: number; n: number }[] } {
  const norm = (s: string) => {
    let t = s.replace(/ /g, " ");
    if (opts.ignoreAccents) t = foldAccents(t.normalize("NFC")).toLowerCase();
    else if (!opts.matchCase) t = t.toLowerCase();
    return t;
  };
  const q = norm(query.replace(/\s+/g, " ").trim());
  const hits: { i: number; at: number; len: number; n: number }[] = [];
  let total = 0;
  if (!q) return { total, hits };
  const isWord = (c: string | undefined) => !!c && /[\p{L}\p{N}]/u.test(c);
  for (const r of rows) {
    if (r.deleted || !r.text) continue;
    const t = norm(r.text);
    // Bỏ dấu và đổi hoa thường có thể làm lệch độ dài chuỗi: khi đó chỉ báo đoạn khớp, đặt vị trí về đầu đoạn
    const aligned = t.length === r.text.length;
    let n = 0;
    let first = -1;
    for (let pos = 0; ; ) {
      const k = t.indexOf(q, pos);
      if (k < 0) break;
      pos = k + q.length;
      if (opts.wholeWord && (isWord(t[k - 1]) || isWord(t[k + q.length]))) continue;
      if (first < 0) first = k;
      n++;
    }
    if (n) {
      total += n;
      hits.push({ i: r.i, at: aligned ? first : 0, len: aligned ? q.length : 0, n });
    }
  }
  return { total, hits };
}

/** Tìm theo TỪNG TỪ khi không đoạn nào chứa nguyên cụm: xếp các đoạn theo số từ của câu hỏi có mặt (không phân biệt hoa thường và dấu).
 *  Chỉ nhận đoạn có ít nhất nửa số từ. Mỗi đoạn trả câu chứa nhiều từ khớp nhất để làm trích dẫn. */
export function searchRowsByWords(
  rows: { i: number; text: string; deleted?: boolean }[],
  query: string
): { words: string[]; hits: { i: number; matched: number; sentence: string }[] } {
  const fold = (s: string) => foldAccents(s.normalize("NFC")).toLowerCase();
  const words = Array.from(new Set((fold(query).match(/[\p{L}\p{N}]+/gu) || []).filter((w) => w.length >= 2)));
  const hits: { i: number; matched: number; sentence: string }[] = [];
  if (words.length < 2) return { words, hits };
  const has = (t: string, w: string) => new RegExp(`(?<![\\p{L}\\p{N}])${w}(?![\\p{L}\\p{N}])`, "u").test(t);
  for (const r of rows) {
    if (r.deleted || !r.text) continue;
    const t = fold(r.text);
    const matched = words.filter((w) => has(t, w)).length;
    if (matched * 2 < words.length || matched < 2) continue;
    let best = "";
    let bestN = -1;
    for (const sent of r.text.split(/(?<=[.;:!?])\s+/)) {
      const fs = fold(sent);
      const n = words.filter((w) => has(fs, w)).length;
      if (n > bestN) {
        bestN = n;
        best = sent;
      }
    }
    hits.push({ i: r.i, matched, sentence: best });
  }
  hits.sort((a, b) => b.matched - a.matched || a.i - b.i);
  return { words, hits };
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
