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
