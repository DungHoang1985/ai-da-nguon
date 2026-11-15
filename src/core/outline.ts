/** Mục lục điều khoản của một văn bản: số do Word tự đánh (nhãn danh sách) và số gõ tay, cấp, tiêu đề, phạm vi đoạn.
 *  Hàm thuần, không cần Office — kiểm thử bằng npm test (test/unit/outline.test.ts).
 *  Kết quả thăm dò 15/11/2026: Word trả nhãn của từng cấp riêng lẻ ("Điều 1.", "1.1.", "a)") nên phải tự ghép với cấp cha. */

export interface ParaRow {
  /** Chỉ số đoạn trong thân văn bản (0-based) */
  i: number;
  /** Chữ của đoạn theo bản hiện tại (không gồm chữ đã bị xóa có theo dõi) */
  text: string;
  style?: string;
  /** 1–9 = cấp tiêu đề; 10 hoặc không có = thân văn bản */
  outlineLevel?: number;
  /** Nhãn số do Word tự đánh ("Điều 1.", "1.1.", "a)"); không có = đoạn không thuộc danh sách */
  listString?: string;
  listLevel?: number;
  listId?: number;
  inTable?: boolean;
  /** Đoạn đã bị xóa có theo dõi (bản hiện tại rỗng) */
  deleted?: boolean;
}

export type ClauseKind = "chapter" | "section" | "article" | "dotted" | "single" | "letter" | "roman" | "heading";

export interface OutlineItem {
  /** Chỉ số đoạn mở đầu điều khoản */
  i: number;
  /** Số điều đầy đủ để người đọc dùng: "Điều 5", "5.3", "5.3(a)", "Chương II"; rỗng với tiêu đề không đánh số */
  num: string;
  /** Khóa chuẩn để tra: "5", "5.3", "5.3.a" */
  key: string;
  /** Nhãn đúng như văn bản hiện ("Điều 5.", "5.3.", "a)") */
  label: string;
  /** Độ sâu trong cây điều khoản, 0 = cấp cao nhất */
  level: number;
  title: string;
  /** list = số do Word tự đánh; typed = số gõ tay; heading = tiêu đề không đánh số */
  source: "list" | "typed" | "heading";
  kind: ClauseKind;
  /** Chỉ số đoạn cuối cùng thuộc điều khoản này (gồm cả các khoản con) */
  end: number;
  /** Vị trí của mục cha trong mảng kết quả; -1 = không có */
  parent: number;
}

interface Parsed {
  kind: ClauseKind;
  label: string;
  /** Phần chữ sau số (với số gõ tay) */
  rest: string;
  word?: string;
  id?: string;
  n?: number;
  parts?: number[];
  letter?: string;
}

const ROMAN = /^(?=[ivxlcdm])m{0,3}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$/i;
const CHAPTER_WORDS = "chương|phần|phụ lục|chapter|part|appendix|annex|schedule";
const ARTICLE_WORDS = "điều|article|clause|section";
const LETTERS = "abcdđefghijklmnopqrstuvwxyz";
/** Đơn vị đứng ngay sau một số thập phân: dấu hiệu đó là số đo, không phải số khoản */
const UNIT = /^(triệu|tỷ|tỉ|nghìn|ngàn|trăm|đồng|vnđ|vnd|usd|eur|m2|m²|m3|m³|ha|kg|tấn|km|lần|năm|tháng|tuần|ngày|giờ|phút|%|‰)(?![\p{L}\p{N}])/iu;

/** Giá trị của số La Mã (I = 1, IV = 4…) */
export function romanValue(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  let total = 0;
  const t = s.toLowerCase();
  for (let k = 0; k < t.length; k++) {
    const a = v[t[k]] || 0;
    const b = v[t[k + 1]] || 0;
    total += a < b ? -a : a;
  }
  return total;
}

const cap = (s: string) => (s ? s[0].toUpperCase() + s.slice(1).toLowerCase() : s);
const cleanTitle = (s: string) => s.replace(/\s+/g, " ").trim();

/** Số có dấu chấm kiểu điều khoản ("5.3", "5.3.1") — loại số tiền ("1.500.000"), ngày ("15.11.2026") */
function dottedParts(s: string): number[] | null {
  const parts = s.split(".");
  if (parts.length < 2 || parts.length > 6) return null;
  if (parts.some((p, k) => !/^\d{1,3}$/.test(p) || (k > 0 && p.length === 3))) return null;
  return parts.map(Number);
}

/** Đọc số ở đầu chuỗi. isLabel = chuỗi là nhãn danh sách của Word (không có chữ theo sau). */
export function parseNumber(raw: string, isLabel: boolean): Parsed | null {
  const s = raw.replace(/ /g, " ").trim();
  if (!s) return null;
  let m: RegExpMatchArray | null;

  // Số gõ tay: phần sau số phải trông như tiêu đề chứ không phải phần tiếp của một câu. "Điều 3 của Hợp đồng này quy định…",
  // "Phụ lục 01 là bộ phận không tách rời…" là câu dẫn chiếu, không phải điều khoản (soát mã độc lập 15/11/2026, mục C2).
  // Nhận khi: hết dòng; có dấu phân cách sau số (. : – -); hoặc phần sau ngắn, viết hoa chữ đầu, không kết thúc như một câu.
  const titleLike = (sep: string | undefined, rest: string) =>
    isLabel || !rest || !!(sep && sep.trim()) || (rest.length <= 80 && /^\p{Lu}/u.test(rest) && !/[.;,:]$/.test(rest));

  // Chương II, Phần 1, Phụ lục 01
  m = s.match(new RegExp(`^(${CHAPTER_WORDS})\\s+([0-9]{1,3}|[ivxlcdm]+)(?![\\p{L}\\p{N}])(\\s*[.:–-])?\\s*(.*)$`, "iu"));
  if (m && (/^\d+$/.test(m[2]) || ROMAN.test(m[2])) && (isLabel || s.length <= 200) && titleLike(m[3], m[4])) {
    return { kind: "chapter", label: s.slice(0, s.length - m[4].length).trim(), rest: m[4], word: cap(m[1]), id: m[2].toUpperCase() };
  }
  // Mục 2
  m = s.match(/^(mục|tiểu mục)\s+([0-9]{1,3}|[ivxlcdm]+)(?![\p{L}\p{N}])(\s*[.:–-])?\s*(.*)$/iu);
  if (m && (/^\d+$/.test(m[2]) || ROMAN.test(m[2])) && (isLabel || s.length <= 200) && titleLike(m[3], m[4])) {
    return { kind: "section", label: s.slice(0, s.length - m[4].length).trim(), rest: m[4], word: cap(m[1]), id: m[2].toUpperCase() };
  }
  // Điều 5.3 / Điều 5
  m = s.match(new RegExp(`^(${ARTICLE_WORDS})\\s+(\\d{1,3}(?:\\.\\d{1,3})*)(?![\\p{L}\\p{N}])(\\s*[.:–-])?\\s*(.*)$`, "iu"));
  if (m && titleLike(m[3], m[4])) {
    const label = s.slice(0, s.length - m[4].length).trim();
    if (m[2].includes(".")) {
      const parts = dottedParts(m[2]);
      if (parts) return { kind: "dotted", label, rest: m[4], parts };
    } else return { kind: "article", label, rest: m[4], word: cap(m[1]), n: Number(m[2]) };
  }
  // 5.3 / 5.3. / 5.3.1  — "1.5 triệu đồng", "2.5%" ở đầu đoạn là số đo, không phải số khoản
  m = s.match(/^(\d{1,3}(?:\.\d{1,3})+)\.?\)?(?:\s+(.*))?$/);
  if (m && (isLabel || m[2])) {
    const parts = dottedParts(m[1]);
    if (parts && (isLabel || !UNIT.test(m[2] || ""))) return { kind: "dotted", label: s.slice(0, s.length - (m[2] || "").length).trim(), rest: m[2] || "", parts };
  }
  // 5. / 5) / (5)
  m = s.match(/^(?:\((\d{1,3})\)|(\d{1,3})[.)])(?:\s+(.*))?$/);
  if (m && (isLabel || m[3])) {
    return { kind: "single", label: s.slice(0, s.length - (m[3] || "").length).trim(), rest: m[3] || "", n: Number(m[1] ?? m[2]) };
  }
  // I. / II. (chữ số La Mã hoa) — mục lớn
  m = s.match(/^([IVXLCDM]{1,6})[.)](?:\s+(.*))?$/);
  if (m && ROMAN.test(m[1]) && (isLabel || (m[2] && s.length <= 150))) {
    return { kind: "chapter", label: s.slice(0, s.length - (m[2] || "").length).trim(), rest: m[2] || "", word: "", id: m[1] };
  }
  // (ii) / ii) / ii.  — chữ số La Mã thường (một chữ cái i, v, x xử lý ở bước dựng cây)
  m = s.match(/^(?:\(([ivxl]{1,6})\)|([ivxl]{1,6})\))(?:\s+(.*))?$/) || (isLabel ? s.match(/^()([ivxl]{2,6})\.()$/) : null);
  if (m && ROMAN.test(m[1] || m[2]) && (isLabel || m[3])) {
    const r = (m[1] || m[2]).toLowerCase();
    const label = s.slice(0, s.length - (m[3] || "").length).trim();
    if (r.length === 1) return { kind: "letter", label, rest: m[3] || "", letter: r };
    return { kind: "roman", label, rest: m[3] || "", id: r };
  }
  // a) / (a) / đ)  — và "a." chỉ khi là nhãn danh sách
  m = s.match(/^(?:\(([a-zđ])\)|([a-zđ])\))(?:\s+(.*))?$/) || (isLabel ? s.match(/^()([a-zđ])\.()$/) : null);
  if (m && (isLabel || m[3])) {
    return { kind: "letter", label: s.slice(0, s.length - (m[3] || "").length).trim(), rest: m[3] || "", letter: m[1] || m[2] };
  }
  // A. / B. — mục lớn đánh chữ hoa (chỉ nhận khi là nhãn danh sách)
  if (isLabel && (m = s.match(/^([A-Z])[.)]$/))) return { kind: "section", label: s, rest: "", word: "", id: m[1] };
  return null;
}

/** Số hiển thị từ khóa: "5.3.a.ii" → "5.3(a)(ii)" */
export function keyToNum(key: string): string {
  return key
    .split(".")
    .map((p, k) => (/^\d+$/.test(p) ? (k ? "." : "") + p : `(${p})`))
    .join("");
}

interface Node extends OutlineItem {
  n?: number;
  letter?: string;
  listId?: number;
  listLevel?: number;
  ol: number;
}

const NUMERIC: ClauseKind[] = ["article", "dotted", "single"];

/** Dựng mục lục điều khoản từ danh sách đoạn. */
export function buildOutline(rows: ParaRow[]): OutlineItem[] {
  const items: Node[] = [];
  /** Đường đi từ gốc tới mục vừa thêm (vị trí trong items) */
  let stack: number[] = [];
  const pathOf = (k: number): number[] => {
    const out: number[] = [];
    for (let x = k; x >= 0; x = items[x].parent) out.unshift(x);
    return out;
  };
  /** Mục gần nhất trên đường đi (từ trong ra ngoài) thỏa điều kiện */
  const nearest = (ok: (it: Node) => boolean): number => {
    for (let s = stack.length - 1; s >= 0; s--) if (ok(items[stack[s]])) return stack[s];
    return -1;
  };
  /** Khóa số đã gặp và tiền tố của các khóa có dấu chấm đã gặp (để nhận số gõ tay theo ngữ cảnh) */
  const seenKeys = new Set<string>();
  const seenPrefixes = new Set<string>();
  let lastTypedRoman = 0;
  const sameList = (a: { source: string; listId?: number }, b: { source: string; listId?: number }) =>
    (a.source === "list") === (b.source === "list") && (a.listId === undefined || b.listId === undefined || a.listId === b.listId);

  for (const row of rows) {
    if (row.inTable || row.deleted) continue;
    const text = cleanTitle(row.text || "");
    const ol = row.outlineLevel && row.outlineLevel >= 1 && row.outlineLevel <= 9 ? row.outlineLevel : 10;
    let p: Parsed | null = null;
    let source: OutlineItem["source"] = "typed";
    if (row.listString && row.listString.trim()) {
      p = parseNumber(row.listString, true);
      if (p) {
        source = "list";
        p.rest = text;
      } else if (/[\p{L}\p{N}]/u.test(row.listString)) p = null; // nhãn lạ: bỏ qua
    }
    if (!p && !(row.listString && row.listString.trim())) {
      p = text ? parseNumber(text, false) : null;
      // Số gõ tay phải khớp ngữ cảnh, nếu không thì là chữ thường của đoạn:
      if (p && p.kind === "dotted") {
        // "5.3" chỉ là khoản khi đã có "5" hoặc một khoản "5.x" đứng trước, hoặc là khoản đầu tiên ("x.1")
        const parts = p.parts as number[];
        const prefix = parts.slice(0, -1).join(".");
        if (!seenKeys.has(prefix) && !seenPrefixes.has(prefix) && parts[parts.length - 1] !== 1) p = null;
      } else if (p && p.kind === "chapter" && !p.word) {
        // "I.", "II." gõ tay: phải bắt đầu từ I và nối tiếp nhau ("C. Công ty X", "D. …" là danh sách chữ cái)
        const v = romanValue(String(p.id));
        if (v === 1 || v === lastTypedRoman + 1) lastTypedRoman = v;
        else p = null;
      }
    }
    if (!p) {
      if (ol <= 9 && text) p = { kind: "heading", label: "", rest: text };
      else continue;
      source = "heading";
    }
    const me = { source, listId: row.listId };
    let parent = -1;
    let key = "";
    let num = "";
    let kind = p.kind;
    let letter = p.letter;

    // "i)", "v)", "x)" là chữ cái nếu nối tiếp h, u, w; ngược lại là số La Mã
    if (kind === "letter" && letter && "ivx".includes(letter)) {
      const sib = nearest((it) => it.kind === "letter" && sameList(it, me));
      const prev = LETTERS[LETTERS.indexOf(letter) - 1];
      const openRoman = nearest((it) => it.kind === "roman" && sameList(it, me)) >= 0;
      if (openRoman || !(sib >= 0 && items[sib].letter === prev)) {
        kind = "roman";
        p = { ...p, kind, id: letter };
        letter = undefined;
      }
    }

    if (kind === "chapter") {
      num = p.word ? `${p.word} ${p.id}` : String(p.id);
      key = num.toLowerCase();
    } else if (kind === "section") {
      parent = nearest((it) => it.kind === "chapter");
      num = p.word ? `${p.word} ${p.id}` : String(p.id);
      key = num.toLowerCase();
    } else if (kind === "article") {
      parent = nearest((it) => it.kind === "chapter" || it.kind === "section" || (it.kind === "heading" && it.ol < ol));
      key = String(p.n);
      num = `${p.word} ${p.n}`;
    } else if (kind === "dotted") {
      const parts = p.parts as number[];
      key = parts.join(".");
      const prefix = parts.slice(0, -1).join(".");
      parent = nearest((it) => NUMERIC.includes(it.kind) && it.key === prefix);
      if (parent < 0) parent = nearest((it) => NUMERIC.includes(it.kind) && it.key === String(parts[0]));
      if (parent < 0) parent = nearest((it) => it.kind === "chapter" || it.kind === "section" || it.kind === "heading");
      num = key;
    } else if (kind === "single") {
      const n = p.n as number;
      // Anh em = khoản cùng loại đứng trước có số nhỏ hơn (không đòi đúng n-1: khoản liền trước có thể đã bị xóa có theo dõi hoặc đánh nhảy số)
      const sib = nearest((it) => it.kind === "single" && (it.n as number) < n && sameList(it, me) && (source !== "list" || it.listLevel === row.listLevel));
      if (sib >= 0) parent = items[sib].parent;
      else if (source === "list" && (row.listLevel || 0) > 0) {
        parent = nearest((it) => it.source === "list" && (it.listLevel ?? 0) < (row.listLevel || 0) && sameList(it, me));
        if (parent < 0) parent = nearest((it) => NUMERIC.includes(it.kind));
      } else {
        parent = nearest((it) => NUMERIC.includes(it.kind));
        if (parent < 0) parent = nearest((it) => it.kind === "chapter" || it.kind === "section" || (it.kind === "heading" && it.ol < ol));
      }
      const pk = parent >= 0 && NUMERIC.includes(items[parent].kind) ? items[parent].key : "";
      key = pk ? `${pk}.${n}` : String(n);
      num = keyToNum(key);
    } else if (kind === "letter" || kind === "roman") {
      const id = String(kind === "letter" ? letter : p.id);
      const sib = nearest((it) => it.kind === kind && sameList(it, me) && (source !== "list" || it.listLevel === row.listLevel));
      if (sib >= 0) parent = items[sib].parent;
      else {
        parent = kind === "roman" ? nearest((it) => it.kind === "letter" || NUMERIC.includes(it.kind)) : nearest((it) => NUMERIC.includes(it.kind));
        if (parent < 0) parent = stack.length ? stack[stack.length - 1] : -1;
      }
      const pk = parent >= 0 && ["article", "dotted", "single", "letter", "roman"].includes(items[parent].kind) ? items[parent].key : "";
      key = pk ? `${pk}.${id}` : id;
      num = pk ? keyToNum(key) : p.label;
    } else {
      // Tiêu đề không đánh số: xếp theo cấp tiêu đề
      parent = nearest((it) => it.ol < ol);
    }

    const node: Node = {
      i: row.i,
      num,
      key,
      label: p.label,
      level: parent >= 0 ? items[parent].level + 1 : 0,
      title: cleanTitle(p.rest).slice(0, 100),
      source,
      kind,
      end: row.i,
      parent,
      n: p.n,
      letter,
      listId: row.listId,
      listLevel: source === "list" ? row.listLevel || 0 : undefined,
      ol,
    };
    items.push(node);
    if (NUMERIC.includes(kind) && key) {
      seenKeys.add(key);
      if (key.includes(".")) seenPrefixes.add(key.slice(0, key.lastIndexOf(".")));
    }
    stack = pathOf(items.length - 1);
  }

  // Phạm vi: tới trước mục kế tiếp cùng cấp hoặc cao hơn
  const last = rows.length ? rows[rows.length - 1].i : -1;
  for (let k = 0; k < items.length; k++) {
    let end = last;
    for (let j = k + 1; j < items.length; j++) {
      if (items[j].level <= items[k].level) {
        end = items[j].i - 1;
        break;
      }
    }
    items[k].end = Math.max(items[k].i, end);
  }
  return items.map(({ i, num, key, label, level, title, source, kind, end, parent }) => ({ i, num, key, label, level, title, source, kind, end, parent }));
}

const stripAccents = (s: string) =>
  s
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");

/** Chuẩn hóa cách người dùng gọi một điều khoản thành khóa: "khoản 2 Điều 5" → "5.2"; "Điều 5.3(a)" → "5.3.a" */
export function clauseKey(query: string): string {
  const q = String(query || "")
    .toLowerCase()
    .replace(/ /g, " ")
    .trim();
  const dieu = q.match(/(?:điều|dieu|article|clause|section)\s+(\d+(?:\.\d+)*)/);
  const khoan = q.match(/(?:khoản|khoan|paragraph)\s+(\d+(?:\.\d+)*)/);
  const diem = q.match(/(?:điểm|diem|point|item)\s+\(?([a-zđ])\)?(?![a-zà-ỹ])/);
  const tiet = q.match(/(?:tiết|tiet)\s+\(?([ivxl]+)\)?/);
  if (khoan || diem || tiet) {
    const parts: string[] = [];
    if (dieu) parts.push(dieu[1]);
    if (khoan) parts.push(dieu && khoan[1].startsWith(dieu[1] + ".") ? khoan[1].slice(dieu[1].length + 1) : khoan[1]);
    if (diem) parts.push(diem[1]);
    if (tiet) parts.push(tiet[1]);
    // "điều 5.3 điểm a": phần số còn lại sau từ "điều" có thể kèm (a)
    return parts.join(".");
  }
  const body = dieu ? q.slice(q.indexOf(dieu[0]) + dieu[0].length - dieu[1].length) : q;
  return body
    .replace(/[()[\],;:]+/g, ".")
    .replace(/\s+/g, ".")
    .replace(/\.{2,}/g, ".")
    .replace(/^\.+|\.+$/g, "");
}

/** Tìm điều khoản theo cách gọi của người dùng. Trả mọi mục khớp (có thể nhiều hơn một khi văn bản đánh số lại từng phần). */
export function findClause(items: OutlineItem[], query: string): OutlineItem[] {
  const q = String(query || "").trim();
  if (!q) return [];
  const flat = stripAccents(q.toLowerCase()).replace(/\s+/g, " ");
  // Chương / Phần / Phụ lục / Mục
  const ch = flat.match(/^(chuong|phan|phu luc|muc|tieu muc|chapter|part|appendix|annex|schedule)\s+([0-9]+|[ivxlcdm]+)$/);
  if (ch) {
    const id = /^\d+$/.test(ch[2]) ? String(Number(ch[2])) : ch[2].toUpperCase();
    return items.filter((it) => {
      if (it.kind !== "chapter" && it.kind !== "section") return false;
      const m = stripAccents(it.num.toLowerCase()).match(/^(.*?)\s*([0-9]+|[ivxlcdm]+)$/);
      if (!m) return false;
      const itId = /^\d+$/.test(m[2]) ? String(Number(m[2])) : m[2].toUpperCase();
      return m[1].trim() === ch[1] && itId === id;
    });
  }
  const key = clauseKey(q);
  if (!key) return [];
  const exact = items.filter((it) => it.kind !== "chapter" && it.kind !== "section" && it.kind !== "heading" && it.key === key);
  if (exact.length) return exact;
  // Số La Mã / chữ cái đứng một mình trùng mục lớn ("II", "A")
  return items.filter((it) => (it.kind === "chapter" || it.kind === "section") && it.num.toLowerCase() === q.toLowerCase());
}

/** Điều khoản trong cùng nhất chứa đoạn p (null nếu đoạn nằm trước mọi điều khoản) */
export function clauseAt(items: OutlineItem[], p: number): OutlineItem | null {
  let lo = 0;
  let hi = items.length - 1;
  let hit = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (items[mid].i <= p) {
      hit = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  for (let k = hit; k >= 0; k = items[k].parent) if (items[k].end >= p) return items[k];
  return null;
}

/** Tên điều khoản để chỉ cho người đọc: "Điều 5", "5.3(a)", hoặc tiêu đề nếu không có số */
export function clauseName(it: OutlineItem | null): string {
  if (!it) return "";
  return it.num || it.title.slice(0, 40);
}

/** Số điều của từng đoạn mở đầu điều khoản: chỉ số đoạn → mục */
export function outlineByParagraph(items: OutlineItem[]): Map<number, OutlineItem> {
  const m = new Map<number, OutlineItem>();
  for (const it of items) if (!m.has(it.i)) m.set(it.i, it);
  return m;
}
