/** Soát nhất quán một văn bản bằng mã (không gọi model, không tốn token): đánh số nhảy hoặc trùng; dẫn chiếu tới điều không
 *  tồn tại; thuật ngữ định nghĩa dùng lệch hoặc dùng mà chưa định nghĩa; một bên được gọi bằng nhiều tên; chỗ trống chưa điền;
 *  số bằng số không khớp bằng chữ. Hàm thuần, không cần Office — kiểm thử bằng npm test (test/unit/dot1b.test.ts).
 *  Đây là các phép dò theo mẫu: kết quả là điểm CẦN XEM, người đọc (hoặc AI) quyết định có phải lỗi không. */
import { OutlineItem, ParaRow, buildOutline, clauseAt, clauseKey, clauseName, findClause, outlineByParagraph, romanValue } from "./outline";

export type IssueKind = "numbering" | "reference" | "term" | "party" | "blank" | "amount";

export interface Issue {
  kind: IssueKind;
  /** Đoạn đầu tiên có điểm cần xem */
  p: number;
  /** Số điều khoản chứa đoạn đó ("" nếu nằm ngoài mọi điều khoản) */
  clause: string;
  message: string;
  /** Trích chữ quanh chỗ đó */
  excerpt?: string;
  /** Các đoạn khác có cùng điểm này */
  also?: number[];
  /** Tổng số chỗ */
  count?: number;
}

export interface ConsistencyResult {
  issues: Issue[];
  /** Thuật ngữ đã định nghĩa và tên các bên nhận ra được (để người đọc biết phép soát dựa vào đâu) */
  terms: string[];
  parties: string[];
}

export const ALL_CHECKS: IssueKind[] = ["numbering", "reference", "term", "party", "blank", "amount"];

const clean = (s: string) =>
  String(s ?? "")
    .replace(/[\u0001\u0002]/g, "")
    .replace(/ /g, " ");
const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const cut = (s: string, n: number) => (s.length > n ? s.slice(0, n) + "…" : s);
/** Trích quanh vị trí at */
function around(text: string, at: number, len: number, pad = 45): string {
  const s = Math.max(0, at - pad);
  const e = Math.min(text.length, at + len + pad);
  return (s > 0 ? "…" : "") + text.slice(s, e) + (e < text.length ? "…" : "");
}

interface Live {
  i: number;
  text: string;
  inTable?: boolean;
}

// ───────────────────────── 1. Đánh số nhảy hoặc trùng ─────────────────────────

const LETTERS_EN = "abcdefghijklmnopqrstuvwxyz";
/** Bảng chữ cái dùng để đánh điểm trong văn bản tiếng Việt (không có f, j, w, z; có đ) */
const LETTERS_VN = "abcdđeghiklmnopqrstuvxy";

/** So hai chữ cái đánh điểm liền nhau: "next" = nối tiếp đúng (theo bảng chữ cái tiếng Anh hoặc tiếng Việt), "same", "jump", "restart" */
function letterStep(prev: string, cur: string): "next" | "same" | "jump" | "restart" {
  if (prev === cur) return "same";
  for (const abc of [LETTERS_EN, LETTERS_VN]) {
    const a = abc.indexOf(prev);
    const b = abc.indexOf(cur);
    if (a >= 0 && b === a + 1) return "next";
  }
  const pos = (c: string) => (c === "đ" ? LETTERS_EN.indexOf("d") + 0.5 : LETTERS_EN.indexOf(c));
  return pos(cur) > pos(prev) ? "jump" : "restart";
}

function ordinalOf(it: OutlineItem): number | null {
  const last = it.key.split(".").pop() || "";
  if (it.kind === "article" || it.kind === "dotted" || it.kind === "single") {
    const n = Number(last);
    return Number.isFinite(n) ? n : null;
  }
  if (it.kind === "roman") return romanValue(last) || null;
  if (it.kind === "chapter" || it.kind === "section") {
    const m = it.num.match(/([0-9]+|[IVXLCDM]+)$/i);
    if (!m) return null;
    return /^\d+$/.test(m[1]) ? Number(m[1]) : romanValue(m[1]) || null;
  }
  return null;
}

export function checkNumbering(rows: ParaRow[], outline: OutlineItem[]): Issue[] {
  const issues: Issue[] = [];
  const last = new Map<string, OutlineItem>();
  /** Giữa hai mục có đoạn nào đã bị xóa có theo dõi không (số nhảy do khoản bị xóa chờ duyệt) */
  const deletedBetween = (a: number, b: number) => rows.some((r) => r.i > a && r.i < b && r.deleted);
  const NUMERIC = ["article", "dotted", "single"];
  for (const it of outline) {
    if (it.kind === "heading" || !it.key) continue;
    const prefix = it.kind === "dotted" ? it.key.slice(0, it.key.lastIndexOf(".")) : "";
    const word = it.kind === "chapter" || it.kind === "section" ? it.num.replace(/\s*([0-9]+|[IVXLCDM]+)$/i, "").toLowerCase() : "";
    // Điều được đánh số liên tục qua các chương: so theo thứ tự trong văn bản, không theo mục cha
    const group = it.kind === "article" ? "article" : `${it.parent}|${it.kind}|${prefix}|${word}|${it.source === "list" ? "L" : "T"}`;
    const prev = last.get(group);
    last.set(group, it);
    const where = clauseName(it);
    if (it.kind === "letter") {
      const cur = it.key.split(".").pop() || "";
      if (!prev) {
        if (cur !== "a" && it.parent >= 0 && NUMERIC.includes(outline[it.parent].kind)) {
          issues.push({ kind: "numbering", p: it.i, clause: where, message: `Đánh số bắt đầu từ “${it.label}” (${where}), không có điểm a đứng trước.` });
        }
        continue;
      }
      const step = letterStep(prev.key.split(".").pop() || "", cur);
      if (step === "same") issues.push({ kind: "numbering", p: it.i, clause: where, message: `Đánh số trùng: ${where} xuất hiện hai lần.`, also: [prev.i] });
      else if (step === "jump") {
        issues.push({
          kind: "numbering",
          p: it.i,
          clause: where,
          message: `Đánh số nhảy: sau ${clauseName(prev)} là ${where}${deletedBetween(prev.i, it.i) ? " (ở giữa có đoạn đã bị xóa có theo dõi, chờ duyệt)" : ""}.`,
        });
      }
      continue;
    }
    const n = ordinalOf(it);
    if (n === null) continue;
    if (!prev) {
      // Khoản đầu tiên của một điều mà không mang số 1 ("5.2" đứng ngay sau "Điều 5")
      if (it.kind === "dotted" && n > 1 && it.parent >= 0 && outline[it.parent].key === prefix) {
        issues.push({ kind: "numbering", p: it.i, clause: where, message: `Đánh số bắt đầu từ ${where}: không có ${prefix}.1${n > 2 ? ` … ${prefix}.${n - 1}` : ""} đứng trước.` });
      }
      continue;
    }
    const pn = ordinalOf(prev);
    if (pn === null) continue;
    if (n === pn) issues.push({ kind: "numbering", p: it.i, clause: where, message: `Đánh số trùng: ${where} xuất hiện hai lần.`, also: [prev.i] });
    else if (n > pn + 1) {
      const missing = n - pn - 1;
      issues.push({
        kind: "numbering",
        p: it.i,
        clause: where,
        message: `Đánh số nhảy: sau ${clauseName(prev)} là ${where} (thiếu ${missing} số)${deletedBetween(prev.i, it.i) ? "; ở giữa có đoạn đã bị xóa có theo dõi, chờ duyệt" : ""}.`,
      });
    }
    // n < pn: danh sách đánh số lại từ đầu (phụ lục, danh sách mới) — không báo
  }
  return issues;
}

// ───────────────────────── 2. Dẫn chiếu tới điều không tồn tại ─────────────────────────

const LAW_WORDS =
  "bộ luật|luật|nghị định|thông tư|quyết định|nghị quyết|pháp lệnh|hiến pháp|quy chế|điều lệ|quy chuẩn|tiêu chuẩn|công ước|hiệp định|văn bản hợp nhất|chỉ thị|law|decree|circular|code|act";
const LAW_AFTER = new RegExp(`^\\s*(?:(?:của|tại|theo|thuộc|trong)\\s+)?(?:${LAW_WORDS})(?![\\p{L}\\p{N}])`, "iu");
const LAW_BEFORE = new RegExp(`(?:${LAW_WORDS})[^.;!?]{0,40}$`, "iu");
/** Văn bản khác (hợp đồng số…, thỏa thuận khung…): dẫn chiếu ra ngoài, không soát */
const OTHER_DOC_AFTER =
  /^\s*(?:(?:của|tại|theo|thuộc|trong)\s+)?(?:hợp đồng|thỏa thuận|thoả thuận|biên bản|quy định|nội quy|quy trình|văn bản|công văn|agreement|contract)\s+(?:số|khung|nguyên tắc|gốc|chính|mua bán|chuyển nhượng|thế chấp|tín dụng|ngày|no\.?|dated|[0-9])/iu;
/** Phần nối trong một chuỗi dẫn chiếu ("…, Điều 6 và Điều 7 của Luật…") — bỏ qua để nhìn thấy tên văn bản phía sau */
const REF_CHAIN = /^(?:\s*(?:,|;|và|hoặc|đến|tới|-|–)\s*(?:các\s+)?(?:(?:điều|khoản|điểm|article|clause)\s+)?[0-9a-zđ.()]+(?![\p{L}]))+/iu;
const FIELD_ERROR = /Error! Reference source not found\.?|Lỗi! Không tìm thấy nguồn tham chiếu\.?|Error! Bookmark not defined\.?|Lỗi! Thẻ đánh dấu không được xác định\.?/g;

const REF_ARTICLE = new RegExp(
  "(?:(điểm|tiết|point|item)\\s+\\(?([a-zđ]|[ivx]{1,5})\\)?\\s*,?\\s*)?" +
    "(?:(khoản|paragraph)\\s+(\\d{1,3})\\s*,?\\s*)?" +
    "(?<![\\p{L}\\p{N}])(điều|article|clause)\\s+(này|\\d{1,3}(?:\\.\\d{1,3})*)((?:\\s?\\([a-zđ]\\)|\\s?\\([ivx]{1,5}\\))*)(?![\\p{L}\\p{N}])",
  "giu"
);
const REF_DOTTED = /(?<![\p{L}\p{N}])(mục|khoản|điểm|tiểu mục)\s+(\d{1,3}(?:\.\d{1,3})+)(?![\p{L}\p{N}.]|\.\d)/giu;
const REF_BIG = /(?<![\p{L}\p{N}])(phụ lục|chương|mục)\s+(\d{1,3}|[IVXLCDM]{1,7})(?![\p{L}\p{N}]|\.\d)/giu;

export function checkReferences(rows: ParaRow[], outline: OutlineItem[]): Issue[] {
  const live: Live[] = rows.filter((r) => !r.deleted && r.text && r.text.trim()).map((r) => ({ i: r.i, text: clean(r.text), inTable: r.inTable }));
  const byPara = outlineByParagraph(outline);
  const numeric = outline.filter((o) => ["article", "dotted", "single", "letter", "roman"].includes(o.kind));
  const keys = new Set(numeric.map((o) => o.key));
  const textOf = new Map(live.map((r) => [r.i, r.text] as [number, string]));
  const found = new Map<string, Issue>();
  const add = (target: string, row: Live, at: number, len: number, message: string) => {
    const hit = found.get(target);
    if (hit) {
      if (hit.p !== row.i && !(hit.also || []).includes(row.i)) hit.also = [...(hit.also || []), row.i];
      hit.count = (hit.count || 1) + 1;
      return;
    }
    found.set(target, { kind: "reference", p: row.i, clause: clauseName(clauseAt(outline, row.i)), message, excerpt: around(row.text, at, len), count: 1 });
  };
  const hasArticles = outline.filter((o) => o.kind === "article" || (o.kind === "single" && o.level === 0) || o.kind === "dotted").length >= 2;
  const appendices = outline.filter((o) => o.kind === "chapter" && /^phụ lục/i.test(o.num));
  const chapters = outline.filter((o) => o.kind === "chapter" && /^chương/i.test(o.num));
  const sections = outline.filter((o) => o.kind === "section" && /^mục/i.test(o.num));
  const external = (text: string, start: number, end: number) => {
    const after = text.slice(end).replace(REF_CHAIN, "");
    return LAW_AFTER.test(after) || OTHER_DOC_AFTER.test(after) || LAW_BEFORE.test(text.slice(Math.max(0, start - 60), start));
  };
  /** Điểm (a), (ii) viết liền trong chữ của khoản cha, không tách thành đoạn riêng */
  const inlineMarker = (parentKey: string, mark: string) =>
    numeric
      .filter((o) => o.key === parentKey)
      .some((o) => {
        for (let p = o.i; p <= o.end; p++) if (new RegExp(`(?:\\(${esc(mark)}\\)|(?<![\\p{L}\\p{N}])${esc(mark)}\\))`, "u").test(textOf.get(p) || "")) return true;
        return false;
      });
  const exists = (key: string): boolean => {
    if (keys.has(key)) return true;
    const cutAt = key.lastIndexOf(".");
    if (cutAt < 0) return false;
    const lastPart = key.slice(cutAt + 1);
    return /^[a-zđ]+$/.test(lastPart) && keys.has(key.slice(0, cutAt)) && inlineMarker(key.slice(0, cutAt), lastPart);
  };
  /** Phần dài nhất của khóa có tồn tại ("7.6" → "7") để thông báo nói rõ thiếu ở cấp nào */
  const deepest = (key: string): string => {
    const parts = key.split(".");
    for (let n = parts.length - 1; n > 0; n--) if (keys.has(parts.slice(0, n).join("."))) return parts.slice(0, n).join(".");
    return "";
  };
  /** Điều (cấp số ngoài cùng) chứa đoạn p — để hiểu "khoản 3 Điều này" */
  const articleOf = (p: number): OutlineItem | null => {
    let hit: OutlineItem | null = null;
    for (let it = clauseAt(outline, p); it; it = it.parent >= 0 ? outline[it.parent] : null) if (["article", "dotted", "single"].includes(it.kind)) hit = it;
    return hit;
  };
  const scopedAppendix = (text: string, end: number): OutlineItem | null | undefined => {
    const m = text.slice(end).match(/^\s*(?:(?:của|tại|thuộc)\s+)?phụ lục\s+(\d{1,3}|[IVXLCDM]{1,7})(?![\p{L}\p{N}])/iu);
    if (!m) return undefined;
    return findClause(outline, `Phụ lục ${m[1]}`)[0] || null;
  };

  for (const row of live) {
    const head = byPara.get(row.i);
    for (const m of row.text.matchAll(FIELD_ERROR)) {
      add(`field@${row.i}`, row, m.index as number, m[0].length, `Trường dẫn chiếu hỏng (Word hiện “${m[0]}”): mục được dẫn chiếu đã bị xóa.`);
    }
    if (hasArticles) {
      for (const m of row.text.matchAll(REF_ARTICLE)) {
        const at = m.index as number;
        const end = at + m[0].length;
        // Chính dòng tiêu đề của điều ("Điều 7. Thời hạn") không phải dẫn chiếu
        if (at === 0 && head && head.source === "typed" && !m[1] && !m[3]) continue;
        let key: string;
        let shown = m[0].replace(/\s+/g, " ");
        if (m[6].toLowerCase() === "này") {
          if (!m[3]) continue; // "Điều này" đứng một mình
          const art = articleOf(row.i);
          if (!art) continue;
          key = [art.key, m[4], m[2] && /^(điểm|point|item)$/i.test(m[1]) ? m[2].toLowerCase() : ""].filter(Boolean).join(".");
          shown += ` (= ${clauseName(art)})`;
        } else {
          if (external(row.text, at, end)) continue;
          key = clauseKey(m[0]);
          // "Điều 5 Phụ lục 01": số nằm trong phụ lục
          const app = scopedAppendix(row.text, end);
          if (app === null) continue; // phụ lục không nằm trong văn bản này: không soát được
          if (app) {
            const inside = numeric.some((o) => o.key === key && o.i >= app.i && o.i <= app.end);
            if (!inside) add(`${key}@${app.num}`, row, at, m[0].length, `Dẫn chiếu “${shown}” của ${app.num}: trong ${app.num} không có mục mang số này.`);
            continue;
          }
        }
        if (!key || exists(key)) continue;
        const base = deepest(key);
        add(
          key,
          row,
          at,
          m[0].length,
          base
            ? `Dẫn chiếu “${shown}”: văn bản có ${clauseName(numeric.find((o) => o.key === base) as OutlineItem)} nhưng không có mục mang số ${key.slice(base.length + 1).replace(/\./g, " → ")} trong đó.`
            : `Dẫn chiếu “${shown}”: văn bản không có điều khoản mang số này.`
        );
      }
      for (const m of row.text.matchAll(REF_DOTTED)) {
        const at = m.index as number;
        const end = at + m[0].length;
        if (at === 0 && head) continue;
        if (external(row.text, at, end)) continue;
        // "khoản 5.2 Điều 5" đã được mẫu trên xét
        if (/^\s*,?\s*(?:điều|article|clause)\s/iu.test(row.text.slice(end))) continue;
        const key = m[2];
        const app = scopedAppendix(row.text, end);
        if (app === null) continue;
        const ok = app ? numeric.some((o) => o.key === key && o.i >= app.i && o.i <= app.end) : exists(key);
        if (!ok) add(`${key}${app ? "@" + app.num : ""}`, row, at, m[0].length, `Dẫn chiếu “${m[0].replace(/\s+/g, " ")}”${app ? ` của ${app.num}` : ""}: không có mục mang số này.`);
      }
    }
    for (const m of row.text.matchAll(REF_BIG)) {
      const at = m.index as number;
      const end = at + m[0].length;
      const word = m[1].toLowerCase();
      const pool = word === "phụ lục" ? appendices : word === "chương" ? chapters : sections;
      if (!pool.length) continue; // văn bản không có phụ lục / chương / mục nào nằm trong tệp: không soát được
      if (at === 0 && head) continue;
      if (word !== "phụ lục" && !/^\d+$/.test(m[2]) && m[2] !== m[2].toUpperCase()) continue;
      if (external(row.text, at, end)) continue;
      if (findClause(outline, `${m[1]} ${m[2]}`).length) continue;
      add(
        `${word} ${m[2]}`,
        row,
        at,
        m[0].length,
        `Dẫn chiếu “${m[0].replace(/\s+/g, " ")}”: trong văn bản chỉ có ${pool
          .slice(0, 6)
          .map((o) => o.num)
          .join(", ")}.`
      );
    }
  }
  return [...found.values()];
}

// ───────────────────────── 3. Thuật ngữ định nghĩa, 4. Tên các bên ─────────────────────────

const QO = "“\"«‘";
const QC = "”\"»’";
const DEF_QUOTED = new RegExp(
  `[${QO}]([^${QO}${QC}\\n]{2,60})[${QC}]\\s*(?:\\([^)]{0,40}\\)\\s*)?(?:là|được hiểu là|có nghĩa là|nghĩa là|được định nghĩa là|bao gồm|gồm|means|shall mean|has the meaning)(?![\\p{L}\\p{N}])`,
  "giu"
);
const DEF_CALLED = new RegExp(
  `(?:sau đây|dưới đây|hereinafter)[^${QO}.;]{0,40}?(?:gọi|called|referred to as)[^${QO}.;]{0,20}?[${QO}]([^${QO}${QC}\\n]{1,60})[${QC}](?:\\s*(?:hoặc|hay|or|,)\\s*[${QO}]([^${QO}${QC}\\n]{1,60})[${QC}])?`,
  "giu"
);
const DEF_PAREN = new RegExp(`\\(\\s*(?:gọi tắt là|gọi là|sau đây gọi là|the)?\\s*[${QO}]([^${QO}${QC}\\n]{1,60})[${QC}]\\s*\\)`, "giu");
const DEF_CALLED_BARE = /(?:sau đây|dưới đây)\s+(?:được\s+)?gọi\s+(?:tắt\s+|chung\s+)?là\s*:?\s*(Bên\s+(?:\p{Lu}(?![\p{L}])|\p{Lu}\p{Ll}+(?:\s+\p{Lu}\p{Ll}+){0,2}))/gu;
/** Dòng giới thiệu bên: "BÊN CHO THUÊ: …", "BÊN A (BÊN BÁN): …" */
const PARTY_INTRO = /^\s*(BÊN\s+[\p{Lu}\s]{1,30}?)\s*(?:\(\s*(BÊN\s+[\p{Lu}\s]{1,30}?)\s*\))?\s*[:：]/u;

const hasUpper = (s: string) => /\p{Lu}/u.test(s);
const isAllCaps = (s: string) => hasUpper(s) && s === s.toUpperCase() && s.replace(/[^\p{L}]/gu, "").length > 1;
/** Từ viết hoa chữ đầu (không phải viết hoa toàn bộ): "Ngày", "Bàn", "A" */
const isCapWord = (w: string) => /^\p{Lu}(?:[\p{Ll}\p{N}]*)$/u.test(w);
const titleCase = (s: string) =>
  s
    .toLowerCase()
    .split(/\s+/)
    .map((w) => (w ? w[0].toUpperCase() + w.slice(1) : w))
    .join(" ");
const termRegex = (t: string) => new RegExp(`(?<![\\p{L}\\p{N}])${t.split(/\s+/).map(esc).join("\\s+")}(?![\\p{L}\\p{N}])`, "giu");
/** Các cụm sau "Bên" không phải tên một bên của hợp đồng */
const NOT_PARTY = /^Bên (Thứ Ba|Kia|Vi Phạm|Bị Vi Phạm|Liên Quan|Còn Lại|Nhận|Giao|Dưới|Trên|Trong|Ngoài|Cạnh)$/;

interface Defs {
  /** thuật ngữ → đoạn định nghĩa */
  terms: Map<string, number>;
  parties: Map<string, number>;
  /** Một đoạn định nghĩa hai tên cho cùng một bên: [tên 1, tên 2, đoạn] */
  aliases: [string, string, number][];
}

export function findDefinitions(live: Live[]): Defs {
  const terms = new Map<string, number>();
  const parties = new Map<string, number>();
  const aliases: [string, string, number][] = [];
  const norm = (s: string) => s.replace(/\s+/g, " ").replace(/^[\s.,;:]+|[\s.,;:]+$/g, "");
  const put = (raw: string | undefined, i: number): string => {
    const t = norm(raw || "");
    if (!t || t.length > 60 || !hasUpper(t) || /^\d/.test(t)) return "";
    if (!terms.has(t)) terms.set(t, i);
    if (/^Bên\s/.test(t) && !parties.has(t)) parties.set(t, i);
    return t;
  };
  for (const r of live) {
    for (const m of r.text.matchAll(DEF_QUOTED)) put(m[1], r.i);
    for (const m of r.text.matchAll(DEF_CALLED)) {
      const first = put(m[1], r.i);
      const second = m[2] ? put(m[2], r.i) : "";
      // "sau đây gọi là “Bên A” hoặc “Bên Bán”": một bên, hai tên
      if (/^Bên\s/.test(first) && /^Bên\s/.test(second) && first !== second) aliases.push([first, second, r.i]);
    }
    for (const m of r.text.matchAll(DEF_PAREN)) put(m[1], r.i);
    for (const m of r.text.matchAll(DEF_CALLED_BARE)) put(m[1], r.i);
    const intro = r.text.match(PARTY_INTRO);
    if (intro) {
      const names: string[] = [];
      for (const g of [intro[1], intro[2]]) {
        if (!g) continue;
        const name = titleCase(norm(g));
        if (!parties.has(name)) parties.set(name, r.i);
        names.push(name);
      }
      // "BÊN A (BÊN BÁN): …": một bên, hai tên
      if (names.length === 2 && names[0] !== names[1]) aliases.push([names[0], names[1], r.i]);
    }
  }
  return { terms, parties, aliases };
}

export function checkTerms(rows: ParaRow[], outline: OutlineItem[], which: { term?: boolean; party?: boolean } = { term: true, party: true }): { issues: Issue[]; terms: string[]; parties: string[] } {
  const live: Live[] = rows.filter((r) => !r.deleted && r.text && r.text.trim()).map((r) => ({ i: r.i, text: clean(r.text), inTable: r.inTable }));
  const defs = findDefinitions(live);
  const byPara = outlineByParagraph(outline);
  const issues: Issue[] = [];
  const clauseOf = (p: number) => clauseName(clauseAt(outline, p));
  /** Dòng tiêu đề (tên điều, tên chương, dòng viết hoa toàn bộ): cách viết hoa ở đó theo thể thức tiêu đề, không soát */
  const isTitle = (r: Live) => isAllCaps(r.text) || (byPara.has(r.i) && r.text.length <= 100 && !/[.;:!?]$/.test(r.text.trim()));
  const termList = [...defs.terms.keys()].sort((a, b) => b.length - a.length);
  const partyNames = [...defs.parties.keys()].sort((a, b) => b.length - a.length);
  const heads = new Set(termList.filter((t) => t.includes(" ")).map((t) => t.split(/\s+/)[0]).filter(isCapWord));

  // Dùng lệch cách viết hoa: “Ngày Bàn Giao” được định nghĩa nhưng có chỗ viết “ngày bàn giao”
  const variants = new Map<string, { term: string; form: string; ps: number[]; at: number; text: string }>();
  // Cụm viết hoa như thuật ngữ nhưng chưa định nghĩa, và tên bên chưa định nghĩa
  const undefinedTerms = new Map<string, { ps: number[]; at: number; text: string }>();
  const strayParties = new Map<string, { ps: number[]; at: number; text: string }>();
  const usedParty = new Map<string, number>();

  for (const r of live) {
    if (isTitle(r)) continue;
    let masked = r.text;
    for (const t of termList) {
      const multi = t.includes(" ");
      masked = masked.replace(termRegex(t), (m, off: number) => {
        const form = m.replace(/\s+/g, " ");
        if (form === t) {
          if (defs.parties.has(t) && defs.terms.get(t) !== r.i) usedParty.set(t, (usedParty.get(t) || 0) + 1);
        } else if (which.term && multi && !isAllCaps(form) && defs.terms.get(t) !== r.i) {
          // Một chữ cái đầu câu viết hoa khác đi thì không tính là lệch
          const k = `${t}→${form}`;
          const v = variants.get(k);
          if (v) v.ps.push(r.i);
          else variants.set(k, { term: t, form, ps: [r.i], at: off, text: r.text });
        }
        return " ".repeat(m.length);
      });
    }
    // Các cụm viết hoa còn lại
    const words = [...masked.matchAll(/[\p{L}\p{N}]+/gu)].map((m) => ({ w: m[0], at: m.index as number }));
    for (let k = 0; k < words.length; k++) {
      const w = words[k];
      const next = words[k + 1];
      const adjacent = (a: { w: string; at: number }, b?: { w: string; at: number }) => !!b && /^ +$/.test(masked.slice(a.at + a.w.length, b.at));
      const isParty = /^[Bb]ên$/.test(w.w) && w.w === "Bên";
      if (!(isParty || heads.has(w.w)) || !adjacent(w, next) || !isCapWord(next.w)) continue;
      let e = k + 1;
      while (e + 1 < words.length && e - k < 4 && adjacent(words[e], words[e + 1]) && isCapWord(words[e + 1].w)) e++;
      const phrase = words
        .slice(k, e + 1)
        .map((x) => x.w)
        .join(" ");
      const store = isParty ? strayParties : undefinedTerms;
      if (isParty ? !which.party || !defs.parties.size || NOT_PARTY.test(phrase) : !which.term) {
        k = e;
        continue;
      }
      // Tên bên đã biết qua dòng giới thiệu viết hoa ("BÊN A:")
      if (isParty && partyNames.some((p) => phrase === p || phrase.startsWith(p + " "))) {
        const hit = partyNames.find((p) => phrase === p || phrase.startsWith(p + " ")) as string;
        usedParty.set(hit, (usedParty.get(hit) || 0) + 1);
        k = e;
        continue;
      }
      const hit = store.get(phrase);
      if (hit) hit.ps.push(r.i);
      else store.set(phrase, { ps: [r.i], at: w.at, text: r.text });
      k = e;
    }
  }

  if (which.term) {
    for (const v of variants.values()) {
      const ps = Array.from(new Set(v.ps));
      issues.push({
        kind: "term",
        p: ps[0],
        clause: clauseOf(ps[0]),
        message: `Thuật ngữ “${v.term}” (định nghĩa ở [¶${defs.terms.get(v.term)}]) có chỗ viết “${v.form}”: ${v.ps.length} chỗ.`,
        excerpt: around(v.text, v.at, v.form.length),
        ...(ps.length > 1 ? { also: ps.slice(1, 12) } : {}),
        count: v.ps.length,
      });
    }
    for (const [phrase, u] of undefinedTerms) {
      const head = phrase.split(" ")[0];
      const near = termList.filter((t) => t.split(/\s+/)[0] === head).slice(0, 4);
      const ps = Array.from(new Set(u.ps));
      issues.push({
        kind: "term",
        p: ps[0],
        clause: clauseOf(ps[0]),
        message: `“${phrase}” viết hoa như một thuật ngữ nhưng không được định nghĩa${near.length ? ` (thuật ngữ đã định nghĩa gần giống: ${near.map((t) => `“${t}”`).join(", ")})` : ""}: ${u.ps.length} chỗ.`,
        excerpt: around(u.text, u.at, phrase.length),
        ...(ps.length > 1 ? { also: ps.slice(1, 12) } : {}),
        count: u.ps.length,
      });
    }
  }
  if (which.party) {
    const known = [...defs.parties.keys()];
    for (const [phrase, u] of strayParties) {
      const ps = Array.from(new Set(u.ps));
      issues.push({
        kind: "party",
        p: ps[0],
        clause: clauseOf(ps[0]),
        message: `“${phrase}” được dùng ${u.ps.length} chỗ nhưng không nằm trong các tên bên đã định nghĩa (${known.join(", ")}): có thể một bên đang được gọi bằng nhiều tên.`,
        excerpt: around(u.text, u.at, phrase.length),
        ...(ps.length > 1 ? { also: ps.slice(1, 12) } : {}),
        count: u.ps.length,
      });
    }
    for (const [a, b, p] of defs.aliases) {
      if ((usedParty.get(a) || 0) > 0 && (usedParty.get(b) || 0) > 0) {
        issues.push({
          kind: "party",
          p,
          clause: clauseOf(p),
          message: `Một bên được định nghĩa bằng hai tên “${a}” và “${b}”, và cả hai tên đều đang được dùng (${usedParty.get(a)} và ${usedParty.get(b)} chỗ): nên dùng thống nhất một tên.`,
        });
      }
    }
  }
  return { issues, terms: termList.slice().sort(), parties: [...defs.parties.keys()] };
}

// ───────────────────────── 5. Chỗ trống chưa điền ─────────────────────────

const BLANKS: RegExp[] = [
  /(?<![\p{L}\p{N}.…])(?:\.{3,}|…+)(?:\s*(?:\.{2,}|…+))*/gu,
  /_{3,}/g,
  /\[\s*(?:●|•|\*+|\.{2,}|…+|_+|x{2,})?\s*\]/giu,
  /\[\s*(?:điền|ghi|nhập|chèn|bổ sung|insert|tbd|to be)[^\]]{0,60}\]/giu,
  /<\s*(?:điền|ghi|nhập|chèn|…|\.{3,})[^>]{0,60}>/giu,
  /(?<![\p{L}\p{N}])(?:TBD|TBC|XXX+|xxx+)(?![\p{L}\p{N}])/gu,
  /ngày\s{2,}tháng\s{2,}năm/giu,
];
/** Dòng ký tên: gạch chân, chấm chấm ở đó là chỗ ký, không phải chỗ quên điền */
const SIGN_LINE = /ký tên|chữ ký|ký, ghi rõ|ký và ghi rõ|ký, đóng dấu|đóng dấu|signature|signed by/i;

export function checkBlanks(rows: ParaRow[], outline: OutlineItem[]): Issue[] {
  const issues: Issue[] = [];
  for (const r of rows) {
    if (r.deleted || !r.text) continue;
    const text = clean(r.text);
    if (SIGN_LINE.test(text)) continue;
    const hits: { at: number; len: number }[] = [];
    for (const re of BLANKS) for (const m of text.matchAll(re)) hits.push({ at: m.index as number, len: m[0].length });
    if (!hits.length) continue;
    hits.sort((a, b) => a.at - b.at);
    // Gộp các chỗ chồng nhau
    const merged: { at: number; len: number }[] = [];
    for (const h of hits) {
      const lastHit = merged[merged.length - 1];
      if (lastHit && h.at < lastHit.at + lastHit.len) lastHit.len = Math.max(lastHit.len, h.at + h.len - lastHit.at);
      else merged.push({ ...h });
    }
    const first = merged[0];
    const s = Math.max(0, first.at - 40);
    const e = Math.min(text.length, merged[merged.length - 1].at + merged[merged.length - 1].len + 30);
    issues.push({
      kind: "blank",
      p: r.i,
      clause: clauseName(clauseAt(outline, r.i)),
      message: `Chỗ trống chưa điền${merged.length > 1 ? ` (${merged.length} chỗ trong đoạn)` : ""}.`,
      excerpt: cut((s > 0 ? "…" : "") + text.slice(s, e) + (e < text.length ? "…" : ""), 220),
      count: merged.length,
    });
  }
  return issues;
}

// ───────────────────────── 6. Số bằng số không khớp bằng chữ ─────────────────────────

const DIGIT: Record<string, number> = {
  không: 0,
  một: 1,
  mốt: 1,
  hai: 2,
  ba: 3,
  bốn: 4,
  tư: 4,
  năm: 5,
  lăm: 5,
  nhăm: 5,
  sáu: 6,
  bảy: 7,
  bẩy: 7,
  tám: 8,
  chín: 9,
};
const SCALE: Record<string, number> = { nghìn: 1e3, ngàn: 1e3, triệu: 1e6, tỷ: 1e9, tỉ: 1e9 };
const isNumberWord = (w: string) => w in DIGIT || w in SCALE || ["mười", "mươi", "trăm", "linh", "lẻ", "rưỡi", "chẵn"].includes(w);

/** Đọc số viết bằng chữ tiếng Việt ở đầu dãy từ. Trả giá trị và số từ đã dùng; null nếu từ đầu tiên không phải chữ số.
 *  "một tỷ năm trăm triệu" → 1.500.000.000 · "hai mươi lăm" → 25 · "một trăm linh năm" → 105 · "mười lăm" → 15 */
export function parseVnNumber(words: string[]): { value: number; used: number } | null {
  let total = 0;
  let group = 0;
  let d: number | null = null;
  let lastUnit = 0;
  let used = 0;
  let any = false;
  for (const raw of words) {
    const w = raw.toLowerCase();
    if (w in DIGIT) {
      if (d !== null) break; // hai chữ số liền nhau ("ba năm" = 3 năm): dừng
      d = DIGIT[w];
      any = true;
    } else if (w === "mười") {
      if (d !== null) break;
      group += 10;
      any = true;
    } else if (w === "mươi") {
      if (d === null) break;
      group += d * 10;
      d = null;
    } else if (w === "trăm") {
      if (d === null) break;
      group += d * 100;
      d = null;
      lastUnit = 100;
    } else if (w === "linh" || w === "lẻ") {
      if (d !== null || !any) break;
    } else if (w in SCALE) {
      if (!any) break;
      group += d ?? 0;
      d = null;
      const s = SCALE[w];
      if (s === 1e9) total = (total + (group || (total ? 0 : 1))) * 1e9;
      else total += (group || 1) * s;
      group = 0;
      lastUnit = s;
    } else if (w === "rưỡi") {
      if (!lastUnit) break;
      if (lastUnit === 100) group += 50;
      else total += lastUnit / 2;
    } else if (w === "chẵn") {
      used++;
      break;
    } else break;
    used++;
  }
  if (!any) return null;
  return { value: total + group + (d ?? 0), used };
}

/** Giá trị của một chuỗi số viết bằng chữ số: "1.500.000" → 1500000; "450,000" → 450000; "1,5" → 1.5; không rõ → null */
export function parseFigure(s: string): number | null {
  if (/^\d+$/.test(s)) return Number(s);
  if (/^\d{1,3}(?:\.\d{3})+$/.test(s) || /^\d{1,3}(?:,\d{3})+$/.test(s)) return Number(s.replace(/[.,]/g, ""));
  if (/^\d+[.,]\d{1,2}$/.test(s)) return Number(s.replace(",", "."));
  return null;
}

const fmt = (n: number) => (Number.isInteger(n) ? n.toLocaleString("vi-VN") : String(n).replace(".", ","));
const FIGURE_THEN_WORDS = /(?<![\p{L}\p{N}.,])(\d[\d.,]*\d|\d)\s*(nghìn|ngàn|triệu|tỷ|tỉ)?\s*((?:[^\s()\d,.;:]{1,12}\s*){0,3}?)[,;:]?\s*\(\s*(?:(?:viết\s+)?bằng chữ|in words)?\s*:?\s*([^()]{2,200})\)/giu;
const WORDS_THEN_FIGURE = /\(\s*(\d[\d.,]*\d|\d)\s*\)/g;

export function checkAmounts(rows: ParaRow[], outline: OutlineItem[]): Issue[] {
  const issues: Issue[] = [];
  const seen = new Set<string>();
  const push = (r: ParaRow, text: string, at: number, len: number, figure: number, figText: string, wordsText: string, value: number) => {
    const k = `${r.i}:${at}`;
    if (seen.has(k)) return;
    seen.add(k);
    issues.push({
      kind: "amount",
      p: r.i,
      clause: clauseName(clauseAt(outline, r.i)),
      message: `Số bằng số và bằng chữ không khớp: ${figText} (= ${fmt(figure)}) nhưng chữ ghi “${wordsText.trim()}” (= ${fmt(value)}).`,
      excerpt: around(text, at, len, 30),
    });
  };
  for (const r of rows) {
    if (r.deleted || !r.text) continue;
    const text = clean(r.text);
    for (const m of text.matchAll(FIGURE_THEN_WORDS)) {
      let figure = parseFigure(m[1]);
      if (figure === null) continue;
      if (m[2]) figure *= SCALE[m[2].toLowerCase()];
      const toks = m[4].trim().split(/\s+/);
      const all = toks.map((t) => t.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, ""));
      if (all.some((t) => /^phẩy$/i.test(t))) continue; // số thập phân viết bằng chữ: không soát
      const parsed = parseVnNumber(all);
      if (!parsed) continue;
      // Phần còn lại sau con số phải ngắn (đơn vị: "đồng", "ngày làm việc", "phần trăm"), nếu không thì đây không phải số viết bằng chữ
      if (all.length - parsed.used > 4) continue;
      const values = [parsed.value];
      // "năm" đứng cuối có thể là đơn vị thời gian: "(hai mươi năm)" = 20 năm hoặc 25
      if (all[parsed.used - 1]?.toLowerCase() === "năm" && parsed.used > 1) {
        const alt = parseVnNumber(all.slice(0, parsed.used - 1));
        if (alt && alt.used === parsed.used - 1) values.push(alt.value);
      }
      if (values.some((v) => Math.abs(v - (figure as number)) < 1e-9)) continue;
      push(r, text, m.index as number, m[0].length, figure, `${m[1]}${m[2] ? " " + m[2] : ""}`, m[4], parsed.value);
    }
    for (const m of text.matchAll(WORDS_THEN_FIGURE)) {
      const figure = parseFigure(m[1]);
      if (figure === null) continue;
      const before = text
        .slice(0, m.index as number)
        .trim()
        .split(/\s+/)
        .slice(-14)
        .map((t) => t.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, ""));
      let start = before.length;
      while (start > 0 && isNumberWord(before[start - 1].toLowerCase())) start--;
      const numWords = before.slice(start);
      if (!numWords.length) continue;
      // Chỉ một chữ "năm" đứng trước "(2026)" là "năm 2026", không phải số 5
      if (numWords.length === 1 && numWords[0].toLowerCase() === "năm") continue;
      const parsed = parseVnNumber(numWords);
      if (!parsed || parsed.used !== numWords.length) continue;
      if (Math.abs(parsed.value - figure) < 1e-9) continue;
      const at = text.lastIndexOf(numWords[0], m.index as number);
      push(r, text, Math.max(0, at), (m.index as number) + m[0].length - Math.max(0, at), figure, m[1], numWords.join(" "), parsed.value);
    }
  }
  return issues;
}

// ───────────────────────── Gộp ─────────────────────────

export function checkConsistency(rows: ParaRow[], opts: { checks?: IssueKind[]; outline?: OutlineItem[] } = {}): ConsistencyResult {
  const outline = opts.outline || buildOutline(rows);
  const want = new Set(opts.checks && opts.checks.length ? opts.checks : ALL_CHECKS);
  const issues: Issue[] = [];
  if (want.has("numbering")) issues.push(...checkNumbering(rows, outline));
  if (want.has("reference")) issues.push(...checkReferences(rows, outline));
  let terms: string[] = [];
  let parties: string[] = [];
  if (want.has("term") || want.has("party")) {
    const t = checkTerms(rows, outline, { term: want.has("term"), party: want.has("party") });
    issues.push(...t.issues);
    terms = t.terms;
    parties = t.parties;
  }
  if (want.has("blank")) issues.push(...checkBlanks(rows, outline));
  if (want.has("amount")) issues.push(...checkAmounts(rows, outline));
  const order = (k: IssueKind) => ALL_CHECKS.indexOf(k);
  issues.sort((a, b) => order(a.kind) - order(b.kind) || a.p - b.p);
  return { issues, terms, parties };
}
