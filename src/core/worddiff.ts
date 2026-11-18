/** So khác ở mức từ và lập kế hoạch sửa một đoạn Word sao cho mỗi chỗ đổi là một cặp xóa/thêm nhỏ.
 *  Hàm thuần, không cần Office — kiểm thử bằng npm test (test/unit/worddiff.test.ts).
 *
 *  Kết quả thăm dò 15/11/2026 quyết định cách làm:
 *  - `paragraph.search` nhìn chuỗi GỒM CẢ chữ đã bị xóa có theo dõi (chuỗi "đánh dấu" = paragraph.text), nên mọi phép định vị
 *    tính trên chuỗi đánh dấu M; còn so khác thì tính trên bản hiện tại C (= M bỏ các khúc đã xóa).
 *  - Nhiều chỗ sửa trong một đoạn, làm trong một lần sync, Word ghi đúng từng cặp xóa/thêm và giữ định dạng. */

export interface Hunk {
  /** Vị trí trong chuỗi cũ */
  at: number;
  del: string;
  ins: string;
}

const TOKEN = /\s+|[\p{L}\p{N}]+(?:[.,/][\p{N}]+)*%?|[^\s\p{L}\p{N}]/gu;

/** Trích chữ GỐC của một đoạn đang có sửa đổi, quanh chỗ đầu tiên khác với chữ hiện tại (để chỗ bị sửa luôn nằm trong phần trích).
 *  Đoạn ngắn hơn max thì trả nguyên; phần bị cắt đánh dấu bằng "…". */
export function beforeExcerpt(original: string, current: string, max = 400): string {
  const o = String(original ?? "");
  if (o.length <= max) return o;
  const c = String(current ?? "");
  let d = 0;
  const n = Math.min(o.length, c.length);
  while (d < n && o[d] === c[d]) d++;
  // Giống hệt nhau tới hết (chỉ khác ở đuôi) hoặc không có chữ hiện tại: lấy quanh điểm đó
  const start = Math.max(0, Math.min(d - Math.floor(max * 0.35), o.length - max));
  return (start > 0 ? "…" : "") + o.slice(start, start + max) + (start + max < o.length ? "…" : "");
}

export function tokenize(s: string): string[] {
  return s.match(TOKEN) || [];
}

/** Khác biệt theo từ giữa câu cũ và câu mới. Các chỗ đổi chỉ cách nhau bởi khoảng trắng được gộp làm một. */
export function diffWords(oldText: string, newText: string): Hunk[] {
  if (oldText === newText) return [];
  const a = tokenize(oldText);
  const b = tokenize(newText);
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let s = 0;
  while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
  const am = a.slice(p, a.length - s);
  const bm = b.slice(p, b.length - s);
  const n = am.length;
  const m = bm.length;
  // Đoạn giữa quá lớn: coi là một khối
  type Op = { t: "eq" | "del" | "ins"; v: string };
  const ops: Op[] = [];
  if (n * m > 4_000_000) {
    am.forEach((v) => ops.push({ t: "del", v }));
    bm.forEach((v) => ops.push({ t: "ins", v }));
  } else {
    const w = m + 1;
    const dp = new Uint16Array((n + 1) * w);
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        dp[i * w + j] = am[i] === bm[j] ? dp[(i + 1) * w + j + 1] + 1 : Math.max(dp[(i + 1) * w + j], dp[i * w + j + 1]);
      }
    }
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (am[i] === bm[j]) {
        ops.push({ t: "eq", v: am[i] });
        i++;
        j++;
      } else if (dp[(i + 1) * w + j] >= dp[i * w + j + 1]) ops.push({ t: "del", v: am[i++] });
      else ops.push({ t: "ins", v: bm[j++] });
    }
    while (i < n) ops.push({ t: "del", v: am[i++] });
    while (j < m) ops.push({ t: "ins", v: bm[j++] });
  }
  // Gom thành khối; khoảng trắng không đổi nằm giữa hai chỗ đổi thì nhập vào khối
  const hunks: Hunk[] = [];
  let at = a.slice(0, p).join("").length;
  let cur: Hunk | null = null;
  for (let k = 0; k < ops.length; k++) {
    const o = ops[k];
    if (o.t === "eq") {
      const next = ops[k + 1];
      if (cur && /^\s+$/.test(o.v) && next && next.t !== "eq") {
        cur.del += o.v;
        cur.ins += o.v;
      } else {
        if (cur) hunks.push(cur);
        cur = null;
      }
      at += o.v.length;
    } else {
      if (!cur) cur = { at, del: "", ins: "" };
      if (o.t === "del") {
        cur.del += o.v;
        at += o.v.length;
      } else cur.ins += o.v;
    }
  }
  if (cur) hunks.push(cur);
  return hunks.filter((h) => h.del !== h.ins);
}

/** Áp các khối vào chuỗi cũ (để kiểm thử: phải ra đúng chuỗi mới) */
export function applyHunks(oldText: string, hunks: Hunk[]): string {
  let out = "";
  let pos = 0;
  for (const h of hunks) {
    out += oldText.slice(pos, h.at) + h.ins;
    pos = h.at + h.del.length;
  }
  return out + oldText.slice(pos);
}

/** Tỷ lệ chữ bị đổi (0–1) */
export function changeRatio(oldText: string, newText: string, hunks: Hunk[]): number {
  const total = oldText.length + newText.length;
  if (!total) return 0;
  return hunks.reduce((n, h) => n + h.del.length + h.ins.length, 0) / total;
}

/** Vị trí trong chuỗi đánh dấu M của từng ký tự thuộc bản hiện tại C. spans = các khúc [đầu, cuối) của M đã bị xóa có theo dõi
 *  (lấy từ vị trí thật trong Word khi có); không có hoặc không khớp thì dò C như một dãy con của M. Trả null nếu C không nằm trong M. */
export function liveMap(M: string, C: string, spans?: [number, number][]): number[] | null {
  if (M === C) return Array.from({ length: C.length }, (_, k) => k);
  if (spans && spans.length) {
    const dead = new Uint8Array(M.length);
    for (const [s, e] of spans) for (let k = Math.max(0, s); k < Math.min(M.length, e); k++) dead[k] = 1;
    const map: number[] = [];
    let ok = true;
    for (let k = 0; k < M.length && ok; k++) {
      if (dead[k]) continue;
      if (M[k] !== C[map.length]) ok = false;
      else map.push(k);
    }
    if (ok && map.length === C.length) return map;
  }
  // Không có vị trí thật: dò C như một dãy con của M, và CHỈ nhận khi cách gán là duy nhất (dò từ trái và dò từ phải ra cùng kết quả).
  // Gán nhập nhằng có thể trỏ vào chữ đã xóa ("là 12 tháng18 tháng" với bản hiện tại "là 18 tháng") và làm sửa nhầm chữ.
  const map: number[] = [];
  let i = 0;
  for (let j = 0; j < C.length; j++) {
    while (i < M.length && M[i] !== C[j]) i++;
    if (i >= M.length) return null;
    map.push(i++);
  }
  let r = M.length - 1;
  for (let j = C.length - 1; j >= 0; j--) {
    while (r >= 0 && M[r] !== C[j]) r--;
    if (r !== map[j]) return null;
    r--;
  }
  return map;
}

/** Số lần `s` xuất hiện trong M (quét trái sang phải, không chồng nhau, như Word tìm) và thứ tự của lần bắt đầu đúng tại `at` (-1 nếu không rơi đúng) */
export function occurrenceAt(M: string, s: string, at: number): { index: number; total: number } {
  let pos = 0;
  let total = 0;
  let index = -1;
  if (!s) return { index: -1, total: 0 };
  for (;;) {
    const k = M.indexOf(s, pos);
    if (k < 0) break;
    if (k === at) index = total;
    total++;
    pos = k + s.length;
  }
  return { index, total };
}

export interface Locator {
  /** Chuỗi tìm trong đoạn (phân biệt hoa thường) */
  search: string;
  /** Lần xuất hiện thứ mấy (0-based) trong các kết quả tìm của đoạn */
  occurrence: number;
  /** Tổng số lần phải thấy; lệch thì không sửa */
  total: number;
  /** Vị trí trong chuỗi đánh dấu (để đối chiếu bằng vị trí thật khi số lần tìm lệch) */
  at: number;
}

export type EditOp =
  /** Thay (hoặc xóa khi text rỗng) một khúc: từ `from` tới hết `to` (không có `to` = chính `from`) */
  | { kind: "replace"; from: Locator; to?: Locator; text: string }
  | { kind: "insertAfter"; from: Locator; text: string }
  | { kind: "insertBefore"; from: Locator; text: string }
  /** Đoạn đang rỗng: chèn vào đầu đoạn */
  | { kind: "insertStart"; text: string };

export interface EditPlan {
  ops: EditOp[];
  /** Số chỗ đổi (cặp xóa/thêm) */
  hunks: number;
  ratio: number;
}

const MAX_SEARCH = 255;
const LONG_PIECE = 120;

function locator(M: string, start: number, end: number): Locator | null {
  const s = M.slice(start, end);
  // Word coi ^ là ký tự đặc biệt nên chuỗi gửi đi có ^ nhân đôi: giới hạn 255 tính trên chuỗi đã nhân đôi
  if (!s || s.length + (s.match(/\^/g) || []).length > MAX_SEARCH) return null;
  const o = occurrenceAt(M, s, start);
  if (o.index < 0) return null;
  return { search: s, occurrence: o.index, total: o.total, at: start };
}

/** Khúc liền nhau trong M của các ký tự C[a..b) */
function liveRuns(map: number[], a: number, b: number): [number, number][] {
  const runs: [number, number][] = [];
  for (let j = a; j < b; j++) {
    const m = map[j];
    const lastRun = runs[runs.length - 1];
    if (lastRun && lastRun[1] === m) lastRun[1] = m + 1;
    else runs.push([m, m + 1]);
  }
  return runs;
}

/** Lập kế hoạch sửa đoạn: M = chuỗi đánh dấu (paragraph.text), C = bản hiện tại, N = câu mới.
 *  Trả lỗi (chuỗi) khi không định vị chắc chắn được — khi đó KHÔNG được thay cả đoạn nếu đoạn đang có sửa đổi chờ duyệt. */
export function planParagraphEdit(M: string, C: string, N: string, spans?: [number, number][]): EditPlan | string {
  const map = liveMap(M, C, spans);
  if (!map) return "Không xác định chắc được chữ nào trong đoạn đang sống, chữ nào đã bị xóa có theo dõi.";
  const hunks = diffWords(C, N);
  const ops: EditOp[] = [];
  for (const h of hunks) {
    if (h.del) {
      const runs = liveRuns(map, h.at, h.at + h.del.length);
      for (let r = 0; r < runs.length; r++) {
        const [s, e] = runs[r];
        const text = r === 0 ? h.ins : "";
        if (e - s <= MAX_SEARCH) {
          const from = locator(M, s, e);
          if (!from) return `Không định vị chắc chắn được chỗ cần sửa “${M.slice(s, Math.min(e, s + 40))}”.`;
          ops.push({ kind: "replace", from, text });
        } else {
          const from = locator(M, s, s + LONG_PIECE);
          const to = locator(M, e - LONG_PIECE, e);
          if (!from || !to) return `Không định vị chắc chắn được khúc dài cần sửa “${M.slice(s, s + 40)}…”.`;
          ops.push({ kind: "replace", from, to, text });
        }
      }
    } else if (!C.length) {
      ops.push({ kind: "insertStart", text: h.ins });
    } else if (h.at > 0) {
      // Neo vào từ đứng ngay trước (tối đa 40 ký tự, liền nhau trong M)
      const endM = map[h.at - 1] + 1;
      let j = h.at - 1;
      while (j > 0 && h.at - j < 40 && map[j - 1] === map[j] - 1 && !/\s/.test(C[j - 1])) j--;
      let from = locator(M, map[j], endM);
      // Từ neo trùng lặp kiểu chồng nhau: nới rộng dần về trước
      while (!from && j > 0 && map[j - 1] === map[j] - 1 && h.at - j < 200) from = locator(M, map[--j], endM);
      if (!from) return `Không định vị chắc chắn được chỗ chèn sau “${C.slice(Math.max(0, h.at - 30), h.at)}”.`;
      ops.push({ kind: "insertAfter", from, text: h.ins });
    } else {
      // Chèn vào đầu đoạn: neo vào từ đầu tiên
      const startM = map[0];
      let j = 0;
      while (j + 1 < C.length && j < 40 && map[j + 1] === map[j] + 1 && !/\s/.test(C[j + 1])) j++;
      const from = locator(M, startM, map[j] + 1);
      if (!from) return "Không định vị chắc chắn được đầu đoạn để chèn.";
      ops.push({ kind: "insertBefore", from, text: h.ins });
    }
  }
  return { ops, hunks: hunks.length, ratio: changeRatio(C, N, hunks) };
}

// ───────────────────────── Hoàn tác: sửa đổi nào là của lượt vừa rồi ─────────────────────────

/** Khóa nhận diện một sửa đổi được theo dõi. Ngày trong Word chỉ chính xác tới phút nên không dùng mốc thời gian để lọc. */
export function changeKey(t: { type: string; author: string; date: any; text: string }): string {
  const text = String(t.text ?? "");
  const d = t.date ? new Date(t.date as any).getTime() : 0;
  return `${t.type}|${t.author}|${d}|${text.length}|${text.slice(0, 200)}`;
}

/** Các khóa có trong `after` mà không có trong `before`, so theo thứ tự trong văn bản (dãy con chung dài nhất). */
export function addedKeys(before: string[], after: string[]): string[] {
  let p = 0;
  while (p < before.length && p < after.length && before[p] === after[p]) p++;
  let s = 0;
  while (s < before.length - p && s < after.length - p && before[before.length - 1 - s] === after[after.length - 1 - s]) s++;
  const a = before.slice(p, before.length - s);
  const b = after.slice(p, after.length - s);
  if (!b.length) return [];
  if (!a.length) return b;
  if (a.length * b.length > 4_000_000) {
    // Quá lớn: so theo số lượng
    const count = new Map<string, number>();
    a.forEach((k) => count.set(k, (count.get(k) || 0) + 1));
    return b.filter((k) => {
      const c = count.get(k) || 0;
      if (c > 0) {
        count.set(k, c - 1);
        return false;
      }
      return true;
    });
  }
  const w = b.length + 1;
  const dp = new Uint16Array((a.length + 1) * w);
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      dp[i * w + j] = a[i] === b[j] ? dp[(i + 1) * w + j + 1] + 1 : Math.max(dp[(i + 1) * w + j], dp[i * w + j + 1]);
    }
  }
  const out: string[] = [];
  let i = 0;
  let j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) {
      i++;
      j++;
    } else if (dp[(i + 1) * w + j] >= dp[i * w + j + 1]) i++;
    else out.push(b[j++]);
  }
  while (j < b.length) out.push(b[j++]);
  return out;
}

export interface ChangeSnap {
  key: string;
  type: string;
  author: string;
  text: string;
}

/** Hai đoạn chữ của hai sửa đổi có phải cùng một sửa đổi đã bị nới rộng hoặc sửa bên trong không:
 *  cái này chứa cái kia, hoặc phần đầu và phần cuối giống nhau chiếm từ nửa chuỗi ngắn hơn trở lên */
function related(x: string, y: string): boolean {
  if (!x || !y) return false;
  if (x.includes(y) || y.includes(x)) return true;
  const n = Math.min(x.length, y.length);
  let p = 0;
  while (p < n && x[p] === y[p]) p++;
  let s = 0;
  while (s < n - p && x[x.length - 1 - s] === y[y.length - 1 - s]) s++;
  return p + s >= Math.max(4, n / 2);
}

export function snapChange(t: { type: string; author: string; date: any; text: string }): ChangeSnap {
  return { key: changeKey(t), type: String(t.type), author: String(t.author), text: String(t.text ?? "") };
}

/** So hai lần chụp danh sách sửa đổi quanh một công cụ ghi.
 *  added  = sửa đổi mới hẳn: hoàn tác bằng cách từ chối đúng các sửa đổi này.
 *  merged = sửa đổi mới nhưng là bản nới rộng / thay thế của một sửa đổi có từ trước (Word gộp phần chèn liền kề của cùng tác giả,
 *           và sửa bên trong một phần chèn đang chờ duyệt thì không sinh sửa đổi mới): từ chối sẽ kéo theo cả phần có từ trước.
 *  removed = số sửa đổi có từ trước đã biến mất hoặc đổi nội dung. merged hoặc removed khác 0 thì phải hoàn tác bằng cách khác. */
export function diffChanges(before: ChangeSnap[], after: ChangeSnap[]): { added: string[]; merged: string[]; removed: number } {
  const bk = before.map((x) => x.key);
  const ak = after.map((x) => x.key);
  const addedK = addedKeys(bk, ak);
  const removedK = addedKeys(ak, bk);
  if (!removedK.length) return { added: addedK, merged: [], removed: 0 };
  const byKey = (list: ChangeSnap[]) => {
    const m = new Map<string, ChangeSnap>();
    list.forEach((x) => !m.has(x.key) && m.set(x.key, x));
    return m;
  };
  const bm = byKey(before);
  const am = byKey(after);
  const gone = removedK.map((k) => bm.get(k) as ChangeSnap);
  const added: string[] = [];
  const merged: string[] = [];
  for (const k of addedK) {
    const a = am.get(k) as ChangeSnap;
    const kin = gone.some((r) => r.type === a.type && r.author === a.author && related(r.text, a.text));
    (kin ? merged : added).push(k);
  }
  return { added, merged, removed: removedK.length };
}

/** Chọn trong danh sách sửa đổi hiện có những mục khớp các khóa cần hoàn tác (mỗi khóa dùng một lần, ưu tiên mục đứng sau). Trả chỉ số, giảm dần. */
export function pickByKeys(current: string[], keys: string[]): number[] {
  const want = new Map<string, number>();
  keys.forEach((k) => want.set(k, (want.get(k) || 0) + 1));
  const picked: number[] = [];
  for (let i = current.length - 1; i >= 0; i--) {
    const c = want.get(current[i]) || 0;
    if (c > 0) {
      want.set(current[i], c - 1);
      picked.push(i);
    }
  }
  return picked;
}

/** Tìm đúng MỘT tác giả theo tên người dùng nêu, cho việc duyệt hàng loạt: trùng hẳn tên (không phân biệt hoa thường) thì lấy;
 *  không thì chỉ nhận khi đúng một tên chứa chuỗi đó. Trả tên thật, hoặc danh sách ứng viên (rỗng = không ai khớp).
 *  Khớp kiểu "chứa chuỗi" với nhiều tên ("An" khớp cả "Hoàng Anh") sẽ duyệt nhầm sửa đổi của người khác. */
export function resolveAuthor(authors: string[], want: string): string | string[] {
  const w = String(want || "")
    .trim()
    .toLowerCase();
  const uniq = Array.from(new Set(authors));
  if (!w) return [];
  const exact = uniq.filter((a) => a.trim().toLowerCase() === w);
  if (exact.length === 1) return exact[0];
  const part = uniq.filter((a) => a.toLowerCase().includes(w));
  return part.length === 1 ? part[0] : part;
}

/** Một sửa đổi cần hoàn tác: khóa, và (khi biết) định danh đoạn chứa nó — để không chọn nhầm một sửa đổi trùng khóa
 *  do chính người dùng tạo ở đoạn khác trong cùng phút. */
export interface Wanted {
  key: string;
  para?: string;
}

/** Chọn trong danh sách sửa đổi hiện có những mục cần hoàn tác. Mỗi mục cần một sửa đổi cùng khóa; biết đoạn thì ưu tiên sửa đổi nằm
 *  trong đúng đoạn đó, không biết thì lấy mục đứng sau cùng chưa dùng. Trả chỉ số, giảm dần. */
export function pickWanted(current: { key: string; para?: string }[], wants: Wanted[]): number[] {
  const used = new Set<number>();
  const take = (w: Wanted, exact: boolean): boolean => {
    for (let i = current.length - 1; i >= 0; i--) {
      if (used.has(i) || current[i].key !== w.key) continue;
      if (exact && current[i].para !== w.para) continue;
      used.add(i);
      return true;
    }
    return false;
  };
  const loose: Wanted[] = [];
  // Mục biết đoạn chọn trước, đúng đoạn; không thấy đúng đoạn thì xếp vào nhóm chọn theo khóa
  for (const w of wants) if (!(w.para && take(w, true))) loose.push(w);
  for (const w of loose) take(w, false);
  return Array.from(used).sort((a, b) => b - a);
}

// ───────────────────────── Chống sửa nhầm đoạn ─────────────────────────

const squash = (s: string) =>
  String(s ?? "")
    .replace(/[\s ]+/g, " ")
    .replace(/[…]+$|\.{3,}$/g, "")
    .trim();

/** Đoạn có đúng là đoạn AI đã đọc không: `expect` là mấy chữ đầu của đoạn (hoặc một cụm nằm trong đoạn). */
export function matchesExpect(text: string, expect: string): boolean {
  const e = squash(expect);
  if (!e) return true;
  const t = squash(text);
  return t.startsWith(e) || (e.length >= 12 && t.includes(e));
}

export type Located =
  | { ok: true; index: number; moved: boolean }
  | { ok: false; reason: string; candidates: number[] };

/** Định vị đoạn cần sửa: đúng chỉ số và khớp `expect` thì dùng; lệch thì tìm đoạn duy nhất khớp `expect` (ưu tiên gần chỉ số cũ);
 *  không có hoặc có nhiều đoạn khớp thì từ chối. Không có `expect` thì dùng chỉ số như cũ. */
export function locateParagraph(texts: string[], index: any, expect?: string): Located {
  const i = Number(index);
  const valid = Number.isInteger(i) && i >= 0 && i < texts.length;
  const e = squash(expect || "");
  if (!e) {
    if (!valid) return { ok: false, reason: `Không có đoạn số ${index}. Văn bản có ${texts.length} đoạn (đánh số từ 0).`, candidates: [] };
    return { ok: true, index: i, moved: false };
  }
  if (valid && matchesExpect(texts[i], e)) return { ok: true, index: i, moved: false };
  const starts: number[] = [];
  const contains: number[] = [];
  texts.forEach((t, k) => {
    const q = squash(t);
    if (q.startsWith(e)) starts.push(k);
    else if (e.length >= 12 && q.includes(e)) contains.push(k);
  });
  const hits = starts.length ? starts : contains;
  if (hits.length === 1) return { ok: true, index: hits[0], moved: true };
  if (!hits.length) {
    return {
      ok: false,
      reason: `Đoạn ${index} không bắt đầu bằng “${e.slice(0, 60)}” và không tìm thấy đoạn nào như vậy: văn bản đã đổi so với lúc đọc. Đọc lại rồi mới sửa.`,
      candidates: [],
    };
  }
  return {
    ok: false,
    reason: `Đoạn ${index} không bắt đầu bằng “${e.slice(0, 60)}”, và có ${hits.length} đoạn khác khớp nên không tự chọn được. Đọc lại các đoạn ứng viên rồi gọi lại với chỉ số đúng.`,
    candidates: hits.slice(0, 8),
  };
}
