/** So hai phiên bản của một văn bản theo từng đoạn và gắn mỗi chỗ khác với số điều khoản.
 *  Hàm thuần, không cần Office — kiểm thử bằng npm test. A = văn bản đang mở, B = phiên bản khác (tệp đính kèm).
 *  Thăm dò 15/11/2026 (tasks/dot/dot-1b/dot-1b.md): kết quả so sánh của chính Word (compareFromBase64) nằm ở một tài liệu khác mà add-in
 *  không đọc được, còn ghi thẳng vào tài liệu đang mở thì làm mất sửa đổi chờ duyệt — nên phần tóm tắt cho AI tự so bằng mã. */
import { OutlineItem, ParaRow, buildOutline, clauseAt, clauseLabel } from "./outline";
import { diffWords } from "./worddiff";

export interface VersionDiff {
  /** changed = đoạn có ở cả hai bản nhưng chữ khác; only_current = chỉ có ở bản đang mở; only_other = chỉ có ở phiên bản khác */
  type: "changed" | "only_current" | "only_other";
  /** Số điều khoản theo bản đang mở (với only_other: điều khoản đứng ngay trước chỗ đoạn đó sẽ nằm) */
  clause: string;
  /** Số điều khoản theo phiên bản khác, khi khác với bản đang mở */
  clause_other?: string;
  /** Chỉ số đoạn trong bản đang mở (với only_other: đoạn đứng ngay trước) */
  p?: number;
  /** Chỉ số đoạn trong phiên bản khác */
  p_other?: number;
  /** Từng chỗ đổi trong đoạn: "“cũ” → “mới”" (cũ = bản đang mở, mới = phiên bản khác) */
  changes?: string[];
  /** Chữ của đoạn (rút gọn): bản đang mở và phiên bản khác */
  current?: string;
  other?: string;
}

export interface CompareResult {
  /** Số đoạn có chữ của mỗi bản */
  paragraphs_current: number;
  paragraphs_other: number;
  identical: boolean;
  diffs: VersionDiff[];
  /** Điều khoản giữ nguyên chữ nhưng mang số khác ở hai bản (hệ quả của việc thêm, bớt khoản) */
  renumbered: { current: string; other: string; p: number }[];
}

interface Line {
  /** Chỉ số đoạn gốc */
  i: number;
  text: string;
  norm: string;
}

const norm = (s: string) =>
  String(s ?? "")
    .replace(/[\u0001\u0002\u0007]/g, "")
    .replace(/[ \t]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

function toLines(rows: ParaRow[]): Line[] {
  const out: Line[] = [];
  for (const r of rows) {
    if (r.deleted) continue;
    const n = norm(r.text);
    if (n) out.push({ i: r.i, text: n, norm: n });
  }
  return out;
}

/** Các cặp chỉ số khớp nhau (dãy con chung dài nhất) giữa hai dãy chuỗi */
export function lcsPairs(a: string[], b: string[]): [number, number][] {
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let s = 0;
  while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
  const pairs: [number, number][] = [];
  for (let k = 0; k < p; k++) pairs.push([k, k]);
  const n = a.length - p - s;
  const m = b.length - p - s;
  if (n > 0 && m > 0) {
    if (n * m <= 9_000_000) {
      const w = m + 1;
      const dp = new Uint16Array((n + 1) * w);
      for (let i = n - 1; i >= 0; i--) {
        for (let j = m - 1; j >= 0; j--) {
          dp[i * w + j] = a[p + i] === b[p + j] ? dp[(i + 1) * w + j + 1] + 1 : Math.max(dp[(i + 1) * w + j], dp[i * w + j + 1]);
        }
      }
      let i = 0;
      let j = 0;
      while (i < n && j < m) {
        if (a[p + i] === b[p + j]) pairs.push([p + i++, p + j++]);
        else if (dp[(i + 1) * w + j] >= dp[i * w + j + 1]) i++;
        else j++;
      }
    } else {
      // Văn bản rất dài: neo vào các đoạn chỉ xuất hiện đúng một lần ở mỗi bản, theo thứ tự tăng dần, rồi so từng khúc giữa hai mốc
      const count = (list: string[], from: number, len: number) => {
        const c = new Map<string, number[]>();
        for (let k = 0; k < len; k++) c.set(list[from + k], [...(c.get(list[from + k]) || []), from + k]);
        return c;
      };
      const ca = count(a, p, n);
      const cb = count(b, p, m);
      const cand: [number, number][] = [];
      for (const [t, ia] of ca) {
        const ib = cb.get(t);
        if (ia.length === 1 && ib && ib.length === 1) cand.push([ia[0], ib[0]]);
      }
      cand.sort((x, y) => x[0] - y[0]);
      // Dãy con tăng dài nhất theo chỉ số ở b
      const tails: number[] = [];
      const prev: number[] = new Array(cand.length).fill(-1);
      for (let k = 0; k < cand.length; k++) {
        let lo = 0;
        let hi = tails.length;
        while (lo < hi) {
          const mid = (lo + hi) >> 1;
          if (cand[tails[mid]][1] < cand[k][1]) lo = mid + 1;
          else hi = mid;
        }
        prev[k] = lo ? tails[lo - 1] : -1;
        tails[lo] = k;
      }
      const anchors: [number, number][] = [];
      for (let k = tails.length ? tails[tails.length - 1] : -1; k >= 0; k = prev[k]) anchors.unshift(cand[k]);
      let ai = p;
      let bi = p;
      const gap = (ae: number, be: number) => {
        // Khúc giữa hai mốc: nếu vẫn quá lớn thì bỏ qua (coi là khác hết), nếu không thì so bằng bảng
        if ((ae - ai) * (be - bi) <= 9_000_000) for (const [x, y] of lcsPairs(a.slice(ai, ae), b.slice(bi, be))) pairs.push([ai + x, bi + y]);
      };
      for (const [x, y] of anchors) {
        gap(x, y);
        pairs.push([x, y]);
        ai = x + 1;
        bi = y + 1;
      }
      gap(p + n, p + m);
    }
  }
  for (let k = 0; k < s; k++) pairs.push([a.length - s + k, b.length - s + k]);
  return pairs;
}

const words = (s: string) => s.toLowerCase().match(/[\p{L}\p{N}]+/gu) || [];
/** Độ giống nhau của hai đoạn theo từ (0–1) */
export function similarity(x: string, y: string): number {
  const a = words(x);
  const b = words(y);
  if (!a.length || !b.length) return 0;
  const count = new Map<string, number>();
  a.forEach((w) => count.set(w, (count.get(w) || 0) + 1));
  let common = 0;
  for (const w of b) {
    const c = count.get(w) || 0;
    if (c > 0) {
      common++;
      count.set(w, c - 1);
    }
  }
  return (2 * common) / (a.length + b.length);
}

export const PAIR_THRESHOLD = 0.5;

/** Ghép các đoạn không khớp của một khúc theo thứ tự: cặp nào đủ giống thì là "đoạn bị sửa", còn lại là thêm / bớt */
function pairGap(ga: Line[], gb: Line[]): [number, number][] {
  const n = ga.length;
  const m = gb.length;
  if (!n || !m) return [];
  if (n * m > 40_000) {
    // Khúc quá lớn: ghép tuần tự
    const out: [number, number][] = [];
    let j = 0;
    for (let i = 0; i < n && j < m; i++) {
      for (let k = j; k < Math.min(m, j + 8); k++) {
        if (similarity(ga[i].norm, gb[k].norm) >= PAIR_THRESHOLD) {
          out.push([i, k]);
          j = k + 1;
          break;
        }
      }
    }
    return out;
  }
  const sim: number[][] = ga.map((x) => gb.map((y) => similarity(x.norm, y.norm)));
  const w = m + 1;
  const best = new Float64Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      let v = Math.max(best[(i + 1) * w + j], best[i * w + j + 1]);
      if (sim[i][j] >= PAIR_THRESHOLD) v = Math.max(v, sim[i][j] + best[(i + 1) * w + j + 1]);
      best[i * w + j] = v;
    }
  }
  const out: [number, number][] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    const here = best[i * w + j];
    if (sim[i][j] >= PAIR_THRESHOLD && Math.abs(here - (sim[i][j] + best[(i + 1) * w + j + 1])) < 1e-9) out.push([i++, j++]);
    else if (Math.abs(here - best[(i + 1) * w + j]) < 1e-9) i++;
    else j++;
  }
  return out;
}

const cut = (s: string, n: number) => (s.length > n ? s.slice(0, n) + "…" : s);

/** Mô tả các chỗ đổi trong một đoạn: "“cũ” → “mới”", "thêm “…”", "bỏ “…”" */
export function describeChanges(current: string, other: string, max = 6): string[] {
  const hunks = diffWords(current, other);
  // Từ bị đổi ngắn quá ("5" → "7") thì kèm vài chữ đứng trước cho dễ nhận ra
  const ctx = (at: number) => {
    const before = current.slice(Math.max(0, at - 24), at);
    const k = before.indexOf(" ");
    return (k >= 0 && at > 24 ? before.slice(k + 1) : before).trimStart();
  };
  const out = hunks.slice(0, max).map((h) => {
    const del = h.del.trim();
    const ins = h.ins.trim();
    const lead = del.length + ins.length < 24 ? ctx(h.at) : "";
    const pre = lead ? `…${lead}` : "";
    if (del && ins) return `${pre}“${cut(del, 140)}” → “${cut(ins, 140)}”`;
    if (ins) return `${pre ? pre + " " : ""}thêm “${cut(ins, 220)}”`;
    return `${pre ? pre + " " : ""}bỏ “${cut(del, 220)}”`;
  });
  if (hunks.length > max) out.push(`… và ${hunks.length - max} chỗ đổi nữa trong đoạn`);
  return out;
}

export function compareVersions(current: ParaRow[], other: ParaRow[], opts: { outlineCurrent?: OutlineItem[]; outlineOther?: OutlineItem[] } = {}): CompareResult {
  const oa = opts.outlineCurrent || buildOutline(current);
  const ob = opts.outlineOther || buildOutline(other);
  const A = toLines(current);
  const B = toLines(other);
  const pairs = lcsPairs(
    A.map((l) => l.norm),
    B.map((l) => l.norm)
  );
  const nameA = (i: number) => clauseLabel(oa, clauseAt(oa, i));
  const nameB = (i: number) => clauseLabel(ob, clauseAt(ob, i));
  const diffs: VersionDiff[] = [];
  const renumbered: CompareResult["renumbered"] = [];
  const headA = new Set(oa.map((o) => o.i));
  let ai = 0;
  let bi = 0;
  const flush = (ae: number, be: number) => {
    const ga = A.slice(ai, ae);
    const gb = B.slice(bi, be);
    const matched = pairGap(ga, gb);
    let x = 0;
    let y = 0;
    const onlyCurrent = (l: Line) => diffs.push({ type: "only_current", clause: nameA(l.i), p: l.i, current: cut(l.text, 400) });
    const onlyOther = (l: Line) => {
      // Chỗ đoạn này sẽ nằm trong bản đang mở: ngay sau đoạn đứng trước nó
      const prevLine = x > 0 ? ga[x - 1] : ai > 0 ? A[ai - 1] : null;
      const cb = nameB(l.i);
      diffs.push({ type: "only_other", clause: prevLine ? nameA(prevLine.i) : "", ...(cb ? { clause_other: cb } : {}), ...(prevLine ? { p: prevLine.i } : {}), p_other: l.i, other: cut(l.text, 400) });
    };
    for (const [mx, my] of matched) {
      while (x < mx) onlyCurrent(ga[x++]);
      while (y < my) onlyOther(gb[y++]);
      const la = ga[mx];
      const lb = gb[my];
      const ca = nameA(la.i);
      const cb = nameB(lb.i);
      diffs.push({ type: "changed", clause: ca, ...(cb && cb !== ca ? { clause_other: cb } : {}), p: la.i, p_other: lb.i, changes: describeChanges(la.text, lb.text), current: cut(la.text, 260), other: cut(lb.text, 260) });
      x = mx + 1;
      y = my + 1;
    }
    while (x < ga.length) onlyCurrent(ga[x++]);
    while (y < gb.length) onlyOther(gb[y++]);
  };
  for (const [pa, pb] of pairs) {
    flush(pa, pb);
    // Đoạn giống hệt nhưng mang số điều khác
    const la = A[pa];
    if (headA.has(la.i)) {
      const ca = nameA(la.i);
      const cb = nameB(B[pb].i);
      if (ca && cb && ca !== cb) renumbered.push({ current: ca, other: cb, p: la.i });
    }
    ai = pa + 1;
    bi = pb + 1;
  }
  flush(A.length, B.length);
  return { paragraphs_current: A.length, paragraphs_other: B.length, identical: diffs.length === 0, diffs, renumbered };
}
