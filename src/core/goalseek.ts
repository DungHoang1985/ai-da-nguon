/** Giải ngược (tìm đầu vào để đầu ra đạt mục tiêu) thay cho Goal Seek mà Excel không mở cho add-in (Đợt 4). Hàm thuần: hàm đánh
 *  giá f(x) do bên gọi đưa vào (trong Office: ghi ô đầu vào, tính lại, đọc ô kết quả) nên thử được ngoài Office. Bên gọi chịu trách
 *  nhiệm trả ô về giá trị gốc. */

export interface GoalSeekOptions {
  target: number;
  /** Giá trị hiện tại của ô đầu vào */
  start: number;
  min?: number;
  max?: number;
  /** Sai số tương đối chấp nhận so với mục tiêu (mặc định 0,001% — chặt hơn tiêu chí 0,1%) */
  tolRel?: number;
  maxIter?: number;
}

export interface GoalSeekResult {
  ok: boolean;
  /** Đầu vào tốt nhất tìm được */
  x: number;
  /** Đầu ra tại x */
  y: number;
  iterations: number;
  /** Sai số tương đối so với mục tiêu */
  error: number;
  reason?: string;
}

const fin = (v: any): v is number => typeof v === "number" && Number.isFinite(v);

export async function goalSeek(f: (x: number) => Promise<number>, o: GoalSeekOptions): Promise<GoalSeekResult> {
  const T = o.target;
  const tolRel = fin(o.tolRel) && o.tolRel > 0 ? o.tolRel : 1e-5;
  // Thang đo sai số: theo mục tiêu; mục tiêu = 0 (hòa vốn, NPV = 0) thì theo độ lớn đầu ra ở điểm xuất phát — đặt sau lần tính đầu
  let scale = Math.abs(T);
  let absTol = tolRel * scale;
  const maxIter = Math.max(5, Math.min(o.maxIter ?? 60, 200));
  const num = (v: any) => (v === null || v === undefined || v === "" ? NaN : Number(v));
  const lo = fin(num(o.min)) ? num(o.min) : -Infinity;
  const hi = fin(num(o.max)) ? num(o.max) : Infinity;
  if (lo > hi) return { ok: false, x: o.start, y: NaN, iterations: 0, error: NaN, reason: "min lớn hơn max" };
  const clamp = (x: number) => Math.min(hi, Math.max(lo, x));
  let n = 0;
  let best = { x: NaN, y: NaN, d: Infinity };
  const g = async (x: number) => {
    n++;
    const y = await f(x);
    if (!fin(y)) return NaN;
    const d = y - T;
    if (Math.abs(d) < Math.abs(best.d)) best = { x, y, d };
    return d;
  };
  const done = (ok: boolean, reason?: string): GoalSeekResult => ({
    ok,
    x: best.x,
    y: best.y,
    iterations: n,
    error: fin(best.d) ? Math.abs(best.d) / Math.max(scale, 1e-12) : NaN,
    ...(reason ? { reason } : {}),
  });

  let x0 = clamp(fin(o.start) ? o.start : 0);
  let d0 = await g(x0);
  if (Number.isNaN(d0)) return done(false, `Ô kết quả không ra số khi ô đầu vào = ${x0}`);
  if (scale === 0) {
    scale = Math.max(Math.abs(d0), 1e-9);
    absTol = tolRel * scale;
  }
  if (Math.abs(d0) <= absTol) return done(true);

  // Tìm hai điểm kẹp nghiệm (đầu ra nằm hai phía của mục tiêu) bằng cách nới dần quanh điểm xuất phát
  let a = x0;
  let da = d0;
  let b = NaN;
  let db = NaN;
  const span = Math.abs(x0) > 1e-9 ? Math.abs(x0) * 0.1 : 1;
  let bracket = false;
  for (let k = 0; k < 40 && n < maxIter && !bracket; k++) {
    const step = span * Math.pow(2, k);
    for (const cand of [clamp(x0 + step), clamp(x0 - step)]) {
      if (cand === a || cand === b) continue;
      const d = await g(cand);
      if (Number.isNaN(d)) continue;
      if (Math.abs(d) <= absTol) return done(true);
      if (d * d0 < 0) {
        b = cand;
        db = d;
        bracket = true;
        break;
      }
      // Giữ điểm cùng phía gần nghiệm hơn làm đầu kẹp còn lại
      if (Math.abs(d) < Math.abs(da)) {
        a = cand;
        da = d;
      }
    }
    if (clamp(x0 + step) === hi && clamp(x0 - step) === lo && !bracket) break;
  }
  if (!bracket) return done(false, "Không tìm được đầu vào nào (trong khoảng cho phép) làm ô kết quả vượt qua mục tiêu: mục tiêu có thể không đạt được, hoặc ô kết quả không phụ thuộc ô đầu vào");

  // Dây cung có chặn (Illinois), rơi về chia đôi khi dây cung ra ngoài khoảng kẹp
  let side = 0;
  while (n < maxIter) {
    let x = (a * db - b * da) / (db - da);
    if (!fin(x) || x <= Math.min(a, b) || x >= Math.max(a, b)) x = (a + b) / 2;
    const d = await g(x);
    if (Number.isNaN(d)) {
      x = (a + b) / 2;
      const d2 = await g(x);
      if (Number.isNaN(d2)) return done(false, `Ô kết quả không ra số khi ô đầu vào = ${x}`);
      if (Math.abs(d2) <= absTol) return done(true);
      if (d2 * da < 0) [b, db] = [x, d2];
      else [a, da] = [x, d2];
      continue;
    }
    if (Math.abs(d) <= absTol) return done(true);
    if (d * da < 0) {
      b = x;
      db = d;
      if (side === -1) da /= 2;
      side = -1;
    } else {
      a = x;
      da = d;
      if (side === 1) db /= 2;
      side = 1;
    }
    if (Math.abs(b - a) <= 1e-14 * Math.max(1, Math.abs(a))) break;
  }
  // Hết số vòng: chấp nhận nếu đã trong sai số 0,1% (hàm bậc thang, làm tròn)
  const rel = Math.abs(best.d) / Math.max(scale, 1e-12);
  return rel <= 1e-3 ? done(true) : done(false, "Chưa hội tụ trong số vòng cho phép; đây là đầu vào gần mục tiêu nhất tìm được");
}

/** Làm gọn số tìm được để hiện cho người đọc (10 chữ số có nghĩa) */
export function tidy(x: number): number {
  return fin(x) ? Number(x.toPrecision(10)) : x;
}

/** Danh sách giá trị thử của bảng nhạy cảm: chỉ giữ số hữu hạn, tối đa `max` giá trị. */
export function cleanNumberList(values: any, max = 15): number[] {
  const arr: any[] = Array.isArray(values) ? values : values === undefined || values === null ? [] : [values];
  return arr
    .map((v) => (typeof v === "string" ? Number(v.replace(/%$/, "")) / (v.trim().endsWith("%") ? 100 : 1) : Number(v)))
    .filter((v) => Number.isFinite(v))
    .slice(0, max);
}
