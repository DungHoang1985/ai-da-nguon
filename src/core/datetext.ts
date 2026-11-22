/** Chữ dạng ngày → số ngày của Excel (Đợt 12, phần B — bài X3 ngày 22/11/2026: cột ngày ở ô định dạng chữ "@", AI đặt định dạng
 *  dd/mm/yyyy nhưng ô vẫn là chữ nên không đổi gì). Hàm thuần, kiểm thử được.
 *  Đọc theo thứ tự NGÀY / THÁNG / NĂM (cách viết ở Việt Nam), hoặc NĂM-THÁNG-NGÀY khi năm đứng đầu. Ngày không có thật (31/02) → null. */

const DMY = /^\s*(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{4})\s*$/;
const YMD = /^\s*(\d{4})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*$/;
const EPOCH = Date.UTC(1899, 11, 30);

/** Số ngày Excel của chữ dạng ngày, hoặc null nếu không phải ngày có thật */
export function dateTextToSerial(s: unknown): number | null {
  if (typeof s !== "string") return null;
  let y: number, m: number, d: number;
  const a = DMY.exec(s);
  const b = a ? null : YMD.exec(s);
  if (a) [d, m, y] = [Number(a[1]), Number(a[2]), Number(a[3])];
  else if (b) [y, m, d] = [Number(b[1]), Number(b[2]), Number(b[3])];
  else return null;
  if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const t = Date.UTC(y, m - 1, d);
  const back = new Date(t);
  if (back.getUTCFullYear() !== y || back.getUTCMonth() !== m - 1 || back.getUTCDate() !== d) return null;
  return Math.round((t - EPOCH) / 86400000);
}

/** Định dạng số có phải định dạng ngày: có ký hiệu ngày (d) hoặc năm (y), không tính phần trong ngoặc vuông và trong nháy
 *  ("mm:ss" chỉ là giờ phút, không tính) */
export function isDateNumberFormat(nf: unknown): boolean {
  const s = String(nf || "");
  if (!s || /^general$/i.test(s) || s === "@") return false;
  return /[dy]/i.test(s.replace(/\[[^\]]*\]|"[^"]*"/g, ""));
}

export interface DateFix {
  /** Ma trận công thức mới (chỉ ô chữ dạng ngày được đổi thành số ngày) */
  formulas: any[][];
  /** Số ô đã đổi */
  converted: number;
  /** Vị trí (hàng, cột từ 0) các ô chữ không đọc được thành ngày */
  leftText: [number, number][];
}

/** Đổi các ô chữ dạng ngày trong ma trận công thức; ô rỗng, ô số, ô công thức ("=…") giữ nguyên */
export function convertDateTexts(formulas: any[][]): DateFix {
  let converted = 0;
  const leftText: [number, number][] = [];
  const out = formulas.map((row, i) =>
    row.map((v, j) => {
      if (typeof v !== "string" || !v.trim() || v.startsWith("=")) return v;
      const n = dateTextToSerial(v);
      if (n === null) {
        leftText.push([i, j]);
        return v;
      }
      converted++;
      return n;
    }),
  );
  return { formulas: out, converted, leftText };
}
