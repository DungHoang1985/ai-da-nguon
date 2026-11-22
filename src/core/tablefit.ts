/** Bảng PowerPoint vừa slide (22/11/2026): chia bề ngang cột theo độ dài chữ và chọn cỡ chữ nhỏ dần khi bảng tràn đáy slide.
 *  Lõi thuần, không gọi Office.js; phần đo chiều cao thật nằm ở `ppt_add_table` (`src/office/pptTools.ts`).
 *  Đã thăm dò trong PowerPoint (22/11/2026, bảng 7 × 4 của bài P3): cỡ mặc định 18 cao 504 pt; cỡ 12 cao 266 pt; đặt cỡ chữ từng ô
 *  sau khi thêm bảng giữ kiểu bảng và phông chủ đề, còn `uniformCellProperties` làm mất kiểu bảng (ô hết tô màu). */

/** Cỡ chữ thử lần lượt khi bảng tràn; không xuống dưới 10 (nhỏ hơn thì khó đọc khi trình chiếu). */
export const TABLE_FONT_STEPS = [16, 14, 13, 12, 11, 10];
export const TABLE_MIN_FONT = 10;
/** Khoảng chừa ở đáy slide (điểm). */
export const TABLE_BOTTOM_GAP = 8;

/** Cỡ chữ kế tiếp nhỏ hơn `cur` để thử; null khi đã ở mức nhỏ nhất. */
export function nextTableFont(cur: number): number | null {
  for (const s of TABLE_FONT_STEPS) if (s < cur) return s;
  return null;
}

/** Bảng có lọt trong slide không (đáy bảng cách đáy slide ít nhất TABLE_BOTTOM_GAP). */
export function tableFits(top: number, height: number, slideH: number): boolean {
  return top + height <= slideH - TABLE_BOTTOM_GAP + 0.5;
}

/** Bề ngang từng cột theo độ dài chữ dài nhất của cột (đo theo từ dài nhất và cả ô), tổng đúng bằng `total`.
 *  Cột ngắn như "Hạn" không chiếm ngang bằng cột "Thủ tục"; cột nào cũng có ít nhất 8% bề ngang. */
export function tableColumnWidths(values: string[][], total: number): number[] {
  const n = values[0]?.length || 0;
  if (!n) return [];
  const weight: number[] = [];
  for (let k = 0; k < n; k++) {
    let cell = 0;
    let word = 0;
    for (const r of values) {
      const s = String(r[k] ?? "");
      cell = Math.max(cell, s.length);
      for (const w of s.split(/\s+/)) word = Math.max(word, w.length);
    }
    // ô dài sẽ xuống dòng nên không cần ngang hết cỡ: lấy căn bậc hai cho ô dài, nhưng không hẹp hơn từ dài nhất
    weight.push(Math.max(word * 1.2, Math.sqrt(Math.max(cell, 1)) * 4, 4));
  }
  const sum = weight.reduce((a, b) => a + b, 0);
  // Soát mã 22/11: từ 13 cột trở lên, sàn 8% × số cột vượt 100% → cột cuối ra âm. Sàn không quá phần chia đều; chia đều khi sàn chiếm hết.
  const floor = Math.min(total * 0.08, total / n);
  if (floor * n >= total - 0.01) {
    const eq = Array(n).fill(Math.round((total / n) * 10) / 10);
    eq[n - 1] = Math.round((total - eq.slice(0, -1).reduce((a: number, b: number) => a + b, 0)) * 10) / 10;
    return eq;
  }
  let w = weight.map((x) => Math.max(floor, (total * x) / sum));
  const over = w.reduce((a, b) => a + b, 0) - total;
  if (over > 0) {
    const room = w.reduce((a, x) => a + Math.max(0, x - floor), 0) || 1;
    w = w.map((x) => x - (over * Math.max(0, x - floor)) / room);
  }
  const out = w.map((x) => Math.round(x * 10) / 10);
  out[n - 1] = Math.round((total - out.slice(0, -1).reduce((a, b) => a + b, 0)) * 10) / 10;
  return out;
}
