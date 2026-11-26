/** Đã đo trong Excel 26/11/2026: Office.js `range.text` trả chữ đã định dạng ("123.456.789") nhưng KHÔNG trả "####" khi cột hẹp
 *  (cột 12pt vẫn ra đủ chữ). Vì vậy ô cột hẹp phải tự ước: độ rộng chữ theo cỡ chữ so với độ rộng cột. */

/** Chữ hiển thị sau khi định dạng ô Excel (mức 2 việc 5, 26/11/2026): lệnh định dạng làm đổi cách hiện (định dạng số, độ rộng cột,
 *  cỡ chữ) tự trả chữ người dùng thấy và các ô hiện "####" (cột hẹp), để model không phải gọi thêm lệnh đọc — `excel_read_range`
 *  trả giá trị, không trả chữ hiển thị. Lời nhắc hệ thống coi kết quả lệnh ghi có nội dung sau khi ghi là đã đọc lại. */

/** Phần xem trước gửi model: tối đa ngần này hàng × cột (giữ kết quả công cụ nhỏ cho chế độ Miễn phí) */
export const DISPLAY_PREVIEW_ROWS = 6;
export const DISPLAY_PREVIEW_COLS = 8;
/** Tối đa số ô "####" nêu tên trong kết quả */
export const NARROW_LIST_MAX = 8;

export interface DisplayCheck {
  /** Chữ hiển thị của góc trên trái vùng */
  preview: string[][];
  /** Địa chỉ các ô số mà chữ ước rộng hơn cột (Excel sẽ hiện "####") */
  narrow: string[];
  narrowTotal: number;
}

/** Độ rộng ước (pt) của chữ một ô số, theo số đo "tự giãn cột" trong Excel 26/11/2026 (Calibri cỡ 11 → hệ số ×cỡ chữ):
 *  chữ số 5,4pt (0,49·cỡ); dấu chấm, phẩy, khoảng trắng 2,7–3pt (0,26); "/" và ":" 4,2pt (0,38); "%" 7,5pt (0,68); ký tự khác 0,55;
 *  lề ô 5,4pt (0,49). Đo: "12345678" 48,6 · "09/12/2025" 57 · "123.456.789" 60 · "1.234,50" 43,2 · "12,5%" 31,8 · cỡ 20 "12345678" 90. */
export function estimateWidth(text: string, size = 11): number {
  let em = 0.49;
  for (const ch of String(text)) em += /[0-9]/.test(ch) ? 0.49 : /[.,\s]/.test(ch) ? 0.26 : /[/:]/.test(ch) ? 0.38 : ch === "%" ? 0.68 : 0.55;
  return em * size;
}

/** text, values: ma trận cùng cỡ của vùng (Office.js `range.text`, `range.values`); addr(i, j): địa chỉ ô ở hàng i, cột j của vùng;
 *  widths[j]: độ rộng cột j (pt), null = không biết (không xét cột đó); size: cỡ chữ của vùng (null = trộn nhiều cỡ → 11).
 *  Chỉ xét ô có giá trị là số (số, ngày, giờ): chữ thường tràn sang ô bên cạnh chứ không hiện "####". Dư 1pt để không báo nhầm sát mép. */
export function checkDisplay(text: string[][], values: any[][], addr: (i: number, j: number) => string, widths: (number | null)[] = [], size: number | null = 11): DisplayCheck {
  const narrow: string[] = [];
  let narrowTotal = 0;
  text.forEach((r, i) =>
    r.forEach((t, j) => {
      const w = widths[j];
      if (typeof values?.[i]?.[j] !== "number" || typeof w !== "number" || w <= 0) return;
      if (estimateWidth(String(t ?? ""), size || 11) > w + 1) {
        narrowTotal++;
        if (narrow.length < NARROW_LIST_MAX) narrow.push(addr(i, j));
      }
    })
  );
  const preview = text.slice(0, DISPLAY_PREVIEW_ROWS).map((r) => r.slice(0, DISPLAY_PREVIEW_COLS).map((t) => String(t ?? "").slice(0, 40)));
  return { preview, narrow, narrowTotal };
}

/** Câu nối vào kết quả lệnh định dạng khi có ô "####" */
export function narrowNote(d: DisplayCheck): string {
  if (!d.narrowTotal) return "";
  return `; CHÚ Ý: ${d.narrowTotal} ô số có thể hiện "####" vì cột hẹp hơn chữ (ước theo cỡ chữ) (${d.narrow.join(", ")}${d.narrowTotal > d.narrow.length ? ", …" : ""}) — nới cột (column_width hoặc autofit_columns) rồi mới báo xong`;
}
