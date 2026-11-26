/** Mức 2 việc 5 (26/11/2026): lệnh định dạng Excel trả chữ hiển thị và ô số có thể hiện "####" (src/core/exceldisplay.ts).
 *  Office.js không trả "####" (đã đo) nên ô cột hẹp được ước theo cỡ chữ và độ rộng cột. */
import * as assert from "assert";
import { DISPLAY_PREVIEW_COLS, DISPLAY_PREVIEW_ROWS, NARROW_LIST_MAX, checkDisplay, estimateWidth, narrowNote } from "../../src/core/exceldisplay";

const addr = (i: number, j: number) => `${String.fromCharCode(66 + j)}${2 + i}`; // vùng bắt đầu ở B2

describe("exceldisplay — chữ hiển thị sau khi định dạng", () => {
  it("khớp số đo tự giãn cột trong Excel 26/11/2026 (sai số dưới 2pt)", () => {
    const measured: [string, number, number][] = [
      ["12345678", 11, 48.6],
      ["09/12/2025", 11, 57],
      ["123.456.789", 11, 60],
      ["1.234,50", 11, 43.2],
      ["1", 11, 10.8],
      ["11111111111", 11, 64.8],
      ["12345678", 20, 90],
      ["12,5%", 11, 31.8],
    ];
    for (const [t, sz, w] of measured) assert.ok(Math.abs(estimateWidth(t, sz) - w) < 2, `${t} cỡ ${sz}: ước ${estimateWidth(t, sz).toFixed(1)}, đo ${w}`);
  });

  it("bắt ô số rộng hơn cột và nêu đúng địa chỉ", () => {
    const d = checkDisplay([["1.234,50", "123.456.789"], ["09/12/2025", "abc"]], [[1234.5, 123456789], [46000, "abc"]], addr, [48, 12]);
    // Ngày "09/12/2025" cần 57pt (đo) nên cột mặc định 48pt hiện "####"
    assert.deepStrictEqual(d.narrow, ["C2", "B3"]);
    assert.strictEqual(d.narrowTotal, 2);
    assert.match(narrowNote(d), /2 ô số có thể hiện "####".*C2, B3.*column_width hoặc autofit_columns/);
  });

  it("ô chữ dài không tính (chữ tràn sang ô bên, không hiện ####); cột không biết độ rộng thì bỏ qua", () => {
    const d = checkDisplay([["một đoạn chữ rất dài", "123.456.789"]], [["một đoạn chữ rất dài", 123456789]], addr, [12, null]);
    assert.strictEqual(d.narrowTotal, 0);
    assert.strictEqual(narrowNote(d), "");
  });

  it("cỡ chữ trộn (null) thì ước theo cỡ 11", () => {
    const d = checkDisplay([["12345678"]], [[12345678]], addr, [48], null);
    assert.strictEqual(d.narrowTotal, 0);
  });

  it("xem trước bị cắt theo hàng, cột và độ dài chữ", () => {
    const big = Array.from({ length: 20 }, () => Array.from({ length: 20 }, () => "x".repeat(60)));
    const vals = big.map((r) => r.map(() => "x"));
    const d = checkDisplay(big, vals, addr);
    assert.strictEqual(d.preview.length, DISPLAY_PREVIEW_ROWS);
    assert.strictEqual(d.preview[0].length, DISPLAY_PREVIEW_COLS);
    assert.strictEqual(d.preview[0][0].length, 40);
  });

  it("danh sách ô có trần, câu nhắc ghi tổng và dấu …", () => {
    const text = [Array.from({ length: 15 }, () => "1.000.000.000")];
    const vals = [Array.from({ length: 15 }, () => 1e9)];
    const d = checkDisplay(text, vals, addr, Array(15).fill(20));
    assert.strictEqual(d.narrow.length, NARROW_LIST_MAX);
    assert.strictEqual(d.narrowTotal, 15);
    assert.match(narrowNote(d), /15 ô.*, …\)/);
  });
});
