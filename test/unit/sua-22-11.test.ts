/** Hai lỗi tìm ra ở buổi so sánh chế độ Miễn phí với add-in Claude chính thức (22/11/2026):
 *  1. ppt_add_table không có cỡ chữ: bảng 7 × 4 của bài P3 tràn slide, model xóa dựng lại 7 lần (P3-D1).
 *  2. Cảnh báo "AI nói đã sửa mà không ghi" báo nhầm ở lượt sau "Làm tiếp" chỉ đọc (X6-D3, W4-B3). */
import * as assert from "assert";
import { TABLE_MIN_FONT, nextTableFont, tableColumnWidths, tableFits } from "../../src/core/tablefit";
import { mergeTaskFacts, taskSoFar, unverifiedWriteClaim } from "../../src/core/claims";
import { PPT_SPECS } from "../../src/office/pptTools";

const P3 = [
  ["Thủ tục", "Cơ quan", "Trạng thái", "Hạn"],
  ["Giấy phép xây dựng khối B", "Sở Xây dựng", "Đang thẩm định", "30/12/2026"],
  ["Giấy phép xây dựng khối C", "Sở Xây dựng", "Chưa nộp", "31/03/2027"],
  ["Thẩm duyệt phòng cháy", "Cảnh sát PCCC", "Bổ sung hồ sơ", "15/12/2026"],
  ["Bán nhà ở hình thành trong tương lai", "Sở Xây dựng", "Chuẩn bị hồ sơ", "28/02/2027"],
  ["Đánh giá tác động môi trường giai đoạn 2", "Sở Tài nguyên và Môi trường", "Đang thẩm định", "20/01/2027"],
  ["Bàn giao mặt bằng đợt 2", "UBND phường", "Vướng 6 hộ", "31/12/2026"],
];

describe("22/11 · ppt_add_table vừa slide", () => {
  it("Khai báo có font_size và mô tả nói công cụ tự co, báo khung thật", () => {
    const s = PPT_SPECS.find((x) => x.name === "ppt_add_table")!;
    assert.ok(s.parameters.properties.font_size);
    assert.match(s.description, /fit/i);
  });
  it("Cỡ chữ thử giảm dần, dừng ở 10", () => {
    const seq: number[] = [];
    for (let s = nextTableFont(18); s !== null; s = nextTableFont(s)) seq.push(s);
    assert.deepStrictEqual(seq, [16, 14, 13, 12, 11, 10]);
    assert.strictEqual(nextTableFont(TABLE_MIN_FONT), null);
    assert.strictEqual(nextTableFont(15), 14);
  });
  it("Lọt slide: đáy bảng cách đáy slide ≥ 8 pt (số đo thật của P3: 504 pt tràn, 266 pt lọt)", () => {
    assert.strictEqual(tableFits(100, 504, 405), false);
    assert.strictEqual(tableFits(100, 266, 405), true);
    assert.strictEqual(tableFits(100, 300, 405), false, "đáy 400 pt, chỉ chừa 5 pt");
  });
  it("Bề ngang cột theo độ dài chữ: tổng đúng, cột “Hạn” hẹp hơn cột “Thủ tục”, không cột nào dưới 8%", () => {
    const w = tableColumnWidths(P3, 640);
    assert.strictEqual(w.length, 4);
    assert.ok(Math.abs(w.reduce((a, b) => a + b, 0) - 640) < 0.01, String(w));
    assert.ok(w[0] > w[3], String(w));
    assert.ok(w.every((x) => x >= 640 * 0.08 - 0.01), String(w));
    // Cột "Hạn" đủ chỗ cho "30/12/2026" ở cỡ 12 (khoảng 6 pt mỗi ký tự + lề ô 14 pt)
    assert.ok(w[3] >= 10 * 6 + 14, String(w));
  });
  it("Bảng một cột, cột rỗng: không lỗi, tổng đúng", () => {
    assert.deepStrictEqual(tableColumnWidths([["A"], ["B"]], 300), [300]);
    const w = tableColumnWidths([["Mục", ""], ["A", ""]], 400);
    assert.ok(Math.abs(w[0] + w[1] - 400) < 0.01 && w[1] >= 32, String(w));
    assert.deepStrictEqual(tableColumnWidths([], 300), []);
  });
});

describe("22/11 · Cảnh báo 'nói đã sửa mà không ghi' gộp cả nhiệm vụ qua Làm tiếp", () => {
  const ask =
    "Dựa trên sheet Giả định, lập sheet mới \"Dòng tiền\" theo quý Q1–Q8 gồm: số căn bán, doanh thu, chi phí đất, chi phí xây dựng, chi phí bán hàng, dòng tiền thuần. Tính NPV bằng hàm NPV của Excel.";
  // Câu trả lời thật của lượt sau Làm tiếp ở X6-D3 (lượt đó chỉ đọc lại Dòng tiền!A1:I13)
  const answer =
    "Đã hoàn tất tạo sheet \"Dòng tiền\" theo quý Q1–Q8.\nKiểm tra: `excel_find_errors` — không lỗi công thức. Đã định dạng số `#,##0.00`, tỷ lệ `%`, in đậm tiêu đề, viền, căn lề.";
  const first = { userText: ask, answer: "(chạm trần số bước)", wrote: true, toolsCalled: ["excel_read_range", "excel_add_sheet", "excel_write_range", "excel_format_range"], failedWrites: 0 };
  const cont = { userText: ask, answer, wrote: false, toolsCalled: ["excel_read_range"], failedWrites: 0 };

  it("Ca báo nhầm X6-D3: lượt trước của cùng nhiệm vụ đã ghi → không cảnh báo", () => {
    // Bằng chứng cách cũ báo nhầm: chỉ xét lượt Làm tiếp
    assert.ok(unverifiedWriteClaim(cont));
    assert.strictEqual(unverifiedWriteClaim(mergeTaskFacts(taskSoFar(first), cont)), null);
  });
  it("Ca báo nhầm W4-B3 (Word, câu có chủ ngữ “tôi”): lượt trước đã ghi → không cảnh báo", () => {
    const prev = taskSoFar({ userText: "Soạn công văn đề nghị gia hạn", answer: "", wrote: true, toolsCalled: ["word_build_admin_doc"] });
    const cur = { userText: "", answer: "Tôi đã soạn xong công văn và đã căn lề theo thể thức.", wrote: false, toolsCalled: ["word_read_document"] };
    assert.strictEqual(unverifiedWriteClaim(mergeTaskFacts(prev, cur)), null);
  });
  it("Ca phải bắt: cả nhiệm vụ không có lệnh ghi nào thành công → vẫn cảnh báo, nói rõ tính cả lượt trước", () => {
    const prev = taskSoFar({ ...first, wrote: false, toolsCalled: ["excel_read_range"] });
    const w = unverifiedWriteClaim(mergeTaskFacts(prev, cont));
    assert.ok(w);
    assert.match(w!.text, /cả nhiệm vụ \(tính cả các lượt trước khi bấm Làm tiếp\)/);
    assert.match(w!.text, /chỉ gọi công cụ đọc/);
  });
  it("Ca phải bắt: lệnh ghi của các lượt đều hỏng → cộng số lệnh hỏng của cả nhiệm vụ", () => {
    const prev = taskSoFar({ ...first, wrote: false, failedWrites: 2 });
    const w = unverifiedWriteClaim(mergeTaskFacts(prev, { ...cont, failedWrites: 1 }));
    assert.ok(w && /3 thao tác ghi của cả nhiệm vụ/.test(w.text), w?.text);
  });
  it("Lượt Làm tiếp dùng yêu cầu gốc để xét ý nhờ sửa (câu “Đã …” không chủ ngữ chỉ tính khi người dùng nhờ sửa)", () => {
    const prev = taskSoFar({ userText: "Bảng này có bao nhiêu dòng?", answer: "", wrote: false, toolsCalled: ["excel_read_range"] });
    const m = mergeTaskFacts(prev, { userText: "Làm tiếp nhiệm vụ, không làm lại các bước", answer: "Đã định dạng tiêu đề.", wrote: false, toolsCalled: [] });
    assert.strictEqual(m.userText, "Bảng này có bao nhiêu dòng?");
    assert.strictEqual(unverifiedWriteClaim(m), null);
  });
  it("Nhiệm vụ mới (không có lượt trước): như cũ, câu cảnh báo nói “lượt này”", () => {
    const cur = { userText: "Sửa mức phạt ở Điều 5", answer: "Đã sửa mức phạt tại Điều 5.", wrote: false, toolsCalled: [] };
    assert.deepStrictEqual(mergeTaskFacts(undefined, cur), cur);
    assert.match(unverifiedWriteClaim(cur)!.text, /lượt này AI không gọi công cụ nào/);
  });
  it("Bản lưu nhiệm vụ cộng dồn qua nhiều lần Làm tiếp", () => {
    const t1 = taskSoFar({ ...first, wrote: false, toolsCalled: ["a"] });
    const t2 = taskSoFar(mergeTaskFacts(t1, { ...cont, wrote: true, toolsCalled: ["b"] }));
    const t3 = mergeTaskFacts(t2, { ...cont, toolsCalled: ["c"] });
    assert.strictEqual(t3.wrote, true);
    assert.deepStrictEqual(t3.toolsCalled, ["a", "b", "c"]);
    assert.strictEqual(unverifiedWriteClaim(t3), null);
  });
});
