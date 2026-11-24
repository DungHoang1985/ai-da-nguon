/** Bản vá 4.11.13 — nhóm "an toàn dữ liệu" của rà soát kiến trúc 20/11/2026: B1, B3 (Excel), B4, B23 (PowerPoint), B6 (hàng đợi lệnh ghi).
 *  Phần cần Office (câu hỏi thật trong hộp xác nhận, xóa slide) kiểm bằng Tự kiểm tra; ở đây kiểm các hàm thuần và cách nối. */
import * as assert from "assert";
import { LONG_CELL_TEXT, LONG_SHAPE_TEXT, MAX_UNDO_CELLS, longOverwrite, noUndoSuffix, OVERWRITE_UNCHECKED, undoNote } from "../../src/office/util";

describe("B1: câu hỏi Excel nói đúng việc có Hoàn tác hay không", () => {
  it("tới 10.000 ô: có thể Hoàn tác; trên 10.000 ô: KHÔNG hoàn tác được, kèm số ô", () => {
    assert.strictEqual(MAX_UNDO_CELLS, 10000);
    assert.match(undoNote(10000), /có thể Hoàn tác/);
    const big = undoNote(250000);
    assert.match(big, /KHÔNG hoàn tác được/);
    assert.match(big, /250\.000 ô/);
    assert.doesNotMatch(big, /có thể Hoàn tác/);
  });

  it("kết quả báo khi không có bản chụp; có bản chụp thì không thêm gì", () => {
    assert.strictEqual(noUndoSuffix(true), "");
    assert.match(noUndoSuffix(false), /KHÔNG có Hoàn tác/);
  });

  it("B3: câu khi không kiểm được vùng đích nói rõ có thể ghi đè", () => {
    assert.match(OVERWRITE_UNCHECKED, /Không kiểm được/);
    assert.match(OVERWRITE_UNCHECKED, /ghi đè/);
  });

  it("sắp xếp được xếp vào nhóm hỏi trước (vùng lớn thì hỏi), xóa vùng vẫn hỏi", () => {
    const { excelTools } = require("../../src/office/excelTools");
    const risky = (n: string) => typeof excelTools.specs.find((s: any) => s.name === n)?.risky === "function";
    assert.ok(risky("excel_sort_range"));
    assert.ok(risky("excel_clear_range"));
    assert.ok(risky("excel_write_range"));
  });
});

describe("B23: PowerPoint chỉ hỏi khi thay chữ dài đang có (anh chọn 25/11/2026)", () => {
  const long = "Nội dung dài ".repeat(30); // ~390 ký tự
  it("thay chữ ngắn (tiêu đề, ô vài chữ): không hỏi", () => {
    assert.strictEqual(longOverwrite([{ where: "hình Title", old: "Tiêu đề cũ", next: "Tiêu đề mới", limit: LONG_SHAPE_TEXT }]), null);
    assert.strictEqual(longOverwrite([{ where: "ô", old: "", next: "mới", limit: LONG_CELL_TEXT }]), null);
  });

  it("thay hoặc xóa chữ dài: hỏi, nêu chỗ, số ký tự và mấy chữ đầu", () => {
    const m = longOverwrite([{ where: 'hình "Nội dung" slide 2', old: long, next: "Ngắn gọn", limit: LONG_SHAPE_TEXT }]);
    assert.ok(m);
    assert.match(m!, /hình "Nội dung" slide 2/);
    assert.match(m!, /\d+ ký tự/);
    assert.match(m!, /Nội dung dài/);
    assert.ok(longOverwrite([{ where: "x", old: long, next: "", limit: LONG_SHAPE_TEXT }]), "xóa trắng chữ dài cũng hỏi");
  });

  it("chỉ thêm vào sau chữ cũ hoặc ghi lại y nguyên: không hỏi", () => {
    assert.strictEqual(longOverwrite([{ where: "x", old: long, next: long + " Thêm một câu.", limit: LONG_SHAPE_TEXT }]), null);
    assert.strictEqual(longOverwrite([{ where: "x", old: long, next: long, limit: LONG_SHAPE_TEXT }]), null);
  });

  it("ô bảng dùng ngưỡng riêng (ngắn hơn hình); nhiều ô thì gộp một câu", () => {
    const cell = "x".repeat(LONG_CELL_TEXT + 1);
    const m = longOverwrite([
      { where: "ô (0,1)", old: cell, next: "a", limit: LONG_CELL_TEXT },
      { where: "ô (0,2)", old: cell, next: "b", limit: LONG_CELL_TEXT },
      { where: "ô (0,3)", old: "ngắn", next: "c", limit: LONG_CELL_TEXT },
    ]);
    assert.match(m!, /ô \(0,1\)/);
    assert.match(m!, /và 1 chỗ khác/);
    assert.strictEqual(longOverwrite([{ where: "hình", old: cell, next: "a", limit: LONG_SHAPE_TEXT }]), null, "101 ký tự trong hình chưa là chữ dài");
  });

  it("ghi chữ vào hình và sửa bảng được nối vào bước hỏi; xóa slide / hình vẫn hỏi", () => {
    const { pptTools } = require("../../src/office/pptTools");
    const risky = (n: string) => typeof pptTools.specs.find((s: any) => s.name === n)?.risky === "function";
    ["ppt_set_shape_text", "ppt_edit_table", "ppt_delete_slide", "ppt_delete_shape"].forEach((n) => assert.ok(risky(n), n));
  });
});

describe("B4: PowerPoint không xóa khi chưa chụp được bản Hoàn tác", () => {
  it("mã xóa slide / hình dừng trước khi xóa nếu không xuất được slide", () => {
    const src = require("fs").readFileSync(require("path").join(__dirname, "../../src/office/pptAdvanced.ts"), "utf8");
    const guards = src.match(/if \(!b64\) throw new Error\(NO_SNAPSHOT_DELETE\);\s*(?:const nm = shp\.name;\s*)?(?:slide|shp)\.delete\(\)/g) || [];
    assert.strictEqual(guards.length, 2, "cả ppt_delete_slide và ppt_delete_shape phải chặn trước lệnh delete()");
    assert.doesNotMatch(src, /không hoàn tác được\)" \}`/);
  });
});

describe("B6: lệnh ghi Excel, PowerPoint chạy lần lượt", () => {
  const { withWriteQueue } = require("../../src/office/tools");
  const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
  function fakeSet(order: string[]) {
    return {
      host: "Thu",
      specs: [{ name: "ghi", write: true }, { name: "doc" }],
      async execute(name: string, a: any) {
        order.push(`${name}${a.n}:bắt đầu`);
        if (a.fail) throw new Error("hỏng");
        await sleep(a.ms || 0);
        order.push(`${name}${a.n}:xong`);
        return a.n;
      },
      overview: async () => ({}),
      systemPrompt: () => "",
    };
  }

  it("lệnh ghi thứ hai chỉ bắt đầu khi lệnh thứ nhất xong; lệnh đọc không phải chờ", async () => {
    const order: string[] = [];
    const set = withWriteQueue(fakeSet(order) as any, "thu-1");
    const p1 = set.execute("ghi", { n: 1, ms: 30 }, { log: [] });
    const p2 = set.execute("ghi", { n: 2, ms: 0 }, { log: [] });
    const p3 = set.execute("doc", { n: 3, ms: 0 }, { log: [] });
    assert.deepStrictEqual(await Promise.all([p1, p2, p3]), [1, 2, 3]);
    assert.ok(order.indexOf("ghi2:bắt đầu") > order.indexOf("ghi1:xong"), order.join(" "));
    assert.ok(order.indexOf("doc3:xong") < order.indexOf("ghi1:xong"), "lệnh đọc chạy ngay");
  });

  it("lệnh ghi lỗi không chặn hàng đợi; hàng đợi dùng chung giữa hai lần lấy bộ công cụ cùng ứng dụng", async () => {
    const order: string[] = [];
    const a = withWriteQueue(fakeSet(order) as any, "thu-2");
    const b = withWriteQueue(fakeSet(order) as any, "thu-2");
    const p1 = a.execute("ghi", { n: 1, ms: 20 }, { log: [] });
    const p2 = b.execute("ghi", { n: 2, fail: true }, { log: [] });
    const p3 = a.execute("ghi", { n: 3 }, { log: [] });
    await p1;
    await assert.rejects(p2, /hỏng/);
    assert.strictEqual(await p3, 3);
    assert.ok(order.indexOf("ghi2:bắt đầu") > order.indexOf("ghi1:xong"));
  });
});
