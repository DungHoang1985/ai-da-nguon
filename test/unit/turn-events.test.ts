/** Khối sự kiện của lượt (core/turnEvents.ts) — tách khỏi Chat.tsx 27/11/2026, phần để sau của nhóm 2 */
import * as assert from "assert";
import { TurnEffect, TurnEventTracker, TurnItem } from "../../src/core/turnEvents";

const WRITE = new Set(["word_replace_text"]);
const mk = (rotation = false) =>
  new TurnEventTracker({ rotation: () => rotation, isWrite: (n) => WRITE.has(n), describe: (c) => `bước ${c.name}` });

/** Áp việc như khung chat: patch lên mục, nối chữ vào khối chữ cuối, ghi thông báo */
function apply(it: TurnItem, fxs: TurnEffect[], log: string[] = []) {
  for (const fx of fxs) {
    if (fx.kind === "patch") it = fx.fn(it);
    else if (fx.kind === "text") {
      const segs = (it.segments || []).slice();
      const last = segs[segs.length - 1];
      if (last && last.type === "text") segs[segs.length - 1] = { ...last, text: last.text + fx.text };
      else segs.push({ type: "text", text: fx.text.replace(/^\n+/, "") });
      it = { ...it, segments: segs };
    } else log.push(fx.kind === "notice" ? "notice:" + fx.text : "clear");
  }
  return it;
}

describe("Khối sự kiện của lượt (TurnEventTracker)", () => {
  it("chữ của bước sau được ngăn bằng xuống dòng; bước đầu thì không", () => {
    const t = mk();
    let it: TurnItem = { text: "" };
    for (const ev of [
      { type: "step", step: 1 },
      { type: "text", delta: "Một" },
      { type: "step", step: 2 },
      { type: "text", delta: "Hai" },
      { type: "text", delta: " tiếp" },
    ])
      it = apply(it, t.handle(ev));
    assert.strictEqual(it.segments![0].text, "Một\nHai tiếp");
    assert.strictEqual(it.calls, 2);
  });

  it("thẻ bước: đang chạy → ok / lỗi; đếm lệnh ghi thành công, thất bại và giữ tham số lệnh ghi", () => {
    const t = mk();
    let it: TurnItem = { text: "" };
    const call = { id: "c1", name: "word_replace_text", args: { find: "12", replace: "6" } };
    it = apply(it, t.handle({ type: "tool_start", call }));
    assert.strictEqual(it.segments![0].status, "running");
    it = apply(it, t.handle({ type: "tool_end", call, result: { ok: true, message: "Đã thay 1 chỗ", brief: "1 chỗ" } }));
    assert.deepStrictEqual([it.segments![0].status, it.segments![0].brief, it.wrote], ["ok", "1 chỗ", true]);
    const bad = { id: "c2", name: "word_replace_text", args: {} };
    it = apply(it, t.handle({ type: "tool_start", call: bad }));
    it = apply(it, t.handle({ type: "tool_end", call: bad, error: "Không thấy chữ" }));
    assert.strictEqual(it.segments![1].status, "error");
    assert.deepStrictEqual(t.called, ["word_replace_text", "word_replace_text"]);
    assert.strictEqual(t.failedWrites, 1);
    assert.deepStrictEqual(t.writtenArgs, [JSON.stringify(call.args)]);
    // Lệnh đọc không tính là ghi
    apply(it, t.handle({ type: "tool_end", call: { id: "c3", name: "word_read" }, result: { ok: true } }));
    assert.strictEqual(t.failedWrites, 1);
  });

  it("thử lại: hiện thông báo, rồi tự bỏ đúng một lần khi có chữ / bước / số token mới", () => {
    const t = mk();
    const log: string[] = [];
    let it: TurnItem = { text: "" };
    it = apply(it, t.handle({ type: "retry", reason: "máy chủ bận", attempt: 1, max: 3, waitMs: 2000 }), log);
    assert.match(log[0], /^notice:Máy chủ bận: đang thử lại lần 1\/3 sau 2 giây/);
    it = apply(it, t.handle({ type: "text", delta: "x" }), log);
    it = apply(it, t.handle({ type: "text", delta: "y" }), log);
    assert.deepStrictEqual(log.slice(1), ["clear"]);
    assert.strictEqual(t.retrying, false);
  });

  it("thứ tự việc: bỏ thông báo thử lại TRƯỚC khi sửa mục (tool_start)", () => {
    const t = mk();
    t.handle({ type: "retry", reason: "x", attempt: 1, max: 2, waitMs: 0 });
    assert.deepStrictEqual(
      t.handle({ type: "tool_start", call: { id: "a", name: "word_read" } }).map((f) => f.kind),
      ["clear_retry", "patch"]
    );
  });

  it("cổng cuối lượt ghi vị trí chữ lúc nhắc; mục giữ nguyên", () => {
    const t = mk();
    let it: TurnItem = { text: "", segments: [{ type: "text", text: "Đã sửa xong" }] };
    const before = it;
    it = apply(it, t.handle({ type: "error_gate" }));
    assert.strictEqual(it, before);
    assert.strictEqual(t.gateCut, "Đã sửa xong".length);
  });

  it("đổi nguồn: ghi model trả lời và nối dòng báo; cắt chữ: báo đúng câu", () => {
    const t = mk();
    let it: TurnItem = { text: "" };
    it = apply(it, t.handle({ type: "switch", from: "A", to: "B", reason: "hết lượt", partial: true }));
    assert.strictEqual(it.answeredBy, "B");
    assert.match(it.segments![0].text, /↪ A: hết lượt — đổi sang B \(trả lời lại bước này\)\./);
    assert.match(t.handle({ type: "truncated" })[0].kind === "text" ? (t.handle({ type: "truncated" })[0] as any).text : "", /bị cắt vì quá dài/);
    assert.strictEqual(t.handle({ type: "truncated", skipped: true }).length, 0);
    assert.match((t.handle({ type: "truncated", skipped: true, gaveUp: true })[0] as any).text, /Dừng lượt/);
  });

  it("cỡ lượt lớn: có báo; lượt xoay vòng Miễn phí không kèm cửa sổ ngữ cảnh của một model", () => {
    const big = { type: "input_size", tokens: 10_000_000, context: 200_000 };
    const a = mk(false).handle(big).filter((f) => f.kind === "notice") as any[];
    const b = mk(true).handle(big).filter((f) => f.kind === "notice") as any[];
    assert.strictEqual(a.length, 1);
    assert.strictEqual(b.length, 1);
    assert.notStrictEqual(a[0].text, b[0].text);
    assert.deepStrictEqual(mk().handle({ type: "input_size", tokens: 10 }).map((f) => f.kind), ["patch"]);
  });

  it("số token cộng dồn; firstInput chỉ lấy ở bước 1", () => {
    const t = mk();
    let it: TurnItem = { text: "" };
    it = apply(it, t.handle({ type: "usage", step: 1, usage: { input: 100, output: 10 } }));
    it = apply(it, t.handle({ type: "usage", step: 2, usage: { input: 50, output: 5 } }));
    assert.strictEqual(it.firstInput, 100);
    assert.strictEqual((it.usage as any).input, 150);
  });

  it("dừng vì chạm trần / kẹt lỗi: ghi thẻ Làm tiếp; sự kiện lạ bỏ qua", () => {
    const t = mk();
    let it: TurnItem = { text: "" };
    it = apply(it, t.handle({ type: "stuck", step: 5, writes: 2, tool: "word_find", times: 3 }));
    assert.deepStrictEqual(it.paused, { steps: 5, writes: 2, stuck: "word_find báo cùng một lỗi 3 bước liên tiếp" });
    assert.deepStrictEqual(t.handle({ type: "khong_co" }), []);
    assert.deepStrictEqual(t.handle(undefined), []);
  });
});
