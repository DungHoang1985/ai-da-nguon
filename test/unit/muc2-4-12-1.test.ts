/** Mức 2, bản 4.12.1: khối "đoạn người dùng tự sửa giữa hai lượt" (src/core/userChanges.ts) và móc chụp cuối lượt (afterTurn). */
import * as assert from "assert";
import { CLIP, MAX_LISTED, SnapRow, diffUserChanges, snapshotOf } from "../../src/core/userChanges";
import { getProvider } from "../../src/core/providers";
import { runTurn } from "../../src/core/agent/turn";
import { HostToolSet } from "../../src/core/agent/types";
import { overviewForPrompt } from "../../src/office/util";

const rows = (xs: [string, string][], deleted: string[] = []): SnapRow[] => xs.map(([id, text], i) => ({ id, i, text, deleted: deleted.includes(id) }));

describe("4.12.1 · Đoạn người dùng tự sửa giữa hai lượt", () => {
  const before = rows([
    ["a", "Điều 1. Phạm vi"],
    ["b", "Thời hạn bảo hành là 12 tháng."],
    ["c", "Điều 2. Thanh toán"],
    ["d", "Bên B thanh toán trong 30 ngày."],
  ]);
  const snap = snapshotOf(before);
  it("Chưa có bản chụp (lượt đầu) hoặc không đổi gì → không báo", () => {
    assert.strictEqual(diffUserChanges(null, before), null);
    assert.strictEqual(diffUserChanges(snap, before), null);
  });
  it("Sửa một đoạn → báo đúng chỉ số hiện tại, chữ cũ và chữ mới", () => {
    const now = rows([
      ["a", "Điều 1. Phạm vi"],
      ["b", "Thời hạn bảo hành là 24 tháng."],
      ["c", "Điều 2. Thanh toán"],
      ["d", "Bên B thanh toán trong 30 ngày."],
    ]);
    const d = diffUserChanges(snap, now)!;
    assert.deepStrictEqual(d.changed, [{ index: 1, before: "Thời hạn bảo hành là 12 tháng.", now: "Thời hạn bảo hành là 24 tháng." }]);
    assert.deepStrictEqual(d.inserted, []);
    assert.strictEqual(d.removed, 0);
    assert.match(d.note, /không ghi đè/);
  });
  it("Thêm đoạn có chữ, xóa đoạn → báo; thêm đoạn trống (Enter thừa) → không báo; chỉ số dời theo văn bản hiện tại", () => {
    const now = rows([
      ["a", "Điều 1. Phạm vi"],
      ["x", "Khoản mới người dùng gõ."],
      ["y", "   "],
      ["c", "Điều 2. Thanh toán"],
      ["d", "Bên B thanh toán trong 30 ngày."],
    ]);
    const d = diffUserChanges(snap, now)!;
    assert.deepStrictEqual(d.inserted, [{ index: 1, text: "Khoản mới người dùng gõ." }]);
    assert.strictEqual(d.removed, 1);
    assert.strictEqual(d.changed.length, 0);
  });
  it("Đoạn bị xóa có theo dõi (còn chờ duyệt) tính là đã xóa; đoạn không có định danh bỏ qua", () => {
    const now = rows(
      [
        ["a", "Điều 1. Phạm vi"],
        ["b", "Thời hạn bảo hành là 12 tháng."],
        ["c", "Điều 2. Thanh toán"],
        ["d", "Bên B thanh toán trong 30 ngày."],
      ],
      ["d"]
    );
    now.push({ i: 4, text: "Đoạn của Word cũ không có định danh" });
    const d = diffUserChanges(snap, now)!;
    assert.strictEqual(d.removed, 1);
    assert.strictEqual(d.inserted.length, 0);
  });
  it("Nhiều chỗ: liệt kê tối đa MAX_LISTED, cắt chữ dài, báo số còn lại", () => {
    const many: [string, string][] = Array.from({ length: 20 }, (_, k) => [`p${k}`, `Đoạn ${k}`]);
    const s = snapshotOf(rows(many));
    const now = rows(many.map(([id, t]) => [id, t + " sửa " + "x".repeat(300)]));
    const d = diffUserChanges(s, now)!;
    assert.strictEqual(d.changed.length, MAX_LISTED);
    assert.strictEqual(d.more, 20 - MAX_LISTED);
    assert.ok(d.changed[0].now.length <= CLIP + 1 && d.changed[0].now.endsWith("…"));
  });
  it("Khối này giữ nguyên khi phần trạng thái gửi model bị cắt cho vừa trần", () => {
    const d = diffUserChanges(snap, rows([["a", "Điều 1. Phạm vi (sửa)"], ["b", "Thời hạn bảo hành là 12 tháng."], ["c", "Điều 2. Thanh toán"], ["d", "Bên B thanh toán trong 30 ngày."]]))!;
    const ov = { paragraphs: 400, headings: Array.from({ length: 300 }, (_, k) => ({ index: k, text: "Điều " + k + " ".repeat(40) })), first_paragraphs: [], user_changes: d };
    const s = overviewForPrompt(ov);
    assert.ok(s.length <= 6000);
    assert.match(s, /"user_changes":\{"note"/);
    assert.match(s, /Điều 1\. Phạm vi \(sửa\)/);
  });
});

describe("4.12.1 · Móc chụp cuối lượt (afterTurn)", () => {
  const P = (id: string) => getProvider(id)!;
  const say = (text: string) => new Response(`data: ${JSON.stringify({ choices: [{ delta: { content: text } }] })}\n\n`, { status: 200 });
  const base = (afterTurn: () => Promise<void>): HostToolSet => ({
    host: "Word",
    specs: [],
    overview: async () => ({}),
    systemPrompt: () => "S",
    execute: async () => ({ ok: true }),
    afterTurn,
  });
  const opts = (tools: HostToolSet, fetchFn: any) =>
    ({ provider: P("groq"), key: "K", model: "m", host: "Word", text: "x", history: [], onEvent: () => {}, fetchFn, retryDelaysMs: [], agent: { tools, log: [] } }) as any;
  it("Lượt xong → gọi afterTurn đúng một lần", async () => {
    let n = 0;
    await runTurn(opts(base(async () => void n++), (async () => say("Chào.")) as any));
    assert.strictEqual(n, 1);
  });
  it("Lượt lỗi khi gọi nhà cung cấp → vẫn gọi afterTurn; afterTurn lỗi không làm hỏng lượt", async () => {
    let n = 0;
    const r1 = await runTurn(opts(base(async () => void n++), (async () => new Response("x", { status: 500 })) as any));
    assert.strictEqual(n, 1);
    assert.ok(r1.error);
    const r2 = await runTurn(
      opts(
        base(async () => {
          throw new Error("Word bận");
        }),
        (async () => say("Chào.")) as any
      )
    );
    assert.ok(!r2.error, String(r2.error));
  });
});
