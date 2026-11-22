/** Đợt 12 (tasks/dot-12.md). Phần A: đo token đầu vào trước khi gửi, chọn model đủ sức chứa. */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { runTurn } from "../../src/core/agent/turn";
import { FreeRotation } from "../../src/core/agent/rotation";
import { HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { biggerModels, estimateInputTokens, fitsContext, sizeNote, tokenRatio, tooLargeMessage } from "../../src/core/inputsize";
import { pptTools, PPT_SYSTEM_PROMPT } from "../../src/office/pptTools";

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200, headers: { "content-type": "text/event-stream" } });
}
const oaText = (t: string) => sse([{ choices: [{ delta: { content: t } }] }, { choices: [{ delta: {}, finish_reason: "stop" }] }, "[DONE]"]);
const counting = () => {
  const calls: string[] = [];
  const fetchFn = (async (url: string) => (calls.push(String(url)), oaText("OK"))) as any;
  return { calls, fetchFn };
};
const host: HostToolSet = { host: "PowerPoint", specs: pptTools.specs, execute: async () => ({ ok: true }), overview: async () => ({}), systemPrompt: () => PPT_SYSTEM_PROMPT };
const memStore = () => {
  const m = new Map<string, string>();
  return { get: (k: string) => m.get(k) ?? null, set: (k: string, v: string) => void m.set(k, v) };
};

describe("Đợt 12 · A · đo token đầu vào", () => {
  it("Tính cả lời nhắc hệ thống, lệnh gọi công cụ và bộ công cụ (trước chỉ đếm tin nhắn); hai hệ số: JSON / lời nhắc 2,6 · chữ 3,6", () => {
    const msgs: any[] = [{ role: "system", content: "s".repeat(300) }, { role: "user", content: "u".repeat(300) }];
    const base = estimateInputTokens(msgs);
    assert.strictEqual(base, Math.ceil(300 / 2.6 + 300 / 3.6));
    const withTools = estimateInputTokens(msgs, pptTools.specs);
    const toolChars = pptTools.specs.reduce((n, t) => n + JSON.stringify({ name: t.name, description: t.description, parameters: t.parameters }).length, 0);
    assert.strictEqual(withTools, Math.ceil((300 + toolChars) / 2.6 + 300 / 3.6));
    // Kết quả công cụ (JSON) tính theo hệ số của JSON
    assert.strictEqual(estimateInputTokens([{ role: "tool", content: "j".repeat(260) } as any]), 100);
    assert.ok(withTools - base > 3000, `bộ công cụ PowerPoint phải đáng kể: ${withTools - base}`);
    const call: any = { role: "assistant", content: "", toolCalls: [{ id: "1", name: "x", args: { a: "b".repeat(60) } }] };
    assert.ok(estimateInputTokens([call]) > 20);
  });
  it("Hệ số theo họ model: Claude đếm chữ tiếng Việt dày hơn (đo 22/11: Haiku 5.5 đếm 23.378 token khi hệ số chung ước 13.689)", () => {
    for (const m of ["claude-haiku-5-5", "anthropic/claude-sonnet-5.5", "us.anthropic.claude-opus-5-5"]) assert.deepStrictEqual(tokenRatio(m), { struct: 2.1, text: 1.5 });
    for (const m of ["nvidia/nemotron-3-ultra-550b-a55b:free", "gemini-3.8-flash", "openai/gpt-oss-120b", undefined]) assert.deepStrictEqual(tokenRatio(m), { struct: 2.6, text: 3.6 });
    const msgs: any[] = [{ role: "user", content: "c".repeat(3000) }];
    assert.strictEqual(estimateInputTokens(msgs, [], "claude-haiku-5-5"), 2000);
    assert.strictEqual(estimateInputTokens(msgs, [], "gemini-3.8-flash"), Math.ceil(3000 / 3.6));
  });
  it("Sức chứa: chừa 10%; không biết sức chứa thì coi như chứa nổi", () => {
    assert.strictEqual(fitsContext(90000, 100000), true);
    assert.strictEqual(fitsContext(90001, 100000), false);
    assert.strictEqual(fitsContext(10 ** 7), true);
  });
  it("Gợi ý model cùng nguồn chứa nổi: nhỏ nhất trước, bỏ model đang chọn và model không rõ sức chứa", () => {
    const models = [
      { id: "a", context: 32000 },
      { id: "b", name: "B lớn", context: 1000000 },
      { id: "c", context: 200000 },
      { id: "d" },
      { id: "cur", context: 500000 },
    ];
    assert.deepStrictEqual(biggerModels(models, 100000, "cur").map((m) => m.id), ["c", "b"]);
    const t = tooLargeMessage({ tokens: 150000, context: 128000, model: "m-128k", alternatives: biggerModels(models, 150000, "cur") });
    assert.match(t, /^Chưa gửi: lượt này khoảng 150\.000 token/);
    assert.match(t, /m-128k \(128\.000 token\)/);
    assert.match(t, /c \(200\.000\), B lớn \(1\.000\.000\)/);
    assert.match(tooLargeMessage({ tokens: 9, context: 1, model: "x", alternatives: [] }), /không có model nào đủ sức chứa/);
    assert.strictEqual(sizeNote(25000, 128000), "Lượt này khoảng 25.000 token đầu vào — model chứa 128.000 token.");
  });

  it("Trả phí, model không chứa nổi: KHÔNG gọi nhà cung cấp, báo lý do too_large kèm gợi ý, lịch sử giữ nguyên", async () => {
    const f = counting();
    const ev: any[] = [];
    const history: any[] = [{ role: "user", content: "trước" }, { role: "assistant", content: "đã trả lời" }];
    const r = await runTurn({
      provider: getProvider("groq")!,
      key: "K",
      model: "nho",
      context: 8000,
      alternatives: [{ id: "nho", context: 8000 }, { id: "lon", name: "Lớn", context: 131072 }],
      host: "PowerPoint",
      history,
      text: "Tóm tắt tệp đính kèm\n" + "chữ ".repeat(8000),
      onEvent: (e) => ev.push(e),
      fetchFn: f.fetchFn,
      agent: { tools: host, log: [] },
    });
    assert.strictEqual(f.calls.length, 0);
    assert.strictEqual((r.error as any)?.reason, "too_large");
    assert.match(String((r.error as any)?.message), /^Chưa gửi.*Lớn \(131\.072\)/);
    assert.strictEqual(r.history, history);
    const sz = ev.find((e) => e.type === "input_size");
    assert.ok(sz && sz.tokens > 8000 && sz.context === 8000);
  });
  it("Trả phí, model chứa nổi: gửi bình thường, vẫn báo số đo", async () => {
    const f = counting();
    const ev: any[] = [];
    const r = await runTurn({ provider: getProvider("groq")!, key: "K", model: "lon", context: 131072, host: "PowerPoint", history: [], text: "Xin chào", onEvent: (e) => ev.push(e), fetchFn: f.fetchFn, agent: { tools: host, log: [] } });
    assert.strictEqual(r.error, undefined);
    assert.strictEqual(f.calls.length, 1);
    assert.ok(ev.some((e) => e.type === "input_size" && e.tokens > 3000));
  });
  it("Trò chuyện thường (không tác tử) cũng đo và chặn khi không chứa nổi", async () => {
    const f = counting();
    const r = await runTurn({ provider: getProvider("groq")!, key: "K", model: "nho", context: 1000, host: "Word", history: [], text: "x".repeat(6000), onEvent: () => {}, fetchFn: f.fetchFn });
    assert.strictEqual(f.calls.length, 0);
    assert.strictEqual((r.error as any)?.reason, "too_large");
  });
  it("Miễn phí: không chặn ở lượt mà để vòng xoay bỏ nguồn; vòng xoay tính cả bộ công cụ", async () => {
    // Tin nhắn nhỏ nhưng bộ công cụ lớn: nguồn sức chứa 3.000 token phải bị bỏ qua, không gọi thử
    const calls: string[] = [];
    const fetchFn = (async (url: string) => (calls.push(String(url)), oaText("vừa"))) as any;
    const P = (id: string) => getProvider(id)!;
    const rot = new FreeRotation(
      [
        { provider: P("groq"), key: "K1", model: "m", context: 3000 },
        { provider: P("openrouter"), key: "K2", model: "m", context: 200000 },
      ],
      { store: memStore() as any },
    );
    const res = await rot.call({ messages: [{ role: "user", content: "Xin chào" }], tools: pptTools.specs as ToolSpec[], onText: () => {}, fetchFn });
    assert.strictEqual(res.text, "vừa");
    assert.ok(calls.length === 1 && /openrouter/.test(calls[0]), calls.join(" "));
  });
  it("Miễn phí, mọi nguồn quá nhỏ: câu lỗi nói cách làm nhỏ lượt, không tự sang trả phí", async () => {
    const fetchFn = (async () => oaText("không được gọi")) as any;
    const rot = new FreeRotation([{ provider: getProvider("groq")!, key: "K", model: "m", context: 1000 }], { store: memStore() as any });
    await assert.rejects(rot.call({ messages: [{ role: "user", content: "x".repeat(9000) }], tools: [], onText: () => {}, fetchFn }), (e: any) => {
      assert.match(e.message, /quá lớn cho mọi nguồn miễn phí: bớt tệp đính kèm/);
      assert.match(e.message, /không tự chuyển sang nguồn trả phí/);
      return true;
    });
  });
});
