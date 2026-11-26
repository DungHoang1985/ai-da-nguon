/** B7 phần 2 (nhóm 2 rà soát kiến trúc, bước 4b, 26/11/2026): chuẩn bị vòng xoay Miễn phí — dời từ Chat.tsx, kiểm bằng nguồn giả.
 *  Chạy thật vòng xoay cần nguồn miễn phí chạy được (OpenRouter, Cerebras, Cohere còn key từ 26/11). */
import * as assert from "assert";
import { prepareFreeRotation } from "../../src/core/agent/rotation";
import { getProvider } from "../../src/core/providers";

const OR = getProvider("openrouter")!;
const CO = getProvider("cohere")!;
const FREE_OR = { id: "nvidia/nemotron-3-ultra-550b-a55b:free", name: "Nemotron free", free: true, tools: true };
const CO_M = { id: "command-a-03-2025", name: "Command A", tools: true };

function setup(o: { cache?: any; fetchModels?: any; keys?: Record<string, string>; waitMs?: number; abort?: boolean; catalogFail?: Record<string, number> } = {}) {
  const cache: Record<string, any> = o.cache ?? { openrouter: { at: Date.now(), models: [FREE_OR] }, cohere: { at: Date.now(), models: [CO_M] } };
  const ac = new AbortController();
  if (o.abort) ac.abort();
  const fetched: string[] = [];
  const catalogFail = o.catalogFail ?? {};
  const run = () =>
    prepareFreeRotation({
      selected: { provider: OR, modelId: FREE_OR.id },
      keys: o.keys ?? { openrouter: "k1", cohere: "k2" },
      order: ["openrouter", "gemini", "groq", "cohere"],
      getProvider,
      modelCache: () => cache,
      setModels: (pid, ms) => (cache[pid] = { at: Date.now(), models: ms }),
      fetchModels: o.fetchModels ?? (async (p: any) => (fetched.push(p.id), [CO_M])),
      isStale: (at) => !at || Date.now() - at > 86400000,
      catalogFail,
      remembered: () => "",
      agentOn: true,
      deepThink: false,
      cooldownStore: { get: () => ({}), set: () => {} },
      signal: ac.signal,
      waitMs: o.waitMs ?? 8000,
    });
  return { run, fetched, cache, catalogFail };
}

describe("B7 · prepareFreeRotation (chuẩn bị vòng xoay Miễn phí)", () => {
  it("danh mục đủ và còn mới: không tải lại; nguồn đang chọn đứng đầu, nguồn có key khác theo thứ tự", async () => {
    const t = setup();
    const r: any = await t.run();
    assert.strictEqual(r.kind, "ok");
    assert.deepStrictEqual(t.fetched, []);
    assert.deepStrictEqual(r.candidates.map((c: any) => c.provider.id), ["openrouter", "cohere"]);
  });

  it("nguồn khác chưa có danh mục: tải trước rồi mới dựng (nguồn đang chọn không tải)", async () => {
    const t = setup({ cache: { openrouter: { at: Date.now(), models: [FREE_OR] } } });
    const r: any = await t.run();
    assert.deepStrictEqual(t.fetched, ["cohere"]);
    assert.strictEqual(r.candidates.length, 2);
  });

  it("tải danh mục lỗi: ghi mốc lỗi, bỏ nguồn đó; trong 10 phút không thử tải lại", async () => {
    let calls = 0;
    const t = setup({ cache: { openrouter: { at: Date.now(), models: [FREE_OR] } }, fetchModels: async () => (calls++, Promise.reject(new Error("503"))) });
    const r1: any = await t.run();
    assert.strictEqual(calls, 1);
    assert.ok(t.catalogFail.cohere > 0);
    assert.deepStrictEqual(r1.candidates.map((c: any) => c.provider.id), ["openrouter"]);
    await t.run();
    assert.strictEqual(calls, 1, "trong 10 phút sau lần lỗi không tải lại");
  });

  it("tải danh mục treo: hết thời gian chờ thì vẫn dựng với những gì đã có", async () => {
    const t = setup({ cache: { openrouter: { at: Date.now(), models: [FREE_OR] } }, fetchModels: () => new Promise(() => {}), waitMs: 30 });
    const t0 = Date.now();
    const r: any = await t.run();
    assert.ok(Date.now() - t0 < 2000);
    assert.strictEqual(r.kind, "ok");
    assert.deepStrictEqual(r.candidates.map((c: any) => c.provider.id), ["openrouter"]);
  });

  it("đã bấm Dừng: trả aborted, không dựng vòng xoay", async () => {
    const r = await setup({ abort: true }).run();
    assert.strictEqual(r.kind, "aborted");
  });

  it("không nguồn nào dùng được: trả lỗi (bên gọi báo và trả lại chữ đã gõ)", async () => {
    const r: any = await setup({ keys: {}, cache: {} }).run();
    assert.strictEqual(r.kind, "error");
    assert.match(String(r.error?.message), /chưa có nguồn nào có key/);
  });

  it("Cohere có bậc miễn phí được nhận vào vòng; nguồn không có key thì bỏ", async () => {
    const r: any = await setup({ keys: { openrouter: "k1" } }).run();
    assert.deepStrictEqual(r.candidates.map((c: any) => c.provider.id), ["openrouter"]);
    void CO;
  });
});
