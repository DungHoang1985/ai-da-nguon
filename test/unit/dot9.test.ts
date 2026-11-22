/** Đợt 9: hai chế độ Miễn phí / Trả phí và xoay vòng nguồn miễn phí (tasks/dot-9.md, tiêu chí T1–T6). */
import * as assert from "assert";
import { allProviders, customProvider, getProvider } from "../../src/core/providers";
import { DEFAULT_FREE_ORDER, freeKeyWarning, modelInMode, normalizeFreeOrder, providerInMode } from "../../src/core/modes";
import { filterModels } from "../../src/core/catalog";
import {
  Candidate,
  Cooldown,
  FREE_STALL_MS,
  FreeRotation,
  SHORT_COOLDOWN_MS,
  buildFreeCandidates,
  classifyFailure,
  dailyResetAt,
  estimateTokens,
  retryHintMs,
} from "../../src/core/agent/rotation";
import { buildStepRequest } from "../../src/core/agent/llm";
import { runAgent } from "../../src/core/agent/loop";
import { runTurn } from "../../src/core/agent/turn";
import { AgentEvent, AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { ApiError } from "../../src/core/errors";
import { Storage } from "../../src/core/storage";

const P = (id: string) => getProvider(id)!;
const FREE_IDS = ["gemini", "groq", "openrouter", "cohere"];

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200, headers: { "content-type": "text/event-stream" } });
}
const oaText = (t: string) => sse([{ choices: [{ delta: { content: t } }] }, { choices: [{ delta: {}, finish_reason: "stop" }] }, "[DONE]"]);
const oaTool = (id: string, name: string, args: any) =>
  sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: JSON.stringify(args) } }] } }] }, { choices: [{ delta: {}, finish_reason: "tool_calls" }] }, "[DONE]"]);
const gmText = (t: string) => sse([{ candidates: [{ content: { role: "model", parts: [{ text: t }] }, finishReason: "STOP" }] }]);
const err = (status: number, message: string, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify({ error: { message } }), { status, headers: { "content-type": "application/json", ...headers } });
/** Nhà cung cấp nhận yêu cầu rồi im lặng */
const hang = (init: any) =>
  new Promise<Response>((_res, rej) => init.signal.addEventListener("abort", () => rej(new DOMException("aborted", "AbortError"))));

/** Máy giả lập nhiều nguồn: định tuyến theo tên miền, mỗi nguồn một hàng đợi phản hồi; ghi lại nguồn nào bị gọi */
function router(plan: Record<string, ((init: any, body: any) => Response | Promise<Response>)[]>) {
  const hosts: Record<string, string> = {
    "generativelanguage.googleapis.com": "gemini",
    "api.groq.com": "groq",
    "openrouter.ai": "openrouter",
    "api.cohere.ai": "cohere",
    "api.anthropic.com": "anthropic",
  };
  const calls: string[] = [];
  const bodies: Record<string, any[]> = {};
  const fetchFn = (async (url: string, init: any) => {
    const host = new URL(url).host;
    const pid = hosts[host] || host;
    calls.push(pid);
    const body = init?.body ? JSON.parse(init.body) : undefined;
    (bodies[pid] ||= []).push(body);
    const q = plan[pid];
    if (!q || !q.length) throw new Error(`Bài thử không chờ lần gọi tới ${pid}`);
    return q.shift()!(init, body);
  }) as any;
  return { fetchFn, calls, bodies };
}

const cand = (pid: string, model = "m", extra: Partial<Candidate> = {}): Candidate => ({ provider: P(pid), key: "K-" + pid, model, ...extra });
const memStore = () => {
  let v: Record<string, Cooldown> = {};
  return { get: () => v, set: (x: Record<string, Cooldown>) => (v = x), peek: () => v };
};
const USER: AgentMessage[] = [
  { role: "system", content: "SYS" },
  { role: "user", content: "Xin chào" },
];

describe("Đợt 9 — chế độ Miễn phí / Trả phí (T1, T4)", () => {
  it("Nguồn có bậc miễn phí đúng bảng mục 3; các nguồn còn lại chỉ ở Trả phí", () => {
    const free = allProviders().filter((p) => providerInMode(p, "free")).map((p) => p.id).sort();
    assert.deepStrictEqual(free, [...FREE_IDS].sort());
    for (const id of ["anthropic", "openrouter-claude", "deepseek", "deepseek-anthropic", "minimax", "kimi", "openai", "mistral", "cerebras"]) {
      assert.ok(!providerInMode(P(id), "free"), id);
      assert.ok(providerInMode(P(id), "paid"), id);
    }
    for (const id of FREE_IDS) assert.ok(providerInMode(P(id), "paid"), id + " có cả bậc trả phí");
  });
  it("OpenRouter: Miễn phí chỉ model :free, Trả phí chỉ model trả tiền", () => {
    const or = P("openrouter");
    assert.ok(modelInMode(or, { id: "x/y:free", free: true }, "free"));
    assert.ok(!modelInMode(or, { id: "anthropic/claude-sonnet-5", free: false }, "free"));
    assert.ok(modelInMode(or, { id: "anthropic/claude-sonnet-5", free: false }, "paid"));
    assert.ok(!modelInMode(or, { id: "x/y:free", free: true }, "paid"));
  });
  it("Gemini ở chế độ Miễn phí: chỉ Flash, Flash-Lite, Gemma; không Pro, ảnh, giọng nói", () => {
    const g = P("gemini");
    for (const id of ["gemini-3.8-flash", "gemini-3.8-flash-lite", "models/gemini-2.5-flash", "gemma-3-27b-it", "gemini-flash-latest"]) assert.ok(modelInMode(g, { id }, "free"), id);
    for (const id of ["gemini-3.8-pro", "gemini-2.5-pro", "gemini-2.5-flash-image", "gemini-2.5-flash-preview-tts", "gemini-live-2.5-flash", "text-embedding-004"])
      assert.ok(!modelInMode(g, { id }, "free"), id);
    assert.ok(modelInMode(g, { id: "gemini-3.8-pro" }, "paid"), "Trả phí: key đã bật thanh toán dùng được mọi model");
  });
  it("Lọc danh sách model theo chế độ (bộ chọn model)", () => {
    const models = [
      { id: "a:free", name: "A", free: true },
      { id: "b", name: "B", free: false },
    ];
    const base = { providerId: "openrouter", freeOnly: false, toolsOnly: false, agentOn: false, search: "" };
    assert.deepStrictEqual(filterModels(models, { ...base, mode: "free", provider: P("openrouter") }).map((m) => m.id), ["a:free"]);
    assert.deepStrictEqual(filterModels(models, { ...base, mode: "paid", provider: P("openrouter") }).map((m) => m.id), ["b"]);
    assert.strictEqual(filterModels(models, base).length, 2, "chưa chọn chế độ: như bản cũ");
  });
  it("Nguồn tự khai: mặc định Trả phí; người dùng khai miễn phí thì chỉ ở Miễn phí", () => {
    const paid = customProvider({ id: "tuy-chinh-a", name: "A", baseUrl: "https://a.example.com", models: ["m"] })!;
    const free = customProvider({ id: "tuy-chinh-b", name: "B", baseUrl: "https://b.example.com", models: ["m"], free: true })!;
    assert.ok(!providerInMode(paid, "free") && providerInMode(paid, "paid"));
    assert.ok(providerInMode(free, "free") && !providerInMode(free, "paid"));
  });
  it("Cảnh báo key có thể bị tính tiền (R1): Gemini, Groq, Cohere có; OpenRouter, nguồn trả phí không", () => {
    for (const id of ["gemini", "groq", "cohere"]) assert.match(freeKeyWarning(P(id)), /bật thanh toán/, id);
    assert.match(freeKeyWarning(P("cohere")), /không cho dùng vào mục đích thương mại/);
    assert.strictEqual(freeKeyWarning(P("openrouter")), "");
    assert.strictEqual(freeKeyWarning(P("anthropic")), "");
  });
  it("Thứ tự xoay vòng: mặc định OpenRouter → Gemini → Groq → Cohere (từ 22/11/2026); bỏ mã lạ, nguồn trả phí, mã trùng", () => {
    assert.deepStrictEqual(DEFAULT_FREE_ORDER, ["openrouter", "gemini", "groq", "cohere"]);
    assert.deepStrictEqual(normalizeFreeOrder(undefined, allProviders()), DEFAULT_FREE_ORDER);
    assert.deepStrictEqual(normalizeFreeOrder(["cohere", "anthropic", "la", "cohere", "groq"], allProviders()), ["cohere", "groq", "openrouter", "gemini"]);
  });
  it("Lưu chế độ, thứ tự xoay vòng, thời điểm nghỉ trong kho cục bộ", () => {
    const mem: Record<string, string> = {};
    const st = new Storage({ getItem: (k) => mem[k] ?? null, setItem: (k, v) => void (mem[k] = v), removeItem: (k) => void delete mem[k] });
    assert.strictEqual(st.getSelection().mode, undefined, "chưa chọn → khung hiện thẻ chọn chế độ");
    st.setSelection({ mode: "free", freeOrder: ["groq", "gemini"] });
    st.setCooldowns({ gemini: { until: 5, reason: "x" } });
    const st2 = new Storage({ getItem: (k) => mem[k] ?? null, setItem: (k, v) => void (mem[k] = v), removeItem: (k) => void delete mem[k] });
    assert.strictEqual(st2.getSelection().mode, "free");
    assert.deepStrictEqual(st2.getSelection().freeOrder, ["groq", "gemini"]);
    assert.strictEqual(st2.getCooldowns().gemini.until, 5);
  });
});

describe("Đợt 9 — dựng vòng xoay (T4)", () => {
  const cache = {
    gemini: { models: [{ id: "gemini-3.8-pro", name: "Pro", tools: true }, { id: "gemini-3.8-flash", name: "Flash", tools: true, context: 1000000 }] },
    groq: { models: [{ id: "openai/gpt-oss-120b", name: "OSS", tools: true, context: 131072 }] },
    openrouter: { models: [{ id: "anthropic/claude-sonnet-5", name: "S", free: false, tools: true }, { id: "nvidia/nemotron-3-ultra:free", name: "N", free: true, tools: true }] },
    cohere: { models: [{ id: "command-a-03-2025", name: "C", tools: true }] },
    anthropic: { models: [{ id: "claude-sonnet-5", name: "S", tools: true }] },
  };
  const keys = { gemini: "g", groq: "q", openrouter: "o", cohere: "c", anthropic: "a" };
  it("Chỉ nguồn miễn phí có key, model miễn phí; nguồn đang chọn đứng đầu", () => {
    const list = buildFreeCandidates({
      order: DEFAULT_FREE_ORDER,
      selected: { providerId: "groq", modelId: "openai/gpt-oss-120b" },
      keys,
      getProvider,
      modelCache: cache as any,
      remembered: () => "",
      agentOn: true,
    });
    assert.deepStrictEqual(list.map((c) => `${c.provider.id}:${c.model}`), [
      "groq:openai/gpt-oss-120b",
      "openrouter:nvidia/nemotron-3-ultra:free",
      "gemini:gemini-3.8-flash",
      "cohere:command-a-03-2025",
    ]);
    assert.ok(list.every((c) => providerInMode(c.provider, "free")), "không có nguồn trả phí nào trong vòng");
  });
  it("Nguồn đang chọn là model trả phí (Gemini Pro đã nhớ) → thay bằng model miễn phí; thiếu key / thiếu danh mục thì bỏ", () => {
    const list = buildFreeCandidates({
      order: DEFAULT_FREE_ORDER,
      selected: { providerId: "gemini", modelId: "gemini-3.8-pro" },
      keys: { gemini: "g", openrouter: "o" },
      getProvider,
      modelCache: { gemini: cache.gemini } as any,
      remembered: () => "gemini-3.8-pro",
      agentOn: true,
    });
    assert.deepStrictEqual(list.map((c) => `${c.provider.id}:${c.model}`), ["gemini:gemini-3.8-flash"]);
  });
  it("Nguồn dự phòng: model đã nhớ bị đánh dấu ⚠ không gọi công cụ đáng tin thì chọn model khác", () => {
    const list = buildFreeCandidates({
      order: ["openrouter"],
      keys,
      getProvider,
      modelCache: { openrouter: { models: [...cache.openrouter.models, { id: "baidu/ernie-5.5-flash-preview:free", name: "E", free: true, tools: true }] } } as any,
      remembered: () => "baidu/ernie-5.5-flash-preview:free",
      agentOn: true,
    });
    assert.strictEqual(list[0].model, "nvidia/nemotron-3-ultra:free");
  });
  it("Nguồn đang chọn là nguồn trả phí (Anthropic) → không vào vòng", () => {
    const list = buildFreeCandidates({ order: DEFAULT_FREE_ORDER, selected: { providerId: "anthropic", modelId: "claude-sonnet-5" }, keys, getProvider, modelCache: cache as any, remembered: () => "", agentOn: true });
    assert.ok(!list.some((c) => c.provider.id === "anthropic"));
  });
  it("Không có nguồn nào → báo rõ", () => {
    assert.throws(() => new FreeRotation([]), /chưa có nguồn nào có key/);
  });
});

describe("Đợt 9 — phân loại lỗi để xoay vòng", () => {
  it("Im lặng, 429 phút, 429 ngày, 429 tháng, 413, vượt ngữ cảnh, lỗi key, 503, mất mạng", () => {
    const stall = new ApiError("im lặng");
    stall.stall = true;
    assert.strictEqual(classifyFailure(stall), "stall");
    assert.strictEqual(classifyFailure(new ApiError("Rate limit: requests per minute", 429)), "rate");
    const d = new ApiError("Quota PerDay", 429);
    d.daily = true;
    assert.strictEqual(classifyFailure(d), "daily");
    assert.strictEqual(classifyFailure(new ApiError("trial key limited to 1000 API calls / month", 429)), "monthly");
    assert.strictEqual(classifyFailure(new ApiError("Request too large", 413)), "too_large");
    assert.strictEqual(classifyFailure(new ApiError("This model's maximum context length is 131072 tokens", 400)), "too_large");
    assert.strictEqual(classifyFailure(new ApiError("Key sai", 401, undefined, "key_invalid")), "key");
    assert.strictEqual(classifyFailure(new ApiError("Overloaded", 503)), "server");
    assert.strictEqual(classifyFailure(new TypeError("Failed to fetch")), "network");
    assert.strictEqual(classifyFailure(new ApiError("Lỗi 404 (No endpoints found that support tool use)", 404)), "tools");
  });
  it("Mốc đặt lại hạn mức ngày: Gemini nửa đêm giờ Thái Bình Dương, OpenRouter nửa đêm UTC, khác 24 giờ", () => {
    const now = Date.UTC(2026, 10, 19, 3, 0, 0); // 19/11 03:00 UTC = 18/11 19:00 PST
    assert.strictEqual(dailyResetAt("gemini", now), Date.UTC(2026, 10, 19, 8, 0, 0));
    assert.strictEqual(dailyResetAt("openrouter", now), Date.UTC(2026, 10, 20, 0, 0, 0));
    assert.strictEqual(dailyResetAt("groq", now), now + 86400000);
  });
  it("Ước lượng token để bỏ nguồn quá nhỏ", () => {
    assert.strictEqual(estimateTokens([{ role: "user", content: "x".repeat(300) }]), 100);
  });
});

describe("Đợt 9 — xoay vòng khi gọi model (T3, T4, T5)", () => {
  it("Hạn chờ im lặng mặc định ở chế độ Miễn phí là 60 giây (anh chốt 19/11/2026)", () => {
    assert.strictEqual(FREE_STALL_MS, 60000);
  });
  it("Im lặng → chuyển ngay sang nguồn kế tiếp, không gọi lại nguồn im lặng; nguồn đó nghỉ 60 giây", async () => {
    const r = router({ gemini: [(init) => hang(init)], groq: [() => oaText("Chào anh")] });
    const store = memStore();
    const t0 = 1_000_000;
    const rot = new FreeRotation([cand("gemini"), cand("groq")], { store, now: () => t0, stallMs: 20 });
    const sw: any[] = [];
    const res = await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }, (i) => sw.push(i));
    assert.strictEqual(res.text, "Chào anh");
    assert.deepStrictEqual(r.calls, ["gemini", "groq"]);
    assert.strictEqual(sw.length, 1);
    assert.match(sw[0].reason, /không phản hồi/);
    assert.strictEqual(store.peek()["gemini|m"].until, t0 + SHORT_COOLDOWN_MS);
    assert.strictEqual(rot.current().provider.id, "groq", "bước sau của lượt bắt đầu từ nguồn đang chạy được");
  });
  it("429 theo phút / 503 / 413 → chuyển ngay, không tự thử lại cùng nguồn", async () => {
    const r = router({
      gemini: [() => err(429, "Resource exhausted: requests per minute")],
      groq: [() => err(413, "Request too large for model in organization on tokens per minute (TPM)")],
      openrouter: [() => err(503, "Service unavailable")],
      cohere: [() => oaText("OK")],
    });
    const rot = new FreeRotation([cand("gemini"), cand("groq"), cand("openrouter"), cand("cohere")], { store: memStore() });
    const res = await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.strictEqual(res.text, "OK");
    assert.deepStrictEqual(r.calls, ["gemini", "groq", "openrouter", "cohere"]);
  });
  it("429 theo phút có retry-after dài hơn 60 giây → nghỉ theo retry-after", async () => {
    const r = router({ gemini: [() => err(429, "Too many requests per minute", { "retry-after": "120" })], groq: [() => oaText("OK")] });
    const store = memStore();
    const rot = new FreeRotation([cand("gemini"), cand("groq")], { store, now: () => 0 });
    await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.strictEqual(store.peek()["gemini|m"].until, 120000);
  });
  it("429 theo ngày → nguồn nghỉ đến mốc đặt lại; lượt sau (vòng mới) bỏ qua nguồn đó, kể cả khi mở lại khung (T5)", async () => {
    const r = router({ gemini: [() => err(429, "Quota exceeded for metric generate_content_free_tier_requests, limit: GenerateRequestsPerDayPerProjectPerModel")], groq: [() => oaText("1"), () => oaText("2")] });
    const store = memStore();
    const now = Date.UTC(2026, 10, 19, 3, 0, 0);
    await new FreeRotation([cand("gemini"), cand("groq")], { store, now: () => now }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.strictEqual(store.peek()["gemini|m"].until, Date.UTC(2026, 10, 19, 8, 0, 0));
    // Lượt sau: vòng mới, cùng kho lưu → không gọi Gemini
    const res = await new FreeRotation([cand("gemini"), cand("groq")], { store, now: () => now + 3600000 }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.strictEqual(res.text, "2");
    assert.deepStrictEqual(r.calls, ["gemini", "groq", "groq"]);
    // Qua mốc đặt lại → Gemini vào lại vòng
    const r2 = router({ gemini: [() => gmText("lại rồi")] });
    const back = await new FreeRotation([cand("gemini"), cand("groq")], { store, now: () => Date.UTC(2026, 10, 19, 8, 0, 1) }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r2.fetchFn });
    assert.strictEqual(back.text, "lại rồi");
  });
  it("Hết mọi nguồn → dừng, báo từng nguồn; không gọi nguồn nào ngoài vòng (không sang trả phí)", async () => {
    const r = router({ gemini: [() => err(503, "x")], groq: [() => err(401, "Invalid API Key")] });
    const rot = new FreeRotation([cand("gemini"), cand("groq")], { store: memStore() });
    await assert.rejects(rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }), (e: any) => {
      assert.match(e.message, /Đã thử hết các nguồn miễn phí/);
      assert.match(e.message, /Google Gemini: máy chủ lỗi hoặc quá tải/);
      assert.match(e.message, /Groq: lỗi key/);
      assert.match(e.message, /không tự chuyển sang nguồn trả phí/);
      return true;
    });
    assert.deepStrictEqual(r.calls, ["gemini", "groq"]);
    assert.ok(!r.calls.includes("anthropic"));
  });
  it("Nguồn quá nhỏ cho yêu cầu → bỏ qua, không gọi thử (R4)", async () => {
    const r = router({ gemini: [() => gmText("vừa")] });
    const big: AgentMessage[] = [{ role: "user", content: "x".repeat(30000) }];
    const rot = new FreeRotation([cand("groq", "m", { context: 8000 }), cand("gemini", "m", { context: 1000000 })], { store: memStore() });
    const res = await rot.call({ messages: big, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.strictEqual(res.text, "vừa");
    assert.deepStrictEqual(r.calls, ["gemini"]);
  });
  it("Người dùng bấm Dừng → không xoay, báo dừng", async () => {
    const ctl = new AbortController();
    const r = router({ gemini: [(init) => (setTimeout(() => ctl.abort(), 5), hang(init))], groq: [() => oaText("không được gọi")] });
    const rot = new FreeRotation([cand("gemini"), cand("groq")], { store: memStore(), stallMs: 5000 });
    await assert.rejects(rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn, signal: ctl.signal }), (e: any) => e.name === "AbortError");
    assert.deepStrictEqual(r.calls, ["gemini"]);
  });
  it("Chữ dở đã hiện rồi nguồn đứt → vẫn xoay, báo partial (R6)", async () => {
    const r = router({
      groq: [
        (init) => {
          const enc = new TextEncoder();
          const body = new ReadableStream({
            start(c) {
              c.enqueue(enc.encode('data: {"choices":[{"delta":{"content":"Đang vi"}}]}\n\n'));
              init.signal.addEventListener("abort", () => c.error(new DOMException("aborted", "AbortError")));
            },
          });
          return new Response(body, { status: 200, headers: { "content-type": "text/event-stream" } });
        },
      ],
      cohere: [() => oaText("Trả lời đầy đủ")],
    });
    const rot = new FreeRotation([cand("groq"), cand("cohere")], { store: memStore(), stallMs: 20 });
    const sw: any[] = [];
    const res = await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }, (i) => sw.push(i));
    assert.strictEqual(res.text, "Trả lời đầy đủ");
    assert.strictEqual(sw[0].partial, true);
  });
});

describe("Đợt 9 — đổi nguồn giữa lượt có lệnh gọi công cụ (T6)", () => {
  const TOOLS: ToolSpec[] = [{ name: "word_read", description: "đọc", parameters: { type: "object", properties: {} } }];
  const hist: AgentMessage[] = [
    { role: "system", content: "SYS" },
    { role: "user", content: "Đọc văn bản" },
    { role: "assistant", content: "", toolCalls: [{ id: "call_1", name: "word_read", args: { from: 1 } }] },
    { role: "tool", toolCallId: "call_1", name: "word_read", content: '{"ok":true,"text":"Điều 1"}' },
  ];
  it("Gemini: lệnh gọi do nguồn khác sinh ra gửi thành chữ (không có functionCall thiếu chữ ký), kết quả kèm theo", () => {
    const { init } = buildStepRequest(P("gemini"), "K", "gemini-3.8-flash", hist, TOOLS, { foreignToolsAsText: true });
    const body = JSON.parse(String(init.body));
    const all = JSON.stringify(body.contents);
    assert.ok(!/functionCall/.test(all) && !/functionResponse/.test(all), all);
    assert.match(all, /đã gọi công cụ word_read/);
    assert.match(all, /Kết quả công cụ word_read/);
    assert.match(all, /Điều 1/);
  });
  it("Không bật cờ (chọn nguồn bằng tay như bản cũ): giữ nguyên cách gửi functionCall", () => {
    const { init } = buildStepRequest(P("gemini"), "K", "gemini-3.8-flash", hist, TOOLS);
    assert.match(String(init.body), /functionCall/);
  });
  it("Lệnh gọi của chính Gemini (có phần gốc) vẫn gửi nguyên văn", () => {
    const own: AgentMessage[] = [
      ...hist.slice(0, 2),
      { role: "assistant", content: "", toolCalls: [{ id: "g1", name: "word_read", args: {} }], raw: { kind: "gemini", parts: [{ functionCall: { name: "word_read", args: {} }, thoughtSignature: "SIG" }] } } as any,
      { role: "tool", toolCallId: "g1", name: "word_read", content: "{}" },
    ];
    const { init } = buildStepRequest(P("gemini"), "K", "gemini-3.8-flash", own, TOOLS, { foreignToolsAsText: true });
    assert.match(String(init.body), /"thoughtSignature":"SIG"/);
    assert.match(String(init.body), /functionResponse/);
  });
  it("Vòng tác tử: Groq gọi công cụ, bước sau Groq hết lượt → Gemini làm tiếp, lượt xong; khung nhận sự kiện đổi nguồn", async () => {
    const r = router({
      groq: [() => oaTool("call_1", "word_read", {}), () => err(429, "Rate limit reached: requests per minute (RPM)")],
      gemini: [() => gmText("Xong: Điều 1")],
    });
    const executed: string[] = [];
    const tools: HostToolSet = {
      host: "Word",
      specs: TOOLS,
      overview: async () => ({}),
      systemPrompt: () => "SYS",
      execute: async (name) => {
        executed.push(name);
        return { ok: true, text: "Điều 1" };
      },
    };
    const events: AgentEvent[] = [];
    const rot = new FreeRotation([cand("groq"), cand("gemini")], { store: memStore() });
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools,
      system: "SYS",
      messages: [{ role: "user", content: "Đọc văn bản" }],
      log: [],
      onEvent: (e) => events.push(e),
      fetchFn: r.fetchFn,
      rotation: rot,
    });
    assert.deepStrictEqual(r.calls, ["groq", "groq", "gemini"]);
    assert.deepStrictEqual(executed, ["word_read"]);
    assert.strictEqual(msgs[msgs.length - 1].content, "Xong: Điều 1");
    assert.ok(events.some((e) => e.type === "switch" && /Groq/.test(e.from) && /Gemini/.test(e.to)));
    const gemBody = JSON.stringify(r.bodies.gemini[0]);
    assert.ok(!/functionCall/.test(gemBody), "lệnh gọi của Groq gửi cho Gemini dưới dạng chữ");
  });
  it("Trò chuyện thường (không tác tử) cũng đi qua vòng xoay; chữ dở không vào lịch sử", async () => {
    const r = router({ gemini: [() => err(503, "x")], groq: [() => oaText("Chào")] });
    const events: any[] = [];
    const res = await runTurn({
      provider: P("gemini"),
      key: "K",
      model: "m",
      host: "Word",
      history: [],
      text: "Xin chào",
      onEvent: (e) => events.push(e),
      fetchFn: r.fetchFn,
      rotation: new FreeRotation([cand("gemini"), cand("groq")], { store: memStore() }),
    });
    assert.strictEqual(res.error, undefined);
    assert.strictEqual(res.history[res.history.length - 1].content, "Chào");
    assert.ok(events.some((e) => e.type === "switch"));
  });
  it("Không có vòng xoay (chế độ Trả phí): lỗi 503 vẫn tự thử lại cùng nguồn như cũ, không đổi nguồn", async () => {
    const r = router({ anthropic: [() => err(503, "x"), () => sse([{ type: "message_start", message: { usage: {} } }, { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } }, { type: "content_block_delta", index: 0, delta: { type: "text_delta", text: "OK" } }, { type: "message_stop" }])] });
    const res = await runTurn({
      provider: P("anthropic"),
      key: "K",
      model: "claude-sonnet-5",
      host: "Word",
      history: [],
      text: "Xin chào",
      onEvent: () => {},
      fetchFn: r.fetchFn,
      retryDelaysMs: [1],
      sleep: async () => {},
    });
    assert.strictEqual(res.error, undefined);
    assert.deepStrictEqual(r.calls, ["anthropic", "anthropic"]);
  });
});

describe("Đợt 9 — sửa theo soát mã độc lập", () => {
  it("H1: chế độ Miễn phí không bao giờ gửi plugin tìm web (tính tiền) của OpenRouter, kể cả khi bật tìm web", async () => {
    const r = router({ openrouter: [() => oaText("OK")] });
    await new FreeRotation([cand("openrouter", "x:free")], { store: memStore() }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn, webSearch: true });
    assert.strictEqual(r.bodies.openrouter[0].plugins, undefined);
  });
  it("H2: nguồn nghỉ 60 giây được dùng lại trong CÙNG lượt khi hết hạn nghỉ", async () => {
    let t = 0;
    const r = router({
      gemini: [() => err(429, "requests per minute"), () => gmText("Gemini lại chạy")],
      groq: [() => oaText("Groq bước 1"), () => err(429, "Rate limit reached: requests per minute (RPM)")],
    });
    const rot = new FreeRotation([cand("gemini"), cand("groq")], { store: memStore(), now: () => t });
    assert.strictEqual((await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn })).text, "Groq bước 1");
    t = 3 * 60000; // ba phút sau, Groq hết lượt theo phút; Gemini đã hết hạn nghỉ
    assert.strictEqual((await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn })).text, "Gemini lại chạy");
  });
  it("H2: lỗi key / yêu cầu quá lớn thì bỏ nguồn cả lượt", async () => {
    const r = router({ groq: [() => err(413, "Request too large")], gemini: [() => gmText("1"), () => gmText("2")] });
    const rot = new FreeRotation([cand("groq"), cand("gemini")], { store: memStore() });
    await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    await rot.call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.deepStrictEqual(r.calls, ["groq", "gemini", "gemini"]);
  });
  it("M1: nghỉ tính theo model — Gemini Flash hết lượt ngày thì Flash-Lite vẫn dùng được", async () => {
    const store = memStore();
    const now = Date.UTC(2026, 10, 19, 3, 0, 0);
    const d = () => err(429, "Quota exceeded for metric generate_content_free_tier_requests, limit: GenerateRequestsPerDayPerProjectPerModel");
    const r = router({ gemini: [d, () => gmText("Lite")] });
    await assert.rejects(new FreeRotation([cand("gemini", "gemini-3.8-flash")], { store, now: () => now }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }));
    const res = await new FreeRotation([cand("gemini", "gemini-3.8-flash-lite")], { store, now: () => now }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn });
    assert.strictEqual(res.text, "Lite");
  });
  it("M1: OpenRouter hết lượt ngày tính chung cả tài khoản cho mọi model :free", async () => {
    const store = memStore();
    const r = router({ openrouter: [() => err(429, "Rate limit exceeded: free-models-per-day")] });
    await assert.rejects(new FreeRotation([cand("openrouter", "a:free")], { store, now: () => 0 }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }));
    await assert.rejects(new FreeRotation([cand("openrouter", "b:free")], { store, now: () => 1000 }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }), /dùng lại lúc/);
    assert.deepStrictEqual(r.calls, ["openrouter"]);
  });
  it("M1: đọc thời gian chờ nhà cung cấp hẹn (Groq 'try again in 14m22.5s', Gemini 'retry in 37s') thay cho nghỉ 24 giờ", async () => {
    assert.strictEqual(retryHintMs("Please try again in 14m22.5s."), 862500);
    assert.strictEqual(retryHintMs("Please retry in 37.4s"), 37400);
    assert.strictEqual(retryHintMs("no hint"), 0);
    const store = memStore();
    const r = router({ groq: [() => err(429, "Rate limit reached on tokens per day (TPD): Limit 200000. Please try again in 14m22.5s.")] });
    await assert.rejects(new FreeRotation([cand("groq")], { store, now: () => 0 }).call({ messages: USER, tools: [], onText: () => {}, fetchFn: r.fetchFn }));
    assert.strictEqual(store.peek()["groq|m"].until, 862500);
  });
  it("L1: 429 theo phút kèm câu hạn mức tháng của key dùng thử không bị coi là hết lượt tháng", () => {
    assert.strictEqual(classifyFailure(new ApiError("You are using a Trial key, limited to 20 API calls / minute and 1000 calls / month", 429)), "rate");
  });
});
