/** Các mục soát mã "để lại" được làm nốt trước khi phát hành 4.7.2 (18/11/2026): tự thử lại lỗi tạm, không thử lại khi hết hạn mức theo ngày. */
import * as assert from "assert";
import { customProvider, getProvider } from "../../src/core/providers";
import { ApiError, ensureOk, fetchDiagnosed, isRetryable } from "../../src/core/errors";
import { callModelStep } from "../../src/core/agent/llm";
import { normalizeSlideList } from "../../src/office/pptTools";
import { runAgent } from "../../src/core/agent/loop";

const res = (status: number, body: any) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });
const thrown = async (p: Promise<any>) => {
  try {
    await p;
  } catch (e) {
    return e;
  }
  throw new Error("không ném lỗi");
};

describe("Nguồn tự khai không theo chuyển hướng (key x-api-key không đi sang miền khác)", () => {
  it("Yêu cầu tới nguồn tự khai mang redirect: error; nguồn có sẵn giữ mặc định", async () => {
    const seen: any[] = [];
    const fetchFn = (async (_u: string, init: any) => {
      seen.push(init);
      return new Response("{}", { status: 200 });
    }) as any;
    const custom = customProvider({ id: "tuy-chinh-x", name: "X", baseUrl: "https://llm.example.com/anthropic", anthropicHeaders: true, models: ["m"] } as any)!;
    await fetchDiagnosed(custom, custom.baseUrl + "/messages", { method: "POST" }, fetchFn);
    await fetchDiagnosed(getProvider("anthropic")!, "https://api.anthropic.com/v1/messages", { method: "POST" }, fetchFn);
    assert.strictEqual(seen[0].redirect, "error");
    assert.strictEqual(seen[1].redirect, undefined);
  });
});

describe("Tự thử lại: lỗi tạm và hết hạn mức theo ngày", () => {
  it("408 (hết giờ chờ phía máy chủ) được tự thử lại", () => {
    assert.ok(isRetryable(new ApiError("x", 408)));
  });
  it("Lỗi mạng đã qua bước chẩn đoán (máy chủ không với tới được) được tự thử lại; người dùng bấm Dừng thì không", () => {
    assert.ok(isRetryable(new TypeError("Failed to fetch")));
    assert.ok(isRetryable(new TypeError("NetworkError when attempting to fetch resource.")));
    assert.ok(!isRetryable(new TypeError("Cannot read properties of undefined")), "lỗi lập trình không phải lỗi mạng");
    const abort = new Error("aborted");
    abort.name = "AbortError";
    assert.ok(!isRetryable(abort));
  });
  it("429 hết hạn mức THEO NGÀY không thử lại (thử lại 3 lần cũng vậy, chỉ làm người dùng chờ thêm 23 giây)", async () => {
    const ca: [string, any][] = [
      ["gemini", { error: { code: 429, message: "Quota exceeded for metric: generativelanguage.googleapis.com/generate_content_free_tier_requests, limit: 250, model: gemini-3.8-flash. Please retry in 20h.", status: "RESOURCE_EXHAUSTED", details: [{ violations: [{ quotaId: "GenerateRequestsPerDayPerProjectPerModel-FreeTier" }] }] } }],
      ["openrouter", { error: { message: "Rate limit exceeded: free-models-per-day. Add 10 credits to unlock 1000 free model requests per day", code: 429 } }],
      ["groq", { error: { message: "Rate limit reached for model `llama-4` in organization `org_x` service tier `on_demand` on requests per day (RPD): Limit 1000, Used 1000, Requested 1.", type: "requests", code: "rate_limit_exceeded" } }],
    ];
    for (const [id, body] of ca) {
      const e: any = await thrown(ensureOk(res(429, body), getProvider(id as any)));
      assert.ok(e instanceof ApiError, id);
      assert.ok(!isRetryable(e), `${id}: hết hạn mức theo ngày mà vẫn thử lại`);
    }
  });
  it("Mạng đứt một lần (cả bước dò máy chủ) rồi có lại: tự thử lại, báo đúng lý do, trả lời được", async () => {
    const sse = (t: string) => new Response(`data: ${JSON.stringify({ choices: [{ delta: { content: t } }] })}\n\n`, { status: 200 });
    let calls = 0;
    const fetchFn = (async () => {
      calls++;
      if (calls <= 2) throw new TypeError("Failed to fetch"); // lần gọi + lần dò máy chủ đều hỏng
      return sse("Chào.");
    }) as any;
    const retries: any[] = [];
    let text = "";
    await callModelStep({
      provider: getProvider("groq")!,
      key: "K",
      model: "m",
      messages: [{ role: "user", content: "x" }],
      tools: [],
      onText: (t: string) => (text += t),
      fetchFn,
      sleep: async () => {},
      onRetry: (r: any) => retries.push(r),
    } as any);
    assert.strictEqual(text, "Chào.");
    assert.strictEqual(retries.length, 1);
    assert.strictEqual(retries[0].reason, "mất kết nối tới nhà cung cấp");
  });
  it("Nhà cung cấp chặn gọi từ khung add-in (máy chủ với tới được nhưng yêu cầu hỏng): không thử lại", async () => {
    let calls = 0;
    const fetchFn = (async (_u: string, init: any) => {
      calls++;
      if (init?.mode === "no-cors") return new Response("", { status: 200 });
      throw new TypeError("Failed to fetch");
    }) as any;
    await assert.rejects(
      callModelStep({ provider: getProvider("groq")!, key: "K", model: "m", messages: [{ role: "user", content: "x" }], tools: [], onText: () => {}, fetchFn, sleep: async () => {} } as any)
    );
    assert.strictEqual(calls, 2, "một lần gọi + một lần dò, không thử lại");
  });
  it("429 theo phút vẫn thử lại (không báo nhầm là theo ngày)", async () => {
    const ca: [string, any][] = [
      ["groq", { error: { message: "Rate limit reached for model `llama-4` on tokens per minute (TPM): Limit 8000, Used 7000, Requested 2000. Please try again in 7.5s.", code: "rate_limit_exceeded" } }],
      ["anthropic", { type: "error", error: { type: "rate_limit_error", message: "This request would exceed the rate limit for your organization of 50,000 input tokens per minute." } }],
      ["openrouter", { error: { message: "Rate limit exceeded: free-models-per-min.", code: 429 } }],
      ["openrouter", { error: { message: "Rate limit exceeded: free-models-per-min. Add 10 credits to unlock 1000 free model requests per day", code: 429 } }],
      ["gemini", { error: { code: 429, message: "Resource has been exhausted (e.g. check quota).", status: "RESOURCE_EXHAUSTED" } }],
    ];
    for (const [id, body] of ca) {
      const e: any = await thrown(ensureOk(res(429, body), getProvider(id as any)));
      assert.ok(isRetryable(e), `${id}: hạn mức theo phút bị coi là theo ngày`);
    }
  });
});

describe("ppt_add_slides: đếm phần dư đúng (soát mã Đợt 7, #8)", () => {
  it("Phần tử rỗng không tính vào phần dư; chuỗi JSON quá 10 slide được tính", () => {
    const eleven = Array.from({ length: 11 }, (_, i) => ({ title: `S${i + 1}` }));
    const withEmpty = [...Array.from({ length: 10 }, (_, i) => ({ title: `S${i + 1}` })), {}, null];
    assert.strictEqual(normalizeSlideList(eleven, Infinity).length - normalizeSlideList(eleven).length, 1);
    assert.strictEqual(normalizeSlideList(withEmpty, Infinity).length - normalizeSlideList(withEmpty).length, 0);
    const s = JSON.stringify(eleven);
    assert.strictEqual(normalizeSlideList(s, Infinity).length - normalizeSlideList(s).length, 1);
  });
});

describe("Dò lỗi lặp (soát mã Đợt 7, #6)", () => {
  const sseCall = (id: string, name: string, args: string) =>
    new Response(`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: args } }] }, finish_reason: "tool_calls" }] })}\n\n`, { status: 200 });
  const tools = (execute: any): any => ({
    host: "Word",
    specs: [{ name: "thay", description: "", parameters: { type: "object", properties: {} }, write: true }],
    overview: async () => ({}),
    systemPrompt: () => "S",
    execute,
  });
  const run = async (o: any) => {
    const events: any[] = [];
    let n = 0;
    await runAgent({
      provider: getProvider("groq")!,
      key: "K",
      model: "m",
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      maxSteps: 6,
      onEvent: (e: any) => events.push(e),
      fetchFn: (async () => sseCall(`c${++n}`, "thay", '{"find":"abc"}')) as any,
      ...o,
    } as any).catch(() => undefined);
    return events;
  };
  it("Công cụ trả {ok:false, message} y nguyên tham số 2 bước liền: dừng sớm", async () => {
    const ev = await run({ tools: tools(async () => ({ ok: false, found: 0, message: 'Không tìm thấy "abc"' })) });
    const stuck = ev.find((e) => e.type === "stuck");
    assert.ok(stuck, "không dừng sớm");
    assert.match(stuck.error, /Không tìm thấy/);
  });
  it("Người dùng từ chối cùng thao tác nhiều lần: không tính là lỗi lặp", async () => {
    const ev = await run({ tools: tools(async () => ({ ok: true })), askBeforeWrite: true, approve: async () => false });
    assert.ok(!ev.some((e) => e.type === "stuck"), "từ chối bị tính là lỗi lặp");
  });
});
