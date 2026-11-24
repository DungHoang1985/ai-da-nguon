/** Đợt 2 — các ca thêm sau lần soát mã độc lập 16/11/2026 (tasks/dot/dot-2/bang-chung/soat-ma-doc-lap.md).
 *  Mọi "key" trong tệp này là chuỗi giả ghép từ ký tự lặp, không phải key thật. */
import * as assert from "assert";
import { ApiError, classifyError, describeError, ensureOk, explainApiError, isRetryable, maskSecrets } from "../../src/core/errors";
import { extractCode, looksLikeApiKey, redact } from "../../src/core/oauth";
import { chatCreditWarning, describeKeyInfo, fetchKeyInfo, parseKeyInfo } from "../../src/core/credits";
import { createStepAccumulator } from "../../src/core/agent/llm";
import { getProvider } from "../../src/core/providers";

const OR = getProvider("openrouter")!;
const GEMINI = getProvider("gemini")!;
const OPENAI = getProvider("openai")!;

const GIA = {
  gemini: "AIza" + "x".repeat(35),
  openai: "sk-proj-" + "0".repeat(40),
  anthropic: "sk-ant-api03-" + "0".repeat(40),
  openrouter: "sk-or-v1-" + "0".repeat(40),
  groq: "gsk_" + "0".repeat(40),
};

const json = (status: number, body: any) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

describe("Đợt 2 (soát mã T1, Th1) — thông báo lỗi không bao giờ mang key", () => {
  it("maskSecrets che mọi dạng key đã biết, kể cả dạng nhà cung cấp đã che dở", () => {
    for (const k of Object.values(GIA)) {
      const out = maskSecrets(`lỗi với key ${k} của anh`);
      assert.ok(!out.includes(k), k.slice(0, 8));
      assert.ok(!out.includes(k.slice(-12)), "còn sót đuôi key");
    }
    const dotted = maskSecrets("Incorrect API key provided: sk-proj-****abcd. You can find your API key at …");
    assert.ok(!dotted.includes("abcd") && dotted.startsWith("Incorrect API key provided: sk-…") && dotted.includes("You can find"));
    assert.ok(!maskSecrets("key sk-ab...wxyz1234 is wrong").includes("wxyz1234"));
    assert.ok(!maskSecrets("Authorization: Bearer abcdefghijklmnop0123456789").includes("abcdefghijklmnop"));
    assert.ok(!maskSecrets("Consumer 'api_key:" + "q".repeat(30) + "' has been suspended").includes("q".repeat(30)));
    // Câu thường không bị đổi
    const plain = "max_tokens: 99999 > 32000, which is the maximum allowed; tokens used: 12";
    assert.strictEqual(maskSecrets(plain), plain);
  });

  it("câu lỗi của nhà cung cấp nhắc lại key: thông báo, chi tiết lỗi đều đã che", async () => {
    const d = `Permission denied: Consumer 'api_key:${GIA.gemini}' has been suspended.`;
    const x = explainApiError(403, d, { provider: GEMINI });
    assert.ok(!x.text.includes(GIA.gemini));
    const expired = `API key expired. Please renew the API key ${GIA.gemini}.`;
    const y = explainApiError(400, expired, { provider: GEMINI, body: { error: { message: expired, details: [{ reason: "API_KEY_INVALID" }] } } });
    assert.ok(y.reason);
    assert.ok(!y.text.includes(GIA.gemini));
    let caught: any;
    try {
      await ensureOk(json(401, { error: { message: `Incorrect API key provided: ${GIA.openai}.`, code: "invalid_api_key" } }), OPENAI);
    } catch (e) {
      caught = e;
    }
    assert.ok(caught instanceof ApiError);
    const info = describeError(caught, { provider: OPENAI });
    assert.ok(!info.text.includes(GIA.openai) && !info.detail.includes(GIA.openai));
    assert.strictEqual(info.reason, "key_invalid");
  });

  it("lớp chặn cuối: lỗi bất kỳ mang key trong message cũng bị che ở describeError", () => {
    const info = describeError(new ApiError(`Nhà cung cấp nói: ${GIA.anthropic}`, 500));
    assert.ok(!info.text.includes(GIA.anthropic) && !info.detail.includes(GIA.anthropic));
    const app = describeError(new Error(`hỏng với ${GIA.groq}`));
    assert.ok(!app.text.includes(GIA.groq) && !app.detail.includes(GIA.groq));
  });

  it("lỗi giữa luồng nhắc lại key cũng được che", () => {
    const acc = createStepAccumulator("openai", () => {}, undefined, OR);
    assert.throws(
      () => acc.push(JSON.stringify({ error: { code: 500, message: `upstream rejected ${GIA.openrouter}` } })),
      (e: any) => e instanceof ApiError && !e.message.includes(GIA.openrouter)
    );
  });

  it("redact của luồng đăng nhập che cả AIza…, gsk_… chứ không chỉ sk-…", () => {
    const out = redact(`a ${GIA.gemini} b ${GIA.groq} c ${GIA.openrouter}`, []);
    assert.ok(!out.includes(GIA.gemini) && !out.includes(GIA.groq) && !out.includes(GIA.openrouter));
  });
});

describe("Đợt 2 (soát mã T2) — dán nhầm API key vào ô mã", () => {
  it("key của nhà cung cấp khác không bao giờ được coi là mã đăng nhập", () => {
    for (const k of [GIA.gemini, GIA.openai, GIA.anthropic, GIA.groq]) {
      const r = extractCode(k, OR);
      assert.strictEqual(r.kind, "invalid", k.slice(0, 8));
      assert.match((r as any).reason, /là API key, không phải mã đăng nhập OpenRouter/);
      assert.ok(!(r as any).reason.includes(k));
    }
    assert.ok(looksLikeApiKey(GIA.gemini) && looksLikeApiKey(GIA.groq) && !looksLikeApiKey("ma-gia-0123456789abcdef"));
  });

  it("key OpenRouter dán vào ô mã của OpenRouter thì vẫn nhận là key", () => {
    assert.deepStrictEqual(extractCode(GIA.openrouter, OR), { kind: "key", key: GIA.openrouter });
  });
});

describe("Đợt 2 (soát mã Th2, Th3, Th4) — phân loại lỗi không báo nhầm, không làm mất tự thử lại", () => {
  const retryable = (status: number, detail: string, body?: any) => {
    const x = explainApiError(status, detail, { provider: OR, body });
    const e = new ApiError(x.text, status, undefined, x.reason);
    if (x.transient) e.transient = true;
    return isRetryable(e);
  };

  it("429 kèm lời mời nạp tiền vẫn là hết hạn mức tạm thời, tự thử lại được", () => {
    for (const d of [
      "Too many requests. Purchase credits to raise limits.",
      "Rate limit exceeded. Add more credits to increase your limit.",
      "You exceeded your current quota, please check your plan and billing details.",
      "Rate limit exceeded: free-models-per-min. Add 10 credits to unlock 1000 free model requests per day",
    ]) {
      assert.deepStrictEqual(classifyError(429, d), {}, d);
      assert.ok(retryable(429, d), d);
    }
  });

  it("429 có tín hiệu chắc chắn thì vẫn nhận ra hết tiền / hết lượt miễn phí trong ngày", () => {
    assert.strictEqual(classifyError(429, "Your prepayment credits are depleted.").reason, "no_credit");
    assert.strictEqual(classifyError(429, "x", { error: { code: "insufficient_quota", message: "x" } }).reason, "no_credit");
    assert.strictEqual(classifyError(429, "Insufficient credits").reason, "no_credit");
    assert.strictEqual(classifyError(429, "Rate limit exceeded: free-models-per-day. Add 10 credits to unlock 1000 free model requests per day").reason, "free_limit");
  });

  it("400 / 403 chỉ vì có chữ “api key” đâu đó thì không bị đổ cho key", () => {
    for (const [st, d] of [
      [400, "tools: web_search is disabled for this model (api key scope)"],
      [403, "Unauthorized: model access is disabled for this organization (api key ok)"],
      [400, "tool_choice: unauthorized key param"],
      [403, "This model is disabled in your region"],
      [400, "thinking: budget_tokens is disabled when api key policy forbids it"],
    ] as [number, string][]) {
      assert.deepStrictEqual(classifyError(st, d), {}, `${st} ${d}`);
    }
  });

  it("400 / 403 nói rõ về key thì vẫn nhận ra", () => {
    assert.strictEqual(classifyError(400, "API key expired. Please renew the API key.").reason, "key_expired");
    assert.strictEqual(classifyError(403, "Your API key was reported as leaked. Please use another API key.").reason, "key_expired");
    assert.strictEqual(classifyError(403, "This key has been disabled").reason, "key_expired");
    assert.strictEqual(classifyError(400, "API key not valid. Please pass a valid API key.").reason, "key_invalid");
    assert.strictEqual(classifyError(401, "Unauthorized").reason, "key_invalid");
  });
});

describe("Đợt 2 (soát mã T4b, Th7) — hạn mức", () => {
  const sample = { data: { limit: null, limit_remaining: null, usage: 1, is_free_tier: true, free_model_daily_requests: { used: 1, limit: 50, remaining: 49 } } };

  it("/credits lỗi mạng hoặc 5xx: chưa biết có được xem hay không (không kết luận là bị cấm)", async () => {
    const mk = (second: () => Promise<Response>) =>
      (async (url: any) => (String(url).endsWith("/key") ? json(200, sample) : second())) as any;
    const a = await fetchKeyInfo(OR, "k", { tryBalance: true, fetchFn: mk(async () => json(503, { error: { message: "down" } })) });
    assert.strictEqual(a.balanceAllowed, undefined);
    const b = await fetchKeyInfo(OR, "k", {
      tryBalance: true,
      fetchFn: mk(async () => {
        throw new TypeError("Failed to fetch");
      }),
    });
    assert.strictEqual(b.balanceAllowed, undefined);
    const c = await fetchKeyInfo(OR, "k", { tryBalance: true, fetchFn: mk(async () => json(403, { error: { message: "Only management keys can perform this operation" } })) });
    assert.strictEqual(c.balanceAllowed, false);
    // Chưa biết thì thẻ không ghi "key này không được xem"
    assert.ok(!describeKeyInfo(a.info, { balanceAllowed: a.balanceAllowed }).lines.some((l) => /không được xem/.test(l)));
    assert.ok(describeKeyInfo(c.info, { balanceAllowed: c.balanceAllowed }).lines.some((l) => /không được xem/.test(l)));
  });

  it("key có trần nhưng máy chủ không báo còn lại bao nhiêu: không coi là đã chạm trần", () => {
    const info = parseKeyInfo({ data: { ...sample.data, limit: 5, limit_remaining: null } });
    const v = describeKeyInfo(info);
    assert.strictEqual(v.level, "ok");
    assert.ok(v.lines.some((l) => /Trần của key: 5,00 USD/.test(l)));
    assert.strictEqual(chatCreditWarning(info, { freeModel: false }), undefined);
  });
});
