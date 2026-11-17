/** Đợt 6: nguồn nói cùng kiểu với Anthropic — phương ngữ (tiêu đề, bộ nhớ đệm, suy nghĩ, PDF), key dùng chung, nguồn tự khai */
import * as assert from "assert";
import { CUSTOM_PREFIX, PROVIDERS, allProviders, customProvider, getProvider, normalizeBaseUrl, setCustomSources } from "../../src/core/providers";
import { anthropicHeaders, fetchModels, modelsRequest, normalizeModels, staticModels } from "../../src/core/models";
import { buildStepRequest } from "../../src/core/agent/llm";
import { dialectThinkingParams, pdfCapable, thinkingToggleWorks } from "../../src/core/agent/thinking";
import { Storage, KVStore } from "../../src/core/storage";
import { AgentMessage, ToolSpec } from "../../src/core/agent/types";

const P = (id: string) => getProvider(id)!;
const HISTORY: AgentMessage[] = [
  { role: "system", content: "Bạn là trợ lý." } as any,
  { role: "user", content: "Chào" } as any,
];
const TOOLS: ToolSpec[] = [{ name: "doc", description: "đọc", parameters: { type: "object", properties: { a: { type: "string" } } } } as any];
const body = (r: { init: RequestInit }) => JSON.parse(String(r.init.body));
const headers = (r: { init: RequestInit }) => r.init.headers as Record<string, string>;
const lower = (h: Record<string, string>) => Object.keys(h).map((k) => k.toLowerCase()).sort();

function memStore(): KVStore {
  const m: Record<string, string> = {};
  return { getItem: (k) => (k in m ? m[k] : null), setItem: (k, v) => void (m[k] = v), removeItem: (k) => void delete m[k] };
}

describe("Đợt 6 — nguồn Anthropic gốc không đổi", () => {
  it("tiêu đề và thân yêu cầu của api.anthropic.com giữ nguyên như trước", () => {
    const r = buildStepRequest(P("anthropic"), "K", "claude-sonnet-4-5", HISTORY, TOOLS);
    assert.strictEqual(r.url, "https://api.anthropic.com/v1/messages");
    assert.deepStrictEqual(headers(r), {
      "x-api-key": "K",
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
      "content-type": "application/json",
    });
    const b = body(r);
    assert.deepStrictEqual(b.system, [{ type: "text", text: "Bạn là trợ lý.", cache_control: { type: "ephemeral" } }]);
    assert.deepStrictEqual(b.messages[0].content[0].cache_control, { type: "ephemeral" });
    assert.deepStrictEqual(anthropicHeaders("K"), headers(r));
  });

  it("các nguồn có sẵn không mang cấu hình phương ngữ", () => {
    for (const id of ["anthropic", "openrouter", "gemini", "groq", "openai", "deepseek"]) assert.strictEqual(P(id).dialect, undefined, id);
    assert.strictEqual(pdfCapable(P("anthropic")), true);
  });
});

describe("Đợt 6 — tiêu đề theo từng nguồn", () => {
  it("OpenRouter kênh Claude, MiniMax, Kimi: chỉ Authorization và content-type (bước dò trước của trình duyệt không cho tiêu đề anthropic-*)", () => {
    for (const id of ["openrouter-claude", "minimax", "kimi"]) {
      const r = buildStepRequest(P(id), "K", "m", HISTORY, TOOLS, { think: true });
      assert.deepStrictEqual(lower(headers(r)), ["authorization", "content-type"], id);
      assert.strictEqual(headers(r).Authorization, "Bearer K", id);
    }
  });

  it("DeepSeek: x-api-key kèm anthropic-version, không kèm tiêu đề gọi-từ-trình-duyệt của Anthropic", () => {
    const r = buildStepRequest(P("deepseek-anthropic"), "K", "deepseek-flash", HISTORY, TOOLS);
    assert.deepStrictEqual(lower(headers(r)), ["anthropic-version", "content-type", "x-api-key"]);
    assert.strictEqual(r.url, "https://api.deepseek.com/anthropic/v1/messages");
  });

  it("địa chỉ gọi đúng tài liệu của từng nguồn", () => {
    assert.strictEqual(buildStepRequest(P("openrouter-claude"), "K", "anthropic/claude-sonnet-5.5", HISTORY, []).url, "https://openrouter.ai/api/v1/messages");
    assert.strictEqual(buildStepRequest(P("minimax"), "K", "MiniMax-M3", HISTORY, []).url, "https://api.minimax.io/anthropic/v1/messages");
    assert.strictEqual(buildStepRequest(P("kimi"), "K", "kimi-k3.1-pro", HISTORY, []).url, "https://api.moonshot.ai/anthropic/v1/messages");
  });
});

describe("Đợt 6 — bộ nhớ đệm, suy nghĩ, PDF, tìm web", () => {
  it("nguồn bỏ qua cache_control thì không gắn mốc; OpenRouter kênh Claude vẫn gắn", () => {
    for (const id of ["deepseek-anthropic", "minimax", "kimi"]) {
      const b = body(buildStepRequest(P(id), "K", "m", HISTORY, TOOLS));
      assert.ok(!JSON.stringify(b).includes("cache_control"), id);
      assert.deepStrictEqual(b.system, [{ type: "text", text: "Bạn là trợ lý." }], id);
    }
    const b = body(buildStepRequest(P("openrouter-claude"), "K", "anthropic/claude-sonnet-5.5", HISTORY, TOOLS));
    assert.deepStrictEqual(b.system[0].cache_control, { type: "ephemeral" });
    assert.deepStrictEqual(b.messages[0].content[0].cache_control, { type: "ephemeral" });
  });

  it("DeepSeek: chỉ bật / tắt suy nghĩ, không kèm budget hay beta", () => {
    assert.deepStrictEqual(dialectThinkingParams(P("deepseek-anthropic"), "deepseek-v4-pro", true, 16000), { thinking: { type: "enabled" } });
    assert.deepStrictEqual(dialectThinkingParams(P("deepseek-anthropic"), "deepseek-v4-pro", false, 16000), { thinking: { type: "disabled" } });
    assert.strictEqual(thinkingToggleWorks(P("deepseek-anthropic"), "deepseek-flash"), true);
  });

  it("MiniMax, Kimi: không gửi tham số suy nghĩ nào (M3.1-Flash từ chối 'disabled'), công tắc ghi rõ là không có tác dụng", () => {
    for (const id of ["minimax", "kimi"]) {
      const b = body(buildStepRequest(P(id), "K", "m", HISTORY, TOOLS, { think: true }));
      assert.strictEqual(b.thinking, undefined, id);
      assert.strictEqual(b.output_config, undefined, id);
      assert.strictEqual(thinkingToggleWorks(P(id), "m"), false, id);
    }
  });

  it("OpenRouter kênh Claude: suy nghĩ theo đời model Claude (bỏ tiền tố anthropic/), không có block_binding và không có tiêu đề beta", () => {
    const r = buildStepRequest(P("openrouter-claude"), "K", "anthropic/claude-sonnet-5.5", HISTORY, TOOLS, { think: true });
    const b = body(r);
    assert.deepStrictEqual(b.thinking, { type: "adaptive" });
    assert.ok(b.output_config?.effort);
    assert.ok(!lower(headers(r)).includes("anthropic-beta"));
    // Model đời cũ: budget
    const old = dialectThinkingParams(P("openrouter-claude"), "anthropic/claude-haiku-4.5", true, 16000);
    assert.strictEqual(old.thinking.type, "enabled");
    assert.ok(old.thinking.budget_tokens >= 1024);
    assert.strictEqual(old.beta, undefined);
  });

  it("PDF nguyên tệp: chỉ nguồn ghi rõ là nhận; tìm web kiểu Anthropic: chỉ nguồn gốc", () => {
    assert.strictEqual(pdfCapable(P("openrouter-claude")), true);
    for (const id of ["deepseek-anthropic", "minimax", "kimi"]) assert.strictEqual(pdfCapable(P(id)), false, id);
    for (const id of ["openrouter-claude", "deepseek-anthropic", "minimax", "kimi"]) {
      const b = body(buildStepRequest(P(id), "K", "m", HISTORY, TOOLS, { webSearch: true }));
      assert.ok(!b.tools.some((t: any) => t.type), id);
    }
    assert.ok(body(buildStepRequest(P("anthropic"), "K", "claude-sonnet-4-5", HISTORY, TOOLS, { webSearch: true })).tools.some((t: any) => t.type === "web_search_20250305"));
  });
});

describe("Đợt 6 — danh sách model", () => {
  it("nguồn có danh sách cố định: không gọi mạng", async () => {
    let called = 0;
    const fake: any = async () => (called++, { ok: false, status: 500, text: async () => "", json: async () => ({}) });
    const list = await fetchModels(P("minimax"), "K", fake);
    assert.strictEqual(called, 0);
    assert.deepStrictEqual(list.map((m) => m.id), ["MiniMax-M3", "MiniMax-M3.1-Flash"]);
    assert.ok(list.every((m) => m.tools === true && m.vision === true));
    assert.deepStrictEqual(staticModels(P("deepseek-anthropic")).map((m) => [m.id, m.vision]), [["deepseek-flash", true], ["deepseek-v4-pro", false]]);
  });

  it("OpenRouter kênh Claude: lấy danh mục OpenRouter bằng Bearer và chỉ giữ model anthropic/", () => {
    const req = modelsRequest(P("openrouter-claude"), "K");
    assert.strictEqual(req.url, "https://openrouter.ai/api/v1/models");
    assert.deepStrictEqual(req.init.headers, { Authorization: "Bearer K" });
    const list = normalizeModels(P("openrouter-claude"), {
      data: [
        { id: "anthropic/claude-sonnet-5.5", name: "Claude Sonnet 5.5", architecture: { input_modalities: ["text", "image"] }, supported_parameters: ["tools"] },
        { id: "openai/gpt-x", name: "GPT" },
        { id: "x/anthropic/giả", name: "giả" },
      ],
    });
    assert.deepStrictEqual(list.map((m) => m.id), ["anthropic/claude-sonnet-5.5"]);
    assert.strictEqual(list[0].vision, true);
  });

  it("Kimi: địa chỉ liệt kê model lỗi thì dùng danh sách cố định", async () => {
    const fail: any = async () => {
      throw new TypeError("Failed to fetch");
    };
    const list = await fetchModels(P("kimi"), "K", fail);
    assert.deepStrictEqual(list.map((m) => m.id), ["kimi-k3.1-pro", "kimi-k3.1-pro-highspeed"]);
  });
});

describe("Đợt 6 — key dùng chung", () => {
  it("kênh kiểu Anthropic nhận key của nguồn gốc mà không lưu thêm bản sao; xóa key gốc thì kênh cũng mất", () => {
    const store = memStore();
    const st = new Storage(store);
    assert.strictEqual(st.getKey("openrouter-claude"), "");
    st.setKey("openrouter", "sk-or-abc");
    st.setKey("deepseek", "sk-ds");
    assert.strictEqual(st.getKey("openrouter-claude"), "sk-or-abc");
    assert.strictEqual(st.getKey("deepseek-anthropic"), "sk-ds");
    // Kho chỉ có hai key thật
    assert.deepStrictEqual(Object.keys(JSON.parse(store.getItem("aidn.keys")!)).sort(), ["deepseek", "openrouter"]);
    st.setKey("groq", "gsk");
    st.removeKey("groq");
    assert.deepStrictEqual(Object.keys(JSON.parse(store.getItem("aidn.keys")!)).sort(), ["deepseek", "openrouter"]);
    st.removeKey("openrouter");
    assert.strictEqual(st.getKey("openrouter-claude"), "");
    assert.strictEqual(st.getKeys()["openrouter-claude"], undefined);
  });

  it("mọi nguồn keyFrom trỏ tới một nguồn có thật và không tự có đăng nhập riêng", () => {
    for (const p of PROVIDERS.filter((x) => x.keyFrom)) {
      assert.ok(PROVIDERS.some((x) => x.id === p.keyFrom && !x.keyFrom), p.id);
      assert.strictEqual(p.oauth, undefined, p.id);
    }
  });
});

describe("Đợt 6 — nguồn tự khai", () => {
  afterEach(() => setCustomSources([]));

  it("chuẩn hóa địa chỉ: thêm /v1, bỏ / cuối và /messages thừa", () => {
    assert.strictEqual(normalizeBaseUrl("https://api.minimax.io/anthropic").url, "https://api.minimax.io/anthropic/v1");
    assert.strictEqual(normalizeBaseUrl(" https://api.siliconflow.com/ ").url, "https://api.siliconflow.com/v1");
    assert.strictEqual(normalizeBaseUrl("https://x.example/anthropic/v1/messages").url, "https://x.example/anthropic/v1");
    assert.strictEqual(normalizeBaseUrl("https://x.example/api/v2/").url, "https://x.example/api/v2");
    assert.strictEqual(normalizeBaseUrl("http://localhost:11434").url, "http://localhost:11434/v1");
    assert.strictEqual(normalizeBaseUrl("https://x.example/a").host, "x.example");
  });

  it("từ chối địa chỉ không an toàn: http ngoài máy này, kèm tài khoản, kèm tham số, không phải địa chỉ", () => {
    for (const bad of ["http://x.example/anthropic", "https://user:pw@x.example", "https://x.example/a?key=1", "https://x.example/a#b", "x.example", "javascript:alert(1)", "file:///c:/a", ""])
      assert.ok(normalizeBaseUrl(bad).error, bad);
  });

  it("nguồn tự khai mặc định chỉ gửi Authorization; bật kiểu gốc thì x-api-key kèm anthropic-version", () => {
    const a = customProvider({ id: CUSTOM_PREFIX + "a", name: "Thử", baseUrl: "https://x.example/anthropic", models: ["m1", "m2"] })!;
    const ra = buildStepRequest(a, "K", "m1", HISTORY, TOOLS, { think: true, webSearch: true });
    assert.strictEqual(ra.url, "https://x.example/anthropic/v1/messages");
    assert.deepStrictEqual(lower(headers(ra)), ["authorization", "content-type"]);
    assert.strictEqual(body(ra).thinking, undefined);
    assert.strictEqual(pdfCapable(a), false);
    assert.ok(a.note!.includes("x.example"));
    const b = customProvider({ id: CUSTOM_PREFIX + "b", name: "", baseUrl: "https://y.example", models: ["m"], anthropicHeaders: true })!;
    assert.deepStrictEqual(lower(headers(buildStepRequest(b, "K", "m", HISTORY, []))), ["anthropic-version", "content-type", "x-api-key"]);
    assert.strictEqual(b.name, "y.example");
  });

  it("nguồn tự khai sai địa chỉ hoặc sai tiền tố mã thì bị bỏ, không đè được nguồn có sẵn", () => {
    assert.strictEqual(customProvider({ id: CUSTOM_PREFIX + "x", name: "x", baseUrl: "http://x.example", models: ["m"] }), undefined);
    assert.strictEqual(customProvider({ id: "anthropic", name: "giả", baseUrl: "https://x.example", models: ["m"] }), undefined);
    setCustomSources([{ id: "anthropic", name: "giả", baseUrl: "https://x.example", models: ["m"] }] as any);
    assert.strictEqual(getProvider("anthropic")!.baseUrl, "https://api.anthropic.com/v1");
    assert.strictEqual(allProviders().length, PROVIDERS.length);
  });

  it("lưu và đọc lại qua kho; key của nguồn tự khai lưu như mọi nguồn", async () => {
    const store = memStore();
    const st = new Storage(store);
    const id = CUSTOM_PREFIX + "sf";
    st.setCustomSources([{ id, name: "SiliconFlow", baseUrl: "https://api.siliconflow.com/v1", models: ["deepseek-ai/DeepSeek-V4"] }]);
    st.setKey(id, "sk-sf");
    const again = new Storage(store);
    assert.strictEqual(again.getKeys()[id], "sk-sf");
    const p = getProvider(id)!;
    assert.strictEqual(p.custom, true);
    assert.deepStrictEqual((await fetchModels(p, "sk-sf")).map((m) => m.id), ["deepseek-ai/DeepSeek-V4"]);
    again.setCustomSources([]);
    assert.strictEqual(getProvider(id), undefined);
  });
});
