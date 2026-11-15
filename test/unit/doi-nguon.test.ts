/** Đợt 0A – bước 4: đổi nhà cung cấp giữa hội thoại.
 *  Bảng thử lấy từ tasks/ra-soat-2026-11-15/ma-tran-doi-nha-cung-cap.txt (10 kiểu lịch sử × 11 nguồn đích). */
import * as assert from "assert";
import { PROVIDERS, getProvider } from "../../src/core/providers";
import { buildStepRequest, isToolUnsupportedError } from "../../src/core/agent/llm";
import { prepareHistory, remapCallId } from "../../src/core/agent/history";
import { AgentMessage, ToolSpec } from "../../src/core/agent/types";
import { ApiError, cleanKey, describeError, keyProblem } from "../../src/core/errors";
import { normalizeModels } from "../../src/core/models";
import { applyFailed, applyLoaded, filterModels, pickModel, sendBlockReason, startCatalog } from "../../src/core/catalog";
import { Storage, KVStore } from "../../src/core/storage";
import { checkProvider, checkSwitch } from "../../src/core/providerCheck";

const P = (id: string) => getProvider(id)!;
const TOOLS: ToolSpec[] = [
  { name: "excel_get_overview", description: "o", parameters: { type: "object", properties: {} } },
  { name: "excel_read_range", description: "r", parameters: { type: "object", properties: { address: { type: "string" } } } },
];

/** Một lượt: hỏi → model gọi 2 công cụ → 2 kết quả → trả lời. */
function turn(ids: [string, string], opts: { content?: string; img?: boolean; raw?: boolean; native?: boolean; endWithTool?: boolean } = {}): AgentMessage[] {
  const user: AgentMessage = opts.img
    ? { role: "user", content: "Xem ảnh", images: [{ mime: "image/png", data: "iVBORw0KGgo=" }] }
    : { role: "user", content: "Tổng quan sổ tính" };
  const asst: AgentMessage = {
    role: "assistant",
    content: opts.content ?? "",
    toolCalls: [
      { id: ids[0], name: "excel_get_overview", args: {}, nativeId: opts.native },
      { id: ids[1], name: "excel_read_range", args: { address: "A1" }, nativeId: opts.native },
    ],
    ...(opts.raw
      ? {
          raw: {
            kind: "gemini",
            parts: [
              { functionCall: { name: "excel_get_overview", args: {}, ...(opts.native ? { id: ids[0] } : {}) }, thoughtSignature: "SIG" },
              { functionCall: { name: "excel_read_range", args: { address: "A1" }, ...(opts.native ? { id: ids[1] } : {}) } },
            ],
          },
        }
      : {}),
  };
  const out: AgentMessage[] = [
    user,
    asst,
    { role: "tool", toolCallId: ids[0], name: "excel_get_overview", content: '{"sheets":["S1"]}' },
    { role: "tool", toolCallId: ids[1], name: "excel_read_range", content: '{"values":[[42]]}' },
  ];
  if (!opts.endWithTool) out.push({ role: "assistant", content: "Ô A1 = 42.", ...(opts.raw ? { raw: { kind: "gemini", parts: [{ text: "Ô A1 = 42." }] } } : {}) });
  return out;
}

const SHAPES: Record<string, AgentMessage[]> = {
  "OpenAI (call_+24, content rỗng)": turn(["call_AAAAAAAAAAAAAAAAAAAAAAAA", "call_BBBBBBBBBBBBBBBBBBBBBBBB"]),
  "Anthropic (toolu_01…)": turn(["toolu_01A09q90qw90lq917835lq9", "toolu_01B09q90qw90lq917835lq9"], { content: "Để tôi xem." }),
  "Gemini 2.5 (không id gốc, raw.parts)": turn(["mM8r35lyu", "MGYGHLAf5"], { raw: true }),
  "Gemini 3 (id gốc, raw.parts)": turn(["fc_gem_0001", "fc_gem_0002"], { raw: true, native: true }),
  "Mistral (9 ký tự)": turn(["D681PevKs", "aB3dE6gH9"]),
  "OpenRouter→Kimi (functions.ten:0)": turn(["functions.excel_get_overview:0", "functions.excel_read_range:1"]),
  "OpenRouter/vLLM (chatcmpl-tool-, chữ trắng)": turn(
    ["chatcmpl-tool-2f96fb218bfb3e3f459161a11397960c", "chatcmpl-tool-32379b976be27706ad1e07bf5e0c27d7"],
    { content: "\n\n" }
  ),
  "Groq (fc_+uuid)": turn(["fc_3bd6b390-5f8c-2238-a55a-e7e1ed5350dc", "fc_35865389-c217-da22-39f2-b3c1ed1db574"]),
  "OpenAI có ảnh ở lượt đầu": turn(["call_CCCCCCCCCCCCCCCCCCCCCCCC", "call_DDDDDDDDDDDDDDDDDDDDDDDD"], { img: true }),
  "OpenAI chạm giới hạn bước (kết thúc bằng tool)": turn(["call_EEEEEEEEEEEEEEEEEEEEEEEE", "call_FFFFFFFFFFFFFFFFFFFFFFFF"], { endWithTool: true }),
};

const NEW_Q: AgentMessage = { role: "user", content: "Câu hỏi mới ở nguồn khác" };

/** Kiểm yêu cầu dựng cho nguồn đích theo luật của từng nhà cung cấp; trả về danh sách vi phạm. */
function violations(pid: string, body: any, vision?: boolean): string[] {
  const p = P(pid);
  const bad: string[] = [];
  if (p.kind === "anthropic") {
    const useIds: string[] = [];
    body.messages.forEach((m: any, i: number) => {
      for (const b of m.content) {
        if (b.type === "text" && !String(b.text).trim()) bad.push(`messages.${i}: khối chữ trắng`);
        if (b.type === "tool_use") {
          if (!/^[a-zA-Z0-9_-]+$/.test(b.id)) bad.push(`tool_use.id sai: ${b.id}`);
          useIds.push(b.id);
          const next = body.messages[i + 1];
          if (!next || !next.content.some((x: any) => x.type === "tool_result" && x.tool_use_id === b.id)) bad.push(`tool_use ${b.id} thiếu kết quả ngay sau`);
        }
        if (b.type === "tool_result" && !useIds.includes(b.tool_use_id)) bad.push(`tool_result mồ côi ${b.tool_use_id}`);
      }
    });
    if (body.messages[0]?.role !== "user") bad.push("lượt đầu không phải user");
  } else if (p.kind === "gemini") {
    const contents = body.contents;
    contents.forEach((c: any, i: number) => {
      for (const part of c.parts) {
        if (part.functionCall) {
          const next = contents[i + 1];
          if (!next || !next.parts.some((x: any) => x.functionResponse?.name === part.functionCall.name)) bad.push(`functionCall ${part.functionCall.name} thiếu kết quả`);
        }
      }
    });
  } else {
    const msgs = body.messages;
    const callIds: string[] = [];
    msgs.forEach((m: any, i: number) => {
      for (const tc of m.tool_calls || []) {
        callIds.push(tc.id);
        if (tc.id.length > 40) bad.push(`tool_calls.id dài ${tc.id.length}`);
        if (pid === "mistral" && !/^[a-zA-Z0-9]{9}$/.test(tc.id)) bad.push(`Mistral id sai: ${tc.id}`);
        if (pid !== "openrouter" && !/^[a-zA-Z0-9_-]+$/.test(tc.id)) bad.push(`id có ký tự lạ: ${tc.id}`);
      }
      if (m.role === "tool") {
        if (!callIds.includes(m.tool_call_id)) bad.push(`tool mồ côi ${m.tool_call_id}`);
        if (pid === "mistral" && !/^[a-zA-Z0-9]{9}$/.test(m.tool_call_id)) bad.push(`Mistral tool_call_id sai`);
        if (pid === "mistral" && msgs[i + 1]?.role === "user") bad.push("Mistral: user ngay sau tool");
      }
      if (m.role === "assistant" && m.tool_calls) {
        for (const tc of m.tool_calls) {
          const has = msgs.slice(i + 1).some((x: any) => x.role === "tool" && x.tool_call_id === tc.id);
          if (!has) bad.push(`tool_call ${tc.id} thiếu kết quả`);
        }
      }
    });
    const userImgs = msgs.filter((m: any) => m.role === "user" && Array.isArray(m.content) && m.content.some((c: any) => c.type === "image_url"));
    if (vision === false && userImgs.length) bad.push("gửi ảnh cho model không nhận ảnh");
    if (vision === undefined && userImgs.some((m: any) => m !== msgs[msgs.length - 1])) bad.push("gửi lại ảnh cũ cho model chưa rõ có nhận ảnh");
  }
  return bad;
}

describe("Đổi nguồn: chuẩn hóa lịch sử cho nguồn đích (10 kiểu × 11 nguồn)", () => {
  for (const [shape, hist] of Object.entries(SHAPES)) {
    it(`${shape} → mọi nguồn đích: không vi phạm luật của nguồn đích`, () => {
      const fails: string[] = [];
      for (const p of PROVIDERS) {
        const vision = p.id === "deepseek" ? false : undefined;
        const r = buildStepRequest(p, "K", "m", [{ role: "system", content: "S" }, ...hist, NEW_Q], TOOLS, { vision });
        const v = violations(p.id, JSON.parse(r.init.body as string), vision);
        if (v.length) fails.push(`${p.id}: ${v.join("; ")}`);
      }
      assert.deepStrictEqual(fails, []);
    });
  }

  it("Đổi mã giữ đúng cặp gọi – kết quả; mã mới ổn định (cùng mã gốc → cùng mã mới)", () => {
    const h = SHAPES["OpenAI (call_+24, content rỗng)"];
    const a = prepareHistory([...h, NEW_Q], { provider: P("mistral") });
    const b = prepareHistory([...h, NEW_Q], { provider: P("mistral") });
    const calls = (a[1] as any).toolCalls.map((c: any) => c.id);
    assert.deepStrictEqual(calls, (b[1] as any).toolCalls.map((c: any) => c.id));
    assert.deepStrictEqual([(a[2] as any).toolCallId, (a[3] as any).toolCallId], calls);
    assert.notStrictEqual(calls[0], calls[1]);
    assert.ok(calls.every((id: string) => /^[a-zA-Z0-9]{9}$/.test(id)));
    assert.strictEqual(remapCallId("call_X", new Set()), remapCallId("call_X", new Set()));
  });

  it("Không đụng tới lịch sử gốc (chỉ đổi trong yêu cầu gửi đi)", () => {
    const h = SHAPES["OpenRouter→Kimi (functions.ten:0)"];
    const before = JSON.stringify(h);
    prepareHistory(h, { provider: P("anthropic") });
    assert.strictEqual(JSON.stringify(h), before);
  });

  it("Mã hợp lệ thì giữ nguyên; Gemini giữ mã gốc khi quay lại Gemini", () => {
    const h = SHAPES["Mistral (9 ký tự)"];
    const out = prepareHistory(h, { provider: P("anthropic") });
    assert.strictEqual((out[1] as any).toolCalls[0].id, "D681PevKs");
    const g = SHAPES["Gemini 3 (id gốc, raw.parts)"];
    const r = buildStepRequest(P("gemini"), "K", "m", [...g, NEW_Q], TOOLS);
    const body = JSON.parse(r.init.body as string);
    const resp = body.contents.find((c: any) => c.parts.some((x: any) => x.functionResponse));
    assert.strictEqual(resp.parts[0].functionResponse.id, "fc_gem_0001");
  });

  it("Lệnh gọi bị dừng giữa chừng (không có kết quả) → thêm kết quả thay thế; kết quả mồ côi → bỏ", () => {
    const h: AgentMessage[] = [
      { role: "user", content: "a" },
      { role: "assistant", content: "", toolCalls: [{ id: "x1", name: "excel_get_overview", args: {} }] },
      { role: "tool", toolCallId: "zz", name: "excel_get_overview", content: "{}" },
      NEW_Q,
    ];
    const out = prepareHistory(h, { provider: P("openai") });
    const tools = out.filter((m) => m.role === "tool") as any[];
    assert.strictEqual(tools.length, 1);
    assert.strictEqual(tools[0].toolCallId, "x1");
    assert.match(tools[0].content, /không có kết quả/);
  });

  it("Dữ liệu cũ lệch chuẩn (content/args undefined) không làm sập", () => {
    const h: any[] = [
      { role: "user", content: undefined },
      { role: "assistant", content: undefined, toolCalls: [{ id: "t1", name: "excel_get_overview", args: undefined }] },
      { role: "tool", toolCallId: "t1", name: "excel_get_overview", content: undefined },
      NEW_Q,
    ];
    for (const p of PROVIDERS) {
      const r = buildStepRequest(p, "K", "m", h, TOOLS);
      assert.deepStrictEqual(violations(p.id, JSON.parse(r.init.body as string)), [], p.id);
    }
  });

  it("Ảnh: model không nhận ảnh → bỏ ảnh, ghi chú thay thế; model nhận ảnh → giữ", () => {
    const h = [...SHAPES["OpenAI có ảnh ở lượt đầu"], NEW_Q];
    const no = prepareHistory(h, { provider: P("groq"), vision: false });
    assert.ok(!(no[0] as any).images);
    assert.match((no[0] as any).content, /ảnh/);
    const yes = prepareHistory(h, { provider: P("openrouter"), vision: true });
    assert.strictEqual((yes[0] as any).images.length, 1);
  });
});

describe("Đổi nguồn: thông báo lỗi đúng nguyên nhân", () => {
  it("Chỉ lỗi fetch thật mới là 'không kết nối được'; nguồn chặn add-in nêu rõ lý do", () => {
    assert.strictEqual(describeError(new TypeError("Failed to fetch")).kind, "network");
    assert.strictEqual(describeError(new TypeError("network error")).kind, "network");
    const s = describeError(new TypeError("Failed to fetch"), { provider: P("sambanova") });
    assert.match(s.text, /SambaNova/);
    assert.match(s.text, /chặn/);
  });
  it("Lỗi lập trình hiện nội dung thật, không đổ cho mạng", () => {
    const d = describeError(new TypeError("Cannot read properties of undefined (reading 'length')"));
    assert.strictEqual(d.kind, "app");
    assert.match(d.text, /reading 'length'/);
    assert.doesNotMatch(d.text, /Không kết nối/);
    assert.match(d.detail, /TypeError/);
  });
  it("Key có ký tự lạ → báo dán lại key", () => {
    const d = describeError(new TypeError("Cannot convert argument to a ByteString because the character at index 13 has a value of 8203"));
    assert.strictEqual(d.kind, "key");
    assert.match(d.text, /dán lại key/);
    assert.strictEqual(describeError(new TypeError('Headers.append: "Bearer sk-abc\ndef" is an invalid header value.')).kind, "key");
  });
  it("Làm sạch key khi lưu: bỏ ký tự ẩn, khoảng trắng, dấu nháy; báo key có chữ lạ", () => {
    assert.strictEqual(cleanKey(" ​sk-ab​c﻿\n"), "sk-abc");
    assert.strictEqual(cleanKey('"sk-abc"'), "sk-abc");
    assert.strictEqual(keyProblem("sk-abc"), null);
    assert.ok(keyProblem("sk-abđ"));
    assert.ok(keyProblem("sk-a bc"));
    const store: any = { d: {} as Record<string, string>, getItem(k: string) { return this.d[k] ?? null; }, setItem(k: string, v: string) { this.d[k] = v; }, removeItem(k: string) { delete this.d[k]; } };
    const s = new Storage(store as KVStore);
    s.setKey("openai", "​sk-1​ ");
    assert.strictEqual(s.getKey("openai"), "sk-1");
    store.d["aidn.keys"] = JSON.stringify({ groq: "gsk_1​" });
    assert.strictEqual(s.getKey("groq"), "gsk_1");
  });
  it("Lỗi API về mã lời gọi công cụ không bị nhận nhầm là 'model không gọi được công cụ'", () => {
    assert.ok(!isToolUnsupportedError(new ApiError("Lỗi 400 (Tool call id has to be a string of length 9)", 400)));
    assert.ok(!isToolUnsupportedError(new ApiError("Lỗi 400 (messages.1.content.0.tool_use.id: String should match pattern '^[a-zA-Z0-9_-]+$')", 400)));
    assert.ok(!isToolUnsupportedError(new ApiError("Lỗi 400 (Invalid 'messages[2].tool_calls[0].id': string too long)", 400)));
    assert.ok(isToolUnsupportedError(new ApiError("Lỗi 404 (No endpoints found that support tool use)", 404)));
    assert.ok(isToolUnsupportedError(new ApiError("Lỗi 400 (model does not support tools)", 400)));
  });
});

const M = (id: string, extra: any = {}) => ({ id, name: id, ...extra });

describe("Đổi nguồn: trạng thái danh mục model (tách khỏi Chat.tsx)", () => {
  it("Đổi nguồn: xóa ngay danh sách cũ; chưa có bộ nhớ đệm thì 'đang tải'", () => {
    const s = startCatalog("anthropic", undefined, 2);
    assert.strictEqual(s.status, "loading");
    assert.deepStrictEqual(s.models, []);
    const c = startCatalog("groq", { at: 1, models: [M("llama")] }, 3);
    assert.strictEqual(c.status, "ready");
    assert.strictEqual(c.models[0].id, "llama");
  });
  it("Bỏ kết quả tải về muộn của nguồn khác hoặc lần tải cũ", () => {
    const s = startCatalog("anthropic", undefined, 5);
    assert.strictEqual(applyLoaded(s, "openrouter", 4, [M("x")], 1), s);
    assert.strictEqual(applyLoaded(s, "anthropic", 4, [M("x")], 1), s);
    const ok = applyLoaded(s, "anthropic", 5, [M("claude")], 1);
    assert.strictEqual(ok.status, "ready");
    const f = applyFailed(s, "anthropic", 5, "Key không đúng");
    assert.strictEqual(f.status, "error");
    assert.strictEqual(f.error, "Key không đúng");
  });
  it("Lọc 'Miễn phí' và 'Chỉ model thao tác được' chỉ áp dụng cho OpenRouter (lỗi tái hiện 15/11: Anthropic hiện 0/15)", () => {
    const ms = [M("claude-a"), M("claude-b")];
    assert.strictEqual(filterModels(ms, { providerId: "anthropic", freeOnly: true, toolsOnly: true, agentOn: true, search: "" }).length, 2);
    const or = [M("a:free", { free: true, tools: true }), M("b", { tools: true }), M("c:free", { free: true, tools: false })];
    assert.deepStrictEqual(filterModels(or, { providerId: "openrouter", freeOnly: true, toolsOnly: true, agentOn: true, search: "" }).map((m) => m.id), ["a:free"]);
  });
  it("Chọn model: nhớ theo từng nguồn; nguồn mới chọn model gọi được công cụ, không lấy mục đầu theo chữ cái", () => {
    const ms = [M("aaa-guard", { tools: false }), M("claude-haiku-4-5", { created: 3 }), M("claude-sonnet-4-5", { created: 2 }), M("claude-sonnet-4-6", { created: 4 })];
    assert.strictEqual(pickModel(ms, { providerId: "anthropic", remembered: "claude-haiku-4-5", agentOn: true, freeOnly: false }), "claude-haiku-4-5");
    assert.strictEqual(pickModel(ms, { providerId: "anthropic", remembered: "khong-con", agentOn: true, freeOnly: false }), "claude-sonnet-4-6");
    const or = [M("a/x", { tools: false, free: true }), M("b/y:free", { tools: true, free: true }), M("c/z", { tools: true })];
    assert.strictEqual(pickModel(or, { providerId: "openrouter", agentOn: true, freeOnly: true }), "b/y:free");
    assert.strictEqual(pickModel([], { providerId: "groq", agentOn: true, freeOnly: false }), "");
  });
  it("Không cho gửi khi đang tải, khi lỗi, khi model không thuộc nguồn đang chọn", () => {
    const loading = startCatalog("anthropic", undefined, 1);
    assert.match(sendBlockReason({ provider: P("anthropic"), catalog: loading, modelId: "x" })!, /đang tải/i);
    const err = applyFailed(loading, "anthropic", 1, "Key không đúng");
    assert.match(sendBlockReason({ provider: P("anthropic"), catalog: err, modelId: "x" })!, /Key không đúng/);
    const ready = applyLoaded(loading, "anthropic", 1, [M("claude")], 1);
    assert.match(sendBlockReason({ provider: P("anthropic"), catalog: ready, modelId: "nvidia/nemotron:free" })!, /chọn model/i);
    assert.strictEqual(sendBlockReason({ provider: P("anthropic"), catalog: ready, modelId: "claude" }), null);
  });
  it("Nhớ model theo nguồn trong bộ nhớ máy; chuyển dữ liệu cũ (một model chung) sang dạng mới", () => {
    const d: Record<string, string> = { "aidn.selection": JSON.stringify({ providerId: "groq", modelId: "llama", freeOnly: true }) };
    const store: KVStore = { getItem: (k) => d[k] ?? null, setItem: (k, v) => void (d[k] = v), removeItem: (k) => void delete d[k] };
    const s = new Storage(store);
    assert.strictEqual(s.getModelFor("groq"), "llama");
    s.setModelFor("anthropic", "claude-sonnet-4-6");
    s.setModelFor("groq", "gpt-oss");
    assert.strictEqual(s.getModelFor("anthropic"), "claude-sonnet-4-6");
    assert.strictEqual(new Storage(store).getModelFor("groq"), "gpt-oss");
  });
  it("Danh mục ghi thời điểm tạo model (để chọn bản mới nhất) và khả năng nhận ảnh khi nhà cung cấp có báo", () => {
    const or = normalizeModels(P("openrouter"), {
      data: [{ id: "a/b", name: "B", created: 100, supported_parameters: ["tools"], architecture: { input_modalities: ["text", "image"] } }],
    });
    assert.strictEqual(or[0].created, 100);
    assert.strictEqual(or[0].vision, true);
    const an = normalizeModels(P("anthropic"), { data: [{ id: "claude-x", display_name: "X", created_at: "2026-01-01T00:00:00Z" }] });
    assert.ok(an[0].created! > 0);
  });
});

/** Máy giả lập nhà cung cấp kiểu OpenAI: lượt 1 gọi công cụ, lượt 2 trả lời theo kết quả công cụ. */
function fakeOpenAI(answer: (body: any) => string, opts: { toolFirst?: boolean } = {}): typeof fetch {
  return (async (_url: any, init: any) => {
    const body = JSON.parse(init.body);
    const hasToolResult = body.messages.some((m: any) => m.role === "tool");
    const last = body.messages[body.messages.length - 1];
    const lines =
      opts.toolFirst !== false && !hasToolResult && last.role === "user"
        ? [{ choices: [{ delta: { tool_calls: [{ index: 0, id: "call_1", function: { name: body.tools[0].function.name, arguments: "{}" } }] } }] }]
        : [{ choices: [{ delta: { content: answer(body) } }] }];
    const text = lines.map((l) => `data: ${JSON.stringify(l)}\n\n`).join("") + "data: [DONE]\n\n";
    return new Response(text, { status: 200 });
  }) as any;
}

describe("Nút Kiểm tra nhà cung cấp", () => {
  it("Chạy thật chuỗi gọi công cụ → nhận kết quả → trả lời có chứa kết quả: Đạt", async () => {
    const f = fakeOpenAI((body) => {
      const tool = body.messages.find((m: any) => m.role === "tool");
      return `Mã là ${JSON.parse(tool.content).ma}`;
    });
    const r = await checkProvider({ provider: P("groq"), key: "K", model: "m", fetchFn: f, code: "482915" });
    assert.strictEqual(r.ok, true, r.message);
    assert.ok(r.history && r.history.length >= 4);
  });
  it("Model trả lời không có mã / không gọi công cụ → Không đạt kèm lý do tiếng Việt", async () => {
    const f = fakeOpenAI(() => "Tôi không biết", { toolFirst: false });
    const r = await checkProvider({ provider: P("groq"), key: "K", model: "m", fetchFn: f, code: "482915" });
    assert.strictEqual(r.ok, false);
    assert.match(r.message, /không gọi công cụ/);
  });
  it("Lỗi HTTP → lý do tiếng Việt", async () => {
    const f = (async () => new Response('{"error":{"message":"invalid key"}}', { status: 401 })) as any;
    const r = await checkProvider({ provider: P("groq"), key: "K", model: "m", fetchFn: f, code: "1" });
    assert.strictEqual(r.ok, false);
    assert.match(r.message, /Key không đúng/);
  });
  it("Đổi nguồn: nguồn B đọc lịch sử của nguồn A và nhắc lại đúng mã", async () => {
    const a = await checkProvider({
      provider: P("openrouter"),
      key: "K",
      model: "m",
      code: "777123",
      fetchFn: fakeOpenAI((b) => `Mã là ${JSON.parse(b.messages.find((m: any) => m.role === "tool").content).ma}`),
    });
    assert.ok(a.ok, a.message);
    let seen: any;
    const fb = fakeOpenAI(
      (body) => {
        seen = body;
        const tool = body.messages.find((m: any) => m.role === "tool");
        return JSON.parse(tool.content).ma;
      },
      { toolFirst: false }
    );
    const r = await checkSwitch({ history: a.history!, code: "777123", to: { provider: P("mistral"), key: "K", model: "m" }, fetchFn: fb });
    assert.strictEqual(r.ok, true, r.message);
    const ids = seen.messages.flatMap((m: any) => (m.tool_calls || []).map((t: any) => t.id));
    assert.ok(ids.every((id: string) => /^[a-zA-Z0-9]{9}$/.test(id)), "Mistral nhận mã 9 ký tự");
  });
});

import { fetchDiagnosed } from "../../src/core/errors";
import { callModelStep } from "../../src/core/agent/llm";

describe("Đổi nguồn: máy chủ trả lỗi nhưng giấu chi tiết (thiếu CORS trên trang lỗi — OpenAI key sk-…, Cerebras; đo 15/11)", () => {
  it("fetch lỗi nhưng máy chủ vẫn tới được → báo 'trả về lỗi nhưng không cho add-in đọc chi tiết', không báo mất mạng", async () => {
    const calls: any[] = [];
    const f = (async (url: string, init: any) => {
      calls.push({ url, mode: init?.mode });
      if (init?.mode === "no-cors") return new Response(null, { status: 200 });
      throw new TypeError("Failed to fetch");
    }) as any;
    await assert.rejects(
      callModelStep({ provider: P("openai"), key: "sk-x", model: "gpt-5.5", messages: [{ role: "user", content: "x" }], tools: [], onText: () => {}, fetchFn: f }),
      (e: any) => {
        assert.match(e.message, /không cho add-in đọc chi tiết/);
        assert.match(e.message, /gpt-5\.5/);
        assert.strictEqual(describeError(e).kind, "api");
        return true;
      }
    );
    assert.strictEqual(calls[1].mode, "no-cors");
    assert.strictEqual(calls[1].url, "https://api.openai.com/");
  });
  it("Không tới được máy chủ (mất mạng) → vẫn là lỗi kết nối", async () => {
    const f = (async () => {
      throw new TypeError("Failed to fetch");
    }) as any;
    await assert.rejects(fetchDiagnosed(P("cerebras"), "https://api.cerebras.ai/v1/models", {}, f), (e: any) => describeError(e).kind === "network");
  });
  it("Nguồn đã biết bị chặn (SambaNova) → không thử lại, báo lý do bị chặn", async () => {
    let n = 0;
    const f = (async () => {
      n++;
      throw new TypeError("Failed to fetch");
    }) as any;
    await assert.rejects(fetchDiagnosed(P("sambanova"), "https://api.sambanova.ai/v1/models", {}, f), (e: any) =>
      /chặn/.test(describeError(e, { provider: P("sambanova") }).text)
    );
    assert.strictEqual(n, 1);
  });
});
