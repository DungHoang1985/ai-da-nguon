/** Đợt 0B: lịch sử chỉ thêm, tự thử lại, lệnh bị cắt dở, bộ nhớ đệm Anthropic, đếm token, ngưỡng tóm tắt, an toàn. */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { anthropicMaxTokens, buildStepRequest, callModelStep, createStepAccumulator } from "../../src/core/agent/llm";
import { NOT_RUN_STOPPED, NOT_RUN_TRUNCATED, STOP_NOTE, runAgent, withTimeout } from "../../src/core/agent/loop";
import {
  COMPACT_AT_CHARS,
  MAX_COMPACT_CHARS,
  MIN_COMPACT_CHARS,
  SUMMARY_PREFIX,
  compactIfNeeded,
  compactLimitChars,
  needsCompaction,
  transcriptOf,
} from "../../src/core/agent/compaction";
import { SAFETY_RULES, buildSystemPrompt, buildTurnMessage, stripTurnContext } from "../../src/core/agent/prompt";
import { addUsage, formatUsage, runTurn } from "../../src/core/agent/turn";
import { AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { ApiError, isRetryable, messageForStatus } from "../../src/core/errors";
import { normalizeModels } from "../../src/core/models";
import { pickModel } from "../../src/core/catalog";
import { planTableFill } from "../../src/office/util";

const P = (id: string) => getProvider(id)!;

function sse(events: any[], init: ResponseInit = { status: 200 }): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  return new Response(text, init);
}
const oaText = (t: string, extra: any = {}) => sse([{ choices: [{ delta: { content: t }, ...extra }] }]);
const oaCall = (id: string, name: string, args: string, finish?: string) =>
  sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: args } }] }, ...(finish ? { finish_reason: finish } : {}) }] }]);
const overloaded = () => new Response('{"error":{"type":"overloaded_error","message":"Overloaded"}}', { status: 529 });

const TOOLS: ToolSpec[] = [
  { name: "doc_read", description: "đọc", parameters: { type: "object", properties: {} } },
  { name: "doc_write", description: "ghi", write: true, parameters: { type: "object", properties: { text: { type: "string" } } } },
];

function fakeTools(done: string[], state = { n: 0 }): HostToolSet {
  return {
    host: "Word",
    specs: TOOLS,
    overview: async () => ({ doan: ++state.n }),
    systemPrompt: () => "LOI NHAC CO DINH",
    stateText: (ov) => `Trạng thái: ${JSON.stringify(ov)}`,
    execute: async (name, args, { log }) => {
      done.push(`${name}:${args?.text ?? ""}`);
      if (name === "doc_write") log.push({ label: "ghi", undo: async () => {} });
      return { ok: true, message: "xong" };
    },
  };
}

const noSleep = async () => {};

describe("Đợt 0B: tự thử lại khi nhà cung cấp quá tải", () => {
  it("Quá tải hai lần rồi thành công: tự thử lại hai lần, báo 'lần 1', 'lần 2', kết quả đúng", async () => {
    const responses = [overloaded(), overloaded(), oaText("Xong.")];
    const retries: string[] = [];
    const waits: number[] = [];
    const res = await callModelStep({
      provider: P("groq"),
      key: "K",
      model: "m",
      messages: [{ role: "user", content: "x" }],
      tools: [],
      onText: () => {},
      fetchFn: (async () => responses.shift()) as any,
      onRetry: (r) => retries.push(`lần ${r.attempt}/${r.max}: ${r.reason}`),
      sleep: async (ms) => {
        waits.push(ms);
      },
    });
    assert.strictEqual(res.text, "Xong.");
    assert.deepStrictEqual(retries, ["lần 1/3: nhà cung cấp quá tải", "lần 2/3: nhà cung cấp quá tải"]);
    assert.deepStrictEqual(waits, [2000, 6000]);
  });
  it("Trong vòng lặp tác tử: phát sự kiện retry để khung chat hiện 'đang thử lại lần n'", async () => {
    const responses = [new Response('{"error":{"message":"rate limit"}}', { status: 429 }), oaText("Chào.")];
    const events: any[] = [];
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools([]),
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      onEvent: (e) => events.push(e),
      fetchFn: (async () => responses.shift()) as any,
      sleep: noSleep,
    });
    const r = events.find((e) => e.type === "retry");
    assert.ok(r && r.attempt === 1 && r.reason === "hết hạn mức tạm thời", JSON.stringify(r));
    assert.strictEqual((msgs[msgs.length - 1] as any).content, "Chào.");
  });
  it("Bấm Dừng trong lúc chờ thử lại: hủy được, không gọi thêm", async () => {
    const ac = new AbortController();
    let calls = 0;
    const p = callModelStep({
      provider: P("groq"),
      key: "K",
      model: "m",
      messages: [{ role: "user", content: "x" }],
      tools: [],
      onText: () => {},
      signal: ac.signal,
      fetchFn: (async () => {
        calls++;
        return overloaded();
      }) as any,
      onRetry: () => setTimeout(() => ac.abort(), 5),
      retryDelaysMs: [60000],
    });
    await assert.rejects(p, (e: any) => e.name === "AbortError");
    assert.strictEqual(calls, 1);
  });
  it("Quá số lần thử thì báo lỗi; lỗi 400/401 không thử lại; chờ quá 60 giây thì không tự chờ", async () => {
    let calls = 0;
    const always = (r: () => Response) =>
      (async () => {
        calls++;
        return r();
      }) as any;
    const base = { provider: P("groq"), key: "K", model: "m", messages: [{ role: "user", content: "x" }] as AgentMessage[], tools: [], onText: () => {}, sleep: noSleep };
    await assert.rejects(callModelStep({ ...base, fetchFn: always(overloaded) }), /quá tải/);
    assert.strictEqual(calls, 4, "1 lần gọi + 3 lần thử lại");
    calls = 0;
    await assert.rejects(callModelStep({ ...base, fetchFn: always(() => new Response("{}", { status: 401 })) }), /Key không đúng/);
    assert.strictEqual(calls, 1);
    calls = 0;
    await assert.rejects(
      callModelStep({ ...base, fetchFn: always(() => new Response("{}", { status: 429, headers: { "retry-after": "3600" } })) }),
      /Hết hạn mức/
    );
    assert.strictEqual(calls, 1);
  });
  it("Đã hiện chữ cho người dùng rồi mới lỗi giữa luồng: không thử lại (tránh lặp chữ)", async () => {
    let calls = 0;
    const f = (async () => {
      calls++;
      return sse([{ choices: [{ delta: { content: "Một nửa" } }] }, { error: { type: "overloaded_error", message: "Overloaded" } }]);
    }) as any;
    await assert.rejects(
      callModelStep({ provider: P("groq"), key: "K", model: "m", messages: [{ role: "user", content: "x" }], tools: [], onText: () => {}, fetchFn: f, sleep: noSleep }),
      /Overloaded/
    );
    assert.strictEqual(calls, 1);
  });
  it("Phân loại lỗi tạm thời", () => {
    assert.ok(isRetryable(new ApiError("x", 429)) && isRetryable(new ApiError("x", 529)) && isRetryable(new ApiError("x", 503)));
    assert.ok(!isRetryable(new ApiError("x", 400)) && !isRetryable(new ApiError("x")));
    // Lỗi mạng thô (đã qua fetchDiagnosed, máy chủ không với tới được): từ 4.7.2 được thử lại — xem dot7-de-lai.test.ts
  });
});

describe("Đợt 0B: lệnh công cụ bị cắt dở không được thực thi", () => {
  it("OpenAI finish_reason=length kèm lệnh gọi: không chạy, báo model chia nhỏ, rồi lượt sau chạy bình thường", async () => {
    const responses = [oaCall("c1", "doc_write", '{"text":"rất dài bị c', "length"), oaCall("c2", "doc_write", '{"text":"ngắn"}'), oaText("Đã ghi.")];
    const bodies: any[] = [];
    const done: string[] = [];
    const events: any[] = [];
    const log: any[] = [];
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools(done),
      system: "S",
      messages: [{ role: "user", content: "viết" }],
      log,
      onEvent: (e) => events.push(e),
      fetchFn: (async (_u: string, init: any) => {
        bodies.push(JSON.parse(init.body));
        return responses.shift();
      }) as any,
    });
    assert.deepStrictEqual(done, ["doc_write:ngắn"], "lệnh bị cắt dở không được chạy");
    assert.strictEqual(log.length, 1);
    const t = events.find((e) => e.type === "truncated");
    assert.ok(t && t.skipped === 1 && !t.gaveUp);
    const firstResult = bodies[1].messages.find((m: any) => m.role === "tool");
    assert.ok(firstResult.content.includes("chia nhỏ"), firstResult.content);
    assert.strictEqual((msgs[msgs.length - 1] as any).content, "Đã ghi.");
  });
  it("Anthropic stop_reason=max_tokens: nhận ra bị cắt; cắt hai lần liên tiếp thì dừng lượt", async () => {
    const acc = createStepAccumulator("anthropic", () => {});
    [
      { type: "message_start", message: { usage: { input_tokens: 10 } } },
      { type: "content_block_start", index: 0, content_block: { type: "tool_use", id: "t1", name: "doc_write" } },
      { type: "content_block_delta", index: 0, delta: { type: "input_json_delta", partial_json: '{"text":"dở' } },
      { type: "message_delta", delta: { stop_reason: "max_tokens" }, usage: { output_tokens: 99 } },
    ].forEach((e) => acc.push(JSON.stringify(e)));
    const r = acc.result();
    assert.strictEqual(r.truncated, true);
    assert.strictEqual(r.toolCalls.length, 1);

    const done: string[] = [];
    const events: any[] = [];
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools(done),
      system: "S",
      messages: [{ role: "user", content: "viết" }],
      log: [],
      onEvent: (e) => events.push(e),
      fetchFn: (async () => oaCall("c", "doc_write", '{"text":"dở', "length")) as any,
    });
    assert.deepStrictEqual(done, []);
    assert.strictEqual(events.filter((e) => e.type === "step").length, 2, "dừng sau hai lần bị cắt liên tiếp");
    assert.ok(events.some((e) => e.type === "truncated" && e.gaveUp));
    const toolResults = msgs.filter((m) => m.role === "tool");
    assert.ok(toolResults.length === 2 && toolResults.every((m: any) => m.content.includes(NOT_RUN_TRUNCATED.slice(0, 20))));
    assert.strictEqual(msgs[msgs.length - 1].role, "assistant");
  });
  it("Bị cắt khi chỉ có chữ: giữ phần chữ, báo sự kiện, kết thúc lượt", async () => {
    const events: any[] = [];
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools([]),
      system: "S",
      messages: [{ role: "user", content: "viết" }],
      log: [],
      onEvent: (e) => events.push(e),
      fetchFn: (async () => oaText("Nửa câu", { finish_reason: "length" })) as any,
    });
    assert.strictEqual((msgs[msgs.length - 1] as any).content, "Nửa câu");
    assert.ok(events.some((e) => e.type === "truncated" && e.skipped === 0));
  });
});

describe("Đợt 0B: lịch sử chỉ thêm, giữ phần đã làm khi dừng hoặc lỗi", () => {
  it("Bấm Dừng giữa chuỗi sửa: bước đã chạy còn kết quả, bước chưa chạy ghi 'chưa thực hiện', có mốc dừng", async () => {
    const ac = new AbortController();
    const done: string[] = [];
    let latest: AgentMessage[] = [];
    // Một bước model trả ba lệnh ghi; người dùng bấm Dừng ngay sau lệnh thứ nhất
    const three = sse([
      {
        choices: [
          {
            delta: {
              tool_calls: [
                { index: 0, id: "a1", function: { name: "doc_write", arguments: '{"text":"một"}' } },
                { index: 1, id: "a2", function: { name: "doc_write", arguments: '{"text":"hai"}' } },
                { index: 2, id: "a3", function: { name: "doc_write", arguments: '{"text":"ba"}' } },
              ],
            },
          },
        ],
      },
    ]);
    const log: any[] = [];
    await assert.rejects(
      runAgent({
        provider: P("groq"),
        key: "K",
        model: "m",
        tools: fakeTools(done),
        system: "S",
        messages: [{ role: "user", content: "sửa ba chỗ" }],
        log,
        signal: ac.signal,
        onHistory: (h) => (latest = h),
        onEvent: (e) => {
          if (e.type === "tool_end") ac.abort();
        },
        fetchFn: (async () => three) as any,
      }),
      (e: any) => e.name === "AbortError"
    );
    assert.deepStrictEqual(done, ["doc_write:một"]);
    assert.strictEqual(log.length, 1, "bước đã làm vẫn hoàn tác được");
    const tools = latest.filter((m) => m.role === "tool") as any[];
    assert.deepStrictEqual(
      tools.map((m) => m.toolCallId),
      ["a1", "a2", "a3"],
      "mọi lệnh gọi đều có kết quả"
    );
    assert.ok(tools[0].content.includes('"ok":true'));
    assert.ok(tools[1].content.includes(NOT_RUN_STOPPED) && tools[2].content.includes(NOT_RUN_STOPPED));
    assert.strictEqual((latest[latest.length - 1] as any).content, STOP_NOTE);
  });
  it("runTurn: bị dừng giữa chừng vẫn trả lịch sử có phần đã làm; lượt sau gửi đủ cho model", async () => {
    const ac = new AbortController();
    const done: string[] = [];
    const tools = fakeTools(done);
    const log: any[] = [];
    const r1 = await runTurn({
      provider: P("groq"),
      key: "K",
      model: "m",
      host: "Word",
      history: [],
      text: "sửa đoạn 1",
      agent: { tools, log },
      signal: ac.signal,
      onEvent: (e) => {
        if (e.type === "tool_end") ac.abort();
      },
      fetchFn: (async () => oaCall("w1", "doc_write", '{"text":"đoạn 1 mới"}')) as any,
    });
    assert.strictEqual((r1.error as any)?.name, "AbortError");
    assert.deepStrictEqual(done, ["doc_write:đoạn 1 mới"]);
    assert.strictEqual(log.length, 1);
    assert.ok(r1.history.length >= 4, "câu hỏi, lệnh gọi, kết quả, mốc dừng");

    let sent: any;
    const r2 = await runTurn({
      provider: P("groq"),
      key: "K",
      model: "m",
      host: "Word",
      history: r1.history,
      text: "anh vừa làm gì?",
      agent: { tools, log: [] },
      onEvent: () => {},
      fetchFn: (async (_u: string, init: any) => {
        sent = JSON.parse(init.body);
        return oaText("Tôi đã sửa đoạn 1.");
      }) as any,
    });
    assert.strictEqual(r2.error, undefined);
    const toolMsg = sent.messages.find((m: any) => m.role === "tool");
    assert.ok(toolMsg && toolMsg.content.includes('"ok":true'), "model thấy kết quả của bước đã làm");
    assert.ok(sent.messages.some((m: any) => m.role === "assistant" && String(m.content).includes("bấm Dừng")));
  });
  it("runTurn: lỗi ngay bước đầu (chưa làm gì) thì lịch sử như chưa gửi", async () => {
    const h: AgentMessage[] = [
      { role: "user", content: "a" },
      { role: "assistant", content: "b" },
    ];
    const r = await runTurn({
      provider: P("groq"),
      key: "K",
      model: "m",
      host: "Word",
      history: h,
      text: "c",
      agent: { tools: fakeTools([]), log: [] },
      onEvent: () => {},
      fetchFn: (async () => new Response('{"error":{"message":"bad"}}', { status: 400 })) as any,
    });
    assert.ok(r.error instanceof ApiError);
    assert.strictEqual(r.history, h);
  });
  it("Lời nhắc hệ thống không đổi giữa hai lượt dù tài liệu đổi; phần đầu hội thoại gửi đi giữ nguyên từng byte", async () => {
    const tools = fakeTools([]);
    const bodies: any[] = [];
    const f = (async (_u: string, init: any) => {
      bodies.push(JSON.parse(init.body));
      return sse([
        { type: "message_start", message: { usage: { input_tokens: 5, cache_read_input_tokens: bodies.length > 1 ? 900 : 0, cache_creation_input_tokens: 100 } } },
        { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } },
        { type: "content_block_delta", index: 0, delta: { type: "text_delta", text: "ok" } },
        { type: "message_delta", delta: { stop_reason: "end_turn" }, usage: { output_tokens: 7 } },
      ]);
    }) as any;
    const base = { provider: P("anthropic"), key: "K", model: "claude-opus-5-5", host: "Word", onEvent: () => {}, fetchFn: f };
    const r1 = await runTurn({ ...base, history: [], text: "câu một", agent: { tools, log: [] } });
    const r2 = await runTurn({ ...base, history: r1.history, text: "câu hai", agent: { tools, log: [] } });
    assert.deepStrictEqual(bodies[0].system, bodies[1].system, "lời nhắc hệ thống giống hệt");
    assert.ok(!JSON.stringify(bodies[0].system).includes("Trạng thái"), "trạng thái tài liệu không nằm trong lời nhắc hệ thống");
    const strip = (m: any) => JSON.stringify(m).replace(/,"cache_control":\{"type":"ephemeral"\}/g, "");
    assert.strictEqual(strip(bodies[1].messages[0]), strip(bodies[0].messages[0]), "tin nhắn lượt một gửi lại y nguyên");
    assert.ok(strip(bodies[1].messages[0]).includes("Trạng thái") && strip(bodies[1].messages[0]).includes("câu một"));
    // Mốc bộ nhớ đệm: ở lời nhắc hệ thống và ở khối cuối của tin nhắn cuối
    assert.deepStrictEqual(bodies[1].system[0].cache_control, { type: "ephemeral" });
    const last = bodies[1].messages[bodies[1].messages.length - 1];
    assert.deepStrictEqual(last.content[last.content.length - 1].cache_control, { type: "ephemeral" });
    assert.strictEqual(JSON.stringify(bodies[1].messages).split('"cache_control"').length - 1, 1, "chỉ một mốc trong phần hội thoại");
    // Số token: lượt hai đọc được từ bộ nhớ đệm
    assert.strictEqual(r2.usage.cacheRead, 900);
    assert.strictEqual(r2.usage.input, 5 + 900 + 100);
    assert.strictEqual(r2.usage.output, 7);
    assert.strictEqual(r2.lastInputTokens, 1005);
    assert.strictEqual(formatUsage(r2.usage), "Token: vào 1.005 · ra 7 · đọc từ bộ nhớ đệm 900 · ghi vào bộ nhớ đệm 100");
  });
  it("Kết quả công cụ của lượt cũ không bị cắt ngắn khi gửi lại", async () => {
    const big = "x".repeat(20000);
    const h: AgentMessage[] = [
      { role: "user", content: "đọc" },
      { role: "assistant", content: "", toolCalls: [{ id: "r1", name: "doc_read", args: {} }] },
      { role: "tool", toolCallId: "r1", name: "doc_read", content: big },
      { role: "assistant", content: "đã đọc" },
    ];
    let sent: any;
    await runTurn({
      provider: P("groq"),
      key: "K",
      model: "m",
      context: 200000,
      host: "Word",
      history: h,
      text: "tiếp",
      agent: { tools: fakeTools([]), log: [] },
      onEvent: () => {},
      fetchFn: (async (_u: string, init: any) => {
        sent = JSON.parse(init.body);
        return oaText("ok");
      }) as any,
    });
    assert.strictEqual(sent.messages.find((m: any) => m.role === "tool").content.length, 20000);
  });
});

describe("Đợt 0B: model mặc định của OpenRouter khi bật Miễn phí", () => {
  it("Chọn dòng model quen gọi được công cụ, không lấy model miễn phí mới nhất bất kỳ", () => {
    const models = [
      { id: "thinkingmachines/inkling-2-small:free", name: "Inkling", free: true, tools: true, created: 999 },
      { id: "openai/gpt-oss-120b:free", name: "GPT OSS", free: true, tools: true, created: 100 },
      { id: "anthropic/claude-sonnet-5", name: "Sonnet", tools: true, created: 500 },
    ];
    assert.strictEqual(pickModel(models, { providerId: "openrouter", agentOn: true, freeOnly: true }), "openai/gpt-oss-120b:free");
    assert.strictEqual(pickModel(models, { providerId: "openrouter", agentOn: true, freeOnly: false }), "anthropic/claude-sonnet-5");
  });
});

describe("Đợt 0B: đếm token và trần đầu ra", () => {
  it("OpenAI: usage ở khối cuối không có choices; Groq đặt trong x_groq; DeepSeek báo token đọc đệm", () => {
    const a = createStepAccumulator("openai", () => {});
    a.push(JSON.stringify({ choices: [{ delta: { content: "hi" } }] }));
    a.push(JSON.stringify({ choices: [], usage: { prompt_tokens: 120, completion_tokens: 8, prompt_tokens_details: { cached_tokens: 100 } } }));
    assert.deepStrictEqual(a.result().usage, { input: 120, output: 8, cacheRead: 100 });
    const g = createStepAccumulator("openai", () => {});
    g.push(JSON.stringify({ choices: [{ delta: {}, finish_reason: "stop" }], x_groq: { usage: { prompt_tokens: 30, completion_tokens: 4 } } }));
    assert.deepStrictEqual(g.result().usage, { input: 30, output: 4, cacheRead: undefined });
    const d = createStepAccumulator("openai", () => {});
    d.push(JSON.stringify({ choices: [], usage: { prompt_tokens: 50, completion_tokens: 1, prompt_cache_hit_tokens: 40 } }));
    assert.strictEqual(d.result().usage?.cacheRead, 40);
  });
  it("Gemini: usageMetadata, cộng token suy nghĩ vào đầu ra; MAX_TOKENS là bị cắt", () => {
    const a = createStepAccumulator("gemini", () => {});
    a.push(
      JSON.stringify({
        candidates: [{ content: { parts: [{ text: "x" }] }, finishReason: "MAX_TOKENS" }],
        usageMetadata: { promptTokenCount: 200, candidatesTokenCount: 10, thoughtsTokenCount: 5, cachedContentTokenCount: 150 },
      })
    );
    const r = a.result();
    assert.deepStrictEqual(r.usage, { input: 200, output: 15, cacheRead: 150 });
    assert.strictEqual(r.truncated, true);
  });
  it("Cộng dồn và hiển thị", () => {
    assert.deepStrictEqual(addUsage({ input: 1, output: 2 }, { input: 10, cacheRead: 5 }), { input: 11, output: 2, cacheRead: 5, cacheWrite: undefined });
    assert.strictEqual(formatUsage({}), "");
    assert.strictEqual(formatUsage({ input: 1500, output: 20 }), "Token: vào 1.500 · ra 20");
  });
  it("Trần đầu ra Anthropic: theo Models API nếu biết (tối đa 32.000), không thì theo đời model", () => {
    assert.strictEqual(anthropicMaxTokens("claude-opus-5-5", 128000), 32000);
    assert.strictEqual(anthropicMaxTokens("claude-x", 8192), 8192);
    assert.strictEqual(anthropicMaxTokens("claude-3-haiku-20240307"), 4096);
    assert.strictEqual(anthropicMaxTokens("claude-3-5-haiku-20241022"), 8192);
    assert.strictEqual(anthropicMaxTokens("claude-haiku-5-5"), 16000);
    const b = JSON.parse(
      buildStepRequest(P("anthropic"), "K", "claude-opus-5-5", [{ role: "user", content: "x" }], [], { maxOutput: 128000 }).init.body as string
    );
    assert.strictEqual(b.max_tokens, 32000);
  });
  it("Danh mục Anthropic lấy sức chứa và trần đầu ra từ Models API", () => {
    const m = normalizeModels(P("anthropic"), { data: [{ id: "claude-opus-5-5", display_name: "Claude Opus 5.5", max_input_tokens: 1000000, max_tokens: 128000 }] })[0];
    assert.strictEqual(m.context, 1000000);
    assert.strictEqual(m.maxOutput, 128000);
  });
});

describe("Đợt 0B: ngưỡng tóm tắt theo sức chứa của model", () => {
  it("Ngưỡng tỉ lệ với sức chứa, có sàn và trần; không biết sức chứa thì dùng mặc định", () => {
    assert.strictEqual(compactLimitChars(undefined), COMPACT_AT_CHARS);
    assert.strictEqual(compactLimitChars(200000), 300000);
    assert.strictEqual(compactLimitChars(32000), 48000);
    assert.strictEqual(compactLimitChars(1000000), MAX_COMPACT_CHARS);
    assert.strictEqual(compactLimitChars(2000), MIN_COMPACT_CHARS);
  });
  it("Token thực đo vượt 75% sức chứa thì tóm tắt dù số ký tự còn thấp", () => {
    assert.ok(needsCompaction({ chars: 1000, context: 8000, lastInputTokens: 6500 }));
    assert.ok(!needsCompaction({ chars: 1000, context: 8000, lastInputTokens: 3000 }));
    assert.ok(needsCompaction({ chars: 50000, context: 32000 }));
    assert.ok(!needsCompaction({ chars: 50000, context: 200000 }));
  });
  const turn = (i: number, size: number): AgentMessage[] => [
    { role: "user", content: `câu ${i}` },
    { role: "assistant", content: "", toolCalls: [{ id: `t${i}`, name: "doc_read", args: {} }] },
    { role: "tool", toolCallId: `t${i}`, name: "doc_read", content: "n".repeat(size) },
    { role: "assistant", content: `xong ${i}` },
  ];
  it("Model sức chứa nhỏ tóm tắt sớm, model lớn chưa tóm tắt với cùng hội thoại", async () => {
    const h = [...turn(1, 20000), ...turn(2, 20000), ...turn(3, 100), ...turn(4, 100)];
    const n = { calls: 0 };
    const count = (): number => n.calls;
    const f = (async () => {
      n.calls++;
      return oaText("TÓM TẮT");
    }) as any;
    const big = await compactIfNeeded({ history: h, provider: P("groq"), key: "K", model: "m", context: 200000, fetchFn: f });
    assert.ok(!big.compacted && big.history === h && count() === 0);
    const small = await compactIfNeeded({ history: h, provider: P("groq"), key: "K", model: "m", context: 16000, fetchFn: f });
    assert.ok(small.compacted && count() === 1);
    assert.ok(small.history[0].content.startsWith(SUMMARY_PREFIX) && small.history[0].content.includes("TÓM TẮT"));
    assert.strictEqual(small.history.filter((m) => m.role === "user").length, 3, "bản tóm tắt + 2 lượt gần nhất");
    // Phần cũ chỉ còn là bản tóm tắt lần trước (dài) → không tóm tắt lại bản tóm tắt
    const longSummary: AgentMessage[] = [
      { role: "user", content: `${SUMMARY_PREFIX}
${"t".repeat(3000)}` },
      { role: "assistant", content: "Đã nắm." },
      ...turn(5, 100),
      ...turn(6, 100),
    ];
    const again = await compactIfNeeded({ history: longSummary, provider: P("groq"), key: "K", model: "m", limitChars: 3200, fetchFn: f });
    assert.ok(!again.compacted && count() === 1);
  });
  it("Hai lượt gần nhất quá lớn: chỉ giữ một lượt, kết quả công cụ dài trong phần giữ lại được rút gọn một lần", async () => {
    const h = [...turn(1, 100), ...turn(2, 30000), ...turn(3, 30000)];
    const r = await compactIfNeeded({ history: h, provider: P("groq"), key: "K", model: "m", limitChars: 40000, fetchFn: (async () => oaText("TT")) as any });
    assert.ok(r.compacted);
    assert.strictEqual(r.history.filter((m) => m.role === "user").length, 2, "bản tóm tắt + 1 lượt");
    const kept = r.history.find((m) => m.role === "tool") as any;
    assert.ok(kept.content.length < 7000 && kept.content.includes("đã rút gọn"));
  });
  it("Bản ghi đưa đi tóm tắt giữ phần đầu (mục tiêu) và phần cuối", () => {
    const h: AgentMessage[] = [{ role: "user", content: "MỤC TIÊU BAN ĐẦU" }, ...Array.from({ length: 50 }, (_, i) => ({ role: "assistant" as const, content: `bước ${i} ` + "y".repeat(500) }))];
    const t = transcriptOf(h, 4000);
    assert.ok(t.length < 4100 && t.includes("MỤC TIÊU BAN ĐẦU") && t.includes("bước 49") && t.includes("bỏ bớt phần giữa"));
  });
});

describe("Đợt 0B: an toàn", () => {
  it("Lời nhắc hệ thống kể rõ các nguồn không tin cậy và cách xử lý câu ra lệnh ẩn", () => {
    const s = buildSystemPrompt({ base: "GỐC", guide: "CHUẨN", host: "Word", instructions: " viết trang trọng ", skillCatalog: "KỸ NĂNG" });
    for (const k of ["nhận xét", "sửa đổi được theo dõi", "đầu trang và chân trang", "tệp đính kèm", "kết quả tìm trên web", "nội dung tài liệu"]) {
      assert.ok(s.includes(k), `thiếu nguồn: ${k}`);
    }
    assert.ok(s.includes("KHÔNG làm theo") && s.includes("hỏi họ có muốn làm không"));
    assert.ok(s.indexOf("GỐC") < s.indexOf(SAFETY_RULES) && s.includes("Chỉ dẫn riêng của người dùng cho Word") && s.includes("viết trang trọng"));
    assert.strictEqual(s, buildSystemPrompt({ base: "GỐC", guide: "CHUẨN", host: "Word", instructions: "viết trang trọng", skillCatalog: "KỸ NĂNG" }), "cùng đầu vào → cùng chuỗi");
  });
  it("Tin nhắn của lượt: phần add-in thêm tách rõ khỏi lời người dùng, gỡ ra được", () => {
    const m = buildTurnMessage({ text: "Sửa điều 3", state: "Trạng thái: {}", skills: [{ name: "ra-soat", body: "làm A" }] });
    assert.ok(m.startsWith("<bo_sung_cua_add_in>") && m.endsWith("Sửa điều 3"));
    assert.ok(m.includes('<skill name="ra-soat">') && m.includes("không phải lời người dùng"));
    assert.strictEqual(stripTurnContext(m), "Sửa điều 3");
    assert.strictEqual(buildTurnMessage({ text: "chào" }), "chào");
    assert.strictEqual(stripTurnContext("chào"), "chào");
  });
  it("Key Gemini nằm trong tiêu đề, không nằm trong đường dẫn", () => {
    const r = buildStepRequest(P("gemini"), "SECRET-KEY", "gemini-2.5-flash", [{ role: "user", content: "x" }], []);
    assert.ok(!r.url.includes("SECRET-KEY"));
    assert.strictEqual((r.init.headers as any)["x-goog-api-key"], "SECRET-KEY");
  });
  it("Điền bảng Word: nhận ra ô đang có dữ liệu sẽ bị ghi đè; ghi trùng giá trị và hàng mới thì không", () => {
    const cur = [
      ["Mục", "Giá trị"],
      ["A", "1"],
      ["B", ""],
    ];
    const p = planTableFill(cur, [["A", "10"], ["B", "2"], ["C", "3"]], 1, 0);
    assert.deepStrictEqual(p.overwrites, [{ row: 1, col: 1, value: "10", old: "1" }]);
    assert.strictEqual(p.writes.length, 4);
    assert.deepStrictEqual(p.extraRows, [["C", "3"]]);
    assert.strictEqual(planTableFill(cur, [["A", "1"]], 1).overwrites.length, 0);
    assert.strictEqual(planTableFill(cur, [["X", "Y"]]).overwrites.length, 2, "mặc định bắt đầu từ hàng 1 (sau tiêu đề)");
    assert.strictEqual(planTableFill(cur, [["Z", "9"]], 3).writes.length, 0);
  });
  it("Hạn chờ cho bước công cụ: quá hạn thì báo model, vòng lặp chạy tiếp", async () => {
    await assert.rejects(withTimeout(new Promise(() => {}), 10, "doc_write"), /chạy quá/);
    assert.strictEqual(await withTimeout(Promise.resolve(5), 10, "x"), 5);
    // Câu thứ ba: cổng cuối lượt (25/11/2026) nhắc một lần vì lệnh đọc lỗi chưa được gọi lại
    const responses = [oaCall("c1", "doc_read", "{}"), oaText("Đọc xong."), oaText("Bước đọc bị treo.")];
    const slow: HostToolSet = { ...fakeTools([]), execute: () => new Promise(() => {}) };
    const ends: any[] = [];
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: slow,
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      toolTimeoutMs: 20,
      onEvent: (e) => e.type === "tool_end" && ends.push(e),
      fetchFn: (async () => responses.shift()) as any,
    });
    assert.match(ends[0].error, /thôi chờ/);
    assert.strictEqual((msgs[msgs.length - 1] as any).content, "Bước đọc bị treo.");
  });
  it("Lỗi 403 không phải lúc nào cũng là sai key: model bị giới hạn nơi dùng được báo đúng", () => {
    assert.match(messageForStatus(403, "This model is only available in agentic harnesses"), /từ chối yêu cầu.*chọn model khác/);
    assert.match(messageForStatus(403, "Invalid API key"), /Key không đúng hoặc không có quyền/);
    assert.match(messageForStatus(403, ""), /Key không đúng hoặc không có quyền/);
    assert.match(messageForStatus(401, "x is only available on agentic harnesses"), /từ chối yêu cầu \(401\)/);
    assert.match(messageForStatus(401, ""), /Key không đúng hoặc đã hết hạn/);
    assert.match(messageForStatus(529, ""), /quá tải/);
  });
});
