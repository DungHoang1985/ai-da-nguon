/** Đợt 1B — phần gọi model: gửi PDF quét nguyên tệp, công tắc Suy nghĩ kỹ, giữ và gửi lại khối suy nghĩ (lịch sử vẫn chỉ thêm),
 *  soát lập luận trước khi ghi, tệp đính kèm cho so sánh phiên bản. */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { buildStepRequest, callModelStep, createStepAccumulator } from "../../src/core/agent/llm";
import { MAX_REVIEW_BLOCKS, runAgent } from "../../src/core/agent/loop";
import { prepareHistory } from "../../src/core/agent/history";
import { compactIfNeeded } from "../../src/core/agent/compaction";
import { REVIEW_SYSTEM, parseReviewVerdict, runTurn } from "../../src/core/agent/turn";
import { THINKING_BINDING_BETA, anthropicThinkingKind, anthropicThinkingParams, geminiThinkingConfig, isThinkingParamError, pdfCapable, thinkingToggleWorks } from "../../src/core/agent/thinking";
import { AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { slimForStorage } from "../../src/core/conversations";
import { DOCX_MIME, withAttachments } from "../../src/core/attachments";
import { clearAttachedFiles, findAttachedFile, listAttachedFiles, registerAttachedFile } from "../../src/core/attachedFiles";
import { ApiError } from "../../src/core/errors";
import { DEFAULT_SETTINGS } from "../../src/core/storage";
import { WORD_SPECS, WORD_SYSTEM_PROMPT } from "../../src/office/wordTools";
import { describeCall } from "../../src/office/tools";

const P = (id: string) => getProvider(id)!;
const sse = (events: any[], init: ResponseInit = { status: 200 }) => new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), init);
const bodyOf = (r: { init: RequestInit }) => JSON.parse(String(r.init.body));
const TOOLS: ToolSpec[] = [
  { name: "doc_read", description: "đọc", parameters: { type: "object", properties: {} } },
  { name: "doc_write", description: "ghi", write: true, reviewable: true, parameters: { type: "object", properties: { text: { type: "string" } } } },
];
const PDF = { mime: "application/pdf", data: "JVBERi0x", name: "quet.pdf" };
const SIG = "c2lnbmF0dXJl";
/** Câu trả lời Anthropic: suy nghĩ → chữ → gọi công cụ */
const anthThinkingStep = (toolId = "toolu_1") =>
  sse([
    { type: "message_start", message: { usage: { input_tokens: 10 } } },
    { type: "content_block_start", index: 0, content_block: { type: "thinking", thinking: "", signature: "" } },
    { type: "content_block_delta", index: 0, delta: { type: "thinking_delta", thinking: "nghĩ" } },
    { type: "content_block_delta", index: 0, delta: { type: "signature_delta", signature: SIG } },
    { type: "content_block_start", index: 1, content_block: { type: "text", text: "" } },
    { type: "content_block_delta", index: 1, delta: { type: "text_delta", text: "Tôi đọc trước." } },
    { type: "content_block_start", index: 2, content_block: { type: "tool_use", id: toolId, name: "doc_read" } },
    { type: "content_block_delta", index: 2, delta: { type: "input_json_delta", partial_json: "{}" } },
    { type: "message_delta", delta: { stop_reason: "tool_use" }, usage: { output_tokens: 5 } },
  ]);
const anthText = (t: string) =>
  sse([
    { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } },
    { type: "content_block_delta", index: 0, delta: { type: "text_delta", text: t } },
    { type: "message_delta", delta: { stop_reason: "end_turn" } },
  ]);

describe("Đợt 1B – PDF quét gửi nguyên tệp", () => {
  const msgs: AgentMessage[] = [
    { role: "system", content: "S" },
    { role: "user", content: "Tóm tắt tệp này", files: [PDF] },
  ];
  it("Anthropic: khối document (base64) đứng trước khối chữ; mốc bộ nhớ đệm ở khối cuối", () => {
    const b = bodyOf(buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", msgs, []));
    assert.deepStrictEqual(b.messages[0].content[0], { type: "document", source: { type: "base64", media_type: "application/pdf", data: "JVBERi0x" } });
    assert.strictEqual(b.messages[0].content[1].type, "text");
    assert.deepStrictEqual(b.messages[0].content[1].cache_control, { type: "ephemeral" });
  });
  it("Gemini: phần inlineData kiểu application/pdf đứng trước chữ", () => {
    const b = bodyOf(buildStepRequest(P("gemini"), "K", "gemini-3-flash", msgs, []));
    assert.deepStrictEqual(b.contents[0].parts[0], { inlineData: { mimeType: "application/pdf", data: "JVBERi0x" } });
    assert.strictEqual(b.contents[0].parts[1].text, "Tóm tắt tệp này");
  });
  it("OpenRouter: phần file dạng data URL và chỉ rõ model tự đọc tệp (không tự chuyển sang OCR tính phí)", () => {
    const b = bodyOf(buildStepRequest(P("openrouter"), "K", "anthropic/claude-haiku-5.5", msgs, [], { webSearch: true }));
    assert.deepStrictEqual(b.messages[1].content[1], { type: "file", file: { filename: "quet.pdf", file_data: "data:application/pdf;base64,JVBERi0x" } });
    assert.deepStrictEqual(b.plugins, [{ id: "web", max_results: 5 }, { id: "file-parser", pdf: { engine: "native" } }]);
  });
  it("Nguồn không đọc được PDF quét: tệp không gửi đi, tin nhắn ghi rõ để model báo người dùng; lịch sử gốc không bị sửa", () => {
    assert.strictEqual(pdfCapable(P("groq")), false);
    assert.ok(pdfCapable(P("anthropic")) && pdfCapable(P("gemini")) && pdfCapable(P("openrouter")));
    const out = prepareHistory(msgs, { provider: P("groq") });
    assert.strictEqual((out[0] as any).files, undefined);
    assert.ok(out[0].content.includes("“quet.pdf”") && out[0].content.includes("KHÔNG gửi được") && out[0].content.includes("Anthropic hoặc Gemini"));
    assert.strictEqual((msgs[1] as any).files.length, 1);
    const b = bodyOf(buildStepRequest(P("groq"), "K", "m", msgs, []));
    assert.strictEqual(typeof b.messages[1].content, "string");
  });
  it("Nguồn đọc được tệp: tệp của lượt cũ vẫn được gửi ở lượt sau (không sửa tin nhắn cũ)", () => {
    const later: AgentMessage[] = [...msgs, { role: "assistant", content: "Tóm tắt…" }, { role: "user", content: "Trang 2 nói gì?" }];
    const b = bodyOf(buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", later, []));
    assert.strictEqual(b.messages[0].content[0].type, "document");
  });
  it("Bản lưu hội thoại bỏ tệp, ghi chú rõ", () => {
    const s = slimForStorage(msgs.slice(1));
    assert.strictEqual((s[0] as any).files, undefined);
    assert.ok(s[0].content.includes("“quet.pdf”") && s[0].content.includes("đính kèm lại"));
  });
});

describe("Đợt 1B – công tắc Suy nghĩ kỹ", () => {
  it("phân loại model Claude theo cách nhận tham số suy nghĩ", () => {
    const k = anthropicThinkingKind;
    assert.strictEqual(k("claude-haiku-5-5"), "bound");
    assert.strictEqual(k("claude-opus-5-5"), "bound");
    assert.strictEqual(k("claude-sonnet-5-5"), "bound");
    assert.strictEqual(k("claude-fable-5-2"), "bound");
    assert.strictEqual(k("claude-opus-5"), "adaptive_on");
    assert.strictEqual(k("claude-sonnet-5"), "adaptive_on");
    assert.strictEqual(k("claude-opus-4-8"), "adaptive_off");
    assert.strictEqual(k("claude-sonnet-4-6"), "adaptive_46");
    assert.strictEqual(k("claude-haiku-4-5"), "budget");
    assert.strictEqual(k("claude-sonnet-4-5-20250929"), "budget");
    assert.strictEqual(k("claude-opus-4-1-20250805"), "budget");
    assert.strictEqual(k("claude-3-7-sonnet-20250219"), "budget");
    assert.strictEqual(k("claude-3-5-haiku-20241022"), "none");
  });
  it("Anthropic, model đời mới: luôn kèm drop_block và tiêu đề beta; bật công tắc thì nâng mức suy nghĩ", () => {
    const off = anthropicThinkingParams("claude-haiku-5-5", false, 32000);
    assert.deepStrictEqual(off, { thinking: { type: "adaptive", block_binding: { prefix_mismatch_behavior: "drop_block" } }, beta: THINKING_BINDING_BETA });
    assert.deepStrictEqual(anthropicThinkingParams("claude-haiku-5-5", true, 32000).output_config, { effort: "high" });
    assert.deepStrictEqual(anthropicThinkingParams("claude-opus-5-5", true, 32000).output_config, { effort: "high" });
    assert.deepStrictEqual(anthropicThinkingParams("claude-sonnet-5-5", true, 32000).output_config, { effort: "xhigh" });
    const r = buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", [{ role: "user", content: "x" }], [], { think: true });
    assert.strictEqual((r.init.headers as any)["anthropic-beta"], THINKING_BINDING_BETA);
    assert.strictEqual(bodyOf(r).output_config.effort, "high");
  });
  it("Anthropic, model cũ hơn: tắt công tắc thì không gửi gì; bật thì adaptive hoặc budget_tokens nhỏ hơn max_tokens", () => {
    assert.deepStrictEqual(anthropicThinkingParams("claude-opus-4-8", false, 32000), {});
    assert.deepStrictEqual(anthropicThinkingParams("claude-opus-4-8", true, 32000), { thinking: { type: "adaptive" } });
    assert.deepStrictEqual(anthropicThinkingParams("claude-opus-5", true, 32000), { thinking: { type: "adaptive" }, output_config: { effort: "xhigh" } });
    const b = bodyOf(buildStepRequest(P("anthropic"), "K", "claude-haiku-4-5", [{ role: "user", content: "x" }], [], { think: true }));
    assert.strictEqual(b.thinking.type, "enabled");
    assert.ok(b.thinking.budget_tokens >= 1024 && b.thinking.budget_tokens < b.max_tokens);
    assert.deepStrictEqual(anthropicThinkingParams("claude-3-5-haiku-20241022", true, 8192), {});
    const none = buildStepRequest(P("anthropic"), "K", "claude-opus-4-8", [{ role: "user", content: "x" }], []);
    assert.strictEqual(bodyOf(none).thinking, undefined);
    assert.strictEqual((none.init.headers as any)["anthropic-beta"], undefined);
  });
  it("Bật hay tắt công tắc không đổi lời nhắc hệ thống, bộ công cụ, tin nhắn (bộ nhớ đệm lời nhắc còn nguyên)", () => {
    const msgs: AgentMessage[] = [
      { role: "system", content: "LOI NHAC" },
      { role: "user", content: "một" },
      { role: "assistant", content: "hai" },
      { role: "user", content: "ba" },
    ];
    for (const [p, m] of [
      ["anthropic", "claude-haiku-5-5"],
      ["gemini", "gemini-3-flash"],
      ["openrouter", "x/y"],
    ]) {
      const a = bodyOf(buildStepRequest(P(p), "K", m, msgs, TOOLS, { think: false }));
      const b = bodyOf(buildStepRequest(P(p), "K", m, msgs, TOOLS, { think: true }));
      for (const k of ["system", "systemInstruction", "tools", "messages", "contents"]) assert.deepStrictEqual(a[k], b[k], `${p}: ${k} đổi khi bật Suy nghĩ kỹ`);
    }
  });
  it("Gemini: thinkingLevel cho dòng 3 trở lên, thinkingBudget cho dòng 2.5, model khác thì không gửi", () => {
    assert.deepStrictEqual(geminiThinkingConfig("gemini-3-flash-preview"), { thinkingLevel: "high" });
    assert.deepStrictEqual(geminiThinkingConfig("gemini-4-carbon"), { thinkingLevel: "high" });
    assert.deepStrictEqual(geminiThinkingConfig("gemini-2.5-pro"), { thinkingBudget: 32768 });
    assert.deepStrictEqual(geminiThinkingConfig("gemini-2.5-flash"), { thinkingBudget: 24576 });
    assert.strictEqual(geminiThinkingConfig("gemini-2.0-flash"), undefined);
    const b = bodyOf(buildStepRequest(P("gemini"), "K", "gemini-3-flash", [{ role: "user", content: "x" }], [], { think: true }));
    assert.deepStrictEqual(b.generationConfig, { thinkingConfig: { thinkingLevel: "high" } });
    assert.strictEqual(bodyOf(buildStepRequest(P("gemini"), "K", "gemini-3-flash", [{ role: "user", content: "x" }], [])).generationConfig, undefined);
  });
  it("OpenRouter và OpenAI: reasoning / reasoning_effort; nguồn khác không gửi gì", () => {
    const m: AgentMessage[] = [{ role: "user", content: "x" }];
    assert.deepStrictEqual(bodyOf(buildStepRequest(P("openrouter"), "K", "a/b", m, [], { think: true })).reasoning, { effort: "high" });
    assert.strictEqual(bodyOf(buildStepRequest(P("openai"), "K", "gpt", m, [], { think: true })).reasoning_effort, "high");
    const g = bodyOf(buildStepRequest(P("groq"), "K", "m", m, [], { think: true }));
    assert.ok(g.reasoning === undefined && g.reasoning_effort === undefined);
    assert.ok(thinkingToggleWorks(P("anthropic"), "claude-haiku-5-5") && thinkingToggleWorks(P("openrouter"), "a/b") && thinkingToggleWorks(P("gemini"), "gemini-3-flash"));
    assert.ok(!thinkingToggleWorks(P("groq"), "m") && !thinkingToggleWorks(P("anthropic"), "claude-3-5-haiku-20241022") && !thinkingToggleWorks(P("gemini"), "gemini-2.0-flash"));
  });
  it("Cài đặt mặc định: Suy nghĩ kỹ và Soát lập luận đều tắt", () => {
    assert.strictEqual(DEFAULT_SETTINGS.deepThink, false);
    assert.strictEqual(DEFAULT_SETTINGS.reviewEdits, false);
  });
});

describe("Đợt 1B – giữ khối suy nghĩ trong lịch sử (lịch sử vẫn chỉ thêm)", () => {
  it("Anthropic: gom khối thinking kèm chữ ký và thứ tự các khối; báo đang suy nghĩ đúng một lần", async () => {
    let thinking = 0;
    const acc = createStepAccumulator("anthropic", () => {}, () => thinking++);
    for (const line of (await anthThinkingStep().text()).split("\n\n")) if (line.startsWith("data: ")) acc.push(line.slice(6));
    const r = acc.result();
    assert.strictEqual(thinking, 1);
    assert.strictEqual(r.text, "Tôi đọc trước.");
    assert.deepStrictEqual(r.raw, { kind: "anthropic", blocks: [{ type: "thinking", thinking: "nghĩ", signature: SIG }, { type: "text", text: "Tôi đọc trước." }, { type: "tool_use" }] });
    assert.strictEqual(r.toolCalls[0].id, "toolu_1");
  });
  it("Khối suy nghĩ chưa có chữ ký (câu trả lời bị cắt) không được giữ; câu trả lời không suy nghĩ thì không có raw", () => {
    const acc = createStepAccumulator("anthropic", () => {});
    acc.push(JSON.stringify({ type: "content_block_start", index: 0, content_block: { type: "thinking", thinking: "" } }));
    acc.push(JSON.stringify({ type: "content_block_delta", index: 0, delta: { type: "thinking_delta", thinking: "dở" } }));
    acc.push(JSON.stringify({ type: "content_block_start", index: 1, content_block: { type: "text", text: "a" } }));
    assert.strictEqual(acc.result().raw, undefined);
  });
  it("Gửi lại: khối suy nghĩ nguyên văn đứng đúng chỗ, trước chữ và lệnh gọi công cụ; khối redacted giữ nguyên", () => {
    const history: AgentMessage[] = [
      { role: "user", content: "đọc đi" },
      {
        role: "assistant",
        content: "Tôi đọc trước.",
        toolCalls: [{ id: "toolu_1", name: "doc_read", args: {} }],
        raw: { kind: "anthropic", blocks: [{ type: "thinking", thinking: "", signature: SIG }, { type: "redacted_thinking", data: "ZGF0YQ==" }, { type: "text" }, { type: "tool_use" }] },
      },
      { role: "tool", toolCallId: "toolu_1", name: "doc_read", content: '{"ok":true}' },
    ];
    const b = bodyOf(buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", history, TOOLS));
    assert.deepStrictEqual(
      b.messages[1].content.map((x: any) => x.type),
      ["thinking", "redacted_thinking", "text", "tool_use"]
    );
    assert.deepStrictEqual(b.messages[1].content[0], { type: "thinking", thinking: "", signature: SIG });
    assert.deepStrictEqual(b.messages[1].content[1], { type: "redacted_thinking", data: "ZGF0YQ==" });
    // Chế độ thường (sau khi nhà cung cấp từ chối): không gửi khối suy nghĩ, không tham số, không tiêu đề beta
    const plain = buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", history, TOOLS, { think: true, plain: true });
    assert.deepStrictEqual(
      bodyOf(plain).messages[1].content.map((x: any) => x.type),
      ["text", "tool_use"]
    );
    assert.ok(bodyOf(plain).thinking === undefined && bodyOf(plain).output_config === undefined && (plain.init.headers as any)["anthropic-beta"] === undefined);
    // Nguồn khác bỏ qua khối suy nghĩ của Anthropic
    const or = bodyOf(buildStepRequest(P("openrouter"), "K", "a/b", history, TOOLS));
    assert.strictEqual(or.messages[1].reasoning_details, undefined);
  });
  it("Vòng lặp tác tử: khối suy nghĩ được ghi vào lịch sử và gửi lại ở bước sau; phần đầu hội thoại giống hệt giữa hai lần gửi", async () => {
    const sent: any[] = [];
    const replies = [anthThinkingStep(), anthText("Xong.")];
    const tools: HostToolSet = { host: "Word", specs: TOOLS, overview: async () => ({}), systemPrompt: () => "S", execute: async () => ({ ok: true }) };
    const out = await runAgent({
      provider: P("anthropic"),
      key: "K",
      model: "claude-haiku-5-5",
      tools,
      system: "S",
      messages: [{ role: "user", content: "đọc đi" }],
      log: [],
      onEvent: () => {},
      fetchFn: (async (_u: string, init: RequestInit) => {
        sent.push(JSON.parse(String(init.body)));
        return replies.shift();
      }) as any,
    });
    assert.strictEqual((out[1] as any).raw.blocks[0].signature, SIG);
    assert.strictEqual(sent.length, 2);
    assert.deepStrictEqual(sent[1].messages[1].content[0], { type: "thinking", thinking: "nghĩ", signature: SIG });
    // Lịch sử chỉ thêm: tin nhắn đầu của lần gửi sau giống lần trước (chỉ khác mốc bộ nhớ đệm được dời về cuối)
    const strip = (m: any) => JSON.stringify(m).replace(/,"cache_control":\{"type":"ephemeral"\}/g, "");
    assert.strictEqual(strip(sent[1].messages[0]), strip(sent[0].messages[0]));
    assert.deepStrictEqual(sent[1].system, sent[0].system);
    assert.deepStrictEqual(sent[1].tools, sent[0].tools);
    assert.deepStrictEqual(sent[1].thinking, sent[0].thinking);
  });
  it("Nhà cung cấp từ chối khối suy nghĩ (400): gửi lại một lần không kèm, báo sự kiện, các bước sau của lượt gửi thẳng chế độ thường", async () => {
    const sent: any[] = [];
    const bad = () => new Response('{"error":{"type":"invalid_request_error","message":"messages.1.content.0: Invalid `signature` in `thinking` block. The block is bound to a different conversation."}}', { status: 400 });
    const replies: (() => Response)[] = [bad, () => anthThinkingStep("toolu_2"), () => anthText("Xong.")];
    const events: string[] = [];
    const history: AgentMessage[] = [
      { role: "user", content: "một" },
      { role: "assistant", content: "hai", raw: { kind: "anthropic", blocks: [{ type: "thinking", thinking: "", signature: "CU" }, { type: "text" }] } },
      { role: "user", content: "ba" },
    ];
    const tools: HostToolSet = { host: "Word", specs: TOOLS, overview: async () => ({}), systemPrompt: () => "S", execute: async () => ({ ok: true }) };
    await runAgent({
      provider: P("anthropic"),
      key: "K",
      model: "claude-haiku-5-5",
      tools,
      system: "S",
      messages: history,
      log: [],
      think: true,
      onEvent: (e) => events.push(e.type),
      fetchFn: (async (_u: string, init: RequestInit) => {
        sent.push({ body: JSON.parse(String(init.body)), beta: (init.headers as any)["anthropic-beta"] });
        return (replies.shift() as () => Response)();
      }) as any,
    });
    assert.strictEqual(sent.length, 3);
    assert.ok(sent[0].beta && sent[0].body.thinking && sent[0].body.messages[1].content[0].type === "thinking");
    assert.ok(!sent[1].beta && !sent[1].body.thinking && !sent[1].body.output_config);
    assert.deepStrictEqual(
      sent[1].body.messages[1].content.map((x: any) => x.type),
      ["text"]
    );
    assert.ok(!sent[2].beta && !sent[2].body.thinking, "bước sau của lượt vẫn ở chế độ thường");
    assert.ok(events.includes("thinking_off"));
    assert.strictEqual(history[1].role === "assistant" && history[1].raw?.blocks?.length, 2, "lịch sử gốc không bị sửa");
  });
  it("Nhận biết lỗi do tham số suy nghĩ; lỗi 400 khác thì không gửi lại", async () => {
    assert.ok(isThinkingParamError(new ApiError("thinking.block_binding: Extra inputs are not permitted", 400)));
    assert.ok(isThinkingParamError(new ApiError("Unsupported parameter: reasoning_effort", 400)));
    assert.ok(!isThinkingParamError(new ApiError("max_tokens: too large", 400)));
    assert.ok(!isThinkingParamError(new ApiError("thinking is overloaded", 529)));
    let calls = 0;
    await assert.rejects(
      callModelStep({
        provider: P("anthropic"),
        key: "K",
        model: "claude-haiku-5-5",
        messages: [{ role: "user", content: "x" }],
        tools: [],
        onText: () => {},
        fetchFn: (async () => {
          calls++;
          return new Response('{"error":{"message":"max_tokens: too large"}}', { status: 400 });
        }) as any,
      })
    );
    assert.strictEqual(calls, 1);
  });
  it("OpenRouter: gom reasoning_details theo từng mẩu và gửi lại nguyên vẹn trên tin nhắn trợ lý", () => {
    const acc = createStepAccumulator("openai", () => {});
    acc.push(JSON.stringify({ choices: [{ delta: { reasoning_details: [{ type: "reasoning.text", index: 0, text: "Xét ", format: "anthropic-claude-v1" }] } }] }));
    acc.push(JSON.stringify({ choices: [{ delta: { reasoning_details: [{ type: "reasoning.text", index: 0, text: "điều 5", signature: "SG" }] } }] }));
    acc.push(JSON.stringify({ choices: [{ delta: { content: "Trả lời", tool_calls: [{ index: 0, id: "c1", function: { name: "doc_read", arguments: "{}" } }] } }] }));
    const r = acc.result();
    assert.deepStrictEqual(r.raw, { kind: "openrouter", reasoning: [{ type: "reasoning.text", index: 0, text: "Xét điều 5", format: "anthropic-claude-v1", signature: "SG" }] });
    const history: AgentMessage[] = [
      { role: "user", content: "x" },
      { role: "assistant", content: r.text, toolCalls: r.toolCalls, raw: r.raw },
      { role: "tool", toolCallId: "c1", name: "doc_read", content: "{}" },
    ];
    assert.deepStrictEqual(bodyOf(buildStepRequest(P("openrouter"), "K", "a/b", history, TOOLS)).messages[1].reasoning_details, r.raw?.reasoning);
    assert.strictEqual(bodyOf(buildStepRequest(P("groq"), "K", "m", history, TOOLS)).messages[1].reasoning_details, undefined);
    assert.strictEqual(bodyOf(buildStepRequest(P("openrouter"), "K", "a/b", history, TOOLS, { plain: true })).messages[1].reasoning_details, undefined);
  });
  it("Tóm tắt hội thoại: phần giữ lại bỏ khối suy nghĩ của Anthropic (đã hết hợp lệ vì phần đứng trước đổi)", async () => {
    const big = "x".repeat(3000);
    const h: AgentMessage[] = [];
    for (let k = 0; k < 4; k++) {
      h.push({ role: "user", content: `câu ${k} ${big}` });
      h.push({ role: "assistant", content: `đáp ${k}`, raw: { kind: "anthropic", blocks: [{ type: "thinking", thinking: "", signature: "S" + k }, { type: "text" }] } });
    }
    const r = await compactIfNeeded({ history: h, provider: P("anthropic"), key: "K", model: "claude-haiku-5-5", limitChars: 8000, fetchFn: (async () => anthText("tóm tắt")) as any });
    assert.strictEqual(r.compacted, true);
    assert.ok(r.history.every((m) => m.role !== "assistant" || !m.raw));
    assert.ok(h.every((m) => m.role !== "assistant" || !!m.raw), "lịch sử gốc không bị sửa");
  });
  it("Bản lưu: từ tin nhắn đầu tiên bị rút gọn trở đi thì bỏ khối suy nghĩ; trước đó giữ nguyên", () => {
    const raw = { kind: "anthropic", blocks: [{ type: "thinking", thinking: "", signature: "S" }, { type: "text" }] };
    const h: AgentMessage[] = [
      { role: "user", content: "a" },
      { role: "assistant", content: "b", raw },
      { role: "user", content: "c", images: [{ mime: "image/png", data: "AA" }] },
      { role: "assistant", content: "d", raw },
    ];
    const s = slimForStorage(h);
    assert.ok((s[1] as any).raw && !(s[3] as any).raw);
  });
  it("Ảnh của lượt cũ: Anthropic vẫn gửi khi chưa rõ model có nhận ảnh (không sửa tin nhắn cũ); nguồn khác như trước", () => {
    const h: AgentMessage[] = [
      { role: "user", content: "ảnh 1", images: [{ mime: "image/png", data: "AA" }] },
      { role: "assistant", content: "ok" },
      { role: "user", content: "tiếp" },
    ];
    assert.strictEqual((prepareHistory(h, { provider: P("anthropic") })[0] as any).images.length, 1);
    assert.strictEqual((prepareHistory(h, { provider: P("groq") })[0] as any).images, undefined);
  });
});

describe("Đợt 1B – soát lập luận trước khi ghi", () => {
  const oaCall = (id: string, text: string) => sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name: "doc_write", arguments: JSON.stringify({ text }) } }] } }] }]);
  const oaText = (t: string) => sse([{ choices: [{ delta: { content: t } }] }]);
  const host = (done: string[]): HostToolSet => ({
    host: "Word",
    specs: TOOLS,
    overview: async () => ({}),
    systemPrompt: () => "S",
    reviewContext: async () => "[¶3] 4.2 Đơn giá thuê là 450.000 đồng.",
    execute: async (name, args) => {
      done.push(`${name}:${args?.text ?? ""}`);
      return { ok: true };
    },
  });

  it("đọc kết luận: chỉ “CHƯA ĐẠT” mới chặn", () => {
    assert.deepStrictEqual(parseReviewVerdict("ĐẠT"), { ok: true });
    assert.deepStrictEqual(parseReviewVerdict("Đạt. Đề xuất đúng yêu cầu."), { ok: true });
    assert.deepStrictEqual(parseReviewVerdict("CHƯA ĐẠT: số tiền bằng chữ không khớp."), { ok: false, note: "số tiền bằng chữ không khớp." });
    assert.deepStrictEqual(parseReviewVerdict("**Chưa đạt** – đổi cả chủ thể."), { ok: false, note: "đổi cả chủ thể." });
    assert.deepStrictEqual(parseReviewVerdict(""), { ok: true });
    assert.ok(REVIEW_SYSTEM.includes("ĐÚNG MỘT DÒNG"));
  });

  it("bị chặn: lệnh ghi không chạy, model nhận lý do và sửa đề xuất; lượt soát không nằm trong lịch sử", async () => {
    const done: string[] = [];
    const reviews: string[] = [];
    const replies = [oaCall("c1", "Đơn giá 420.000 đồng (Bốn trăm nghìn đồng)"), oaCall("c2", "Đơn giá 420.000 đồng"), oaText("Đã sửa.")];
    const events: string[] = [];
    const out = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: host(done),
      system: "S",
      messages: [{ role: "user", content: "giảm đơn giá còn 420.000" }],
      log: [],
      onEvent: (e) => events.push(e.type),
      review: async (call) => {
        reviews.push(call.args.text);
        return call.id === "c1" ? { ok: false, note: "số bằng chữ không khớp" } : { ok: true };
      },
      fetchFn: (async () => replies.shift()) as any,
    });
    assert.deepStrictEqual(done, ["doc_write:Đơn giá 420.000 đồng"]);
    assert.strictEqual(reviews.length, 2);
    const first = out.find((m) => m.role === "tool" && m.toolCallId === "c1") as any;
    assert.ok(first.content.includes("CHƯA GHI") && first.content.includes("số bằng chữ không khớp"));
    assert.ok(events.includes("review_start") && events.includes("review_end"));
    assert.strictEqual(out.filter((m) => m.role === "assistant").length, 3);
  });

  it("công cụ chỉ đọc không bị soát; chặn tối đa hai lần một lượt; lỗi khi soát không chặn lệnh", async () => {
    const done: string[] = [];
    let n = 0;
    const replies = [oaCall("c1", "a"), oaCall("c2", "b"), oaCall("c3", "c"), oaText("xong")];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: host(done),
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      onEvent: () => {},
      review: async () => {
        n++;
        return { ok: false, note: "luôn chê" };
      },
      fetchFn: (async () => replies.shift()) as any,
    });
    assert.strictEqual(n, MAX_REVIEW_BLOCKS);
    assert.deepStrictEqual(done, ["doc_write:c"]);
    const done2: string[] = [];
    const r2 = [oaCall("c1", "a"), oaText("xong")];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: host(done2),
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      onEvent: () => {},
      review: async () => {
        throw new Error("mạng lỗi");
      },
      fetchFn: (async () => r2.shift()) as any,
    });
    assert.deepStrictEqual(done2, ["doc_write:a"]);
  });

  it("runTurn: bật tùy chọn thì gọi một lượt model riêng kèm yêu cầu, chữ hiện tại và đề xuất; tắt thì không gọi", async () => {
    const bodies: any[] = [];
    const mk = () => {
      const replies = [oaCall("c1", "Đơn giá 420.000 đồng"), oaText("CHƯA ĐẠT: thiếu đơn vị tính."), oaCall("c2", "Đơn giá 420.000 đồng mỗi mét vuông"), oaText("ĐẠT"), oaText("Đã sửa.")];
      return (async (_u: string, init: RequestInit) => {
        bodies.push(JSON.parse(String(init.body)));
        return replies.shift();
      }) as any;
    };
    const done: string[] = [];
    const r = await runTurn({ provider: P("groq"), key: "K", model: "m", host: "Word", history: [], text: "giảm đơn giá còn 420.000", agent: { tools: host(done), log: [], reviewEdits: true }, onEvent: () => {}, fetchFn: mk() });
    assert.strictEqual(r.error, undefined);
    assert.deepStrictEqual(done, ["doc_write:Đơn giá 420.000 đồng mỗi mét vuông"]);
    assert.strictEqual(bodies.length, 5);
    const review = bodies[1];
    assert.strictEqual(review.tools, undefined);
    assert.strictEqual(review.messages[0].content, REVIEW_SYSTEM);
    assert.ok(review.messages[1].content.includes("giảm đơn giá còn 420.000") && review.messages[1].content.includes("450.000 đồng") && review.messages[1].content.includes("doc_write"));
    // Lịch sử của hội thoại chính không chứa lượt soát
    assert.ok(!JSON.stringify(r.history).includes("thiếu đơn vị tính") || r.history.some((m) => m.role === "tool" && m.content.includes("thiếu đơn vị tính")));
    assert.ok(!r.history.some((m) => m.role === "assistant" && m.content.includes("CHƯA ĐẠT")));
    bodies.length = 0;
    const replies = [oaCall("c1", "x"), oaText("Đã sửa.")];
    await runTurn({ provider: P("groq"), key: "K", model: "m", host: "Word", history: [], text: "sửa", agent: { tools: host([]), log: [] }, onEvent: () => {}, fetchFn: (async (_u: string, init: RequestInit) => (bodies.push(JSON.parse(String(init.body))), replies.shift())) as any });
    assert.strictEqual(bodies.length, 2);
  });
});

describe("Đợt 1B – tệp đính kèm và công cụ Word mới", () => {
  it("PDF quét và tệp .docx trong tin nhắn: ghi rõ tệp gửi nguyên vẹn; ở Word kèm cách gọi công cụ so sánh", () => {
    const t = withAttachments(
      "So hai bản",
      [
        { name: "quet.pdf", kind: "file", mime: "application/pdf", data: "AA", note: "PDF quét, 3 trang, gửi nguyên tệp" },
        { name: "ban-2.docx", kind: "text", text: "nội dung", mime: DOCX_MIME, data: "BB" },
      ],
      { compare: true }
    );
    assert.ok(t.includes("quet.pdf (PDF quét, 3 trang, gửi nguyên tệp)") && t.includes("đọc trực tiếp"));
    assert.ok(t.includes('word_compare_versions với file = "ban-2.docx"'));
    assert.ok(!withAttachments("x", [{ name: "ban-2.docx", kind: "text", text: "nội dung", mime: DOCX_MIME, data: "BB" }]).includes("word_compare_versions"));
  });
  it("Kho tệp đính kèm của phiên: tìm theo tên, thiếu đuôi, không nêu tên thì lấy tệp gần nhất; giữ tối đa 6 tệp", () => {
    clearAttachedFiles();
    assert.strictEqual(findAttachedFile(), null);
    registerAttachedFile("Hop dong v1.docx", "A");
    registerAttachedFile("Hop dong v2.docx", "B");
    assert.strictEqual(findAttachedFile()?.data, "B");
    assert.strictEqual(findAttachedFile("hop dong v1")?.data, "A");
    assert.strictEqual(findAttachedFile("HOP DONG V2.DOCX")?.data, "B");
    assert.strictEqual(findAttachedFile("hop dong"), null, "tên khớp nhiều tệp thì không đoán");
    assert.strictEqual(findAttachedFile("khong co"), null);
    registerAttachedFile("Hop dong v1.docx", "A2");
    assert.deepStrictEqual(listAttachedFiles(), ["Hop dong v2.docx", "Hop dong v1.docx"]);
    for (let k = 0; k < 8; k++) registerAttachedFile(`t${k}.docx`, "x");
    assert.strictEqual(listAttachedFiles().length, 6);
    clearAttachedFiles();
  });
  it("Bốn công cụ soát và tra: có trong bộ công cụ Word, đều chỉ đọc, có nhãn tiếng Việt; lời nhắc hệ thống nhắc tới", () => {
    for (const name of ["word_check_consistency", "word_read_parts", "word_check_page_format", "word_compare_versions"]) {
      const sp = WORD_SPECS.find((x) => x.name === name);
      assert.ok(sp, name);
      assert.ok(!sp!.write && !sp!.reviewable, `${name} phải là công cụ chỉ đọc`);
      assert.notStrictEqual(describeCall({ id: "1", name, args: {} }), name);
      assert.ok(WORD_SYSTEM_PROMPT.includes(name), `lời nhắc hệ thống chưa nhắc ${name}`);
    }
    assert.strictEqual(WORD_SPECS.length, 29);
    assert.strictEqual(new Set(WORD_SPECS.map((x) => x.name)).size, 29);
  });
  it("Công cụ ghi chữ vào điều khoản được đánh dấu để soát lập luận; công cụ khác thì không", () => {
    const reviewable = WORD_SPECS.filter((x) => x.reviewable).map((x) => x.name).sort();
    assert.deepStrictEqual(reviewable, ["word_insert_clause", "word_replace_selection", "word_replace_text", "word_set_paragraph_text"]);
    assert.ok(reviewable.every((n) => WORD_SPECS.find((x) => x.name === n)!.write));
  });
});

describe("Đợt 1B – sửa theo lần soát mã độc lập", () => {
  it("M5: câu trả lời xen chữ giữa các khối suy nghĩ được gửi lại đúng từng khối, đúng thứ tự, đúng nguyên văn", () => {
    const acc = createStepAccumulator("anthropic", () => {});
    const ev = [
      { type: "content_block_start", index: 0, content_block: { type: "thinking", thinking: "", signature: "S1" } },
      { type: "content_block_start", index: 1, content_block: { type: "text", text: "" } },
      { type: "content_block_delta", index: 1, delta: { type: "text_delta", text: "Đoạn A." } },
      { type: "content_block_start", index: 2, content_block: { type: "thinking", thinking: "", signature: "S2" } },
      { type: "content_block_start", index: 3, content_block: { type: "text", text: "" } },
      { type: "content_block_delta", index: 3, delta: { type: "text_delta", text: "Đoạn B." } },
      { type: "content_block_start", index: 4, content_block: { type: "tool_use", id: "t1", name: "doc_read" } },
      { type: "content_block_delta", index: 4, delta: { type: "input_json_delta", partial_json: "{}" } },
    ];
    ev.forEach((e) => acc.push(JSON.stringify(e)));
    const r = acc.result();
    assert.strictEqual(r.text, "Đoạn A.Đoạn B.");
    const history: AgentMessage[] = [
      { role: "user", content: "x" },
      { role: "assistant", content: r.text, toolCalls: r.toolCalls, raw: r.raw },
      { role: "tool", toolCallId: "t1", name: "doc_read", content: "{}" },
    ];
    const sent = bodyOf(buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", history, TOOLS)).messages[1].content;
    assert.deepStrictEqual(
      sent.map((b: any) => (b.type === "text" ? `text:${b.text}` : b.type === "thinking" ? `thinking:${b.signature}` : b.type)),
      ["thinking:S1", "text:Đoạn A.", "thinking:S2", "text:Đoạn B.", "tool_use"]
    );
    // Dựng hai lần ra cùng một kết quả (phần đầu hội thoại ổn định)
    assert.deepStrictEqual(bodyOf(buildStepRequest(P("anthropic"), "K", "claude-haiku-5-5", history, TOOLS)).messages[1].content, sent);
  });

  it("M4: model đã biết chỉ đọc chữ thì không gửi PDF quét, kể cả ở nguồn nhận tệp; tổng cỡ tệp trong hội thoại đếm được", () => {
    const h: AgentMessage[] = [{ role: "user", content: "đọc tệp", files: [PDF] }];
    const out = prepareHistory(h, { provider: P("openrouter"), vision: false });
    assert.strictEqual((out[0] as any).files, undefined);
    assert.ok(out[0].content.includes("chỉ đọc chữ"));
    assert.strictEqual((prepareHistory(h, { provider: P("openrouter"), vision: true })[0] as any).files.length, 1);
    assert.strictEqual((prepareHistory(h, { provider: P("openrouter") })[0] as any).files.length, 1);
    const { filesBase64Chars, MAX_FILES_TOTAL_B64 } = require("../../src/core/attachments");
    assert.strictEqual(filesBase64Chars([...h, { role: "assistant", content: "ok" }, { role: "user", content: "nữa", files: [PDF, PDF] }]), 3 * PDF.data.length);
    assert.ok(MAX_FILES_TOTAL_B64 < 32 * 1024 * 1024);
  });

  it("M6: gỡ được một tệp khỏi kho tệp đính kèm của phiên", () => {
    const { unregisterAttachedFile } = require("../../src/core/attachedFiles");
    clearAttachedFiles();
    registerAttachedFile("that.docx", "A");
    registerAttachedFile("aidn-thu-so-sanh-X.docx", "B");
    unregisterAttachedFile("aidn-thu-so-sanh-X.docx");
    assert.deepStrictEqual(listAttachedFiles(), ["that.docx"]);
    assert.strictEqual(findAttachedFile()?.data, "A");
    clearAttachedFiles();
  });

  it("T7: nguồn mà add-in không gửi tham số suy nghĩ thì lỗi 400 có chữ “reasoning” không kích hoạt gửi lại", async () => {
    let calls = 0;
    const events: string[] = [];
    await assert.rejects(
      callModelStep({
        provider: P("deepseek"),
        key: "K",
        model: "deepseek-reasoner",
        messages: [{ role: "user", content: "x" }],
        tools: [],
        onText: () => {},
        onThinkingOff: () => events.push("off"),
        fetchFn: (async () => {
          calls++;
          return new Response('{"error":{"message":"Missing reasoning_content field in the assistant message"}}', { status: 400 });
        }) as any,
      })
    );
    assert.strictEqual(calls, 1);
    assert.deepStrictEqual(events, []);
  });

  it("T6 và thứ tự: soát lập luận xong mới hiện bước ghi; bấm Dừng trong lúc soát thì có sự kiện kết thúc và mọi lệnh đều có kết quả", async () => {
    const oaCall = (id: string) => sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name: "doc_write", arguments: '{"text":"a"}' } }] } }] }]);
    const tools: HostToolSet = { host: "Word", specs: TOOLS, overview: async () => ({}), systemPrompt: () => "S", execute: async () => ({ ok: true }) };
    const events: string[] = [];
    const replies = [oaCall("c1"), sse([{ choices: [{ delta: { content: "xong" } }] }])];
    await runAgent({ provider: P("groq"), key: "K", model: "m", tools, system: "S", messages: [{ role: "user", content: "x" }], log: [], onEvent: (e) => events.push(e.type), review: async () => ({ ok: true }), fetchFn: (async () => replies.shift()) as any });
    const order = events.filter((e) => /^(review|tool)_/.test(e));
    assert.deepStrictEqual(order, ["review_start", "review_end", "tool_start", "tool_end"]);
    const ac = new AbortController();
    const ev2: string[] = [];
    let latest: AgentMessage[] = [];
    await assert.rejects(
      runAgent({
        provider: P("groq"),
        key: "K",
        model: "m",
        tools,
        system: "S",
        messages: [{ role: "user", content: "x" }],
        log: [],
        signal: ac.signal,
        onHistory: (h) => (latest = h),
        onEvent: (e) => ev2.push(e.type),
        review: async () => {
          ac.abort();
          throw new DOMException("Đã dừng", "AbortError");
        },
        fetchFn: (async () => oaCall("c9")) as any,
      }),
      (e: any) => e.name === "AbortError"
    );
    assert.deepStrictEqual(
      ev2.filter((e) => /^(review|tool)_/.test(e)),
      ["review_start", "review_end"]
    );
    const tool = latest.find((m) => m.role === "tool") as any;
    assert.ok(tool && tool.toolCallId === "c9" && tool.content.includes("Chưa thực hiện"));
  });
});
