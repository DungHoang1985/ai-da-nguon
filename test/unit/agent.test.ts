import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { buildStepRequest, callModelStep, createStepAccumulator, isToolUnsupportedError, toGeminiSchema } from "../../src/core/agent/llm";
import { compactHistory, runAgent } from "../../src/core/agent/loop";
import { AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { ApiError } from "../../src/core/errors";
import {
  colToLetters,
  diffMiddle,
  escapeWordSearch,
  lettersToCol,
  normalize2D,
  parseMarkdownLine,
  sanitizeSheetName,
  splitAddress,
  topLeft,
} from "../../src/office/util";

const P = (id: string) => getProvider(id)!;

function sse(events: any[], split = 9): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  const bytes = new TextEncoder().encode(text);
  return new Response(
    new ReadableStream<Uint8Array>({
      start(c) {
        for (let i = 0; i < bytes.length; i += split) c.enqueue(bytes.slice(i, i + split));
        c.close();
      },
    }),
    { status: 200 }
  );
}

const TOOLS: ToolSpec[] = [
  { name: "excel_get_overview", description: "overview", parameters: { type: "object", properties: {} } },
  {
    name: "excel_write_range",
    description: "write",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string" },
        values: { type: "array", items: { type: "array", items: { type: "string" } } },
      },
      required: ["address", "values"],
      additionalProperties: false,
    },
    write: true,
  },
];

const HISTORY: AgentMessage[] = [
  { role: "system", content: "SYS" },
  { role: "user", content: "Ghi tổng" },
  {
    role: "assistant",
    content: "Tôi sẽ ghi.",
    toolCalls: [
      { id: "abc123XYZ", name: "excel_write_range", args: { address: "A1", values: [["1"]] } },
      { id: "def456UVW", name: "excel_get_overview", args: {} },
    ],
  },
  { role: "tool", toolCallId: "abc123XYZ", name: "excel_write_range", content: '{"ok":true}' },
  { role: "tool", toolCallId: "def456UVW", name: "excel_get_overview", content: '{"sheets":[]}' },
];

describe("Tác tử: dựng yêu cầu gọi công cụ", () => {
  it("OpenAI-compatible: tools, tool_calls, tool_call_id đúng chuẩn", () => {
    const r = buildStepRequest(P("openrouter"), "K", "m", HISTORY, TOOLS);
    const b = JSON.parse(r.init.body as string);
    assert.strictEqual(r.url, "https://openrouter.ai/api/v1/chat/completions");
    assert.strictEqual(b.tools[1].function.name, "excel_write_range");
    assert.strictEqual(b.messages[0].role, "system");
    const asst = b.messages[2];
    assert.strictEqual(asst.tool_calls.length, 2);
    assert.strictEqual(JSON.parse(asst.tool_calls[0].function.arguments).address, "A1");
    assert.deepStrictEqual(
      b.messages.slice(3).map((m: any) => [m.role, m.tool_call_id]),
      [
        ["tool", "abc123XYZ"],
        ["tool", "def456UVW"],
      ]
    );
  });
  it("Anthropic: tool_use + gom tool_result vào một lượt user, có input_schema", () => {
    const r = buildStepRequest(P("anthropic"), "K", "claude-sonnet-4-5", HISTORY, TOOLS);
    const b = JSON.parse(r.init.body as string);
    assert.strictEqual(b.system, "SYS");
    assert.strictEqual(b.tools[1].input_schema.required[0], "address");
    assert.deepStrictEqual(
      b.messages.map((m: any) => m.role),
      ["user", "assistant", "user"]
    );
    assert.strictEqual(b.messages[1].content[1].type, "tool_use");
    assert.deepStrictEqual(
      b.messages[2].content.map((c: any) => c.tool_use_id),
      ["abc123XYZ", "def456UVW"]
    );
    assert.strictEqual(b.max_tokens, 8192);
  });
  it("Gemini: functionDeclarations, schema rút gọn, functionResponse chung một lượt, giữ nguyên parts gốc", () => {
    const hist: AgentMessage[] = [
      ...HISTORY.slice(0, 2),
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "g1", name: "excel_get_overview", args: {}, nativeId: true }],
        raw: { kind: "gemini", parts: [{ functionCall: { name: "excel_get_overview", args: {}, id: "g1" }, thoughtSignature: "SIG" }] },
      },
      { role: "tool", toolCallId: "g1", name: "excel_get_overview", content: '{"sheets":["A"]}' },
    ];
    const r = buildStepRequest(P("gemini"), "K", "gemini-2.5-flash", hist, TOOLS);
    const b = JSON.parse(r.init.body as string);
    assert.ok(r.url.includes(":streamGenerateContent?alt=sse&key=K"));
    const decls = b.tools[0].functionDeclarations;
    assert.strictEqual(decls[0].parameters, undefined, "công cụ không tham số thì bỏ parameters");
    assert.strictEqual(decls[1].parameters.additionalProperties, undefined);
    assert.strictEqual(b.contents[1].parts[0].thoughtSignature, "SIG");
    assert.deepStrictEqual(b.contents[2].parts[0].functionResponse, {
      name: "excel_get_overview",
      response: { sheets: ["A"] },
      id: "g1",
    });
  });
  it("toGeminiSchema: bỏ khóa không hỗ trợ, mảng luôn có items", () => {
    assert.deepStrictEqual(toGeminiSchema({ type: "object", additionalProperties: false, properties: { a: { type: "array", default: [] } } }), {
      type: "object",
      properties: { a: { type: "array", items: { type: "string" } } },
    });
  });
});

describe("Tác tử: đọc luồng trả về", () => {
  it("OpenAI: ghép tham số bị cắt nhiều mảnh, 2 lệnh gọi song song", () => {
    const got: string[] = [];
    const acc = createStepAccumulator("openai", (d) => got.push(d));
    [
      { choices: [{ delta: { content: "Đang " } }] },
      { choices: [{ delta: { content: "làm" } }] },
      { choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "excel_write_range", arguments: '{"addr' } }] } }] },
      { choices: [{ delta: { tool_calls: [{ index: 0, function: { arguments: 'ess":"A1","values":[["x"]]}' } }] } }] },
      { choices: [{ delta: { tool_calls: [{ index: 1, id: "c2", function: { name: "excel_get_overview", arguments: "" } }] } }] },
      { choices: [{ delta: {}, finish_reason: "tool_calls" }] },
    ].forEach((e) => acc.push(JSON.stringify(e)));
    const r = acc.result();
    assert.strictEqual(r.text, "Đang làm");
    assert.deepStrictEqual(got, ["Đang ", "làm"]);
    assert.deepStrictEqual(r.toolCalls.map((c) => [c.id, c.name]), [
      ["c1", "excel_write_range"],
      ["c2", "excel_get_overview"],
    ]);
    assert.deepStrictEqual(r.toolCalls[0].args, { address: "A1", values: [["x"]] });
    assert.deepStrictEqual(r.toolCalls[1].args, {});
  });
  it("OpenAI: nhà cung cấp không gửi index, gửi arguments dạng object", () => {
    const acc = createStepAccumulator("openai", () => {});
    acc.push(JSON.stringify({ choices: [{ delta: { tool_calls: [{ id: "aaaaaaaaa", function: { name: "t1", arguments: { a: 1 } } }] } }] }));
    acc.push(JSON.stringify({ choices: [{ delta: { tool_calls: [{ id: "bbbbbbbbb", function: { name: "t2", arguments: "{}" } }] } }] }));
    const r = acc.result();
    assert.deepStrictEqual(r.toolCalls.map((c) => c.name), ["t1", "t2"]);
    assert.deepStrictEqual(r.toolCalls[0].args, { a: 1 });
  });
  it("Anthropic: text_delta + tool_use với input_json_delta", () => {
    const acc = createStepAccumulator("anthropic", () => {});
    [
      { type: "message_start", message: {} },
      { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } },
      { type: "content_block_delta", index: 0, delta: { type: "text_delta", text: "Để tôi xem." } },
      { type: "content_block_stop", index: 0 },
      { type: "content_block_start", index: 1, content_block: { type: "tool_use", id: "toolu_1", name: "excel_read_range", input: {} } },
      { type: "content_block_delta", index: 1, delta: { type: "input_json_delta", partial_json: '{"address": "Sheet1' } },
      { type: "content_block_delta", index: 1, delta: { type: "input_json_delta", partial_json: '!A1:B9"}' } },
      { type: "content_block_stop", index: 1 },
      { type: "message_delta", delta: { stop_reason: "tool_use" } },
    ].forEach((e) => acc.push(JSON.stringify(e)));
    const r = acc.result();
    assert.strictEqual(r.text, "Để tôi xem.");
    assert.strictEqual(r.finishReason, "tool_use");
    assert.deepStrictEqual(r.toolCalls, [{ id: "toolu_1", name: "excel_read_range", args: { address: "Sheet1!A1:B9" } }]);
  });
  it("Gemini: functionCall nguyên khối, giữ parts gốc (chữ ký suy nghĩ)", () => {
    const acc = createStepAccumulator("gemini", () => {});
    acc.push(JSON.stringify({ candidates: [{ content: { parts: [{ text: "Ok. " }] } }] }));
    acc.push(
      JSON.stringify({
        candidates: [{ content: { parts: [{ functionCall: { name: "ppt_add_slide", args: { title: "A" } }, thoughtSignature: "S" }] }, finishReason: "STOP" }],
      })
    );
    const r = acc.result();
    assert.strictEqual(r.text, "Ok. ");
    assert.strictEqual(r.toolCalls[0].name, "ppt_add_slide");
    assert.strictEqual(r.toolCalls[0].nativeId, false);
    assert.strictEqual(r.raw!.parts!.length, 2);
    assert.strictEqual(r.raw!.parts![1].thoughtSignature, "S");
  });
  it("Lỗi giữa luồng → báo lỗi tiếng Việt", () => {
    const acc = createStepAccumulator("anthropic", () => {});
    assert.throws(() => acc.push(JSON.stringify({ type: "error", error: { message: "Overloaded" } })), /Nhà cung cấp báo lỗi: Overloaded/);
  });
  it("Nhận biết lỗi model không hỗ trợ công cụ", () => {
    assert.ok(isToolUnsupportedError(new ApiError("Lỗi 404 (No endpoints found that support tool use)", 404)));
    assert.ok(isToolUnsupportedError(new ApiError("Lỗi 400 (model does not support tools)", 400)));
    assert.ok(!isToolUnsupportedError(new ApiError("Key không đúng", 401)));
  });
  it("callModelStep đọc luồng HTTP bị cắt vụn", async () => {
    const f = (async () =>
      sse([
        { choices: [{ delta: { content: "Xin chào" } }] },
        { choices: [{ delta: { tool_calls: [{ index: 0, id: "q", function: { name: "excel_get_overview", arguments: "{}" } }] } }] },
        "[DONE]",
      ], 5)) as any;
    const r = await callModelStep({ provider: P("groq"), key: "K", model: "m", messages: HISTORY.slice(0, 2), tools: TOOLS, onText: () => {}, fetchFn: f });
    assert.strictEqual(r.text, "Xin chào");
    assert.strictEqual(r.toolCalls[0].name, "excel_get_overview");
  });
});

function fakeTools(calls: string[]): HostToolSet {
  return {
    host: "Excel",
    specs: TOOLS,
    overview: async () => ({}),
    systemPrompt: () => "SYS",
    execute: async (name, args, { log }) => {
      calls.push(name);
      if (name === "excel_write_range") {
        log.push({ label: "w", undo: async () => {} });
        return { ok: true, message: `ghi ${args.address}` };
      }
      return { ok: true, sheets: ["Sheet1"] };
    },
  };
}

describe("Tác tử: vòng lặp", () => {
  it("Gọi công cụ → nhận kết quả → trả lời cuối; ghi nhật ký hoàn tác", async () => {
    const bodies: any[] = [];
    const responses = [
      sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "excel_get_overview", arguments: "{}" } }] } }] }]),
      sse([
        { choices: [{ delta: { tool_calls: [{ index: 0, id: "c2", function: { name: "excel_write_range", arguments: '{"address":"B2","values":[["=1+1"]]}' } }] } }] },
      ]),
      sse([{ choices: [{ delta: { content: "Đã ghi B2." } }] }]),
    ];
    const f = (async (_u: string, init: any) => {
      bodies.push(JSON.parse(init.body));
      return responses.shift();
    }) as any;
    const calls: string[] = [];
    const events: string[] = [];
    const log: any[] = [];
    const msgs = await runAgent({
      provider: P("openrouter"),
      key: "K",
      model: "m",
      tools: fakeTools(calls),
      system: "SYS",
      messages: [{ role: "user", content: "Ghi 1+1 vào B2" }],
      log,
      onEvent: (e) => events.push(e.type),
      fetchFn: f,
    });
    assert.deepStrictEqual(calls, ["excel_get_overview", "excel_write_range"]);
    assert.strictEqual(log.length, 1);
    assert.strictEqual(bodies.length, 3);
    assert.strictEqual(bodies[2].messages.filter((m: any) => m.role === "tool").length, 2);
    const last = msgs[msgs.length - 1];
    assert.strictEqual(last.role, "assistant");
    assert.strictEqual((last as any).content, "Đã ghi B2.");
    assert.ok(events.includes("tool_start") && events.includes("tool_end"));
  });
  it("Người dùng từ chối thao tác ghi → không thực hiện, báo cho model", async () => {
    const responses = [
      sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "excel_write_range", arguments: '{"address":"A1","values":[["x"]]}' } }] } }] }]),
      sse([{ choices: [{ delta: { content: "Tôi không sửa." } }] }]),
    ];
    let toolMsg = "";
    const f = (async (_u: string, init: any) => {
      const b = JSON.parse(init.body);
      const t = b.messages.find((m: any) => m.role === "tool");
      if (t) toolMsg = t.content;
      return responses.shift();
    }) as any;
    const calls: string[] = [];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools(calls),
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      onEvent: () => {},
      approve: async () => false,
      fetchFn: f,
    });
    assert.deepStrictEqual(calls, []);
    assert.match(toolMsg, /không cho phép/);
  });
  it("Công cụ lạ / lỗi khi chạy → trả lỗi cho model, không dừng vòng lặp", async () => {
    const responses = [
      sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "khong_ton_tai", arguments: "{}" } }] } }] }]),
      sse([{ choices: [{ delta: { content: "Xin lỗi." } }] }]),
    ];
    const ends: any[] = [];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools([]),
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      onEvent: (e) => e.type === "tool_end" && ends.push(e),
      fetchFn: (async () => responses.shift()) as any,
    });
    assert.match(ends[0].error, /Không có công cụ/);
  });
  it("Giới hạn số bước → dừng và báo", async () => {
    const f = (async () =>
      sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c", function: { name: "excel_get_overview", arguments: "{}" } }] } }] }])) as any;
    const events: string[] = [];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: fakeTools([]),
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      maxSteps: 3,
      onEvent: (e) => events.push(e.type),
      fetchFn: f,
    });
    assert.strictEqual(events.filter((e) => e === "step").length, 3);
    assert.strictEqual(events[events.length - 1], "limit");
  });
  it("Bấm Dừng → ngừng ngay", async () => {
    const ac = new AbortController();
    ac.abort();
    await assert.rejects(
      runAgent({
        provider: P("groq"),
        key: "K",
        model: "m",
        tools: fakeTools([]),
        system: "S",
        messages: [{ role: "user", content: "x" }],
        log: [],
        signal: ac.signal,
        onEvent: () => {},
        fetchFn: (async () => sse([])) as any,
      }),
      (e: any) => e.name === "AbortError"
    );
  });
  it("Thu gọn kết quả công cụ của lượt cũ", () => {
    const h = compactHistory([{ role: "tool", toolCallId: "a", name: "n", content: "x".repeat(5000) }]);
    assert.ok((h[0] as any).content.length < 1700);
  });
});

describe("Hàm phụ cho Office", () => {
  it("Địa chỉ ô và cột", () => {
    assert.strictEqual(colToLetters(0), "A");
    assert.strictEqual(colToLetters(27), "AB");
    assert.strictEqual(lettersToCol("AB"), 27);
    assert.deepStrictEqual(splitAddress("'Doanh thu Q1'!$B$2:C9"), { sheet: "Doanh thu Q1", address: "B2:C9" });
    assert.deepStrictEqual(splitAddress("Sheet1!A1"), { sheet: "Sheet1", address: "A1" });
    assert.deepStrictEqual(splitAddress("A1:B2", "S"), { sheet: "S", address: "A1:B2" });
    assert.deepStrictEqual(topLeft("Sheet1!C5:D9"), { col: 2, row: 4 });
  });
  it("Chuẩn hóa mảng 2 chiều từ dữ liệu model gửi", () => {
    assert.deepStrictEqual(normalize2D([["a"], ["b", 2]]), [
      ["a", ""],
      ["b", 2],
    ]);
    assert.deepStrictEqual(normalize2D(["x", "y"]), [["x", "y"]]);
    assert.deepStrictEqual(normalize2D('[["1",null]]'), [["1", ""]]);
  });
  it("Tên sheet hợp lệ", () => {
    assert.strictEqual(sanitizeSheetName("Báo cáo [2026]/Q1:*?"), "Báo cáo  2026  Q1");
    assert.strictEqual(sanitizeSheetName("x".repeat(40)).length, 31);
  });
  it("So khác biệt theo từ để Track Changes gọn", () => {
    assert.deepStrictEqual(diffMiddle("Bên A thanh toán trong 30 ngày.", "Bên A thanh toán trong 15 ngày."), {
      prefix: "Bên A thanh toán trong ",
      removed: "30",
      added: "15",
      suffix: " ngày.",
    });
    const d = diffMiddle("Hợp đồng số 02/2026 ký ngày 01/01/2026.", "Hợp đồng số 02/2026 ký ngày 01/01/2026, đã có hiệu lực.");
    assert.strictEqual(d.removed, "01/01/2026.");
    assert.strictEqual(d.added, "01/01/2026, đã có hiệu lực.");
    assert.deepStrictEqual(diffMiddle("abc", "abc"), { prefix: "abc", removed: "", added: "", suffix: "" });
  });
  it("Ký hiệu đầu dòng → kiểu đoạn Word; thoát ký tự ^", () => {
    assert.deepStrictEqual(parseMarkdownLine("## Điều 2"), { text: "Điều 2", style: "Heading2" });
    assert.deepStrictEqual(parseMarkdownLine("- Ý một"), { text: "Ý một", style: "ListBullet" });
    assert.deepStrictEqual(parseMarkdownLine("2. Ý hai"), { text: "Ý hai", style: "ListNumber" });
    assert.deepStrictEqual(parseMarkdownLine("**Đậm** thường"), { text: "Đậm thường", style: "Normal" });
    assert.strictEqual(escapeWordSearch("a^b"), "a^^b");
  });
});

import { settableProps } from "../../src/office/excelTools";
describe("Hoàn tác Excel: chỉ ghi lại thuộc tính hợp lệ", () => {
  it("Bỏ address, viền trong, giá trị null; ô không tô nền → pattern None", () => {
    const got = settableProps([
      [
        {
          address: "Sheet1!A1",
          format: {
            font: { bold: true, italic: false, color: "#000000", size: 11, name: "Aptos", underline: "None" },
            fill: { color: "#FFFFFF", pattern: "None" },
            horizontalAlignment: "General",
            verticalAlignment: "Bottom",
            wrapText: false,
            borders: {
              top: { style: "Continuous", color: "#8EA9C1", weight: "Thin" },
              bottom: { style: "None", color: "#000000", weight: "Thin" },
              horizontal: { style: "None" },
              vertical: { style: "None" },
              diagonalDown: { style: "None" },
            },
          },
        },
      ],
    ]);
    assert.deepStrictEqual(got, [
      [
        {
          format: {
            font: { bold: true, italic: false, color: "#000000", size: 11, name: "Aptos", underline: "None" },
            fill: { pattern: "None" },
            horizontalAlignment: "General",
            verticalAlignment: "Bottom",
            wrapText: false,
            borders: { top: { style: "Continuous", color: "#8EA9C1", weight: "Thin" }, bottom: { style: "None" } },
          },
        },
      ],
    ]);
  });
});
