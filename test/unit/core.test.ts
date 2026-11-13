import * as assert from "assert";
import { createSSEParser } from "../../src/core/sse";
import { normalizeModels, fetchModels, isStale, modelsRequest } from "../../src/core/models";
import { buildChatRequest, extractDelta, streamChat } from "../../src/core/chat";
import { getProvider, PROVIDERS } from "../../src/core/providers";
import { Storage, KVStore } from "../../src/core/storage";
import { parseMarkdownTable, stripMarkdown, valuesToTsv } from "../../src/core/table";
import { messageForStatus, networkErrorMessage } from "../../src/core/errors";

const P = (id: string) => getProvider(id)!;

function fakeStore(): KVStore & { data: Record<string, string> } {
  const data: Record<string, string> = {};
  return {
    data,
    getItem: (k) => (k in data ? data[k] : null),
    setItem: (k, v) => {
      data[k] = v;
    },
    removeItem: (k) => {
      delete data[k];
    },
  };
}

function sseResponse(events: string[], split = 7): Response {
  const text = events.map((e) => `data: ${e}\n\n`).join("") + "data: [DONE]\n\n";
  const bytes = new TextEncoder().encode(text);
  const stream = new ReadableStream<Uint8Array>({
    start(c) {
      for (let i = 0; i < bytes.length; i += split) c.enqueue(bytes.slice(i, i + split));
      c.close();
    },
  });
  return new Response(stream, { status: 200 });
}

describe("Danh sách nhà cung cấp", () => {
  it("id không trùng, URL đều là https", () => {
    const ids = PROVIDERS.map((p) => p.id);
    assert.strictEqual(new Set(ids).size, ids.length);
    PROVIDERS.forEach((p) => assert.ok(p.baseUrl.startsWith("https://") && p.keyUrl.startsWith("https://")));
  });
});

describe("Bộ tách SSE", () => {
  it("ghép được dòng bị cắt giữa chừng và bỏ [DONE]", () => {
    const got: string[] = [];
    const p = createSSEParser((d) => got.push(d));
    p.push('data: {"a"');
    p.push(':1}\r\n\ndata: [DONE]\n');
    p.push(": comment\n");
    p.end();
    assert.deepStrictEqual(got, ['{"a":1}']);
  });
});

describe("Chuẩn hóa danh sách model", () => {
  it("OpenRouter: nhận đúng model miễn phí", () => {
    const list = normalizeModels(P("openrouter"), {
      data: [
        { id: "a/x:free", name: "X", pricing: { prompt: "0", completion: "0" }, context_length: 8000 },
        { id: "b/y", name: "Y", pricing: { prompt: "0.000001", completion: "0.000002" } },
        { id: "c/z", name: "Z", pricing: { prompt: "0", completion: "0" } },
      ],
    });
    assert.deepStrictEqual(
      list.map((m) => [m.id, !!m.free]),
      [
        ["a/x:free", true],
        ["b/y", false],
        ["c/z", true],
      ]
    );
  });
  it("OpenAI-compatible: bỏ model không phải chat, bỏ trùng", () => {
    const list = normalizeModels(P("groq"), {
      data: [{ id: "llama-3.3-70b" }, { id: "whisper-large-v3" }, { id: "llama-3.3-70b" }, { id: "llama-guard-4" }],
    });
    assert.deepStrictEqual(list.map((m) => m.id), ["llama-3.3-70b"]);
  });
  it("Gemini: chỉ lấy model có generateContent, bỏ tiền tố models/", () => {
    const list = normalizeModels(P("gemini"), {
      models: [
        { name: "models/gemini-2.5-flash", displayName: "Gemini 2.5 Flash", supportedGenerationMethods: ["generateContent"] },
        { name: "models/text-embedding-004", supportedGenerationMethods: ["embedContent"] },
      ],
    });
    assert.deepStrictEqual(list.map((m) => m.id), ["gemini-2.5-flash"]);
  });
  it("Anthropic và Cohere", () => {
    assert.strictEqual(normalizeModels(P("anthropic"), { data: [{ id: "claude-x", display_name: "Claude X" }] })[0].id, "claude-x");
    assert.strictEqual(normalizeModels(P("cohere"), { models: [{ name: "command-a" }] })[0].id, "command-a");
  });
  it("Yêu cầu liệt kê model gắn key đúng chỗ", () => {
    assert.ok(modelsRequest(P("gemini"), "K").url.includes("key=K"));
    assert.strictEqual((modelsRequest(P("anthropic"), "K").init.headers as any)["x-api-key"], "K");
    assert.strictEqual((modelsRequest(P("groq"), "K").init.headers as any).Authorization, "Bearer K");
    assert.ok(modelsRequest(P("cohere"), "K").url.startsWith("https://api.cohere.com/v1/models"));
  });
  it("Key sai → báo lỗi tiếng Việt", async () => {
    const f = (async () => new Response('{"error":{"message":"Invalid key"}}', { status: 401 })) as any;
    await assert.rejects(fetchModels(P("groq"), "bad", f), /Key không đúng.*Invalid key/);
  });
  it("Danh mục quá 24 giờ thì cần làm mới", () => {
    const now = Date.now();
    assert.strictEqual(isStale(now - 1000, now), false);
    assert.strictEqual(isStale(now - 25 * 3600 * 1000, now), true);
    assert.strictEqual(isStale(undefined, now), true);
  });
});

describe("Chat streaming", () => {
  const msgs = [
    { role: "system" as const, content: "S" },
    { role: "user" as const, content: "Xin chào" },
  ];
  it("Gemini: system tách riêng, role model", () => {
    const r = buildChatRequest({
      provider: P("gemini"),
      key: "K",
      model: "gemini-2.5-flash",
      messages: [...msgs, { role: "assistant", content: "A" }],
    });
    const body = JSON.parse(r.init.body as string);
    assert.ok(r.url.includes(":streamGenerateContent?alt=sse&key=K"));
    assert.strictEqual(body.systemInstruction.parts[0].text, "S");
    assert.deepStrictEqual(body.contents.map((c: any) => c.role), ["user", "model"]);
  });
  it("Anthropic: system ở trường riêng, có header gọi từ trình duyệt", () => {
    const r = buildChatRequest({ provider: P("anthropic"), key: "K", model: "m", messages: msgs });
    const body = JSON.parse(r.init.body as string);
    assert.strictEqual(body.system, "S");
    assert.strictEqual(body.messages.length, 1);
    assert.strictEqual((r.init.headers as any)["anthropic-dangerous-direct-browser-access"], "true");
  });
  it("Tách phần chữ từ 3 kiểu sự kiện", () => {
    assert.strictEqual(extractDelta(P("groq"), '{"choices":[{"delta":{"content":"ab"}}]}'), "ab");
    assert.strictEqual(extractDelta(P("gemini"), '{"candidates":[{"content":{"parts":[{"text":"cd"}]}}]}'), "cd");
    assert.strictEqual(
      extractDelta(P("anthropic"), '{"type":"content_block_delta","delta":{"type":"text_delta","text":"ef"}}'),
      "ef"
    );
    assert.strictEqual(extractDelta(P("anthropic"), '{"type":"message_start"}'), "");
    assert.throws(() => extractDelta(P("openrouter"), '{"error":{"message":"Rate limit"}}'), /Rate limit/);
  });
  it("Nhận câu trả lời chạy dần, ghép đủ chữ", async () => {
    const deltas: string[] = [];
    const f = (async () =>
      sseResponse([
        '{"choices":[{"delta":{"content":"Xin "}}]}',
        '{"choices":[{"delta":{"content":"chào bạn"}}]}',
      ])) as any;
    const full = await streamChat({ provider: P("openrouter"), key: "K", model: "m", messages: msgs, onDelta: (d) => deltas.push(d), fetchFn: f });
    assert.strictEqual(full, "Xin chào bạn");
    assert.strictEqual(deltas.length, 2);
  });
  it("Hết hạn mức 429 → báo lỗi tiếng Việt", async () => {
    const f = (async () => new Response("slow down", { status: 429 })) as any;
    await assert.rejects(
      streamChat({ provider: P("groq"), key: "K", model: "m", messages: msgs, onDelta: () => {}, fetchFn: f }),
      /Hết hạn mức/
    );
  });
});

describe("Lưu key cục bộ", () => {
  it("Lưu, đọc lại sau khi tạo mới (mô phỏng mở lại Office), xóa", () => {
    const store = fakeStore();
    new Storage(store).setKey("groq", "  gsk_123  ");
    const again = new Storage(store);
    assert.strictEqual(again.getKey("groq"), "gsk_123");
    again.setModels("groq", [{ id: "m", name: "m" }]);
    again.removeKey("groq");
    assert.strictEqual(again.getKey("groq"), "");
    assert.strictEqual(again.getModelCache().groq, undefined);
  });
  it("Dữ liệu hỏng không làm sập add-in", () => {
    const store = fakeStore();
    store.setItem("aidn.keys", "{hong");
    assert.deepStrictEqual(new Storage(store).getKeys(), {});
  });
});

describe("Bảng và văn bản", () => {
  it("Tách bảng Markdown thành mảng ô", () => {
    const t = parseMarkdownTable("Kết quả:\n\n| Tháng | **Doanh thu** |\n|---|---:|\n| 1 | 100 |\n| 2 | 200 | x |\n\nHết");
    assert.deepStrictEqual(t, [
      ["Tháng", "Doanh thu", ""],
      ["1", "100", ""],
      ["2", "200", "x"],
    ]);
    assert.strictEqual(parseMarkdownTable("không có bảng | nào"), null);
  });
  it("Bỏ ký hiệu Markdown", () => {
    assert.strictEqual(stripMarkdown("## Tiêu đề\n**đậm** và `mã`\n- ý 1"), "Tiêu đề\nđậm và mã\n• ý 1");
  });
  it("Vùng ô → văn bản, cắt khi quá lớn", () => {
    assert.deepStrictEqual(valuesToTsv([["a", 1], [null, true]]), { text: "a\t1\n\ttrue", truncated: false });
    assert.strictEqual(valuesToTsv([[1, 2], [3, 4]], 3).truncated, true);
  });
  it("Thông báo lỗi", () => {
    assert.match(messageForStatus(402), /hết tiền/);
    assert.match(networkErrorMessage(new TypeError("Failed to fetch")), /CORS/);
  });
});
