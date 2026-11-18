/** Đợt 7 · F: kiểm thử cho các chỗ sửa theo ba lượt soát mã độc lập (lõi tác tử, Đợt 5, Đợt 6) */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { createStepAccumulator } from "../../src/core/agent/llm";
import { STOPPED_MID_TOOL, runAgent, safeJson, withTimeout } from "../../src/core/agent/loop";
import { compactIfNeeded, splitForCompaction, transcriptOf } from "../../src/core/agent/compaction";
import { MAX_TOOL_IMAGES, oldToolImageMessages } from "../../src/core/agent/history";
import { AgentEvent, AgentMessage, HostToolSet } from "../../src/core/agent/types";
import { forgetSecrets, maskSecrets, rememberSecret } from "../../src/core/errors";
import { fetchModels } from "../../src/core/models";
import { MAX_PAGE_TEXT_CHARS, attachmentPrivacyNote, parsePageSpec } from "../../src/core/pdfPages";
import { MAX_TOOL_RESULT_CHARS } from "../../src/core/agent/loop";
import { beforeExcerpt } from "../../src/core/worddiff";

const P = (id: string) => getProvider(id)!;
const sse = (events: any[]) => new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });

describe("Đợt 7 · F: lõi tác tử", () => {
  const long = (role: "user" | "assistant", n: number, tag: string): AgentMessage => ({ role, content: tag + " " + "x".repeat(n) }) as AgentMessage;
  const history: AgentMessage[] = [long("user", 3000, "Q1"), long("assistant", 3000, "A1"), long("user", 3000, "Q2"), long("assistant", 3000, "A2"), long("user", 3000, "Q3"), long("assistant", 3000, "A3")];

  it("Bản tóm tắt rỗng: giữ nguyên hội thoại, không thay bằng dòng trống nghĩa", async () => {
    const r = await compactIfNeeded({ history, provider: P("groq"), key: "K", model: "m", limitChars: 8000, fetchFn: (async () => sse([{ choices: [{ delta: { content: "   " }, finish_reason: "stop" }] }])) as any });
    assert.strictEqual(r.compacted, false);
    assert.strictEqual(r.history, history);
  });
  it("Bản tóm tắt bị cắt cụt (hết độ dài): cũng giữ nguyên", async () => {
    const r = await compactIfNeeded({ history, provider: P("groq"), key: "K", model: "m", limitChars: 8000, fetchFn: (async () => sse([{ choices: [{ delta: { content: "Người dùng nhờ" }, finish_reason: "length" }] }])) as any });
    assert.strictEqual(r.compacted, false);
  });
  it("Bản tóm tắt bình thường: vẫn tóm tắt như trước", async () => {
    const r = await compactIfNeeded({ history, provider: P("groq"), key: "K", model: "m", limitChars: 8000, fetchFn: (async () => sse([{ choices: [{ delta: { content: "Người dùng nhờ sửa Điều 4." }, finish_reason: "stop" }] }])) as any });
    assert.strictEqual(r.compacted, true);
    assert.match(r.history[0].content as string, /sửa Điều 4/);
  });
  it("Cắt để tóm tắt không rơi vào tin ảnh do công cụ chụp (không phải đầu một lượt)", () => {
    const h: AgentMessage[] = [
      { role: "user", content: "lượt 1" },
      { role: "assistant", content: "a1" },
      { role: "user", content: "lượt 2" },
      { role: "assistant", content: "", toolCalls: [{ id: "c1", name: "t", args: {} }] },
      { role: "tool", toolCallId: "c1", name: "t", content: "{}" },
      { role: "user", content: "[ảnh]", images: [{ mime: "image/png", data: "QQ==" }], fromTool: true } as any,
      { role: "assistant", content: "a2" },
      { role: "user", content: "lượt 3" },
      { role: "assistant", content: "a3" },
    ];
    const { old, recent } = splitForCompaction(h, 2);
    assert.strictEqual(recent[0].content, "lượt 2");
    assert.strictEqual(old.length, 2);
  });
  it("Bản ghi đưa đi tóm tắt giữ phần CUỐI của tin người dùng dài (lời yêu cầu thật nằm ở cuối)", () => {
    const t = transcriptOf([{ role: "user", content: "TRẠNG THÁI " + "s".repeat(9000) + " YÊU CẦU: thêm khoản sau Điều 8.2" }]);
    assert.match(t, /thêm khoản sau Điều 8\.2/);
    assert.match(t, /^NGƯỜI DÙNG: TRẠNG THÁI/);
    assert.ok(t.length < 4300);
  });
  it("Bấm Dừng lúc công cụ đang chạy: thôi chờ ngay, không đợi hết hạn chờ", async () => {
    const ac = new AbortController();
    const t0 = Date.now();
    setTimeout(() => ac.abort(), 20);
    await assert.rejects(withTimeout(new Promise((r) => setTimeout(r, 2000)), 5000, "x", ac.signal), (e: any) => e.message === STOPPED_MID_TOOL);
    assert.ok(Date.now() - t0 < 500);
    // Không có tín hiệu, không có hạn: trả đúng lời hứa gốc
    assert.strictEqual(await withTimeout(Promise.resolve(7), 0, "x"), 7);
    assert.strictEqual(await withTimeout(Promise.resolve(8), 1000, "x", new AbortController().signal), 8);
  });
  it("Dừng giữa bước công cụ: mọi lệnh gọi vẫn có kết quả, lượt kết thúc bằng mốc Dừng", async () => {
    const ac = new AbortController();
    const hist: AgentMessage[][] = [];
    const tools: HostToolSet = {
      host: "Word",
      specs: [{ name: "cham", description: "", parameters: { type: "object", properties: {} }, write: true }],
      overview: async () => ({}),
      systemPrompt: () => "S",
      execute: () => new Promise((r) => setTimeout(() => r({ ok: true }), 1500)),
    };
    setTimeout(() => ac.abort(), 30);
    await assert.rejects(
      runAgent({
        provider: P("groq"),
        key: "K",
        model: "m",
        system: "S",
        messages: [{ role: "user", content: "x" }],
        log: [],
        tools,
        signal: ac.signal,
        onEvent: () => {},
        onHistory: (h) => hist.push([...h]),
        fetchFn: (async () =>
          sse([
            { choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "cham", arguments: "{}" } }, { index: 1, id: "c2", function: { name: "cham", arguments: "{}" } }] }, finish_reason: "tool_calls" }] },
          ])) as any,
      }),
      (e: any) => e.name === "AbortError"
    );
    const last = hist[hist.length - 1];
    const toolMsgs = last.filter((m) => m.role === "tool");
    assert.strictEqual(toolMsgs.length, 2);
    assert.match(toolMsgs[0].content as string, /bấm Dừng khi bước này đang chạy/);
    assert.match(toolMsgs[1].content as string, /Chưa thực hiện/);
  });
  it("Kết quả công cụ không đóng gói được (vòng tham chiếu, BigInt) không làm hỏng lượt", () => {
    const loop: any = { a: 1 };
    loop.self = loop;
    assert.match(safeJson(loop), /không đóng gói được/);
    assert.strictEqual(safeJson({ n: BigInt(12) }), '{"n":"12"}');
    assert.strictEqual(safeJson({ ok: true }), '{"ok":true}');
  });
  it("Luồng đứt khi lệnh công cụ chưa có tham số và chưa có lý do kết thúc → coi là bị cắt dở (không chạy với tham số rỗng)", () => {
    const acc = createStepAccumulator("openai", () => {}, () => {}, P("groq"));
    acc.push(JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "word_replace_text" } }] } }] }));
    assert.strictEqual(acc.result().truncated, true);
    const ok = createStepAccumulator("openai", () => {}, () => {}, P("groq"));
    ok.push(JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "word_get_overview", arguments: "" } }] }, finish_reason: "tool_calls" }] }));
    assert.strictEqual(ok.result().truncated, false);
  });
  it("Model trả lời rỗng: khung nhận một dòng nói rõ, không kết thúc im lặng", async () => {
    const texts: string[] = [];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      tools: { host: "Word", specs: [], overview: async () => ({}), systemPrompt: () => "S", execute: async () => ({}) },
      onEvent: (e: AgentEvent) => e.type === "text" && texts.push(e.delta),
      fetchFn: (async () => sse([{ choices: [{ delta: {}, finish_reason: "content_filter" }] }])) as any,
    });
    assert.match(texts.join(""), /Model không trả lời gì.*content_filter/);
  });
  it("Ảnh do công cụ chụp: quá trần thì gỡ ảnh của các lần chụp cũ, theo lô; lần mới nhất luôn giữ", () => {
    assert.deepStrictEqual(oldToolImageMessages([4, 4, 4, 4]), []);
    const drop = oldToolImageMessages([4, 4, 4, 4, 4]);
    assert.deepStrictEqual(drop, [0, 1]);
    const many = Array.from({ length: 15 }, () => 4);
    const d = oldToolImageMessages(many);
    const kept = many.reduce((a, b) => a + b, 0) - d.length * 4;
    assert.ok(kept <= MAX_TOOL_IMAGES && kept >= 4 && !d.includes(14));
    // Mốc cắt không đổi ở từng bước: thêm một lần chụp chưa chắc đổi danh sách gỡ
    assert.deepStrictEqual(oldToolImageMessages([4, 4, 4, 4, 4, 4]), [0, 1]);
    assert.deepStrictEqual(oldToolImageMessages([20]), []);
  });
});

describe("Đợt 7 · F: Đợt 5 và Đợt 6", () => {
  it("Trần chữ của công cụ đọc PDF nằm dưới trần kết quả công cụ của vòng lặp", () => {
    assert.ok(MAX_PAGE_TEXT_CHARS * 1.25 + 3000 < MAX_TOOL_RESULT_CHARS);
  });
  it('Khoảng trang viết có dấu cách ("1 - 3", "2 – 4, 7") được nhận', () => {
    assert.deepStrictEqual(parsePageSpec("1 - 3", 10).pages, [1, 2, 3]);
    assert.deepStrictEqual(parsePageSpec("2 – 4, 7", 10).pages, [2, 3, 4, 7]);
    assert.ok(parsePageSpec("-1", 10).error);
    assert.ok(parsePageSpec("0", 10).error);
  });
  it("Dòng nhắc riêng tư có cả ở kênh Claude của OpenRouter", () => {
    assert.match(attachmentPrivacyNote({ providerId: "openrouter-claude", providerName: "OpenRouter · Claude", modelId: "anthropic/claude-haiku-5.5" }), /OpenRouter/);
    assert.match(attachmentPrivacyNote({ providerId: "openrouter", providerName: "OpenRouter", modelId: "x:free" }), /miễn phí/);
    assert.strictEqual(attachmentPrivacyNote({ providerId: "anthropic", providerName: "Anthropic", modelId: "claude-haiku-5-5" }), "");
  });
  it("Che đúng giá trị key đang dùng, kể cả key không có tiền tố nhận ra được (JWT, hex)", () => {
    forgetSecrets();
    const jwt = "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ4In0.abcDEF123456";
    const hex = "0123456789abcdef0123456789abcdef";
    rememberSecret(jwt);
    rememberSecret(hex);
    rememberSecret("ngắn");
    const out = maskSecrets(`Unauthorized: token ${jwt} is invalid; key=${hex}; chữ thường project_management_2026 giữ nguyên`);
    assert.ok(!out.includes(jwt) && !out.includes(hex));
    assert.match(out, /project_management_2026/);
    assert.ok(maskSecrets("ngắn").includes("ngắn"));
    forgetSecrets();
  });
  it("Nguồn có danh sách model cố định: key bị từ chối (401) thì báo lỗi, không hiện nhầm là đã kết nối", async () => {
    const kimi = P("kimi");
    await assert.rejects(fetchModels(kimi, "khoa-sai-khoa-sai", (async () => new Response('{"error":{"message":"Invalid Authentication"}}', { status: 401 })) as any));
    // Lỗi khác (máy chủ lỗi tạm) thì vẫn dùng tạm danh sách cố định như trước
    const list = await fetchModels(kimi, "khoa-nao-do-12345", (async () => new Response("oops", { status: 500 })) as any);
    assert.ok(list.length > 0);
    forgetSecrets();
  });
  it("Chữ gốc của đoạn đang sửa được trích quanh chỗ bị sửa, không phải 400 ký tự đầu", () => {
    const head = "Mở đầu. ".repeat(90);
    const orig = head + "Bên B thanh toán trong 45 ngày kể từ ngày nhận hóa đơn. " + "Phần sau. ".repeat(30);
    const cur = orig.replace("45 ngày", "90 ngày");
    const ex = beforeExcerpt(orig, cur, 400);
    assert.match(ex, /45 ngày/);
    assert.ok(ex.startsWith("…") && ex.endsWith("…") && ex.length <= 402);
    assert.strictEqual(beforeExcerpt("ngắn thôi", "khác", 400), "ngắn thôi");
    assert.ok(beforeExcerpt(orig, "", 400).length <= 402);
  });
});
