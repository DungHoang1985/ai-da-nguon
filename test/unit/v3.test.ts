import * as assert from "assert";
import JSZip from "jszip";
import { getProvider } from "../../src/core/providers";
import { buildStepRequest } from "../../src/core/agent/llm";
import { runAgent } from "../../src/core/agent/loop";
import { compactIfNeeded, splitForCompaction } from "../../src/core/agent/compaction";
import { AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { MemoryKV } from "../../src/core/db";
import { ConversationStore, MAX_CONVERSATIONS, slimForStorage, titleFrom } from "../../src/core/conversations";
import {
  SkillStore,
  importZip,
  parseFrontmatter,
  parseSkillMd,
  resolveSlash,
  skillCatalog,
  toSkillName,
  useSkill,
} from "../../src/core/skills";
import { BUILTIN_SKILLS, CORE_GUIDES, autoSelectSkills } from "../../src/core/skillpack";
import { findRefs } from "../../src/core/refs";
import { publishSnapshot, readOtherApp, readSnapshot } from "../../src/core/share";
import { docxXmlToText, withAttachments } from "../../src/core/attachments";
import { EXCEL_SPECS } from "../../src/office/excelTools";
import { EXCEL_ADV_SPECS } from "../../src/office/excelAdvanced";
import { WORD_SPECS } from "../../src/office/wordTools";
import { PPT_SPECS } from "../../src/office/pptTools";
import { PPT_ADV_SPECS, buildChartPptx } from "../../src/office/pptAdvanced";
import { createCommonTools } from "../../src/office/common";
import { describeCall } from "../../src/office/tools";

const P = (id: string) => getProvider(id)!;

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200 });
}

class MemStorage {
  m = new Map<string, string>();
  getItem(k: string) {
    return this.m.has(k) ? this.m.get(k)! : null;
  }
  setItem(k: string, v: string) {
    this.m.set(k, v);
  }
  removeItem(k: string) {
    this.m.delete(k);
  }
}

const ALL_SPECS: ToolSpec[] = [...EXCEL_SPECS, ...EXCEL_ADV_SPECS, ...WORD_SPECS, ...PPT_SPECS, ...PPT_ADV_SPECS];

describe("v3 · Ảnh và tìm kiếm web trong yêu cầu gửi nhà cung cấp", () => {
  const msgs: AgentMessage[] = [{ role: "user", content: "Ảnh này là gì?", images: [{ mime: "image/png", data: "QUJD" }] }];
  it("OpenAI-compatible: ảnh dạng image_url data URI", () => {
    const body = JSON.parse(buildStepRequest(P("groq"), "K", "m", msgs, [], {}).init.body as string);
    const user = body.messages.find((m: any) => m.role === "user");
    assert.ok(Array.isArray(user.content));
    assert.ok(user.content.some((c: any) => c.type === "image_url" && c.image_url.url === "data:image/png;base64,QUJD"));
  });
  it("Anthropic: ảnh base64 + công cụ web_search khi bật", () => {
    const body = JSON.parse(buildStepRequest(P("anthropic"), "K", "m", msgs, [], { webSearch: true }).init.body as string);
    assert.ok(JSON.stringify(body.messages).includes('"media_type":"image/png"'));
    assert.ok(body.tools.some((t: any) => t.type === "web_search_20250305"));
  });
  it("Gemini: ảnh inlineData; google_search chỉ khi không có công cụ", () => {
    const b1 = JSON.parse(buildStepRequest(P("gemini"), "K", "m", msgs, [], { webSearch: true }).init.body as string);
    assert.ok(JSON.stringify(b1.contents).includes("inlineData"));
    assert.deepStrictEqual(b1.tools, [{ google_search: {} }]);
    const b2 = JSON.parse(buildStepRequest(P("gemini"), "K", "m", msgs, [EXCEL_SPECS[0]], { webSearch: true }).init.body as string);
    assert.ok(!JSON.stringify(b2.tools).includes("google_search"));
  });
  it("OpenRouter: plugin web", () => {
    const body = JSON.parse(buildStepRequest(P("openrouter"), "K", "m", msgs, [], { webSearch: true }).init.body as string);
    assert.deepStrictEqual(body.plugins, [{ id: "web", max_results: 5 }]);
  });
});

describe("v3 · Xác nhận thao tác rủi ro và công cụ dùng chung", () => {
  const specs: ToolSpec[] = [
    {
      name: "excel_clear_range",
      description: "clear",
      parameters: { type: "object", properties: { address: { type: "string" } } },
      write: true,
      risky: async (a) => `Xóa ${a.address}`,
    },
  ];
  const tools = (calls: string[]): HostToolSet => ({
    host: "Excel",
    specs,
    overview: async () => ({}),
    systemPrompt: () => "S",
    execute: async (name) => {
      calls.push(name);
      return { ok: true };
    },
  });
  const twoStep = (name: string, args: string) => [
    sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name, arguments: args } }] } }] }]),
    sse([{ choices: [{ delta: { content: "Xong." } }] }]),
  ];
  it("Thao tác rủi ro luôn hỏi kèm lý do, kể cả khi tắt 'hỏi trước mỗi thao tác'; từ chối → không chạy", async () => {
    const responses = twoStep("excel_clear_range", '{"address":"A1:B2"}');
    const calls: string[] = [];
    const reasons: (string | undefined)[] = [];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: tools(calls),
      system: "S",
      messages: [{ role: "user", content: "xóa" }],
      log: [],
      onEvent: () => {},
      approve: async (_c, reason) => {
        reasons.push(reason);
        return false;
      },
      askBeforeWrite: false,
      confirmRisky: true,
      fetchFn: (async () => responses.shift()) as any,
    });
    assert.deepStrictEqual(calls, []);
    assert.deepStrictEqual(reasons, ["Xóa A1:B2"]);
  });
  it("Tắt xác nhận rủi ro → chạy luôn", async () => {
    const responses = twoStep("excel_clear_range", '{"address":"A1"}');
    const calls: string[] = [];
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: tools(calls),
      system: "S",
      messages: [{ role: "user", content: "xóa" }],
      log: [],
      onEvent: () => {},
      approve: async () => false,
      confirmRisky: false,
      fetchFn: (async () => responses.shift()) as any,
    });
    assert.deepStrictEqual(calls, ["excel_clear_range"]);
  });
  it("use_skill chạy qua công cụ dùng chung, trả hướng dẫn kỹ năng cho model", async () => {
    const responses = twoStep("use_skill", '{"name":"ra-soat-hop-dong"}');
    let toolMsg = "";
    const extra = createCommonTools({ host: "Word", skills: BUILTIN_SKILLS, shareContext: true });
    assert.deepStrictEqual(
      extra.specs.map((s) => s.name),
      ["use_skill", "read_other_app"]
    );
    assert.deepStrictEqual(extra.specs[1].parameters.properties.app.enum, ["Excel", "PowerPoint"]);
    await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      tools: tools([]),
      extra,
      system: "S",
      messages: [{ role: "user", content: "rà soát" }],
      log: [],
      onEvent: () => {},
      fetchFn: (async (_u: string, init: any) => {
        const b = JSON.parse(init.body);
        const t = b.messages.find((m: any) => m.role === "tool");
        if (t) toolMsg = t.content;
        return responses.shift();
      }) as any,
    });
    assert.match(toolMsg, /Rà soát hợp đồng/);
  });
});

describe("v3 · Skill: định dạng SKILL.md, lệnh /, gói plugin", () => {
  it("Đọc frontmatter (khối >, danh sách), tên kỹ năng tiếng Việt", () => {
    const { data, body } = parseFrontmatter("---\nname: Rà soát HĐ\ndescription: >\n  Dùng khi\n  rà soát\nhosts: [Word]\n---\nThân bài");
    assert.strictEqual(data.description, "Dùng khi rà soát");
    assert.strictEqual(body, "Thân bài");
    assert.strictEqual(toSkillName("Rà soát Hợp đồng – Đặt cọc"), "ra-soat-hop-dong-dat-coc");
    const sk = parseSkillMd("---\nname: x\ndescription: d\nhosts: Word, excel\n---\nB");
    assert.deepStrictEqual(sk.hosts, ["Word", "Excel"]);
    assert.throws(() => parseSkillMd("---\nname: x\n---\nB"), /description/);
  });
  it("Lệnh / thay $ARGUMENTS; gọi kỹ năng / nạp sẵn hướng dẫn", () => {
    const r1 = resolveSlash("/tom-tat hợp đồng A", [], [{ name: "tom-tat", body: "Tóm tắt: $ARGUMENTS" }]);
    assert.strictEqual(r1.content, "Tóm tắt: hợp đồng A");
    assert.strictEqual(r1.label, "/tom-tat");
    const r2 = resolveSlash("/sang-loc-nda", BUILTIN_SKILLS, []);
    assert.match(r2.content, /<skill name="sang-loc-nda">/);
    assert.strictEqual(resolveSlash("bình thường", BUILTIN_SKILLS, []).content, "bình thường");
  });
  it("use_skill đọc thân và tệp tham chiếu; báo lỗi tên sai", () => {
    const skills = [{ name: "a", description: "d", body: "B", files: { "ref.md": "R" }, source: "user" as const }];
    assert.strictEqual(useSkill(skills, "a").instructions, "B");
    assert.strictEqual(useSkill(skills, "a", "ref.md").content, "R");
    assert.throws(() => useSkill(skills, "zz"), /Không có kỹ năng/);
  });
  it("Nhập gói plugin Claude (.zip): skills, commands, tệp tham chiếu; bỏ qua hooks/scripts", async () => {
    const zip = new JSZip();
    zip.file("my-plugin/.claude-plugin/plugin.json", JSON.stringify({ name: "phap-ly" }));
    zip.file("my-plugin/skills/dat-coc/SKILL.md", "---\nname: dat-coc\ndescription: Soát hợp đồng đặt cọc\n---\nBước 1…");
    zip.file("my-plugin/skills/dat-coc/references/mau.md", "Mẫu");
    zip.file("my-plugin/skills/dat-coc/scripts/run.py", "print(1)");
    zip.file("my-plugin/commands/tom-tat.md", "---\ndescription: Tóm tắt\n---\nTóm tắt $ARGUMENTS");
    zip.file("my-plugin/hooks/hooks.json", "{}");
    const r = await importZip(await zip.generateAsync({ type: "uint8array" }), "goi.zip");
    assert.strictEqual(r.skills.length, 1);
    assert.strictEqual(r.skills[0].origin, "plugin phap-ly");
    assert.deepStrictEqual(Object.keys(r.skills[0].files || {}), ["references/mau.md"]);
    assert.strictEqual(r.commands[0].name, "tom-tat");
    assert.ok(r.ignored.some((x) => x.includes("run.py")));
    assert.ok(r.ignored.some((x) => x.includes("hooks")));
  });
  it("Kho kỹ năng: tích hợp + nhập thêm, bật/tắt, lọc theo ứng dụng", async () => {
    const store = new SkillStore(new MemoryKV());
    await store.addSkills([{ name: "rieng", description: "d", body: "b", source: "user", hosts: ["Excel"] }]);
    assert.strictEqual((await store.all()).length, BUILTIN_SKILLS.length + 1);
    assert.ok((await store.forHost("Excel")).some((s) => s.name === "rieng"));
    assert.ok(!(await store.forHost("Word")).some((s) => s.name === "rieng"));
    await store.setEnabled("sang-loc-nda", false);
    assert.ok(!(await store.forHost("Word")).some((s) => s.name === "sang-loc-nda"));
    assert.match(skillCatalog(await store.forHost("Word")), /ra-soat-hop-dong/);
  });
});

describe("v3 · Bộ kỹ năng tích hợp chuyển thể từ skill/plugin chính thức", () => {
  it("17 kỹ năng hợp lệ: tên ≤ 64 ký tự kiểu a-b-c, có mô tả, ứng dụng hợp lệ, có từ khóa", () => {
    assert.strictEqual(BUILTIN_SKILLS.length, 17);
    for (const s of BUILTIN_SKILLS) {
      assert.match(s.name, /^[a-z0-9]+(-[a-z0-9]+)*$/);
      assert.ok(s.name.length <= 64);
      assert.ok(s.description.length > 40, s.name);
      assert.ok((s.hosts || []).every((h) => ["Excel", "Word", "PowerPoint"].includes(h)), s.name);
      assert.ok(s.triggers && s.triggers.length, s.name);
    }
  });
  it("Mọi công cụ được nhắc trong kỹ năng và hướng dẫn nền đều tồn tại", () => {
    const names = new Set([...ALL_SPECS.map((s) => s.name), "use_skill", "read_other_app"]);
    const text = BUILTIN_SKILLS.map((s) => s.body).join("\n") + Object.values(CORE_GUIDES).join("\n");
    const used = new Set(text.match(/\b(?:excel|word|ppt)_[a-z_]+/g) || []);
    const missing = [...used].filter((u) => !names.has(u));
    assert.deepStrictEqual(missing, []);
  });
  it("Không ghi số hiệu văn bản pháp luật trong kỹ năng (tránh trích từ trí nhớ)", () => {
    const text = BUILTIN_SKILLS.map((s) => s.body + s.description).join("\n");
    assert.ok(!/\b\d{1,3}\/\d{4}\/(NĐ-CP|QH\d+|TT-[A-Z]+)/.test(text));
  });
  it("Tự chọn kỹ năng theo câu hỏi", () => {
    const word = BUILTIN_SKILLS.filter((s) => !s.hosts || s.hosts.includes("Word"));
    const excel = BUILTIN_SKILLS.filter((s) => !s.hosts || s.hosts.includes("Excel"));
    assert.strictEqual(autoSelectSkills("Rà soát hợp đồng mua bán này giúp tôi", word)[0]?.name, "ra-soat-hop-dong");
    assert.strictEqual(autoSelectSkills("Kiểm tra văn bản trước khi trình ký", word)[0]?.name, "kiem-tra-truoc-ky");
    assert.strictEqual(autoSelectSkills("Phương án này có đúng quy định không?", word)[0]?.name, "kiem-tra-tuan-thu");
    assert.strictEqual(autoSelectSkills("Vẽ biểu đồ doanh thu theo tháng", excel)[0]?.name, "thiet-ke-bieu-do");
    assert.deepStrictEqual(autoSelectSkills("xin chào", word), []);
  });
  it("Mỗi công cụ Office mới có nhãn tiếng Việt", () => {
    for (const sp of [...EXCEL_ADV_SPECS, ...PPT_ADV_SPECS, ...WORD_SPECS]) {
      const label = describeCall({ id: "x", name: sp.name, args: {} });
      assert.notStrictEqual(label, sp.name, `thiếu nhãn cho ${sp.name}`);
    }
  });
});

describe("v3 · Biểu đồ gốc PowerPoint (pptxgenjs)", () => {
  it("Tạo slide có biểu đồ cột thật, màu theo chủ đề, kèm dữ liệu Excel nhúng", async () => {
    const b64 = await buildChartPptx(
      { chart_type: "column", categories: ["Q1", "Q2", "Q3"], series: [{ name: "Doanh thu", values: [10, 20, 30] }], title: "Doanh thu tăng" },
      { w: 960, h: 540 }
    );
    const zip = await JSZip.loadAsync(b64, { base64: true });
    const files = Object.keys(zip.files);
    const chart = files.find((f) => /ppt\/charts\/chart\d+\.xml$/.test(f));
    assert.ok(chart, "không có chart xml");
    const xml = await zip.file(chart!)!.async("string");
    assert.ok(xml.includes("<c:barChart>"));
    assert.ok(xml.includes('<c:barDir val="col"/>'));
    assert.ok(xml.includes('<a:schemeClr val="accent1"'));
    assert.ok(files.some((f) => /ppt\/embeddings\/.*\.xlsx$/.test(f)), "không có dữ liệu nhúng");
    const pres = await zip.file("ppt/presentation.xml")!.async("string");
    assert.match(pres, /<p:sldSz cx="12192000" cy="6858000"/);
  });
  it("Biểu đồ tròn và cột chồng", async () => {
    const pie = await JSZip.loadAsync(
      await buildChartPptx({ chart_type: "pie", categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] }, { w: 720, h: 540 }),
      { base64: true }
    );
    const pxml = await pie.file(Object.keys(pie.files).find((f) => /charts\/chart\d+\.xml$/.test(f))!)!.async("string");
    assert.ok(pxml.includes("<c:pieChart>"));
    const st = await JSZip.loadAsync(
      await buildChartPptx(
        { chart_type: "stacked_column", categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }, { name: "S2", values: [3, 4] }] },
        { w: 960, h: 540 }
      ),
      { base64: true }
    );
    const sxml = await st.file(Object.keys(st.files).find((f) => /charts\/chart\d+\.xml$/.test(f))!)!.async("string");
    assert.ok(sxml.includes('<c:grouping val="stacked"/>'));
  });
  it("Mọi loại biểu đồ, 1 hoặc 2 chuỗi: không có mã màu sai (PowerPoint từ chối tệp có srgbClr không phải mã hex)", async () => {
    for (const t of ["column", "stacked_column", "bar", "stacked_bar", "line", "area", "pie", "doughnut"]) {
      for (const n of [1, 2]) {
        const series = Array.from({ length: n }, (_, i) => ({ name: "S" + i, values: [1, -2, 3] }));
        const zip = await JSZip.loadAsync(
          await buildChartPptx({ chart_type: t, categories: ["A", "B", "C"], series, show_values: true, chart_title: "T" }, { w: 960, h: 540 }),
          { base64: true }
        );
        const f = Object.keys(zip.files).find((x) => /ppt\/charts\/chart\d+\.xml$/.test(x))!;
        const xml = await zip.file(f)!.async("string");
        const bad = [...xml.matchAll(/srgbClr val="([^"]*)"/g)].map((m) => m[1]).filter((v) => !/^[0-9A-Fa-f]{6}$/.test(v));
        assert.deepStrictEqual(bad, [], `${t}/${n} chuỗi`);
        if (t.includes("column") || t.includes("bar")) assert.ok(!xml.includes("<c:dPt>"), `${t}/${n}: tô màu từng cột`);
      }
    }
  });
  it("Thiếu dữ liệu → báo lỗi tiếng Việt", async () => {
    await assert.rejects(buildChartPptx({ chart_type: "column", categories: [], series: [] }, { w: 960, h: 540 }), /categories/);
  });
});

describe("v3 · Lịch sử, tóm tắt hội thoại, chia sẻ giữa ứng dụng, tham chiếu", () => {
  it("Lưu, liệt kê theo ứng dụng, mở, xóa; giới hạn số hội thoại", async () => {
    const st = new ConversationStore(new MemoryKV());
    for (let i = 0; i < MAX_CONVERSATIONS + 3; i++) {
      await st.save({ id: "c" + i, host: i % 2 ? "Word" : "Excel", title: "t" + i, updatedAt: i, items: [], history: [] });
    }
    const all = await st.list();
    assert.strictEqual(all.length, MAX_CONVERSATIONS);
    assert.strictEqual(all[0].id, "c" + (MAX_CONVERSATIONS + 2));
    assert.ok((await st.list("Word")).every((c) => c.host === "Word"));
    await st.delete(all[0].id);
    assert.strictEqual(await st.get(all[0].id), undefined);
    assert.strictEqual(titleFrom("  a   b "), "a b");
  });
  it("Bản lưu bỏ ảnh, cắt kết quả công cụ dài", () => {
    const slim = slimForStorage([
      { role: "user", content: "x", images: [{ mime: "image/png", data: "A".repeat(100) }] },
      { role: "tool", toolCallId: "1", name: "t", content: "y".repeat(10000) },
    ]);
    assert.strictEqual((slim[0] as any).images, undefined);
    assert.ok(slim[1].content.length < 5000);
  });
  it("Tóm tắt khi hội thoại quá dài: giữ 2 lượt gần nhất, cắt ở câu hỏi người dùng", async () => {
    const h: AgentMessage[] = [];
    for (let i = 0; i < 6; i++) {
      h.push({ role: "user", content: "hỏi " + i + " " + "x".repeat(300) });
      h.push({ role: "assistant", content: "đáp " + i });
    }
    const { old, recent } = splitForCompaction(h);
    assert.strictEqual(recent[0].role, "user");
    assert.strictEqual(old.length + recent.length, h.length);
    const res = await compactIfNeeded({
      history: h,
      provider: P("groq"),
      key: "K",
      model: "m",
      limitChars: 1000,
      fetchFn: (async () => sse([{ choices: [{ delta: { content: "TÓM TẮT" } }] }])) as any,
    });
    assert.ok(res.compacted);
    assert.match(res.history[0].content, /TÓM TẮT/);
    assert.ok(res.history.length < h.length);
  });
  it("Chia sẻ nội dung giữa Excel – Word – PowerPoint; hết hạn sau 30 phút", () => {
    const kv = new MemStorage();
    publishSnapshot("Excel", { a: 1 }, "Sổ tính", kv);
    assert.deepStrictEqual(readOtherApp("excel", "Word", kv).data, { a: 1 });
    assert.throws(() => readOtherApp("Excel", "Excel", kv), /Đang ở Excel/);
    assert.throws(() => readOtherApp("PowerPoint", "Word", kv), /Chưa có dữ liệu/);
    assert.strictEqual(readSnapshot("Excel", Date.now() + 31 * 60 * 1000, kv), null);
  });
  it("Tham chiếu bấm được: ô Excel, đoạn Word, slide PowerPoint", () => {
    const ex = findRefs("Tổng ở 'Doanh thu'!C12 và B2:B5, không phải COVID", "Excel").filter((p) => p.ref);
    assert.deepStrictEqual(
      ex.map((p) => p.ref),
      [
        { kind: "cell", address: "'Doanh thu'!C12" },
        { kind: "cell", address: "B2:B5" },
      ]
    );
    const wd = findRefs("Xem [¶12] và đoạn 5, đoạn số 7", "Word").filter((p) => p.ref).map((p) => p.ref);
    assert.deepStrictEqual(wd, [
      { kind: "para", index: 12 },
      { kind: "para", index: 5 },
      { kind: "para", index: 7 },
    ]);
    const pp = findRefs("Sửa slide 3 và slide số 4", "PowerPoint").filter((p) => p.ref).map((p) => p.ref);
    assert.deepStrictEqual(pp, [
      { kind: "slide", index: 3 },
      { kind: "slide", index: 4 },
    ]);
  });
  it("Tệp đính kèm: DOCX → chữ, ghép vào câu hỏi", () => {
    const xml =
      '<w:document><w:body><w:p><w:r><w:t>Điều 1.</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Giá </w:t></w:r><w:r><w:t>100</w:t></w:r></w:p></w:body></w:document>';
    const t = docxXmlToText(xml);
    assert.match(t, /Điều 1\./);
    assert.match(t, /Giá 100/);
    const q = withAttachments("Tóm tắt", [{ name: "a.docx", kind: "text", text: t }]);
    assert.match(q, /Tệp đính kèm: a\.docx/);
  });
});
