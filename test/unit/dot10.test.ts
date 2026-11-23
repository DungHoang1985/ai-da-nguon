/** Đợt 10: rút gọn token mỗi lượt Word (tasks/dot-10.md) — viết gọn (mọi chế độ) và nạp công cụ theo việc (chỉ chế độ Miễn phí). */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { runTurn } from "../../src/core/agent/turn";
import { buildSystemPrompt } from "../../src/core/agent/prompt";
import { AgentMessage, HostToolSet } from "../../src/core/agent/types";
import { LOAD_TOOLS, WORD_SPECIAL_LINE, WORD_TOOL_GROUPS, fold, freeBasePrompt, groupsNote, loadToolsSpec, requestGroups, toolGroupsFor, usedGroups, visibleSpecs, withLoadTools } from "../../src/core/toolgroups";
import { WORD_SPECS, WORD_SYSTEM_PROMPT } from "../../src/office/wordTools";
import { createCommonTools } from "../../src/office/common";
import { BUILTIN_SKILLS, CORE_GUIDES } from "../../src/core/skillpack";
import { skillCatalog } from "../../src/core/skills";
import { describeCall } from "../../src/office/tools";

const G = WORD_TOOL_GROUPS;
const ids = (s: Set<string>) => [...s].sort();
const req = (text: string, o: { skills?: string[]; hints?: string[] } = {}) => ids(requestGroups(G, { text, ...o }));

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200, headers: { "content-type": "text/event-stream" } });
}
const oaText = (t: string) => sse([{ choices: [{ delta: { content: t } }] }, { choices: [{ delta: {}, finish_reason: "stop" }] }, "[DONE]"]);
const oaTools = (calls: { id: string; name: string; args: any }[]) =>
  sse([
    { choices: [{ delta: { tool_calls: calls.map((c, index) => ({ index, id: c.id, function: { name: c.name, arguments: JSON.stringify(c.args) } })) } }] },
    { choices: [{ delta: {}, finish_reason: "tool_calls" }] },
    "[DONE]",
  ]);

/** Bộ công cụ Word giả: đúng danh sách công cụ thật, công cụ nào cũng trả ok */
function fakeWord(executed: string[]): HostToolSet {
  return {
    host: "Word",
    specs: WORD_SPECS,
    execute: async (name) => {
      executed.push(name);
      return { ok: true };
    },
    overview: async () => ({ paragraphs: 3 }),
    systemPrompt: () => WORD_SYSTEM_PROMPT,
  };
}

/** Máy giả nhận yêu cầu tới Groq (kiểu OpenAI), ghi lại thân yêu cầu */
function fakeGroq(replies: (() => Response)[]) {
  const bodies: any[] = [];
  const fetchFn = (async (_url: string, init: any) => {
    bodies.push(JSON.parse(init.body));
    const r = replies.shift();
    if (!r) throw new Error("Bài thử không chờ thêm lần gọi");
    return r();
  }) as any;
  return { fetchFn, bodies };
}
const sentTools = (body: any): string[] => (body.tools || []).map((t: any) => t.function?.name || t.name);
const sizeOf = (system: string, specs: any[]) => system.length + specs.reduce((n, t) => n + JSON.stringify({ name: t.name, description: t.description, input_schema: t.parameters || {} }).length, 0);

describe("Đợt 10 — nhóm công cụ Word", () => {
  it("Mọi công cụ khai trong nhóm đều có thật (Word hoặc công cụ dùng chung); không công cụ nào thuộc hai nhóm", () => {
    const common = createCommonTools({ host: "Word", skills: BUILTIN_SKILLS as any, shareContext: true });
    const real = new Set([...WORD_SPECS, ...common.specs].map((t) => t.name));
    const seen = new Set<string>();
    for (const g of G)
      for (const t of g.tools) {
        assert.ok(real.has(t), `${g.id}: không có công cụ ${t}`);
        assert.ok(!seen.has(t), `${t} nằm ở hai nhóm`);
        seen.add(t);
      }
  });
  it("Nhóm lõi luôn gửi: đọc, tìm, sửa, chèn, xóa, định dạng, nhận xét, kỹ năng", () => {
    const common = createCommonTools({ host: "Word", skills: BUILTIN_SKILLS as any, shareContext: true });
    const names = visibleSpecs([...WORD_SPECS, ...common.specs], G, new Set()).map((t) => t.name);
    for (const n of ["word_get_overview", "word_get_outline", "word_find", "word_read_clause", "word_read_paragraphs", "word_replace_text", "word_set_paragraph_text", "word_insert_paragraphs", "word_insert_clause", "word_delete_paragraphs", "word_format_text", "word_add_comment", "use_skill"])
      assert.ok(names.includes(n), n);
    for (const n of ["word_create_admin_document", "word_fill_form", "word_compare_versions", "word_list_revisions", "word_read_tables", "write_other_app"]) assert.ok(!names.includes(n), n);
  });
  it("Từ khóa: bật đúng nhóm theo câu của người dùng (có dấu, không dấu)", () => {
    assert.deepStrictEqual(req("Soạn công văn gửi Sở Xây dựng"), ["hanh_chinh"]);
    assert.deepStrictEqual(req("soan to trinh xin chu truong"), ["hanh_chinh"]);
    assert.deepStrictEqual(req("Điền mẫu này giúp tôi"), ["dien_mau"]);
    assert.deepStrictEqual(req("Kiểm tra trước khi ký"), ["soat"]);
    assert.deepStrictEqual(req("Đối tác đã sửa gì trong bản này?"), ["nhan_xet_sua_doi"]);
    assert.deepStrictEqual(req("Xử lý các nhận xét"), ["nhan_xet_sua_doi"]);
    assert.deepStrictEqual(req("Chèn một bảng 3 cột"), ["bang"]);
    assert.deepStrictEqual(req("Lấy số liệu từ Excel đang mở"), ["ngoai"]);
  });
  it("“bằng” (không phải “bảng”) không bật nhóm bảng; câu thường không bật nhóm nào", () => {
    assert.deepStrictEqual(req("Viết lại Điều 5 bằng giọng văn trang trọng hơn"), []);
    assert.deepStrictEqual(req("Thêm khoản phạt chậm thanh toán vào Điều 8"), []);
    assert.deepStrictEqual(req("Tóm tắt nghĩa vụ của bên B"), []);
  });
  it("Kỹ năng tự chọn và tệp đính kèm kéo theo nhóm", () => {
    assert.deepStrictEqual(req("làm giúp", { skills: ["dien-mau"] }), ["dien_mau", "ngoai"]);
    assert.deepStrictEqual(req("làm giúp", { skills: ["xu-ly-nhan-xet"] }), ["nhan_xet_sua_doi"]);
    assert.deepStrictEqual(req("đọc tệp này", { hints: ["soat"] }), ["soat"]);
  });
  it("Nhóm đã dùng trong lịch sử giữ lại: load_tools, gọi thẳng công cụ của nhóm, use_skill", () => {
    const h: AgentMessage[] = [
      { role: "assistant", content: "", toolCalls: [{ id: "1", name: LOAD_TOOLS, args: { group: "bang" } }] },
      { role: "assistant", content: "", toolCalls: [{ id: "2", name: "word_fill_form", args: {} }] },
      { role: "assistant", content: "", toolCalls: [{ id: "3", name: "use_skill", args: { name: "soan-thao-van-ban-hanh-chinh" } }] },
      { role: "assistant", content: "", toolCalls: [{ id: "4", name: LOAD_TOOLS, args: { group: "khong_co" } }] },
    ] as any;
    assert.deepStrictEqual(ids(usedGroups(G, h)), ["bang", "dien_mau", "hanh_chinh", "soat"]);
  });
  it("load_tools: enum là chuỗi (Gemini chỉ nhận enum chuỗi); nhóm sai báo lỗi kèm danh sách; công cụ dùng chung vẫn chạy qua lớp bọc", async () => {
    const spec = loadToolsSpec(G);
    assert.ok((spec.parameters as any).properties.group.enum.every((x: any) => typeof x === "string"));
    const calls: string[] = [];
    const ex = withLoadTools({ specs: [{ name: "use_skill", description: "", parameters: {} } as any], execute: async (n) => (calls.push(n), "ok") }, G);
    assert.deepStrictEqual(ex.specs.map((s) => s.name), ["use_skill", LOAD_TOOLS]);
    const r = await ex.execute(LOAD_TOOLS, { group: "dien_mau" });
    assert.deepStrictEqual(r.tools, ["word_list_form_fields", "word_fill_form"]);
    await assert.rejects(ex.execute(LOAD_TOOLS, { group: "x" }), /hanh_chinh/);
    assert.strictEqual(await ex.execute("use_skill", {}), "ok");
    assert.deepStrictEqual(calls, ["use_skill"]);
  });
  it("Nhãn bước tiếng Việt cho load_tools", () => {
    assert.match(describeCall({ id: "1", name: LOAD_TOOLS, args: { group: "dien_mau" } }), /Nạp thêm công cụ: điền mẫu/);
  });
  it("fold bỏ dấu và chữ đ", () => assert.strictEqual(fold("Điền  Mẫu Tờ khai"), "dien mau to khai"));
});

describe("Đợt 10 — chạy một lượt", () => {
  it("Miễn phí: bước đầu chỉ gửi nhóm lõi + load_tools; model xin nhóm điền mẫu thì bước sau có công cụ điền mẫu và chạy được", async () => {
    const executed: string[] = [];
    const g = fakeGroq([
      () => oaTools([{ id: "a", name: LOAD_TOOLS, args: { group: "dien_mau" } }]),
      () => oaTools([{ id: "b", name: "word_list_form_fields", args: {} }]),
      () => oaText("Xong"),
    ]);
    const res = await runTurn({
      provider: getProvider("groq")!,
      key: "K",
      model: "m",
      host: "Word",
      history: [],
      text: "Hoàn thiện văn bản này cho tôi", // cố ý không khớp từ khóa nào
      onEvent: () => {},
      fetchFn: g.fetchFn,
      agent: { tools: fakeWord(executed), log: [], toolGroups: G },
    });
    assert.strictEqual(res.error, undefined);
    const first = sentTools(g.bodies[0]);
    assert.ok(first.includes(LOAD_TOOLS) && first.includes("word_find"));
    assert.ok(!first.includes("word_fill_form") && !first.includes("word_create_admin_document"));
    const second = sentTools(g.bodies[1]);
    assert.ok(second.includes("word_list_form_fields") && second.includes("word_fill_form"));
    assert.ok(!second.includes("word_create_admin_document"));
    assert.deepStrictEqual(executed, ["word_list_form_fields"]);
    const sys = g.bodies[0].messages[0].content as string;
    assert.ok(sys.includes("## Nhóm công cụ chuyên") && sys.includes("- dien_mau: điền mẫu"));
    // Lượt sau cùng hội thoại: nhóm đã dùng vẫn được gửi ngay từ bước đầu
    const g2 = fakeGroq([() => oaText("OK")]);
    await runTurn({ provider: getProvider("groq")!, key: "K", model: "m", host: "Word", history: res.history, text: "Cảm ơn", onEvent: () => {}, fetchFn: g2.fetchFn, agent: { tools: fakeWord([]), log: [], toolGroups: G } });
    assert.ok(sentTools(g2.bodies[0]).includes("word_fill_form"));
  });
  it("Miễn phí: câu khớp từ khóa bật nhóm ngay từ bước đầu", async () => {
    const g = fakeGroq([() => oaText("OK")]);
    await runTurn({ provider: getProvider("groq")!, key: "K", model: "m", host: "Word", history: [], text: "Soạn tờ trình xin chủ trương", onEvent: () => {}, fetchFn: g.fetchFn, agent: { tools: fakeWord([]), log: [], toolGroups: G } });
    const t = sentTools(g.bodies[0]);
    assert.ok(t.includes("word_create_admin_document") && !t.includes("word_fill_form"));
  });
  it("Model gọi thẳng công cụ chưa gửi (đoán đúng tên): vẫn chạy, và nhóm đó được gửi ở bước sau", async () => {
    const executed: string[] = [];
    const g = fakeGroq([() => oaTools([{ id: "a", name: "word_read_tables", args: {} }]), () => oaText("Xong")]);
    await runTurn({ provider: getProvider("groq")!, key: "K", model: "m", host: "Word", history: [], text: "xem giúp", onEvent: () => {}, fetchFn: g.fetchFn, agent: { tools: fakeWord(executed), log: [], toolGroups: G } });
    assert.deepStrictEqual(executed, ["word_read_tables"]);
    assert.ok(!sentTools(g.bodies[0]).includes("word_read_tables"));
    assert.ok(sentTools(g.bodies[1]).includes("word_insert_table"));
  });
  it("Trả phí (không có toolGroups): gửi đủ công cụ, không có load_tools, lời nhắc không có danh sách nhóm", async () => {
    const g = fakeGroq([() => oaText("OK")]);
    await runTurn({ provider: getProvider("groq")!, key: "K", model: "m", host: "Word", history: [], text: "Xin chào", onEvent: () => {}, fetchFn: g.fetchFn, agent: { tools: fakeWord([]), log: [] } });
    const t = sentTools(g.bodies[0]);
    assert.strictEqual(t.length, WORD_SPECS.length);
    assert.ok(!t.includes(LOAD_TOOLS));
    assert.ok(!String(g.bodies[0].messages[0].content).includes("Nhóm công cụ chuyên"));
  });
});

describe("Đợt 10 — sửa theo soát mã độc lập", () => {
  const skills = (BUILTIN_SKILLS as any[]).filter((s) => !s.hosts || s.hosts.includes("Word"));
  const common = createCommonTools({ host: "Word", skills, shareContext: true });
  const ex = withLoadTools(common, G);
  it("H1: read_attachment_pages luôn được gửi (quy tắc Đợt 5), kể cả khi không nhóm nào bật", () => {
    assert.ok(visibleSpecs([...WORD_SPECS, ...ex.specs], G, new Set()).some((t) => t.name === "read_attachment_pages"));
    assert.ok(!G.some((g) => g.tools.includes("read_attachment_pages")));
  });
  it("H2: lời nhắc hệ thống ở chế độ Miễn phí không nêu tên công cụ nào có thể đang ẩn", () => {
    const sys = buildSystemPrompt({ base: freeBasePrompt(WORD_SYSTEM_PROMPT), guide: CORE_GUIDES.Word, host: "Word", skillCatalog: skillCatalog(skills), toolGroups: groupsNote(G) });
    for (const t of G.flatMap((g) => g.tools)) assert.ok(!sys.includes(t), `lời nhắc Miễn phí còn nêu ${t}`);
    assert.ok(WORD_SYSTEM_PROMPT.includes(WORD_SPECIAL_LINE), "dòng chỉ đường Trả phí phải nằm nguyên trong lời nhắc Word để thay được");
    assert.ok(sys.includes("CHỈ gọi công cụ có trong danh sách"));
  });
  it("H3: kỹ năng gọi bằng /tên bật nhóm của kỹ năng (groupSkills)", async () => {
    const g = fakeGroq([() => oaText("OK")]);
    await runTurn({ provider: getProvider("groq")!, key: "K", model: "m", host: "Word", history: [], text: "[Người dùng gọi kỹ năng]", request: "/dien-mau", onEvent: () => {}, fetchFn: g.fetchFn, agent: { tools: fakeWord([]), log: [], toolGroups: G, groupSkills: ["dien-mau"] } });
    assert.ok(sentTools(g.bodies[0]).includes("word_fill_form"));
  });
  it("M1: mọi công cụ chuyên mà thân kỹ năng Word dựng sẵn nhắc tới đều thuộc nhóm kỹ năng đó kéo theo", () => {
    for (const sk of skills) {
      const named = (sk.body + JSON.stringify(sk.files || {})).match(/\b(word_[a-z_]+|read_other_app|write_other_app)\b/g) || [];
      for (const t of new Set(named)) {
        const grp = G.find((x) => x.tools.includes(t));
        if (grp) assert.ok(grp.skills?.includes(sk.name), `${sk.name} nhắc ${t} nhưng không kéo nhóm ${grp.id}`);
      }
    }
  });
  it("M2: chữ gõ kiểu tổ hợp (NFD) vẫn khớp từ khóa có dấu", () => {
    assert.deepStrictEqual(req("Chèn bảng 3 cột".normalize("NFD")), ["bang"]);
    assert.deepStrictEqual(req("Soạn công văn".normalize("NFD")), ["hanh_chinh"]);
  });
  it("M3: không khớp chuỗi con — số điện thoại, có chữ ký, lại sửa", () => {
    assert.deepStrictEqual(req("Thêm số điện thoại của bên B vào Điều 1"), []);
    assert.deepStrictEqual(req("Văn bản này có chữ ký chưa?"), []);
    assert.deepStrictEqual(req("Viết lại sửa cho gọn Điều 2"), []);
  });
  it("M4: từ khóa số trang, lề trái, cỡ chữ", () => {
    assert.ok(req("Đánh số trang ở giữa").includes("hanh_chinh"));
    assert.deepStrictEqual(req("Chỉnh lề trái 3 cm"), ["hanh_chinh"]);
    assert.deepStrictEqual(req("Cỡ chữ 14 cho toàn văn"), ["hanh_chinh"]);
  });
  it("M5: chỉ chế độ Miễn phí nạp theo việc (Excel, PowerPoint từ Đợt 11: test/unit/dot11.test.ts)", () => {
    assert.strictEqual(toolGroupsFor("free", "Word"), G);
    assert.strictEqual(toolGroupsFor("paid", "Word"), undefined);
    assert.strictEqual(toolGroupsFor("paid", "Excel"), undefined);
    assert.strictEqual(toolGroupsFor("paid", "PowerPoint"), undefined);
  });
});

describe("Đợt 10 — chặn phình lại (số ký tự phần cố định của lượt Word)", () => {
  const skills = (BUILTIN_SKILLS as any[]).filter((s) => !s.hosts || s.hosts.includes("Word"));
  const common = createCommonTools({ host: "Word", skills, shareContext: true });
  const baseParts = { base: WORD_SYSTEM_PROMPT, guide: CORE_GUIDES.Word, host: "Word", skillCatalog: skillCatalog(skills) };
  // Đợt 12 phần E (23/11/2026): +400 cho công cụ remember (đề xuất nhớ) — Trả phí luôn gửi đủ; Miễn phí để trong nhóm bo_nho
  // Bản 4.11.8 (23/11/2026): +500 cho loại hop_dong của word_create_admin_document và kỹ năng soan-hop-dong (một dòng danh mục)
  it("Trả phí: ≤ 31.600 ký tự (mốc 4.9.0: 40.922; giảm ≥ 22%)", () => {
    const n = sizeOf(buildSystemPrompt(baseParts), [...WORD_SPECS, ...common.specs]);
    assert.ok(n <= 31600, `đang ${n}`);
  });
  it("Miễn phí, nhóm lõi: ≤ 20.400 ký tự (giảm ≥ 50%)", () => {
    const ex = withLoadTools(common, G);
    const n = sizeOf(buildSystemPrompt({ ...baseParts, toolGroups: groupsNote(G) }), visibleSpecs([...WORD_SPECS, ...ex.specs], G, new Set()));
    assert.ok(n <= 20400, `đang ${n}`);
  });
  it("Danh mục kỹ năng: kỹ năng Word dựng sẵn đều có mô tả ngắn ≤ 200 ký tự", () => {
    for (const s of skills) assert.ok(s.brief && s.brief.length <= 200, s.name);
  });
});
