/** Đợt 11: rút gọn token mỗi lượt Excel và PowerPoint (tasks/dot-11.md) — viết gọn (mọi chế độ) và nạp công cụ theo việc (chỉ Miễn phí). */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { runTurn } from "../../src/core/agent/turn";
import { buildSystemPrompt } from "../../src/core/agent/prompt";
import { HostToolSet, ToolSpec } from "../../src/core/agent/types";
import {
  ALL_TOOL_GROUPS,
  EXCEL_SPECIAL_LINE,
  EXCEL_TOOL_GROUPS,
  LOAD_TOOLS,
  PPT_SPECIAL_LINE,
  PPT_TOOL_GROUPS,
  ToolGroup,
  WORD_TOOL_GROUPS,
  freeBasePrompt,
  groupsNote,
  hideGroupTools,
  requestGroups,
  toolGroupsFor,
  visibleSpecs,
  withLoadTools,
} from "../../src/core/toolgroups";
import { EXCEL_SYSTEM_PROMPT, excelTools } from "../../src/office/excelTools";
import { PPT_SYSTEM_PROMPT, pptTools } from "../../src/office/pptTools";
import { createCommonTools } from "../../src/office/common";
import { BUILTIN_SKILLS, CORE_GUIDES } from "../../src/core/skillpack";
import { skillCatalog } from "../../src/core/skills";
import { describeCall } from "../../src/office/tools";

const HOSTS: { host: "Excel" | "PowerPoint"; G: ToolGroup[]; tools: HostToolSet; prompt: string; line: string }[] = [
  { host: "Excel", G: EXCEL_TOOL_GROUPS, tools: excelTools, prompt: EXCEL_SYSTEM_PROMPT, line: EXCEL_SPECIAL_LINE },
  { host: "PowerPoint", G: PPT_TOOL_GROUPS, tools: pptTools, prompt: PPT_SYSTEM_PROMPT, line: PPT_SPECIAL_LINE },
];
const skillsOf = (host: string) => (BUILTIN_SKILLS as any[]).filter((s) => !s.hosts || s.hosts.includes(host));
const commonOf = (host: string) => createCommonTools({ host: host as any, skills: skillsOf(host), shareContext: true });
const req = (G: ToolGroup[], text: string, o: { skills?: string[]; hints?: string[] } = {}) => [...requestGroups(G, { text, ...o })].sort();
const sizeOf = (system: string, specs: ToolSpec[]) =>
  system.length + specs.reduce((n, t) => n + JSON.stringify({ name: t.name, description: t.description, input_schema: t.parameters || {} }).length, 0);

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
/** Bộ công cụ giả: đúng danh sách công cụ và lời nhắc thật, công cụ nào cũng trả ok */
const fake = (h: (typeof HOSTS)[number], executed: string[]): HostToolSet => ({
  host: h.host,
  specs: h.tools.specs,
  execute: async (name) => (executed.push(name), { ok: true }),
  overview: async () => ({ sheets: 1 }),
  systemPrompt: () => h.prompt,
});

describe("Đợt 11 — nhóm công cụ Excel, PowerPoint", () => {
  for (const h of HOSTS) {
    it(`${h.host}: mọi công cụ khai trong nhóm đều có thật; không công cụ nào thuộc hai nhóm; id nhóm là chuỗi chữ thường`, () => {
      const real = new Set([...h.tools.specs, ...commonOf(h.host).specs].map((t) => t.name));
      const seen = new Set<string>();
      for (const g of h.G) {
        assert.match(g.id, /^[a-z_]+$/);
        for (const t of g.tools) {
          assert.ok(real.has(t), `${g.id}: không có công cụ ${t}`);
          assert.ok(!seen.has(t), `${t} nằm ở hai nhóm`);
          seen.add(t);
        }
      }
    });
    it(`${h.host}: read_attachment_pages, use_skill luôn được gửi (nhóm lõi)`, () => {
      const names = visibleSpecs([...h.tools.specs, ...withLoadTools(commonOf(h.host), h.G).specs], h.G, new Set()).map((t) => t.name);
      for (const n of ["read_attachment_pages", "use_skill", LOAD_TOOLS]) assert.ok(names.includes(n), n);
    });
  }
  it("Excel: nhóm lõi giữ đọc, ghi, định dạng, tìm, quét lỗi, lần công thức, xem đối tượng có sẵn; ẩn biểu đồ, PivotTable, in, mô hình", () => {
    const names = visibleSpecs(excelTools.specs, EXCEL_TOOL_GROUPS, new Set()).map((t) => t.name);
    for (const n of ["excel_get_overview", "excel_read_range", "excel_write_range", "excel_format_range", "excel_find", "excel_find_errors", "excel_trace", "excel_list_objects", "excel_insert_delete", "excel_sheet", "excel_select_range"])
      assert.ok(names.includes(n), n);
    for (const n of ["excel_create_chart", "excel_create_pivot", "excel_page_setup", "excel_what_if", "excel_protect", "excel_filter"]) assert.ok(!names.includes(n), n);
  });
  it("PowerPoint: nhóm lõi giữ thêm slide, sửa chữ, hình, bảng mới, định dạng, chụp slide, đọc mẫu; ẩn biểu đồ, sắp xếp, kiểm bài, ảnh", () => {
    const names = visibleSpecs(pptTools.specs, PPT_TOOL_GROUPS, new Set()).map((t) => t.name);
    for (const n of ["ppt_get_overview", "ppt_add_slide", "ppt_add_slides", "ppt_set_shape_text", "ppt_add_table", "ppt_update_shape", "ppt_format_text", "ppt_get_slide_image", "ppt_get_theme", "ppt_get_selection"])
      assert.ok(names.includes(n), n);
    for (const n of ["ppt_add_chart", "ppt_add_process_flow", "ppt_move_slide", "ppt_check_deck", "ppt_insert_image", "ppt_edit_table"]) assert.ok(!names.includes(n), n);
  });
  it("Excel: từ khóa bật đúng nhóm", () => {
    const E = EXCEL_TOOL_GROUPS;
    assert.deepStrictEqual(req(E, "Vẽ biểu đồ cột doanh thu"), ["bieu_do"]);
    assert.deepStrictEqual(req(E, "Lập bảng tổng hợp theo khu vực"), ["pivot"]);
    assert.deepStrictEqual(req(E, "Lọc các dòng có doanh thu > 100"), ["loc_dinh_dang"]);
    assert.deepStrictEqual(req(E, "Thêm danh sách chọn cho cột Trạng thái"), ["loc_dinh_dang"]);
    assert.deepStrictEqual(req(E, "Chuẩn bị in khổ A4 ngang, vừa 1 trang"), ["in_an"]);
    assert.deepStrictEqual(req(E, "Tìm giá bán để IRR đạt 15%"), ["mo_hinh"]);
    assert.deepStrictEqual(req(E, "Xóa các dòng trùng"), ["lam_sach_khoa"]);
    assert.deepStrictEqual(req(E, "Khóa các ô công thức"), ["lam_sach_khoa"]);
    assert.deepStrictEqual(req(E, "Lấy nội dung Word đang mở"), ["ngoai"]);
  });
  it("Excel: câu thường không bật nhóm (trung bình, in đậm, tên riêng Lộc, khoa)", () => {
    const E = EXCEL_TOOL_GROUPS;
    for (const s of ["Tính tổng doanh thu theo tháng", "Tính trung bình cột C", "In đậm dòng tiêu đề", "Anh Lộc gửi số liệu, cập nhật giúp", "Khoa Kế toán cần bảng lương"]) assert.deepStrictEqual(req(E, s), [], s);
  });
  it("PowerPoint: từ khóa bật đúng nhóm; câu thường không bật (anh, lời)", () => {
    const P = PPT_TOOL_GROUPS;
    assert.deepStrictEqual(req(P, "Thêm biểu đồ doanh thu"), ["bieu_do_so_do"]);
    assert.deepStrictEqual(req(P, "Vẽ quy trình 5 bước"), ["bieu_do_so_do"]);
    assert.deepStrictEqual(req(P, "Đổi thứ tự slide 3 lên đầu"), ["sap_xep"]);
    assert.deepStrictEqual(req(P, "Kiểm tra bài trước khi gửi"), ["kiem_bai"]);
    assert.deepStrictEqual(req(P, "Chèn ảnh vào slide 2"), ["hinh_anh"]);
    assert.deepStrictEqual(req(P, "Sửa bảng ở slide 4"), ["bang"]);
    assert.deepStrictEqual(req(P, "Làm slide từ văn bản Word đang mở"), ["ngoai"]);
    for (const s of ["Tạo 3 slide giới thiệu dự án", "Anh Dũng nhờ sửa tiêu đề slide 2", "Lời mở đầu cho slide 1", "Viết lại nội dung slide 3 cho gọn"]) assert.deepStrictEqual(req(P, s), [], s);
  });
  it("PowerPoint: ảnh đính kèm bật nhóm chèn ảnh (gợi ý hinh_anh); gợi ý nhóm không có ở ứng dụng thì bỏ qua", () => {
    assert.deepStrictEqual(req(PPT_TOOL_GROUPS, "đặt vào slide 2", { hints: ["hinh_anh"] }), ["hinh_anh"]);
    assert.deepStrictEqual(req(EXCEL_TOOL_GROUPS, "xem giúp", { hints: ["soat", "hinh_anh"] }), []);
  });
  it("Kỹ năng kéo theo nhóm: mọi công cụ chuyên mà thân kỹ năng Excel, PowerPoint dựng sẵn nhắc tới đều thuộc nhóm kỹ năng đó kéo theo", () => {
    for (const h of HOSTS)
      for (const sk of skillsOf(h.host)) {
        const named = (sk.body + JSON.stringify(sk.files || {})).match(/\b((?:excel|ppt)_[a-z_]+|read_other_app|write_other_app)\b/g) || [];
        for (const t of new Set(named)) {
          const grp = h.G.find((x) => x.tools.includes(t));
          if (grp) assert.ok(grp.skills?.includes(sk.name), `${h.host}: ${sk.name} nhắc ${t} nhưng không kéo nhóm ${grp.id}`);
        }
      }
  });
  it("toolGroupsFor: Miễn phí của cả ba ứng dụng nạp theo việc, Trả phí không", () => {
    assert.strictEqual(toolGroupsFor("free", "Word"), WORD_TOOL_GROUPS);
    assert.strictEqual(toolGroupsFor("free", "Excel"), EXCEL_TOOL_GROUPS);
    assert.strictEqual(toolGroupsFor("free", "PowerPoint"), PPT_TOOL_GROUPS);
    for (const host of ["Word", "Excel", "PowerPoint"]) assert.strictEqual(toolGroupsFor("paid", host), undefined);
    assert.strictEqual(toolGroupsFor("free", "Outlook"), undefined);
  });
  it("Cùng id nhóm ở hai ứng dụng thì cùng nhãn (nhãn bước load_tools không biết ứng dụng)", () => {
    const byId = new Map<string, string>();
    for (const g of ALL_TOOL_GROUPS) {
      if (byId.has(g.id)) assert.strictEqual(g.label, byId.get(g.id), g.id);
      byId.set(g.id, g.label);
    }
    assert.match(describeCall({ id: "1", name: LOAD_TOOLS, args: { group: "pivot" } }), /Nạp thêm công cụ: PivotTable/);
    assert.match(describeCall({ id: "1", name: LOAD_TOOLS, args: { group: "bieu_do_so_do" } }), /biểu đồ, sơ đồ/);
  });
});

describe("Đợt 11 — lời nhắc", () => {
  for (const h of HOSTS) {
    it(`${h.host}, Trả phí: dòng chỉ đường theo tên công cụ nằm nguyên trong lời nhắc; mọi công cụ nhắc tới đều có thật`, () => {
      assert.ok(h.prompt.includes(h.line));
      assert.strictEqual(h.tools.systemPrompt(), h.prompt);
      const names = new Set(h.tools.specs.map((t) => t.name));
      const all = (h.prompt + CORE_GUIDES[h.host]).match(/\b(?:excel|ppt)_[a-z_]+\b/g) || [];
      for (const n of all) assert.ok(names.has(n), `nhắc công cụ không có: ${n}`);
    });
    it(`${h.host}, Miễn phí: lời nhắc hệ thống không nêu tên công cụ nào có thể đang ẩn; có mục nhóm và dòng chỉ đường theo nhóm`, () => {
      const sys = buildSystemPrompt({
        base: freeBasePrompt(h.prompt, h.G),
        guide: hideGroupTools(CORE_GUIDES[h.host], h.G),
        host: h.host,
        skillCatalog: skillCatalog(skillsOf(h.host)),
        toolGroups: groupsNote(h.G),
      });
      for (const t of h.G.flatMap((g) => g.tools)) assert.ok(!sys.includes(t), `lời nhắc Miễn phí còn nêu ${t}`);
      assert.ok(!sys.includes(h.line));
      assert.ok(sys.includes("CHỈ gọi công cụ có trong danh sách") && sys.includes("(xem mục Nhóm công cụ chuyên)"));
      for (const g of h.G) assert.ok(sys.includes(`- ${g.id}: ${g.label}`), g.id);
    });
  }
  it("freeBasePrompt chỉ thay dòng của đúng ứng dụng (nhóm Word không đụng lời nhắc Excel)", () => {
    assert.ok(freeBasePrompt(EXCEL_SYSTEM_PROMPT, WORD_TOOL_GROUPS).includes(EXCEL_SPECIAL_LINE));
  });
  it("Danh mục kỹ năng: mọi kỹ năng dựng sẵn đều có mô tả ngắn ≤ 200 ký tự", () => {
    for (const s of BUILTIN_SKILLS as any[]) assert.ok(s.brief && s.brief.length <= 200, s.name);
  });
});

describe("Đợt 11 — sửa theo soát mã độc lập", () => {
  const fs = require("fs");
  const path = require("path");
  const src = (f: string) => fs.readFileSync(path.join(__dirname, "../../src", f), "utf8") as string;
  const namedIn = (text: string, G: ToolGroup[]) => G.flatMap((g) => g.tools.filter((t) => new RegExp(`\\b${t}\\b`).test(text)).map((t) => `${g.id}:${t}`));
  it("#1: chỗ sửa mà ppt_check_deck gợi ý chỉ nêu công cụ lõi hoặc cùng nhóm kiem_bai", () => {
    assert.deepStrictEqual(namedIn(src("core/deckcheck.ts"), PPT_TOOL_GROUPS.filter((g) => g.id !== "kiem_bai")), []);
  });
  it("#2: câu lỗi của lớp nhật ký Excel (bọc mọi công cụ, kể cả lõi) không nêu tên công cụ chuyên", () => {
    assert.deepStrictEqual(namedIn(src("office/excelTrail.ts"), EXCEL_TOOL_GROUPS), []);
  });
  it("#3, #5: PivotTable viết liền, thả xuống, đổi vị trí slide bật đúng nhóm; #6: gia đình, đặt tên sheet không bật mô hình", () => {
    assert.deepStrictEqual(req(EXCEL_TOOL_GROUPS, "Tạo PivotTable từ bảng A1:D100"), ["pivot"]);
    assert.deepStrictEqual(req(EXCEL_TOOL_GROUPS, "Thêm danh sách thả xuống cho cột C"), ["loc_dinh_dang"]);
    assert.deepStrictEqual(req(PPT_TOOL_GROUPS, "Hoán đổi slide 2 và 3"), ["sap_xep"]);
    assert.deepStrictEqual(req(EXCEL_TOOL_GROUPS, "Lập bảng chi tiêu gia đình tháng này"), []);
    assert.deepStrictEqual(req(EXCEL_TOOL_GROUPS, "Đặt tên sheet là Tổng hợp"), []);
  });
  it("#7: câu ghi chú ảnh đính kèm không nêu tên công cụ chèn ảnh", () => {
    assert.ok(!/Muốn đặt ảnh vào slide thì gọi ppt_insert_image/.test(src("taskpane/components/Chat.tsx")));
  });
});

describe("Đợt 11 — chạy một lượt ở chế độ Miễn phí", () => {
  it("Excel: bước đầu chỉ gửi nhóm lõi + load_tools; model xin nhóm pivot thì bước sau có công cụ PivotTable và chạy được", async () => {
    const executed: string[] = [];
    const g = fakeGroq([
      () => oaTools([{ id: "a", name: LOAD_TOOLS, args: { group: "pivot" } }]),
      () => oaTools([{ id: "b", name: "excel_create_pivot", args: {} }]),
      () => oaText("Xong"),
    ]);
    const res = await runTurn({
      provider: getProvider("groq")!,
      key: "K",
      model: "m",
      host: "Excel",
      history: [],
      text: "Gom số liệu theo khu vực giúp tôi", // cố ý không khớp từ khóa nào
      onEvent: () => {},
      fetchFn: g.fetchFn,
      agent: { tools: fake(HOSTS[0], executed), guide: CORE_GUIDES.Excel, log: [], toolGroups: EXCEL_TOOL_GROUPS },
    });
    assert.strictEqual(res.error, undefined);
    const first = sentTools(g.bodies[0]);
    assert.ok(first.includes(LOAD_TOOLS) && first.includes("excel_read_range"));
    assert.ok(!first.includes("excel_create_pivot") && !first.includes("excel_page_setup"));
    const second = sentTools(g.bodies[1]);
    assert.ok(second.includes("excel_create_pivot") && second.includes("excel_edit_pivot") && !second.includes("excel_page_setup"));
    assert.deepStrictEqual(executed, ["excel_create_pivot"]);
    const sys = g.bodies[0].messages[0].content as string;
    assert.ok(sys.includes("- pivot: PivotTable"));
    for (const t of EXCEL_TOOL_GROUPS.flatMap((x) => x.tools)) assert.ok(!sys.includes(t), `lời nhắc gửi đi còn nêu ${t}`);
  });
  it("PowerPoint: câu khớp từ khóa bật nhóm ngay từ bước đầu; chuẩn làm việc gửi đi không nêu công cụ ẩn", async () => {
    const g = fakeGroq([() => oaText("OK")]);
    await runTurn({
      provider: getProvider("groq")!,
      key: "K",
      model: "m",
      host: "PowerPoint",
      history: [],
      text: "Thêm slide biểu đồ doanh thu 3 năm",
      onEvent: () => {},
      fetchFn: g.fetchFn,
      agent: { tools: fake(HOSTS[1], []), guide: CORE_GUIDES.PowerPoint, log: [], toolGroups: PPT_TOOL_GROUPS },
    });
    const t = sentTools(g.bodies[0]);
    assert.ok(t.includes("ppt_add_chart") && !t.includes("ppt_check_deck"));
    const sys = g.bodies[0].messages[0].content as string;
    for (const n of ["ppt_check_deck", "ppt_move_slide", "ppt_insert_image"]) assert.ok(!sys.includes(n), n);
  });
  it("Trả phí (không có toolGroups): gửi đủ công cụ PowerPoint, không có load_tools", async () => {
    const g = fakeGroq([() => oaText("OK")]);
    await runTurn({ provider: getProvider("groq")!, key: "K", model: "m", host: "PowerPoint", history: [], text: "Xin chào", onEvent: () => {}, fetchFn: g.fetchFn, agent: { tools: fake(HOSTS[1], []), log: [] } });
    const t = sentTools(g.bodies[0]);
    assert.strictEqual(t.length, pptTools.specs.length);
    assert.ok(!t.includes(LOAD_TOOLS));
  });
});

describe("Đợt 11 — chặn phình lại (số ký tự phần cố định của một lượt)", () => {
  // Mốc 4.10.0: Excel 31.754, PowerPoint 23.293 (tasks/dot-11-bang-chung/do-token-*-moc-4100.json)
  // Đợt 12 (22/11/2026): Excel nâng trần +400 / +350 cho tham số `more` của excel_format_range (định dạng nhiều vùng khác kiểu trong một lệnh,
  // bài X6 chạm trần 12 bước vì 7 lệnh định dạng rời) — đo: Trả phí 25.891, Miễn phí 15.135
  const CAP: Record<string, { paid: number; free: number }> = { Excel: { paid: 26000, free: 15250 }, PowerPoint: { paid: 18900, free: 14100 } };
  for (const h of HOSTS) {
    const parts = { base: h.prompt, guide: CORE_GUIDES[h.host], host: h.host, skillCatalog: skillCatalog(skillsOf(h.host)) };
    it(`${h.host}, Trả phí: ≤ ${CAP[h.host].paid} ký tự`, () => {
      const n = sizeOf(buildSystemPrompt(parts), [...h.tools.specs, ...commonOf(h.host).specs]);
      assert.ok(n <= CAP[h.host].paid, `đang ${n}`);
    });
    it(`${h.host}, Miễn phí, nhóm lõi: ≤ ${CAP[h.host].free} ký tự`, () => {
      const ex = withLoadTools(commonOf(h.host), h.G);
      const sys = buildSystemPrompt({ ...parts, base: freeBasePrompt(h.prompt, h.G), guide: hideGroupTools(parts.guide, h.G), toolGroups: groupsNote(h.G) });
      const n = sizeOf(sys, visibleSpecs([...h.tools.specs, ...ex.specs], h.G, new Set()));
      assert.ok(n <= CAP[h.host].free, `đang ${n}`);
    });
  }
});
