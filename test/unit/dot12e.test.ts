/** Đợt 12 phần E: bộ nhớ dài hạn giữa các hội thoại (tasks/dot-12-bo-nho.md) — tiêu chí E1 */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { runTurn } from "../../src/core/agent/turn";
import { buildPlainSystem, buildSystemPrompt } from "../../src/core/agent/prompt";
import { HostToolSet } from "../../src/core/agent/types";
import {
  MEMORY_CAP_FREE,
  MEMORY_CAP_PAID,
  MEMORY_HEADING,
  MEMORY_ITEM_MAX,
  MEMORY_MAX_ITEMS,
  MemoryItem,
  blockedReason,
  buildMemoryBlock,
  memoryCapFor,
  sanitizeMemory,
  validateMemoryText,
} from "../../src/core/memory";
import { Storage } from "../../src/core/storage";
import { createCommonTools } from "../../src/office/common";
import { EXCEL_TOOL_GROUPS, PPT_TOOL_GROUPS, WORD_TOOL_GROUPS, requestGroups, visibleSpecs, withLoadTools } from "../../src/core/toolgroups";
import { BUILTIN_SKILLS } from "../../src/core/skillpack";
import { MemoryKV } from "../../src/core/db";
import { createBackup, restoreBackup, summarizeBackup } from "../../src/core/backup";

const item = (text: string, o: Partial<MemoryItem> = {}): MemoryItem => ({ id: "m" + text.length + text.slice(0, 3), text, scope: "all", source: "user", createdAt: 1, ...o });

function memStore() {
  const d: Record<string, string> = {};
  return { getItem: (k: string) => (k in d ? d[k] : null), setItem: (k: string, v: string) => void (d[k] = v), removeItem: (k: string) => void delete d[k], d };
}

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200, headers: { "content-type": "text/event-stream" } });
}
const oaText = (t: string) => sse([{ choices: [{ delta: { content: t } }] }, { choices: [{ delta: {}, finish_reason: "stop" }] }, "[DONE]"]);
function fakeGroq(n: number) {
  const bodies: any[] = [];
  const fetchFn = (async (_url: string, init: any) => {
    bodies.push(JSON.parse(init.body));
    if (bodies.length > n) throw new Error("Bài thử không chờ thêm lần gọi");
    return oaText("OK");
  }) as any;
  return { fetchFn, bodies };
}
const fakeTools: HostToolSet = {
  host: "Word",
  specs: [{ name: "word_read", description: "đọc", parameters: { type: "object", properties: {} } }],
  execute: async () => ({ ok: true }),
  overview: async () => ({ paragraphs: 1 }),
  systemPrompt: () => "LỜI NHẮC WORD",
};

describe("Đợt 12 E — bộ nhớ vào lời nhắc đúng chỗ", () => {
  const mem = buildMemoryBlock([item("Trả lời kết luận trước.")], "Word", MEMORY_CAP_PAID).text;
  it("tác tử: sau chỉ dẫn riêng, trước danh mục kỹ năng; không có bộ nhớ thì lời nhắc như cũ", () => {
    const sys = buildSystemPrompt({ base: "BASE", host: "Word", instructions: "CHỈ DẪN", memory: mem, skillCatalog: "## DANH MỤC KỸ NĂNG" });
    const a = sys.indexOf("CHỈ DẪN");
    const b = sys.indexOf(MEMORY_HEADING);
    const c = sys.indexOf("## DANH MỤC KỸ NĂNG");
    assert.ok(a > 0 && b > a && c > b, sys);
    assert.ok(sys.includes("- Trả lời kết luận trước."));
    assert.strictEqual(buildSystemPrompt({ base: "BASE", host: "Word", memory: "" }), buildSystemPrompt({ base: "BASE", host: "Word" }));
  });
  it("trò chuyện thường: có bộ nhớ sau chỉ dẫn riêng", () => {
    const sys = buildPlainSystem("CHỈ DẪN", mem);
    assert.ok(sys.indexOf(MEMORY_HEADING) > sys.indexOf("CHỈ DẪN"));
    assert.strictEqual(buildPlainSystem("X", undefined), buildPlainSystem("X"));
  });
  it("runTurn: hai lượt cùng hội thoại gửi đúng một lời nhắc hệ thống có bộ nhớ (đóng băng, bộ nhớ đệm giữ nguyên)", async () => {
    const g = fakeGroq(2);
    const base = { provider: getProvider("groq")!, key: "K", model: "m", host: "Word", onEvent: () => {}, fetchFn: g.fetchFn, memory: mem, agent: { tools: fakeTools, log: [] } };
    const r1 = await runTurn({ ...base, history: [], text: "Câu 1" });
    await runTurn({ ...base, history: r1.history, text: "Câu 2" });
    const s1 = g.bodies[0].messages[0].content;
    const s2 = g.bodies[1].messages[0].content;
    assert.ok(String(s1).includes("Trả lời kết luận trước."));
    assert.strictEqual(s1, s2);
  });
  it("khối bộ nhớ không phụ thuộc thời điểm dựng (cùng danh sách → cùng chuỗi)", () => {
    const list = [item("A một"), item("B hai", { id: "x2", createdAt: 2 })];
    assert.strictEqual(buildMemoryBlock(list, "Word", 2000).text, buildMemoryBlock(list.slice(), "Word", 2000).text);
  });
});

describe("Đợt 12 E — chọn mục và giới hạn độ dài", () => {
  it("lấy mục chung + mục của ứng dụng, bỏ mục đang tắt và mục của ứng dụng khác; cũ trước", () => {
    const list = [
      item("Chung mới", { id: "a", createdAt: 5 }),
      item("Chỉ Excel", { id: "b", scope: "Excel", createdAt: 1 }),
      item("Chỉ Word", { id: "c", scope: "Word", createdAt: 3 }),
      item("Đang tắt", { id: "d", off: true }),
      item("Chung cũ", { id: "e", createdAt: 0 }),
    ];
    const b = buildMemoryBlock(list, "Word", 2000);
    assert.deepStrictEqual(b.text.split("\n").slice(1), ["- Chung cũ", "- Chỉ Word", "- Chung mới"]);
    assert.strictEqual(b.used, 3);
    assert.strictEqual(buildMemoryBlock([], "Word", 2000).text, "");
    assert.strictEqual(buildMemoryBlock([item("x", { off: true })], "Word", 2000).text, "");
  });
  it("trần: Trả phí 2.000, Miễn phí 800 ký tự; khối không bao giờ vượt trần; báo số mục bỏ", () => {
    assert.strictEqual(memoryCapFor("free"), MEMORY_CAP_FREE);
    assert.strictEqual(memoryCapFor("paid"), MEMORY_CAP_PAID);
    assert.strictEqual(memoryCapFor(undefined), MEMORY_CAP_PAID);
    const list = Array.from({ length: 20 }, (_, i) => item(`Quy ước số ${i} `.padEnd(280, "x"), { id: "k" + i, createdAt: i }));
    for (const cap of [MEMORY_CAP_FREE, MEMORY_CAP_PAID]) {
      const b = buildMemoryBlock(list, "Word", cap);
      assert.ok(b.text.length <= cap, `${b.text.length} > ${cap}`);
      assert.ok(b.used >= 1 && b.used + b.skipped === 20);
    }
    assert.ok(buildMemoryBlock(list, "Word", MEMORY_CAP_FREE).used < buildMemoryBlock(list, "Word", MEMORY_CAP_PAID).used);
  });
  it("mục > 300 ký tự, trùng, đủ 60 mục: không nhận", () => {
    assert.match(validateMemoryText("a".repeat(MEMORY_ITEM_MAX + 1))!, /tối đa 300/);
    assert.strictEqual(validateMemoryText("a".repeat(MEMORY_ITEM_MAX)), undefined);
    assert.match(validateMemoryText("Trả lời ngắn", [item("trả lời ngắn")], "all")!, /Đã có/);
    assert.strictEqual(validateMemoryText("Trả lời ngắn", [item("trả lời ngắn", { scope: "Excel" })], "all"), undefined);
    const full = Array.from({ length: MEMORY_MAX_ITEMS }, (_, i) => item("m" + i, { id: "i" + i }));
    assert.match(validateMemoryText("mới", full)!, /đủ 60/);
    assert.match(validateMemoryText("")!, /trống/);
  });
});

describe("Đợt 12 E — bộ lọc key và số định danh", () => {
  const mustBlock = [
    "Key OpenRouter của tôi là sk-or-v1-0123456789abcdef0123",
    "Dùng key AIzaSyA1234567890abcdefghijk",
    "Groq gsk_abcdefghijklmnop1234",
    "Authorization: Bearer abcdef1234567890xyz",
    "api key: abc123def456",
    "mật khẩu: Hoang@2026",
    "Mã truy cập 9f8e7d6c5b4a39281706f5e4d3c2b1a0ffee",
    "Gọi anh Nam số 0912345678",
    "SĐT chị Lan 0912 345 678",
    "Điện thoại +84 912 345 678",
    "Số máy 090.123.4567",
    "CCCD của giám đốc 001085012345",
    "CMND 123456789",
    "CCCD 001 085 012 345",
    "Tài khoản 19034567890123 tại Techcombank",
    "Thẻ 4111 1111 1111 1111",
  ];
  const mustPass = [
    "Căn cứ Nghị định 30/2020/NĐ-CP về công tác văn thư",
    "Quyết định số 1234/QĐ-UBND ngày 23/11/2026",
    "Giá trị hợp đồng ghi dạng 1.500.000.000 đồng (bằng chữ kèm theo)",
    "Số tiền đơn vị triệu đồng, ví dụ 12.345,6 triệu",
    "Dự án H = Khu đô thị mới phía Tây, diện tích 25,6 ha",
    "Trả lời kết luận trước, dùng bảng Tiêu chí · Kết quả",
    "Công văn ký thừa lệnh Giám đốc; số ký hiệu dạng 123/CV-PL",
    "Khoản 2 Điều 15; Điều 0912 không tồn tại",
    "Giai đoạn 2026 2027 2028 ưu tiên pháp lý",
    "Tỷ lệ 0,912345678 chỉ là ví dụ số thập phân",
    "Mã dự án tiếng Việt viết thường, gạch nối: du-an-khu-do-thi-phia-tay",
  ];
  for (const t of mustBlock) it(`chặn: ${t}`, () => assert.ok(blockedReason(t), t));
  for (const t of mustPass) it(`không chặn nhầm: ${t}`, () => assert.strictEqual(blockedReason(t), undefined, t));
});

describe("Đợt 12 E — kho lưu", () => {
  it("addMemory lưu, chặn key, sanitize bỏ mục hỏng; không ảnh hưởng cài đặt", () => {
    const st = new Storage(memStore());
    const r = st.addMemory("  Trả lời   kết luận trước ", "all", "user", 100);
    assert.ok(r.ok);
    assert.strictEqual(st.getMemory()[0].text, "Trả lời kết luận trước");
    const bad = st.addMemory("key sk-ant-api03-abcdefghijklmnop", "all", "ai");
    assert.ok("error" in bad && /key/.test(bad.error));
    assert.strictEqual(st.getMemory().length, 1);
    assert.deepStrictEqual(sanitizeMemory([{ id: 1, text: "x" }, { id: "a", text: "" }, null, { id: "b", text: "ok", scope: "Lạ", source: "khác" }]), [
      { id: "b", text: "ok", scope: "all", source: "user", createdAt: 0 },
    ]);
    assert.deepStrictEqual(st.getSettings().instructions, {});
  });
});

describe("Đợt 12 E — công cụ remember (AI chỉ đề xuất)", () => {
  it("đưa đề xuất lên khung, không tự lưu; this_app → ứng dụng hiện tại", async () => {
    const got: any[] = [];
    const ct = createCommonTools({ host: "Excel", skills: [], shareContext: false, onRemember: (p) => got.push(p), memory: () => [] });
    const r: any = await ct.execute("remember", { text: "Số tiền ghi đơn vị triệu đồng", scope: "this_app" });
    assert.ok(r.ok && /hội thoại sau/.test(r.message));
    assert.deepStrictEqual(got, [{ text: "Số tiền ghi đơn vị triệu đồng", scope: "Excel" }]);
    await ct.execute("remember", { text: "Kết luận trước" });
    assert.strictEqual(got[1].scope, "all");
  });
  it("chặn key / số định danh, mục trùng; không có khung nhận thì báo lỗi", async () => {
    const got: any[] = [];
    const ct = createCommonTools({ host: "Word", skills: [], shareContext: false, onRemember: (p) => got.push(p), memory: () => [item("Kết luận trước")] });
    await assert.rejects(() => ct.execute("remember", { text: "CCCD 001085012345" }), /số/);
    await assert.rejects(() => ct.execute("remember", { text: "kết luận trước" }), /Đã có/);
    assert.strictEqual(got.length, 0);
    const none = createCommonTools({ host: "Word", skills: [], shareContext: false });
    await assert.rejects(() => none.execute("remember", { text: "x" }), /không dùng được/);
  });
  it("Trả phí: luôn có trong bộ công cụ; Miễn phí: chỉ gửi khi người dùng nói tới quy ước lâu dài (nhóm bo_nho)", () => {
    for (const [host, G] of [["Word", WORD_TOOL_GROUPS], ["Excel", EXCEL_TOOL_GROUPS], ["PowerPoint", PPT_TOOL_GROUPS]] as const) {
      const ct = createCommonTools({ host, skills: BUILTIN_SKILLS as any, shareContext: true });
      assert.ok(ct.specs.some((t) => t.name === "remember"), host);
      const ex = withLoadTools(ct, G as any);
      const core = visibleSpecs(ex.specs, G as any, new Set()).map((t) => t.name);
      assert.ok(!core.includes("remember"), host);
      assert.ok(requestGroups(G as any, { text: "Từ giờ luôn trả lời kết luận trước nhé" }).has("bo_nho"), host);
      assert.ok(requestGroups(G as any, { text: "Ghi nhớ: dự án H là khu đô thị phía Tây" }).has("bo_nho"), host);
      assert.ok(!requestGroups(G as any, { text: "Sửa giúp khoản 2 Điều 5 cho gọn" }).has("bo_nho"), host);
    }
  });
});

describe("Đợt 12 E — sao lưu", () => {
  it("bản sao lưu có bộ nhớ; khôi phục gộp giữ mục hiện có, thêm mục chỉ có trong tệp", async () => {
    const mk = (list: MemoryItem[]) => {
      const d: Record<string, string> = { "aidn.memory": JSON.stringify(list) };
      return {
        d,
        get length() {
          return Object.keys(d).length;
        },
        key: (i: number) => Object.keys(d)[i] ?? null,
        getItem: (k: string) => (k in d ? d[k] : null),
        setItem: (k: string, v: string) => void (d[k] = v),
        removeItem: (k: string) => void delete d[k],
      };
    };
    const src = mk([item("A", { id: "a" }), item("B", { id: "b" })]);
    const b = await createBackup(src, new MemoryKV(), { appVersion: "t" });
    assert.strictEqual(summarizeBackup(b).memory, 2);
    const dst = mk([item("A sửa", { id: "a" }), item("C", { id: "c" })]);
    await restoreBackup(b, dst, new MemoryKV(), "merge");
    const after = JSON.parse(dst.d["aidn.memory"]).map((x: any) => x.id + ":" + x.text);
    assert.deepStrictEqual(after, ["a:A sửa", "c:C", "b:B"]);
  });
});
