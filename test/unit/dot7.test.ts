import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { runAgent, stepCapFor } from "../../src/core/agent/loop";
import { AgentEvent, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { ConversationStore, MAX_CONVERSATIONS, TEST_MODE_KEY, testModeOn } from "../../src/core/conversations";
import { MAX_BULK_CELL_ROWS, MAX_CELL_ROWS, buildLogRows } from "../../src/core/auditlog";
import { fitInBox, imageBox, placeRegion } from "../../src/core/imagefit";
import { clearAttachedFiles, findAttachedFile, findAttachedImage, listAttachedImages, registerAttachedFile } from "../../src/core/attachedFiles";
import { lockLayoutRisk } from "../../src/office/excelModel";
import { MAX_SLIDES_PER_CALL, normalizeSlideList } from "../../src/office/pptTools";

const P = (id: string) => getProvider(id)!;

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200 });
}
const toolCall = (name: string, args: string, id = "c") => sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: args } }] } }] }]);
const say = (text: string) => sse([{ choices: [{ delta: { content: text } }] }]);

const TOOLS: ToolSpec[] = [
  { name: "doc", description: "đọc", parameters: { type: "object", properties: {} } },
  { name: "ghi", description: "ghi", parameters: { type: "object", properties: { o: { type: "string" } } }, write: true },
];
function tools(run: (name: string, args: any) => any): HostToolSet {
  return { host: "Excel", specs: TOOLS, overview: async () => ({}), systemPrompt: () => "S", execute: async (name, args) => run(name, args) };
}
const base = (over: any) => ({ provider: P("groq"), key: "K", model: "m", system: "S", messages: [{ role: "user" as const, content: "x" }], log: [], ...over });

describe("Đợt 7 · A: trần số bước và dừng sớm", () => {
  it("Trần theo ứng dụng: Word 12, Excel 12, PowerPoint 20", () => {
    assert.strictEqual(stepCapFor("Word"), 12);
    assert.strictEqual(stepCapFor("Excel"), 12);
    assert.strictEqual(stepCapFor("PowerPoint"), 20);
    assert.strictEqual(stepCapFor("Khác"), 20);
  });
  it("Chạm trần: sự kiện limit mang số lệnh ghi thành công; lịch sử kết thúc bằng mốc chờ làm tiếp", async () => {
    let n = 0;
    const events: AgentEvent[] = [];
    const msgs = await runAgent(
      base({
        tools: tools((name) => (name === "ghi" ? { ok: true } : { ok: true, data: 1 })),
        maxSteps: 4,
        onEvent: (e: AgentEvent) => events.push(e),
        // Ba bước ghi (tham số khác nhau), một bước đọc
        fetchFn: (async () => (++n === 2 ? toolCall("doc", "{}", "c" + n) : toolCall("ghi", `{"o":"A${n}"}`, "c" + n))) as any,
      })
    );
    const lim = events[events.length - 1] as any;
    assert.strictEqual(lim.type, "limit");
    assert.strictEqual(lim.maxSteps, 4);
    assert.strictEqual(lim.writes, 3);
    assert.match(String(msgs[msgs.length - 1].content), /giới hạn 4 bước/);
    // Mọi lệnh gọi đều có kết quả
    const calls = msgs.filter((m) => m.role === "assistant").flatMap((m: any) => m.toolCalls || []).length;
    assert.strictEqual(msgs.filter((m) => m.role === "tool").length, calls);
  });
  it("Lệnh ghi trả ok=false không được tính là đã ghi", async () => {
    let n = 0;
    const events: AgentEvent[] = [];
    await runAgent(
      base({
        tools: tools(() => ({ ok: false, error: "thay 0 chỗ" })),
        maxSteps: 2,
        onEvent: (e: AgentEvent) => events.push(e),
        fetchFn: (async () => toolCall("ghi", `{"o":"A${++n}"}`, "c" + n)) as any,
      })
    );
    assert.strictEqual((events[events.length - 1] as any).writes, 0);
  });
  it("Cùng công cụ, cùng tham số, cùng lỗi 2 bước liên tiếp → dừng sớm (stuck), không đợi hết trần", async () => {
    const events: AgentEvent[] = [];
    let n = 0;
    const msgs = await runAgent(
      base({
        tools: tools(() => {
          throw new Error("Không có sheet Tổng hợp");
        }),
        maxSteps: 12,
        onEvent: (e: AgentEvent) => events.push(e),
        fetchFn: (async () => toolCall("ghi", '{"o":"A1"}', "c" + ++n)) as any,
      })
    );
    const last = events[events.length - 1] as any;
    assert.strictEqual(last.type, "stuck");
    assert.strictEqual(last.tool, "ghi");
    assert.strictEqual(last.step, 2);
    assert.strictEqual(events.filter((e) => e.type === "step").length, 2);
    assert.match(String(msgs[msgs.length - 1].content), /cùng một lỗi 2 bước liên tiếp/);
    assert.strictEqual(msgs.filter((m) => m.role === "tool").length, 2);
  });
  it("Cùng lỗi nhưng tham số có đổi → cho thử tới bước thứ 3 mới dừng", async () => {
    const events: AgentEvent[] = [];
    let n = 0;
    await runAgent(
      base({
        tools: tools(() => {
          throw new Error("Địa chỉ không hợp lệ");
        }),
        maxSteps: 12,
        onEvent: (e: AgentEvent) => events.push(e),
        fetchFn: (async () => toolCall("ghi", `{"o":"A${++n}"}`, "c" + n)) as any,
      })
    );
    const last = events[events.length - 1] as any;
    assert.strictEqual(last.type, "stuck");
    assert.strictEqual(last.step, 3);
  });
  it("Lỗi rồi sửa được ở bước sau → không dừng sớm; lỗi khác nhau xen kẽ cũng không", async () => {
    const events: AgentEvent[] = [];
    let n = 0;
    let run = 0;
    await runAgent(
      base({
        tools: tools(() => {
          run++;
          if (run === 1) throw new Error("lỗi A");
          if (run === 2) return { ok: true };
          if (run === 3) throw new Error("lỗi A");
          return { ok: true };
        }),
        maxSteps: 12,
        onEvent: (e: AgentEvent) => events.push(e),
        fetchFn: (async () => (++n <= 4 ? toolCall("ghi", '{"o":"A1"}', "c" + n) : say("xong"))) as any,
      })
    );
    assert.ok(!events.some((e) => e.type === "stuck" || e.type === "limit"));
  });
});

describe("Đợt 7 · E: chế độ thử và lịch sử đầy", () => {
  const memKV = () => {
    const m = new Map<string, any>();
    return { get: async (k: string) => m.get(k), set: async (k: string, v: any) => void m.set(k, v), del: async (k: string) => void m.delete(k), update: async (k: string, fn: (v: any) => any) => { const v = fn(m.get(k)); m.set(k, v); return v; }, m } as any;
  };
  const conv = (id: string, at = 0) => ({ id, host: "Word", title: "T " + id, updatedAt: at, items: [{ role: "user", text: "x" }], history: [] });

  it("testModeOn đọc đúng khóa; kho lỗi thì coi như tắt", () => {
    assert.strictEqual(testModeOn({ getItem: (k) => (k === TEST_MODE_KEY ? "1" : null) }), true);
    assert.strictEqual(testModeOn({ getItem: () => null }), false);
    assert.strictEqual(testModeOn({ getItem: () => "0" }), false);
    assert.strictEqual(
      testModeOn({
        getItem: () => {
          throw new Error("chặn");
        },
      }),
      false
    );
  });
  it("Chế độ thử bật: 3 hội thoại không làm đổi lịch sử; tắt: lưu như cũ", async () => {
    const kv = memKV();
    let on = false;
    const store = new ConversationStore(kv, () => on);
    await store.save(conv("that-1"));
    on = true;
    for (const id of ["thu-1", "thu-2", "thu-3"]) await store.save(conv(id));
    assert.deepStrictEqual((await store.list()).map((c) => c.id), ["that-1"]);
    assert.strictEqual(await store.get("thu-1"), undefined);
    assert.strictEqual(await store.nextToDrop(), undefined);
    on = false;
    await store.save(conv("that-2"));
    assert.strictEqual((await store.list()).length, 2);
  });
  it("Lịch sử đủ trần: nextToDrop nêu đúng hội thoại cũ nhất; hội thoại đang có trong lịch sử thì không báo", async () => {
    const kv = memKV();
    const index = Array.from({ length: MAX_CONVERSATIONS }, (_, i) => ({ id: "c" + i, host: "Word", title: "Hội thoại " + i, updatedAt: 1000 + i }));
    await kv.set("conv-index", index);
    const store = new ConversationStore(kv);
    const d = await store.nextToDrop();
    assert.strictEqual(d?.id, "c0");
    assert.strictEqual(await store.nextToDrop("c5"), undefined);
    await kv.set("conv-index", index.slice(1));
    assert.strictEqual(await store.nextToDrop(), undefined);
  });
});

describe("Đợt 7 · D: nhật ký Excel cho thao tác hàng loạt, khóa ô", () => {
  const changes = (n: number) => Array.from({ length: n }, (_, i) => ({ cell: "A" + (i + 1), before: "cũ" + i, after: "mới" + i }));
  const input = (n: number, maxRows?: number) => ({ time: new Date(2026, 10, 18, 9, 0, 0), sheet: "S", address: "A1:A" + n, changes: changes(n), action: "Thay hàng loạt", maxRows });
  it("5 ô: đủ 5 dòng có giá trị cũ / mới", () => {
    const rows = buildLogRows(input(5, MAX_BULK_CELL_ROWS));
    assert.strictEqual(rows.length, 5);
    assert.deepStrictEqual(rows[4].slice(2, 5), ["A5", "cũ4", "mới4"]);
  });
  it("300 ô: 200 dòng chi tiết + 1 dòng tổng nêu 100 ô còn lại", () => {
    const rows = buildLogRows(input(300, MAX_BULK_CELL_ROWS));
    assert.strictEqual(rows.length, MAX_BULK_CELL_ROWS + 1);
    assert.match(rows[MAX_BULK_CELL_ROWS][4], /100 ô khác/);
  });
  it("Thao tác thường vẫn giữ trần cũ; maxRows lạ (0, âm, quá lớn, không nguyên) không phá trần", () => {
    assert.strictEqual(buildLogRows(input(300)).length, MAX_CELL_ROWS + 1);
    assert.strictEqual(buildLogRows(input(300, 0)).length, MAX_CELL_ROWS + 1);
    assert.strictEqual(buildLogRows(input(300, -5)).length, MAX_CELL_ROWS + 1);
    assert.strictEqual(buildLogRows(input(300, 1.5)).length, MAX_CELL_ROWS + 1);
    assert.strictEqual(buildLogRows(input(300, 99999)).length, MAX_BULK_CELL_ROWS + 1);
  });
  it("Giá trị cũ / mới là công thức vẫn được ghi thành chữ", () => {
    const rows = buildLogRows({ ...input(1, MAX_BULK_CELL_ROWS), changes: [{ cell: "A1", before: "=WEBSERVICE(\"http://x\")", after: "+1" }] });
    assert.ok(rows[0][3].startsWith("'=") && rows[0][4].startsWith("'+"));
  });
  it("Khóa ô: sheet mặc định không hỏi; có ô mở khóa riêng hoặc lẫn lộn thì hỏi kèm lý do", () => {
    assert.strictEqual(lockLayoutRisk("S", true, true), null);
    assert.match(String(lockLayoutRisk("S", null, true)), /bố trí khóa ô riêng/);
    assert.match(String(lockLayoutRisk("S", true, null)), /bố trí khóa ô riêng/);
    assert.match(String(lockLayoutRisk("Bảng", false, false)), /"Bảng"/);
  });
});

describe("Đợt 7 · B: thêm nhiều slide một lệnh", () => {
  it("Chuẩn hóa danh sách: bỏ phần tử rỗng, nhận chuỗi JSON, tối đa 10", () => {
    assert.deepStrictEqual(normalizeSlideList([{ title: "A" }, null, {}, { body: "x\ny", layout: "Title Only" }, "chữ"]), [
      { layout: undefined, title: "A", body: undefined },
      { layout: "Title Only", title: undefined, body: "x\ny" },
    ]);
    assert.strictEqual(normalizeSlideList('[{"title":"A"},{"title":"B"}]').length, 2);
    assert.deepStrictEqual(normalizeSlideList("không phải json"), []);
    assert.deepStrictEqual(normalizeSlideList(undefined), []);
    assert.deepStrictEqual(normalizeSlideList({ title: "A" }), []);
    assert.strictEqual(normalizeSlideList(Array.from({ length: 25 }, (_, i) => ({ title: "S" + i }))).length, MAX_SLIDES_PER_CALL);
    assert.strictEqual(normalizeSlideList([{ title: 5 }])[0].title, "5");
  });
});

describe("Đợt 7 · C: đặt ảnh vào slide", () => {
  const slide = { w: 960, h: 540 };
  const ratio = (b: { width: number; height: number }) => b.width / b.height;
  const inside = (b: any) => b.left >= -0.01 && b.top >= -0.01 && b.left + b.width <= slide.w + 0.01 && b.top + b.height <= slide.h + 0.01;
  it("Giữ đúng tỷ lệ ảnh (lệch dưới 1%) và nằm trong slide ở mọi vị trí, mọi dáng ảnh", () => {
    for (const img of [
      { w: 1600, h: 900 },
      { w: 900, h: 1600 },
      { w: 500, h: 500 },
      { w: 4000, h: 300 },
      { w: 10, h: 2000 },
    ]) {
      for (const place of ["content", "left", "right", "center", "full", "lạ", undefined]) {
        const b = imageBox(img, slide, { place: place as any });
        assert.ok(Math.abs(ratio(b) / (img.w / img.h) - 1) < 0.01, `tỷ lệ ${img.w}x${img.h} ${place}`);
        assert.ok(inside(b), `trong slide ${img.w}x${img.h} ${place}: ${JSON.stringify(b)}`);
        assert.ok(b.width > 0 && b.height > 0);
      }
    }
  });
  it("Vị trí: right nằm nửa phải, left nửa trái, không đè vùng tiêu đề", () => {
    const img = { w: 800, h: 600 };
    const r = imageBox(img, slide, { place: "right", titleBottom: 110 });
    const l = imageBox(img, slide, { place: "left", titleBottom: 110 });
    assert.ok(r.left >= slide.w / 2 && l.left + l.width <= slide.w / 2);
    assert.ok(r.top >= 110 && l.top >= 110);
    assert.ok(placeRegion("content", slide).top >= 40);
    const full = imageBox({ w: 1920, h: 1080 }, slide, { place: "full" });
    assert.deepStrictEqual(full, { left: 0, top: 0, width: 960, height: 540 });
  });
  it("Tham số do model truyền: chỉ width → height theo tỷ lệ; số âm, 0, chuỗi, quá lớn đều được kẹp", () => {
    const img = { w: 1000, h: 500 };
    const a = imageBox(img, slide, { width: 400, left: 100, top: 100 });
    assert.deepStrictEqual(a, { left: 100, top: 100, width: 400, height: 200 });
    const b = imageBox(img, slide, { width: "300" as any });
    assert.strictEqual(b.height, 150);
    const c = imageBox(img, slide, { width: -1, height: 0 });
    assert.ok(c.width > 0 && inside(c) && Math.abs(ratio(c) - 2) < 0.02);
    const d = imageBox(img, slide, { width: 99999, left: -500, top: 99999 });
    assert.ok(inside(d) && Math.abs(ratio(d) - 2) < 0.02);
    const e = imageBox(img, slide, { width: 400, height: 400 });
    assert.ok(Math.abs(ratio(e) - 2) < 0.02 && e.width <= 400 && e.height <= 400);
    const f = imageBox({ w: 0, h: 0 }, slide, {});
    assert.ok(f.width > 0 && f.height > 0 && inside(f));
    assert.ok(fitInBox(img, { left: 0, top: 0, width: 100, height: 100 }).height === 50);
  });
  it("Tìm ảnh đính kèm: theo tên, thiếu đuôi, số thứ tự, mặc định ảnh mới nhất; không lẫn với PDF và DOCX", () => {
    clearAttachedFiles();
    assert.strictEqual(findAttachedImage(), null);
    registerAttachedFile("hop-dong.pdf", "UERG");
    registerAttachedFile("Mặt bằng.PNG", "QQ==");
    registerAttachedFile("phoi-canh.jpg", "Qg==");
    registerAttachedFile("ban-ve.docx", "RA==");
    assert.deepStrictEqual(listAttachedImages(), ["Mặt bằng.PNG", "phoi-canh.jpg"]);
    assert.strictEqual(findAttachedImage()?.name, "phoi-canh.jpg");
    assert.strictEqual(findAttachedImage("mặt bằng")?.name, "Mặt bằng.PNG");
    assert.strictEqual(findAttachedImage("phoi")?.name, "phoi-canh.jpg");
    assert.strictEqual(findAttachedImage("1")?.name, "Mặt bằng.PNG");
    assert.strictEqual(findAttachedImage(2)?.name, "phoi-canh.jpg");
    assert.strictEqual(findAttachedImage("9"), null);
    assert.strictEqual(findAttachedImage("hop-dong"), null);
    // Đường cũ không đổi: tệp PDF, DOCX vẫn tìm theo đuôi
    assert.strictEqual(findAttachedFile(undefined, ".pdf")?.name, "hop-dong.pdf");
    assert.strictEqual(findAttachedFile(undefined, ".docx")?.name, "ban-ve.docx");
    clearAttachedFiles();
  });
});
