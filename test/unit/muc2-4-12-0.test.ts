/** Mức 2 tối ưu theo mã add-in chính thức, bản 4.12.0 (tasks/viec-le/rea-claude-chinh-thuc/):
 *  cổng cuối lượt "đổi hình thức mà chưa xem lại" (PowerPoint), danh sách lỗi khách quan của ảnh slide. */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { lookGateNote, runAgent } from "../../src/core/agent/loop";
import { AgentEvent, AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { PPT_LOOK, PPT_NEEDS_LOOK, pptTools } from "../../src/office/pptTools";
import { SLIDE_CHECKLIST } from "../../src/office/pptDesign";

const P = (id: string) => getProvider(id)!;
function sse(events: any[]): Response {
  return new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });
}
const toolCall = (name: string, args: string, id: string) => sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: args } }] } }] }]);
const say = (text: string) => sse([{ choices: [{ delta: { content: text } }] }]);

const TOOLS: ToolSpec[] = [
  { name: "sua_slide", description: "sửa", parameters: { type: "object", properties: { o: { type: "string" } } }, write: true, needsLook: true },
  { name: "them_ghi_chu", description: "ghi không đổi hình thức", parameters: { type: "object", properties: {} }, write: true },
  { name: "chup", description: "xem", parameters: { type: "object", properties: {} }, look: true },
];
function tools(run: (name: string) => any, specs = TOOLS): HostToolSet {
  return { host: "PowerPoint", specs, overview: async () => ({}), systemPrompt: () => "S", execute: async (name) => run(name) };
}
async function go(responses: Response[], exec: (name: string) => any = () => ({ ok: true }), over: any = {}) {
  const events: AgentEvent[] = [];
  let calls = 0;
  const msgs = await runAgent({
    provider: P("groq"),
    key: "K",
    model: "m",
    system: "S",
    messages: [{ role: "user", content: "x" }],
    log: [],
    tools: tools(exec),
    onEvent: (e) => events.push(e),
    fetchFn: (async () => {
      calls++;
      const r = responses.shift();
      if (!r) throw new Error("model bị gọi quá số câu dựng sẵn");
      return r;
    }) as any,
    ...over,
  });
  return { msgs, events, calls };
}
const lookNotes = (msgs: AgentMessage[]) => msgs.filter((m) => m.role === "user" && (m as any).fromTool && /chưa xem lại kết quả/.test(String(m.content)));

describe("4.12.0 · Cổng cuối lượt: đổi hình thức mà chưa xem lại", () => {
  it("Sửa slide rồi định kết thúc → add-in nhắc một lần, model xem lại rồi mới xong", async () => {
    const { msgs, events, calls } = await go([toolCall("sua_slide", "{}", "c1"), say("Xong."), toolCall("chup", "{}", "c2"), say("Đã xem lại, không có lỗi.")]);
    assert.strictEqual(calls, 4);
    const notes = lookNotes(msgs);
    assert.strictEqual(notes.length, 1);
    assert.match(String(notes[0].content), /\(sua_slide\).*Gọi chup/);
    assert.deepStrictEqual((events.find((e) => e.type === "look_gate") as any).tools, ["sua_slide"]);
  });
  it("Đã xem lại sau lần sửa cuối → không nhắc", async () => {
    const { msgs, calls } = await go([toolCall("sua_slide", "{}", "c1"), toolCall("chup", "{}", "c2"), say("Xong.")]);
    assert.strictEqual(calls, 3);
    assert.strictEqual(lookNotes(msgs).length, 0);
  });
  it("Xem rồi lại sửa tiếp mà không xem lần nữa → vẫn nhắc", async () => {
    const { msgs } = await go([toolCall("sua_slide", "{}", "c1"), toolCall("chup", "{}", "c2"), toolCall("sua_slide", '{"o":"2"}', "c3"), say("Xong."), say("Đã xem.")]);
    assert.strictEqual(lookNotes(msgs).length, 1);
  });
  it("Lệnh sửa báo lỗi, hoặc lệnh ghi không đổi hình thức → không nhắc", async () => {
    const a = await go([toolCall("sua_slide", "{}", "c1"), say("Chưa sửa được vì lỗi.")], () => ({ ok: false, error: "hỏng" }));
    assert.strictEqual(lookNotes(a.msgs).length, 0);
    const b = await go([toolCall("them_ghi_chu", "{}", "c1"), say("Xong.")]);
    assert.strictEqual(lookNotes(b.msgs).length, 0);
  });
  it("Tối đa một lần mỗi lượt; không nhắc ở bước cuối của trần", async () => {
    const a = await go([toolCall("sua_slide", "{}", "c1"), say("Xong."), say("Vẫn xong.")]);
    assert.strictEqual(a.calls, 3);
    assert.strictEqual(lookNotes(a.msgs).length, 1);
    const b = await go([toolCall("sua_slide", "{}", "c1"), say("Xong.")], undefined, { maxSteps: 2 });
    assert.strictEqual(b.calls, 2);
    assert.strictEqual(lookNotes(b.msgs).length, 0);
  });
  it("Bộ công cụ không có công cụ xem lại (Word, Excel) → không nhắc", async () => {
    const events: AgentEvent[] = [];
    const msgs = await runAgent({
      provider: P("groq"),
      key: "K",
      model: "m",
      system: "S",
      messages: [{ role: "user", content: "x" }],
      log: [],
      tools: tools(() => ({ ok: true }), [TOOLS[0]]),
      onEvent: (e) => events.push(e),
      fetchFn: (async () => (msgsCount++ === 0 ? toolCall("sua_slide", "{}", "c1") : say("Xong."))) as any,
    });
    assert.strictEqual(lookNotes(msgs).length, 0);
  });
  it("Câu nhắc nêu công cụ đã đổi và công cụ để xem", () => {
    assert.match(lookGateNote(["ppt_format_text"], ["ppt_get_slide_image", "ppt_check_deck"]), /\(ppt_format_text\).*ppt_get_slide_image hoặc ppt_check_deck/);
  });
});
let msgsCount = 0;

describe("4.12.0 · PowerPoint: cờ cổng xem lại, danh sách lỗi khách quan", () => {
  const spec = (n: string) => pptTools.specs.find((s) => s.name === n);
  it("Lệnh đổi bố cục có needsLook, ảnh slide và kiểm bài có look; mọi tên trong hai danh sách đều là công cụ thật", () => {
    for (const n of PPT_NEEDS_LOOK) assert.ok(spec(n)?.needsLook && spec(n)?.write, n);
    for (const n of PPT_LOOK) assert.ok(spec(n)?.look && !spec(n)?.write, n);
    assert.ok(!spec("ppt_delete_slide")?.needsLook, "Xóa slide không cần xem bố cục");
    assert.ok(spec("ppt_delete_slide")?.risky, "Giữ nguyên kiểm rủi ro khi gắn cờ");
  });
  it("Danh sách lỗi chỉ lỗi khách quan, có quy tắc tối đa 3 vòng", () => {
    assert.ok(SLIDE_CHECKLIST.length >= 15);
    assert.ok(SLIDE_CHECKLIST.some((x) => /tràn/.test(x)) && SLIDE_CHECKLIST.some((x) => /14pt/.test(x)) && SLIDE_CHECKLIST.some((x) => /placeholder/.test(x)));
    assert.ok(SLIDE_CHECKLIST.some((x) => /3 vòng/.test(x)));
  });
});
