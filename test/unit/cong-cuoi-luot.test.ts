/** Mức 1 tối ưu theo mã add-in chính thức (25/11/2026, tasks/viec-le/rea-claude-chinh-thuc/):
 *  cổng cuối lượt cho lệnh lỗi chưa xử lý, quy tắc đọc lại sau khi ghi, quy tắc báo cáo đúng việc đã làm. */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { errorGateNote, runAgent } from "../../src/core/agent/loop";
import { EFFICIENCY_RULES, REPORT_RULES, buildSystemPrompt } from "../../src/core/agent/prompt";
import { CORE_GUIDES } from "../../src/core/skillpack";
import { AgentEvent, AgentMessage, HostToolSet, ToolSpec } from "../../src/core/agent/types";

const P = (id: string) => getProvider(id)!;
function sse(events: any[]): Response {
  return new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });
}
const toolCall = (name: string, args: string, id: string) => sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: args } }] } }] }]);
const say = (text: string) => sse([{ choices: [{ delta: { content: text } }] }]);

const TOOLS: ToolSpec[] = [
  { name: "doc", description: "đọc", parameters: { type: "object", properties: {} } },
  { name: "ghi", description: "ghi", parameters: { type: "object", properties: { o: { type: "string" } } }, write: true, risky: async () => "rủi ro" },
];
function tools(run: (name: string, args: any) => any): HostToolSet {
  return { host: "Word", specs: TOOLS, overview: async () => ({}), systemPrompt: () => "S", execute: async (name, args) => run(name, args) };
}
/** Chạy với danh sách câu trả lời dựng sẵn; trả lịch sử, sự kiện và số lần gọi model */
async function run(responses: Response[], exec: (name: string, args: any) => any, over: any = {}) {
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
const gateNotes = (msgs: AgentMessage[]) => msgs.filter((m) => m.role === "user" && (m as any).fromTool && /Add-in nhắc: trong lượt này/.test(String(m.content)));

describe("Cổng cuối lượt: lệnh lỗi chưa xử lý", () => {
  it("Lệnh ghi lỗi rồi model định kết thúc → add-in nhắc một lần, model trả lời tiếp", async () => {
    const { msgs, events, calls } = await run(
      [toolCall("ghi", '{"o":"A1"}', "c1"), say("Đã sửa xong."), say("Chưa sửa được A1 vì không tìm thấy chỗ cần sửa.")],
      () => ({ ok: false, error: "Không tìm thấy" })
    );
    assert.strictEqual(calls, 3);
    const notes = gateNotes(msgs);
    assert.strictEqual(notes.length, 1);
    assert.match(String(notes[0].content), /1 lệnh công cụ báo lỗi.*\(ghi\)/);
    assert.match(String(notes[0].content), /không phải yêu cầu mới của người dùng/);
    assert.deepStrictEqual((events.find((e) => e.type === "error_gate") as any).tools, ["ghi"]);
    // Lịch sử chỉ thêm: câu "Đã sửa xong." vẫn còn nguyên, nhắc nằm sau nó, câu cuối là câu trả lời mới
    const i = msgs.findIndex((m) => m.role === "assistant" && m.content === "Đã sửa xong.");
    assert.ok(i > 0 && msgs[i + 1] === notes[0]);
    assert.strictEqual(msgs[msgs.length - 1].content, "Chưa sửa được A1 vì không tìm thấy chỗ cần sửa.");
  });
  it("Chỉ nhắc tối đa một lần mỗi lượt", async () => {
    const { msgs, calls } = await run([toolCall("doc", "{}", "c1"), say("Xong."), say("Vẫn xong.")], () => {
      throw new Error("Đọc hỏng");
    });
    assert.strictEqual(calls, 3);
    assert.strictEqual(gateNotes(msgs).length, 1);
  });
  it("Lỗi rồi gọi lại cùng công cụ thành công → không nhắc", async () => {
    let n = 0;
    const { msgs, calls } = await run([toolCall("doc", "{}", "c1"), toolCall("doc", "{}", "c2"), say("Đã đọc.")], () => {
      if (++n === 1) throw new Error("Bận");
      return { ok: true };
    });
    assert.strictEqual(calls, 3);
    assert.strictEqual(gateNotes(msgs).length, 0);
  });
  it("Người dùng từ chối lệnh → không tính là lỗi, không nhắc", async () => {
    const { msgs, calls } = await run([toolCall("ghi", '{"o":"A1"}', "c1"), say("Anh không cho phép nên em dừng.")], () => ({ ok: true }), {
      confirmRisky: true,
      approve: async () => false,
    });
    assert.strictEqual(calls, 2);
    assert.strictEqual(gateNotes(msgs).length, 0);
  });
  it("Đang ở bước cuối của trần → không nhắc (không vượt trần số bước)", async () => {
    const { msgs, calls } = await run([toolCall("doc", "{}", "c1"), say("Xong.")], () => ({ ok: false, error: "x" }), { maxSteps: 2 });
    assert.strictEqual(calls, 2);
    assert.strictEqual(gateNotes(msgs).length, 0);
  });
  it("Câu nhắc gộp theo tên công cụ, có số lần", () => {
    const s = errorGateNote(
      new Map([
        ["word_replace_text", 2],
        ["doc", 1],
      ])
    );
    assert.match(s, /3 lệnh công cụ báo lỗi/);
    assert.match(s, /\(doc, word_replace_text \(×2\)\)/);
  });
});

describe("Lời nhắc: đọc lại sau khi ghi, báo cáo đúng việc đã làm", () => {
  const sys = buildSystemPrompt({ base: "B", host: "Word", guide: CORE_GUIDES.Word });
  it("Bỏ quy tắc 'KHÔNG ĐỌC LẠI', thay bằng đọc lại một lần sau khi ghi", () => {
    assert.ok(!/KHÔNG ĐỌC LẠI/.test(EFFICIENCY_RULES));
    assert.match(EFFICIENCY_RULES, /ĐỌC LẠI SAU KHI GHI/);
    assert.match(EFFICIENCY_RULES, /MỘT lần cho mọi chỗ vừa ghi/);
  });
  it("Lời nhắc hệ thống có quy tắc báo cáo, lỗi giữa chừng và lời người dùng thắng kết quả công cụ", () => {
    assert.ok(sys.includes(REPORT_RULES));
    assert.match(REPORT_RULES, /“tất cả”/);
    assert.match(REPORT_RULES, /không chạy lại nguyên lệnh/);
    assert.match(REPORT_RULES, /tin lời họ hơn kết quả công cụ/);
  });
  it("Chuẩn Word: sửa hẹp tới cụm chữ, không thay hàng loạt chữ số trơn", () => {
    assert.match(CORE_GUIDES.Word, /Sửa hẹp tới cụm chữ đổi/);
    assert.match(CORE_GUIDES.Word, /không thay hàng loạt một chữ số trơn/);
  });
  it("Lời nhắc hệ thống không chứa thứ hay đổi (giữ bộ nhớ đệm)", () => {
    assert.strictEqual(buildSystemPrompt({ base: "B", host: "Word", guide: CORE_GUIDES.Word }), sys);
  });
});
