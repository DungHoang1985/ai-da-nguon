/** Sau Đợt 3 (bản 4.3.1): một nhiệm vụ, ít lượt gọi model — quy tắc gộp lệnh, chụp nhiều slide một lần, nhiều lệnh trong một lượt. */
import * as assert from "assert";
import { EFFICIENCY_RULES, SAFETY_RULES, buildSystemPrompt } from "../../src/core/agent/prompt";
import { PPT_DESIGN_SPECS } from "../../src/office/pptDesign";
import { pptTools } from "../../src/office/pptTools";
import { runAgent } from "../../src/core/agent/loop";

describe("Bản 4.3.1 · Một nhiệm vụ, ít lượt gọi model", () => {
  it("Lời nhắc hệ thống của mọi ứng dụng có quy tắc gộp lệnh và không đọc lại; vẫn cố định (cùng đầu vào → cùng chuỗi)", () => {
    for (const host of ["Word", "Excel", "PowerPoint"]) {
      const a = buildSystemPrompt({ base: "BASE " + host, host });
      const b = buildSystemPrompt({ base: "BASE " + host, host });
      assert.strictEqual(a, b);
      assert.ok(a.includes(EFFICIENCY_RULES) && a.includes(SAFETY_RULES));
      assert.ok(a.indexOf(EFFICIENCY_RULES) < a.indexOf(SAFETY_RULES));
    }
    assert.match(EFFICIENCY_RULES, /CÙNG MỘT lượt/);
    assert.match(EFFICIENCY_RULES, /không gọi lại công cụ tổng quan/);
    assert.ok(!/\d{4}/.test(EFFICIENCY_RULES), "không chứa ngày giờ hay thứ hay đổi");
  });

  it("Công cụ chụp ảnh slide nhận nhiều slide trong một lần gọi; lời nhắc PowerPoint dặn làm xong hết rồi kiểm một lần", () => {
    const sp = PPT_DESIGN_SPECS.find((s) => s.name === "ppt_get_slide_image")!;
    assert.strictEqual(sp.parameters.properties.slides.type, "array");
    assert.match(sp.description, /up to 4 in ONE call/);
    const p = pptTools.systemPrompt();
    assert.ok(/CÙNG một lượt/.test(p) && /slides: \[\.\.\]/.test(p) && /Không gọi ppt_get_overview sau mỗi bước/.test(p));
  });

  it("Model gọi ba lệnh trong một lượt: cả ba chạy theo thứ tự trong MỘT bước, nhiệm vụ xong sau hai lần gọi model", async () => {
    const sse = (events: any[]) => new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });
    let calls = 0;
    const fetchFn: any = async () => {
      calls++;
      return calls === 1
        ? sse([
            {
              choices: [
                {
                  delta: {
                    tool_calls: [
                      { index: 0, id: "c1", function: { name: "ppt_add_slide", arguments: JSON.stringify({ title: "Một" }) } },
                      { index: 1, id: "c2", function: { name: "ppt_add_slide", arguments: JSON.stringify({ title: "Hai" }) } },
                      { index: 2, id: "c3", function: { name: "ppt_add_slide", arguments: JSON.stringify({ title: "Ba" }) } },
                    ],
                  },
                },
              ],
            },
          ])
        : sse([{ choices: [{ delta: { content: "Đã thêm ba slide." } }] }]);
    };
    const ran: string[] = [];
    const steps: number[] = [];
    const tools: any = {
      host: "PowerPoint",
      specs: [{ name: "ppt_add_slide", description: "x", parameters: { type: "object", properties: {} }, write: true }],
      execute: async (_n: string, a: any) => {
        ran.push(a.title);
        return { ok: true, slide: ran.length };
      },
      overview: async () => ({}),
      systemPrompt: () => "x",
    };
    const msgs = await runAgent({
      provider: { id: "openrouter", name: "OpenRouter", kind: "openai", baseUrl: "https://x.invalid/v1" } as any,
      key: "k",
      model: "m",
      tools,
      system: "x",
      messages: [{ role: "user", content: "thêm ba slide" }],
      log: [],
      onEvent: (e) => {
        if (e.type === "step") steps.push(e.step);
      },
      fetchFn,
      retryDelaysMs: [],
    });
    assert.deepStrictEqual(ran, ["Một", "Hai", "Ba"]);
    assert.strictEqual(calls, 2);
    assert.deepStrictEqual(steps, [1, 2]);
    assert.strictEqual(msgs.filter((m) => m.role === "tool").length, 3);
  });
});
