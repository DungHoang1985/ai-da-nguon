/** Kiểm thử của Đợt 0A (bản 3.6). */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { buildStepRequest, callModelStep } from "../../src/core/agent/llm";
import { displayVersion, APP_VERSION } from "../../src/core/version";

const P = (id: string) => getProvider(id)!;

describe("Đợt 0A – bước 3: số phiên bản một nguồn", () => {
  it("3.6.0 hiện là 3.6; bản vá giữ đủ; chuỗi lạ giữ nguyên", () => {
    assert.strictEqual(displayVersion("3.6.0"), "3.6");
    assert.strictEqual(displayVersion("3.6.1"), "3.6.1");
    assert.strictEqual(displayVersion("3.6.0-beta"), "3.6.0-beta");
    assert.strictEqual(displayVersion("dev"), "dev");
  });
  it("Ngoài webpack (kiểm thử) không có số phiên bản thật thì không sập", () => {
    assert.ok(typeof APP_VERSION === "string" && APP_VERSION.length > 0);
  });
  it("Anthropic: system ở trường riêng, có header cho phép gọi từ trình duyệt (thay kiểm thử của chat.ts)", () => {
    const r = buildStepRequest(P("anthropic"), "K", "m", [
      { role: "system", content: "S" },
      { role: "user", content: "Xin chào" },
    ], []);
    const body = JSON.parse(r.init.body as string);
    assert.strictEqual(body.system, "S");
    assert.strictEqual(body.messages.length, 1);
    assert.strictEqual((r.init.headers as any)["anthropic-dangerous-direct-browser-access"], "true");
  });
  it("Hết hạn mức 429 → báo lỗi tiếng Việt (thay kiểm thử của chat.ts)", async () => {
    const f = (async () => new Response("slow down", { status: 429 })) as any;
    await assert.rejects(
      callModelStep({ provider: P("groq"), key: "K", model: "m", messages: [{ role: "user", content: "x" }], tools: [], onText: () => {}, fetchFn: f }),
      /Hết hạn mức/
    );
  });
});
