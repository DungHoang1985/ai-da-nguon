/** Kiểm thử của Đợt 0A (bản 3.6). */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { buildStepRequest, callModelStep } from "../../src/core/agent/llm";
import { displayVersion, APP_VERSION } from "../../src/core/version";
import { replaceLimit } from "../../src/office/util";

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
    assert.strictEqual(body.system[0].text, "S");
    assert.strictEqual(body.messages.length, 1);
    assert.strictEqual((r.init.headers as any)["anthropic-dangerous-direct-browser-access"], "true");
  });
  it("Hết hạn mức 429 → báo lỗi tiếng Việt (thay kiểm thử của chat.ts)", async () => {
    const f = (async () => new Response("slow down", { status: 429 })) as any;
    await assert.rejects(
      callModelStep({ provider: P("groq"), key: "K", model: "m", messages: [{ role: "user", content: "x" }], tools: [], onText: () => {}, fetchFn: f, retryDelaysMs: [] }),
      /Hết hạn mức/
    );
  });
});

import { inputKeyAction } from "../../src/core/keys";

describe("Đợt 0A – bước 5: Enter để gửi", () => {
  const idle = { menuOpen: false, busy: false };
  it("Enter gửi; Shift+Enter xuống dòng; Ctrl+Enter vẫn gửi", () => {
    assert.strictEqual(inputKeyAction({ key: "Enter" }, idle), "send");
    assert.strictEqual(inputKeyAction({ key: "Enter", shiftKey: true }, idle), "newline");
    assert.strictEqual(inputKeyAction({ key: "Enter", ctrlKey: true }, idle), "send");
    assert.strictEqual(inputKeyAction({ key: "a" }, idle), "none");
  });
  it("Bộ gõ tiếng Việt đang ghép chữ (isComposing / mã 229): không gửi", () => {
    assert.strictEqual(inputKeyAction({ key: "Enter", isComposing: true }, idle), "none");
    assert.strictEqual(inputKeyAction({ key: "Enter", keyCode: 229 }, idle), "none");
  });
  it("Danh sách '/' đang mở: mũi tên di chuyển, Enter/Tab chọn, Esc đóng", () => {
    const menu = { menuOpen: true, busy: false };
    assert.strictEqual(inputKeyAction({ key: "ArrowDown" }, menu), "menu-next");
    assert.strictEqual(inputKeyAction({ key: "ArrowUp" }, menu), "menu-prev");
    assert.strictEqual(inputKeyAction({ key: "Enter" }, menu), "menu-pick");
    assert.strictEqual(inputKeyAction({ key: "Tab" }, menu), "menu-pick");
    assert.strictEqual(inputKeyAction({ key: "Escape" }, menu), "menu-close");
  });
  it("Đang chạy: Enter không gửi, Esc = Dừng", () => {
    const busy = { menuOpen: false, busy: true };
    assert.strictEqual(inputKeyAction({ key: "Enter" }, busy), "blocked");
    assert.strictEqual(inputKeyAction({ key: "Escape" }, busy), "stop");
    assert.strictEqual(inputKeyAction({ key: "Escape" }, idle), "none");
  });
});

import { overviewForPrompt, packParagraphs, paragraphIndex, takeWithinBudget } from "../../src/office/util";
import { MAX_TOOL_RESULT_CHARS, clip } from "../../src/core/agent/loop";

describe("Đợt 0A – bước 8: ba lỗi Word", () => {
  it("Văn bản 400 đoạn × 370 ký tự: đọc tiếp theo next_start thì đọc hết, tổng số đoạn đã đọc = tổng số đoạn", () => {
    const doc = Array.from({ length: 400 }, (_, i) => ({ style: "Normal", text: `Đoạn ${i}: ` + "nội dung hợp đồng ".repeat(20).slice(0, 360) }));
    const seen: number[] = [];
    let start = 0;
    for (let guard = 0; guard < 100; guard++) {
      const r: any = packParagraphs(doc, start, start + 200);
      const json = JSON.stringify(r);
      assert.ok(json.length < MAX_TOOL_RESULT_CHARS, `JSON ${json.length} ký tự, vượt mức cắt`);
      // next_start nằm đầu kết quả: dù vòng lặp có cắt bớt thì vẫn còn
      if (r.next_start !== undefined) assert.ok(clip(json, 2000).includes('"next_start"'));
      seen.push(...r.paragraphs.map((p: any) => p.i));
      if (r.next_start === undefined) break;
      start = r.next_start;
    }
    assert.strictEqual(seen.length, 400);
    assert.deepStrictEqual(seen, Array.from({ length: 400 }, (_, i) => i));
  });
  it("Một đoạn dài hơn ngân sách: vẫn trả về (rút gọn, có ghi chú), không làm vỡ JSON", () => {
    const doc = [{ style: "Normal", text: "x".repeat(50000) }, { style: "Normal", text: "b" }];
    const r: any = packParagraphs(doc, 0, 2);
    assert.ok(JSON.stringify(r).length < MAX_TOOL_RESULT_CHARS);
    assert.match(r.paragraphs[0].text, /đã cắt/);
    assert.strictEqual(r.next_start, 1);
    const t = takeWithinBudget([1, 2, 3], 1000);
    assert.deepStrictEqual(t, { picked: [1, 2, 3], truncated: false });
  });
  it("Số đoạn của nhận xét / sửa đổi tìm theo định danh đoạn, đúng cả khi có hai đoạn trùng nội dung", () => {
    const texts = ["Bên A:", "Nội dung", "Bên A:"];
    const ids = ["id-0", "id-1", "id-2"];
    assert.strictEqual(paragraphIndex("id-2", "Bên A:", ids, texts), 2);
    assert.strictEqual(paragraphIndex(undefined, "Bên A:", null, texts), 0, "Word cũ (không có định danh) thì so chữ như trước");
    assert.strictEqual(paragraphIndex("khong-co", "Nội dung", ids, texts), 1);
  });
  it("Trạng thái văn bản trong lời nhắc: JSON hợp lệ, không quá 6.000 ký tự, luôn còn vùng chọn và chế độ theo dõi", () => {
    const ov = {
      paragraphs: 900,
      characters: 300000,
      headings: Array.from({ length: 100 }, (_, i) => ({ index: i * 9, style: "Heading 2", text: `Điều ${i}. ` + "Quy định về nghĩa vụ ".repeat(5) })),
      first_paragraphs: Array.from({ length: 6 }, (_, i) => ({ index: i, style: "Normal", text: "Mở đầu ".repeat(30) })),
      selection: "Đoạn đang chọn để sửa",
      track_changes: "TrackAll",
      comments: 3,
    };
    const s = overviewForPrompt(ov);
    assert.ok(s.length <= 6000, `dài ${s.length}`);
    const j = JSON.parse(s);
    assert.strictEqual(j.selection, "Đoạn đang chọn để sửa");
    assert.strictEqual(j.track_changes, "TrackAll");
    assert.strictEqual(j.comments, 3);
    assert.strictEqual(j.headings_truncated, true);
    assert.strictEqual(j.headings_total, 100);
    assert.ok(Object.keys(j).indexOf("selection") < Object.keys(j).indexOf("headings"), "vùng chọn đứng trước danh sách tiêu đề");
  });
});

describe("Đợt 0A – bước 9: số chỗ thay tối đa của word_replace_text", () => {
  it("Model truyền -1, 0, chuỗi lạ hay bỏ trống thì dùng mặc định 50, không ra số âm", () => {
    for (const v of [-1, 0, undefined, null, "", "all", NaN, -Infinity]) assert.strictEqual(replaceLimit(v), 50, String(v));
  });
  it("Số nguyên dương thì giữ; số lẻ làm tròn xuống; chuỗi số vẫn nhận", () => {
    assert.strictEqual(replaceLimit(1), 1);
    assert.strictEqual(replaceLimit(3.9), 3);
    assert.strictEqual(replaceLimit("7"), 7);
    assert.strictEqual(replaceLimit(0.5), 50);
  });
});
