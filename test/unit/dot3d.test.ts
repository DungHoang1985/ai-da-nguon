/** Đợt 3 — phần B: các ca từ lần soát mã độc lập (kênh lệnh ghi chéo, ảnh của công cụ, kiểm bài). */
import * as assert from "assert";
import { CMD_PREFIX, RES_PREFIX, beat, buildPayload, claimCommand, describePayload, sendCommand, takeCommand } from "../../src/core/crossapp";
import { createCommonTools } from "../../src/office/common";
import { checkDeck } from "../../src/core/deckcheck";
import { prepareHistory } from "../../src/core/agent/history";

function memStore() {
  const m: Record<string, string> = {};
  return { m, getItem: (k: string) => (k in m ? m[k] : null), setItem: (k: string, v: string) => void (m[k] = v), removeItem: (k: string) => void delete m[k] };
}

describe("Đợt 3 · B1 · Sửa theo soát mã: không gửi lệnh chưa qua bước hỏi", () => {
  it("Lúc kiểm rủi ro đích chưa mở (không hỏi), ngay sau đó đích mở: công cụ KHÔNG gửi lệnh", async () => {
    const s = memStore();
    const t = createCommonTools({ host: "Excel", skills: [], shareContext: true, store: s, sendTimeoutMs: 50 });
    const w = t.specs.find((x) => x.name === "write_other_app")!;
    const args = { app: "Word", text: "Tóm tắt" };
    assert.strictEqual(await w.risky!(args), null);
    beat(s, { host: "Word", accept: true });
    await assert.rejects(() => t.execute("write_other_app", args), /chưa qua bước hỏi người dùng/);
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null);
    // Hỏi một nội dung, gửi nội dung khác: không gửi
    assert.ok(await w.risky!(args));
    await assert.rejects(() => t.execute("write_other_app", { app: "Word", text: "Nội dung khác hẳn" }), /chưa qua bước hỏi người dùng/);
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null);
    // Hỏi đúng nội dung rồi gửi: lệnh rời khung (ở đây không có khung nhận nên hết hạn và được rút lại)
    assert.ok(await w.risky!(args));
    await assert.rejects(() => t.execute("write_other_app", args), /rút lại, chưa ghi gì/);
    // Một lần hỏi chỉ dùng cho một lần gửi
    await assert.rejects(() => t.execute("write_other_app", args), /chưa qua bước hỏi người dùng/);
  });

  it("Hộp hỏi hiện gần đủ nội dung và nói rõ phần bị cắt; có câu nhắc nội dung do AI soạn", async () => {
    const s = memStore();
    beat(s, { host: "Word", accept: true, title: "bao-cao.docx" });
    const w = createCommonTools({ host: "Excel", skills: [], shareContext: true, store: s }).specs.find((x) => x.name === "write_other_app")!;
    const long = "Dòng mở đầu vô hại.\n" + "x".repeat(900) + "\nDONG-AN-O-GIUA\n" + "y".repeat(2000);
    const q = (await w.risky!({ app: "Word", text: long }))!;
    assert.ok(q.includes("DONG-AN-O-GIUA"), "phần giữa phải hiện trong hộp hỏi");
    assert.match(q, /còn \d+ ký tự nữa chưa hiện ở đây/);
    assert.match(q, new RegExp(`${long.length} ký tự`));
    assert.match(q, /Nội dung do AI soạn/);
    const x = describePayload(buildPayload("Excel", { values: [["A", "B"], ["1", "2"], ["3", "4"], ["5", "6"], ["7", "8"], ["cuối", "9"]] }), true);
    assert.ok(/hàng 2: 1 \| 2/.test(x) && /hàng cuối: cuối \| 9/.test(x), x);
  });
});

describe("Đợt 3 · B1 · Sửa theo soát mã: hai tài liệu cùng ứng dụng, lệnh bị bỏ", () => {
  it("Lệnh ghi cho tệp A: khung của tệp B không lấy, khung của tệp A lấy", async () => {
    const s = memStore();
    beat(s, { host: "Word", accept: true, title: "A.docx" });
    const p = sendCommand(s, "Excel", "Word", buildPayload("Word", { text: "x" }), { timeoutMs: 300, pollMs: 5 });
    await new Promise((r) => setTimeout(r, 15));
    assert.strictEqual(takeCommand(s, "Word", { accept: true, seen: new Set(), title: "B.docx" }), null);
    assert.ok(s.getItem(CMD_PREFIX + "Word"), "lệnh phải còn trong hộp thư cho khung của tệp A");
    const cmd = takeCommand(s, "Word", { accept: true, seen: new Set(), title: "A.docx" })!;
    assert.strictEqual(cmd.toTitle, "A.docx");
    s.setItem(RES_PREFIX + cmd.id, JSON.stringify({ id: cmd.id, ok: true, at: Date.now(), message: "ok" }));
    assert.ok((await p).ok);
  });

  it("Hai khung cùng giành một lệnh: chỉ một khung thắng", async () => {
    const s = memStore();
    const [a, b] = await Promise.all([claimCommand(s, "abcd1234efgh", "t1", () => new Promise((r) => setTimeout(r, 20))), claimCommand(s, "abcd1234efgh", "t2", () => new Promise((r) => setTimeout(r, 20)))]);
    assert.strictEqual([a, b].filter(Boolean).length, 1);
  });

  it("Lệnh sai hình dạng bị bỏ thì khung gửi được báo ngay là chưa ghi (không phải chờ hết hạn)", () => {
    const s = memStore();
    const bad = { id: "abcd1234efgh", from: "Excel", to: "Word", at: Date.now(), payload: { op: "word_delete" } };
    s.setItem(CMD_PREFIX + "Word", JSON.stringify(bad));
    assert.strictEqual(takeCommand(s, "Word", { accept: true, seen: new Set() }), null);
    const res = JSON.parse(s.getItem(RES_PREFIX + bad.id)!);
    assert.ok(res.ok === false && /chưa ghi gì/.test(res.error));
  });
});

describe("Đợt 3 · B3 · Sửa theo soát mã: ảnh của công cụ và chữ mẫu", () => {
  it("Model chưa rõ có nhận ảnh: ảnh người dùng đính kèm ở đầu lượt vẫn được giữ sau khi công cụ chụp ảnh; ảnh công cụ chỉ giữ lần chụp gần nhất", () => {
    const img = (d: string) => [{ mime: "image/png", data: d }];
    const h: any[] = [
      { role: "user", content: "xem ảnh này và slide 2", images: img("NGUOI_DUNG") },
      { role: "assistant", content: "", toolCalls: [{ id: "c1", name: "ppt_get_slide_image", args: {} }] },
      { role: "tool", toolCallId: "c1", name: "ppt_get_slide_image", content: "{}" },
      { role: "user", content: "[ảnh công cụ 1]", images: img("CONG_CU_1"), fromTool: true },
      { role: "assistant", content: "", toolCalls: [{ id: "c2", name: "ppt_get_slide_image", args: {} }] },
      { role: "tool", toolCallId: "c2", name: "ppt_get_slide_image", content: "{}" },
      { role: "user", content: "[ảnh công cụ 2]", images: img("CONG_CU_2"), fromTool: true },
    ];
    const out = prepareHistory(h, { provider: { id: "openrouter", name: "OpenRouter", kind: "openai", baseUrl: "x" } as any, vision: undefined } as any);
    const kept = out.filter((m: any) => m.images?.length).map((m: any) => m.images[0].data);
    assert.deepStrictEqual(kept, ["NGUOI_DUNG", "CONG_CU_2"]);
  });

  it("Chữ mẫu trong ngoặc vuông được báo", () => {
    const sh = (text: string) => ({ name: "Hộp", type: "TextBox", box: [40, 120, 300, 80] as [number, number, number, number], paragraphs: [{ text, font: "Calibri", size: 18 }], autoSize: "AutoSizeShapeToFitText", wordWrap: true });
    const title = { name: "Title", type: "Placeholder", isTitle: true, box: [40, 30, 600, 60] as [number, number, number, number], paragraphs: [{ text: "T", font: "Calibri", size: 30 }] };
    for (const t of ["[cần bổ sung]", "Doanh thu: [cần bổ sung số liệu quý 4]", "Click to add text", "Lorem ipsum dolor"])
      assert.ok(checkDeck([{ slide: 1, shapes: [title, sh(t)] }], { slideW: 720, slideH: 405, themeFonts: ["Calibri"] }).issues.some((i) => i.kind === "placeholder_text"), t);
    assert.ok(!checkDeck([{ slide: 1, shapes: [title, sh("Tiến độ quý 4 [xem phụ lục 2]")] }], { slideW: 720, slideH: 405, themeFonts: ["Calibri"] }).issues.some((i) => i.kind === "placeholder_text"));
  });
});
