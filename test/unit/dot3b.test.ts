/** Đợt 3 — phần B: ghi chéo ứng dụng (kênh lệnh cục bộ qua kho dùng chung). */
import * as assert from "assert";
import {
  CMD_PREFIX,
  CMD_MAX_AGE_MS,
  LIVE_MAX_AGE_MS,
  RES_PREFIX,
  beat,
  blockReason,
  buildPayload,
  clearBeat,
  connectionLine,
  connections,
  describePayload,
  liveOf,
  reply,
  sendCommand,
  takeCommand,
  validCommand,
} from "../../src/core/crossapp";
import { createCommonTools } from "../../src/office/common";
import { cellAsText } from "../../src/office/crossExec";
import { runAgent } from "../../src/core/agent/loop";
import { describeCall } from "../../src/office/tools";
import { writeSucceeded } from "../../src/core/claims";

function memStore() {
  const m: Record<string, string> = {};
  return {
    m,
    getItem: (k: string) => (k in m ? m[k] : null),
    setItem: (k: string, v: string) => void (m[k] = v),
    removeItem: (k: string) => void delete m[k],
    get length() {
      return Object.keys(m).length;
    },
    key: (i: number) => Object.keys(m)[i] ?? null,
  };
}

describe("Đợt 3 · B1 · Kênh lệnh ghi chéo ứng dụng", () => {
  it("Nhịp sống: khung đang mở, quá hạn thì coi là đã đóng; chỉ báo kết nối", () => {
    const s = memStore();
    beat(s, { host: "Excel", accept: true, title: "tien-do.xlsx" }, 1000);
    assert.deepStrictEqual(liveOf(s, "Excel", 2000), { host: "Excel", at: 1000, title: "tien-do.xlsx", accept: true });
    assert.strictEqual(liveOf(s, "Excel", 1000 + LIVE_MAX_AGE_MS + 1), null);
    assert.strictEqual(liveOf(s, "Word", 2000), null);
    const cs = connections(s, "PowerPoint", 2000);
    assert.deepStrictEqual(cs.map((c) => [c.app, c.open]), [["Excel", true], ["Word", false]]);
    assert.strictEqual(connectionLine(cs), "Excel ✓ · Word —");
    clearBeat(s, "Excel");
    assert.strictEqual(liveOf(s, "Excel", 2000), null);
    // Khóa nhịp sống bị ghi bậy không làm sập
    s.setItem("aidn.live.Word", "{hỏng");
    assert.strictEqual(liveOf(s, "Word", 2000), null);
  });

  it("Khung đích chưa mở hoặc tắt nhận ghi: báo rõ ngay, không đặt lệnh", async () => {
    const s = memStore();
    const r = await sendCommand(s, "Excel", "Word", buildPayload("Word", { text: "Tóm tắt" }), { timeoutMs: 50 });
    assert.ok(!r.ok && /chưa mở/.test(r.error!) && /bấm nút AI Đa Nguồn/.test(r.error!));
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null);
    beat(s, { host: "Word", accept: false });
    assert.match(blockReason(s, "Excel", "Word")!, /Nhận ghi từ ứng dụng khác/);
    assert.match(blockReason(s, "Word", "Word")!, /chính ứng dụng này/);
  });

  it("Gửi → nhận → trả lời: đúng nội dung, lệnh rời hộp thư, kết quả được dọn", async () => {
    const s = memStore();
    beat(s, { host: "Word", accept: true, title: "bao-cao.docx" });
    const p = sendCommand(s, "Excel", "Word", buildPayload("Word", { text: "  Dòng 1\nDòng 2  " }), { timeoutMs: 2000, pollMs: 5 });
    await new Promise((r) => setTimeout(r, 20));
    const seen = new Set<string>();
    const cmd = takeCommand(s, "Word", { accept: true, seen })!;
    assert.deepStrictEqual([cmd.from, cmd.to, cmd.payload], ["Excel", "Word", { op: "word_append", text: "Dòng 1\nDòng 2" }]);
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null);
    reply(s, cmd.id, { ok: true, message: "Đã chèn 2 đoạn" });
    const res = await p;
    assert.deepStrictEqual([res.ok, res.message], [true, "Đã chèn 2 đoạn"]);
    assert.strictEqual(s.getItem(RES_PREFIX + cmd.id), null);
    // Lệnh đã chạy bị đặt lại vào hộp thư: không chạy lần hai
    s.setItem(CMD_PREFIX + "Word", JSON.stringify(cmd));
    assert.strictEqual(takeCommand(s, "Word", { accept: true, seen }), null);
  });

  it("Khung đích không lấy lệnh: hết hạn thì rút lệnh và báo chưa ghi; đã lấy mà không trả lời: báo có thể đã ghi", async () => {
    const s = memStore();
    beat(s, { host: "PowerPoint", accept: true });
    const r = await sendCommand(s, "Excel", "PowerPoint", buildPayload("PowerPoint", { title: "T" }), { timeoutMs: 40, pollMs: 5 });
    assert.ok(!r.ok && /rút lại, chưa ghi gì/.test(r.error!));
    assert.strictEqual(s.getItem(CMD_PREFIX + "PowerPoint"), null);
    const p = sendCommand(s, "Excel", "PowerPoint", buildPayload("PowerPoint", { title: "T" }), { timeoutMs: 60, pollMs: 5 });
    await new Promise((r2) => setTimeout(r2, 15));
    assert.ok(takeCommand(s, "PowerPoint", { accept: true, seen: new Set() }));
    const r2 = await p;
    assert.ok(!r2.ok && /Có thể nội dung ĐÃ được ghi/.test(r2.error!) && /không gửi lại ngay/.test(r2.error!));
  });

  it("Bấm Dừng trong lúc chờ: thôi chờ, rút lệnh", async () => {
    const s = memStore();
    beat(s, { host: "Word", accept: true });
    const signal = { aborted: false };
    const p = sendCommand(s, "Excel", "Word", buildPayload("Word", { text: "x" }), { timeoutMs: 5000, pollMs: 5, signal });
    await new Promise((r) => setTimeout(r, 20));
    signal.aborted = true;
    const r = await p;
    assert.ok(!r.ok);
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null);
  });

  it("Khung nhận kiểm lại lệnh: sai ứng dụng, sai thao tác, quá hạn, quá cỡ, thừa trường, tự gửi cho mình đều bị bỏ", () => {
    const now = 100000;
    const ok = { id: "abcd1234efgh", from: "Excel", to: "Word", at: now, payload: { op: "word_append", text: "Xin chào" } };
    assert.ok(validCommand(ok, "Word", now));
    const bad: any[] = [
      { ...ok, to: "Excel" },
      { ...ok, from: "Word" },
      { ...ok, from: "Outlook" },
      { ...ok, id: "../x" },
      { ...ok, at: now - CMD_MAX_AGE_MS - 1 },
      { ...ok, at: now + CMD_MAX_AGE_MS + 1 },
      { ...ok, payload: { op: "excel_new_sheet", name: "a", values: [["1"]] } },
      { ...ok, payload: { op: "word_append", text: "" } },
      { ...ok, payload: { op: "word_append", text: "x".repeat(30000) } },
      { ...ok, payload: { op: "word_append", text: "a", delete_all: true } },
      { ...ok, payload: { op: "word_delete", start: 0, end: 99 } },
      null,
      "chuỗi",
    ];
    for (const c of bad) assert.ok(!validCommand(c, "Word", now), JSON.stringify(c));
    // takeCommand gỡ lệnh xấu khỏi hộp thư
    const s = memStore();
    s.setItem(CMD_PREFIX + "Word", JSON.stringify(bad[9]));
    assert.strictEqual(takeCommand(s, "Word", { accept: true, seen: new Set(), now }), null);
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null);
  });

  it("Khung nhận đang tắt nhận ghi (hoặc đang bận): trả lời từ chối, không chạy", () => {
    const s = memStore();
    const cmd = { id: "abcd1234efgh", from: "Excel", to: "Word", at: Date.now(), payload: { op: "word_append", text: "x" } };
    s.setItem(CMD_PREFIX + "Word", JSON.stringify(cmd));
    assert.strictEqual(takeCommand(s, "Word", { accept: false, seen: new Set() }), null);
    const res = JSON.parse(s.getItem(RES_PREFIX + cmd.id)!);
    assert.ok(res.ok === false && /Nhận ghi từ ứng dụng khác/.test(res.error));
  });

  it("Nội dung ghi hộ: chỉ ba thao tác thêm mới; bảng Excel chuẩn hóa thành chữ nhật; công thức thành chữ", () => {
    assert.deepStrictEqual(buildPayload("Excel", { sheet_name: "Tóm/tắt:1", values: [["A", "B"], ["1"]] }), { op: "excel_new_sheet", name: "Tóm tắt 1", values: [["A", "B"], ["1", ""]] });
    assert.throws(() => buildPayload("Excel", { values: "A,B" }), /values/);
    assert.throws(() => buildPayload("Excel", { values: Array.from({ length: 200 }, () => Array(40).fill("x")) }), /quá lớn/);
    assert.throws(() => buildPayload("Word", {}), /text/);
    assert.deepStrictEqual(buildPayload("PowerPoint", { title: "Tiến độ", bullets: ["a", "b"] }), { op: "ppt_add_slide", title: "Tiến độ", body: "a\nb" });
    assert.match(describePayload(buildPayload("Word", { text: "Một\nHai" })), /thêm 2 đoạn \(7 ký tự\) vào cuối văn bản/);
    assert.match(describePayload(buildPayload("Excel", { values: [["A", "B"]] })), /sheet mới .* 1×2 ô/);
    assert.deepStrictEqual(["=SUM(A1)", "+1", "@x", "-5", "-abc", "12", "Chữ"].map(cellAsText), ["'=SUM(A1)", "'+1", "'@x", "-5", "'-abc", "12", "Chữ"]);
  });
});

describe("Đợt 3 · B1 · Công cụ write_other_app: luôn hỏi trước khi ghi sang ứng dụng khác", () => {
  const skills: any[] = [];
  it("Có trong công cụ dùng chung khi bật kết nối; là công cụ ghi, luôn hỏi; tắt kết nối với một ứng dụng thì không còn trong danh sách", () => {
    const t = createCommonTools({ host: "Excel", skills, shareContext: true, store: memStore() });
    const w = t.specs.find((x) => x.name === "write_other_app")!;
    assert.ok(w.write && w.alwaysAsk && typeof w.risky === "function");
    assert.deepStrictEqual(w.parameters.properties.app.enum, ["Word", "PowerPoint"]);
    const t2 = createCommonTools({ host: "Excel", skills, shareContext: true, apps: ["Word"], store: memStore() });
    assert.deepStrictEqual(t2.specs.find((x) => x.name === "write_other_app")!.parameters.properties.app.enum, ["Word"]);
    assert.ok(!createCommonTools({ host: "Excel", skills, shareContext: false }).specs.some((x) => /other_app/.test(x.name)));
    assert.notStrictEqual(describeCall({ id: "x", name: "write_other_app", args: { app: "Word" } }), "write_other_app");
  });

  it("Lời hỏi nêu ứng dụng đích, tên tệp và nội dung; đích chưa mở hoặc lệnh sai thì không hỏi (công cụ tự báo lỗi)", async () => {
    const s = memStore();
    const t = createCommonTools({ host: "Excel", skills, shareContext: true, store: s, sendTimeoutMs: 50 });
    const w = t.specs.find((x) => x.name === "write_other_app")!;
    assert.strictEqual(await w.risky!({ app: "Word", text: "Tóm tắt" }), null);
    await assert.rejects(() => t.execute("write_other_app", { app: "Word", text: "Tóm tắt" }), /chưa mở/);
    beat(s, { host: "Word", accept: true, title: "bao-cao.docx" });
    const q = (await w.risky!({ app: "Word", text: "Tóm tắt quý 3" }))!;
    assert.ok(/Ghi sang Word \(tệp “bao-cao\.docx”\)/.test(q) && /Tóm tắt quý 3/.test(q) && /Hoàn tác/.test(q), q);
    assert.strictEqual(await w.risky!({ app: "Word" }), null);
    await assert.rejects(() => t.execute("write_other_app", { app: "Word" }), /Thiếu text/);
    await assert.rejects(() => t.execute("write_other_app", { app: "Excel", values: [["a"]] }), /chính ứng dụng này/);
    const t2 = createCommonTools({ host: "Excel", skills, shareContext: true, apps: ["PowerPoint"], store: s });
    await assert.rejects(() => t2.execute("write_other_app", { app: "Word", text: "x" }), /đang tắt/);
  });

  it("Trong vòng lặp tác tử: hỏi cả khi người dùng tắt 'hỏi trước thao tác rủi ro'; không cho phép thì lệnh không rời khung", async () => {
    const s = memStore();
    beat(s, { host: "Word", accept: true });
    const extra = createCommonTools({ host: "Excel", skills, shareContext: true, store: s, sendTimeoutMs: 80 });
    let step = 0;
    const asked: string[] = [];
    const sse = (events: any[]) => new Response(events.map((e) => `data: ${JSON.stringify(e)}

`).join(""), { status: 200 });
    const fetchFn: any = async () => {
      step++;
      return step === 1
        ? sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "write_other_app", arguments: JSON.stringify({ app: "Word", text: "Tóm tắt" }) } }] } }] }])
        : sse([{ choices: [{ delta: { content: "Chưa ghi được." } }] }]);
    };
    const tools: any = { host: "Excel", specs: [], execute: async () => ({ ok: true }), overview: async () => ({}), systemPrompt: () => "x" };
    const msgs = await runAgent({
      provider: { id: "groq", name: "Groq", kind: "openai", baseUrl: "https://x.invalid/v1" } as any,
      key: "k",
      model: "m",
      tools,
      extra,
      system: "x",
      messages: [{ role: "user", content: "ghi tóm tắt sang Word" }],
      log: [],
      onEvent: () => undefined,
      confirmRisky: false,
      askBeforeWrite: false,
      approve: async (_c, reason) => {
        asked.push(reason || "");
        return false;
      },
      fetchFn,
      retryDelaysMs: [],
    });
    assert.strictEqual(asked.length, 1);
    assert.match(asked[0], /Ghi sang Word/);
    assert.strictEqual(s.getItem(CMD_PREFIX + "Word"), null, "bị từ chối mà lệnh vẫn rời khung");
    const toolMsg = msgs.find((m) => m.role === "tool") as any;
    assert.match(toolMsg.content, /không cho phép/);
    assert.ok(!writeSucceeded(JSON.parse(toolMsg.content), "lỗi"));
  });
});
