/** Đợt 12 (tasks/dot-12.md). Phần A: đo token đầu vào trước khi gửi, chọn model đủ sức chứa. */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { runTurn } from "../../src/core/agent/turn";
import { FreeRotation } from "../../src/core/agent/rotation";
import { HostToolSet, ToolSpec } from "../../src/core/agent/types";
import { biggerModels, estimateInputTokens, fitsContext, sizeNote, tokenRatio, tooLargeMessage } from "../../src/core/inputsize";
import { pptTools, PPT_SYSTEM_PROMPT } from "../../src/office/pptTools";

function sse(events: any[]): Response {
  const text = events.map((e) => `data: ${typeof e === "string" ? e : JSON.stringify(e)}\n\n`).join("");
  return new Response(text, { status: 200, headers: { "content-type": "text/event-stream" } });
}
const oaText = (t: string) => sse([{ choices: [{ delta: { content: t } }] }, { choices: [{ delta: {}, finish_reason: "stop" }] }, "[DONE]"]);
const counting = () => {
  const calls: string[] = [];
  const fetchFn = (async (url: string) => (calls.push(String(url)), oaText("OK"))) as any;
  return { calls, fetchFn };
};
const host: HostToolSet = { host: "PowerPoint", specs: pptTools.specs, execute: async () => ({ ok: true }), overview: async () => ({}), systemPrompt: () => PPT_SYSTEM_PROMPT };
const memStore = () => {
  const m = new Map<string, string>();
  return { get: (k: string) => m.get(k) ?? null, set: (k: string, v: string) => void m.set(k, v) };
};

describe("Đợt 12 · A · đo token đầu vào", () => {
  it("Tính cả lời nhắc hệ thống, lệnh gọi công cụ và bộ công cụ (trước chỉ đếm tin nhắn); hai hệ số: JSON / lời nhắc 2,6 · chữ 3,6", () => {
    const msgs: any[] = [{ role: "system", content: "s".repeat(300) }, { role: "user", content: "u".repeat(300) }];
    const base = estimateInputTokens(msgs);
    assert.strictEqual(base, Math.ceil(300 / 2.6 + 300 / 3.6));
    const withTools = estimateInputTokens(msgs, pptTools.specs);
    const toolChars = pptTools.specs.reduce((n, t) => n + JSON.stringify({ name: t.name, description: t.description, parameters: t.parameters }).length, 0);
    assert.strictEqual(withTools, Math.ceil((300 + toolChars) / 2.6 + 300 / 3.6));
    // Kết quả công cụ (JSON) tính theo hệ số của JSON
    assert.strictEqual(estimateInputTokens([{ role: "tool", content: "j".repeat(260) } as any]), 100);
    assert.ok(withTools - base > 3000, `bộ công cụ PowerPoint phải đáng kể: ${withTools - base}`);
    const call: any = { role: "assistant", content: "", toolCalls: [{ id: "1", name: "x", args: { a: "b".repeat(60) } }] };
    assert.ok(estimateInputTokens([call]) > 20);
  });
  it("Hệ số theo họ model: Claude đếm chữ tiếng Việt dày hơn (đo 22/11: Haiku 5.5 đếm 23.378 token khi hệ số chung ước 13.689)", () => {
    for (const m of ["claude-haiku-5-5", "anthropic/claude-sonnet-5.5", "us.anthropic.claude-opus-5-5"]) assert.deepStrictEqual(tokenRatio(m), { struct: 2.1, text: 1.5 });
    for (const m of ["nvidia/nemotron-3-ultra-550b-a55b:free", "gemini-3.8-flash", "openai/gpt-oss-120b", undefined]) assert.deepStrictEqual(tokenRatio(m), { struct: 2.6, text: 3.6 });
    // Chữ tiếng Việt (≈ 1/5 ký tự ngoài ASCII): Claude 1,5 ký tự một token
    const vi = "Bên Thuê có nghĩa vụ thanh toán đúng hạn ".repeat(75);
    const share = [...vi].filter((c) => c.charCodeAt(0) > 127).length / vi.length;
    assert.ok(share >= 0.2, `tỷ lệ ${share}`);
    assert.strictEqual(estimateInputTokens([{ role: "user", content: vi } as any], [], "claude-haiku-5-5"), Math.ceil(vi.length / 1.5));
    assert.strictEqual(estimateInputTokens([{ role: "user", content: vi } as any], [], "gemini-3.8-flash"), Math.ceil(vi.length / 3.6));
    // Soát mã 22/11: chữ tiếng Anh với Claude khoảng 3,8 ký tự một token (không chặn nhầm hợp đồng tiếng Anh dài)
    const en = "The Lessee shall pay the rent on time ".repeat(80);
    assert.strictEqual(estimateInputTokens([{ role: "user", content: en } as any], [], "claude-haiku-5-5"), Math.ceil(en.length / 3.8));
  });
  it("Sức chứa: chừa 10%; không biết sức chứa thì coi như chứa nổi", () => {
    assert.strictEqual(fitsContext(90000, 100000), true);
    assert.strictEqual(fitsContext(90001, 100000), false);
    assert.strictEqual(fitsContext(10 ** 7), true);
  });
  it("Gợi ý model cùng nguồn chứa nổi: nhỏ nhất trước, bỏ model đang chọn và model không rõ sức chứa", () => {
    const models = [
      { id: "a", context: 32000 },
      { id: "b", name: "B lớn", context: 1000000 },
      { id: "c", context: 200000 },
      { id: "d" },
      { id: "cur", context: 500000 },
    ];
    assert.deepStrictEqual(biggerModels(models, 100000, "cur").map((m) => m.id), ["c", "b"]);
    const t = tooLargeMessage({ tokens: 150000, context: 128000, model: "m-128k", alternatives: biggerModels(models, 150000, "cur") });
    assert.match(t, /^Chưa gửi: lượt này khoảng 150\.000 token/);
    assert.match(t, /m-128k \(128\.000 token\)/);
    assert.match(t, /c \(200\.000\), B lớn \(1\.000\.000\)/);
    assert.match(tooLargeMessage({ tokens: 9, context: 1, model: "x", alternatives: [] }), /không có model nào đủ sức chứa/);
    assert.strictEqual(sizeNote(25000, 128000), "Lượt này khoảng 25.000 token đầu vào — model chứa 128.000 token.");
  });

  it("Trả phí, model không chứa nổi: KHÔNG gọi nhà cung cấp, báo lý do too_large kèm gợi ý, lịch sử giữ nguyên", async () => {
    const f = counting();
    const ev: any[] = [];
    const history: any[] = [{ role: "user", content: "trước" }, { role: "assistant", content: "đã trả lời" }];
    const r = await runTurn({
      provider: getProvider("groq")!,
      key: "K",
      model: "nho",
      context: 8000,
      alternatives: [{ id: "nho", context: 8000 }, { id: "lon", name: "Lớn", context: 131072 }],
      host: "PowerPoint",
      history,
      text: "Tóm tắt tệp đính kèm\n" + "chữ ".repeat(8000),
      onEvent: (e) => ev.push(e),
      fetchFn: f.fetchFn,
      agent: { tools: host, log: [] },
    });
    assert.strictEqual(f.calls.length, 0);
    assert.strictEqual((r.error as any)?.reason, "too_large");
    assert.match(String((r.error as any)?.message), /^Chưa gửi.*Lớn \(131\.072\)/);
    assert.strictEqual(r.history, history);
    const sz = ev.find((e) => e.type === "input_size");
    assert.ok(sz && sz.tokens > 8000 && sz.context === 8000);
  });
  it("Trả phí, model chứa nổi: gửi bình thường, vẫn báo số đo", async () => {
    const f = counting();
    const ev: any[] = [];
    const r = await runTurn({ provider: getProvider("groq")!, key: "K", model: "lon", context: 131072, host: "PowerPoint", history: [], text: "Xin chào", onEvent: (e) => ev.push(e), fetchFn: f.fetchFn, agent: { tools: host, log: [] } });
    assert.strictEqual(r.error, undefined);
    assert.strictEqual(f.calls.length, 1);
    assert.ok(ev.some((e) => e.type === "input_size" && e.tokens > 3000));
  });
  it("Trò chuyện thường (không tác tử) cũng đo và chặn khi không chứa nổi", async () => {
    const f = counting();
    const r = await runTurn({ provider: getProvider("groq")!, key: "K", model: "nho", context: 1000, host: "Word", history: [], text: "x".repeat(6000), onEvent: () => {}, fetchFn: f.fetchFn });
    assert.strictEqual(f.calls.length, 0);
    assert.strictEqual((r.error as any)?.reason, "too_large");
  });
  it("Miễn phí: không chặn ở lượt mà để vòng xoay bỏ nguồn; vòng xoay tính cả bộ công cụ", async () => {
    // Tin nhắn nhỏ nhưng bộ công cụ lớn: nguồn sức chứa 3.000 token phải bị bỏ qua, không gọi thử
    const calls: string[] = [];
    const fetchFn = (async (url: string) => (calls.push(String(url)), oaText("vừa"))) as any;
    const P = (id: string) => getProvider(id)!;
    const rot = new FreeRotation(
      [
        { provider: P("groq"), key: "K1", model: "m", context: 3000 },
        { provider: P("openrouter"), key: "K2", model: "m", context: 200000 },
      ],
      { store: memStore() as any },
    );
    const res = await rot.call({ messages: [{ role: "user", content: "Xin chào" }], tools: pptTools.specs as ToolSpec[], onText: () => {}, fetchFn });
    assert.strictEqual(res.text, "vừa");
    assert.ok(calls.length === 1 && /openrouter/.test(calls[0]), calls.join(" "));
  });
  it("Miễn phí, mọi nguồn quá nhỏ: câu lỗi nói cách làm nhỏ lượt, không tự sang trả phí", async () => {
    const fetchFn = (async () => oaText("không được gọi")) as any;
    const rot = new FreeRotation([{ provider: getProvider("groq")!, key: "K", model: "m", context: 1000 }], { store: memStore() as any });
    await assert.rejects(rot.call({ messages: [{ role: "user", content: "x".repeat(9000) }], tools: [], onText: () => {}, fetchFn }), (e: any) => {
      assert.match(e.message, /quá lớn cho mọi nguồn miễn phí: bớt tệp đính kèm/);
      assert.match(e.message, /không tự chuyển sang nguồn trả phí/);
      return true;
    });
  });
});

import { convertDateTexts, dateTextToSerial, isDateNumberFormat } from "../../src/core/datetext";

describe("Đợt 12 · B · X3: chữ dạng ngày thành ngày thật khi đặt định dạng ngày", () => {
  it("Đọc theo ngày/tháng/năm; năm đứng đầu thì năm-tháng-ngày; dữ liệu thật của bài X3", () => {
    const s = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
    assert.strictEqual(dateTextToSerial("05/03/2026"), s(2026, 3, 5));
    for (const [t, d] of [["2026-03-07", 7], ["7-3-2026", 7], ["12/3/2026", 12], ["2026/03/15", 15], ["18.03.2026", 18], ["25-03-2026", 25], [" 20/03/2026 ", 20]] as [string, number][])
      assert.strictEqual(dateTextToSerial(t), s(2026, 3, d), t);
    assert.strictEqual(dateTextToSerial("01/01/1900"), 2);
  });
  it("Không phải ngày có thật hoặc không phải chữ dạng ngày → null (không đổi nhầm)", () => {
    for (const t of ["31/02/2026", "00/03/2026", "12/13/2026", "0912345678", "Ngày đặt cọc", "05/03/26", "2026-3", "1.234.567", "5/3", "", "=TODAY()"]) assert.strictEqual(dateTextToSerial(t), null, t);
    assert.strictEqual(dateTextToSerial(45000 as any), null);
  });
  it("Định dạng ngày: có d hoặc y; giờ phút, số, phần trăm, chữ thì không", () => {
    for (const f of ["dd/mm/yyyy", "m/d/yyyy", "[$-vi-VN]dd/mm/yyyy", "d-mmm-yy", "yyyy"]) assert.ok(isDateNumberFormat(f), f);
    for (const f of ["General", "@", "#,##0", "0.0%", "mm:ss", "h:mm", '"ngày"0', "[Red]0.00", ""]) assert.ok(!isDateNumberFormat(f), f);
  });
  it("Chỉ đổi ô chữ dạng ngày; ô số, công thức, rỗng giữ nguyên; báo ô chữ không đọc được", () => {
    const fix = convertDateTexts([["Ngày đặt cọc"], ["05/03/2026"], [46000], ["=TODAY()"], [""], ["không rõ"], ["31/02/2026"]]);
    assert.strictEqual(fix.converted, 1);
    assert.deepStrictEqual(fix.formulas.map((r) => r[0]), ["Ngày đặt cọc", dateTextToSerial("05/03/2026"), 46000, "=TODAY()", "", "không rõ", "31/02/2026"]);
    assert.deepStrictEqual(fix.leftText, [[0, 0], [5, 0], [6, 0]]);
  });
});

import { EXCEL_SPECS } from "../../src/office/excelTools";

describe("Đợt 12 · B · X6: định dạng nhiều vùng khác kiểu trong một lệnh", () => {
  it("excel_format_range có `more` (mục cần address); mô tả nói cùng một lệnh", () => {
    const s = EXCEL_SPECS.find((x) => x.name === "excel_format_range")!;
    const more = s.parameters.properties.more;
    assert.strictEqual(more.type, "array");
    assert.deepStrictEqual(more.items.required, ["address"]);
    assert.match(s.description, /more/);
  });
});

import { auditModel } from "../../src/core/modelaudit";
import { cellAddress } from "../../src/office/util";

describe("Đợt 12 · B · X5: kiểm toán mô hình bắt đủ 5 lỗi cài sẵn, không báo nhầm", () => {
  // Lưới R1C1 đúng như may/xl/mo-hinh-loi.js của buổi so sánh 22/11 (A1:K13)
  const C = 8;
  const row = (f: (k: number) => any) => Array.from({ length: C }, (_, k) => f(k));
  const g: any[][] = [
    ["Chỉ tiêu (tỷ đồng)", "", "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Tổng"],
    ["Số căn bán", "", 0, 10, 15, 12, 20, 18, 14, 9, ""],
    ["Giá bán bình quân (tỷ/căn)", "", ...row(() => "=R12C2"), ""],
    ["Doanh thu", "", ...row((k) => (k === 3 ? 38.4 : "=R[-2]C*R[-1]C")), "=SUM(RC[-8]:RC[-1])"],
    ["Chi phí xây dựng", "", 45, 40, 35, 30, 20, 10, 5, 2, "=SUM(RC[-8]:RC[-1])"],
    ["Chi phí bán hàng", "", ...row((k) => (k === 5 ? "=R[-2]C*0.03" : "=R[-2]C*R13C2")), "=SUM(RC[-8]:RC[-1])"],
    ["Tổng chi phí", "", ...row((k) => (k === 4 ? "=R[-2]C+R[-1]C[-1]" : "=R[-2]C+R[-1]C")), "=SUM(RC[-8]:RC[-1])"],
    ["Dòng tiền thuần", "", ...row(() => "=R[-4]C-R[-1]C"), "=SUM(RC[-8]:RC[-2])"],
    ["Lũy kế", "", ...row((k) => (k === 0 ? "=R[-1]C" : k === 2 ? "=RC[-2]+R[-1]C" : "=RC[-1]+R[-1]C")), ""],
    Array(11).fill(""),
    Array(11).fill(""),
    ["Giá bán (tỷ/căn)", 3.2, ...Array(9).fill("")],
    ["Tỷ lệ chi phí bán hàng", 0.03, ...Array(9).fill("")],
  ];
  const run = (grid: any[][]) => auditModel({ r1c1: grid, origin: { col: 0, row: 0 }, addr: cellAddress });

  it("F4 số gõ cứng (ô dưới là hàng nhập tay), H6 tỷ lệ gõ, G7 lệch cột, E9 lũy kế bỏ quý, K8 SUM thiếu cột J", () => {
    const r = run(g);
    assert.deepStrictEqual(r.findings.filter((f) => f.kind === "hardcoded").map((f) => f.cell), ["F4"]);
    assert.deepStrictEqual(r.findings.filter((f) => f.kind === "inconsistent").map((f) => f.cell).sort(), ["E9", "G7", "H6", "K8"]);
    assert.deepStrictEqual(r.findings.filter((f) => f.kind === "embedded_rate").map((f) => f.cell), ["H6"]);
    assert.match(r.findings.find((f) => f.cell === "K8")!.detail, /cuối dãy/);
  });
  it("Mô hình đã sửa hết: không báo gì", () => {
    const ok = g.map((r) => r.slice());
    ok[3][5] = "=R[-2]C*R[-1]C";
    ok[5][7] = "=R[-2]C*R13C2";
    ok[6][6] = "=R[-2]C+R[-1]C";
    ok[7][10] = "=SUM(RC[-8]:RC[-1])";
    ok[8][4] = "=RC[-1]+R[-1]C";
    assert.deepStrictEqual(run(ok).findings, []);
  });
  it("Ca dễ báo nhầm của luật ô cuối dãy: dòng tổng khác hàm, dãy ngắn (2 ô), ô cuối còn công thức phía sau", () => {
    assert.strictEqual(run([["=R[-1]C*2"], ["=R[-1]C*2"], ["=R[-1]C*2"], ["=SUM(R[-3]C:R[-1]C)"]]).findings.length, 0);
    assert.strictEqual(run([["=R[-1]C*2"], ["=R[-1]C*2"], ["=R[-1]C*3"]]).findings.length, 0);
    assert.strictEqual(run([["x", "=RC[-1]*2", "=RC[-1]*2", "=RC[-1]*2", "=RC[-1]*2", "=AVERAGE(RC[-4]:RC[-1])"]]).findings.length, 0);
    // Cột nhập tay xen giữa các cột công thức vẫn không bị báo (ô liền kề cũng kẹp giữa công thức)
    const colInput = [["KH1", "TH1", "KH2"], ["=RC[5]*2", 10, "=RC[5]*2"], ["=RC[5]*2", 20, "=RC[5]*2"], ["=RC[5]*2", 30, "=RC[5]*2"]];
    assert.strictEqual(run(colInput).counts.hardcoded, 0);
  });
});

import { targetsLogSheet } from "../../src/core/auditlog";
import { tableColumnWidths } from "../../src/core/tablefit";

describe("Đợt 12 · Soát mã độc lập 22/11 (mã từ 4.11.0)", () => {
  it("C1: bảng nhiều cột không ra bề ngang âm; tổng vẫn đúng; không cột nào dưới phần chia đều khi sàn chiếm hết", () => {
    for (const n of [12, 13, 14, 20, 40]) {
      const vals = [Array.from({ length: n }, () => "ab"), Array.from({ length: n }, (_, k) => "x".repeat(k % 5))];
      const w = tableColumnWidths(vals, 600);
      assert.strictEqual(w.length, n);
      assert.ok(w.every((x) => x > 0), `${n} cột: ${w}`);
      assert.ok(Math.abs(w.reduce((a, b) => a + b, 0) - 600) < 0.05, `${n} cột tổng ${w.reduce((a, b) => a + b, 0)}`);
    }
  });
  it("C2: mục trong `more` nhắm sheet nhật ký cũng bị chặn; mục bình thường thì không", () => {
    assert.ok(targetsLogSheet({ address: "A1", more: [{ address: "'Nhật ký AI'!A1:F50", fill_color: "none" }] }));
    assert.ok(targetsLogSheet({ address: "A1", more: [{ address: "B2", sheet: "Nhật ký AI" }] }));
    assert.ok(!targetsLogSheet({ address: "A1", more: [{ address: "B2:B9", bold: true }, { address: "'Dòng tiền'!C3" }] }));
  });
  it("L6: ô cuối dãy cố ý khác HẰNG SỐ (tăng giá năm cuối) không bị báo là vùng hụt; khác địa chỉ thì vẫn báo", () => {
    const run = (g: any[][]) => auditModel({ r1c1: g, origin: { col: 0, row: 0 }, addr: cellAddress });
    assert.strictEqual(run([["=RC[-1]*1.03", "=RC[-1]*1.03", "=RC[-1]*1.03", "=RC[-1]*1.05"]].map((r) => ["Giá", ...r])).counts.inconsistent, 0);
    assert.strictEqual(run([["x", "=SUM(R1C2:R5C2)", "=SUM(R1C3:R5C3)", "=SUM(R1C4:R5C4)", "=SUM(R1C5:R4C5)"]]).counts.inconsistent, 0, "R1C1 tuyệt đối mỗi cột một khác: không phải một dãy cùng mẫu");
    assert.strictEqual(run([["=SUM(RC[-8]:RC[-1])"], ["=SUM(RC[-8]:RC[-1])"], ["=SUM(RC[-8]:RC[-1])"], ["=SUM(RC[-8]:RC[-2])"]]).counts.inconsistent, 1);
  });
});

import { DOCX_INLINE_MAX_CHARS, docxOutlineText, docxParagraphs, formatNumber, parseParaRange, readDocxSection } from "../../src/core/docxparas";
import { buildOutline } from "../../src/core/outline";
import { DOCX_MIME, withAttachments } from "../../src/core/attachments";

describe("Đợt 12 · D · .docx đính kèm: số điều như Word hiện, mục lục, đọc theo điều", () => {
  // Cấu trúc như hợp đồng thử (kiem-thu/hop-dong-thu.mjs): số điều do danh sách tự động gắn vào kiểu đoạn Heading1/2/3
  const numbering =
    '<w:numbering><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Điều %1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%3)"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
  const styles =
    '<w:styles><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="H3con"><w:basedOn w:val="Heading3"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style></w:styles>';
  const p = (style: string, t: string) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const doc =
    "<w:document><w:body>" +
    "<w:p><w:r><w:t>HỢP ĐỒNG THUÊ</w:t></w:r></w:p>" +
    p("Heading1", "Định nghĩa") + p("Heading2", "“Bên Thuê” là Công ty A &amp; B") +
    p("Heading1", "Tiền thuê") + p("Heading2", "Đơn giá là 450.000 đồng") +
    '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Tăng giá </w:t></w:r><w:del><w:r><w:delText>5%</w:delText></w:r></w:del><w:ins><w:r><w:t>3%</w:t></w:r></w:ins><w:r><w:tab/><w:t>mỗi hai năm</w:t></w:r></w:p>' +
    p("H3con", "điện bị cắt quá 72 giờ") + p("H3con", "nước bị cắt quá 72 giờ") +
    p("Heading1", "Thanh toán") + p("Heading2", "Trả theo quý") +
    "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>ô bảng</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
    "<w:p/>" +
    "</w:body></w:document>";
  const rows = docxParagraphs({ document: doc, styles, numbering });
  const items = buildOutline(rows);

  it("Số như Word hiện: Điều 1., 1.1., Điều 2., 2.1., 2.2., a), b) (kiểu đoạn kế thừa basedOn); bộ đếm cấp dưới đặt lại", () => {
    assert.deepStrictEqual(rows.map((r) => r.listString || ""), ["", "Điều 1.", "1.1.", "Điều 2.", "2.1.", "2.2.", "a)", "b)", "Điều 3.", "3.1.", "", ""]);
    assert.strictEqual(rows[2].text, "“Bên Thuê” là Công ty A & B");
  });
  it("Chữ hiện hành: bỏ chữ đã xóa có theo dõi, giữ chữ chèn, tab; đoạn trong bảng được đánh dấu", () => {
    assert.strictEqual(rows[5].text, "Tăng giá 3%\tmỗi hai năm");
    assert.strictEqual(rows[10].inTable, true);
    assert.strictEqual(rows.length, 12);
  });
  it("Mục lục và tra điều: '2.2(b)' → đúng đoạn; dải đoạn; phần đọc dài có `next`", () => {
    const sec = readDocxSection(rows, items, { clause: "2.2(b)" });
    assert.strictEqual(sec.text, "[¶7] b) nước bị cắt quá 72 giờ");
    const art = readDocxSection(rows, items, { clause: "Điều 2" });
    assert.deepStrictEqual([art.from, art.to], [3, 7]);
    assert.match(art.text, /^\[¶3\] Điều 2\. Tiền thuê\n\[¶4\] 2\.1\. Đơn giá/);
    assert.deepStrictEqual([readDocxSection(rows, items, { paragraphs: "8-9" }).from, readDocxSection(rows, items, { paragraphs: "8-9" }).to], [8, 9]);
    const cut = readDocxSection(rows, items, { paragraphs: "1-9" }, 60);
    assert.ok(cut.next !== undefined && cut.to < 9, JSON.stringify(cut));
    assert.throws(() => readDocxSection(rows, items, { clause: "Điều 9" }), /Không thấy "Điều 9"/);
    assert.throws(() => readDocxSection(rows, items, {}), /Cần clause/);
    const ol = docxOutlineText(items);
    assert.match(ol, /^Điều 1 \[¶1–2\] Định nghĩa$/m);
    assert.match(ol, /^2\.2 \[¶5–7\] Tăng giá/m);
    assert.ok(!/^\(?a\)?/m.test(ol.split("\n").filter((l) => /^a/.test(l)).join("")), "mục lục gọn: không kê điểm a), b)");
  });
  it("Dải đoạn, định dạng số của Word", () => {
    assert.deepStrictEqual(parseParaRange("120-160", 200), [120, 160]);
    assert.deepStrictEqual(parseParaRange("¶87", 200), [87, 87]);
    assert.deepStrictEqual(parseParaRange("190-260", 200), [190, 199]);
    assert.strictEqual(parseParaRange("5-2", 200), null);
    assert.strictEqual(parseParaRange("abc", 200), null);
    assert.deepStrictEqual([1, 4, 9, 14, 27].map((n) => formatNumber(n, "lowerRoman")), ["i", "iv", "ix", "xiv", "xxvii"]);
    assert.deepStrictEqual([1, 26, 27].map((n) => formatNumber(n, "lowerLetter")), ["a", "z", "aa"]);
    assert.strictEqual(formatNumber(3, "decimalZero"), "03");
  });
  it("Tệp dài ở Word (tác tử): gửi mục lục + cách đọc thay cho cả tệp; tệp ngắn, ngoài Word, không có bản gốc: như cũ", () => {
    const long = { name: "HĐ.docx", kind: "text" as const, mime: DOCX_MIME, data: "UEsD", text: "x".repeat(DOCX_INLINE_MAX_CHARS + 10), outline: "Điều 1 [¶1–2] Định nghĩa", paraCount: 165, fullChars: 168817 };
    const out = withAttachments("Điều 18.2(b) nói gì?", [long], { docTool: true, compare: true });
    assert.match(out, /nội dung CHƯA được gửi kèm vì tệp dài/);
    assert.match(out, /read_attachment_text với file = "HĐ\.docx"/);
    assert.match(out, /Điều 1 \[¶1–2\] Định nghĩa/);
    assert.match(out, /word_compare_versions/);
    assert.ok(!out.includes("xxxxxxxxxx"));
    assert.match(withAttachments("q", [long], { docTool: false }), /x{50}/);
    assert.match(withAttachments("q", [{ ...long, text: "ngắn", fullChars: 4 }], { docTool: true }), /ngắn/);
    assert.match(withAttachments("q", [{ ...long, data: undefined }], { docTool: true }), /x{50}/);
  });
});

describe("Đợt 12 · D · mục lục dài vẫn đủ cả văn bản", () => {
  it("Quá giới hạn thì chỉ kê cấp Điều, có đủ tới Điều cuối và Phụ lục", () => {
    const items: any[] = [];
    for (let a = 1; a <= 34; a++) {
      items.push({ i: a * 4, end: a * 4 + 3, num: `Điều ${a}`, level: 0, title: `Tiêu đề điều ${a} khá dài để chiếm chỗ trong mục lục` });
      for (let k = 1; k <= 3; k++) items.push({ i: a * 4 + k, end: a * 4 + k, num: `${a}.${k}`, level: 1, title: "Nội dung khoản rất dài ".repeat(4) });
    }
    items.push({ i: 149, end: 162, num: "Phụ lục 01", level: 0, title: "MÔ TẢ MẶT BẰNG" });
    const ol = docxOutlineText(items, 6000);
    assert.ok(ol.length <= 6200, String(ol.length));
    assert.match(ol, /^Điều 34 \[¶136–139\]/m);
    assert.match(ol, /^Phụ lục 01 \[¶149–162\]/m);
    assert.ok(!/^\d+\.\d+ /m.test(ol), "không kê khoản khi đã quá giới hạn");
    assert.match(ol, /chỉ kê cấp Điều/);
    // Vừa giới hạn thì kê cả khoản cấp một như trước
    assert.match(docxOutlineText(items.slice(0, 8), 6000), /^1\.1 /m);
  });
});
