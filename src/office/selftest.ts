/* global Excel, Word, PowerPoint, Office */
/** Tự kiểm tra: chạy thật từng công cụ trên tài liệu đang mở rồi dọn sạch. Không cần API key. */
import { HostToolSet, UndoEntry } from "../core/agent/types";
import { runWordSelfTest } from "./selftestWord";
import { APPS, AppName, CMD_PREFIX, beat, buildPayload, reply, sendCommand, takeCommand } from "../core/crossapp";
import { cellAsText, executeIncoming } from "./crossExec";
import { excelTools } from "./excelTools";
import { withAuditTrail } from "./excelTrail";
import { runExcelModelSelfTest } from "./selftestExcel";

export interface TestStep {
  name: string;
  /** true = đạt, false = lỗi, null = bỏ qua (máy không hỗ trợ) */
  ok: boolean | null;
  detail: string;
}

export class Skip extends Error {}

export async function runSelfTest(tools: HostToolSet, onProgress: (steps: TestStep[]) => void): Promise<TestStep[]> {
  const steps: TestStep[] = [];
  const log: UndoEntry[] = [];
  // Excel: chạy trên bộ công cụ gốc với nhật ký và tô sáng TẮT, để bài thử không ghi vào sheet "Nhật ký AI" thật của người dùng
  const target = tools.host === "Excel" ? withAuditTrail(excelTools, { settings: () => ({}), describe: (c) => c.name }) : tools;
  const run = (name: string, args: any = {}) => target.execute(name, args, { log });
  const step = async (name: string, fn: () => Promise<string | void>) => {
    try {
      const d = await fn();
      steps.push({ name, ok: true, detail: d || "" });
    } catch (e: any) {
      if (e instanceof Skip) steps.push({ name, ok: null, detail: e.message });
      else {
        let detail = e?.message || String(e);
        if (e?.debugInfo?.errorLocation) detail += ` (${e.debugInfo.errorLocation})`;
        steps.push({ name, ok: false, detail });
      }
    }
    onProgress([...steps]);
  };
  const expect = (cond: any, msg: string) => {
    if (!cond) throw new Error(msg);
  };
  /** Lời cảnh báo rủi ro của một lần gọi công cụ (null = không rủi ro), đúng hàm mà vòng lặp tác tử dùng để quyết định có hỏi người dùng */
  const riskOf = async (name: string, args: any): Promise<string | null> => {
    const spec = tools.specs.find((sp) => sp.name === name);
    if (!spec?.risky) throw new Error(`Công cụ ${name} chưa có kiểm tra rủi ro`);
    return spec.risky(args);
  };

  if (tools.host === "Excel") {
    const sheet = "AIDN_KT_" + Date.now().toString(36).slice(-4);
    // Dọn sheet thử còn sót từ lần chạy trước
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets;
      ws.load("items/name");
      await ctx.sync();
      ws.items.filter((w) => /^AIDN_KT_/.test(w.name)).forEach((w) => w.delete());
      await ctx.sync();
    }).catch(() => undefined);
    await step("Xem tổng quan sổ tính", async () => {
      const ov = await run("excel_get_overview");
      expect(ov.sheets?.length, "Không đọc được danh sách sheet");
      return `${ov.sheets.length} sheet, đang mở "${ov.activeSheet}"`;
    });
    await step("Tạo sheet thử", async () => (await run("excel_add_sheet", { name: sheet })).message);
    await step("Ghi giá trị + công thức", async () => {
      const r = await run("excel_write_range", {
        sheet,
        address: "A1",
        values: [["Tháng", "Doanh thu"], ["1", "100"], ["2", "200"], ["Tổng", "=SUM(B2:B3)"]],
      });
      const total = r.result_preview?.[3]?.[1];
      expect(total === 300, `Ô tổng B4 = ${total}, cần 300`);
      return "B4 = SUM(B2:B3) = 300";
    });
    await step("Đọc lại dữ liệu", async () => {
      const r = await run("excel_read_range", { sheet, address: "A1:B4" });
      expect(r.formulas?.B4 === "=SUM(B2:B3)", `Công thức đọc được: ${JSON.stringify(r.formulas)}`);
      expect(r.values?.[2]?.[1] === 200, "Giá trị B3 sai");
      return "Đọc đúng giá trị và công thức";
    });
    await step("Sắp xếp", async () => {
      await run("excel_write_range", { sheet, address: "D1", values: [["Số"], ["3"], ["1"], ["2"]] });
      await run("excel_sort_range", { sheet, address: "D1:D4", column: "D", ascending: true });
      const r = await run("excel_read_range", { sheet, address: "D2:D4" });
      const got = r.values.map((x: any[]) => x[0]).join(",");
      expect(got === "1,2,3", `Sau sắp xếp: ${got}`);
      return "3,1,2 → 1,2,3";
    });
    await step("Định dạng", async () => {
      await run("excel_format_range", {
        sheet,
        address: "A1:B1",
        bold: true,
        fill_color: "#DDEBF7",
        borders: "all",
        autofit_columns: true,
      });
      await run("excel_format_range", { sheet, address: "B2:B4", number_format: "#,##0" });
      const bold = await Excel.run(async (ctx) => {
        const f = ctx.workbook.worksheets.getItem(sheet).getRange("A1:B1").format.font;
        f.load("bold");
        await ctx.sync();
        return f.bold;
      });
      expect(bold === true, "Chữ tiêu đề chưa đậm");
      return "Đậm, tô nền, kẻ viền, định dạng số";
    });
    await step("Tạo bảng (Table)", async () => (await run("excel_create_table", { sheet, address: "A1:B4" })).message);
    await step("Vẽ biểu đồ", async () =>
      (await run("excel_create_chart", { sheet, source_address: "A1:B3", chart_type: "column", title: "Thử" })).message
    );
    await step("Xóa vùng", async () => {
      await run("excel_clear_range", { sheet, address: "D1:D4" });
      const r = await run("excel_read_range", { sheet, address: "D1:D4" });
      expect(r.values.every((x: any[]) => x[0] === ""), "Vùng D1:D4 chưa trống");
      return "D1:D4 đã trống";
    });
    let pivotSheet = "";
    await step("Cảnh báo trước khi ghi đè ô có dữ liệu", async () => {
      const spec = tools.specs.find((x) => x.name === "excel_write_range");
      expect(spec?.risky, "Công cụ ghi chưa có kiểm tra rủi ro");
      const msg = await spec!.risky!({ sheet, address: "A1", values: [["x", "y"]] });
      expect(msg && /Ghi đè 2 ô/.test(msg), `Không cảnh báo đúng: ${msg}`);
      const none = await spec!.risky!({ sheet, address: "Z50", values: [["x"]] });
      expect(!none, "Cảnh báo nhầm ở vùng trống");
      return msg!;
    });
    await step("Tìm kiếm", async () => {
      const r = await run("excel_find", { text: "Doanh thu", sheet });
      expect(r.total_cells >= 1, r.message);
      return `Thấy ở ${r.results[0].addresses}`;
    });
    await step("Quét lỗi công thức", async () => {
      await run("excel_write_range", { sheet, address: "F1", values: [["=1/0"]] });
      const r = await run("excel_find_errors", { sheet });
      expect(r.total_errors >= 1, r.message);
      const c = r.sheets[0].cells[0];
      return `${r.message}: ${c.cell} = ${c.error}`;
    });
    await step("Lần ô được công thức tham chiếu", async () => {
      const r = await run("excel_trace", { cell: `${sheet}!B4`, direction: "precedents" });
      const txt = JSON.stringify(r.nodes);
      expect(/B2/.test(txt) && /B3/.test(txt), `Kết quả: ${txt.slice(0, 200)}`);
      return r.message;
    });
    await step("Định dạng có điều kiện", async () => {
      await run("excel_conditional_format", { sheet, address: "B2:B3", type: "data_bar" });
      const n = await Excel.run(async (ctx) => {
        const cf = ctx.workbook.worksheets.getItem(sheet).getRange("B2:B3").conditionalFormats;
        cf.load("items");
        await ctx.sync();
        return cf.items.length;
      });
      expect(n >= 1, "Chưa có quy tắc định dạng có điều kiện");
      return `Thanh dữ liệu cho B2:B3 (${n} quy tắc)`;
    });
    await step("Danh sách chọn", async () => {
      await run("excel_data_validation", { sheet, address: "G2:G4", list: ["Đạt", "Chưa đạt"] });
      const t = await Excel.run(async (ctx) => {
        const dv = ctx.workbook.worksheets.getItem(sheet).getRange("G2").dataValidation;
        dv.load("type");
        await ctx.sync();
        return dv.type;
      });
      expect(String(t) === "List", `Kiểu kiểm soát dữ liệu: ${t}`);
      return "G2:G4 chọn Đạt / Chưa đạt";
    });
    await step("Ghi chú vào ô", async () => (await run("excel_add_comment", { cell: `${sheet}!B4`, text: "Nguồn: tự kiểm tra" })).message);
    await step("Lọc dữ liệu", async () => {
      await run("excel_write_range", { sheet, address: "K1", values: [["Nhóm", "Giá trị"], ["A", "100"], ["B", "200"], ["C", "300"]] });
      await run("excel_filter", { sheet, address: "K1:L4", column: "L", condition: ">150" });
      const hidden = await Excel.run(async (ctx) => {
        const r = ctx.workbook.worksheets.getItem(sheet).getRange("K2");
        r.load("rowHidden");
        await ctx.sync();
        return r.rowHidden;
      });
      await run("excel_filter", { sheet, address: "K1:L4", clear: true });
      expect(hidden === true, "Dòng có giá trị 100 chưa bị ẩn");
      return "Lọc > 150 ẩn đúng dòng 100, rồi bỏ lọc";
    });
    await step("PivotTable", async () => {
      const r = await run("excel_create_pivot", {
        sheet,
        source_address: "K1:L4",
        rows: ["Nhóm"],
        values: [{ field: "Giá trị", summarize: "sum" }],
      });
      pivotSheet = r.sheet || "";
      const n = await Excel.run(async (ctx) => {
        const pts = ctx.workbook.worksheets.getItem(r.sheet).pivotTables;
        pts.load("items");
        await ctx.sync();
        return pts.items.length;
      });
      expect(n >= 1, `Sheet ${r.sheet} chưa có PivotTable`);
      return r.message;
    });
    await step("Chèn hàng", async () => (await run("excel_insert_delete", { sheet, action: "insert_rows", at: "8", count: 2 })).message);
    await step("Cố định dòng tiêu đề", async () => (await run("excel_sheet", { sheet, action: "freeze_rows", count: 1 })).message);
    await step("Chọn vùng kết quả", async () => (await run("excel_select_range", { sheet, address: "A1:B4" })).message);
    await step("Hoàn tác từng thao tác", async () => {
      const first = log[0];
      const errs: string[] = [];
      for (const u of log.slice(1).reverse()) {
        try {
          await u.undo();
        } catch (e: any) {
          errs.push(`${u.label}: ${e?.message || e}`);
        }
      }
      expect(!errs.length, `Hoàn tác lỗi: ${errs.join("; ")}`);
      const state = await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        const r = ws.getRange("A1:D4");
        r.load("values");
        ws.charts.load("items");
        ws.tables.load("items");
        await ctx.sync();
        const font = ws.getRange("A1").format.font;
        font.load("bold");
        await ctx.sync();
        return {
          empty: r.values.every((row) => row.every((v) => v === "")),
          charts: ws.charts.items.length,
          tables: ws.tables.items.length,
          bold: font.bold,
        };
      });
      expect(state.empty, "Sheet thử chưa trống sau khi hoàn tác");
      expect(state.charts === 0 && state.tables === 0, `Còn ${state.charts} biểu đồ, ${state.tables} bảng`);
      expect(state.bold === false, "Định dạng (chữ đậm) chưa được hoàn tác");
      if (first) await first.undo();
      const gone = await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(sheet);
        await ctx.sync();
        return ws.isNullObject;
      });
      expect(gone, "Sheet thử chưa được xóa");
      return `Hoàn tác ${log.length} thao tác (cả nội dung lẫn định dạng), sổ tính trở lại như cũ`;
    });
    await runExcelModelSelfTest({ base: excelTools, step, expect, riskOf });
    // Bảo đảm không để lại sheet thử dù bước nào lỗi
    for (const name of [pivotSheet, sheet].filter(Boolean)) {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
        await ctx.sync();
        if (!ws.isNullObject) {
          ws.delete();
          await ctx.sync();
        }
      }).catch(() => undefined);
    }
  } else if (tools.host === "Word") {
    await runWordSelfTest({ tools, step, run, expect, riskOf, log, Skip });
  } else if (tools.host === "PowerPoint") {
    let before = 0;
    let slideNo = 0;
    let titleShape = "";
    await step("Xem tổng quan bài trình chiếu", async () => {
      const ov = await run("ppt_get_overview");
      before = ov.slide_count;
      return `${ov.slide_count} slide, ${ov.layouts.length} bố cục`;
    });
    await step("Thêm slide có tiêu đề + nội dung", async () => {
      const r = await run("ppt_add_slide", { layout: "Title and Content", title: "AIDN kiểm tra", body: "Dòng 1\nDòng 2" });
      slideNo = r.slide;
      const ov = await run("ppt_get_overview");
      const s = ov.slides.find((x: any) => x.slide === slideNo);
      const t = s?.shapes.find((x: any) => x.text === "AIDN kiểm tra");
      expect(t, "Không thấy tiêu đề trên slide mới");
      titleShape = t.name;
      const b = s.shapes.find((x: any) => String(x.text || "").includes("Dòng 2"));
      expect(b, "Không thấy nội dung trên slide mới");
      return r.message;
    });
    await step("Sửa chữ trong hình", async () => {
      await run("ppt_set_shape_text", { slide: slideNo, shape: titleShape, text: "AIDN kiểm tra 2", bold: true });
      const ov = await run("ppt_get_overview");
      const s = ov.slides.find((x: any) => x.slide === slideNo);
      expect(s.shapes.some((x: any) => x.text === "AIDN kiểm tra 2"), "Chữ chưa đổi");
      return "Đã đổi tiêu đề";
    });
    await step("Thêm ô chữ", async () =>
      (await run("ppt_add_textbox", { slide: slideNo, text: "Ô chữ thử", left: 600, top: 440, width: 300, height: 40, font_size: 14 })).message
    );
    await step("Thêm hình", async () =>
      (
        await run("ppt_add_shape", {
          slide: slideNo,
          shape_type: "rounded_rectangle",
          text: "Hình thử",
          left: 60,
          top: 430,
          width: 200,
          height: 60,
          fill_color: "#DDEBF7",
          font_color: "#1F4E79",
        })
      ).message
    );
    await step("Thêm bảng", async () => {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) throw new Skip("PowerPoint chưa hỗ trợ chèn bảng");
      return (await run("ppt_add_table", { slide: slideNo, values: [["Mục", "Giá trị"], ["A", "1"]], left: 300, top: 300 })).message;
    });
    const has18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
    let dupNo = 0;
    const layoutCount = () =>
      PowerPoint.run(async (ctx) => {
        const l = ctx.presentation.slideMasters.getItemAt(0).layouts;
        l.load("items/name");
        await ctx.sync();
        return l.items.length;
      });
    await step("Biểu đồ gốc (sửa được dữ liệu)", async () => {
      const nLayouts = await layoutCount();
      const before: string[] = await PowerPoint.run(async (ctx) => {
        const ls = ctx.presentation.slideMasters.getItemAt(0).layouts;
        ls.load("items/name");
        await ctx.sync();
        return ls.items.map((l) => l.name);
      });
      const r = await run("ppt_add_chart", {
        chart_type: "column",
        categories: ["Q1", "Q2", "Q3"],
        series: [{ name: "Doanh thu", values: [10, 20, 15] }],
        title: "AIDN biểu đồ",
        after_slide: slideNo,
      });
      let parts = "không kiểm được tệp";
      if (has18) {
        const b64 = await PowerPoint.run(async (ctx) => {
          const sl = ctx.presentation.slides;
          sl.load("items");
          await ctx.sync();
          const out = sl.items[r.slide - 1].exportAsBase64();
          await ctx.sync();
          return out.value;
        });
        const JSZip = (await import("jszip")).default;
        const zip = await JSZip.loadAsync(b64, { base64: true });
        const files = Object.keys(zip.files);
        const chart = files.filter((f) => /ppt\/charts\/chart\d+\.xml$/.test(f));
        const data = files.filter((f) => /ppt\/embeddings\/.+\.xlsx$/.test(f));
        expect(chart.length >= 1, `Slide ${r.slide} không chứa biểu đồ gốc`);
        parts = `${chart.length} biểu đồ, ${data.length} bảng dữ liệu nhúng`;
      }
      const ov = await run("ppt_get_overview");
      const sl = ov.slides.find((x: any) => x.slide === r.slide);
      const n2 = await layoutCount();
      // Điều người dùng thấy: slide biểu đồ phải dùng một bố cục CÓ SẴN của mẫu (không lệch mẫu).
      expect(before.includes(String(sl?.layout)), `Slide biểu đồ dùng bố cục lạ "${sl?.layout}" (mẫu có: ${before.join(", ")})`);
      // Tổng quan phải báo vị trí thật của ô tiêu đề (trước 4.7.2 báo 0, 0, 0, 0 nên AI tốn một lượt đặt lại)
      const tt = (sl?.shapes || []).find((x: any) => /AIDN biểu đồ/.test(String(x.text || "")));
      expect(tt && tt.box[2] > 0 && tt.box[3] > 0, `Ô tiêu đề slide biểu đồ báo toạ độ ${JSON.stringify(tt?.box)}`);
      // PowerPoint có lần tự thêm một bố cục khi chèn (lần đầu trên bài mới); add-in không xóa được bố cục (không có API) nên chỉ ghi nhận.
      return `${r.message} (${parts}; slide dùng bố cục có sẵn "${sl?.layout}"${n2 === nLayouts ? ", mẫu không bị thêm bố cục" : `; LƯU Ý: PowerPoint đã thêm ${n2 - nLayouts} bố cục không dùng vào mẫu (${nLayouts} → ${n2}), add-in không xóa được`})`;
    });
    await step("Sơ đồ quy trình tự dàn trang", async () => {
      const r = await run("ppt_add_process_flow", {
        slide: slideNo,
        steps: ["Chuẩn bị hồ sơ", "Nộp hồ sơ", "Thẩm định", "Phê duyệt"],
        style: "chevron",
        top: 170,
        height: 80,
      });
      const ov = await run("ppt_get_overview");
      const sh = ov.slides.find((x: any) => x.slide === slideNo).shapes;
      expect(sh.some((x: any) => x.type === "Group") || sh.filter((x: any) => /Chuẩn bị hồ sơ|Phê duyệt/.test(x.text || "")).length >= 2, "Không thấy sơ đồ");
      return r.message;
    });
    await step("Sửa vị trí, màu, chữ của hình", async () => {
      const ov = await run("ppt_get_overview");
      const h = ov.slides.find((x: any) => x.slide === slideNo).shapes.find((x: any) => x.text === "Hình thử");
      expect(h, "Không thấy hình thử");
      await run("ppt_update_shape", { slide: slideNo, shape: h.name, left: 90, fill_color: "#FFE699", font_size: 12 });
      const ov2 = await run("ppt_get_overview");
      const h2 = ov2.slides.find((x: any) => x.slide === slideNo).shapes.find((x: any) => x.name === h.name);
      expect(Math.abs(h2.box[0] - 90) <= 1, `Vị trí trái: ${h2.box[0]}`);
      return `Dời ${h.box[0]} → 90 điểm, đổi màu nền, cỡ chữ 12`;
    });
    await step("Nhân bản và di chuyển slide", async () => {
      if (!has18) throw new Skip("PowerPoint chưa hỗ trợ");
      const nLayouts = await layoutCount();
      const r = await run("ppt_duplicate_slide", { slide: slideNo });
      const n2 = await layoutCount();
      expect(n2 === nLayouts, `Nhân bản làm mẫu thêm bố cục (${nLayouts} → ${n2})`);
      const ov = await run("ppt_get_overview");
      const last = ov.slide_count;
      await run("ppt_move_slide", { slide: r.slide, to: last });
      const ov2 = await run("ppt_get_overview");
      const end = ov2.slides.find((x: any) => x.slide === last);
      expect(end?.shapes.some((x: any) => x.text === "AIDN kiểm tra 2"), "Slide cuối không phải bản sao");
      dupNo = last;
      return `Bản sao ở slide ${r.slide} → chuyển xuống cuối (slide ${last})`;
    });
    await step("Xóa hình (hỏi trước, có hoàn tác)", async () => {
      const ov = await run("ppt_get_overview");
      const t = ov.slides.find((x: any) => x.slide === slideNo).shapes.find((x: any) => x.text === "Ô chữ thử");
      expect(t, "Không thấy ô chữ thử");
      const spec = tools.specs.find((x) => x.name === "ppt_delete_shape");
      const warn = spec?.risky ? await spec.risky({ slide: slideNo, shape: t.name }) : null;
      expect(warn, "Không có cảnh báo trước khi xóa");
      await run("ppt_delete_shape", { slide: slideNo, shape: t.name });
      const ov2 = await run("ppt_get_overview");
      expect(!ov2.slides.find((x: any) => x.slide === slideNo).shapes.some((x: any) => x.text === "Ô chữ thử"), "Ô chữ vẫn còn");
      return `Cảnh báo: "${warn}" → đã xóa`;
    });
    await step("Xóa slide (hỏi trước, có hoàn tác)", async () => {
      if (!dupNo) throw new Skip("Không có bản sao để xóa");
      const ov = await run("ppt_get_overview");
      await run("ppt_delete_slide", { slide: dupNo });
      const ov2 = await run("ppt_get_overview");
      expect(ov2.slide_count === ov.slide_count - 1, `Còn ${ov2.slide_count} slide`);
      return `Đã xóa slide ${dupNo}`;
    });
    await step("Chuyển tới slide", async () => (await run("ppt_go_to_slide", { slide: slideNo })).message);
    await step("Hoàn tác toàn bộ", async () => {
      const errs: string[] = [];
      for (const u of log.slice().reverse()) {
        try {
          await u.undo();
        } catch (e: any) {
          errs.push(`${u.label}: ${e?.message || e}`);
        }
      }
      const ov = await run("ppt_get_overview");
      if (ov.slide_count !== before) {
        // dọn slide thử còn sót (nhận ra theo chữ "AIDN")
        await PowerPoint.run(async (ctx) => {
          const sl = ctx.presentation.slides;
          sl.load("items/id");
          await ctx.sync();
          ov.slides
            .filter((x: any) => x.shapes.some((y: any) => /^AIDN (kiểm tra|biểu đồ)/.test(String(y.text || ""))))
            .forEach((x: any) => sl.items[x.slide - 1]?.delete());
          await ctx.sync();
        }).catch(() => undefined);
      }
      expect(!errs.length, `Hoàn tác lỗi: ${errs.join("; ")}`);
      expect(ov.slide_count === before, `Còn ${ov.slide_count} slide, ban đầu ${before}`);
      return `Hoàn tác ${log.length} thao tác, bài trình chiếu trở lại như cũ`;
    });
    // ----- Đợt 3: đọc mẫu, slide theo mẫu, kiểm bài theo mẫu, định dạng theo đoạn, bảng, bố cục, vùng chọn, ảnh slide -----
    {
      const mine: UndoEntry[] = [];
      const go = (name: string, args: any = {}) => tools.execute(name, args, { log: mine });
      const tag = "AIDN3" + Date.now().toString(36).slice(-4).toUpperCase();
      let base = 0;
      let theme: any = null;
      const countSlides = async () => (await go("ppt_get_overview")).slide_count as number;
      const shapeInfo = (slide: number, part: string) =>
        PowerPoint.run(async (ctx) => {
          const s = ctx.presentation.slides.getItemAt(slide - 1);
          s.layout.load("name");
          s.shapes.load("items/id,items/name,items/type");
          await ctx.sync();
          const out: any = { layout: s.layout.name, shapes: s.shapes.items.map((x) => x.name) };
          for (const sh of s.shapes.items) {
            if (!["TextBox", "GeometricShape", "Placeholder"].includes(String(sh.type))) continue;
            try {
              const tr = sh.textFrame.textRange;
              tr.load("text");
              sh.fill.load("foregroundColor,type");
              await ctx.sync();
              if (tr.text.includes(part)) {
                out.name = sh.name;
                out.text = tr.text;
                out.fill = sh.fill.foregroundColor;
                const lines = tr.text.split(/\r\n|\r|\n|\v/);
                let at = 0;
                out.paras = [];
                for (const line of lines) {
                  if (line.length) {
                    const r = tr.getSubstring(at, line.length);
                    r.font.load("bold,size,name");
                    await ctx.sync();
                    out.paras.push({ text: line, bold: r.font.bold, size: r.font.size, font: r.font.name });
                  }
                  at += line.length + 1;
                }
                break;
              }
            } catch {
              /* hình không có chữ */
            }
          }
          return out;
        });
      const cleanup = async () => {
        await PowerPoint.run(async (ctx) => {
          const sl = ctx.presentation.slides;
          sl.load("items/id");
          await ctx.sync();
          for (let k = sl.items.length - 1; k >= base && k >= 1; k--) sl.items[k].delete();
          await ctx.sync();
        }).catch(() => undefined);
      };

      await step("Đọc mẫu: màu chủ đề, phông, bố cục và ô giữ chỗ", async () => {
        base = await countSlides();
        theme = await go("ppt_get_theme");
        expect(/^#[0-9A-F]{6}$/i.test(theme.colors?.Accent1 || ""), `Không đọc được màu Accent1: ${JSON.stringify(theme.colors)}`);
        expect(theme.layouts?.length > 0 && theme.layouts.some((l: any) => l.placeholders.some((p: string) => /Title/.test(p))), "Không đọc được bố cục có ô tiêu đề");
        expect(theme.slide_size?.width > 100, "Không đọc được kích thước slide");
        return `${theme.brief}; Accent1 ${theme.colors.Accent1}; slide ${theme.slide_size.width}×${theme.slide_size.height}`;
      });

      await step("Slide mới dùng đúng bố cục và màu của mẫu", async () => {
        expect(theme, "Chưa đọc được mẫu");
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title") && l.placeholders.some((p: string) => /Body|Content|Object/.test(p))) || theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title"));
        const r = await go("ppt_add_slide", { layout: lay.name, title: `Theo mẫu ${tag}`, body: "Ý một\nÝ hai" });
        await go("ppt_add_shape", { slide: r.slide, shape_type: "rectangle", text: `Khối màu ${tag}`, left: 60, top: 300, width: 200, height: 60, fill_color: theme.colors.Accent1 });
        const info = await shapeInfo(r.slide, `Khối màu ${tag}`);
        expect(info.layout === lay.name, `Slide mới dùng bố cục “${info.layout}”, phải là “${lay.name}”`);
        expect(String(info.fill || "").toUpperCase() === String(theme.colors.Accent1).toUpperCase(), `Màu khối là ${info.fill}, phải là Accent1 ${theme.colors.Accent1}`);
        return `Slide ${r.slide}: bố cục “${info.layout}”, khối màu ${info.fill} = Accent1 của mẫu`;
      });

      await step("Kiểm bài theo mẫu: cài 5 lỗi (2 phông lạ, 1 slide thiếu tiêu đề, 2 chỗ chữ dưới 12 pt) → tìm đủ 5, không báo nhầm", async () => {
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title")) || theme.layouts[0];
        const a = (await go("ppt_add_slide", { layout: lay.name, title: `Slide A ${tag}`, body: "Nội dung đúng mẫu" })).slide;
        const b = (await go("ppt_add_slide", { layout: lay.name, title: `Slide B ${tag}`, body: "Nội dung đúng mẫu" })).slide;
        const c = (await go("ppt_add_slide", { layout: lay.name, title: "", body: "Slide này cố ý không có tiêu đề" })).slide;
        // Hộp chữ thử đặt sát đáy slide, không đè lên ô giữ chỗ nào của bố cục
        const bottom = theme.slide_size.height - 44;
        await go("ppt_add_textbox", { slide: a, text: `Phông lạ một ${tag}`, left: 40, top: bottom, width: 260, height: 36 });
        await go("ppt_format_text", { slide: a, shape: (await shapeInfo(a, `Phông lạ một ${tag}`)).name, font_name: "Comic Sans MS" });
        await go("ppt_add_textbox", { slide: a, text: `Chữ nhỏ một ${tag}`, left: 340, top: bottom, width: 260, height: 36, font_size: 9 });
        await go("ppt_add_textbox", { slide: b, text: `Phông lạ hai ${tag}`, left: 40, top: bottom, width: 260, height: 36 });
        await go("ppt_format_text", { slide: b, shape: (await shapeInfo(b, `Phông lạ hai ${tag}`)).name, font_name: "Courier New" });
        await go("ppt_add_textbox", { slide: b, text: `Chữ nhỏ hai ${tag}`, left: 340, top: bottom, width: 260, height: 36, font_size: 10 });
        await go("ppt_add_textbox", { slide: c, text: `Chữ đúng mẫu ${tag}`, left: 40, top: bottom, width: 260, height: 36, font_size: 18 });
        const r = await go("ppt_check_deck", { slides: [a, b, c] });
        const of = (kind: string) => r.issues.filter((x: any) => x.kind === kind);
        const fonts = of("font");
        const small = of("small_text");
        const noTitle = of("missing_title");
        expect(fonts.length === 2 && fonts.some((x: any) => x.slide === a && /Comic Sans MS/.test(x.message)) && fonts.some((x: any) => x.slide === b && /Courier New/.test(x.message)), `Phông lạ: ${JSON.stringify(fonts.map((x: any) => x.message))}`);
        expect(small.length === 2 && small.some((x: any) => x.slide === a && /cỡ 9 pt/.test(x.message)) && small.some((x: any) => x.slide === b && /cỡ 10 pt/.test(x.message)), `Chữ nhỏ: ${JSON.stringify(small.map((x: any) => x.message))}`);
        expect(noTitle.length === 1 && noTitle[0].slide === c, `Thiếu tiêu đề: ${JSON.stringify(noTitle.map((x: any) => x.message))}`);
        const other = r.issues.filter((x: any) => !["font", "small_text", "missing_title"].includes(x.kind));
        expect(other.length === 0, `Báo thêm điểm không cài: ${JSON.stringify(other.map((x: any) => x.message))}`);
        return `Tìm đủ 5/5: ${r.brief}; phông mẫu ${r.template_fonts.join(", ")}`;
      });

      await step("Định dạng chữ theo đoạn (có Hoàn tác); sửa bảng có sẵn; đổi bố cục; đọc vùng chọn", async () => {
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title")) || theme.layouts[0];
        const n = (await go("ppt_add_slide", { layout: lay.name, title: `Slide D ${tag}` })).slide;
        await go("ppt_add_textbox", { slide: n, text: `Đoạn một ${tag}\nĐoạn hai cần đậm\nĐoạn ba`, left: 60, top: 150, width: 300, height: 110, font_size: 16 });
        const box = (await shapeInfo(n, `Đoạn một ${tag}`)).name;
        const mark = mine.length;
        await go("ppt_format_text", { slide: n, shape: box, paragraph: 2, bold: true, font_size: 22 });
        let info = await shapeInfo(n, `Đoạn một ${tag}`);
        expect(info.paras[1].bold === true && info.paras[1].size === 22 && info.paras[0].bold !== true && info.paras[0].size === 16 && info.paras[2].size === 16, `Định dạng theo đoạn sai: ${JSON.stringify(info.paras)}`);
        await mine[mine.length - 1].undo();
        mine.length = mark;
        info = await shapeInfo(n, `Đoạn một ${tag}`);
        expect(info.paras[1].bold !== true && info.paras[1].size === 16, `Hoàn tác định dạng chưa về như cũ: ${JSON.stringify(info.paras[1])}`);
        await go("ppt_format_text", { slide: n, shape: info.name, find: "cần đậm", italic: true });
        let bad = "";
        try {
          await go("ppt_format_text", { slide: n, shape: info.name, paragraph: 9, bold: true });
        } catch (e: any) {
          bad = e.message;
        }
        expect(/không có đoạn số 9/.test(bad), "Đoạn không tồn tại mà không báo lỗi");
        // Bảng
        await go("ppt_add_table", { slide: n, values: [["Hạng mục", "Kế hoạch"], ["Móng", "100"]], left: 400, top: 150, width: 280, height: 80 });
        const ov = await go("ppt_get_overview");
        const tb = ov.slides.find((x: any) => x.slide === n).shapes.find((x: any) => x.type === "Table");
        expect(tb && tb.table && tb.table[1][1] === "100", `Tổng quan không kèm nội dung bảng: ${JSON.stringify(tb)}`);
        const t = await go("ppt_edit_table", { slide: n, shape: tb.name, cells: [{ row: 1, col: 1, text: "120" }], add_rows: [["Thân", "200"]] });
        expect(t.values[1][1] === "120" && t.rows === 3 && t.values[2][0] === "Thân" && t.values[2][1] === "200", `Sửa bảng sai: ${JSON.stringify(t.values)}`);
        let out = "";
        try {
          await go("ppt_edit_table", { slide: n, shape: tb.name, cells: [{ row: 9, col: 0, text: "x" }] });
        } catch (e: any) {
          out = e.message;
        }
        expect(/ngoài bảng/.test(out), "Ô ngoài bảng mà không báo lỗi");
        // Bố cục
        const other = theme.layouts.find((l: any) => l.name !== lay.name && l.placeholders.some((p: string) => /Title/.test(p))) || theme.layouts.find((l: any) => l.name !== lay.name);
        const L = await go("ppt_set_layout", { slide: n, layout: other.name });
        expect((await shapeInfo(n, tag)).layout === other.name, `Đổi bố cục không ăn: ${L.message}`);
        // Vùng chọn
        const sel = await go("ppt_get_selection");
        expect(Array.isArray(sel.slides) && Array.isArray(sel.shapes), "Không đọc được vùng chọn");
        return `Đoạn 2 đậm cỡ 22 rồi hoàn tác; bảng: ô (1,1) = 120, thêm hàng “Thân”; ${L.message}; vùng chọn: ${sel.brief}`;
      });

      await step("Ảnh slide cho AI tự nhìn; chữ cố ý tràn khung được đo ra, sửa xong thì hết", async () => {
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title")) || theme.layouts[0];
        const n = (await go("ppt_add_slide", { layout: lay.name, title: `Slide E ${tag}` })).slide;
        await PowerPoint.run(async (ctx) => {
          const s = ctx.presentation.slides.getItemAt(n - 1);
          const tb = s.shapes.addTextBox(`Tràn ${tag}: ` + "chữ rất dài cố ý để tràn khỏi khung. ".repeat(10), { left: 60, top: 200, width: 220, height: 50 });
          tb.name = `Tran ${tag}`;
          tb.textFrame.autoSizeSetting = "AutoSizeNone" as any;
          tb.textFrame.wordWrap = true;
          await ctx.sync();
        });
        const r = await go("ppt_get_slide_image", { slide: n });
        const img = r.__images?.[0];
        if (!img) throw new Skip(r.message || "Máy không chụp được ảnh slide");
        expect(img.mime === "image/png" && img.data.startsWith("iVBOR") && img.data.length > 2000, "Ảnh slide không phải PNG hợp lệ");
        expect(r.measured.some((m: string) => m.includes(`Tran ${tag}`) && /tràn khung/.test(m)), `Không đo ra chữ tràn: ${JSON.stringify(r.measured)}`);
        await go("ppt_format_text", { slide: n, shape: `Tran ${tag}`, font_size: 10 });
        await go("ppt_update_shape", { slide: n, shape: `Tran ${tag}`, width: 560, height: 110 });
        const again = await go("ppt_get_slide_image", { slide: n });
        expect(!again.measured.some((m: string) => m.includes(`Tran ${tag}`)), `Sửa rồi vẫn đo ra tràn: ${JSON.stringify(again.measured)}`);
        return `Ảnh PNG ${Math.round((img.data.length * 3) / 4 / 1024)} KB; đo ra tràn khung → giảm cỡ chữ và nới khung → hết tràn`;
      });

      await step("Đợt 7 · Thêm 3 slide trong một lệnh (Hoàn tác gỡ cả lô); đặt ảnh đính kèm vào slide đúng tỷ lệ, có Hoàn tác", async () => {
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title")) || theme.layouts[0];
        const before = await countSlides();
        const mark = mine.length;
        const r = await go("ppt_add_slides", {
          slides: [
            { layout: lay.name, title: `Slide F1 ${tag}`, body: "Ý một\nÝ hai" },
            { layout: lay.name, title: `Slide F2 ${tag}` },
            { layout: lay.name, title: `Slide F3 ${tag}` },
          ],
        });
        expect(r.ok && r.slides?.length === 3 && r.slides.map((x: any) => x.slide).join() === [before + 1, before + 2, before + 3].join(), `Thêm 3 slide: ${JSON.stringify(r)}`);
        expect(mine.length === mark + 1, `Cả lô phải là MỘT mục Hoàn tác, có ${mine.length - mark}`);
        expect((await countSlides()) === before + 3, "Số slide không tăng 3");

        // Ảnh thử 400 × 200 (tỷ lệ 2:1) dựng ngay trong khung, đăng ký như một ảnh người dùng đính kèm
        const { registerAttachedFile, unregisterAttachedFile } = await import("../core/attachedFiles");
        const cv = document.createElement("canvas");
        cv.width = 400;
        cv.height = 200;
        const g = cv.getContext("2d")!;
        g.fillStyle = "#1F6FB5";
        g.fillRect(0, 0, 400, 200);
        g.fillStyle = "#FFFFFF";
        g.fillRect(20, 20, 120, 60);
        const name = `anh-thu-${tag}.png`;
        registerAttachedFile(name, cv.toDataURL("image/png").split(",")[1]);
        try {
          const im = await go("ppt_insert_image", { slide: before + 1, image: `anh-thu-${tag}`, place: "right" });
          const box = await PowerPoint.run(async (ctx) => {
            const s = ctx.presentation.slides.getItemAt(before);
            s.shapes.load("items/name,items/left,items/top,items/width,items/height");
            await ctx.sync();
            const sh = s.shapes.items.find((x) => x.name === im.shape);
            return sh ? { left: sh.left, top: sh.top, width: sh.width, height: sh.height } : null;
          });
          expect(box, `Không thấy hình ảnh "${im.shape}" trên slide`);
          const size = theme.slide_size || { width: 960, height: 540 };
          const W = size.width || size.w || 960;
          const H = size.height || size.h || 540;
          expect(Math.abs(box!.width / box!.height / 2 - 1) < 0.01, `Sai tỷ lệ ảnh: ${JSON.stringify(box)}`);
          expect(box!.left >= W / 2 - 1 && box!.left + box!.width <= W + 1 && box!.top >= 0 && box!.top + box!.height <= H + 1, `Ảnh không nằm nửa phải trong slide (${W}×${H}): ${JSON.stringify(box)}`);
          await mine.pop()!.undo();
          const gone = await PowerPoint.run(async (ctx) => {
            const s = ctx.presentation.slides.getItemAt(before);
            s.shapes.load("items/name");
            await ctx.sync();
            return !s.shapes.items.some((x) => x.name === im.shape);
          });
          expect(gone, "Hoàn tác chưa gỡ ảnh");
          let missing = "";
          try {
            await go("ppt_insert_image", { slide: before + 1, image: "khong-co-anh-nay" });
          } catch (e: any) {
            missing = e.message;
          }
          expect(/Không thấy ảnh đính kèm/.test(missing), `Ảnh không có mà không báo rõ: ${missing}`);
          await mine.pop()!.undo();
          expect((await countSlides()) === before, "Hoàn tác cả lô chưa gỡ đủ 3 slide");
          return `3 slide (${r.slides.map((x: any) => x.slide).join(", ")}) trong một lệnh, một mục Hoàn tác; ảnh 400×200 đặt ${Math.round(box!.width)}×${Math.round(box!.height)} tại (${Math.round(box!.left)}, ${Math.round(box!.top)}), đúng tỷ lệ, nửa phải; Hoàn tác gỡ ảnh rồi gỡ cả 3 slide`;
        } finally {
          unregisterAttachedFile(name);
        }
      });

      await step("Đợt 7 · Chữ có “ư” trong tiêu đề, nội dung và tiêu đề slide biểu đồ vẫn theo phông của mẫu (không thành Times New Roman)", async () => {
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title") && l.placeholders.some((p: string) => /Body|Content/.test(p))) || theme.layouts[0];
        const before = await countSlides();
        /** Phông của từng dòng chữ trong các ô giữ chỗ có chữ của một slide */
        const fontsOf = (n: number) =>
          PowerPoint.run(async (ctx) => {
            const s = ctx.presentation.slides.getItemAt(n - 1);
            s.shapes.load("items/id,items/name,items/type");
            await ctx.sync();
            const out: { shape: string; fonts: string[] }[] = [];
            for (const sh of s.shapes.items.filter((x) => String(x.type) === "Placeholder")) {
              const tr = sh.textFrame.textRange;
              tr.load("text");
              await ctx.sync();
              const text = String(tr.text || "");
              if (!text.trim()) continue;
              const rs: PowerPoint.TextRange[] = [];
              let at = 0;
              for (const line of text.split(/\r\n|\r|\n|\v/)) {
                if (line.length) {
                  const r = tr.getSubstring(at, line.length);
                  r.font.load("name");
                  rs.push(r);
                }
                at += line.length + 1;
              }
              await ctx.sync();
              out.push({ shape: sh.name, fonts: rs.map((r) => String(r.font.name)) });
            }
            return out;
          });
        const r = await go("ppt_add_slides", { slides: [{ layout: lay.name, title: `Nhận xét xu hướng ${tag}`, body: "Tăng trưởng tương đối đều\nKhông dấu" }] });
        const n = r.slides[0].slide;
        const f1 = await fontsOf(n);
        const all1 = f1.flatMap((x) => x.fonts);
        expect(all1.length >= 3, `Không đọc được phông của slide thử: ${JSON.stringify(f1)}`);
        expect(!all1.some((f) => /times new roman/i.test(f)), `Chữ có “ư” bị đổi sang Times New Roman: ${JSON.stringify(f1)}`);
        // Sửa chữ một ô có sẵn bằng ppt_set_shape_text
        const titleName = f1[0].shape;
        await go("ppt_set_shape_text", { slide: n, shape: titleName, text: `Xu hướng thị trường ${tag}` });
        const f2 = await fontsOf(n);
        expect(!f2.flatMap((x) => x.fonts).some((f) => /times new roman/i.test(f)), `Sau ppt_set_shape_text chữ thành Times New Roman: ${JSON.stringify(f2)}`);
        // Tiêu đề của slide biểu đồ
        const c = await go("ppt_add_chart", { title: `Doanh thu theo quý ${tag}`, chart_type: "column", categories: ["Q1", "Q2"], series: [{ name: "Doanh thu", values: [1, 2] }] });
        const f3 = await fontsOf(c.slide);
        expect(f3.length >= 1, `Slide biểu đồ không có ô tiêu đề có chữ: ${JSON.stringify(f3)}`);
        expect(!f3.flatMap((x) => x.fonts).some((f) => /times new roman/i.test(f)), `Tiêu đề slide biểu đồ mang Times New Roman: ${JSON.stringify(f3)}`);
        expect((await countSlides()) === before + 2, "Số slide thử không đúng");
        return `Tiêu đề và nội dung có “ư”: ${[...new Set(all1)].join(", ")}; sau khi sửa chữ: ${[...new Set(f2.flatMap((x) => x.fonts))].join(", ")}; tiêu đề slide biểu đồ: ${[...new Set(f3.flatMap((x) => x.fonts))].join(", ")}`;
      });

      await step("Chữ có “ư” trong ô chữ, hình, sơ đồ quy trình và bảng vẫn theo phông của mẫu", async () => {
        const lay = theme.layouts.find((l: any) => l.placeholders.some((p: string) => p === "Title")) || theme.layouts[0];
        const r = await go("ppt_add_slides", { slides: [{ layout: lay.name, title: `Ô chữ ${tag}` }] });
        const n = r.slides[0].slide;
        // Ô chữ không dấu làm mốc: phông mà chữ có “ư” phải giống
        await go("ppt_add_textbox", { slide: n, text: `Tong quan ${tag}`, left: 40, top: 120, width: 300, height: 36 });
        await go("ppt_add_textbox", { slide: n, text: `Nhận xét xu hướng ${tag}`, left: 360, top: 120, width: 300, height: 36 });
        await go("ppt_add_shape", { slide: n, shape_type: "rounded_rectangle", text: `Khối tăng trưởng ${tag}`, left: 40, top: 170, width: 260, height: 50 });
        await go("ppt_add_process_flow", { slide: n, steps: ["Chuẩn bị hồ sơ", "Thẩm định dự án", "Phê duyệt chủ trương"], style: "chevron", top: 240, height: 60 });
        await go("ppt_add_process_flow", { slide: n, steps: ["Khởi công", "Hoàn thiện phần thô", "Bàn giao nhà ở"], style: "timeline", top: 310, height: 80 });
        await go("ppt_add_table", { slide: n, values: [["Hạng mục", "Tăng trưởng"], ["Quý 1", "Tư vấn"]], left: 420, top: 170 });
        const seen = await PowerPoint.run(async (ctx) => {
          const s = ctx.presentation.slides.getItemAt(n - 1);
          s.shapes.load("items/id,items/name,items/type");
          await ctx.sync();
          let list: PowerPoint.Shape[] = [];
          for (const sh of s.shapes.items) {
            if (String(sh.type) === "Group") {
              sh.group.shapes.load("items/id,items/name,items/type");
              await ctx.sync();
              list = list.concat(sh.group.shapes.items);
            } else list.push(sh);
          }
          const out: { text: string; font: string }[] = [];
          for (const sh of list.filter((x) => ["TextBox", "GeometricShape"].includes(String(x.type)))) {
            const tr = sh.textFrame.textRange;
            tr.load("text");
            await ctx.sync();
            const text = tr.text || "";
            if (!text.trim()) continue;
            // Phông theo từng ký tự (phông cả khung có thể che chỗ lệch), gom thành các đoạn liền cùng phông
            const chars = [...text].map((_, i) => {
              const ch = tr.getSubstring(i, 1);
              ch.font.load("name");
              return ch;
            });
            await ctx.sync();
            const runs: { text: string; font: string }[] = [];
            chars.forEach((ch, i) => {
              const last = runs[runs.length - 1];
              if (last && last.font === String(ch.font.name)) last.text += text[i];
              else runs.push({ text: text[i], font: String(ch.font.name) });
            });
            out.push(...runs);
          }
          for (const sh of list.filter((x) => String(x.type) === "Table")) {
            const t = sh.getTable();
            const cells = [0, 1].flatMap((row) => [0, 1].map((col) => t.getCellOrNullObject(row, col)));
            cells.forEach((c) => {
              c.load("text");
              c.font.load("name");
            });
            await ctx.sync();
            cells.forEach((c) => out.push({ text: c.text, font: String(c.font.name) }));
          }
          return out;
        });
        const ref = seen.find((x) => x.text.startsWith("Tong quan"))?.font;
        expect(!!ref, `Không đọc được phông ô chữ mốc: ${JSON.stringify(seen)}`);
        const off = seen.filter((x) => x.font !== ref);
        expect(!off.length, `Chữ lệch phông mẫu (${ref}): ${JSON.stringify(off)}`);
        expect(seen.length >= 10, `Đọc được quá ít đoạn chữ: ${JSON.stringify(seen)}`);
        return `${seen.length} đoạn chữ (ô chữ, hình, 2 sơ đồ, bảng) đều ${ref}`;
      });

      await step("Dọn các slide thử của Đợt 3", async () => {
        await cleanup();
        const n = await countSlides();
        expect(n === base, `Còn ${n} slide, ban đầu ${base}`);
        return `Bài trở lại ${n} slide`;
      });
    }

  } else {
    steps.push({ name: "Ứng dụng", ok: null, detail: "Chỉ chạy được trong Excel, Word, PowerPoint" });
  }
  if (APPS.includes(tools.host as AppName)) {
    await step("Ghi chéo ứng dụng: khung đích chưa mở → báo rõ, không treo; nhận lệnh ghi hộ → ghi đúng, có Hoàn tác; lệnh sai bị bỏ", async () => {
      const self = tools.host as AppName;
      const other = APPS.find((a) => a !== self)!;
      // Kho giả trong bộ nhớ: bước này không đụng kho dùng chung thật, không gửi gì sang ứng dụng khác
      const mem: Record<string, string> = {};
      const kv = { getItem: (k: string) => (k in mem ? mem[k] : null), setItem: (k: string, v: string) => void (mem[k] = v), removeItem: (k: string) => void delete mem[k] };
      const code = "AIDNCHEO" + Date.now().toString(36).slice(-5).toUpperCase();
      const sample = (app: AppName): any =>
        app === "Word"
          ? { text: `Đoạn ghi chéo ${code}\nDòng hai ${code}` }
          : app === "Excel"
            ? { sheet_name: `KT ${code}`.slice(0, 28), values: [["Hạng mục", "Giá trị"], ["Móng", "100"], ["Công thức lạ", "=1+1"]] }
            : { title: `Slide ghi chéo ${code}`, body: "Ý một\nÝ hai" };
      // 1. Khung đích chưa mở: trả lời ngay, nêu cách mở
      let t0 = Date.now();
      const r0 = await sendCommand(kv, self, other, buildPayload(other, sample(other)), { timeoutMs: 1500 });
      expect(!r0.ok && /chưa mở/.test(r0.error || ""), `Đích chưa mở mà không báo rõ: ${JSON.stringify(r0)}`);
      expect(Date.now() - t0 < 800, "Đích chưa mở mà phải chờ");
      // 2. Khung đích có nhịp sống nhưng không lấy lệnh: hết hạn thì rút lệnh, không treo
      beat(kv, { host: other, accept: true, title: "tệp thử" });
      t0 = Date.now();
      const r1 = await sendCommand(kv, self, other, buildPayload(other, sample(other)), { timeoutMs: 1200, pollMs: 100 });
      expect(!r1.ok && /rút lại/.test(r1.error || ""), `Hết hạn mà không rút lệnh: ${JSON.stringify(r1)}`);
      expect(kv.getItem(CMD_PREFIX + other) === null && Date.now() - t0 < 3000, "Lệnh hết hạn còn nằm trong hộp thư");
      // 3. Lệnh sai ứng dụng hoặc sai hình dạng bị bỏ
      kv.setItem(CMD_PREFIX + self, JSON.stringify({ id: "saihinhdang1", from: other, to: self, at: Date.now(), payload: buildPayload(other, sample(other)) }));
      expect(takeCommand(kv, self, { accept: true, seen: new Set() }) === null && kv.getItem(CMD_PREFIX + self) === null, "Lệnh dành cho ứng dụng khác không bị bỏ");
      expect(cellAsText("=1+1") === "'=1+1" && cellAsText("-5") === "-5" && cellAsText("@x") === "'@x", "Ô bắt đầu bằng = hoặc @ phải được ghi thành chữ");
      // 4. Nhận một lệnh gửi tới chính ứng dụng này và chạy bằng công cụ ghi thật
      const count = async (): Promise<{ n: number; has: boolean; extra?: string }> => {
        if (self === "Excel")
          return Excel.run(async (ctx) => {
            const ws = ctx.workbook.worksheets;
            ws.load("items/name");
            await ctx.sync();
            const sh = ws.items.find((w) => w.name.includes(code));
            if (!sh) return { n: ws.items.length, has: false };
            const r = sh.getRange("A1:B3");
            r.load("values,formulas");
            await ctx.sync();
            return { n: ws.items.length, has: true, extra: JSON.stringify(r.formulas) };
          });
        if (self === "Word")
          return Word.run(async (ctx) => {
            const ps = ctx.document.body.paragraphs;
            ps.load("items");
            await ctx.sync();
            const ts = ps.items.map((p) => p.getReviewedText("Current"));
            await ctx.sync();
            return { n: ts.length, has: ts.some((t) => t.value.includes(`Đoạn ghi chéo ${code}`)) && ts.some((t) => t.value.includes(`Dòng hai ${code}`)) };
          });
        const ov = await run("ppt_get_overview");
        return { n: ov.slide_count, has: JSON.stringify(ov.slides).includes(`Slide ghi chéo ${code}`) };
      };
      const before = await count();
      beat(kv, { host: self, accept: true });
      const pending = sendCommand(kv, other, self, buildPayload(self, sample(self)), { timeoutMs: 60000, pollMs: 100 });
      await new Promise((r) => setTimeout(r, 200));
      const seen = new Set<string>();
      const cmd = takeCommand(kv, self, { accept: true, seen });
      expect(!!cmd, "Khung nhận không lấy được lệnh hợp lệ");
      const mine: UndoEntry[] = [];
      let message = "";
      try {
        message = await executeIncoming(tools, cmd!.payload, mine);
        reply(kv, cmd!.id, { ok: true, message });
        const res = await pending;
        expect(res.ok && !!res.message, `Khung gửi không nhận được kết quả: ${JSON.stringify(res)}`);
        const mid = await count();
        expect(mid.has, "Nội dung ghi chéo không có trong tài liệu");
        if (self === "Excel") expect(/'?=1\+1/.test(mid.extra || "") && !/"2"|,2\]/.test(mid.extra || ""), `Công thức từ ứng dụng khác phải thành chữ: ${mid.extra}`);
        expect(mine.length > 0, "Ghi chéo không để lại mục Hoàn tác");
        // Lệnh đã chạy không chạy lại
        kv.setItem(CMD_PREFIX + self, JSON.stringify(cmd));
        expect(takeCommand(kv, self, { accept: true, seen }) === null, "Lệnh đã chạy bị chạy lại");
      } finally {
        for (const u of mine.slice().reverse()) await u.undo();
      }
      const after = await count();
      expect(!after.has && after.n === before.n, `Hoàn tác ghi chéo chưa sạch: trước ${before.n}, sau ${after.n}, còn nội dung: ${after.has}`);
      return `${other} chưa mở → “${(r0.error || "").slice(0, 60)}…”; hết hạn → rút lệnh; nhận lệnh → ${message.slice(0, 70)}; hoàn tác sạch`;
    });
  }
  // Dùng chung mọi ứng dụng: đọc tệp PDF đính kèm theo trang (không đụng tài liệu đang mở)
  await step("Đọc tệp PDF đính kèm theo trang: trang có chữ ra chữ, trang quét ra ảnh", async () => {
    const { registerAttachedFile, unregisterAttachedFile } = await import("../core/attachedFiles");
    const { createCommonTools } = await import("./common");
    const { testPdf } = await import("./selftestPdf");
    const name = `aidn-tu-kiem-${Date.now().toString(36)}.pdf`;
    registerAttachedFile(name, btoa(testPdf()));
    try {
      const t = createCommonTools({ host: tools.host, skills: [], shareContext: false });
      const r = await t.execute("read_attachment_pages", { file: name, pages: "1-2" });
      expect(r.total_pages === 2, `Đếm ra ${r.total_pages} trang, cần 2`);
      const p1 = r.pages.find((x: any) => x.page === 1);
      const p2 = r.pages.find((x: any) => x.page === 2);
      expect(p1?.kind === "text" && /MOC-TU-KIEM/.test(p1.text), `Trang 1 không ra chữ: ${JSON.stringify(p1).slice(0, 160)}`);
      expect(p2?.kind === "scan", `Trang 2 (không có lớp chữ) không được coi là trang quét: ${JSON.stringify(p2).slice(0, 160)}`);
      const img = r.__images?.[0];
      expect(img?.mime === "image/jpeg" && img.data.startsWith("/9j/") && img.data.length > 1000, "Ảnh trang quét không phải JPEG hợp lệ");
      expect(typeof r.rules === "string" && r.rules.includes("[?]"), "Thiếu lời dặn chép nguyên văn");
      return `2 trang: trang 1 ${p1.text.length} ký tự, trang 2 ảnh JPEG ${Math.round((img.data.length * 3) / 4 / 1024)} KB`;
    } finally {
      unregisterAttachedFile(name);
    }
  });
  onProgress([...steps]);
  return steps;
}
