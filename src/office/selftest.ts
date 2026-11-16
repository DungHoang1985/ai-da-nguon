/* global Excel, Word, PowerPoint, Office */
/** Tự kiểm tra: chạy thật từng công cụ trên tài liệu đang mở rồi dọn sạch. Không cần API key. */
import { HostToolSet, UndoEntry } from "../core/agent/types";
import { runWordSelfTest } from "./selftestWord";
import { APPS, AppName, CMD_PREFIX, beat, buildPayload, reply, sendCommand, takeCommand } from "../core/crossapp";
import { cellAsText, executeIncoming } from "./crossExec";

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
  const run = (name: string, args: any = {}) => tools.execute(name, args, { log });
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
      expect(n2 === nLayouts, `Mẫu bị thêm bố cục mới (${nLayouts} → ${n2}), slide dùng bố cục "${sl?.layout}"`);
      return `${r.message} (${parts}; slide dùng bố cục "${sl?.layout}", mẫu không bị thêm bố cục)`;
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
  onProgress([...steps]);
  return steps;
}
