/* global Excel, Word, PowerPoint, Office */
/** Tự kiểm tra: chạy thật từng công cụ trên tài liệu đang mở rồi dọn sạch. Không cần API key. */
import { HostToolSet, UndoEntry } from "../core/agent/types";

export interface TestStep {
  name: string;
  /** true = đạt, false = lỗi, null = bỏ qua (máy không hỗ trợ) */
  ok: boolean | null;
  detail: string;
}

class Skip extends Error {}

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
      else steps.push({ name, ok: false, detail: e?.message || String(e) });
    }
    onProgress([...steps]);
  };
  const expect = (cond: any, msg: string) => {
    if (!cond) throw new Error(msg);
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
    await step("Chọn vùng kết quả", async () => (await run("excel_select_range", { sheet, address: "A1:B4" })).message);
    await step("Hoàn tác từng thao tác", async () => {
      const first = log[0];
      for (const u of log.slice(1).reverse()) await u.undo();
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
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheet);
      await ctx.sync();
      if (!ws.isNullObject) {
        ws.delete();
        await ctx.sync();
      }
    }).catch(() => undefined);
  } else if (tools.host === "Word") {
    const marker = "AIDN-KT-" + Date.now().toString(36).slice(-5);
    let startCount = 0;
    await step("Xem tổng quan văn bản", async () => {
      const ov = await run("word_get_overview");
      startCount = ov.paragraphs;
      return `${ov.paragraphs} đoạn, Track Changes: ${ov.track_changes ?? "không rõ"}`;
    });
    await step("Chèn tiêu đề + đoạn + gạch đầu dòng", async () => {
      await run("word_insert_paragraphs", {
        position: "end",
        text: `## Mục kiểm tra ${marker}\nHợp đồng số 01/2026 ký ngày 01/01/2026.\n- Ý thứ nhất`,
      });
      // Văn bản trống: Word đặt nội dung mới trước đoạn trống ban đầu → đọc từ đầu
      const r = await run("word_read_paragraphs", { start: 0, end: startCount + 10 });
      const texts = r.paragraphs.map((p: any) => p.text).join(" | ");
      expect(texts.includes(marker) && texts.includes("Hợp đồng số 01/2026"), `Đọc lại: ${texts}`);
      const h = r.paragraphs.find((p: any) => p.text.includes(marker));
      expect(/2|heading|tiêu đề/i.test(String(h?.style)), `Kiểu tiêu đề: ${h?.style}`);
      const isBullet = await Word.run(async (ctx) => {
        const hits = ctx.document.body.search("Ý thứ nhất", { matchCase: true });
        hits.load("items");
        await ctx.sync();
        const p = hits.items[0].paragraphs.getFirst();
        p.load("isListItem");
        await ctx.sync();
        return p.isListItem;
      });
      expect(isBullet, "Dòng '- Ý thứ nhất' chưa thành gạch đầu dòng");
      return `Đã chèn; tiêu đề kiểu "${h?.style}", dòng gạch đầu dòng là danh sách thật`;
    });
    await step("Ghi bằng Theo dõi thay đổi", async () => {
      if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) throw new Skip("Word chưa hỗ trợ đếm Track Changes");
      const n = await Word.run(async (ctx) => {
        const tc = ctx.document.body.getTrackedChanges();
        tc.load("items");
        await ctx.sync();
        return tc.items.length;
      });
      expect(n > 0, "Không thấy sửa đổi được theo dõi");
      return `${n} sửa đổi đang chờ duyệt`;
    });
    await step("Tìm và thay chữ", async () => {
      const r = await run("word_replace_text", { find: "01/2026 ký", replace: "02/2026 ký" });
      expect(r.replaced === 1, `Thay ${r.replaced} chỗ`);
      return "01/2026 → 02/2026";
    });
    await step("Viết lại một đoạn (chỉ sửa phần khác)", async () => {
      const r = await run("word_read_paragraphs", { start: 0, end: startCount + 10 });
      const p = r.paragraphs.find((x: any) => x.text.includes("Hợp đồng số"));
      expect(p, "Không thấy đoạn hợp đồng");
      const res = await run("word_set_paragraph_text", {
        index: p.i,
        text: "Hợp đồng số 02/2026 ký ngày 01/01/2026, đã có hiệu lực.",
      });
      const r2 = await run("word_read_paragraphs", { start: p.i, end: p.i + 1 });
      expect(r2.paragraphs[0].text.includes("đã có hiệu lực"), `Đoạn sau sửa: ${r2.paragraphs[0].text}`);
      return res.message;
    });
    await step("Định dạng chữ", async () => (await run("word_format_text", { find: marker, bold: true, font_color: "#C00000" })).message);
    await step("Thêm nhận xét", async () => {
      if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) throw new Skip("Word chưa hỗ trợ nhận xét");
      return (await run("word_add_comment", { find: marker, text: "Nhận xét thử của AI Đa Nguồn" })).message;
    });
    await step("Chèn bảng", async () => {
      const r = await run("word_read_paragraphs", { start: 0, end: startCount + 10 });
      const last = r.paragraphs.find((x: any) => x.text.includes("Ý thứ nhất")) || r.paragraphs[r.paragraphs.length - 1];
      return (
        await run("word_insert_table", {
          position: "after_paragraph",
          paragraph_index: last.i,
          values: [["Mục", "Giá trị"], ["A", "1"]],
        })
      ).message;
    });
    await step("Dọn sạch nội dung thử", async () => {
      await Word.run(async (ctx) => {
        const hits = ctx.document.body.search(marker, { matchCase: true });
        hits.load("items");
        await ctx.sync();
        expect(hits.items.length, "Không tìm thấy đoạn thử để dọn");
        const startPara = hits.items[0].paragraphs.getFirst();
        const rng = startPara.getRange("Whole").expandTo(ctx.document.body.getRange("End"));
        if (Office.context.requirements.isSetSupported("WordApi", "1.4")) {
          const cs = rng.getComments();
          cs.load("items");
          await ctx.sync();
          cs.items.forEach((c) => c.delete());
          ctx.document.load("changeTrackingMode");
          await ctx.sync();
          const mode = ctx.document.changeTrackingMode;
          ctx.document.changeTrackingMode = "Off" as any;
          rng.delete();
          await ctx.sync();
          ctx.document.changeTrackingMode = mode;
        } else {
          rng.delete();
        }
        await ctx.sync();
        const again = ctx.document.body.search(marker, { matchCase: true });
        again.load("items");
        await ctx.sync();
        expect(again.items.length === 0, "Vẫn còn nội dung thử");
      });
      return "Đã xóa toàn bộ nội dung thử";
    });
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
    await step("Chuyển tới slide", async () => (await run("ppt_go_to_slide", { slide: slideNo })).message);
    await step("Hoàn tác toàn bộ", async () => {
      for (const u of log.slice().reverse()) await u.undo();
      const ov = await run("ppt_get_overview");
      if (ov.slide_count !== before) {
        // dọn slide thử còn sót
        await PowerPoint.run(async (ctx) => {
          const sl = ctx.presentation.slides;
          sl.load("items/id");
          await ctx.sync();
          if (slideNo > 0 && sl.items[slideNo - 1]) sl.items[slideNo - 1].delete();
          await ctx.sync();
        }).catch(() => undefined);
      }
      expect(ov.slide_count === before, `Còn ${ov.slide_count} slide, ban đầu ${before}`);
      return `Hoàn tác ${log.length} thao tác, bài trình chiếu trở lại như cũ`;
    });
  } else {
    steps.push({ name: "Ứng dụng", ok: null, detail: "Chỉ chạy được trong Excel, Word, PowerPoint" });
  }
  onProgress([...steps]);
  return steps;
}
