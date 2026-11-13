/* global Excel */
/** Công cụ để AI thao tác trực tiếp trên sổ tính Excel. Mỗi thao tác ghi đều lưu ảnh chụp để Hoàn tác. */
import { HostToolSet, ToolSpec, UndoEntry } from "../core/agent/types";
import { cellAddress, lettersToCol, normalize2D, sanitizeSheetName, splitAddress, topLeft, truncate } from "./util";
import { getRange, restoreFrom, snapshot } from "./excelCore";
import { EXCEL_ADV_SPECS, excelRisk, executeExcelAdvanced } from "./excelAdvanced";

export { getRange, restoreFrom, settableProps, snapshot } from "./excelCore";
export type { Snapshot } from "./excelCore";

const MAX_READ_CELLS = 2500;

const addressProp = {
  type: "string",
  description: "A1-style range, optionally with sheet name, e.g. 'Sheet1!A1:D20' or 'B2'.",
};
const sheetProp = { type: "string", description: "Sheet name (optional; defaults to the active sheet)." };

export const EXCEL_SPECS: ToolSpec[] = [
  {
    name: "excel_get_overview",
    description:
      "Get a workbook overview: sheet names, used range of each sheet, active sheet, current selection and the first rows of every sheet.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "excel_read_range",
    description:
      "Read cell values (and formulas) of a range. Use the used range from the overview to read a whole table. Large ranges are truncated.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        include_formulas: { type: "boolean", description: "Also return formulas (default true)." },
      },
      required: ["address"],
    },
  },
  {
    name: "excel_write_range",
    description:
      "Write a 2D block of values starting at a cell. Strings starting with '=' are formulas (English function names, comma separators, e.g. '=SUM(B2:B10)'). Numbers as plain digits. Prefix with an apostrophe to force text (e.g. \"'0912\"). Write whole tables in one call.",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string", description: "Top-left cell where writing starts, e.g. 'Sheet1!B2'." },
        sheet: sheetProp,
        values: {
          type: "array",
          description: "Rows of cells, e.g. [[\"Month\",\"Revenue\"],[\"1\",\"100\"],[\"Total\",\"=SUM(B2:B2)\"]].",
          items: { type: "array", items: { type: "string" } },
        },
      },
      required: ["address", "values"],
    },
    write: true,
  },
  {
    name: "excel_format_range",
    description: "Format a range: font, fill colour, number format, alignment, wrap, borders, column width / autofit.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_color: { type: "string", description: "Hex colour like #1F4E79." },
        font_size: { type: "number" },
        font_name: { type: "string" },
        fill_color: { type: "string", description: "Hex colour like #DDEBF7, or 'none' to clear." },
        number_format: { type: "string", description: "Excel number format, e.g. '#,##0', '0.0%', 'dd/mm/yyyy'." },
        horizontal_alignment: { type: "string", enum: ["Left", "Center", "Right"] },
        wrap_text: { type: "boolean" },
        borders: { type: "string", enum: ["all", "outside", "none"] },
        autofit_columns: { type: "boolean" },
        column_width: { type: "number", description: "Column width in points." },
      },
      required: ["address"],
    },
    write: true,
  },
  {
    name: "excel_add_sheet",
    description: "Add a new worksheet (and activate it).",
    parameters: {
      type: "object",
      properties: { name: { type: "string" }, activate: { type: "boolean", description: "Default true." } },
      required: ["name"],
    },
    write: true,
  },
  {
    name: "excel_create_table",
    description: "Convert a range into a formatted Excel table (with filter buttons).",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        has_headers: { type: "boolean", description: "First row contains headers (default true)." },
        name: { type: "string", description: "Table name without spaces (optional)." },
        style: { type: "string", description: "Table style, e.g. TableStyleMedium2 (optional)." },
      },
      required: ["address"],
    },
    write: true,
  },
  {
    name: "excel_create_chart",
    description: "Create a chart from a data range (first row/column used as labels). The chart is placed to the right of the data.",
    parameters: {
      type: "object",
      properties: {
        source_address: addressProp,
        sheet: sheetProp,
        chart_type: { type: "string", enum: ["column", "bar", "line", "pie", "area", "scatter", "doughnut"] },
        title: { type: "string" },
        series_by: { type: "string", enum: ["auto", "rows", "columns"] },
      },
      required: ["source_address", "chart_type"],
    },
    write: true,
  },
  {
    name: "excel_sort_range",
    description: "Sort the rows of a range by one column.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        column: { type: "string", description: "Column letter of the sort key, e.g. 'C'." },
        ascending: { type: "boolean", description: "Default true." },
        has_headers: { type: "boolean", description: "Keep the first row as header (default true)." },
      },
      required: ["address", "column"],
    },
    write: true,
  },
  {
    name: "excel_clear_range",
    description: "Clear contents and/or formats of a range. Only when the user asked to delete or clean data.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        what: { type: "string", enum: ["contents", "formats", "all"], description: "Default contents." },
      },
      required: ["address"],
    },
    write: true,
  },
  {
    name: "excel_select_range",
    description: "Select a range (and its sheet) so the user sees it. Use at the end to show the result.",
    parameters: { type: "object", properties: { address: addressProp, sheet: sheetProp }, required: ["address"] },
  },
];

function isDateFormat(nf: any): boolean {
  const s = String(nf || "");
  if (!s || s === "General") return false;
  return /[dmyhs]/i.test(s.replace(/\[[^\]]*\]|"[^"]*"/g, ""));
}

const ERROR_RE = /^#(NAME\?|VALUE!|REF!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|FIELD!|BUSY!)/;

async function overview() {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const sheets = wb.worksheets;
    sheets.load("items/name");
    const active = sheets.getActiveWorksheet();
    active.load("name");
    await ctx.sync();
    let selection: string | null = null;
    try {
      const sel = wb.getSelectedRange();
      sel.load("address");
      await ctx.sync();
      selection = sel.address;
    } catch {
      selection = null;
    }
    const list = sheets.items.slice(0, 30);
    const used = list.map((ws) => {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load("address,rowCount,columnCount");
      return u;
    });
    await ctx.sync();
    const previews = used.map((u) => {
      if (u.isNullObject) return null;
      const r = u.getCell(0, 0).getResizedRange(Math.min(u.rowCount, 4) - 1, Math.min(u.columnCount, 12) - 1);
      r.load("values");
      return r;
    });
    await ctx.sync();
    return {
      activeSheet: active.name,
      selection,
      sheets: list.map((ws, i) => ({
        name: ws.name,
        usedRange: used[i].isNullObject ? null : used[i].address,
        rows: used[i].isNullObject ? 0 : used[i].rowCount,
        columns: used[i].isNullObject ? 0 : used[i].columnCount,
        firstRows: previews[i] ? previews[i]!.values.map((row) => row.map((v) => truncate(v, 40))) : [],
      })),
      moreSheets: sheets.items.length > list.length ? sheets.items.length - list.length : 0,
    };
  });
}

async function execute(name: string, a: any, { log }: { log: UndoEntry[] }): Promise<any> {
  switch (name) {
    case "excel_get_overview":
      return overview();

    case "excel_read_range":
      return Excel.run(async (ctx) => {
        const { range } = getRange(ctx, a.address, a.sheet);
        range.load("address,rowCount,columnCount");
        await ctx.sync();
        const cols = Math.min(range.columnCount, 60);
        const maxRows = Math.max(1, Math.floor(MAX_READ_CELLS / cols));
        let r = range;
        let truncated = false;
        if (range.rowCount > maxRows || range.columnCount > cols) {
          r = range.getCell(0, 0).getResizedRange(Math.min(range.rowCount, maxRows) - 1, cols - 1);
          truncated = true;
        }
        r.load("address,values,formulas,text,numberFormat");
        await ctx.sync();
        const tl = topLeft(r.address);
        const values = r.values.map((row, i) =>
          row.map((v, j) =>
            typeof v === "number" && isDateFormat(r.numberFormat[i][j]) ? r.text[i][j] : typeof v === "string" ? truncate(v, 300) : v
          )
        );
        const formulas: Record<string, string> = {};
        if (a.include_formulas !== false) {
          r.formulas.forEach((row, i) =>
            row.forEach((f, j) => {
              if (typeof f === "string" && f.startsWith("=")) formulas[cellAddress(tl.col + j, tl.row + i)] = f;
            })
          );
        }
        return {
          ok: true,
          address: r.address,
          values,
          ...(Object.keys(formulas).length ? { formulas } : {}),
          ...(truncated
            ? { note: `Vùng ${range.address} có ${range.rowCount}×${range.columnCount} ô; chỉ trả về ${r.address}. Đọc tiếp phần còn lại nếu cần.` }
            : {}),
        };
      });

    case "excel_write_range":
      return Excel.run(async (ctx) => {
        const vals = normalize2D(a.values);
        if (!vals.length || !vals[0].length) throw new Error("values rỗng.");
        const { range } = getRange(ctx, a.address, a.sheet);
        const target = range.getCell(0, 0).getResizedRange(vals.length - 1, vals[0].length - 1);
        const snap = await snapshot(ctx, target);
        target.formulas = vals.map((row) => row.map((v) => (typeof v === "number" || typeof v === "boolean" ? v : String(v))));
        target.load("address,values");
        await ctx.sync();
        const undo = restoreFrom(snap, `Ghi ${target.address}`);
        if (undo) log.push(undo);
        const tl = topLeft(target.address);
        const errors: string[] = [];
        target.values.forEach((row, i) =>
          row.forEach((v, j) => {
            if (typeof v === "string" && ERROR_RE.test(v)) errors.push(`${cellAddress(tl.col + j, tl.row + i)}: ${v}`);
          })
        );
        return {
          ok: true,
          message: `Đã ghi ${vals.length * vals[0].length} ô vào ${target.address}`,
          address: target.address,
          result_preview: target.values.slice(0, 12).map((row) => row.slice(0, 10)),
          ...(errors.length ? { formula_errors: errors.slice(0, 20) } : {}),
          ...(snap ? {} : { note: "Vùng lớn nên không lưu được bản để Hoàn tác." }),
        };
      });

    case "excel_format_range":
      return Excel.run(async (ctx) => {
        const { range } = getRange(ctx, a.address, a.sheet);
        range.load("address,rowCount,columnCount");
        await ctx.sync();
        const snap = await snapshot(ctx, range, !!(a.autofit_columns || a.column_width));
        const f = range.format;
        if (a.bold !== undefined) f.font.bold = !!a.bold;
        if (a.italic !== undefined) f.font.italic = !!a.italic;
        if (a.underline !== undefined) f.font.underline = a.underline ? "Single" : "None";
        if (a.font_color) f.font.color = a.font_color;
        if (a.font_size) f.font.size = Number(a.font_size);
        if (a.font_name) f.font.name = a.font_name;
        if (a.fill_color) {
          if (/^none$/i.test(a.fill_color)) f.fill.clear();
          else f.fill.color = a.fill_color;
        }
        if (a.number_format) {
          if (range.rowCount * range.columnCount > 200000) throw new Error("Vùng quá lớn để đặt định dạng số.");
          range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(a.number_format));
        }
        if (a.horizontal_alignment) f.horizontalAlignment = a.horizontal_alignment;
        if (a.wrap_text !== undefined) f.wrapText = !!a.wrap_text;
        if (a.borders) {
          const edges: string[] = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
          const inside: string[] = [];
          if (range.columnCount > 1) inside.push("InsideVertical");
          if (range.rowCount > 1) inside.push("InsideHorizontal");
          if (a.borders === "none") {
            [...edges, ...inside].forEach((e) => (f.borders.getItem(e as any).style = "None"));
          } else {
            const use = a.borders === "all" ? [...edges, ...inside] : edges;
            use.forEach((e) => {
              const b = f.borders.getItem(e as any);
              b.style = "Continuous";
              b.color = "#8EA9C1";
            });
          }
        }
        if (a.column_width) f.columnWidth = Number(a.column_width);
        if (a.autofit_columns) f.autofitColumns();
        await ctx.sync();
        const undo = restoreFrom(snap, `Định dạng ${range.address}`);
        if (undo) log.push(undo);
        return { ok: true, message: `Đã định dạng ${range.address}` };
      });

    case "excel_add_sheet":
      return Excel.run(async (ctx) => {
        const nm = sanitizeSheetName(a.name);
        const exists = ctx.workbook.worksheets.getItemOrNullObject(nm);
        await ctx.sync();
        if (!exists.isNullObject) throw new Error(`Sheet "${nm}" đã có. Chọn tên khác hoặc ghi vào sheet đó.`);
        const ws = ctx.workbook.worksheets.add(nm);
        if (a.activate !== false) ws.activate();
        await ctx.sync();
        log.push({
          label: `Tạo sheet ${nm}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.worksheets.getItem(nm).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã tạo sheet "${nm}"`, sheet: nm };
      });

    case "excel_create_table":
      return Excel.run(async (ctx) => {
        const { ws, range } = getRange(ctx, a.address, a.sheet);
        const snap = await snapshot(ctx, range);
        const t = ws.tables.add(range, a.has_headers !== false);
        if (a.name) t.name = String(a.name).replace(/[^\p{L}\p{N}_]/gu, "_").replace(/^([^\p{L}_])/u, "_$1");
        t.style = a.style || "TableStyleMedium2";
        t.load("name");
        await ctx.sync();
        const tname = t.name;
        log.push({
          label: `Tạo bảng ${tname}`,
          undo: async () => {
            await Excel.run(async (c) => {
              c.workbook.tables.getItem(tname).convertToRange();
              await c.sync();
            });
            const r = restoreFrom(snap, "");
            if (r) await r.undo();
          },
        });
        return { ok: true, message: `Đã tạo bảng ${tname} tại ${range.address}`, table: tname };
      });

    case "excel_create_chart":
      return Excel.run(async (ctx) => {
        const types: Record<string, string> = {
          column: "ColumnClustered",
          bar: "BarClustered",
          line: "Line",
          pie: "Pie",
          area: "Area",
          scatter: "XYScatter",
          doughnut: "Doughnut",
        };
        const by: Record<string, string> = { auto: "Auto", rows: "Rows", columns: "Columns" };
        const { ws, range } = getRange(ctx, a.source_address, a.sheet);
        range.load("address,rowIndex,columnIndex,columnCount");
        ws.load("name");
        await ctx.sync();
        const chart = ws.charts.add(
          (types[a.chart_type] || "ColumnClustered") as any,
          range,
          (by[a.series_by] || "Auto") as any
        );
        if (a.title) chart.title.text = a.title;
        chart.setPosition(ws.getCell(range.rowIndex, range.columnIndex + range.columnCount + 1));
        chart.load("name");
        await ctx.sync();
        const sheetName = ws.name;
        const chartName = chart.name;
        log.push({
          label: `Biểu đồ ${chartName}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.worksheets.getItem(sheetName).charts.getItem(chartName).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã vẽ biểu đồ ${a.chart_type} từ ${range.address}`, chart: chartName };
      });

    case "excel_sort_range":
      return Excel.run(async (ctx) => {
        const { range } = getRange(ctx, a.address, a.sheet);
        range.load("address,columnIndex,columnCount");
        await ctx.sync();
        const col = String(a.column ?? "").trim();
        const key = /^[A-Za-z]{1,3}$/.test(col) ? lettersToCol(col) - range.columnIndex : Number(col);
        if (!Number.isFinite(key) || key < 0 || key >= range.columnCount) {
          throw new Error(`Cột "${col}" không nằm trong vùng ${range.address}.`);
        }
        const snap = await snapshot(ctx, range);
        range.sort.apply([{ key, ascending: a.ascending !== false }], false, a.has_headers !== false);
        await ctx.sync();
        const undo = restoreFrom(snap, `Sắp xếp ${range.address}`);
        if (undo) log.push(undo);
        return { ok: true, message: `Đã sắp xếp ${range.address} theo cột ${col}` };
      });

    case "excel_clear_range":
      return Excel.run(async (ctx) => {
        const { range } = getRange(ctx, a.address, a.sheet);
        const snap = await snapshot(ctx, range);
        const what = a.what === "formats" ? "Formats" : a.what === "all" ? "All" : "Contents";
        range.clear(what as any);
        await ctx.sync();
        const undo = restoreFrom(snap, `Xóa ${range.address}`);
        if (undo) log.push(undo);
        return { ok: true, message: `Đã xóa (${a.what || "contents"}) ${range.address}` };
      });

    case "excel_select_range":
      return Excel.run(async (ctx) => {
        const { ws, range } = getRange(ctx, a.address, a.sheet);
        ws.activate();
        range.select();
        range.load("address");
        await ctx.sync();
        return { ok: true, message: `Đã chọn ${range.address}` };
      });
  }
  throw new Error(`Không có công cụ ${name}`);
}

const RISKY_NAMES = new Set(["excel_write_range", "excel_clear_range", "excel_insert_delete", "excel_sheet"]);

/** Ảnh chụp nội dung để ứng dụng khác (Word, PowerPoint) đọc chéo. */
async function shareSnapshot() {
  const ov = await overview();
  const data = await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const u = ws.getUsedRangeOrNullObject(true);
    u.load("address,rowCount,columnCount");
    await ctx.sync();
    if (u.isNullObject) return null;
    const r = u.getCell(0, 0).getResizedRange(Math.min(u.rowCount, 200) - 1, Math.min(u.columnCount, 30) - 1);
    r.load("address,values");
    await ctx.sync();
    return { address: r.address, values: r.values.map((row) => row.map((v) => truncate(v, 80))) };
  });
  return { title: "Sổ tính Excel đang mở", data: { overview: ov, activeSheetData: data } };
}

export const excelTools: HostToolSet = {
  host: "Excel",
  specs: [...EXCEL_SPECS, ...EXCEL_ADV_SPECS].map((sp) =>
    RISKY_NAMES.has(sp.name) ? { ...sp, risky: (args: any) => excelRisk(sp.name, args) } : sp
  ),
  execute: async (name, a, ctx) => {
    const adv = EXCEL_ADV_SPECS.some((sp) => sp.name === name);
    return adv ? executeExcelAdvanced(name, a || {}, ctx.log) : execute(name, a || {}, ctx);
  },
  snapshot: shareSnapshot,
  overview,
  systemPrompt: (ov) =>
    [
      "Bạn là trợ lý AI được tích hợp trong Microsoft Excel, có thể thao tác trực tiếp trên sổ tính đang mở bằng các công cụ (excel_*).",
      "Nguyên tắc:",
      "- Khi người dùng muốn tạo, sửa, tính toán, định dạng, vẽ biểu đồ… hãy GỌI CÔNG CỤ để làm thật trên sổ tính, không chỉ hướng dẫn cách làm.",
      "- Chưa rõ dữ liệu thì đọc trước (excel_read_range) rồi mới sửa.",
      "- Ưu tiên công thức Excel (tên hàm tiếng Anh, dấu phẩy) thay vì số tính sẵn để bảng tự cập nhật; ghi cả khối bằng một lần excel_write_range.",
      "- Không ghi đè hay xóa dữ liệu có sẵn nếu người dùng không yêu cầu; cần chỗ mới thì dùng vùng trống hoặc sheet mới.",
      "- Kiểm tra formula_errors trong kết quả ghi; có lỗi thì sửa lại.",
      "- Công cụ nâng cao: excel_find (tìm), excel_find_errors (quét lỗi công thức), excel_trace (lần ô phụ thuộc/được tham chiếu), excel_filter, excel_create_pivot, excel_conditional_format, excel_data_validation (danh sách chọn), excel_add_comment, excel_insert_delete (chèn/xóa hàng cột), excel_sheet (đổi tên, cố định dòng/cột, xóa sheet).",
      "- Khi nhắc tới ô hay vùng, ghi rõ dạng Sheet!A1 hoặc Sheet!A1:B5 để người dùng bấm vào xem.",
      "- Xong việc: dùng excel_select_range để chỉ cho người dùng thấy kết quả, rồi trả lời ngắn gọn bằng tiếng Việt: đã làm gì, ở sheet/ô nào.",
      "",
      "Trạng thái hiện tại của sổ tính (JSON):",
      JSON.stringify(ov).slice(0, 6000),
    ].join("\n"),
};
