/* global Excel */
/** Công cụ Excel nâng cao: tìm kiếm, quét lỗi, lần công thức, lọc, PivotTable, định dạng có điều kiện, danh sách chọn, nhận xét ô, chèn/xóa hàng cột, thao tác sheet. */
import { ToolSpec, UndoEntry } from "../core/agent/types";
import { getRange, restoreFrom, snapshot } from "./excelCore";
import { lettersToCol, sanitizeSheetName, splitAddress, truncate } from "./util";

const addressProp = { type: "string", description: "A1-style range, optionally with sheet, e.g. 'Sheet1!A1:D20'." };
const sheetProp = { type: "string", description: "Sheet name (optional; defaults to the active sheet)." };

export const EXCEL_ADV_SPECS: ToolSpec[] = [
  {
    name: "excel_find",
    description: "Find cells containing a text or number across the workbook (or one sheet). Returns addresses.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string" },
        sheet: sheetProp,
        match_case: { type: "boolean" },
        whole_cell: { type: "boolean", description: "Match the entire cell content." },
      },
      required: ["text"],
    },
  },
  {
    name: "excel_find_errors",
    description: "Scan all sheets (or one) for cells whose formulas return errors (#REF!, #DIV/0!, #N/A, #VALUE!, #NAME?…).",
    parameters: { type: "object", properties: { sheet: sheetProp } },
  },
  {
    name: "excel_trace",
    description:
      "Trace how a cell is calculated (precedents) or which cells depend on it (dependents), several levels deep, with formulas and values. Use to explain a number or find the root cause of an error.",
    parameters: {
      type: "object",
      properties: {
        cell: { type: "string", description: "Single cell, e.g. 'Sheet1!C42'." },
        direction: { type: "string", enum: ["precedents", "dependents"], description: "Default precedents." },
        depth: { type: "number", description: "Levels to follow (default 3, max 6)." },
      },
      required: ["cell"],
    },
  },
  {
    name: "excel_filter",
    description:
      "Apply an AutoFilter on a data range by one column, or clear filters. Criteria: list of values to show, or a condition like '>100', '<=5', '*text*' (contains), or top N.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        column: { type: "string", description: "Column letter to filter on, e.g. 'C'." },
        values: { type: "array", items: { type: "string" }, description: "Show only rows with these values." },
        condition: { type: "string", description: "e.g. '>100', '<>0', '*Cần Thơ*'." },
        top: { type: "number", description: "Show top N items." },
        clear: { type: "boolean", description: "Remove all filters on the sheet." },
      },
    },
    write: true,
  },
  {
    name: "excel_create_pivot",
    description:
      "Create a PivotTable from a data range with headers. Field names must match the header texts exactly. Placed on a new sheet by default.",
    parameters: {
      type: "object",
      properties: {
        source_address: addressProp,
        sheet: sheetProp,
        rows: { type: "array", items: { type: "string" }, description: "Header names for row fields." },
        columns: { type: "array", items: { type: "string" }, description: "Header names for column fields (optional)." },
        values: {
          type: "array",
          description: "Value fields with summary function.",
          items: {
            type: "object",
            properties: {
              field: { type: "string" },
              summarize: { type: "string", enum: ["sum", "count", "average", "max", "min"] },
            },
            required: ["field"],
          },
        },
        filters: { type: "array", items: { type: "string" }, description: "Header names for report filters (optional)." },
        destination: { type: "string", description: "'new_sheet' (default) or a top-left cell like 'Sheet2!H2'." },
        name: { type: "string" },
      },
      required: ["source_address", "rows", "values"],
    },
    write: true,
  },
  {
    name: "excel_conditional_format",
    description:
      "Add conditional formatting: color_scale, data_bar, cell_value (compare with value1/value2), text_contains, top_bottom (top/bottom N), duplicates.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        type: { type: "string", enum: ["color_scale", "data_bar", "cell_value", "text_contains", "top_bottom", "duplicates"] },
        operator: {
          type: "string",
          enum: ["GreaterThan", "LessThan", "Between", "EqualTo", "NotEqualTo", "GreaterThanOrEqual", "LessThanOrEqual"],
        },
        value1: { type: "string" },
        value2: { type: "string" },
        text: { type: "string" },
        rank: { type: "number", description: "N for top_bottom (default 10)." },
        bottom: { type: "boolean", description: "Bottom N instead of top N." },
        fill_color: { type: "string", description: "Highlight fill, e.g. #FFC7CE." },
        font_color: { type: "string", description: "Highlight font colour, e.g. #9C0006." },
      },
      required: ["address", "type"],
    },
    write: true,
  },
  {
    name: "excel_data_validation",
    description: "Add a dropdown list or number limits to cells, or clear validation.",
    parameters: {
      type: "object",
      properties: {
        address: addressProp,
        sheet: sheetProp,
        list: { type: "array", items: { type: "string" }, description: "Dropdown items." },
        source_range: { type: "string", description: "Dropdown items from a range, e.g. 'Lists!A2:A10'." },
        min: { type: "number" },
        max: { type: "number" },
        whole_number: { type: "boolean", description: "With min/max: integers only." },
        error_message: { type: "string" },
        clear: { type: "boolean" },
      },
      required: ["address"],
    },
    write: true,
  },
  {
    name: "excel_add_comment",
    description: "Add a comment (note) to a cell, e.g. to document a source or an assumption, or a review remark.",
    parameters: {
      type: "object",
      properties: { cell: { type: "string", description: "e.g. 'Sheet1!B5'." }, text: { type: "string" } },
      required: ["cell", "text"],
    },
    write: true,
  },
  {
    name: "excel_insert_delete",
    description: "Insert or delete whole rows or columns.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["insert_rows", "delete_rows", "insert_columns", "delete_columns"] },
        sheet: sheetProp,
        at: { type: "string", description: "Row number (e.g. '5') or column letter (e.g. 'C') where the action starts." },
        count: { type: "number", description: "Default 1." },
      },
      required: ["action", "at"],
    },
    write: true,
  },
  {
    name: "excel_sheet",
    description: "Sheet operations: rename, freeze top rows / first columns, unfreeze, activate, delete.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["rename", "freeze_rows", "freeze_columns", "unfreeze", "activate", "delete"] },
        sheet: sheetProp,
        new_name: { type: "string" },
        count: { type: "number", description: "Rows/columns to freeze (default 1)." },
      },
      required: ["action"],
    },
    write: true,
  },
];

function sheetOf(ctx: Excel.RequestContext, sheet?: string) {
  return sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
}

function parseAt(at: string): { kind: "row"; n: number } | { kind: "col"; letters: string } {
  const s = String(at || "").trim();
  if (/^\d+$/.test(s)) return { kind: "row", n: Number(s) };
  if (/^[A-Za-z]{1,3}$/.test(s)) return { kind: "col", letters: s.toUpperCase() };
  throw new Error(`"at" phải là số hàng (vd 5) hoặc chữ cột (vd C), nhận "${at}".`);
}

function colPlus(letters: string, n: number): string {
  let c = lettersToCol(letters) + n;
  let out = "";
  c += 1;
  while (c > 0) {
    const r = (c - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    c = Math.floor((c - 1) / 26);
  }
  return out;
}

/** Cảnh báo cho thao tác rủi ro (hiện hộp xác nhận) */
export async function excelRisk(name: string, a: any): Promise<string | null> {
  if (name === "excel_write_range") {
    return Excel.run(async (ctx) => {
      const rows = Array.isArray(a.values) ? a.values.length : 1;
      const cols = Array.isArray(a.values) && Array.isArray(a.values[0]) ? Math.max(...a.values.map((r: any[]) => (Array.isArray(r) ? r.length : 1))) : 1;
      const { range } = getRange(ctx, a.address, a.sheet);
      const target = range.getCell(0, 0).getResizedRange(Math.max(rows, 1) - 1, Math.max(cols, 1) - 1);
      target.load("address,formulas");
      await ctx.sync();
      const filled = target.formulas.flat().filter((v) => v !== "" && v !== null).length;
      return filled ? `Ghi đè ${filled} ô đang có dữ liệu tại ${target.address} (có thể Hoàn tác).` : null;
    }).catch(() => null);
  }
  if (name === "excel_clear_range") return `Xóa ${a.what === "all" ? "nội dung và định dạng" : a.what === "formats" ? "định dạng" : "nội dung"} vùng ${a.sheet ? a.sheet + "!" : ""}${a.address} (có thể Hoàn tác).`;
  if (name === "excel_insert_delete" && String(a.action).startsWith("delete"))
    return `Xóa ${a.count || 1} ${String(a.action).includes("rows") ? "hàng" : "cột"} từ ${a.at}${a.sheet ? " ở sheet " + a.sheet : ""}. Công thức đang tham chiếu tới đó sẽ báo #REF!; Hoàn tác chỉ khôi phục nội dung.`;
  if (name === "excel_sheet" && a.action === "delete") return `Xóa sheet "${a.sheet || "đang mở"}" — KHÔNG hoàn tác được.`;
  return null;
}

export async function executeExcelAdvanced(name: string, a: any, log: UndoEntry[]): Promise<any> {
  switch (name) {
    case "excel_find":
      return Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();
        const list = a.sheet ? sheets.items.filter((w) => w.name === a.sheet) : sheets.items;
        const found: { sheet: string; areas: Excel.RangeAreas }[] = [];
        for (const ws of list) {
          const r = ws.findAllOrNullObject(String(a.text), { completeMatch: !!a.whole_cell, matchCase: !!a.match_case });
          r.load("address,cellCount");
          found.push({ sheet: ws.name, areas: r });
        }
        await ctx.sync();
        const hits = found.filter((f) => !f.areas.isNullObject);
        return {
          ok: true,
          total_cells: hits.reduce((n, f) => n + f.areas.cellCount, 0),
          results: hits.map((f) => ({ sheet: f.sheet, addresses: truncate(f.areas.address, 1500) })),
          message: hits.length ? `Tìm thấy ở ${hits.length} sheet` : `Không thấy "${a.text}"`,
        };
      });

    case "excel_find_errors":
      return Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();
        const list = a.sheet ? sheets.items.filter((w) => w.name === a.sheet) : sheets.items;
        const out: any[] = [];
        let total = 0;
        for (const ws of list) {
          const used = ws.getUsedRangeOrNullObject();
          await ctx.sync();
          if (used.isNullObject) continue;
          const errs = used.getSpecialCellsOrNullObject("Formulas" as any, "Errors" as any);
          errs.load("address,cellCount");
          await ctx.sync();
          if (errs.isNullObject) continue;
          total += errs.cellCount;
          const areas = errs.areas;
          areas.load("items/address,items/values,items/formulas");
          await ctx.sync();
          const cells: any[] = [];
          for (const ar of areas.items) {
            const { address } = splitAddress(ar.address);
            const [tl] = address.split(":");
            const m = tl.match(/^([A-Z]+)(\d+)$/)!;
            ar.values.forEach((row, i) =>
              row.forEach((v, j) => {
                if (cells.length < 80) cells.push({ cell: `${ws.name}!${colPlus(m[1], j)}${Number(m[2]) + i}`, error: v, formula: ar.formulas[i][j] });
              })
            );
          }
          out.push({ sheet: ws.name, count: errs.cellCount, cells });
        }
        return { ok: true, total_errors: total, sheets: out, message: total ? `Có ${total} ô lỗi` : "Không có ô lỗi công thức" };
      });

    case "excel_trace":
      return Excel.run(async (ctx) => {
        const dir = a.direction === "dependents" ? "dependents" : "precedents";
        const maxDepth = Math.min(Math.max(Number(a.depth) || 3, 1), 6);
        const nodes: any[] = [];
        const seen = new Set<string>();
        const queue: { address: string; level: number }[] = [];
        const first = getRange(ctx, a.cell);
        first.range.load("address");
        await ctx.sync();
        queue.push({ address: first.range.address, level: 0 });
        while (queue.length && nodes.length < 60) {
          const cur = queue.shift()!;
          if (seen.has(cur.address)) continue;
          seen.add(cur.address);
          const { range } = getRange(ctx, cur.address);
          range.load("address,formulas,values,cellCount");
          await ctx.sync();
          const node: any = {
            cell: range.address,
            level: cur.level,
            ...(range.cellCount === 1 ? { value: range.values[0][0], formula: range.formulas[0][0] } : { cells: range.cellCount }),
          };
          let links: string[] = [];
          if (range.cellCount === 1 && cur.level < maxDepth) {
            try {
              const r = dir === "precedents" ? range.getDirectPrecedents() : range.getDirectDependents();
              r.load("addresses");
              await ctx.sync();
              links = r.addresses.flatMap((x) => x.split(",")).map((x) => x.trim()).filter(Boolean);
            } catch {
              links = [];
            }
          }
          node[dir] = links;
          nodes.push(node);
          for (const l of links) if (!seen.has(l) && !l.includes(":")) queue.push({ address: l, level: cur.level + 1 });
        }
        const arrow = dir === "precedents" ? "←" : "→";
        const chain = nodes
          .filter((n) => n[dir]?.length)
          .slice(0, 4)
          .map((n) => `${n.cell.split("!").pop()} ${arrow} ${n[dir].map((x: string) => x.split("!").pop()).join(", ")}`)
          .join("; ");
        return {
          ok: true,
          direction: dir,
          nodes,
          message: `${dir === "precedents" ? "Nguồn" : "Ô phụ thuộc"} của ${first.range.address}: ${chain || "không có"}`,
        };
      });

    case "excel_filter":
      return Excel.run(async (ctx) => {
        const ws = sheetOf(ctx, a.sheet ? splitAddress(a.address || "", a.sheet).sheet : splitAddress(a.address || "").sheet);
        if (a.clear) {
          ws.autoFilter.remove();
          await ctx.sync();
          return { ok: true, message: "Đã bỏ lọc" };
        }
        const { range } = getRange(ctx, a.address, a.sheet);
        range.load("address,columnIndex,columnCount");
        await ctx.sync();
        const key = lettersToCol(String(a.column || "A")) - range.columnIndex;
        if (key < 0 || key >= range.columnCount) throw new Error(`Cột ${a.column} không nằm trong ${range.address}.`);
        let criteria: any;
        if (Array.isArray(a.values) && a.values.length) criteria = { filterOn: "Values", values: a.values.map(String) };
        else if (a.top) criteria = { filterOn: "TopItems", criterion1: String(a.top) };
        else if (a.condition) criteria = { filterOn: "Custom", criterion1: String(a.condition) };
        else throw new Error("Cần values, condition hoặc top.");
        ws.autoFilter.apply(range, key, criteria);
        await ctx.sync();
        const wsName = (ws.load("name"), await ctx.sync(), ws.name);
        log.push({
          label: `Lọc ${range.address}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.worksheets.getItem(wsName).autoFilter.remove();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã lọc ${range.address} theo cột ${a.column}` };
      });

    case "excel_create_pivot":
      return Excel.run(async (ctx) => {
        const { range } = getRange(ctx, a.source_address, a.sheet);
        const header = range.getRow(0);
        header.load("values");
        range.load("address");
        await ctx.sync();
        const headers = header.values[0].map((v) => String(v).trim());
        const need = [...(a.rows || []), ...(a.columns || []), ...(a.filters || []), ...(a.values || []).map((v: any) => v.field)];
        const missing = need.filter((f: string) => !headers.includes(String(f).trim()));
        if (missing.length) throw new Error(`Không có cột: ${missing.join(", ")}. Tiêu đề hiện có: ${headers.join(", ")}`);
        let destWs: Excel.Worksheet;
        let dest: Excel.Range;
        let createdSheet: string | null = null;
        if (!a.destination || a.destination === "new_sheet") {
          const sheets = ctx.workbook.worksheets;
          sheets.load("items/name");
          await ctx.sync();
          let nm = sanitizeSheetName(a.name || "Pivot");
          let i = 2;
          while (sheets.items.some((s) => s.name === nm)) nm = sanitizeSheetName(`${a.name || "Pivot"} ${i++}`);
          destWs = ctx.workbook.worksheets.add(nm);
          dest = destWs.getRange("A3");
          createdSheet = nm;
        } else {
          const d = getRange(ctx, a.destination);
          destWs = d.ws;
          dest = d.range;
        }
        const pname = String(a.name || "PivotAIDN" + Date.now().toString(36).slice(-4)).replace(/[^\p{L}\p{N}_]/gu, "_");
        const pt = destWs.pivotTables.add(pname, range, dest);
        for (const f of a.rows || []) pt.rowHierarchies.add(pt.hierarchies.getItem(String(f).trim()));
        for (const f of a.columns || []) pt.columnHierarchies.add(pt.hierarchies.getItem(String(f).trim()));
        for (const f of a.filters || []) pt.filterHierarchies.add(pt.hierarchies.getItem(String(f).trim()));
        const agg: Record<string, string> = { sum: "Sum", count: "Count", average: "Average", max: "Max", min: "Min" };
        for (const v of a.values || []) {
          const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(String(v.field).trim()));
          dh.summarizeBy = (agg[v.summarize] || "Sum") as any;
        }
        destWs.activate();
        destWs.load("name");
        await ctx.sync();
        const wsName = destWs.name;
        log.push({
          label: `PivotTable ${pname}`,
          undo: () =>
            Excel.run(async (c) => {
              if (createdSheet) c.workbook.worksheets.getItem(createdSheet).delete();
              else c.workbook.pivotTables.getItem(pname).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã tạo PivotTable "${pname}" ở sheet ${wsName}`, sheet: wsName, pivot: pname };
      });

    case "excel_conditional_format":
      return Excel.run(async (ctx) => {
        const { ws, range } = getRange(ctx, a.address, a.sheet);
        range.load("address");
        ws.load("name");
        await ctx.sync();
        const fill = a.fill_color || "#FFC7CE";
        const font = a.font_color || "#9C0006";
        let cf: Excel.ConditionalFormat;
        switch (a.type) {
          case "color_scale":
            cf = range.conditionalFormats.add("ColorScale" as any);
            cf.colorScale.criteria = {
              minimum: { formula: null as any, type: "LowestValue" as any, color: "#F8696B" },
              midpoint: { formula: "50", type: "Percentile" as any, color: "#FFEB84" },
              maximum: { formula: null as any, type: "HighestValue" as any, color: "#63BE7B" },
            };
            break;
          case "data_bar":
            cf = range.conditionalFormats.add("DataBar" as any);
            cf.dataBar.positiveFormat.fillColor = a.fill_color || "#5B9BD5";
            break;
          case "cell_value":
            cf = range.conditionalFormats.add("CellValue" as any);
            cf.cellValue.format.fill.color = fill;
            cf.cellValue.format.font.color = font;
            cf.cellValue.rule = {
              formula1: String(a.value1 ?? "0"),
              ...(a.value2 !== undefined ? { formula2: String(a.value2) } : {}),
              operator: (a.operator || "GreaterThan") as any,
            };
            break;
          case "text_contains":
            cf = range.conditionalFormats.add("ContainsText" as any);
            cf.textComparison.format.fill.color = fill;
            cf.textComparison.format.font.color = font;
            cf.textComparison.rule = { operator: "Contains" as any, text: String(a.text || "") };
            break;
          case "top_bottom":
            cf = range.conditionalFormats.add("TopBottom" as any);
            cf.topBottom.format.fill.color = fill;
            cf.topBottom.rule = { rank: Number(a.rank) || 10, type: (a.bottom ? "BottomItems" : "TopItems") as any };
            break;
          case "duplicates":
            cf = range.conditionalFormats.add("PresetCriteria" as any);
            cf.preset.format.fill.color = fill;
            cf.preset.format.font.color = font;
            cf.preset.rule = { criterion: "DuplicateValues" as any };
            break;
          default:
            throw new Error(`Kiểu định dạng có điều kiện không hỗ trợ: ${a.type}`);
        }
        cf.load("id");
        await ctx.sync();
        const id = cf.id;
        const wsName = ws.name;
        const { address } = splitAddress(range.address);
        log.push({
          label: `Định dạng có điều kiện ${range.address}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.worksheets.getItem(wsName).getRange(address).conditionalFormats.getItem(id).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã thêm định dạng có điều kiện (${a.type}) cho ${range.address}` };
      });

    case "excel_data_validation":
      return Excel.run(async (ctx) => {
        const { ws, range } = getRange(ctx, a.address, a.sheet);
        range.load("address");
        ws.load("name");
        range.dataValidation.load("type");
        await ctx.sync();
        const hadRule = range.dataValidation.type !== "None";
        range.dataValidation.clear();
        if (!a.clear) {
          if (Array.isArray(a.list) && a.list.length) {
            range.dataValidation.rule = { list: { inCellDropDown: true, source: a.list.map(String).join(",") } };
          } else if (a.source_range) {
            const src = String(a.source_range).startsWith("=") ? a.source_range : `=${a.source_range}`;
            range.dataValidation.rule = { list: { inCellDropDown: true, source: src } };
          } else if (a.min !== undefined || a.max !== undefined) {
            const rule = {
              formula1: a.min ?? -1e15,
              formula2: a.max ?? 1e15,
              operator: "Between" as any,
            };
            range.dataValidation.rule = a.whole_number ? { wholeNumber: rule } : { decimal: rule };
          } else throw new Error("Cần list, source_range, hoặc min/max.");
          if (a.error_message) {
            range.dataValidation.errorAlert = { message: a.error_message, showAlert: true, style: "Stop" as any, title: "Giá trị không hợp lệ" };
          }
        }
        await ctx.sync();
        const wsName = ws.name;
        const { address } = splitAddress(range.address);
        log.push({
          label: `Kiểm soát nhập liệu ${range.address}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.worksheets.getItem(wsName).getRange(address).dataValidation.clear();
              await c.sync();
            }),
        });
        return {
          ok: true,
          message: a.clear ? `Đã bỏ kiểm soát nhập liệu ở ${range.address}` : `Đã đặt kiểm soát nhập liệu cho ${range.address}`,
          ...(hadRule ? { note: "Ô đã có quy tắc cũ; Hoàn tác sẽ xóa quy tắc (không khôi phục quy tắc cũ)." } : {}),
        };
      });

    case "excel_add_comment":
      return Excel.run(async (ctx) => {
        const { ws, range } = getRange(ctx, a.cell);
        ws.load("name");
        range.load("address");
        await ctx.sync();
        const cellAddr = range.address.split(":")[0];
        const full = cellAddr.includes("!") ? cellAddr : `${ws.name}!${cellAddr}`;
        ctx.workbook.comments.add(full, String(a.text || ""));
        await ctx.sync();
        log.push({
          label: `Nhận xét ${full}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.comments.getItemByCell(full).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã thêm nhận xét vào ${full}` };
      });

    case "excel_insert_delete":
      return Excel.run(async (ctx) => {
        const ws = sheetOf(ctx, a.sheet);
        ws.load("name");
        const at = parseAt(a.at);
        const n = Math.max(1, Number(a.count) || 1);
        const isRows = String(a.action).includes("rows");
        if (isRows !== (at.kind === "row")) throw new Error(isRows ? "Thao tác hàng cần at là số hàng." : "Thao tác cột cần at là chữ cột.");
        const addr = at.kind === "row" ? `${at.n}:${at.n + n - 1}` : `${at.letters}:${colPlus(at.letters, n - 1)}`;
        await ctx.sync();
        const wsName = ws.name;
        if (String(a.action).startsWith("insert")) {
          ws.getRange(addr).insert((isRows ? "Down" : "Right") as any);
          await ctx.sync();
          log.push({
            label: `Chèn ${addr}`,
            undo: () =>
              Excel.run(async (c) => {
                c.workbook.worksheets.getItem(wsName).getRange(addr).delete((isRows ? "Up" : "Left") as any);
                await c.sync();
              }),
          });
          return { ok: true, message: `Đã chèn ${n} ${isRows ? "hàng" : "cột"} tại ${a.at}` };
        }
        // Xóa: chụp phần có dữ liệu để Hoàn tác
        const used = ws.getUsedRangeOrNullObject();
        used.load("address,columnCount,rowCount,columnIndex,rowIndex");
        await ctx.sync();
        let snap = null as any;
        if (!used.isNullObject) {
          const part =
            at.kind === "row"
              ? ws.getRangeByIndexes(at.n - 1, 0, n, used.columnIndex + used.columnCount)
              : ws.getRangeByIndexes(0, lettersToCol(at.letters), used.rowIndex + used.rowCount, n);
          snap = await snapshot(ctx, part);
        }
        ws.getRange(addr).delete((isRows ? "Up" : "Left") as any);
        await ctx.sync();
        log.push({
          label: `Xóa ${addr}`,
          undo: async () => {
            await Excel.run(async (c) => {
              c.workbook.worksheets.getItem(wsName).getRange(addr).insert((isRows ? "Down" : "Right") as any);
              await c.sync();
            });
            const r = restoreFrom(snap, "");
            if (r) await r.undo();
          },
        });
        return { ok: true, message: `Đã xóa ${n} ${isRows ? "hàng" : "cột"} tại ${a.at}` };
      });

    case "excel_sheet":
      return Excel.run(async (ctx) => {
        const ws = sheetOf(ctx, a.sheet);
        ws.load("name");
        await ctx.sync();
        const old = ws.name;
        switch (a.action) {
          case "rename": {
            const nm = sanitizeSheetName(a.new_name);
            ws.name = nm;
            await ctx.sync();
            log.push({
              label: `Đổi tên sheet ${old}`,
              undo: () =>
                Excel.run(async (c) => {
                  c.workbook.worksheets.getItem(nm).name = old;
                  await c.sync();
                }),
            });
            return { ok: true, message: `Đã đổi tên "${old}" → "${nm}"` };
          }
          case "freeze_rows":
          case "freeze_columns": {
            const n = Math.max(1, Number(a.count) || 1);
            if (a.action === "freeze_rows") ws.freezePanes.freezeRows(n);
            else ws.freezePanes.freezeColumns(n);
            await ctx.sync();
            log.push({
              label: `Cố định ${old}`,
              undo: () =>
                Excel.run(async (c) => {
                  c.workbook.worksheets.getItem(old).freezePanes.unfreeze();
                  await c.sync();
                }),
            });
            return { ok: true, message: `Đã cố định ${n} ${a.action === "freeze_rows" ? "hàng đầu" : "cột đầu"} ở "${old}"` };
          }
          case "unfreeze":
            ws.freezePanes.unfreeze();
            await ctx.sync();
            return { ok: true, message: `Đã bỏ cố định ở "${old}"` };
          case "activate":
            ws.activate();
            await ctx.sync();
            return { ok: true, message: `Đã mở sheet "${old}"` };
          case "delete":
            ws.delete();
            await ctx.sync();
            return { ok: true, message: `Đã xóa sheet "${old}"` };
        }
        throw new Error(`Thao tác sheet không hỗ trợ: ${a.action}`);
      });
  }
  return undefined;
}
