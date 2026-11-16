/* global Excel */
/** Công cụ Excel của Đợt 4: liệt kê và sửa biểu đồ / PivotTable có sẵn, chuẩn bị in, tên vùng, kiểm toán mô hình, giải ngược và
 *  bảng nhạy cảm, làm sạch dữ liệu, bảo vệ ô công thức. Các API dùng ở đây đã thăm dò trong Excel thật 16/11/2026
 *  (tasks/dot-4.md, bảng thăm dò; bằng chứng tasks/dot-4-bang-chung/01). */
import { ToolSpec, UndoEntry } from "../core/agent/types";
import { getRange, restoreFrom, snapshot } from "./excelCore";
import { cellAddress, lettersToCol, splitAddress, topLeft, truncate } from "./util";
import { formulaRisksOf } from "../core/formularisk";
import { auditModel } from "../core/modelaudit";
import { cleanNumberList, goalSeek, tidy } from "../core/goalseek";

const addressProp = { type: "string", description: "A1-style range, optionally with sheet, e.g. 'Sheet1!A1:D20'." };
const sheetProp = { type: "string", description: "Sheet name (optional; defaults to the active sheet)." };
const strList = { type: "array", items: { type: "string" } };

export const EXCEL_MODEL_SPECS: ToolSpec[] = [
  {
    name: "excel_list_objects",
    description:
      "List what already exists in the workbook: charts, PivotTables (with their fields), tables, named ranges, and the print setup and protection of each sheet. Call it before editing an existing chart or PivotTable.",
    parameters: { type: "object", properties: { sheet: sheetProp } },
  },
  {
    name: "excel_edit_chart",
    description:
      "Edit an EXISTING chart: change type, title, data source, axis titles, legend; or delete it. Identify the chart by name (from excel_list_objects); omit when the sheet has a single chart.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        chart: { type: "string", description: "Chart name, e.g. 'Chart 1'." },
        chart_type: { type: "string", enum: ["column", "bar", "line", "pie", "area", "scatter", "doughnut", "stacked_column", "stacked_bar"] },
        title: { type: "string" },
        source_address: { type: "string", description: "New contiguous data range (with header row)." },
        series_by: { type: "string", enum: ["auto", "rows", "columns"] },
        x_axis_title: { type: "string" },
        y_axis_title: { type: "string" },
        legend: { type: "string", enum: ["top", "bottom", "left", "right", "none"] },
        delete: { type: "boolean", description: "Delete the chart (asks the user)." },
      },
    },
    write: true,
  },
  {
    name: "excel_edit_pivot",
    description:
      "Edit an EXISTING PivotTable: add or remove fields, change the summary function, keep only some items of a field, clear filters, refresh after the source data changed. Field names = header texts of the source data.",
    parameters: {
      type: "object",
      properties: {
        pivot: { type: "string", description: "PivotTable name (from excel_list_objects); omit when the workbook has a single PivotTable." },
        add_rows: strList,
        add_columns: strList,
        add_filters: strList,
        add_values: {
          type: "array",
          description: "Value fields to add, or existing ones whose summary function should change.",
          items: {
            type: "object",
            properties: { field: { type: "string" }, summarize: { type: "string", enum: ["sum", "count", "average", "max", "min"] } },
            required: ["field"],
          },
        },
        remove: { type: "array", items: { type: "string" }, description: "Source field names to remove from rows, columns, filters and values." },
        filter: {
          type: "object",
          description: "Show only these items of a row/column/filter field.",
          properties: { field: { type: "string" }, items: strList },
          required: ["field", "items"],
        },
        clear_filters: { type: "boolean" },
        refresh: { type: "boolean" },
      },
    },
    write: true,
  },
  {
    name: "excel_page_setup",
    description:
      "Prepare a sheet for printing in ONE call: print area, repeated title rows/columns, orientation, paper size, fit to pages, margins, centering, header/footer, gridlines. Only the given options change.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        print_area: { type: "string", description: "Range to print, e.g. 'A1:F40'; 'none' clears it." },
        title_rows: { type: "string", description: "Rows repeated on every page, e.g. '1:1' or '1:2'; 'none' clears." },
        title_columns: { type: "string", description: "Columns repeated on every page, e.g. 'A:A'; 'none' clears." },
        orientation: { type: "string", enum: ["portrait", "landscape"] },
        paper: { type: "string", enum: ["A4", "A3", "A5", "Letter", "Legal"] },
        fit_width: { type: "number", description: "Fit to this many pages wide (1 = one page wide); 0 = no fit." },
        fit_height: { type: "number", description: "Fit to this many pages tall; 0 or omitted = as many as needed." },
        scale: { type: "number", description: "Print scale in percent (10–400); cancels fit_width/fit_height." },
        margins_cm: {
          type: "object",
          properties: { top: { type: "number" }, bottom: { type: "number" }, left: { type: "number" }, right: { type: "number" }, header: { type: "number" }, footer: { type: "number" } },
        },
        center_horizontally: { type: "boolean" },
        center_vertically: { type: "boolean" },
        header: { type: "string", description: "Centre header text; '' clears." },
        footer: { type: "string", description: "Centre footer text; use &P for page number and &N for page count, e.g. 'Trang &P / &N'." },
        gridlines: { type: "boolean", description: "Print gridlines." },
      },
    },
    write: true,
  },
  {
    name: "excel_named_range",
    description:
      "Create or delete a workbook name for a cell/range (e.g. name the assumption cell 'Lai_suat' so formulas read =Gia_ban*Lai_suat). Names: letters, digits, underscore, no spaces.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["add", "delete"] },
        name: { type: "string" },
        address: { type: "string", description: "For add: the cell or range, e.g. 'Giả định!B3'." },
        comment: { type: "string" },
      },
      required: ["action", "name"],
    },
    write: true,
  },
  {
    name: "excel_audit_model",
    description:
      "Audit a model by code (no reading needed first): hard-coded numbers sitting inside a run of formulas, formulas that differ from their neighbours in the same row/column, rates typed inside formulas (e.g. *0.05), links to other files. Returns points to check — read the cells before telling the user.",
    parameters: { type: "object", properties: { sheet: sheetProp, address: { type: "string", description: "Range to audit (default: the used range of the sheet)." } } },
  },
  {
    name: "excel_what_if",
    description:
      "What-if analysis that leaves the workbook unchanged (inputs are restored afterwards). mode 'goal_seek': find the value of input_cell that makes target_cell equal target_value. mode 'sensitivity': table of output_cell for a list of values of input_cell (and optionally input_cell_2). Input cells must hold plain numbers, not formulas. To keep a result, write it afterwards with excel_write_range.",
    parameters: {
      type: "object",
      properties: {
        mode: { type: "string", enum: ["goal_seek", "sensitivity"] },
        input_cell: { type: "string", description: "e.g. 'Mô hình!B3'." },
        target_cell: { type: "string", description: "goal_seek: formula cell that must reach target_value." },
        target_value: { type: "number" },
        min: { type: "number", description: "goal_seek: lowest allowed input (optional)." },
        max: { type: "number", description: "goal_seek: highest allowed input (optional)." },
        output_cell: { type: "string", description: "sensitivity: formula cell to observe." },
        values: { type: "array", items: { type: "number" }, description: "sensitivity: values to try for input_cell (max 15)." },
        input_cell_2: { type: "string" },
        values_2: { type: "array", items: { type: "number" }, description: "sensitivity: values for input_cell_2 (max 15)." },
      },
      required: ["mode", "input_cell"],
    },
  },
  {
    name: "excel_clean_data",
    description:
      "Clean a data range. 'remove_duplicates': delete rows that repeat (by the given key columns; asks the user). 'replace_all': replace a text everywhere in the range. 'fill_series': extend a series (dates, numbers, text pattern) from the cells in address down/right to fill_to.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["remove_duplicates", "replace_all", "fill_series"] },
        address: addressProp,
        sheet: sheetProp,
        columns: { type: "array", items: { type: "string" }, description: "remove_duplicates: key column letters, e.g. ['A']; default all columns." },
        has_headers: { type: "boolean", description: "remove_duplicates: first row is a header (default true)." },
        find: { type: "string" },
        replace: { type: "string" },
        match_case: { type: "boolean" },
        whole_cell: { type: "boolean" },
        fill_to: { type: "string", description: "fill_series: destination range that includes address, e.g. 'A2:A31'." },
        fill_type: { type: "string", enum: ["default", "series", "days", "weekdays", "months", "years", "copy"] },
      },
      required: ["action", "address"],
    },
    write: true,
  },
  {
    name: "excel_protect",
    description:
      "'lock_formulas': lock every formula cell of the sheet, leave the other cells (inputs) editable, then protect the sheet WITHOUT a password (formatting, sorting and filtering stay allowed). 'unprotect': remove protection (only works when the sheet has no password).",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["lock_formulas", "unprotect"] },
        sheet: sheetProp,
        input_ranges: { type: "array", items: { type: "string" }, description: "lock_formulas: if given, ONLY these ranges stay editable (everything else is locked)." },
      },
      required: ["action"],
    },
    write: true,
  },
];

const CHART_TYPES: Record<string, string> = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "Line",
  pie: "Pie",
  area: "Area",
  scatter: "XYScatter",
  doughnut: "Doughnut",
  stacked_column: "ColumnStacked",
  stacked_bar: "BarStacked",
};
const AGG: Record<string, string> = { sum: "Sum", count: "Count", average: "Average", max: "Max", min: "Min" };
const PAPER: Record<string, string> = { A4: "A4", A3: "A3", A5: "A5", Letter: "Letter", Legal: "Legal" };
const FILL: Record<string, string> = {
  default: "FillDefault",
  series: "FillSeries",
  days: "FillDays",
  weekdays: "FillWeekdays",
  months: "FillMonths",
  years: "FillYears",
  copy: "FillCopy",
};
const CM = 28.3464567;
const cm = (pt: number) => Math.round((pt / CM) * 100) / 100;
/** Chỉ chữ "none" (hoặc "không", "clear") mới là lệnh bỏ; chuỗi rỗng model điền cho tham số không dùng thì bỏ qua tham số đó */
const isNone = (v: any) => /^(none|không|khong|clear)$/i.test(String(v ?? "").trim());
const given = (v: any) => v !== undefined && v !== null && String(v).trim() !== "";
const list = (v: any): string[] => (Array.isArray(v) ? v.map((x) => String(x).trim()).filter(Boolean) : typeof v === "string" && v.trim() ? [v.trim()] : []);

function sheetOf(ctx: Excel.RequestContext, sheet?: string) {
  return sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
}

/** Tên vùng hợp lệ: chữ, số, gạch dưới; không bắt đầu bằng số; không trông như địa chỉ ô (A1, R1C1). */
export function cleanName(raw: any): string {
  let n = String(raw ?? "")
    .trim()
    .replace(/[^\p{L}\p{N}_.]/gu, "_")
    .replace(/_+/g, "_");
  if (!n) throw new Error("Thiếu tên vùng (name).");
  if (/^[\p{N}.]/u.test(n)) n = "_" + n;
  if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[Rr]\d*[Cc]\d*$/.test(n) || /^[RrCc]$/.test(n)) n = "_" + n;
  return n.slice(0, 200);
}

async function findChart(ctx: Excel.RequestContext, a: any) {
  const ws = sheetOf(ctx, a.sheet);
  ws.load("name");
  ws.charts.load("items/name");
  await ctx.sync();
  const items = ws.charts.items;
  if (!items.length) throw new Error(`Sheet "${ws.name}" không có biểu đồ nào. Gọi excel_list_objects để xem biểu đồ nằm ở sheet nào.`);
  const want = String(a.chart ?? "").trim();
  let ch = want ? items.find((c) => c.name === want) || items.find((c) => c.name.toLowerCase() === want.toLowerCase()) : undefined;
  if (!ch && /^\d+$/.test(want)) ch = items[Number(want) - 1];
  if (!ch && !want && items.length === 1) ch = items[0];
  if (!ch) throw new Error(`Không rõ biểu đồ nào${want ? ` (không có "${want}")` : ""}. Sheet "${ws.name}" có: ${items.map((c) => c.name).join(", ")}.`);
  return { ws, chart: ch, sheetName: ws.name };
}

async function findPivot(ctx: Excel.RequestContext, a: any) {
  const pts = ctx.workbook.pivotTables;
  pts.load("items/name");
  await ctx.sync();
  if (!pts.items.length) throw new Error("Sổ tính chưa có PivotTable nào. Tạo bằng excel_create_pivot.");
  const want = String(a.pivot ?? "").trim();
  const pt = want ? pts.items.find((p) => p.name === want) || pts.items.find((p) => p.name.toLowerCase() === want.toLowerCase()) : pts.items.length === 1 ? pts.items[0] : undefined;
  if (!pt) throw new Error(`Không rõ PivotTable nào${want ? ` (không có "${want}")` : ""}. Đang có: ${pts.items.map((p) => p.name).join(", ")}.`);
  return pt;
}

/** Trường của Pivot theo tên tiêu đề nguồn, báo lỗi kèm danh sách tên đúng */
async function pivotFieldNames(ctx: Excel.RequestContext, pt: Excel.PivotTable) {
  pt.hierarchies.load("items/name");
  pt.rowHierarchies.load("items/name");
  pt.columnHierarchies.load("items/name");
  pt.filterHierarchies.load("items/name");
  pt.dataHierarchies.load("items/name,items/summarizeBy");
  await ctx.sync();
  const dataFields: { name: string; source: string; summarize: string }[] = [];
  for (const d of pt.dataHierarchies.items) {
    d.field.load("name");
  }
  await ctx.sync();
  pt.dataHierarchies.items.forEach((d) => dataFields.push({ name: d.name, source: d.field.name, summarize: String(d.summarizeBy) }));
  return {
    all: pt.hierarchies.items.map((h) => h.name),
    rows: pt.rowHierarchies.items.map((h) => h.name).filter((n) => n !== "Values"),
    columns: pt.columnHierarchies.items.map((h) => h.name).filter((n) => n !== "Values"),
    filters: pt.filterHierarchies.items.map((h) => h.name),
    values: dataFields,
  };
}

async function readPageSetup(ctx: Excel.RequestContext, ws: Excel.Worksheet) {
  const pl = ws.pageLayout;
  pl.load("orientation,paperSize,zoom,centerHorizontally,centerVertically,printGridlines,topMargin,bottomMargin,leftMargin,rightMargin,headerMargin,footerMargin");
  const pa = pl.getPrintAreaOrNullObject();
  pa.load("address");
  const tr = pl.getPrintTitleRowsOrNullObject();
  tr.load("address");
  const tc = pl.getPrintTitleColumnsOrNullObject();
  tc.load("address");
  const hf = pl.headersFooters.defaultForAllPages;
  hf.load("centerHeader,centerFooter");
  await ctx.sync();
  const local = (addr: string) => splitAddress(addr).address;
  return {
    orientation: String(pl.orientation).toLowerCase(),
    paper: pl.paperSize ? String(pl.paperSize) : null,
    fit_width: pl.zoom?.horizontalFitToPages ?? null,
    fit_height: pl.zoom?.verticalFitToPages ?? null,
    scale: pl.zoom?.scale ?? null,
    print_area: pa.isNullObject ? null : local(pa.address),
    title_rows: tr.isNullObject ? null : local(tr.address),
    title_columns: tc.isNullObject ? null : local(tc.address),
    margins_cm: { top: cm(pl.topMargin), bottom: cm(pl.bottomMargin), left: cm(pl.leftMargin), right: cm(pl.rightMargin), header: cm(pl.headerMargin), footer: cm(pl.footerMargin) },
    center_horizontally: pl.centerHorizontally,
    center_vertically: pl.centerVertically,
    header: hf.centerHeader || "",
    footer: hf.centerFooter || "",
    gridlines: pl.printGridlines,
  };
}

function applyPageSetup(ws: Excel.Worksheet, a: any) {
  const pl = ws.pageLayout;
  const done: string[] = [];
  if (a.orientation) {
    pl.orientation = (/^land/i.test(a.orientation) ? "Landscape" : "Portrait") as any;
    done.push(/^land/i.test(a.orientation) ? "giấy ngang" : "giấy dọc");
  }
  if (a.paper) {
    const p = PAPER[String(a.paper)] || PAPER[String(a.paper).toUpperCase()] || PAPER[String(a.paper)[0].toUpperCase() + String(a.paper).slice(1).toLowerCase()];
    if (!p) throw new Error(`Khổ giấy không hỗ trợ: ${a.paper} (dùng A4, A3, A5, Letter, Legal).`);
    pl.paperSize = p as any;
    done.push(`khổ ${p}`);
  }
  const hasFit = a.fit_width !== undefined || a.fit_height !== undefined;
  if (a.scale !== undefined && a.scale !== null) {
    const sc = Math.round(Number(a.scale));
    if (!(sc >= 10 && sc <= 400)) throw new Error("scale phải trong khoảng 10–400 (phần trăm).");
    pl.zoom = { scale: sc };
    done.push(`tỷ lệ in ${sc}%`);
  } else if (hasFit) {
    const w = Math.max(0, Math.floor(Number(a.fit_width) || 0));
    const h = Math.max(0, Math.floor(Number(a.fit_height) || 0));
    if (!w && !h) {
      pl.zoom = { scale: 100 };
      done.push("bỏ co vừa trang");
    } else {
      pl.zoom = { ...(w ? { horizontalFitToPages: w } : {}), ...(h ? { verticalFitToPages: h } : {}) } as any;
      done.push(`co vừa ${w ? `${w} trang ngang` : ""}${w && h ? ", " : ""}${h ? `${h} trang dọc` : ""}`);
    }
  }
  if (given(a.print_area)) {
    if (isNone(a.print_area)) pl.setPrintArea("" as any);
    else pl.setPrintArea(ws.getRange(splitAddress(String(a.print_area)).address));
    done.push(isNone(a.print_area) ? "bỏ vùng in" : `vùng in ${splitAddress(String(a.print_area)).address}`);
  }
  if (given(a.title_rows)) {
    if (isNone(a.title_rows)) pl.setPrintTitleRows("" as any);
    else {
      const t = splitAddress(String(a.title_rows)).address;
      pl.setPrintTitleRows(ws.getRange(/^\d+$/.test(t) ? `${t}:${t}` : t));
    }
    done.push(isNone(a.title_rows) ? "bỏ dòng tiêu đề lặp" : `dòng tiêu đề lặp ${a.title_rows}`);
  }
  if (given(a.title_columns)) {
    if (isNone(a.title_columns)) pl.setPrintTitleColumns("" as any);
    else {
      const t = splitAddress(String(a.title_columns)).address;
      pl.setPrintTitleColumns(ws.getRange(/^[A-Za-z]+$/.test(t) ? `${t}:${t}` : t));
    }
    done.push(isNone(a.title_columns) ? "bỏ cột tiêu đề lặp" : `cột tiêu đề lặp ${a.title_columns}`);
  }
  const m = a.margins_cm;
  if (m && typeof m === "object") {
    const opt: any = {};
    for (const k of ["top", "bottom", "left", "right", "header", "footer"]) {
      const v = Number(m[k]);
      if (m[k] !== undefined && m[k] !== null && Number.isFinite(v) && v >= 0 && v <= 10) opt[k] = v;
    }
    if (Object.keys(opt).length) {
      pl.setPrintMargins("Centimeters" as any, opt);
      done.push("lề");
    }
  }
  if (a.center_horizontally !== undefined) pl.centerHorizontally = !!a.center_horizontally;
  if (a.center_vertically !== undefined) pl.centerVertically = !!a.center_vertically;
  if (a.center_horizontally !== undefined || a.center_vertically !== undefined) done.push("căn giữa trang");
  if (a.header !== undefined && a.header !== null) {
    pl.headersFooters.defaultForAllPages.centerHeader = String(a.header);
    done.push("đầu trang");
  }
  if (a.footer !== undefined && a.footer !== null) {
    pl.headersFooters.defaultForAllPages.centerFooter = String(a.footer);
    done.push("chân trang");
  }
  if (a.gridlines !== undefined) {
    pl.printGridlines = !!a.gridlines;
    done.push(a.gridlines ? "in đường lưới" : "không in đường lưới");
  }
  return done;
}

/** Cảnh báo rủi ro cho các công cụ của tệp này (null = không hỏi) */
export async function excelModelRisk(name: string, a: any): Promise<string | null> {
  if (name === "excel_edit_chart" && a.delete) return `Xóa biểu đồ${a.chart ? ` "${a.chart}"` : ""}${a.sheet ? ` ở sheet ${a.sheet}` : ""} — KHÔNG hoàn tác được.`;
  if (name === "excel_named_range" && a.action === "delete") return `Xóa tên vùng "${a.name}": các công thức đang dùng tên này sẽ báo #NAME? (có thể Hoàn tác).`;
  if (name === "excel_clean_data" && a.action === "remove_duplicates")
    return `Xóa các dòng trùng trong ${a.sheet && !String(a.address).includes("!") ? a.sheet + "!" : ""}${a.address}${list(a.columns).length ? ` (so theo cột ${list(a.columns).join(", ")})` : " (so mọi cột)"}; chỉ giữ dòng đầu tiên của mỗi nhóm trùng (Hoàn tác được khi vùng không quá 10.000 ô).`;
  if (name === "excel_clean_data" && a.action === "fill_series") {
    // Tự điền ghi đè mọi ô của vùng đích: có dữ liệu ngoài các ô mẫu thì hỏi (quy tắc ghi đè luôn hỏi)
    return Excel.run(async (ctx) => {
      const { ws, range } = getRange(ctx, a.address, a.sheet);
      ws.load("name");
      await ctx.sync();
      const dest = getRange(ctx, String(a.fill_to || a.address), a.sheet || ws.name).range;
      range.load("formulas");
      dest.load("formulas,address,rowCount,columnCount");
      await ctx.sync();
      const filled = (g: any[][]) => g.flat().filter((v) => v !== "" && v !== null).length;
      const extra = filled(dest.formulas) - filled(range.formulas);
      const big = dest.rowCount * dest.columnCount > 10000;
      return extra > 0 ? `Tự điền sẽ ghi đè ${extra} ô đang có dữ liệu trong ${dest.address}${big ? " — vùng lớn nên KHÔNG hoàn tác được" : " (có thể Hoàn tác)"}.` : null;
    }).catch(() => null);
  }
  if (name === "excel_clean_data" && a.action === "replace_all") {
    // Chữ thay vào là công thức có hàm cần hỏi: hỏi như khi ghi ô (không để thay hàng loạt thành đường vòng)
    const fr = formulaRisksOf(String(a.replace ?? "").trim());
    const frMsg = fr.length ? `Chữ thay vào là công thức có ${fr.map((x) => x.what).join(", ")} (${fr[0].why}): ${truncate(a.replace, 120)}.` : "";
    const many = await Excel.run(async (ctx) => {
      const { range } = getRange(ctx, a.address, a.sheet);
      const used = range.getUsedRangeOrNullObject(true);
      used.load("formulas");
      await ctx.sync();
      const find = a.match_case ? String(a.find ?? "") : String(a.find ?? "").toLowerCase();
      const n =
        used.isNullObject || !find
          ? 0
          : used.formulas.flat().filter((v) => {
              const t = a.match_case ? String(v) : String(v).toLowerCase();
              return a.whole_cell ? t === find : t.includes(find);
            }).length;
      return n > 20 ? `Thay "${truncate(a.find, 40)}" → "${truncate(a.replace, 40)}" ở ${n} ô trong ${a.address} (Hoàn tác được khi vùng không quá 10.000 ô).` : null;
    }).catch(() => null);
    return [many, frMsg].filter(Boolean).join(" ") || null;
  }
  if (name === "excel_protect" && a.action === "unprotect") return `Bỏ bảo vệ sheet "${a.sheet || "đang mở"}": mọi ô, kể cả ô công thức, sẽ sửa được.`;
  return null;
}

export async function executeExcelModel(name: string, a: any, log: UndoEntry[]): Promise<any> {
  switch (name) {
    case "excel_list_objects":
      return Excel.run(async (ctx) => {
        const wss = ctx.workbook.worksheets;
        wss.load("items/name");
        ctx.workbook.names.load("items/name,items/formula,items/visible");
        await ctx.sync();
        const sheets = (a.sheet ? wss.items.filter((w) => w.name === a.sheet) : wss.items).slice(0, 30);
        if (a.sheet && !sheets.length) throw new Error(`Không có sheet "${a.sheet}".`);
        const out: any[] = [];
        for (const ws of sheets) {
          ws.charts.load("items/name,items/chartType");
          ws.pivotTables.load("items/name");
          ws.tables.load("items/name");
          ws.protection.load("protected");
          await ctx.sync();
          const charts: any[] = [];
          for (const c of ws.charts.items.slice(0, 20)) {
            c.title.load("text,visible");
            c.series.load("items/name");
            await ctx.sync();
            let source = "";
            try {
              const s0 = c.series.items[0]?.getDimensionDataSourceString("Values" as any);
              await ctx.sync();
              source = s0?.value || "";
            } catch {
              source = "";
            }
            charts.push({ name: c.name, type: String(c.chartType), title: c.title.visible ? c.title.text : "", series: c.series.items.map((s) => s.name).slice(0, 12), ...(source ? { first_series_source: source } : {}) });
          }
          const pivots: any[] = [];
          for (const p of ws.pivotTables.items.slice(0, 10)) {
            const f = await pivotFieldNames(ctx, p);
            let source = "";
            try {
              const src = p.getDataSourceString();
              await ctx.sync();
              source = src.value;
            } catch {
              source = "";
            }
            pivots.push({ name: p.name, source, rows: f.rows, columns: f.columns, filters: f.filters, values: f.values.map((v) => `${v.source} (${v.summarize})`), unused_fields: f.all.filter((n) => ![...f.rows, ...f.columns, ...f.filters, ...f.values.map((v) => v.source)].includes(n)) });
          }
          const tables: any[] = [];
          for (const t of ws.tables.items.slice(0, 20)) {
            const r = t.getRange();
            r.load("address");
            await ctx.sync();
            tables.push({ name: t.name, address: r.address });
          }
          const print = await readPageSetup(ctx, ws);
          out.push({ sheet: ws.name, charts, pivots, tables, print, protected: ws.protection.protected });
        }
        const names = ctx.workbook.names.items.filter((n) => n.visible).slice(0, 60).map((n) => ({ name: n.name, refers_to: n.formula }));
        const nCharts = out.reduce((n, s) => n + s.charts.length, 0);
        const nPivots = out.reduce((n, s) => n + s.pivots.length, 0);
        return { ok: true, sheets: out, names, brief: `${nCharts} biểu đồ, ${nPivots} PivotTable, ${names.length} tên vùng`, message: `Có ${nCharts} biểu đồ, ${nPivots} PivotTable, ${names.length} tên vùng` };
      });

    case "excel_edit_chart":
      return Excel.run(async (ctx) => {
        const { ws, chart, sheetName } = await findChart(ctx, a);
        const chartName = chart.name;
        if (a.delete) {
          chart.delete();
          await ctx.sync();
          return { ok: true, message: `Đã xóa biểu đồ "${chartName}" ở sheet ${sheetName}` };
        }
        chart.load("chartType");
        chart.title.load("text,visible");
        chart.legend.load("visible,position");
        await ctx.sync();
        const before = { type: chart.chartType, title: chart.title.text, titleVisible: chart.title.visible, legendVisible: chart.legend.visible, legendPosition: chart.legend.position };
        const done: string[] = [];
        let note = "";
        if (a.source_address) {
          const { range } = getRange(ctx, a.source_address, a.sheet);
          const by: Record<string, string> = { auto: "Auto", rows: "Rows", columns: "Columns" };
          chart.setData(range, (by[a.series_by] || "Auto") as any);
          done.push(`nguồn dữ liệu ${a.source_address}`);
          note = "Hoàn tác trả lại kiểu, tiêu đề và chú giải; nguồn dữ liệu cũ không tự trả lại được.";
        }
        if (a.chart_type) {
          const t = CHART_TYPES[String(a.chart_type)];
          if (!t) throw new Error(`Kiểu biểu đồ không hỗ trợ: ${a.chart_type}`);
          chart.chartType = t as any;
          done.push(`kiểu ${a.chart_type}`);
        }
        if (a.title !== undefined && a.title !== null) {
          chart.title.visible = String(a.title) !== "";
          if (String(a.title) !== "") chart.title.text = String(a.title);
          done.push(`tiêu đề "${truncate(a.title, 50)}"`);
        }
        if (a.legend) {
          if (a.legend === "none") chart.legend.visible = false;
          else {
            chart.legend.visible = true;
            chart.legend.position = (String(a.legend)[0].toUpperCase() + String(a.legend).slice(1)) as any;
          }
          done.push(`chú giải ${a.legend}`);
        }
        await ctx.sync();
        // Tiêu đề trục: biểu đồ tròn, vành khuyên không có trục → báo riêng, không làm hỏng phần đã sửa
        for (const [key, axis] of [["x_axis_title", "categoryAxis"], ["y_axis_title", "valueAxis"]] as const) {
          if (a[key] === undefined || a[key] === null) continue;
          try {
            const ax = (chart.axes as any)[axis] as Excel.ChartAxis;
            ax.title.visible = String(a[key]) !== "";
            if (String(a[key]) !== "") ax.title.text = String(a[key]);
            await ctx.sync();
            done.push(`tiêu đề trục ${key === "x_axis_title" ? "ngang" : "dọc"}`);
          } catch {
            note += (note ? " " : "") + `Không đặt được tiêu đề trục (${key}): kiểu biểu đồ này không có trục đó.`;
          }
        }
        if (!done.length) throw new Error("Chưa nêu cần sửa gì ở biểu đồ (chart_type, title, source_address, x_axis_title, y_axis_title, legend hoặc delete).");
        chart.load("chartType");
        chart.title.load("text");
        await ctx.sync();
        log.push({
          label: `Sửa biểu đồ ${chartName}`,
          undo: () =>
            Excel.run(async (c) => {
              const ch = c.workbook.worksheets.getItem(sheetName).charts.getItem(chartName);
              ch.chartType = before.type as any;
              ch.title.visible = before.titleVisible;
              if (before.titleVisible && before.title) ch.title.text = before.title;
              ch.legend.visible = before.legendVisible;
              if (before.legendVisible) ch.legend.position = before.legendPosition as any;
              await c.sync();
            }),
        });
        void ws;
        return { ok: true, message: `Đã sửa biểu đồ "${chartName}" (${sheetName}): ${done.join(", ")}`, chart: chartName, type: String(chart.chartType), title: chart.title.text, ...(note ? { note } : {}) };
      });

    case "excel_edit_pivot":
      return Excel.run(async (ctx) => {
        const pt = await findPivot(ctx, a);
        const pname = pt.name;
        const f = await pivotFieldNames(ctx, pt);
        const known = (n: string) => {
          const hit = f.all.find((x) => x === n) || f.all.find((x) => x.trim().toLowerCase() === n.trim().toLowerCase());
          if (!hit) throw new Error(`PivotTable "${pname}" không có trường "${n}". Các trường: ${f.all.join(", ")}.`);
          return hit;
        };
        // Kiểm mọi tên trường TRƯỚC khi đổi bất cứ gì: tên sai thì báo lỗi khi Pivot còn nguyên (soát mã 16/11/2026, #18)
        for (const n of [...list(a.remove), ...list(a.add_rows), ...list(a.add_columns), ...list(a.add_filters)]) known(n);
        for (const v of Array.isArray(a.add_values) ? a.add_values : []) known(String(v?.field ?? v ?? ""));
        if (a.filter && typeof a.filter === "object") known(String(a.filter.field ?? ""));
        const done: string[] = [];
        const undo: ((p: Excel.PivotTable) => void)[] = [];
        // Bớt trước, thêm sau (một trường có thể đổi từ hàng sang cột trong cùng một lệnh)
        for (const raw of list(a.remove)) {
          const n = known(raw);
          let hit = false;
          if (f.rows.includes(n)) {
            pt.rowHierarchies.remove(pt.rowHierarchies.getItem(n));
            undo.push((p) => p.rowHierarchies.add(p.hierarchies.getItem(n)));
            f.rows = f.rows.filter((x) => x !== n);
            hit = true;
          }
          if (f.columns.includes(n)) {
            pt.columnHierarchies.remove(pt.columnHierarchies.getItem(n));
            undo.push((p) => p.columnHierarchies.add(p.hierarchies.getItem(n)));
            f.columns = f.columns.filter((x) => x !== n);
            hit = true;
          }
          if (f.filters.includes(n)) {
            pt.filterHierarchies.remove(pt.filterHierarchies.getItem(n));
            undo.push((p) => p.filterHierarchies.add(p.hierarchies.getItem(n)));
            f.filters = f.filters.filter((x) => x !== n);
            hit = true;
          }
          for (const v of f.values.filter((x) => x.source === n)) {
            pt.dataHierarchies.remove(pt.dataHierarchies.getItem(v.name));
            const sum = v.summarize;
            undo.push((p) => {
              p.dataHierarchies.add(p.hierarchies.getItem(n)).summarizeBy = sum as any;
            });
            hit = true;
          }
          f.values = f.values.filter((x) => x.source !== n);
          if (!hit) throw new Error(`Trường "${n}" hiện không nằm trong hàng, cột, bộ lọc hay giá trị của PivotTable.`);
          done.push(`bớt ${n}`);
        }
        await ctx.sync();
        const used = () => [...f.rows, ...f.columns, ...f.filters];
        const addAxis = (names: string[], kind: "rows" | "columns" | "filters") => {
          for (const raw of names) {
            const n = known(raw);
            if (used().includes(n)) throw new Error(`Trường "${n}" đã nằm trong PivotTable (hàng, cột hoặc bộ lọc). Muốn chuyển chỗ thì đưa nó vào remove trong cùng lệnh.`);
            const col: any = kind === "rows" ? pt.rowHierarchies : kind === "columns" ? pt.columnHierarchies : pt.filterHierarchies;
            col.add(pt.hierarchies.getItem(n));
            f[kind].push(n);
            undo.push((p) => {
              const c2: any = kind === "rows" ? p.rowHierarchies : kind === "columns" ? p.columnHierarchies : p.filterHierarchies;
              c2.remove(c2.getItem(n));
            });
            done.push(`thêm ${n} vào ${kind === "rows" ? "hàng" : kind === "columns" ? "cột" : "bộ lọc"}`);
          }
        };
        addAxis(list(a.add_rows), "rows");
        addAxis(list(a.add_columns), "columns");
        addAxis(list(a.add_filters), "filters");
        for (const v of Array.isArray(a.add_values) ? a.add_values : []) {
          const n = known(String(v?.field ?? v ?? ""));
          const sum = AGG[String(v?.summarize || "sum").toLowerCase()] || "Sum";
          const have = f.values.find((x) => x.source === n);
          if (have) {
            const dh = pt.dataHierarchies.getItem(have.name);
            dh.summarizeBy = sum as any;
            const old = have.summarize;
            dh.load("name");
            await ctx.sync();
            const newName = dh.name;
            undo.push((p) => {
              p.dataHierarchies.getItem(newName).summarizeBy = old as any;
            });
            done.push(`đổi hàm của ${n} sang ${sum}`);
          } else {
            const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(n));
            dh.summarizeBy = sum as any;
            dh.load("name");
            await ctx.sync();
            const newName = dh.name;
            undo.push((p) => p.dataHierarchies.remove(p.dataHierarchies.getItem(newName)));
            done.push(`thêm giá trị ${n} (${sum})`);
          }
        }
        await ctx.sync();
        if (a.clear_filters) {
          for (const n of used()) pt.hierarchies.getItem(n).fields.getItem(n).clearAllFilters();
          done.push("bỏ mọi lọc");
        }
        if (a.filter && typeof a.filter === "object") {
          const n = known(String(a.filter.field ?? ""));
          if (!used().includes(n)) throw new Error(`Muốn lọc theo "${n}" thì trường đó phải nằm trong hàng, cột hoặc bộ lọc của PivotTable.`);
          const items = list(a.filter.items);
          if (!items.length) throw new Error("filter.items rỗng.");
          const fld = pt.hierarchies.getItem(n).fields.getItem(n);
          fld.items.load("items/name");
          await ctx.sync();
          const have = fld.items.items.map((x) => x.name);
          const missing = items.filter((x) => !have.includes(x));
          if (missing.length) throw new Error(`Trường "${n}" không có mục: ${missing.join(", ")}. Các mục: ${have.slice(0, 40).join(", ")}.`);
          fld.applyFilter({ manualFilter: { selectedItems: items } });
          undo.push((p) => p.hierarchies.getItem(n).fields.getItem(n).clearAllFilters());
          done.push(`lọc ${n} = ${items.join(", ")}`);
        }
        if (a.refresh) {
          pt.refresh();
          done.push("làm mới");
        }
        if (!done.length) throw new Error("Chưa nêu cần sửa gì ở PivotTable (add_rows, add_columns, add_values, add_filters, remove, filter, clear_filters hoặc refresh).");
        await ctx.sync();
        const after = await pivotFieldNames(ctx, pt);
        const rng = pt.layout.getRange();
        rng.load("address");
        await ctx.sync();
        if (undo.length) {
          log.push({
            label: `Sửa PivotTable ${pname}`,
            undo: () =>
              Excel.run(async (c) => {
                const p = c.workbook.pivotTables.getItem(pname);
                for (const u of undo.slice().reverse()) {
                  u(p);
                  await c.sync();
                }
              }),
          });
        }
        return {
          ok: true,
          message: `Đã sửa PivotTable "${pname}": ${done.join("; ")}`,
          pivot: pname,
          address: rng.address,
          rows: after.rows,
          columns: after.columns,
          filters: after.filters,
          values: after.values.map((v) => `${v.source} (${v.summarize})`),
        };
      });

    case "excel_page_setup":
      return Excel.run(async (ctx) => {
        const ws = sheetOf(ctx, a.sheet || (a.print_area ? splitAddress(String(a.print_area)).sheet : undefined));
        ws.load("name");
        const before = await readPageSetup(ctx, ws);
        const done = applyPageSetup(ws, a);
        if (!done.length) throw new Error("Chưa nêu thiết lập in nào (print_area, title_rows, orientation, paper, fit_width, margins_cm…).");
        await ctx.sync();
        const now = await readPageSetup(ctx, ws);
        const sheetName = ws.name;
        log.push({
          label: `Thiết lập in ${sheetName}`,
          undo: () =>
            Excel.run(async (c) => {
              const w = c.workbook.worksheets.getItem(sheetName);
              applyPageSetup(w, {
                orientation: before.orientation,
                ...(before.paper && PAPER[before.paper] ? { paper: before.paper } : {}),
                ...(before.fit_width || before.fit_height ? { fit_width: before.fit_width || 0, fit_height: before.fit_height || 0 } : { scale: before.scale || 100 }),
                print_area: before.print_area || "none",
                title_rows: before.title_rows || "none",
                title_columns: before.title_columns || "none",
                margins_cm: before.margins_cm,
                center_horizontally: before.center_horizontally,
                center_vertically: before.center_vertically,
                header: before.header,
                footer: before.footer,
                gridlines: before.gridlines,
              });
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã thiết lập in cho sheet "${sheetName}": ${done.join(", ")}`, sheet: sheetName, print: now };
      });

    case "excel_named_range":
      return Excel.run(async (ctx) => {
        const names = ctx.workbook.names;
        if (a.action === "delete") {
          const want = String(a.name ?? "").trim();
          const it = names.getItemOrNullObject(want);
          it.load("name,formula,comment");
          await ctx.sync();
          if (it.isNullObject) throw new Error(`Không có tên vùng "${want}". Gọi excel_list_objects để xem các tên đang có.`);
          const old = { name: it.name, formula: it.formula, comment: it.comment };
          it.delete();
          await ctx.sync();
          log.push({
            label: `Xóa tên vùng ${old.name}`,
            undo: () =>
              Excel.run(async (c) => {
                c.workbook.names.add(old.name, old.formula, old.comment || undefined);
                await c.sync();
              }),
          });
          return { ok: true, message: `Đã xóa tên vùng "${old.name}" (trỏ tới ${old.formula})` };
        }
        if (a.action !== "add") throw new Error(`action phải là add hoặc delete, nhận "${a.action}".`);
        const nm = cleanName(a.name);
        if (!a.address) throw new Error("Thiếu address: ô hoặc vùng mà tên trỏ tới.");
        const exists = names.getItemOrNullObject(nm);
        exists.load("formula");
        await ctx.sync();
        if (!exists.isNullObject) throw new Error(`Tên vùng "${nm}" đã có (đang trỏ tới ${exists.formula}). Chọn tên khác, hoặc xóa tên cũ trước.`);
        const { range } = getRange(ctx, a.address);
        range.load("address");
        names.add(nm, range, a.comment ? String(a.comment).slice(0, 250) : undefined);
        await ctx.sync();
        log.push({
          label: `Đặt tên vùng ${nm}`,
          undo: () =>
            Excel.run(async (c) => {
              c.workbook.names.getItem(nm).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã đặt tên "${nm}" cho ${range.address}. Dùng trong công thức: =${nm}`, name: nm, address: range.address };
      });

    case "excel_audit_model":
      return Excel.run(async (ctx) => {
        let range: Excel.Range;
        let ws: Excel.Worksheet;
        if (a.address) {
          const g = getRange(ctx, a.address, a.sheet);
          range = g.range;
          ws = g.ws;
        } else {
          ws = sheetOf(ctx, a.sheet);
          const u = ws.getUsedRangeOrNullObject(true);
          u.load("address");
          await ctx.sync();
          if (u.isNullObject) return { ok: true, total: 0, findings: [], brief: "Sheet trống", message: "Sheet trống, không có gì để kiểm." };
          range = u;
        }
        ws.load("name");
        range.load("address,rowCount,columnCount");
        await ctx.sync();
        let note = "";
        if (range.rowCount * range.columnCount > 60000) {
          const rows = Math.max(1, Math.floor(60000 / Math.min(range.columnCount, 200)));
          range = range.getCell(0, 0).getResizedRange(Math.min(range.rowCount, rows) - 1, Math.min(range.columnCount, 200) - 1);
          note = "Vùng rất lớn: chỉ kiểm phần đầu. Gọi lại với address cụ thể cho phần còn lại.";
        }
        range.load("address,formulasR1C1,formulas");
        await ctx.sync();
        const tl = topLeft(range.address);
        const res = auditModel({ r1c1: range.formulasR1C1, formulas: range.formulas, origin: tl, addr: cellAddress });
        const sheetName = ws.name;
        const pref = (c: string) => `${/[^\p{L}\p{N}_]/u.test(sheetName) ? `'${sheetName.replace(/'/g, "''")}'` : sheetName}!${c}`;
        const findings = res.findings.slice(0, 80).map((f) => ({ kind: f.kind, cell: pref(f.cell), ...(f.count > 1 ? { count: f.count, cells: f.cells.map(pref) } : {}), detail: f.detail }));
        const c = res.counts;
        const brief = `${c.hardcoded} hằng số gõ cứng, ${c.inconsistent} công thức lệch, ${c.embedded_rate} ô có tỷ lệ gõ trong công thức, ${c.external_link} ô liên kết tệp ngoài`;
        return {
          ok: true,
          address: range.address,
          scanned: res.scanned,
          counts: c,
          findings,
          brief,
          message: `Kiểm ${res.scanned.formulas} công thức ở ${range.address}: ${brief}`,
          ...(res.findings.length > 80 ? { truncated: `Còn ${res.findings.length - 80} điểm chưa liệt kê.` } : {}),
          ...(note ? { note } : {}),
          how_to_read:
            "hardcoded = số gõ đè giữa dãy công thức; inconsistent = công thức khác mẫu của các ô cùng hàng/cột; embedded_rate = tỷ lệ gõ thẳng trong công thức (nên đưa ra ô giả định); external_link = lấy số từ tệp khác. Đây là điểm cần xem, đọc ô trước khi kết luận.",
        };
      });

    case "excel_what_if":
      return whatIf(a);

    case "excel_clean_data":
      return Excel.run(async (ctx) => {
        const { ws, range } = getRange(ctx, a.address, a.sheet);
        range.load("address,columnIndex,columnCount,rowCount");
        ws.load("name");
        await ctx.sync();
        if (a.action === "remove_duplicates") {
          const cols = list(a.columns).map((c) => {
            const k = /^[A-Za-z]{1,3}$/.test(c) ? lettersToCol(c) - range.columnIndex : NaN;
            if (!Number.isInteger(k) || k < 0 || k >= range.columnCount) throw new Error(`Cột "${c}" không nằm trong vùng ${range.address}.`);
            return k;
          });
          const keys = cols.length ? cols : Array.from({ length: range.columnCount }, (_, i) => i);
          const snap = await snapshot(ctx, range);
          const rd = range.removeDuplicates(keys, a.has_headers !== false);
          rd.load("removed,uniqueRemaining");
          await ctx.sync();
          const u = restoreFrom(snap, `Xóa dòng trùng ${range.address}`);
          if (u && rd.removed) log.push(u);
          return { ok: true, removed: rd.removed, remaining: rd.uniqueRemaining, message: rd.removed ? `Đã xóa ${rd.removed} dòng trùng ở ${range.address}, còn ${rd.uniqueRemaining} dòng khác nhau` : `Không có dòng trùng ở ${range.address}`, ...(snap ? {} : { note: "Vùng lớn nên không lưu được bản để Hoàn tác." }) };
        }
        if (a.action === "replace_all") {
          const find = String(a.find ?? "");
          if (!find) throw new Error("Thiếu find: chữ cần thay.");
          const snap = await snapshot(ctx, range);
          const n = range.replaceAll(find, String(a.replace ?? ""), { completeMatch: !!a.whole_cell, matchCase: !!a.match_case });
          await ctx.sync();
          const u = restoreFrom(snap, `Thay hàng loạt ${range.address}`);
          if (u && n.value) log.push(u);
          return { ok: true, replaced: n.value, message: n.value ? `Đã thay ${n.value} chỗ "${truncate(find, 40)}" → "${truncate(a.replace, 40)}" ở ${range.address}` : `Không thấy "${truncate(find, 40)}" ở ${range.address}: thay 0 chỗ`, ...(snap ? {} : { note: "Vùng lớn nên không lưu được bản để Hoàn tác." }) };
        }
        if (a.action === "fill_series") {
          if (!a.fill_to) throw new Error("Thiếu fill_to: vùng đích (gồm cả các ô mẫu), ví dụ 'A2:A31'.");
          const dest = getRange(ctx, a.fill_to, a.sheet || ws.name).range;
          dest.load("address,rowCount,columnCount");
          await ctx.sync();
          if (dest.rowCount * dest.columnCount > 20000) throw new Error("Vùng điền quá lớn (trên 20.000 ô).");
          const snap = await snapshot(ctx, dest);
          range.autoFill(dest, (FILL[String(a.fill_type || "default")] || "FillDefault") as any);
          await ctx.sync();
          const first = dest.getCell(0, 0).getResizedRange(Math.min(dest.rowCount, 6) - 1, Math.min(dest.columnCount, 6) - 1);
          first.load("text");
          await ctx.sync();
          const u = restoreFrom(snap, `Tự điền ${dest.address}`);
          if (u) log.push(u);
          return { ok: true, message: `Đã tự điền ${dest.address} theo mẫu ở ${range.address}`, address: dest.address, preview: first.text, ...(snap ? {} : { note: "Vùng lớn nên không lưu được bản để Hoàn tác." }) };
        }
        throw new Error(`action không hỗ trợ: ${a.action}`);
      });

    case "excel_protect":
      return Excel.run(async (ctx) => {
        const ws = sheetOf(ctx, a.sheet);
        ws.load("name");
        ws.protection.load("protected,options");
        await ctx.sync();
        const sheetName = ws.name;
        if (a.action === "unprotect") {
          if (!ws.protection.protected) return { ok: true, message: `Sheet "${sheetName}" đang không được bảo vệ.` };
          const opts = ws.protection.options;
          try {
            ws.protection.unprotect();
            await ctx.sync();
          } catch (e: any) {
            throw new Error(`Không bỏ được bảo vệ sheet "${sheetName}": sheet có mật khẩu. Người dùng tự bỏ ở thẻ Review → Unprotect Sheet. (${e?.code || e?.message || e})`);
          }
          log.push({
            label: `Bỏ bảo vệ ${sheetName}`,
            undo: () =>
              Excel.run(async (c) => {
                c.workbook.worksheets.getItem(sheetName).protection.protect(opts || undefined);
                await c.sync();
              }),
          });
          return { ok: true, message: `Đã bỏ bảo vệ sheet "${sheetName}"` };
        }
        if (a.action !== "lock_formulas") throw new Error(`action phải là lock_formulas hoặc unprotect, nhận "${a.action}".`);
        if (ws.protection.protected) throw new Error(`Sheet "${sheetName}" đang được bảo vệ. Bỏ bảo vệ trước (action unprotect, cần người dùng đồng ý) rồi mới đặt lại.`);
        const used = ws.getUsedRangeOrNullObject();
        used.load("address");
        used.format.protection.load("locked");
        await ctx.sync();
        if (used.isNullObject) throw new Error(`Sheet "${sheetName}" trống.`);
        const prevLocked = used.format.protection.locked;
        const inputs = list(a.input_ranges);
        let nFormula = 0;
        // Trạng thái khóa của CẢ sheet trước khi đổi (true / false = đồng nhất; null = đang lẫn lộn, tức đã có người bố trí khóa riêng)
        const whole = ws.getRange();
        whole.format.protection.load("locked");
        const fs = used.getSpecialCellsOrNullObject("Formulas" as any);
        fs.load("cellCount");
        await ctx.sync();
        const wholeLocked = whole.format.protection.locked;
        // Kiểm xong mọi điều kiện rồi mới đổi khóa (không để dở dang khi báo lỗi)
        if (!inputs.length && fs.isNullObject) throw new Error(`Sheet "${sheetName}" không có ô công thức nào để khóa. Muốn khóa cả sheet trừ ô nhập thì nêu input_ranges.`);
        if (inputs.length) {
          whole.format.protection.locked = true;
          for (const r of inputs) ws.getRange(splitAddress(r).address).format.protection.locked = false;
        } else {
          // Mở khóa cả sheet (kể cả ô ngoài vùng đang dùng, để còn thêm được dòng mới), rồi chỉ khóa ô công thức
          whole.format.protection.locked = false;
          nFormula = fs.cellCount;
          fs.format.protection.locked = true;
        }
        ws.protection.protect({
          allowFormatCells: true,
          allowFormatColumns: true,
          allowFormatRows: true,
          allowAutoFilter: true,
          allowSort: true,
          selectionMode: "Normal" as any,
        });
        await ctx.sync();
        void prevLocked;
        log.push({
          label: `Bảo vệ ${sheetName}`,
          undo: () =>
            Excel.run(async (c) => {
              const w = c.workbook.worksheets.getItem(sheetName);
              w.protection.unprotect();
              // Trạng thái khóa cũ đồng nhất thì trả lại; lẫn lộn (null) thì về mặc định của Excel: mọi ô đều khóa
              w.getRange().format.protection.locked = wholeLocked === false ? false : true;
              await c.sync();
            }),
        });
        return {
          ok: true,
          message: inputs.length
            ? `Đã bảo vệ sheet "${sheetName}" (không mật khẩu): chỉ sửa được ${inputs.join(", ")}`
            : `Đã khóa ${nFormula} ô công thức và bảo vệ sheet "${sheetName}" (không mật khẩu); các ô còn lại vẫn nhập được`,
          note:
            "Không đặt mật khẩu: ai cũng bỏ bảo vệ được ở thẻ Review. Bảo vệ này chống sửa nhầm, không phải bảo mật. Khi sheet đang bảo vệ, AI cũng không ghi được vào ô bị khóa." +
            (wholeLocked === null ? " Sheet trước đó đã có bố trí khóa / mở khóa riêng từng ô: bố trí đó đã bị thay; Hoàn tác chỉ bỏ bảo vệ và đưa mọi ô về trạng thái khóa mặc định." : ""),
        };
      });
  }
  return undefined;
}

/** Giải ngược và bảng nhạy cảm: ghi tạm ô đầu vào, đọc ô kết quả, rồi LUÔN trả ô đầu vào về nội dung gốc (kể cả khi lỗi giữa chừng). */
async function whatIf(a: any): Promise<any> {
  const mode = a.mode === "sensitivity" ? "sensitivity" : a.mode === "goal_seek" ? "goal_seek" : "";
  if (!mode) throw new Error(`mode phải là goal_seek hoặc sensitivity, nhận "${a.mode}".`);
  const outAddr = mode === "goal_seek" ? a.target_cell : a.output_cell;
  if (!outAddr) throw new Error(mode === "goal_seek" ? "Thiếu target_cell." : "Thiếu output_cell.");
  const inputs: string[] = [a.input_cell, ...(mode === "sensitivity" && a.input_cell_2 ? [a.input_cell_2] : [])].map(String);

  // Đọc trạng thái gốc
  const orig = await Excel.run(async (ctx) => {
    ctx.workbook.application.load("calculationMode");
    const cells = inputs.map((x) => {
      const r = getRange(ctx, x).range;
      r.load("address,formulas,values,cellCount");
      return r;
    });
    const out = getRange(ctx, outAddr).range;
    out.load("address,formulas,values,cellCount");
    await ctx.sync();
    return {
      manual: String(ctx.workbook.application.calculationMode) === "Manual",
      inputs: cells.map((r) => ({ address: r.address, formula: r.formulas[0][0], value: r.values[0][0], count: r.cellCount })),
      out: { address: out.address, formula: out.formulas[0][0], value: out.values[0][0], count: out.cellCount },
    };
  });
  for (const c of [...orig.inputs, orig.out]) if (c.count !== 1) throw new Error(`${c.address} phải là MỘT ô.`);
  for (const c of orig.inputs) {
    if (typeof c.formula === "string" && c.formula.startsWith("=")) throw new Error(`Ô đầu vào ${c.address} đang là công thức (${truncate(c.formula, 60)}): chọn ô chứa số gõ trực tiếp (ô giả định) mà công thức này lấy số.`);
    if (typeof c.value !== "number") throw new Error(`Ô đầu vào ${c.address} phải đang chứa một số (hiện là "${truncate(c.value, 40)}").`);
  }
  if (!(typeof orig.out.formula === "string" && orig.out.formula.startsWith("="))) throw new Error(`Ô kết quả ${orig.out.address} phải là ô công thức.`);
  if (new Set([...orig.inputs.map((c) => c.address), orig.out.address]).size !== orig.inputs.length + 1) throw new Error("Ô đầu vào và ô kết quả phải là các ô khác nhau.");

  /** Ghi các ô đầu vào rồi đọc ô kết quả (một lần đồng bộ; chế độ tính tay thì tính lại cả sổ) */
  /** Đã có lần ghi thử nào tới được sổ tính chưa (chưa thì không có gì phải trả lại) */
  let touched = false;
  const evalAt = (vals: number[]) =>
    Excel.run(async (ctx) => {
      vals.forEach((v, i) => (getRange(ctx, orig.inputs[i].address).range.values = [[v]]));
      if (orig.manual) ctx.workbook.application.calculate("Full" as any);
      const out = getRange(ctx, orig.out.address).range;
      out.load("values");
      // Đặt cờ TRƯỚC khi đồng bộ: nếu lần đồng bộ lỗi giữa chừng vẫn phải thử trả lại
      touched = true;
      try {
        await ctx.sync();
      } catch (e: any) {
        // Ô đầu vào bị khóa (sheet bảo vệ): lệnh ghi bị từ chối trọn gói, sổ tính chưa đổi
        if (e?.code === "AccessDenied") touched = false;
        throw e;
      }
      const y = out.values[0][0];
      return typeof y === "number" ? y : NaN;
    });
  const restore = async () => {
    let lastErr: any;
    for (let k = 0; k < 3; k++) {
      try {
        await Excel.run(async (ctx) => {
          orig.inputs.forEach((c) => (getRange(ctx, c.address).range.formulas = [[c.formula]]));
          if (orig.manual) ctx.workbook.application.calculate("Full" as any);
          const rs = orig.inputs.map((c) => {
            const r = getRange(ctx, c.address).range;
            r.load("formulas");
            return r;
          });
          await ctx.sync();
          rs.forEach((r, i) => {
            if (r.formulas[0][0] !== orig.inputs[i].formula) throw new Error(`đọc lại ${orig.inputs[i].address} chưa đúng`);
          });
        });
        return null;
      } catch (e: any) {
        lastErr = e;
        await new Promise((r) => setTimeout(r, 300));
      }
    }
    return `KHÔNG trả lại được ô đầu vào về giá trị gốc (${orig.inputs.map((c) => `${c.address} = ${c.formula}`).join("; ")}): ${lastErr?.message || lastErr}. Báo ngay cho người dùng để họ nhập lại.`;
  };

  let result: any;
  let failure: any;
  try {
    if (mode === "goal_seek") {
      const target = Number(a.target_value);
      if (!Number.isFinite(target)) throw new Error("Thiếu target_value (một số).");
      const gs = await goalSeek((x) => evalAt([x]), { target, start: orig.inputs[0].value as number, min: a.min, max: a.max });
      result = {
        ok: gs.ok,
        mode,
        input_cell: orig.inputs[0].address,
        input_now: orig.inputs[0].value,
        ...(Number.isFinite(gs.x) ? { input_found: tidy(gs.x), target_cell_would_be: tidy(gs.y) } : {}),
        target_cell: orig.out.address,
        target_value: target,
        target_cell_now: orig.out.value,
        relative_error: Number.isFinite(gs.error) ? Number(gs.error.toExponential(2)) : null,
        iterations: gs.iterations,
        ...(gs.ok ? {} : { error: gs.reason }),
        brief: gs.ok ? `${orig.inputs[0].address} = ${tidy(gs.x)} thì ${orig.out.address} = ${tidy(gs.y)}` : "Không tìm được",
        message: gs.ok
          ? `Để ${orig.out.address} đạt ${target}, ${orig.inputs[0].address} cần bằng ${tidy(gs.x)} (đang là ${orig.inputs[0].value}). Sổ tính CHƯA đổi: ô đầu vào đã được trả về giá trị gốc.`
          : `Không tìm được đầu vào đạt mục tiêu: ${gs.reason}. Sổ tính không đổi.`,
      };
    } else {
      const v1 = cleanNumberList(a.values);
      const v2 = orig.inputs.length > 1 ? cleanNumberList(a.values_2) : [];
      if (!v1.length) throw new Error("Thiếu values: danh sách số thử cho input_cell.");
      if (orig.inputs.length > 1 && !v2.length) throw new Error("Thiếu values_2: danh sách số thử cho input_cell_2.");
      const table: any[][] = [];
      if (v2.length) {
        table.push([`${orig.inputs[0].address} \\ ${orig.inputs[1].address}`, ...v2]);
        for (const x of v1) {
          const row: any[] = [x];
          for (const y of v2) {
            const r = await evalAt([x, y]);
            row.push(Number.isFinite(r) ? tidy(r) : "#LỖI");
          }
          table.push(row);
        }
      } else {
        table.push([orig.inputs[0].address, orig.out.address]);
        for (const x of v1) {
          const r = await evalAt([x]);
          table.push([x, Number.isFinite(r) ? tidy(r) : "#LỖI"]);
        }
      }
      result = {
        ok: true,
        mode,
        output_cell: orig.out.address,
        output_now: orig.out.value,
        inputs_now: orig.inputs.map((c) => ({ cell: c.address, value: c.value })),
        table,
        brief: `Bảng nhạy cảm ${v1.length}${v2.length ? ` × ${v2.length}` : ""} giá trị`,
        message: `Đã tính ${v1.length * Math.max(1, v2.length)} phương án cho ${orig.out.address}. Sổ tính CHƯA đổi: các ô đầu vào đã được trả về giá trị gốc. Muốn lưu bảng thì ghi table ra vùng trống hoặc sheet mới bằng excel_write_range.`,
      };
    }
  } catch (e) {
    failure = e;
  }
  const bad = touched ? await restore() : null;
  if (bad) {
    if (failure) throw new Error(`${failure?.message || failure}. ${bad}`);
    return { ...result, ok: false, error: bad, message: bad };
  }
  if (failure) throw failure;
  return result;
}
