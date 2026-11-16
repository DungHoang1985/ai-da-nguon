/* global Excel */
/** Dấu vết soát lại cho Excel (Đợt 4): lớp bọc quanh bộ công cụ Excel.
 *  - Bật "nhật ký": sau mỗi thao tác ghi thành công, thêm dòng vào sheet "Nhật ký AI" (giờ, ô, giá trị cũ, giá trị mới, thao tác và lý do,
 *    yêu cầu của lượt). Bấm Hoàn tác thì ghi thêm một dòng "Hoàn tác".
 *  - Bật "tô sáng": tô nền các ô AI vừa ghi; Hoàn tác trả lại màu cũ (bản chụp của công cụ ghi lấy trước khi tô).
 *  - Luôn luôn: AI không được ghi vào sheet nhật ký; kết quả trả cho model không chứa dữ liệu nội bộ (__trail).
 *  Mặc định cả hai công tắc đều tắt. Ghi nhật ký lỗi thì không làm hỏng thao tác chính. */
import { HostToolSet, ToolCall, UndoEntry } from "../core/agent/types";
import { CellChange, HIGHLIGHT_COLOR, LOG_HEADER, LOG_SHEET, buildLogRows, diffCells, isLogSheet, targetsLogSheet, undoLogRow } from "../core/auditlog";
import { cellAddress, splitAddress, topLeft } from "./util";

export interface TrailOptions {
  /** Đọc cài đặt lúc chạy: { excelLog, excelHighlight } */
  settings: () => any;
  describe: (call: ToolCall) => string;
  /** Tên sheet nhật ký (Tự kiểm tra dùng tên riêng để không đụng nhật ký thật) */
  logSheet?: string;
  now?: () => Date;
}

/** Dữ liệu công cụ ghi gửi kèm cho lớp bọc (bị gỡ khỏi kết quả trước khi trả cho model) */
export interface TrailData {
  sheet: string;
  address: string;
  before: any[][] | null;
  after: any[][];
  /** Có bản chụp để Hoàn tác (không có thì không tô sáng: tô rồi không trả lại màu cũ được) */
  undoable: boolean;
}

interface Highlighted {
  sheet: string;
  address: string;
  fills: any[][];
}
/** Các vùng đã tô trong phiên này kèm màu cũ, để nút "Bỏ tô sáng" trả lại (mất khi tải lại khung) */
const highlighted: Highlighted[] = [];

export function highlightedCount(): number {
  return highlighted.length;
}

async function appendRows(logSheet: string, rows: string[][]): Promise<void> {
  if (!rows.length) return;
  await Excel.run(async (ctx) => {
    let ws = ctx.workbook.worksheets.getItemOrNullObject(logSheet);
    await ctx.sync();
    let next = 2;
    if (ws.isNullObject) {
      const active = ctx.workbook.worksheets.getActiveWorksheet();
      ws = ctx.workbook.worksheets.add(logSheet);
      const head = ws.getRange("A1:G1");
      head.values = [LOG_HEADER];
      head.format.font.bold = true;
      head.format.fill.color = "#DDEBF7";
      ws.getRange("A:A").format.columnWidth = 110;
      ws.getRange("B:C").format.columnWidth = 80;
      ws.getRange("D:E").format.columnWidth = 150;
      ws.getRange("F:G").format.columnWidth = 300;
      ws.freezePanes.freezeRows(1);
      // Tạo sheet nhật ký không được kéo người dùng khỏi sheet đang làm
      active.activate();
    } else {
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("rowIndex,rowCount");
      await ctx.sync();
      next = used.isNullObject ? 2 : used.rowIndex + used.rowCount + 1;
    }
    const target = ws.getRange(`A${next}:G${next + rows.length - 1}`);
    target.numberFormat = rows.map((r) => r.map(() => "@"));
    target.values = rows;
    target.format.wrapText = false;
    await ctx.sync();
  });
}

async function highlight(t: TrailData, cells: CellChange[]): Promise<void> {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(t.sheet);
    const r = ws.getRange(t.address);
    const props = r.getCellProperties({ format: { fill: { color: true, pattern: true } } });
    await ctx.sync();
    highlighted.push({ sheet: t.sheet, address: t.address, fills: props.value });
    if (highlighted.length > 200) highlighted.shift();
    // Chỉ tô các ô đổi nội dung; cả khối đổi thì tô một lần
    const total = (t.after.length || 1) * (t.after[0]?.length || 1);
    if (cells.length >= total || cells.length > 300) r.format.fill.color = HIGHLIGHT_COLOR;
    else for (const c of cells) ws.getRange(c.cell).format.fill.color = HIGHLIGHT_COLOR;
    await ctx.sync();
  });
}

/** Trả lại màu nền cũ cho các ô còn đang mang màu tô sáng (ô người dùng đã tự đổi màu thì để nguyên). Trả số vùng đã xử lý. */
export async function clearHighlights(): Promise<number> {
  let n = 0;
  while (highlighted.length) {
    const h = highlighted.pop()!;
    try {
      await Excel.run(async (ctx) => {
        const r = ctx.workbook.worksheets.getItem(h.sheet).getRange(h.address);
        const cur = r.getCellProperties({ format: { fill: { color: true } } });
        await ctx.sync();
        const next = cur.value.map((row: any[], i: number) =>
          row.map((cell: any, j: number) => {
            if (String(cell?.format?.fill?.color || "").toUpperCase() !== HIGHLIGHT_COLOR) return {};
            const old = h.fills?.[i]?.[j]?.format?.fill;
            return { format: { fill: old?.color && old.pattern !== "None" ? { color: old.color } : { pattern: "None" } } };
          })
        );
        r.setCellProperties(next as any);
        await ctx.sync();
      });
      n++;
    } catch {
      /* sheet đã bị xóa hoặc đổi tên: bỏ qua */
    }
  }
  return n;
}

const whereOf = (a: any): { sheet: string; address: string } => {
  for (const k of ["address", "cell", "source_address", "print_area", "fill_to"]) {
    if (typeof a?.[k] === "string" && a[k].trim()) {
      const s = splitAddress(a[k].split(",")[0], a.sheet);
      return { sheet: s.sheet || "", address: s.address };
    }
  }
  return { sheet: String(a?.sheet || a?.name || ""), address: a?.at ? String(a.at) : "" };
};

export function withAuditTrail(base: HostToolSet, o: TrailOptions): HostToolSet {
  const logSheet = o.logSheet || LOG_SHEET;
  const now = o.now || (() => new Date());
  const isWrite = (name: string) => !!base.specs.find((s) => s.name === name)?.write;
  return {
    ...base,
    overview: async () => {
      const ov = await base.overview();
      // Sheet nhật ký không cần nằm trong trạng thái gửi model mỗi lượt (chứa lời cũ của người dùng và của AI)
      if (ov && Array.isArray(ov.sheets)) {
        ov.sheets = ov.sheets.map((s: any) => (isLogSheet(s?.name, logSheet) ? { ...s, firstRows: [], note: "Sheet nhật ký do add-in tự ghi; AI không sửa sheet này." } : s));
      }
      return ov;
    },
    execute: async (name, args, ctx: { log: UndoEntry[]; said?: string; request?: string }) => {
      const a = args || {};
      const write = isWrite(name);
      if (write && targetsLogSheet(a, logSheet) && !(name === "excel_sheet" && a.action === "delete")) {
        throw new Error(`Sheet “${logSheet}” là nhật ký do add-in tự ghi: AI không ghi, sửa hay đổi tên sheet này. Người dùng muốn bỏ nhật ký thì xóa sheet (excel_sheet, action delete) hoặc tắt nhật ký trong Cài đặt.`);
      }
      const logBefore = ctx.log.length;
      let result: any;
      try {
        result = await base.execute(name, a, ctx);
      } catch (e: any) {
        if (e?.code === "AccessDenied" && write) {
          const err: any = new Error(`${e.message || e} — Sheet có thể đang được bảo vệ (ô bị khóa). Nói rõ với người dùng; chỉ bỏ bảo vệ (excel_protect, action unprotect) khi họ đồng ý.`);
          err.code = e.code;
          throw err;
        }
        throw e;
      }
      let trail: TrailData | undefined;
      if (result && typeof result === "object" && "__trail" in result) {
        const { __trail, ...rest } = result;
        trail = __trail as TrailData;
        result = rest;
      }
      if (!write || result?.ok === false) return result;
      const st = o.settings() || {};
      if (!st.excelLog && !st.excelHighlight) return result;

      const w = trail ? { sheet: trail.sheet, address: trail.address } : whereOf(a);
      let sheet = w.sheet;
      if (!sheet) {
        try {
          sheet = await Excel.run(async (c) => {
            const s = c.workbook.worksheets.getActiveWorksheet();
            s.load("name");
            await c.sync();
            return s.name;
          });
        } catch {
          sheet = "";
        }
      }
      const changes = trail ? diffCells(trail.before, trail.after, topLeft(trail.address), cellAddress) : undefined;
      const action = o.describe({ id: "", name, args: a });
      const notes: string[] = [];
      if (st.excelHighlight && trail?.undoable && changes?.length) {
        try {
          await highlight({ ...trail, sheet }, changes);
        } catch (e: any) {
          notes.push(`Không tô sáng được: ${e?.message || e}`);
        }
      }
      if (st.excelLog) {
        try {
          await appendRows(logSheet, buildLogRows({ time: now(), sheet, address: w.address, changes, action, reason: ctx.said, request: ctx.request }));
        } catch (e: any) {
          notes.push(`Không ghi được nhật ký: ${e?.message || e}`);
        }
        // Hoàn tác các mục do lần gọi này thêm → ghi một dòng "Hoàn tác" (đọc lại cài đặt lúc bấm)
        for (let i = logBefore; i < ctx.log.length; i++) {
          const entry = ctx.log[i];
          const orig = entry.undo;
          entry.undo = async () => {
            await orig();
            if ((o.settings() || {}).excelLog) await appendRows(logSheet, [undoLogRow(now(), sheet, w.address, entry.label || action)]).catch(() => undefined);
          };
        }
      }
      return notes.length ? { ...result, trail_note: notes.join(" ") } : result;
    },
  };
}
