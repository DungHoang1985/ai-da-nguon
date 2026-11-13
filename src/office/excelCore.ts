/* global Excel */
/** Hàm dùng chung cho các công cụ Excel: lấy vùng, chụp trạng thái và tạo thao tác Hoàn tác. */
import { UndoEntry } from "../core/agent/types";
import { splitAddress } from "./util";

const MAX_UNDO_CELLS = 10000;

export function getRange(ctx: Excel.RequestContext, address: string, sheet?: string) {
  if (!address) throw new Error("Thiếu địa chỉ ô (address).");
  const a = splitAddress(address, sheet);
  const ws = a.sheet ? ctx.workbook.worksheets.getItem(a.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
  return { ws, range: ws.getRange(a.address) };
}

export interface Snapshot {
  sheet: string;
  address: string;
  formulas: any[][];
  numberFormat: any[][];
  props: any;
  columns?: any;
}

/** Chụp lại nội dung + định dạng trước khi sửa (giới hạn 10.000 ô). */
export async function snapshot(ctx: Excel.RequestContext, range: Excel.Range, withColumns = false): Promise<Snapshot | null> {
  range.load("address,rowCount,columnCount");
  await ctx.sync();
  if (range.rowCount * range.columnCount > MAX_UNDO_CELLS) return null;
  range.load("formulas,numberFormat");
  const props = range.getCellProperties({
    format: {
      font: { bold: true, italic: true, underline: true, color: true, size: true, name: true },
      fill: { color: true, pattern: true },
      horizontalAlignment: true,
      verticalAlignment: true,
      wrapText: true,
      borders: { style: true, color: true, weight: true },
    },
  });
  const cols = withColumns ? range.getColumnProperties({ format: { columnWidth: true } }) : null;
  await ctx.sync();
  const { sheet, address } = splitAddress(range.address);
  return {
    sheet: sheet || "",
    address,
    formulas: range.formulas,
    numberFormat: range.numberFormat,
    props: props.value,
    columns: cols ? cols.value : undefined,
  };
}

/** Chỉ giữ các thuộc tính ghi lại được (bỏ address, viền trong, giá trị rỗng) — nếu không setCellProperties báo lỗi tham số. */
export function settableProps(props: any[][]): any[][] {
  const clean = (o: any) => {
    const out: any = {};
    for (const [k, v] of Object.entries(o || {})) if (v !== null && v !== undefined && v !== "") out[k] = v;
    return out;
  };
  return (props || []).map((row) =>
    (row || []).map((cell: any) => {
      const f = cell?.format || {};
      const format: any = {};
      const font = clean(f.font);
      if (Object.keys(font).length) format.font = font;
      if (f.fill) format.fill = f.fill.pattern === "None" || !f.fill.color ? { pattern: "None" } : clean({ color: f.fill.color, pattern: f.fill.pattern });
      for (const k of ["horizontalAlignment", "verticalAlignment", "wrapText"]) if (f[k] !== null && f[k] !== undefined) format[k] = f[k];
      if (f.borders) {
        const b: any = {};
        for (const edge of ["top", "bottom", "left", "right"]) {
          const e = f.borders[edge];
          if (!e) continue;
          b[edge] = !e.style || e.style === "None" ? { style: "None" } : clean({ style: e.style, color: e.color, weight: e.weight });
        }
        if (Object.keys(b).length) format.borders = b;
      }
      return { format };
    })
  );
}

export function restoreFrom(snap: Snapshot | null, label: string): UndoEntry | null {
  if (!snap) return null;
  return {
    label,
    undo: async () => {
      // Bước 1: nội dung + định dạng số (bắt buộc)
      await Excel.run(async (ctx) => {
        const ws = snap.sheet ? ctx.workbook.worksheets.getItem(snap.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
        const r = ws.getRange(snap.address);
        r.formulas = snap.formulas;
        r.numberFormat = snap.numberFormat;
        await ctx.sync();
      });
      // Bước 2: định dạng ô và độ rộng cột (cố gắng hết mức)
      await Excel.run(async (ctx) => {
        const ws = snap.sheet ? ctx.workbook.worksheets.getItem(snap.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
        const r = ws.getRange(snap.address);
        r.setCellProperties(settableProps(snap.props));
        if (snap.columns) {
          r.setColumnProperties(
            snap.columns.map((c: any) => (c?.format?.columnWidth ? { format: { columnWidth: c.format.columnWidth } } : {}))
          );
        }
        await ctx.sync();
      }).catch(() => undefined);
    },
  };
}

