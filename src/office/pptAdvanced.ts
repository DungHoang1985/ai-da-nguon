/* global PowerPoint */
/** Công cụ PowerPoint nâng cao: biểu đồ gốc (sửa được dữ liệu), sơ đồ quy trình tự dàn trang, di chuyển/nhân bản/xóa slide, sửa và xóa hình. */
import { ToolSpec, UndoEntry } from "../core/agent/types";
import { nextTableFont, tableFits } from "../core/tablefit";
import { currentSlideId, deleteShapeUndo, findShape, remapSlide, restoreFont, shapeOn, slideAt, supports } from "./pptCore";

const slideProp = { type: "number" };
const shapeProp = { type: "string", description: "Name, or number on the slide." };

export const PPT_ADV_SPECS: ToolSpec[] = [
  {
    name: "ppt_add_chart",
    description:
      "Insert a NATIVE chart (editable, theme colours) on a NEW slide; the slide title states the takeaway.",
    parameters: {
      type: "object",
      properties: {
        chart_type: { type: "string", enum: ["column", "stacked_column", "bar", "stacked_bar", "line", "area", "pie", "doughnut"] },
        categories: { type: "array", items: { type: "string" } },
        series: {
          type: "array",
          description: "Each with as many values as categories.",
          items: {
            type: "object",
            properties: { name: { type: "string" }, values: { type: "array", items: { type: "number" } } },
            required: ["name", "values"],
          },
        },
        title: { type: "string", description: "Slide title, e.g. 'Doanh thu tăng 32% năm 2026'." },
        chart_title: { type: "string" },
        after_slide: { type: "number", description: "Default: last slide." },
        number_format: { type: "string", description: "e.g. '#,##0', '0.0%'." },
        show_values: { type: "boolean", description: "Data labels." },
        layout: { type: "string", description: "Default 'Title Only'." },
      },
      required: ["chart_type", "categories", "series"],
    },
    write: true,
  },
  {
    name: "ppt_add_process_flow",
    description:
      "Draw a process or timeline diagram from ordered steps on an existing slide, laid out automatically. chevron (arrow band), boxes (joined by arrows), timeline (steps like '03/2026: Khởi công').",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        steps: { type: "array", items: { type: "string" }, description: "2–8 short labels." },
        style: { type: "string", enum: ["chevron", "boxes", "timeline"] },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        fill_color: { type: "string", description: "Default: theme colour." },
        font_color: { type: "string" },
        font_size: { type: "number" },
      },
      required: ["slide", "steps"],
    },
    write: true,
  },
  {
    name: "ppt_move_slide",
    description: "Move a slide to a new position.",
    parameters: { type: "object", properties: { slide: slideProp, to: { type: "number", description: "New position." } }, required: ["slide", "to"] },
    write: true,
  },
  {
    name: "ppt_duplicate_slide",
    description: "Duplicate a slide (copy right after it).",
    parameters: { type: "object", properties: { slide: slideProp }, required: ["slide"] },
    write: true,
  },
  {
    name: "ppt_delete_slide",
    description: "Delete a slide.",
    parameters: { type: "object", properties: { slide: slideProp }, required: ["slide"] },
    write: true,
  },
  {
    name: "ppt_update_shape",
    description:
      "Change a shape's position/size (points), fill, outline, text format. Colours #RRGGBB; 'none' removes fill or outline.",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        shape: shapeProp,
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        fill_color: { type: "string" },
        line_color: { type: "string" },
        line_weight: { type: "number" },
        font_color: { type: "string" },
        font_size: { type: "number" },
        bold: { type: "boolean" },
        font_name: { type: "string" },
        text_align: { type: "string", enum: ["left", "center", "right", "justify"] },
        vertical_align: { type: "string", enum: ["top", "middle", "bottom"] },
      },
      required: ["slide", "shape"],
    },
    write: true,
  },
  {
    name: "ppt_delete_shape",
    description: "Delete a shape from a slide.",
    parameters: { type: "object", properties: { slide: slideProp, shape: shapeProp }, required: ["slide", "shape"] },
    write: true,
  },
];

export async function pptRisk(name: string, a: any): Promise<string | null> {
  const undoable = supports("1.8") ? " (có thể Hoàn tác)" : " — KHÔNG hoàn tác được trên máy này";
  if (name === "ppt_delete_slide") return `Xóa slide ${a.slide}${undoable}.`;
  if (name === "ppt_delete_shape") return `Xóa hình "${a.shape}" trên slide ${a.slide}${undoable}.`;
  return null;
}

// ---------- Kích thước slide ----------

let cachedSize: { w: number; h: number } | null = null;

/** Kích thước slide (điểm). Đọc từ tệp xuất của slide đầu; không đọc được thì dùng 960×540. */
export async function slideSizePt(ctx: PowerPoint.RequestContext): Promise<{ w: number; h: number }> {
  if (cachedSize) return cachedSize;
  const fallback = { w: 960, h: 540 };
  if (!supports("1.8")) return fallback;
  try {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    if (!slides.items.length) return fallback;
    const b64 = slides.items[0].exportAsBase64();
    await ctx.sync();
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(b64.value, { base64: true });
    const xml = (await zip.file("ppt/presentation.xml")?.async("string")) || "";
    const tag = /<p:sldSz\b[^>]*>/.exec(xml)?.[0] || "";
    const cx = Number(/\bcx="(\d+)"/.exec(tag)?.[1]);
    const cy = Number(/\bcy="(\d+)"/.exec(tag)?.[1]);
    if (cx > 0 && cy > 0) cachedSize = { w: Math.round(cx / 12700), h: Math.round(cy / 12700) };
    return cachedSize || fallback;
  } catch {
    return fallback;
  }
}

// ---------- Bảng vừa slide (22/11/2026, tasks/sua-22-11.md) ----------

/** Đặt cỡ chữ (nếu có `fontSize`) rồi tự co chữ bảng cho tới khi đáy bảng lọt slide, tối thiểu 10 pt.
 *  Cỡ chữ đặt từng ô (PowerPointApi 1.9): giữ kiểu bảng và phông chủ đề của mẫu — `uniformCellProperties` làm mất kiểu bảng (đã thăm dò).
 *  Dùng chung cho ppt_add_table và ppt_edit_table (thêm hàng). Trả chiều cao thật (`shape.height`), lọt hay không, ghi chú cho model. */
export async function fitTableToSlide(
  ctx: PowerPoint.RequestContext,
  shp: PowerPoint.Shape,
  o: { rows: number; cols: number; top: number; slideH: number; fontSize?: any },
): Promise<{ height: number; fits: boolean; font: number | null; notes: string[] }> {
  shp.load("height");
  await ctx.sync();
  const canFont = supports("1.9");
  const want = Number(o.fontSize);
  let font: number | null = null;
  const setFont = async (s: number) => {
    const t = shp.getTable();
    for (let r = 0; r < o.rows; r++) for (let k = 0; k < o.cols; k++) t.getCellOrNullObject(r, k).font.size = s;
    shp.load("height");
    await ctx.sync();
    font = s;
  };
  const notes: string[] = [];
  if (want && !(want >= 6 && want <= 60)) notes.push(`bỏ qua font_size ${o.fontSize} (chỉ nhận 6–60)`);
  else if (want && canFont) await setFont(want);
  else if (want) notes.push("PowerPoint trên máy này chưa cho đặt cỡ chữ ô bảng (cần PowerPointApi 1.9)");
  let startFont: number | null = font;
  if (canFont && startFont === null) {
    const c = shp.getTable().getCellOrNullObject(Math.min(1, o.rows - 1), 0);
    c.font.load("size");
    await ctx.sync();
    startFont = Number(c.font.size) || 18;
  }
  if (canFont && !tableFits(o.top, shp.height, o.slideH)) {
    for (let s = nextTableFont(startFont ?? 18); s !== null && !tableFits(o.top, shp.height, o.slideH); s = nextTableFont(s)) await setFont(s);
    if (font !== null && startFont !== null && font < startFont) notes.push(`đã tự co chữ từ ${startFont} xuống ${font} pt cho vừa slide`);
  }
  const fits = tableFits(o.top, shp.height, o.slideH);
  if (!fits)
    notes.push(
      `BẢNG VẪN TRÀN: đáy bảng ở ${Math.round(o.top + shp.height)} pt, slide cao ${Math.round(o.slideH)} pt${canFont ? ` dù chữ đã ${font ?? startFont} pt` : ""}. Đừng xóa dựng lại cùng nội dung: kéo bảng lên (top nhỏ hơn), bớt chữ, hoặc chia bảng sang hai slide.`,
    );
  return { height: shp.height, fits, font, notes };
}

// ---------- Chèn slide từ base64 ----------

async function slideIds(ctx: PowerPoint.RequestContext): Promise<string[]> {
  const s = ctx.presentation.slides;
  s.load("items/id");
  await ctx.sync();
  return s.items.map((x) => x.id);
}

/** Chèn các slide trong tệp base64 vào vị trí index0 (0-based); trả về id các slide mới. */
async function insertAt(
  ctx: PowerPoint.RequestContext,
  b64: string,
  index0: number,
  formatting: "UseDestinationTheme" | "KeepSourceFormatting" = "UseDestinationTheme"
): Promise<string[]> {
  const before = await slideIds(ctx);
  const opts: PowerPoint.InsertSlideOptions = { formatting };
  const at = Math.max(0, Math.min(index0, before.length));
  if (at > 0) opts.targetSlideId = before[at - 1];
  ctx.presentation.insertSlidesFromBase64(b64, opts);
  await ctx.sync();
  const after = await slideIds(ctx);
  const fresh = after.filter((id) => !before.includes(id));
  if (supports("1.8") && fresh.length) {
    const pos = after.indexOf(fresh[0]);
    if (pos !== at) {
      fresh.forEach((id, k) => ctx.presentation.slides.getItem(id).moveTo(at + k));
      await ctx.sync();
    }
  }
  return fresh;
}

function deleteSlidesUndo(ids: string[], label: string): UndoEntry {
  return {
    label,
    undo: () =>
      PowerPoint.run(async (c) => {
        ids.forEach((id) => c.presentation.slides.getItem(currentSlideId(id)).delete());
        await c.sync();
      }),
  };
}

/** Hoàn tác bằng bản sao slide: chèn lại bản sao vào đúng chỗ, xóa slide hiện tại (nếu còn), ghi nhận id mới. */
export function restoreSlideUndo(b64: string, oldId: string, index0: number, removeCurrent: boolean, label: string): UndoEntry {
  return {
    label,
    undo: () =>
      PowerPoint.run(async (c) => {
        const ids = await slideIds(c);
        const cur = currentSlideId(oldId);
        const pos = ids.indexOf(cur);
        const at = removeCurrent && pos >= 0 ? pos : Math.min(index0, ids.length);
        const fresh = await insertAt(c, b64, at);
        if (removeCurrent && pos >= 0) {
          c.presentation.slides.getItem(cur).delete();
          await c.sync();
        }
        if (fresh[0]) remapSlide(cur, fresh[0]);
      }),
  };
}

export async function exportSlide(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<string | null> {
  if (!supports("1.8")) return null;
  try {
    const r = slide.exportAsBase64();
    await ctx.sync();
    return r.value;
  } catch {
    return null;
  }
}

// ---------- Biểu đồ gốc ----------

/** Kiểu ô giữ chỗ là tiêu đề (không tính Subtitle — soát mã Đợt 7, #15) */
const TITLE_PH = /^(Title|CenterTitle|VerticalTitle)$/;
const THEME_COLORS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

export interface ChartLayout {
  /** Tên bố cục trong bài đang mở */
  name: string;
  /** Kiểu bố cục theo Office.js (TitleOnly, Object…) — PowerPoint ghép bố cục theo kiểu khi chèn với "Use Destination Theme" */
  type?: string;
  /** Vị trí ô tiêu đề của bố cục đó (điểm) */
  title?: { left: number; top: number; width: number; height: number };
}

/** Chọn bố cục "Chỉ tiêu đề" của bài đang mở cho slide biểu đồ, kèm vị trí ô tiêu đề để đặt biểu đồ ngay bên dưới. */
async function pickChartLayout(ctx: PowerPoint.RequestContext, wanted?: string): Promise<ChartLayout> {
  try {
    const layouts = ctx.presentation.slideMasters.getItemAt(0).layouts;
    layouts.load("items/name");
    await ctx.sync();
    let types: string[] = [];
    if (supports("1.8")) {
      try {
        layouts.load("items/type");
        await ctx.sync();
        types = layouts.items.map((l) => String(l.type));
      } catch {
        types = [];
      }
    }
    const byType = (t: string) => layouts.items.find((_, i) => types[i] === t);
    const find = (re: RegExp) => layouts.items.find((l) => re.test(l.name.trim()));
    const lay =
      (wanted ? layouts.items.find((l) => l.name.trim().toLowerCase() === String(wanted).trim().toLowerCase()) : undefined) ||
      byType("TitleOnly") ||
      find(/^(title only|chỉ tiêu đề|chỉ có tiêu đề)$/i) ||
      find(/title only|chỉ (có )?tiêu đề/i) ||
      byType("Object") ||
      find(/^title and content$|^tiêu đề và nội dung$/i);
    if (!lay) return { name: "Title Only", type: "TitleOnly" };
    const type = types[layouts.items.indexOf(lay)] || (/title only|chỉ (có )?tiêu đề/i.test(lay.name) ? "TitleOnly" : undefined);
    try {
      lay.shapes.load("items/name,items/left,items/top,items/width,items/height");
      await ctx.sync();
      const t = lay.shapes.items.find((x) => /title|tiêu đề/i.test(x.name) && !/sub|phụ/i.test(x.name));
      return { name: lay.name, type, title: t ? { left: t.left, top: t.top, width: t.width, height: t.height } : undefined };
    } catch {
      return { name: lay.name, type };
    }
  } catch {
    return { name: "Title Only", type: "TitleOnly" };
  }
}

/** Ô tiêu đề là placeholder thật, không tự đặt vị trí/kiểu → nhận đúng kiểu tiêu đề của mẫu đang mở. */
function inheritTitle(slideXml: string): string {
  return slideXml.replace(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="title"[\s\S]*?<\/p:sp>/, (sp) =>
    sp
      .replace(/<p:ph\b[^>]*\/>/, '<p:ph type="title"/>')
      .replace(/<p:spPr>[\s\S]*?<\/p:spPr>/, "<p:spPr/>")
      .replace(/<a:pPr indent="0" marL="0"><a:buNone\/><\/a:pPr>/g, "")
  );
}

export async function buildChartPptx(a: any, size: { w: number; h: number }, layout?: ChartLayout): Promise<string> {
  const PptxGenJS = (await import("pptxgenjs")).default;
  const pptx = new PptxGenJS();
  const wIn = size.w / 72;
  const hIn = size.h / 72;
  pptx.defineLayout({ name: "AIDN", width: wIn, height: hIn });
  pptx.layout = "AIDN";
  const m = 0.5;
  const tb = layout?.title;
  const titleBox = tb
    ? { x: tb.left / 72, y: tb.top / 72, w: tb.width / 72, h: tb.height / 72 }
    : { x: m, y: 0.3, w: wIn - 2 * m, h: 0.9 };
  const layoutName = layout?.name || "Title Only";
  pptx.defineSlideMaster({
    title: layoutName,
    objects: [{ placeholder: { options: { name: "title", type: "title", ...titleBox }, text: "" } }],
  } as any);
  const s = pptx.addSlide({ masterName: layoutName });
  let top = m;
  if (a.title) {
    // KHÔNG gắn lang "vi-VN": chủ đề Office có phông riêng cho chữ tiếng Việt (<a:font script="Viet" typeface="Times New Roman"/>
    // trong majorFont / minorFont) nên chữ gắn vi-VN ra Times New Roman thay cho phông tiêu đề của mẫu, và Office.js đặt lại
    // đúng tên phông chủ đề thì PowerPoint bỏ qua (đo 18/11/2026, tham-do-8-phong-bieu-do.js). Mặc định en-US: theo phông của mẫu.
    s.addText(String(a.title), { placeholder: "title" } as any);
    top = titleBox.y + titleBox.h + 0.15;
  }
  const kind = String(a.chart_type || "column");
  const cats: string[] = (a.categories || []).map((c: any) => String(c));
  if (!cats.length) throw new Error("Thiếu categories (nhãn danh mục).");
  const series: any[] = Array.isArray(a.series) ? a.series : [];
  if (!series.length) throw new Error("Thiếu series (chuỗi số liệu).");
  const data = series.map((sr, i) => {
    const vals = (Array.isArray(sr.values) ? sr.values : []).map((v: any) => Number(v) || 0);
    while (vals.length < cats.length) vals.push(0);
    return { name: String(sr.name || `Chuỗi ${i + 1}`), labels: cats, values: vals.slice(0, cats.length) };
  });
  const pie = kind === "pie" || kind === "doughnut";
  const type = pie ? kind : kind === "line" ? "line" : kind === "area" ? "area" : "bar";
  const fmt = a.number_format || (pie ? "0%" : "#,##0");
  const left = tb ? titleBox.x : m;
  const width = tb ? titleBox.w : wIn - 2 * m;
  const opts: any = {
    x: left,
    y: top,
    w: width,
    h: Math.max(1.5, hIn - top - 0.45),
    // Một chuỗi cột/thanh: một màu (màu theo đối tượng, không theo từng cột). pptxgenjs ghi màu từng cột bằng srgbClr
    // nên tên màu chủ đề ở đó làm tệp không hợp lệ (PowerPoint báo InvalidArgument khi chèn).
    chartColors: type === "bar" && data.length === 1 ? THEME_COLORS.slice(0, 1) : THEME_COLORS,
    showLegend: pie || data.length > 1,
    legendPos: "b",
    legendColor: "tx1",
    catAxisLabelColor: "tx1",
    valAxisLabelColor: "tx1",
    dataLabelColor: pie ? "bg1" : "tx1",
    valAxisLabelFormatCode: a.number_format || "#,##0",
    valGridLine: { color: "D9D9D9", size: 0.5 },
    catGridLine: { style: "none" },
  };
  if (a.chart_title) Object.assign(opts, { showTitle: true, title: String(a.chart_title), titleColor: "tx1" });
  if (type === "bar") {
    opts.barDir = kind.includes("bar") ? "bar" : "col";
    opts.barGrouping = kind.startsWith("stacked") ? "stacked" : "clustered";
    if (opts.barGrouping === "stacked") opts.overlap = 100;
  }
  if (type === "line") Object.assign(opts, { lineSize: 2, lineDataSymbol: "circle", lineDataSymbolSize: 6 });
  if (pie) {
    Object.assign(opts, { showPercent: !a.number_format, showValue: !!a.show_values && !!a.number_format, dataLabelFormatCode: fmt });
    if (kind === "doughnut") opts.holeSize = 55;
  } else if (a.show_values) {
    Object.assign(opts, { showValue: true, dataLabelFormatCode: fmt });
  }
  s.addChart(type as any, data, opts);
  const b64 = (await pptx.write({ outputType: "base64" })) as string;
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(b64, { base64: true });
  const f = zip.file("ppt/slides/slide1.xml");
  if (!f) return b64;
  if (a.title) zip.file("ppt/slides/slide1.xml", inheritTitle(await f.async("string")));
  // Gắn kiểu bố cục chuẩn (vd titleOnly) để PowerPoint dùng bố cục cùng kiểu của mẫu thay vì thêm bố cục mới
  const ooxml = OOXML_LAYOUT[layout?.type || "TitleOnly"];
  const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")?.async("string");
  const lf = rels && /slideLayouts\/(slideLayout\d+\.xml)/.exec(rels)?.[1];
  const lx = lf ? zip.file("ppt/slideLayouts/" + lf) : null;
  if (ooxml && lx) {
    const xml = await lx.async("string");
    if (!/<p:sldLayout\b[^>]*\btype="/.test(xml)) zip.file("ppt/slideLayouts/" + lf, xml.replace(/<p:sldLayout\b/, `<p:sldLayout type="${ooxml}"`));
  }
  return zip.generateAsync({ type: "base64", compression: "DEFLATE" });
}

/** Kiểu bố cục Office.js → giá trị sldLayout@type trong OOXML */
const OOXML_LAYOUT: Record<string, string> = {
  TitleOnly: "titleOnly",
  Title: "title",
  Object: "obj",
  Text: "tx",
  Blank: "blank",
  SectionHeader: "secHead",
  TwoObjects: "twoObj",
  Comparison: "twoTxTwoObj",
  ContentWithCaption: "objTx",
  PictureWithCaption: "picTx",
};

// ---------- Sơ đồ quy trình ----------

async function drawFlow(ctx: PowerPoint.RequestContext, a: any) {
  const size = await slideSizePt(ctx);
  const { slide } = await slideAt(ctx, a.slide);
  const steps: string[] = (a.steps || []).map((x: any) => String(x).trim()).filter(Boolean).slice(0, 8);
  if (steps.length < 2) throw new Error("Cần ít nhất 2 bước (steps).");
  const style = ["chevron", "boxes", "timeline"].includes(a.style) ? a.style : "chevron";
  const n = steps.length;
  const W = Number(a.width) || size.w - 80;
  const L = a.left !== undefined ? Number(a.left) : (size.w - W) / 2;
  const H = Number(a.height) || (style === "timeline" ? Math.min(200, size.h * 0.4) : Math.min(110, size.h * 0.2));
  const T = a.top !== undefined ? Number(a.top) : Math.round((size.h * 0.25 + size.h - 40) / 2 - H / 2);
  const fontSize = Number(a.font_size) || (n <= 4 ? 16 : n <= 6 ? 13 : 11);
  const added: PowerPoint.Shape[] = [];
  /** Phông của từng hình đọc TRƯỚC khi ghi chữ (lệnh nạp xếp hàng trước lệnh ghi): chữ có “ư” làm PowerPoint đổi sang Arial
   *  (đo 18/11/2026, tham-do-8-chu-u.js) → trả lại phông này sau khi ghi */
  const fontsBefore: { shp: PowerPoint.Shape; font: PowerPoint.ShapeFont }[] = [];
  const text = (shp: PowerPoint.Shape, t: string) => {
    const tf = shp.textFrame;
    tf.wordWrap = true;
    tf.verticalAlignment = "Middle";
    const tr = tf.textRange;
    const font = tr.font;
    font.load("name");
    fontsBefore.push({ shp, font });
    tr.text = t;
    tr.font.size = fontSize;
    if (a.font_color) tr.font.color = a.font_color;
    tr.paragraphFormat.horizontalAlignment = "Center";
  };
  const fill = (shp: PowerPoint.Shape) => {
    if (a.fill_color) shp.fill.setSolidColor(a.fill_color);
  };

  if (style === "chevron") {
    const gap = 6;
    const w = (W - gap * (n - 1)) / n;
    steps.forEach((t, i) => {
      const shp = slide.shapes.addGeometricShape(i === 0 ? "HomePlate" : "Chevron", { left: L + i * (w + gap), top: T, width: w, height: H });
      fill(shp);
      text(shp, t);
      added.push(shp);
    });
  } else if (style === "boxes") {
    const arrowW = Math.max(16, Math.min(32, W * 0.03));
    const gap = 6;
    const w = (W - (n - 1) * (arrowW + 2 * gap)) / n;
    steps.forEach((t, i) => {
      const x = L + i * (w + arrowW + 2 * gap);
      const box = slide.shapes.addGeometricShape("RoundRectangle", { left: x, top: T, width: w, height: H });
      fill(box);
      text(box, t);
      added.push(box);
      if (i < n - 1) {
        const ar = slide.shapes.addGeometricShape("RightArrow", { left: x + w + gap, top: T + H / 2 - arrowW / 2, width: arrowW, height: arrowW });
        fill(ar);
        added.push(ar);
      }
    });
  } else {
    const mid = T + H / 2;
    const line = slide.shapes.addLine("Straight", { left: L, top: mid, width: W, height: 0 });
    line.lineFormat.weight = 2;
    if (a.fill_color) line.lineFormat.color = a.fill_color;
    added.push(line);
    const slot = W / n;
    const dot = 16;
    steps.forEach((t, i) => {
      const cx = L + slot * (i + 0.5);
      const d = slide.shapes.addGeometricShape("Ellipse", { left: cx - dot / 2, top: mid - dot / 2, width: dot, height: dot });
      fill(d);
      added.push(d);
      const above = i % 2 === 0;
      const tb = slide.shapes.addTextBox("", { left: cx - slot / 2 + 4, top: above ? T : mid + dot, width: slot - 8, height: H / 2 - dot });
      const tf = tb.textFrame;
      const font = tf.textRange.font;
      font.load("name");
      fontsBefore.push({ shp: tb, font });
      tf.textRange.text = t;
      tf.wordWrap = true;
      tf.verticalAlignment = above ? "Bottom" : "Top";
      tf.textRange.font.size = fontSize;
      if (a.font_color) tf.textRange.font.color = a.font_color;
      tf.textRange.paragraphFormat.horizontalAlignment = "Center";
      added.push(tb);
    });
  }
  added.forEach((s) => s.load("id,name"));
  await ctx.sync();
  for (const x of fontsBefore) await restoreFont(ctx, x.shp, x.font.name).catch(() => 0);
  let group: { id: string; name: string } | null = null;
  if (supports("1.8")) {
    try {
      const g = slide.shapes.addGroup(added.map((s) => s.id));
      g.load("id,name");
      await ctx.sync();
      group = { id: g.id, name: g.name };
    } catch {
      group = null;
    }
  }
  return { slide, group, shapes: added.map((s) => ({ id: s.id, name: s.name })), style, n };
}

// ---------- Thực thi ----------

const ALIGN: Record<string, string> = { left: "Left", center: "Center", right: "Right", justify: "Justify" };
const VALIGN: Record<string, string> = { top: "Top", middle: "Middle", bottom: "Bottom" };

export async function executePptAdvanced(name: string, a: any, log: UndoEntry[]): Promise<any> {
  switch (name) {
    case "ppt_add_chart":
      return PowerPoint.run(async (ctx) => {
        const size = await slideSizePt(ctx);
        const ids = await slideIds(ctx);
        const after = a.after_slide !== undefined ? Math.max(0, Math.min(Number(a.after_slide), ids.length)) : ids.length;
        const layout = await pickChartLayout(ctx, a.layout);
        const b64 = await buildChartPptx(a, size, layout);
        // Bố cục có sẵn của bài TRƯỚC khi chèn: nếu PowerPoint không ghép được bố cục và tự thêm bố cục lạ (lỗi cũ "thêm một bố cục"),
        // slide biểu đồ được chuyển về bố cục của mẫu để không lệch mẫu (đã đo 16/11/2026: slide.applyLayout dùng được, PowerPointApi 1.8)
        let known: { id: string; name: string }[] = [];
        if (supports("1.8")) {
          try {
            const ls = ctx.presentation.slideMasters.getItemAt(0).layouts;
            ls.load("items/id,items/name");
            await ctx.sync();
            known = ls.items.map((l) => ({ id: l.id, name: l.name }));
          } catch {
            known = [];
          }
        }
        const fresh = await insertAt(ctx, b64, after);
        if (!fresh.length) throw new Error("PowerPoint không chèn được slide biểu đồ.");
        log.push(deleteSlidesUndo(fresh, `Slide biểu đồ ${after + 1}`));
        let usedLayout = layout.name;
        let layoutNote = "";
        if (known.length) {
          try {
            const sl = ctx.presentation.slides.getItem(fresh[0]);
            sl.layout.load("id,name");
            await ctx.sync();
            usedLayout = sl.layout.name;
            if (!known.some((k) => k.id === sl.layout.id)) {
              const want = known.find((k) => k.name === layout.name) || known.find((k) => /title only|chỉ (có )?tiêu đề/i.test(k.name)) || known.find((k) => /tiêu đề|title/i.test(k.name));
              if (want) {
                const ls = ctx.presentation.slideMasters.getItemAt(0).layouts;
                sl.applyLayout(ls.getItem(want.id));
                await ctx.sync();
                // Ô giữ chỗ rỗng của bố cục mới (ô nội dung) nằm đè lên biểu đồ: gỡ đi, giữ ô tiêu đề
                sl.shapes.load("items/id,items/name,items/type");
                await ctx.sync();
                const holders = sl.shapes.items.filter((x) => String(x.type) === "Placeholder");
                holders.forEach((x) => {
                  x.placeholderFormat.load("type");
                  x.textFrame.textRange.load("text");
                });
                await ctx.sync();
                holders.filter((x) => !TITLE_PH.test(String(x.placeholderFormat.type)) && !x.textFrame.textRange.text.trim()).forEach((x) => x.delete());
                await ctx.sync();
                usedLayout = want.name;
                layoutNote = " (PowerPoint không ghép được bố cục khi chèn nên slide đã được chuyển về bố cục này của mẫu)";
              }
            }
          } catch {
            /* không đổi được bố cục: giữ như PowerPoint đã chèn */
          }
        }
        // Tiêu đề của slide biểu đồ: slide dựng bằng pptxgenjs rồi chèn vào nên chữ tiêu đề mang phông Times New Roman, không theo mẫu
        // (đo 18/11/2026). Lấy phông của ô tiêu đề trên bố cục đang dùng và trả cho tiêu đề.
        if (supports("1.8")) {
          try {
            const sl = ctx.presentation.slides.getItem(fresh[0]);
            sl.shapes.load("items/id,items/name,items/type");
            sl.layout.shapes.load("items/id,items/name,items/type");
            await ctx.sync();
            const isTitle = async (list: PowerPoint.Shape[]) => {
              const hs = list.filter((x) => String(x.type) === "Placeholder");
              hs.forEach((x) => x.placeholderFormat.load("type"));
              await ctx.sync();
              return hs.find((x) => TITLE_PH.test(String(x.placeholderFormat.type)));
            };
            const mine = await isTitle(sl.shapes.items);
            const model = await isTitle(sl.layout.shapes.items);
            if (mine && model) {
              const f = model.textFrame.textRange.font;
              f.load("name");
              mine.load("left,top,width,height");
              model.load("left,top,width,height");
              await ctx.sync();
              await restoreFont(ctx, mine, f.name);
              // Ô tiêu đề để trống vị trí (thừa hưởng từ bố cục) thì Office.js báo toạ độ 0, 0, 0, 0 và AI tốn một lượt "đặt lại"
              // (đo 18/11/2026, tham-do-8-toa-do-b.js). Ghi rõ vị trí của ô tiêu đề trên bố cục: trên slide trông y như cũ.
              if (!(mine.width > 0 && mine.height > 0) && model.width > 0 && model.height > 0) {
                mine.left = model.left;
                mine.top = model.top;
                mine.width = model.width;
                mine.height = model.height;
                await ctx.sync();
              }
            }
          } catch {
            /* không đọc được phông của bố cục: giữ như PowerPoint đã chèn */
          }
        }
        if (supports("1.5")) {
          ctx.presentation.setSelectedSlides([fresh[0]]);
          await ctx.sync();
        }
        return {
          ok: true,
          message: `Đã chèn biểu đồ gốc (sửa được dữ liệu) ở slide ${after + 1}, bố cục "${usedLayout}" của mẫu${layoutNote}`,
          layout: usedLayout,
          slide: after + 1,
        };
      });

    case "ppt_add_process_flow":
      return PowerPoint.run(async (ctx) => {
        const r = await drawFlow(ctx, a);
        if (r.group) log.push(deleteShapeUndo(r.slide.id, r.group.id, `Sơ đồ quy trình slide ${a.slide}`, r.group.name));
        else r.shapes.forEach((x) => log.push(deleteShapeUndo(r.slide.id, x.id, `Hình của sơ đồ slide ${a.slide}`, x.name)));
        return { ok: true, message: `Đã vẽ sơ đồ ${r.style} ${r.n} bước ở slide ${a.slide}${r.group ? " (đã nhóm)" : ""}`, group: r.group?.name };
      });

    case "ppt_move_slide":
      return PowerPoint.run(async (ctx) => {
        if (!supports("1.8")) throw new Error("PowerPoint trên máy này chưa hỗ trợ di chuyển slide qua add-in.");
        const { slide, index, count } = await slideAt(ctx, a.slide);
        const to = Math.max(1, Math.min(Number(a.to) || 1, count));
        const id = slide.id;
        slide.moveTo(to - 1);
        await ctx.sync();
        log.push({
          label: `Di chuyển slide ${index} → ${to}`,
          undo: () =>
            PowerPoint.run(async (c) => {
              c.presentation.slides.getItem(currentSlideId(id)).moveTo(index - 1);
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã chuyển slide ${index} tới vị trí ${to}` };
      });

    case "ppt_duplicate_slide":
      return PowerPoint.run(async (ctx) => {
        if (!supports("1.8")) throw new Error("PowerPoint trên máy này chưa hỗ trợ nhân bản slide qua add-in.");
        const { slide, index } = await slideAt(ctx, a.slide);
        const b64 = await exportSlide(ctx, slide);
        if (!b64) throw new Error("Không xuất được slide để nhân bản.");
        const fresh = await insertAt(ctx, b64, index);
        log.push(deleteSlidesUndo(fresh, `Bản sao slide ${index}`));
        return { ok: true, message: `Đã nhân bản slide ${index} thành slide ${index + 1}`, slide: index + 1 };
      });

    case "ppt_delete_slide":
      return PowerPoint.run(async (ctx) => {
        const { slide, index, count } = await slideAt(ctx, a.slide);
        if (count <= 1) throw new Error("Không xóa slide cuối cùng còn lại.");
        const id = slide.id;
        const b64 = await exportSlide(ctx, slide);
        slide.delete();
        await ctx.sync();
        if (b64) log.push(restoreSlideUndo(b64, id, index - 1, false, `Xóa slide ${index}`));
        return { ok: true, message: `Đã xóa slide ${index}${b64 ? "" : " (không hoàn tác được)"}` };
      });

    case "ppt_update_shape":
      return PowerPoint.run(async (ctx) => {
        const { slide } = await slideAt(ctx, a.slide);
        const shp = await shapeOn(ctx, slide, a.shape);
        shp.load("id,name,left,top,width,height");
        shp.fill.load("type,foregroundColor");
        shp.lineFormat.load("visible,color,weight");
        await ctx.sync();
        const prev: any = {
          left: shp.left,
          top: shp.top,
          width: shp.width,
          height: shp.height,
          fillType: shp.fill.type,
          fillColor: shp.fill.foregroundColor,
          line: { visible: shp.lineFormat.visible, color: shp.lineFormat.color, weight: shp.lineFormat.weight },
        };
        const wantsText = ["font_color", "font_size", "bold", "font_name", "text_align", "vertical_align"].some((k) => a[k] !== undefined);
        if (wantsText) {
          try {
            const tr = shp.textFrame.textRange;
            tr.font.load("color,size,bold,name");
            tr.paragraphFormat.load("horizontalAlignment");
            shp.textFrame.load("verticalAlignment");
            await ctx.sync();
            prev.font = { color: tr.font.color, size: tr.font.size, bold: tr.font.bold, name: tr.font.name };
            prev.align = tr.paragraphFormat.horizontalAlignment;
            prev.valign = shp.textFrame.verticalAlignment;
          } catch {
            throw new Error(`Hình "${shp.name}" không có chữ để định dạng.`);
          }
        }
        for (const k of ["left", "top", "width", "height"]) if (a[k] !== undefined) (shp as any)[k] = Number(a[k]);
        if (a.fill_color !== undefined) {
          if (String(a.fill_color).toLowerCase() === "none") shp.fill.clear();
          else shp.fill.setSolidColor(String(a.fill_color));
        }
        if (a.line_color !== undefined) {
          if (String(a.line_color).toLowerCase() === "none") shp.lineFormat.visible = false;
          else {
            shp.lineFormat.visible = true;
            shp.lineFormat.color = String(a.line_color);
          }
        }
        if (a.line_weight !== undefined) shp.lineFormat.weight = Number(a.line_weight);
        if (wantsText) {
          const tr = shp.textFrame.textRange;
          if (a.font_color !== undefined) tr.font.color = String(a.font_color);
          if (a.font_size !== undefined) tr.font.size = Number(a.font_size);
          if (a.bold !== undefined) tr.font.bold = !!a.bold;
          if (a.font_name !== undefined) tr.font.name = String(a.font_name);
          if (a.text_align && ALIGN[a.text_align]) tr.paragraphFormat.horizontalAlignment = ALIGN[a.text_align] as any;
          if (a.vertical_align && VALIGN[a.vertical_align]) shp.textFrame.verticalAlignment = VALIGN[a.vertical_align] as any;
        }
        await ctx.sync();
        const slideId = slide.id;
        const shapeId = shp.id;
        const shapeName = shp.name;
        log.push({
          label: `Sửa hình ${shp.name}`,
          undo: () =>
            PowerPoint.run(async (c) => {
              const s = await findShape(c, slideId, shapeId, shapeName);
              s.left = prev.left;
              s.top = prev.top;
              s.width = prev.width;
              s.height = prev.height;
              if (a.fill_color !== undefined) {
                if (prev.fillType === "Solid" && prev.fillColor) s.fill.setSolidColor(prev.fillColor);
                else if (prev.fillType === "NoFill") s.fill.clear();
              }
              if (a.line_color !== undefined || a.line_weight !== undefined) {
                s.lineFormat.visible = !!prev.line.visible;
                if (prev.line.visible) {
                  if (prev.line.color) s.lineFormat.color = prev.line.color;
                  if (prev.line.weight) s.lineFormat.weight = prev.line.weight;
                }
              }
              if (prev.font) {
                const tr = s.textFrame.textRange;
                if (a.font_color !== undefined && prev.font.color) tr.font.color = prev.font.color;
                if (a.font_size !== undefined && prev.font.size) tr.font.size = prev.font.size;
                if (a.bold !== undefined && prev.font.bold !== null) tr.font.bold = prev.font.bold;
                if (a.font_name !== undefined && prev.font.name) tr.font.name = prev.font.name;
                if (a.text_align && prev.align) tr.paragraphFormat.horizontalAlignment = prev.align;
                if (a.vertical_align && prev.valign) s.textFrame.verticalAlignment = prev.valign;
              }
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã sửa hình ${shp.name} ở slide ${a.slide}` };
      });

    case "ppt_delete_shape":
      return PowerPoint.run(async (ctx) => {
        const { slide, index } = await slideAt(ctx, a.slide);
        const shp = await shapeOn(ctx, slide, a.shape);
        const b64 = await exportSlide(ctx, slide);
        const nm = shp.name;
        shp.delete();
        await ctx.sync();
        if (b64) log.push(restoreSlideUndo(b64, slide.id, index - 1, true, `Xóa hình ${nm}`));
        return { ok: true, message: `Đã xóa hình ${nm} ở slide ${index}${b64 ? "" : " (không hoàn tác được)"}` };
      });
  }
  return undefined;
}
