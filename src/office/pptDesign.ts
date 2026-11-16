/* global PowerPoint */
/** Công cụ PowerPoint của Đợt 3: ảnh slide cho AI tự nhìn, kiểm bài theo mẫu, đọc mẫu (màu, phông, bố cục, ô giữ chỗ),
 *  định dạng chữ theo đoạn / cụm chữ, sửa bảng có sẵn, đọc vùng đang chọn, đổi bố cục slide.
 *  Các API dùng ở đây đã thăm dò trong PowerPoint 16/11/2026 (kiem-thu/cdp-bai/tham-do-3-ppt.js, bảng ở tasks/dot-3.md).
 *  Không có API cho ghi chú diễn giả và phông chủ đề: phông chủ đề đọc từ gói tệp của một slide xuất ra (chỉ đọc). */
import { ToolSpec, UndoEntry } from "../core/agent/types";
import { DeckPara, DeckShape, DeckSlide, checkDeck, deckBrief } from "../core/deckcheck";
import { exportSlide, restoreSlideUndo, slideSizePt } from "./pptAdvanced";
import { shapeOn, slideAt, supports } from "./pptCore";
import { truncate } from "./util";

const slideProp = { type: "number", description: "Slide number (1-based)." };
const shapeProp = { type: "string", description: "Shape name (from the overview) or its number on the slide." };

export const PPT_DESIGN_SPECS: ToolSpec[] = [
  {
    name: "ppt_get_slide_image",
    description:
      "Take a picture of one slide so you can LOOK at it (text overflowing its box, overlapping shapes, crowded layout), plus a code-based measurement of that slide. Call it after creating or changing a slide to check the result, fix what is wrong, then look once more. Only useful with a model that accepts images; otherwise you get the measurements only.",
    parameters: { type: "object", properties: { slide: slideProp }, required: ["slide"] },
  },
  {
    name: "ppt_check_deck",
    description:
      "Check the whole deck against its template by code (no image needed): fonts that differ from the template fonts, text smaller than the minimum size, slides without a title, text estimated to overflow its box, shapes outside the slide, overlapping text boxes, leftover placeholder text. Returns every finding with slide number and shape name, and how to fix it.",
    parameters: {
      type: "object",
      properties: {
        min_font_size: { type: "number", description: "Smallest acceptable text size in points (default 12)." },
        font: { type: "string", description: "Required font, if the user names one (default: the template's theme fonts)." },
        slides: { type: "array", items: { type: "number" }, description: "Only these slide numbers (default: all, max 60)." },
      },
    },
  },
  {
    name: "ppt_get_theme",
    description: "Read the template: theme colours (accent 1–6, dark, light), theme fonts (headings, body), slide size, and every layout with its placeholders. Use before adding slides or shapes so new content follows the template.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "ppt_format_text",
    description:
      "Format part of the text in a shape: one paragraph (paragraph = 1-based line number in the shape), or every occurrence of `find`, or the whole shape when neither is given. Font options, alignment and bullets.",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        shape: shapeProp,
        paragraph: { type: "number", description: "1-based paragraph (line) of the shape's text." },
        find: { type: "string", description: "Format every occurrence of this text inside the shape." },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_size: { type: "number" },
        font_name: { type: "string" },
        font_color: { type: "string", description: "Hex colour like #C00000; prefer theme colours from ppt_get_theme." },
        alignment: { type: "string", enum: ["Left", "Center", "Right", "Justify"] },
        bullet: { type: "boolean", description: "Show or hide the bullet of the paragraph(s)." },
      },
      required: ["slide", "shape"],
    },
    write: true,
  },
  {
    name: "ppt_edit_table",
    description: "Edit an existing table on a slide: set the text of given cells (0-based row and column; row 0 is the header) and optionally append rows. Read the table first with ppt_get_overview (tables include their values).",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        shape: shapeProp,
        cells: {
          type: "array",
          items: { type: "object", properties: { row: { type: "number" }, col: { type: "number" }, text: { type: "string" } }, required: ["row", "col", "text"] },
        },
        add_rows: { type: "array", items: { type: "array", items: { type: "string" } }, description: "Rows to append at the bottom." },
      },
      required: ["slide", "shape"],
    },
    write: true,
  },
  {
    name: "ppt_get_selection",
    description: "Read what the user has selected in PowerPoint right now: the selected slide(s), shape(s) and text. Use when the user says 'this slide', 'this shape', 'the selected text'.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "ppt_set_layout",
    description: "Change the layout of an existing slide to another layout of the template (names from ppt_get_theme or the overview), e.g. to give a slide a title placeholder.",
    parameters: { type: "object", properties: { slide: slideProp, layout: { type: "string", description: "Layout name." } }, required: ["slide", "layout"] },
    write: true,
  },
];

const TITLE_TYPES = new Set(["Title", "CenterTitle", "VerticalTitle"]);
const TEXTY = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);
const THEME_COLOR_NAMES = ["Dark1", "Light1", "Dark2", "Light2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"];

/** Tách chữ của một khung thành các đoạn kèm vị trí (PowerPoint ngắt đoạn bằng \r, \n hoặc \v) */
export function splitParagraphs(text: string): { text: string; start: number }[] {
  const out: { text: string; start: number }[] = [];
  const re = /\r\n|\r|\n|\v/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    out.push({ text: text.slice(last, m.index), start: last });
    last = m.index + m[0].length;
  }
  out.push({ text: text.slice(last), start: last });
  return out;
}

/** Đọc các slide thành dữ liệu cho phép kiểm bằng mã: hình, khung, từng đoạn chữ với phông / cỡ / màu */
async function readSlides(ctx: PowerPoint.RequestContext, only?: number[]): Promise<DeckSlide[]> {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const picks = (only && only.length ? only.map((n) => Number(n) - 1) : slides.items.map((_, i) => i)).filter((i) => Number.isInteger(i) && i >= 0 && i < slides.items.length).slice(0, 60);
  const out: DeckSlide[] = [];
  for (const i of picks) {
    const s = slides.items[i];
    s.layout.load("name");
    s.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await ctx.sync();
    const shapes = s.shapes.items;
    const texty = shapes.filter((sh) => TEXTY.has(String(sh.type)));
    const holders = shapes.filter((sh) => String(sh.type) === "Placeholder");
    holders.forEach((sh) => sh.placeholderFormat.load("type"));
    const frames = new Map<string, { text: string; auto: string; wrap: boolean }>();
    const reads = texty.map((sh) => {
      sh.textFrame.load("autoSizeSetting,wordWrap,hasText");
      const tr = sh.textFrame.textRange;
      tr.load("text");
      return { sh, tr };
    });
    try {
      await ctx.sync();
      reads.forEach(({ sh, tr }) => frames.set(sh.id, { text: tr.text || "", auto: String(sh.textFrame.autoSizeSetting), wrap: !!sh.textFrame.wordWrap }));
    } catch {
      // Một hình không có khung chữ làm hỏng cả lô: đọc từng hình
      for (const sh of texty) {
        try {
          sh.textFrame.load("autoSizeSetting,wordWrap");
          const tr = sh.textFrame.textRange;
          tr.load("text");
          await ctx.sync();
          frames.set(sh.id, { text: tr.text || "", auto: String(sh.textFrame.autoSizeSetting), wrap: !!sh.textFrame.wordWrap });
        } catch {
          /* hình không có chữ */
        }
      }
      for (const sh of holders) {
        try {
          sh.placeholderFormat.load("type");
          await ctx.sync();
        } catch {
          /* bỏ qua */
        }
      }
    }
    // Phông, cỡ, màu của từng đoạn
    const subs: { id: string; text: string; r: PowerPoint.TextRange }[] = [];
    for (const sh of texty) {
      const f = frames.get(sh.id);
      if (!f || !f.text.trim()) continue;
      for (const p of splitParagraphs(f.text)) {
        if (!p.text.trim()) continue;
        const r = sh.textFrame.textRange.getSubstring(p.start, p.text.length);
        r.font.load("name,size,color");
        subs.push({ id: sh.id, text: p.text, r });
      }
    }
    if (subs.length) await ctx.sync();
    const paras = new Map<string, DeckPara[]>();
    for (const x of subs) {
      const list = paras.get(x.id) || [];
      const size = typeof x.r.font.size === "number" && x.r.font.size > 0 ? x.r.font.size : null;
      list.push({ text: x.text, font: x.r.font.name ? String(x.r.font.name) : null, size, color: x.r.font.color || null });
      paras.set(x.id, list);
    }
    const list: DeckShape[] = shapes.map((sh, k) => {
      let isTitle = false;
      try {
        isTitle = String(sh.type) === "Placeholder" && TITLE_TYPES.has(String(sh.placeholderFormat.type));
      } catch {
        isTitle = false;
      }
      const f = frames.get(sh.id);
      return { name: sh.name, n: k + 1, type: String(sh.type), isTitle, box: [sh.left, sh.top, sh.width, sh.height], paragraphs: paras.get(sh.id) || (f ? [] : undefined), autoSize: f?.auto, wordWrap: f?.wrap };
    });
    out.push({ slide: i + 1, layout: s.layout.name, shapes: list });
  }
  return out;
}

async function themeColors(ctx: PowerPoint.RequestContext): Promise<Record<string, string>> {
  if (!supports("1.10")) return {};
  try {
    const m = ctx.presentation.slideMasters.getItemAt(0);
    const rs = THEME_COLOR_NAMES.map((n) => m.themeColorScheme.getThemeColor(n as any));
    await ctx.sync();
    return Object.fromEntries(THEME_COLOR_NAMES.map((n, i) => [n, rs[i].value]));
  } catch {
    return {};
  }
}

let fontCache: { at: number; fonts: { heading?: string; body?: string } } | null = null;
/** Phông chủ đề: không có API → đọc từ tệp chủ đề trong gói của một slide xuất ra (đã đo: exportAsBase64 trả gói .pptx) */
async function themeFonts(ctx: PowerPoint.RequestContext): Promise<{ heading?: string; body?: string }> {
  if (fontCache && Date.now() - fontCache.at < 60000) return fontCache.fonts;
  const fonts: { heading?: string; body?: string } = {};
  try {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    if (!slides.items.length) return fonts;
    const b64 = await exportSlide(ctx, slides.items[0]);
    if (!b64) return fonts;
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(b64, { base64: true });
    const name = Object.keys(zip.files)
      .filter((f) => /^ppt\/theme\/theme\d*\.xml$/.test(f))
      .sort()[0];
    if (!name) return fonts;
    const xml = await zip.files[name].async("string");
    const pick = (tag: string) => (xml.match(new RegExp(`<a:${tag}>[\\s\\S]*?<a:latin[^>]*typeface="([^"]+)"`)) || [])[1];
    fonts.heading = pick("majorFont");
    fonts.body = pick("minorFont");
    fontCache = { at: Date.now(), fonts };
  } catch {
    /* không đọc được: bỏ qua */
  }
  return fonts;
}

async function deckOptions(ctx: PowerPoint.RequestContext, a: any) {
  const size = await slideSizePt(ctx);
  const tf = a?.font ? { heading: String(a.font), body: String(a.font) } : await themeFonts(ctx);
  const colors = await themeColors(ctx);
  const fonts = [...new Set([tf.heading, tf.body].filter(Boolean) as string[])];
  return { size, fonts, colors };
}

/** Bản sao slide trước khi sửa → Hoàn tác bằng cách chèn lại bản sao */
async function snapshotUndo(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide, index: number, label: string, log: UndoEntry[]) {
  const b64 = await exportSlide(ctx, slide);
  slide.load("id");
  await ctx.sync();
  if (b64) log.push(restoreSlideUndo(b64, slide.id, index - 1, true, label));
  return !!b64;
}

export async function executePptDesign(name: string, a: any, log: UndoEntry[]): Promise<any> {
  switch (name) {
    case "ppt_get_slide_image":
      return PowerPoint.run(async (ctx) => {
        const { slide, index } = await slideAt(ctx, a.slide);
        const o = await deckOptions(ctx, {});
        const data = await readSlides(ctx, [index]);
        const rep = checkDeck(data, { slideW: o.size.w, slideH: o.size.h, themeFonts: o.fonts, minSize: 12 });
        const layout = rep.issues.filter((i) => ["overflow", "off_slide", "overlap"].includes(i.kind));
        let image: string | null = null;
        let note = "";
        if (!supports("1.8")) note = "PowerPoint trên máy này chưa chụp được ảnh slide (cần PowerPointApi 1.8): chỉ có số đo bằng mã.";
        else {
          try {
            const img = slide.getImageAsBase64({ width: 960 });
            await ctx.sync();
            image = img.value;
          } catch (e: any) {
            note = `Không chụp được ảnh slide (${e?.message || e}): chỉ có số đo bằng mã.`;
          }
        }
        return {
          ok: true,
          slide: index,
          brief: `${image ? "đã chụp ảnh slide " + index : "không có ảnh"}${layout.length ? ` · số đo: ${layout.length} điểm bố cục cần xem` : " · số đo: không thấy tràn, chồng"}`,
          message: image
            ? `Ảnh slide ${index} được gửi kèm ngay sau kết quả này: hãy NHÌN ảnh để xem chữ có tràn khung, hình có chồng nhau, bố cục có cân không. Số đo bằng mã ở measured (chữ tràn chỉ là ước lượng — tin ảnh hơn).`
            : note,
          ...(note && image ? { note } : {}),
          measured: layout.map((i) => i.message),
          shapes: data[0]?.shapes.map((s) => ({ n: s.n, name: s.name, ...(s.isTitle ? { title: true } : {}), box: s.box.map((v) => Math.round(v)), ...(s.paragraphs?.length ? { text: truncate(s.paragraphs.map((p) => p.text).join(" / "), 120), size: s.paragraphs[0].size } : {}) })),
          ...(image ? { __images: [{ mime: "image/png", data: image, label: `Ảnh slide ${index}` }] } : {}),
        };
      });

    case "ppt_check_deck":
      return PowerPoint.run(async (ctx) => {
        const o = await deckOptions(ctx, a);
        const data = await readSlides(ctx, Array.isArray(a.slides) ? a.slides : undefined);
        const min = Number(a.min_font_size) > 0 && Number(a.min_font_size) < 100 ? Number(a.min_font_size) : 12;
        const rep = checkDeck(data, { slideW: o.size.w, slideH: o.size.h, themeFonts: o.fonts, minSize: min });
        return {
          ok: true,
          brief: deckBrief(rep),
          slides_checked: rep.slides,
          template_fonts: rep.standardFonts,
          min_font_size: rep.minSize,
          problems: rep.issues.length,
          counts: rep.counts,
          issues: rep.issues.slice(0, 120).map((i) => ({ kind: i.kind, slide: i.slide, ...(i.shape ? { shape: i.shape } : {}), message: i.message, ...(i.fix ? { fix: i.fix } : {}) })),
          message: rep.issues.length ? `Có ${rep.issues.length} điểm cần xem. Chữ tràn khung là ước lượng: muốn chắc thì chụp slide đó bằng ppt_get_slide_image.` : "Không thấy điểm lệch mẫu nào bằng phép kiểm bằng mã.",
        };
      });

    case "ppt_get_theme":
      return PowerPoint.run(async (ctx) => {
        const size = await slideSizePt(ctx);
        const colors = await themeColors(ctx);
        const fonts = await themeFonts(ctx);
        const masters = ctx.presentation.slideMasters;
        masters.load("items/name");
        await ctx.sync();
        const m = masters.items[0];
        m.layouts.load("items/name");
        await ctx.sync();
        const lays = m.layouts.items.slice(0, 25);
        lays.forEach((l) => l.shapes.load("items/name,items/type"));
        await ctx.sync();
        const layouts: any[] = [];
        for (const l of lays) {
          const hs = l.shapes.items.filter((s) => String(s.type) === "Placeholder");
          hs.forEach((s) => s.placeholderFormat.load("type"));
          let types: string[] = [];
          try {
            await ctx.sync();
            types = hs.map((s) => String(s.placeholderFormat.type));
          } catch {
            types = [];
          }
          layouts.push({ name: l.name, placeholders: types.filter((t) => !["Date", "Footer", "SlideNumber"].includes(t)) });
        }
        return {
          ok: true,
          brief: `mẫu “${m.name}”: ${layouts.length} bố cục${fonts.heading ? `, phông ${fonts.heading}${fonts.body && fonts.body !== fonts.heading ? " / " + fonts.body : ""}` : ""}`,
          master: m.name,
          masters: masters.items.length,
          slide_size: { width: Math.round(size.w), height: Math.round(size.h) },
          colors,
          fonts: fonts.heading || fonts.body ? fonts : { note: "Không đọc được phông chủ đề (PowerPoint không có API cho việc này; đọc từ gói tệp không thành)." },
          layouts,
          message: "Màu: dùng Accent1–Accent6 theo thứ tự cho hình và biểu đồ, Dark1 / Light1 cho chữ và nền. Slide mới: chọn bố cục theo tên ở layouts.",
        };
      });

    case "ppt_get_selection":
      return PowerPoint.run(async (ctx) => {
        if (!supports("1.5")) throw new Error("PowerPoint trên máy này chưa cho add-in đọc vùng chọn (cần PowerPointApi 1.5).");
        const all = ctx.presentation.slides;
        all.load("items/id");
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        const shapes = ctx.presentation.getSelectedShapes();
        shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
        const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
        tr.load("text,isNullObject");
        await ctx.sync();
        const texts: Record<string, string> = {};
        for (const sh of shapes.items.filter((s) => TEXTY.has(String(s.type)))) {
          try {
            const t = sh.textFrame.textRange;
            t.load("text");
            await ctx.sync();
            texts[sh.id] = t.text;
          } catch {
            /* không có chữ */
          }
        }
        const nums = sel.items.map((x) => all.items.findIndex((y) => y.id === x.id) + 1).filter((n) => n > 0);
        return {
          ok: true,
          brief: `slide ${nums.join(", ") || "?"}${shapes.items.length ? `, ${shapes.items.length} hình` : ""}${!tr.isNullObject && tr.text ? ", có chữ đang chọn" : ""}`,
          slides: nums,
          shapes: shapes.items.map((s) => ({ name: s.name, type: String(s.type), box: [s.left, s.top, s.width, s.height].map((v) => Math.round(v)), ...(texts[s.id] !== undefined ? { text: truncate(texts[s.id], 600) } : {}) })),
          selected_text: tr.isNullObject ? null : truncate(tr.text, 2000),
          ...(shapes.items.length || !tr.isNullObject ? {} : { note: "Người dùng chưa chọn hình hay chữ nào; chỉ có slide đang mở." }),
        };
      });

    case "ppt_format_text":
      return PowerPoint.run(async (ctx) => {
        const { slide, index } = await slideAt(ctx, a.slide);
        const shp = await shapeOn(ctx, slide, a.shape);
        shp.load("name");
        const whole = shp.textFrame.textRange;
        whole.load("text");
        try {
          await ctx.sync();
        } catch {
          throw new Error(`Hình "${a.shape}" không có chữ để định dạng.`);
        }
        const text = whole.text || "";
        const targets: PowerPoint.TextRange[] = [];
        let where = "cả hình";
        if (a.paragraph !== undefined && a.paragraph !== null) {
          const ps = splitParagraphs(text);
          const k = Number(a.paragraph) - 1;
          if (!Number.isInteger(k) || k < 0 || k >= ps.length) throw new Error(`Hình có ${ps.length} đoạn; không có đoạn số ${a.paragraph}.`);
          if (!ps[k].text.length) throw new Error(`Đoạn ${a.paragraph} rỗng.`);
          targets.push(whole.getSubstring(ps[k].start, ps[k].text.length));
          where = `đoạn ${a.paragraph}`;
        } else if (a.find !== undefined && a.find !== null) {
          const f = String(a.find);
          if (!f) throw new Error("find rỗng: nêu cụm chữ cần định dạng, hoặc bỏ find để định dạng cả hình.");
          for (let at = text.indexOf(f); at >= 0 && targets.length < 50; at = text.indexOf(f, at + f.length)) targets.push(whole.getSubstring(at, f.length));
          if (!targets.length) throw new Error(`Không thấy “${truncate(f, 60)}” trong hình ${shp.name}.`);
          where = `${targets.length} chỗ “${truncate(f, 30)}”`;
        } else targets.push(whole);
        const undoable = await snapshotUndo(ctx, slide, index, `Định dạng chữ ở slide ${index}`, log);
        const skipped: string[] = [];
        const size = a.font_size === undefined || a.font_size === null ? null : Number(a.font_size);
        if (size !== null && !(size >= 4 && size <= 400)) skipped.push("font_size");
        for (const t of targets) {
          if (a.bold !== undefined) t.font.bold = !!a.bold;
          if (a.italic !== undefined) t.font.italic = !!a.italic;
          if (a.underline !== undefined) t.font.underline = (a.underline ? "Single" : "None") as any;
          if (size !== null && size >= 4 && size <= 400) t.font.size = size;
          if (a.font_name) t.font.name = String(a.font_name);
          if (a.font_color) t.font.color = String(a.font_color);
          if (a.alignment) t.paragraphFormat.horizontalAlignment = a.alignment as any;
          if (a.bullet !== undefined) t.paragraphFormat.bulletFormat.visible = !!a.bullet;
        }
        await ctx.sync();
        return { ok: true, message: `Đã định dạng ${where} của hình ${shp.name} ở slide ${index}${skipped.length ? ` (bỏ qua tham số không hợp lệ: ${skipped.join(", ")})` : ""}${undoable ? "" : " (không tạo được bản sao để Hoàn tác)"}` };
      });

    case "ppt_edit_table":
      return PowerPoint.run(async (ctx) => {
        if (!supports("1.8")) throw new Error("PowerPoint trên máy này chưa cho add-in sửa bảng (cần PowerPointApi 1.8).");
        const { slide, index } = await slideAt(ctx, a.slide);
        const shp = await shapeOn(ctx, slide, a.shape);
        shp.load("name,type");
        await ctx.sync();
        if (String(shp.type) !== "Table") throw new Error(`Hình "${shp.name}" không phải bảng (loại ${shp.type}).`);
        const t = shp.getTable();
        t.load("rowCount,columnCount");
        await ctx.sync();
        const cells = Array.isArray(a.cells) ? a.cells : [];
        const rows = Array.isArray(a.add_rows) ? a.add_rows.filter((r: any) => Array.isArray(r)) : [];
        if (!cells.length && !rows.length) throw new Error("Cần cells (ô cần sửa) hoặc add_rows (hàng thêm vào cuối).");
        if (cells.length > 300 || rows.length > 50) throw new Error("Quá nhiều ô hoặc hàng trong một lần gọi (tối đa 300 ô, 50 hàng): chia nhỏ.");
        for (const c of cells) {
          if (c?.row === null || c?.row === undefined || c?.col === null || c?.col === undefined || c.row === "" || c.col === "") throw new Error("Mỗi ô cần row và col (số, đánh từ 0).");
          const r = Number(c?.row);
          const k = Number(c?.col);
          if (!Number.isInteger(r) || !Number.isInteger(k) || r < 0 || k < 0 || r >= t.rowCount || k >= t.columnCount) throw new Error(`Ô (hàng ${c?.row}, cột ${c?.col}) nằm ngoài bảng ${t.rowCount}×${t.columnCount} (đánh số từ 0).`);
        }
        const undoable = await snapshotUndo(ctx, slide, index, `Sửa bảng ở slide ${index}`, log);
        for (const c of cells) t.getCellOrNullObject(Number(c.row), Number(c.col)).text = String(c.text ?? "");
        await ctx.sync();
        if (rows.length) {
          const first = t.rowCount;
          t.rows.add(null as any, rows.length);
          await ctx.sync();
          rows.forEach((r: any[], i: number) => r.slice(0, t.columnCount).forEach((v, k) => (t.getCellOrNullObject(first + i, k).text = String(v ?? ""))));
          await ctx.sync();
        }
        t.load("rowCount,values");
        await ctx.sync();
        return { ok: true, rows: t.rowCount, values: t.values, message: `Đã sửa ${cells.length} ô${rows.length ? `, thêm ${rows.length} hàng` : ""} trong bảng ${shp.name} ở slide ${index}${undoable ? "" : " (không tạo được bản sao để Hoàn tác)"}` };
      });

    case "ppt_set_layout":
      return PowerPoint.run(async (ctx) => {
        if (!supports("1.8")) throw new Error("PowerPoint trên máy này chưa cho add-in đổi bố cục slide (cần PowerPointApi 1.8).");
        const { slide, index } = await slideAt(ctx, a.slide);
        slide.layout.load("name");
        const masters = ctx.presentation.slideMasters;
        masters.load("items");
        await ctx.sync();
        const want = String(a.layout || "").trim().toLowerCase();
        let target: PowerPoint.SlideLayout | undefined;
        const names: string[] = [];
        for (const m of masters.items) {
          m.layouts.load("items/name");
          await ctx.sync();
          names.push(...m.layouts.items.map((l) => l.name));
          target = target || m.layouts.items.find((l) => l.name.toLowerCase() === want) || m.layouts.items.find((l) => want && l.name.toLowerCase().includes(want));
        }
        if (!target) throw new Error(`Không có bố cục "${a.layout}". Có: ${[...new Set(names)].join(", ")}`);
        const before = slide.layout.name;
        if (before.toLowerCase() === target.name.toLowerCase()) return { ok: true, changed: false, message: `Slide ${index} đã dùng bố cục "${before}".` };
        const undoable = await snapshotUndo(ctx, slide, index, `Đổi bố cục slide ${index}`, log);
        slide.applyLayout(target);
        await ctx.sync();
        slide.layout.load("name");
        slide.shapes.load("items/name,items/type");
        await ctx.sync();
        return { ok: true, message: `Slide ${index}: bố cục "${before}" → "${slide.layout.name}"${undoable ? "" : " (không tạo được bản sao để Hoàn tác)"}`, shapes: slide.shapes.items.map((s) => `${s.name} (${s.type})`) };
      });
  }
  throw new Error(`Không có công cụ ${name}`);
}
