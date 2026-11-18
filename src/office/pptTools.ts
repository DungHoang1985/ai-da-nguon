/* global PowerPoint, Office */
/** Công cụ để AI thao tác trực tiếp trên bài trình chiếu PowerPoint. Mỗi thao tác tạo/sửa đều có Hoàn tác. */
import { HostToolSet, ToolSpec, UndoEntry } from "../core/agent/types";
import { normalize2D, truncate } from "./util";
import { currentSlideId, deleteShapeUndo, findShape, loadShapeTexts, restoreFont, setTextKeepFont, shapeOn, slideAt, supports } from "./pptCore";
import { PPT_ADV_SPECS, executePptAdvanced, pptRisk, slideSizePt } from "./pptAdvanced";
import { findAttachedImage, listAttachedImages } from "../core/attachedFiles";
import { imageBox } from "../core/imagefit";
import { PPT_DESIGN_SPECS, executePptDesign } from "./pptDesign";

const slideProp = { type: "number", description: "Slide number (1-based)." };

export const PPT_SPECS: ToolSpec[] = [
  {
    name: "ppt_get_overview",
    description:
      "Get a presentation overview: slide size (points), slides (number, id, layout, shapes with name/type/text/position), selected slides and available layout names.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "ppt_add_slide",
    description:
      "Add a slide using one of the presentation's layouts and fill its title and body placeholders. Body: one bullet per line, short.",
    parameters: {
      type: "object",
      properties: {
        layout: { type: "string", description: "Layout name, e.g. 'Title and Content', 'Title Slide', 'Two Content'." },
        title: { type: "string" },
        body: { type: "string", description: "Bullet lines separated by \\n." },
        after_slide: { type: "number", description: "Insert after this slide number (default: at the end)." },
      },
    },
    write: true,
  },
  {
    name: "ppt_add_slides",
    description:
      "Add SEVERAL slides in one call (use this instead of ppt_add_slide whenever you add 2 or more slides). Slides are appended at the end in the given order; each uses one of the presentation's layouts and fills its title and body placeholders. Max 10 slides per call. Undo removes the whole batch.",
    parameters: {
      type: "object",
      properties: {
        slides: {
          type: "array",
          description: "Slides in order.",
          items: {
            type: "object",
            properties: {
              layout: { type: "string", description: "Layout name, e.g. 'Title and Content', 'Title Slide', 'Title Only'." },
              title: { type: "string" },
              body: { type: "string", description: "Bullet lines separated by \\n. Omit for slides that will get a chart or table." },
            },
          },
        },
      },
      required: ["slides"],
    },
    write: true,
  },
  {
    name: "ppt_insert_image",
    description:
      "Place an image the user ATTACHED in this chat onto a slide, keeping its aspect ratio. Only images attached in this session can be used (not web URLs). By default the image is fitted into the content area below the title; use place or left/top/width/height (points) to position it. Undo removes it.",
    parameters: {
      type: "object",
      properties: {
        slide: { type: "number", description: "Slide number (from 1)." },
        image: { type: "string", description: "Attached image file name (or its position '1', '2'…). Omit to use the most recently attached image." },
        place: { type: "string", enum: ["content", "left", "right", "center", "full"], description: "Where to fit the image: content area (default), left or right half, centered, or the whole slide." },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number", description: "Give width OR height to set the size; the other side follows the aspect ratio." },
        height: { type: "number" },
      },
      required: ["slide"],
    },
    write: true,
  },
  {
    name: "ppt_set_shape_text",
    description: "Replace the text of a shape (title, placeholder, text box) on a slide; optionally change its font.",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        shape: { type: "string", description: "Shape name from the overview, or its 1-based number on the slide." },
        text: { type: "string" },
        font_size: { type: "number" },
        bold: { type: "boolean" },
        font_color: { type: "string", description: "Hex colour like #1F4E79." },
      },
      required: ["slide", "shape", "text"],
    },
    write: true,
  },
  {
    name: "ppt_add_textbox",
    description: "Add a text box on a slide (positions in points; see slide_size in the overview).",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        text: { type: "string" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        font_size: { type: "number" },
        bold: { type: "boolean" },
        font_color: { type: "string" },
        align: { type: "string", enum: ["Left", "Center", "Right"] },
      },
      required: ["slide", "text"],
    },
    write: true,
  },
  {
    name: "ppt_add_shape",
    description: "Add a shape (rectangle, rounded rectangle, ellipse, arrow…) with optional text and fill colour.",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        shape_type: {
          type: "string",
          enum: ["rectangle", "rounded_rectangle", "ellipse", "triangle", "right_arrow", "chevron"],
        },
        text: { type: "string" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        fill_color: { type: "string" },
        font_color: { type: "string" },
        font_size: { type: "number" },
      },
      required: ["slide", "shape_type"],
    },
    write: true,
  },
  {
    name: "ppt_add_table",
    description: "Add a table on a slide. First row is the header.",
    parameters: {
      type: "object",
      properties: {
        slide: slideProp,
        values: { type: "array", items: { type: "array", items: { type: "string" } } },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
      },
      required: ["slide", "values"],
    },
    write: true,
  },
  {
    name: "ppt_go_to_slide",
    description: "Select (show) a slide for the user.",
    parameters: { type: "object", properties: { slide: slideProp }, required: ["slide"] },
  },
];

async function overview() {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const list = slides.items.slice(0, 40);
    list.forEach((s) => {
      s.layout.load("name");
      s.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    });
    await ctx.sync();
    const result: any[] = [];
    for (let i = 0; i < list.length; i++) {
      const s = list[i];
      const texts = await loadShapeTexts(ctx, s.shapes.items);
      // Bảng: kèm nội dung ô để AI đọc và sửa được bảng có sẵn (ppt_edit_table)
      const tabs: Record<string, string[][]> = {};
      if (supports("1.8")) {
        const ts = s.shapes.items.filter((x) => String(x.type) === "Table").map((x) => ({ id: x.id, t: x.getTable() }));
        if (ts.length) {
          ts.forEach((x) => x.t.load("values"));
          try {
            await ctx.sync();
            ts.forEach((x) => (tabs[x.id] = x.t.values.slice(0, 15).map((r) => r.slice(0, 8).map((c) => truncate(String(c ?? ""), 40)))));
          } catch {
            /* không đọc được bảng: bỏ qua */
          }
        }
      }
      result.push({
        slide: i + 1,
        id: s.id,
        layout: s.layout.name,
        shapes: s.shapes.items.map((sh, k) => ({
          n: k + 1,
          name: sh.name,
          type: sh.type,
          ...(texts[sh.id] !== undefined ? { text: truncate(texts[sh.id], 300) } : {}),
          ...(tabs[sh.id] ? { table: tabs[sh.id] } : {}),
          box: [Math.round(sh.left), Math.round(sh.top), Math.round(sh.width), Math.round(sh.height)],
        })),
      });
    }
    let selected: number[] = [];
    if (supports("1.5")) {
      try {
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        selected = sel.items.map((x) => slides.items.findIndex((y) => y.id === x.id) + 1).filter((n) => n > 0);
      } catch {
        selected = [];
      }
    }
    let layouts: string[] = [];
    try {
      const m = ctx.presentation.slideMasters.getItemAt(0);
      m.layouts.load("items/name");
      await ctx.sync();
      layouts = m.layouts.items.map((l) => l.name);
    } catch {
      layouts = [];
    }
    const size = await slideSizePt(ctx);
    return {
      slide_count: slides.items.length,
      slide_size: { width: size.w, height: size.h },
      selected_slides: selected,
      layouts,
      slides: result,
      ...(slides.items.length > list.length ? { note: `Chỉ liệt kê ${list.length} slide đầu.` } : {}),
    };
  });
}

/** Tìm ô tiêu đề và ô nội dung của slide mới */
async function findPlaceholders(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide) {
  const shapes = slide.shapes;
  shapes.load("items/id,items/name,items/type");
  await ctx.sync();
  let title: PowerPoint.Shape | undefined;
  let body: PowerPoint.Shape | undefined;
  const ph = shapes.items.filter((s) => s.type === "Placeholder");
  if (supports("1.8") && ph.length) {
    try {
      ph.forEach((s) => s.placeholderFormat.load("type"));
      await ctx.sync();
      title = ph.find((s) => ["Title", "CenterTitle", "VerticalTitle"].includes(s.placeholderFormat.type));
      body = ph.find((s) => ["Body", "Content", "Subtitle", "VerticalBody", "VerticalContent"].includes(s.placeholderFormat.type));
    } catch {
      /* dùng cách đoán theo tên */
    }
  }
  if (!title) title = shapes.items.find((s) => /title|tiêu đề/i.test(s.name) && !/sub/i.test(s.name));
  if (!body) body = shapes.items.find((s) => s !== title && /content|text|body|subtitle|nội dung|văn bản|phụ đề/i.test(s.name));
  if (!title && ph[0]) title = ph[0];
  if (!body) body = ph.find((s) => s !== title);
  return { title, body };
}

/** Đọc kích thước ảnh đính kèm; ảnh không phải PNG/JPEG (GIF, WebP) hoặc quá lớn thì vẽ lại thành PNG để PowerPoint nhận được.
 *  Trả base64 (không kèm tiền tố data:). */
async function loadImageForSlide(name: string, data: string): Promise<{ w: number; h: number; data: string }> {
  const lower = name.toLowerCase();
  const mime = lower.endsWith(".png") ? "image/png" : /\.jpe?g$/.test(lower) ? "image/jpeg" : lower.endsWith(".gif") ? "image/gif" : "image/webp";
  const el = await new Promise<HTMLImageElement>((resolve, reject) => {
    const im = new Image();
    im.onload = () => resolve(im);
    im.onerror = () => reject(new Error(`Không đọc được ảnh "${name}" (tệp hỏng hoặc không phải ảnh).`));
    im.src = `data:${mime};base64,${data}`;
  });
  const w = el.naturalWidth;
  const h = el.naturalHeight;
  if (!w || !h) throw new Error(`Không đọc được kích thước của ảnh "${name}".`);
  const MAX_SIDE = 2400;
  if ((mime === "image/png" || mime === "image/jpeg") && Math.max(w, h) <= MAX_SIDE) return { w, h, data };
  const k = Math.min(1, MAX_SIDE / Math.max(w, h));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(w * k));
  canvas.height = Math.max(1, Math.round(h * k));
  const g = canvas.getContext("2d");
  if (!g) throw new Error("Không chuyển được ảnh sang PNG trong môi trường này.");
  g.drawImage(el, 0, 0, canvas.width, canvas.height);
  return { w, h, data: canvas.toDataURL("image/png").split(",")[1] };
}

export const MAX_SLIDES_PER_CALL = 10;

/** Chuẩn hóa danh sách slide của ppt_add_slides: nhận mảng hoặc chuỗi JSON, bỏ phần tử rỗng, tối đa MAX_SLIDES_PER_CALL */
export function normalizeSlideList(v: any): { layout?: string; title?: string; body?: string }[] {
  let list = v;
  if (typeof list === "string") {
    try {
      list = JSON.parse(list);
    } catch {
      list = [];
    }
  }
  if (!Array.isArray(list)) return [];
  return list
    .filter((x) => x && typeof x === "object" && (String(x.title ?? "").trim() || String(x.body ?? "").trim() || String(x.layout ?? "").trim()))
    .slice(0, MAX_SLIDES_PER_CALL)
    .map((x) => ({ layout: x.layout === undefined ? undefined : String(x.layout), title: x.title === undefined ? undefined : String(x.title), body: x.body === undefined ? undefined : String(x.body) }));
}

/** Thêm một slide theo bố cục của mẫu và điền tiêu đề, nội dung. select = chuyển tới slide vừa thêm. */
async function addOneSlide(a: any, log: UndoEntry[], select: boolean): Promise<{ slide: number; slide_id: string; layout: string; note: string }> {
      return PowerPoint.run(async (ctx) => {
        const master = ctx.presentation.slideMasters.getItemAt(0);
        master.load("id");
        master.layouts.load("items/id,items/name");
        await ctx.sync();
        const layouts = master.layouts.items;
        const want = String(a.layout || "").toLowerCase().trim();
        let layout =
          (want && (layouts.find((l) => l.name.toLowerCase() === want) || layouts.find((l) => l.name.toLowerCase().includes(want)))) ||
          undefined;
        let note = "";
        if (!layout) {
          if (want) note = `Không có bố cục "${a.layout}", đã dùng bố cục khác. `;
          layout =
            layouts.find((l) => /^title and content$/i.test(l.name)) ||
            layouts.find((l) => /content|nội dung/i.test(l.name)) ||
            layouts[1] ||
            layouts[0];
        }
        ctx.presentation.slides.add({ slideMasterId: master.id, layoutId: layout.id });
        await ctx.sync();
        const slides = ctx.presentation.slides;
        slides.load("items/id");
        await ctx.sync();
        const slide = slides.items[slides.items.length - 1];
        const slideId = slide.id;
        if (a.after_slide && supports("1.8")) {
          const target = Math.min(Math.max(Number(a.after_slide), 0), slides.items.length - 1);
          (slide as any).moveTo(target);
          await ctx.sync();
        }
        const { title, body } = await findPlaceholders(ctx, slide);
        if (a.title) {
          if (title) await setTextKeepFont(ctx, title, String(a.title));
          else {
            const tb = slide.shapes.addTextBox(String(a.title), { left: 40, top: 30, width: 880, height: 70 });
            tb.textFrame.textRange.font.size = 32;
            tb.textFrame.textRange.font.bold = true;
          }
        }
        if (a.body) {
          const text = String(a.body)
            .split(/\r?\n/)
            .map((l) => l.replace(/^\s*[-*•]\s+/, ""))
            .filter((l) => l.trim())
            .join("\n");
          if (body) await setTextKeepFont(ctx, body, text);
          else {
            const tb = slide.shapes.addTextBox(text, { left: 40, top: 120, width: 880, height: 380 });
            tb.textFrame.textRange.font.size = 20;
          }
        }
        await ctx.sync();
        if (select && supports("1.5")) {
          try {
            ctx.presentation.setSelectedSlides([slideId]);
            await ctx.sync();
          } catch {
            /* bỏ qua */
          }
        }
        slides.load("items/id");
        await ctx.sync();
        const index = slides.items.findIndex((s) => s.id === slideId) + 1;
        log.push({
          label: `Slide ${index}`,
          undo: () =>
            PowerPoint.run(async (c) => {
              c.presentation.slides.getItem(currentSlideId(slideId)).delete();
              await c.sync();
            }),
        });
        return { slide: index, slide_id: slideId, layout: layout.name, note };
      });
}

async function execute(name: string, a: any, { log }: { log: UndoEntry[] }): Promise<any> {
  switch (name) {
    case "ppt_get_overview":
      return overview();

    case "ppt_add_slides": {
      const list = normalizeSlideList(a.slides);
      if (!list.length) throw new Error("Thiếu danh sách slides (mỗi phần tử gồm layout, title, body).");
      // Mỗi slide ghi mục Hoàn tác riêng vào sổ tạm; cả lô gộp thành MỘT mục Hoàn tác (gỡ từ slide cuối lên)
      const batch: UndoEntry[] = [];
      const done: any[] = [];
      let failed = "";
      for (const one of list) {
        try {
          done.push(await addOneSlide(one, batch, false));
        } catch (e: any) {
          failed = e?.message || String(e);
          break;
        }
      }
      if (batch.length) {
        log.push({
          label: `Thêm ${batch.length} slide`,
          undo: async () => {
            for (const u of batch.slice().reverse()) await u.undo();
          },
        });
      }
      if (!done.length) throw new Error(failed || "Không thêm được slide nào.");
      const nums = done.map((d) => d.slide);
      return {
        ok: !failed,
        message: `Đã thêm ${done.length}/${list.length} slide: ${nums.join(", ")}${failed ? `. Dừng ở slide thứ ${done.length + 1} vì lỗi: ${failed}` : ""}`,
        slides: done.map((d) => ({ slide: d.slide, slide_id: d.slide_id, layout: d.layout, ...(d.note ? { note: d.note } : {}) })),
        ...(Array.isArray(a.slides) && a.slides.length > list.length ? { note: `Chỉ nhận ${MAX_SLIDES_PER_CALL} slide mỗi lần gọi; ${a.slides.length - list.length} slide còn lại chưa thêm.` } : {}),
      };
    }

    case "ppt_add_slide": {
      const r = await addOneSlide(a, log, true);
      return { ok: true, message: `${r.note}Đã thêm slide ${r.slide} (bố cục "${r.layout}")`, slide: r.slide, slide_id: r.slide_id };
    }


    case "ppt_insert_image": {
      const have = listAttachedImages();
      const file = findAttachedImage(a.image);
      if (!file) {
        throw new Error(
          have.length
            ? `Không thấy ảnh đính kèm "${a.image}". Ảnh đang có: ${have.join(", ")}.`
            : "Chưa có ảnh đính kèm nào trong phiên này (ảnh đính kèm mất khi tải lại khung hoặc mở hội thoại mới). Hãy nói người dùng đính kèm ảnh rồi gửi lại yêu cầu; không dùng được ảnh từ địa chỉ web."
        );
      }
      const img = await loadImageForSlide(file.name, file.data);
      return PowerPoint.run(async (ctx) => {
        const { slide } = await slideAt(ctx, a.slide);
        const size = await slideSizePt(ctx);
        // Mép dưới của ô tiêu đề (nếu có) để ảnh không đè lên tiêu đề
        let titleBottom: number | undefined;
        try {
          const { title } = await findPlaceholders(ctx, slide);
          if (title) {
            title.load("top,height");
            await ctx.sync();
            titleBottom = title.top + title.height;
          }
        } catch {
          titleBottom = undefined;
        }
        const box = imageBox({ w: img.w, h: img.h }, size, { place: a.place, left: a.left, top: a.top, width: a.width, height: a.height, titleBottom });
        const shp = slide.shapes.addGeometricShape("Rectangle" as any, box);
        shp.fill.setImage(img.data);
        shp.lineFormat.visible = false;
        shp.name = `Ảnh ${file.name}`.slice(0, 60);
        shp.load("id,name");
        await ctx.sync();
        log.push(deleteShapeUndo(slide.id, shp.id, `Ảnh ${file.name} ở slide ${a.slide}`, shp.name));
        return {
          ok: true,
          message: `Đã đặt ảnh "${file.name}" vào slide ${a.slide}`,
          shape: shp.name,
          box: [Math.round(box.left), Math.round(box.top), Math.round(box.width), Math.round(box.height)],
          image_px: [img.w, img.h],
          note: "Ảnh là phần tô của một hình chữ nhật: đổi cỡ thì giữ đúng tỷ lệ rộng/cao này; không cắt xén được như ảnh thường.",
        };
      });
    }

    case "ppt_set_shape_text":
      return PowerPoint.run(async (ctx) => {
        const { slide } = await slideAt(ctx, a.slide);
        const shape = await shapeOn(ctx, slide, a.shape);
        const tr = shape.textFrame.textRange;
        tr.load("text");
        tr.font.load("size,bold,color");
        await ctx.sync();
        const old = { text: tr.text, size: tr.font.size, bold: tr.font.bold, color: tr.font.color };
        // Phông của khung trước khi ghi: ghi chữ có “ư” làm PowerPoint đổi cả dòng sang Times New Roman (đo 18/11/2026) → trả lại sau khi ghi
        const fontBefore = shape.textFrame.textRange.font;
        fontBefore.load("name");
        await ctx.sync();
        const wantFont = fontBefore.name;
        tr.text = String(a.text ?? "");
        if (a.font_size) tr.font.size = Number(a.font_size);
        if (a.bold !== undefined) tr.font.bold = !!a.bold;
        if (a.font_color) tr.font.color = a.font_color;
        await ctx.sync();
        await restoreFont(ctx, shape, wantFont).catch(() => 0);
        const slideId = slide.id;
        const shapeId = shape.id;
        const shapeName = shape.name;
        log.push({
          label: `Sửa chữ ${shape.name}`,
          undo: () =>
            PowerPoint.run(async (c) => {
              const t = (await findShape(c, slideId, shapeId, shapeName)).textFrame.textRange;
              t.text = old.text;
              if (old.size) t.font.size = old.size;
              if (old.bold !== null && old.bold !== undefined) t.font.bold = old.bold;
              if (old.color) t.font.color = old.color;
              await c.sync();
            }),
        });
        return { ok: true, message: `Đã sửa chữ "${shape.name}" ở slide ${a.slide}` };
      });

    case "ppt_add_textbox":
      return PowerPoint.run(async (ctx) => {
        const { slide } = await slideAt(ctx, a.slide);
        const shp = slide.shapes.addTextBox(String(a.text ?? ""), {
          left: a.left ?? 60,
          top: a.top ?? 60,
          width: a.width ?? 400,
          height: a.height ?? 60,
        });
        const tr = shp.textFrame.textRange;
        if (a.font_size) tr.font.size = Number(a.font_size);
        if (a.bold !== undefined) tr.font.bold = !!a.bold;
        if (a.font_color) tr.font.color = a.font_color;
        if (a.align) tr.paragraphFormat.horizontalAlignment = a.align;
        shp.load("id,name");
        await ctx.sync();
        log.push(deleteShapeUndo(slide.id, shp.id, `Ô chữ ${shp.name}`, shp.name));
        return { ok: true, message: `Đã thêm ô chữ ở slide ${a.slide}`, shape: shp.name };
      });

    case "ppt_add_shape":
      return PowerPoint.run(async (ctx) => {
        const types: Record<string, string> = {
          rectangle: "Rectangle",
          rounded_rectangle: "RoundRectangle",
          ellipse: "Ellipse",
          triangle: "Triangle",
          right_arrow: "RightArrow",
          chevron: "Chevron",
        };
        const { slide } = await slideAt(ctx, a.slide);
        const shp = slide.shapes.addGeometricShape((types[a.shape_type] || "Rectangle") as any, {
          left: a.left ?? 100,
          top: a.top ?? 100,
          width: a.width ?? 200,
          height: a.height ?? 100,
        });
        if (a.fill_color) shp.fill.setSolidColor(a.fill_color);
        if (a.text) {
          const tr = shp.textFrame.textRange;
          tr.text = String(a.text);
          if (a.font_color) tr.font.color = a.font_color;
          if (a.font_size) tr.font.size = Number(a.font_size);
        }
        shp.load("id,name");
        await ctx.sync();
        log.push(deleteShapeUndo(slide.id, shp.id, `Hình ${shp.name}`, shp.name));
        return { ok: true, message: `Đã thêm hình ${a.shape_type} ở slide ${a.slide}`, shape: shp.name };
      });

    case "ppt_add_table":
      return PowerPoint.run(async (ctx) => {
        if (!supports("1.8")) throw new Error("PowerPoint trên máy này chưa hỗ trợ chèn bảng qua add-in.");
        const vals = normalize2D(a.values).map((r) => r.map((v) => String(v)));
        const { slide } = await slideAt(ctx, a.slide);
        const opts: any = { values: vals, left: a.left ?? 40, top: a.top ?? 120 };
        if (a.width) opts.width = a.width;
        if (a.height) opts.height = a.height;
        const shp = slide.shapes.addTable(vals.length, vals[0].length, opts);
        shp.load("id,name");
        await ctx.sync();
        log.push(deleteShapeUndo(slide.id, shp.id, `Bảng ${shp.name}`, shp.name));
        return { ok: true, message: `Đã thêm bảng ${vals.length}×${vals[0].length} ở slide ${a.slide}`, shape: shp.name };
      });

    case "ppt_go_to_slide":
      return PowerPoint.run(async (ctx) => {
        const { slide } = await slideAt(ctx, a.slide);
        if (supports("1.5")) {
          ctx.presentation.setSelectedSlides([slide.id]);
          await ctx.sync();
        }
        return { ok: true, message: `Đã chuyển tới slide ${a.slide}` };
      });
  }
  throw new Error(`Không có công cụ ${name}`);
}

const PPT_RISKY = new Set(["ppt_delete_slide", "ppt_delete_shape"]);
const ALL_PPT_SPECS = [...PPT_SPECS, ...PPT_ADV_SPECS, ...PPT_DESIGN_SPECS];

export const pptTools: HostToolSet = {
  host: "PowerPoint",
  specs: ALL_PPT_SPECS.map((sp) => (PPT_RISKY.has(sp.name) ? { ...sp, risky: (args: any) => pptRisk(sp.name, args) } : sp)),
  execute: async (name, a, ctx) => {
    if (PPT_DESIGN_SPECS.some((sp) => sp.name === name)) return executePptDesign(name, a || {}, ctx.log);
    const adv = PPT_ADV_SPECS.some((sp) => sp.name === name);
    return adv ? executePptAdvanced(name, a || {}, ctx.log) : execute(name, a || {}, ctx);
  },
  snapshot: async () => ({ title: "Bài trình chiếu PowerPoint đang mở", data: await overview() }),
  overview,
  systemPrompt: () =>
    [
      "Bạn là trợ lý AI được tích hợp trong Microsoft PowerPoint, có thể thao tác trực tiếp trên bài trình chiếu đang mở bằng các công cụ (ppt_*).",
      "Nguyên tắc:",
      "- Khi người dùng muốn tạo slide, sửa chữ, thêm hình/bảng… hãy GỌI CÔNG CỤ để làm thật, không chỉ đưa dàn ý.",
      "- Slide đánh số từ 1; kích thước slide (điểm) ở slide_size. Dùng bố cục có sẵn của tệp (danh sách layouts) để giữ đúng giao diện.",
      "- Số liệu cần so sánh → ppt_add_chart (biểu đồ gốc, sửa được dữ liệu, theo màu chủ đề). Quy trình/các bước/mốc thời gian → ppt_add_process_flow. Sắp xếp slide → ppt_move_slide / ppt_duplicate_slide / ppt_delete_slide. Chỉnh vị trí, màu, chữ của hình → ppt_update_shape.",
      "- THEO MẪU: trước khi thêm slide hoặc hình vào một bài có sẵn, xem mẫu bằng ppt_get_theme (màu chủ đề, phông, bố cục và ô giữ chỗ); dùng bố cục và màu của mẫu, không tự đặt phông hay màu lạ.",
      "- ÍT LƯỢT GỌI: từ 2 slide trở lên thì dùng ppt_add_slides (một lệnh thêm cả lô); gọi các lệnh thêm slide, bảng, biểu đồ trong CÙNG một lượt (slide mới nối vào cuối bài theo thứ tự lệnh, nên biết trước số slide); làm xong HẾT rồi mới kiểm, và kiểm một lần cho mọi slide vừa tạo bằng ppt_get_slide_image với slides: [..]. Không gọi ppt_get_overview sau mỗi bước.",
      "- TỰ KIỂM BẰNG MẮT: sau khi tạo hoặc sửa slide có nhiều chữ, bảng, hình, gọi ppt_get_slide_image để NHÌN slide; thấy chữ tràn khung, hình chồng nhau thì sửa ĐÚNG hình có vấn đề (giảm cỡ chữ bằng ppt_format_text, nới hoặc dời khung bằng ppt_update_shape, tách slide) rồi chụp lại MỘT lần để xác nhận. Không dời, không đổi cỡ ô tiêu đề và ô nội dung của bố cục khi chúng không có vấn đề. Slide có hình trùng tên thì chỉ hình bằng số thứ tự (n). Không chụp lại quá hai lần cho một slide; ảnh cho thấy ổn thì dừng, không sửa tiếp theo số đo.",
      "- KIỂM BÀI: khi người dùng nhờ kiểm tra, soát, hoàn thiện bài, gọi ppt_check_deck trước (phông lệch mẫu, chữ quá nhỏ, slide thiếu tiêu đề, chữ tràn, hình ra ngoài slide), báo đủ từng điểm kèm số slide; chỉ sửa khi được nhờ.",
      "- Định dạng một đoạn hay một cụm chữ trong hình: ppt_format_text. Sửa bảng có sẵn: ppt_edit_table. Người dùng nói 'slide này', 'hình này', 'chữ đang chọn': ppt_get_selection. Đổi bố cục slide: ppt_set_layout.",
      "- PowerPoint không cho add-in ghi chú diễn giả, sửa số liệu của biểu đồ đã có, tạo SmartArt, chuyển cảnh: được nhờ thì nói rõ là không làm được và gợi ý cách thay (biểu đồ: tạo biểu đồ mới bằng ppt_add_chart rồi xóa biểu đồ cũ sau khi người dùng đồng ý).",
      "- Khi nhắc tới slide, ghi \"slide N\" để người dùng bấm vào xem.",
      "- Nội dung slide ngắn gọn: tiêu đề ≤ 10 từ, mỗi slide 3–6 ý, mỗi ý một dòng.",
      "- Đặt hình/ô chữ không chồng lên nhau, cách mép ≥ 40 điểm.",
      "- Xong việc: trả lời ngắn gọn bằng tiếng Việt: đã tạo/sửa slide nào.",
    ].join("\n"),
  stateText: (ov) => `Trạng thái hiện tại của bài trình chiếu (JSON):\n${JSON.stringify(ov).slice(0, 6000)}`,
};
