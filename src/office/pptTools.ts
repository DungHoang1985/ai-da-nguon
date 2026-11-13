/* global PowerPoint, Office */
/** Công cụ để AI thao tác trực tiếp trên bài trình chiếu PowerPoint. Mỗi thao tác tạo/sửa đều có Hoàn tác. */
import { HostToolSet, ToolSpec, UndoEntry } from "../core/agent/types";
import { normalize2D, truncate } from "./util";

const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);

const slideProp = { type: "number", description: "Slide number (1-based)." };

export const PPT_SPECS: ToolSpec[] = [
  {
    name: "ppt_get_overview",
    description:
      "Get a presentation overview: slides (number, id, layout, shapes with name/type/text/position), selected slides and available layout names. Slide size is usually 960x540 points.",
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
    description: "Add a text box on a slide (positions in points; slide is 960x540).",
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

function supports(v: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", v);
  } catch {
    return false;
  }
}

async function slideAt(ctx: PowerPoint.RequestContext, n: number) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const i = Number(n) - 1;
  if (!Number.isInteger(i) || i < 0 || i >= slides.items.length) {
    throw new Error(`Không có slide số ${n}. Bài có ${slides.items.length} slide (đánh số từ 1).`);
  }
  return { slide: slides.items[i], index: i + 1, count: slides.items.length };
}

async function shapeOn(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide, ref: any) {
  const shapes = slide.shapes;
  shapes.load("items/id,items/name");
  await ctx.sync();
  const s = String(ref ?? "").trim();
  let shape = shapes.items.find((x) => x.name.toLowerCase() === s.toLowerCase());
  if (!shape) shape = shapes.items.find((x) => x.name.toLowerCase().includes(s.toLowerCase()) && s);
  if (!shape && /^\d+$/.test(s)) shape = shapes.items[Number(s) - 1];
  if (!shape) throw new Error(`Không thấy hình "${s}" trên slide. Có: ${shapes.items.map((x) => x.name).join(", ")}`);
  return shape;
}

function deleteShapeUndo(slideId: string, shapeId: string, label: string): UndoEntry {
  return {
    label,
    undo: () =>
      PowerPoint.run(async (c) => {
        c.presentation.slides.getItem(slideId).shapes.getItem(shapeId).delete();
        await c.sync();
      }),
  };
}

async function loadShapeTexts(ctx: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  const eligible = shapes.filter((s) => TEXT_TYPES.has(String(s.type)));
  const trs = eligible.map((s) => {
    const tr = s.textFrame.textRange;
    tr.load("text");
    return tr;
  });
  try {
    await ctx.sync();
    eligible.forEach((s, i) => (out[s.id] = trs[i].text));
  } catch {
    // Một hình không có khung chữ → đọc từng hình
    for (const s of eligible) {
      try {
        const tr = s.textFrame.textRange;
        tr.load("text");
        await ctx.sync();
        out[s.id] = tr.text;
      } catch {
        /* bỏ qua */
      }
    }
  }
  return out;
}

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
      result.push({
        slide: i + 1,
        id: s.id,
        layout: s.layout.name,
        shapes: s.shapes.items.map((sh, k) => ({
          n: k + 1,
          name: sh.name,
          type: sh.type,
          ...(texts[sh.id] !== undefined ? { text: truncate(texts[sh.id], 300) } : {}),
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
    return {
      slide_count: slides.items.length,
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

async function execute(name: string, a: any, { log }: { log: UndoEntry[] }): Promise<any> {
  switch (name) {
    case "ppt_get_overview":
      return overview();

    case "ppt_add_slide":
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
          if (title) title.textFrame.textRange.text = String(a.title);
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
          if (body) body.textFrame.textRange.text = text;
          else {
            const tb = slide.shapes.addTextBox(text, { left: 40, top: 120, width: 880, height: 380 });
            tb.textFrame.textRange.font.size = 20;
          }
        }
        await ctx.sync();
        if (supports("1.5")) {
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
              c.presentation.slides.getItem(slideId).delete();
              await c.sync();
            }),
        });
        return { ok: true, message: `${note}Đã thêm slide ${index} (bố cục "${layout.name}")`, slide: index, slide_id: slideId };
      });

    case "ppt_set_shape_text":
      return PowerPoint.run(async (ctx) => {
        const { slide } = await slideAt(ctx, a.slide);
        const shape = await shapeOn(ctx, slide, a.shape);
        const tr = shape.textFrame.textRange;
        tr.load("text");
        tr.font.load("size,bold,color");
        await ctx.sync();
        const old = { text: tr.text, size: tr.font.size, bold: tr.font.bold, color: tr.font.color };
        tr.text = String(a.text ?? "");
        if (a.font_size) tr.font.size = Number(a.font_size);
        if (a.bold !== undefined) tr.font.bold = !!a.bold;
        if (a.font_color) tr.font.color = a.font_color;
        await ctx.sync();
        const slideId = slide.id;
        const shapeId = shape.id;
        log.push({
          label: `Sửa chữ ${shape.name}`,
          undo: () =>
            PowerPoint.run(async (c) => {
              const t = c.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange;
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
        log.push(deleteShapeUndo(slide.id, shp.id, `Ô chữ ${shp.name}`));
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
        log.push(deleteShapeUndo(slide.id, shp.id, `Hình ${shp.name}`));
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
        log.push(deleteShapeUndo(slide.id, shp.id, `Bảng ${shp.name}`));
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

export const pptTools: HostToolSet = {
  host: "PowerPoint",
  specs: PPT_SPECS,
  execute,
  overview,
  systemPrompt: (ov) =>
    [
      "Bạn là trợ lý AI được tích hợp trong Microsoft PowerPoint, có thể thao tác trực tiếp trên bài trình chiếu đang mở bằng các công cụ (ppt_*).",
      "Nguyên tắc:",
      "- Khi người dùng muốn tạo slide, sửa chữ, thêm hình/bảng… hãy GỌI CÔNG CỤ để làm thật, không chỉ đưa dàn ý.",
      "- Slide đánh số từ 1; kích thước thường là 960×540 điểm. Dùng bố cục có sẵn của tệp (danh sách layouts) để giữ đúng giao diện.",
      "- Nội dung slide ngắn gọn: tiêu đề ≤ 10 từ, mỗi slide 3–6 ý, mỗi ý một dòng.",
      "- Đặt hình/ô chữ không chồng lên nhau, cách mép ≥ 40 điểm.",
      "- Xong việc: trả lời ngắn gọn bằng tiếng Việt: đã tạo/sửa slide nào.",
      "",
      "Trạng thái hiện tại của bài trình chiếu (JSON):",
      JSON.stringify(ov).slice(0, 6000),
    ].join("\n"),
};
