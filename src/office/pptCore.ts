/* global PowerPoint, Office */
/** Hàm dùng chung cho các công cụ PowerPoint. */
import { UndoEntry } from "../core/agent/types";

export const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);

export function supports(v: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", v);
  } catch {
    return false;
  }
}

export async function slideAt(ctx: PowerPoint.RequestContext, n: number) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const i = Number(n) - 1;
  if (!Number.isInteger(i) || i < 0 || i >= slides.items.length) {
    throw new Error(`Không có slide số ${n}. Bài có ${slides.items.length} slide (đánh số từ 1).`);
  }
  return { slide: slides.items[i], index: i + 1, count: slides.items.length };
}

export async function shapeOn(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide, ref: any) {
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

/** Khi một slide được khôi phục bằng bản sao, id slide đổi: ghi lại id cũ → id mới để các bước Hoàn tác trước đó vẫn tìm thấy. */
const slideRemap = new Map<string, string>();
export function remapSlide(oldId: string, newId: string) {
  slideRemap.set(oldId, newId);
}
export function currentSlideId(id: string): string {
  let cur = id;
  for (let i = 0; i < 20 && slideRemap.has(cur); i++) cur = slideRemap.get(cur)!;
  return cur;
}

/** Tìm lại hình khi Hoàn tác: theo id, nếu slide đã được khôi phục bằng bản sao (id hình có thể đổi) thì theo tên. */
export async function findShape(c: PowerPoint.RequestContext, slideId: string, shapeId: string, name?: string): Promise<PowerPoint.Shape> {
  const shapes = c.presentation.slides.getItem(currentSlideId(slideId)).shapes;
  shapes.load("items/id,items/name");
  await c.sync();
  const s = shapes.items.find((x) => x.id === shapeId) || (name ? shapes.items.find((x) => x.name === name) : undefined);
  if (!s) throw new Error(`Không còn hình ${name || shapeId} trên slide`);
  return s;
}

export function deleteShapeUndo(slideId: string, shapeId: string, label: string, name?: string): UndoEntry {
  return {
    label,
    undo: () =>
      PowerPoint.run(async (c) => {
        (await findShape(c, slideId, shapeId, name)).delete();
        await c.sync();
      }),
  };
}

export async function loadShapeTexts(ctx: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]): Promise<Record<string, string>> {
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

