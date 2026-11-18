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
  // Bài có nhiều hình trùng tên (đã gặp khi Chạy thật: hai ô giữ chỗ cùng tên "Text 2" làm AI sửa nhầm ô tiêu đề): bắt dùng số thứ tự
  const same = shapes.items.map((x, k) => ({ x, k })).filter((y) => y.x.name.toLowerCase() === s.toLowerCase());
  if (same.length > 1) {
    throw new Error(`Slide có ${same.length} hình cùng tên "${s}" (hình số ${same.map((y) => y.k + 1).join(", ")}). Hãy chỉ hình bằng SỐ THỨ TỰ của nó (trường n trong tổng quan), không dùng tên.`);
  }
  let shape = same[0]?.x;
  // Số thứ tự trước, rồi mới tới khớp một phần của tên ("1" không được trúng hình tên "Rectangle 12")
  if (!shape && /^\d+$/.test(s)) shape = shapes.items[Number(s) - 1];
  if (!shape) shape = shapes.items.find((x) => x.name.toLowerCase().includes(s.toLowerCase()) && s);
  if (!shape) throw new Error(`Không thấy hình "${s}" trên slide. Có: ${shapes.items.map((x) => x.name).join(", ")}`);
  return shape;
}

const sameFont = (a: any, b: any) => String(a ?? "").trim().toLowerCase() === String(b ?? "").trim().toLowerCase();

/** Trả phông `want` cho những dòng chữ của một hình đang mang phông khác. Trả số dòng đã sửa.
 *  Vì sao cần (đo trong PowerPoint 18/11/2026, `kiem-thu/cdp-bai/tham-do-7-phong-tieu-de.js`): ghi chữ có ký tự “ư” vào một ô giữ chỗ
 *  bằng Office.js làm PowerPoint đổi CẢ dòng sang Times New Roman (chữ có “ổ”, “ậ”, “ý” thì không); đặt lại tên phông sau khi ghi thì hết.
 *  Chỉ đụng dòng bị đổi, để các dòng khác vẫn theo phông chủ đề. */
export async function restoreFont(ctx: PowerPoint.RequestContext, shape: PowerPoint.Shape, want: string | null | undefined): Promise<number> {
  if (!want || !String(want).trim()) return 0;
  const tr = shape.textFrame.textRange;
  tr.load("text");
  await ctx.sync();
  const text = String(tr.text || "");
  if (!text.trim()) return 0;
  const parts: PowerPoint.TextRange[] = [];
  let at = 0;
  for (const line of text.split(/\r\n|\r|\n|\v/)) {
    if (line.length) {
      const r = tr.getSubstring(at, line.length);
      r.font.load("name");
      parts.push(r);
    }
    at += line.length + 1;
  }
  await ctx.sync();
  let fixed = 0;
  for (const r of parts) {
    if (!sameFont(r.font.name, want)) {
      r.font.name = String(want);
      fixed++;
    }
  }
  if (fixed) await ctx.sync();
  return fixed;
}

/** Ghi chữ vào một khung (ô tiêu đề, ô nội dung…) mà giữ nguyên phông khung đang có trước khi ghi (thường là phông chủ đề). */
export async function setTextKeepFont(ctx: PowerPoint.RequestContext, shape: PowerPoint.Shape, text: string): Promise<void> {
  const tr = shape.textFrame.textRange;
  tr.font.load("name");
  let want: string | null = null;
  try {
    await ctx.sync();
    want = tr.font.name || null;
  } catch {
    want = null;
  }
  shape.textFrame.textRange.text = text;
  await ctx.sync();
  try {
    await restoreFont(ctx, shape, want);
  } catch {
    /* không đọc được phông từng dòng: chữ đã ghi, giữ như PowerPoint đặt */
  }
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

