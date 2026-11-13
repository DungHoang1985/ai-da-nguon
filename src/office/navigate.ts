/* global Excel, Word, PowerPoint */
/** Bấm vào địa chỉ ô / đoạn / slide trong câu trả lời → chuyển tới đúng chỗ trong tài liệu. */
import { RefTarget } from "../core/refs";
import { getRange } from "./excelCore";

export async function navigateTo(host: string, ref: RefTarget): Promise<void> {
  if (host === "Excel" && ref.kind === "cell") {
    await Excel.run(async (ctx) => {
      const { ws, range } = getRange(ctx, ref.address);
      ws.activate();
      range.select();
      await ctx.sync();
    });
    return;
  }
  if (host === "Word" && ref.kind === "para") {
    await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const p = ps.items[ref.index];
      if (!p) throw new Error(`Không có đoạn số ${ref.index}.`);
      p.select();
      await ctx.sync();
    });
    return;
  }
  if (host === "PowerPoint" && ref.kind === "slide") {
    await PowerPoint.run(async (ctx) => {
      const s = ctx.presentation.slides;
      s.load("items/id");
      await ctx.sync();
      const sl = s.items[ref.index - 1];
      if (!sl) throw new Error(`Không có slide số ${ref.index}.`);
      ctx.presentation.setSelectedSlides([sl.id]);
      await ctx.sync();
    });
  }
}
