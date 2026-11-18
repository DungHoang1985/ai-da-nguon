// Thăm dò sau 4.7.1, bước C: slide thêm bằng Office.js theo bố cục (cách ppt_add_slide làm) — ô giữ chỗ thừa hưởng vị trí từ bố cục
// có bị báo 0, 0, 0, 0 như slide biểu đồ không? Đọc vị trí các ô giữ chỗ của mọi bố cục, rồi xóa slide thử.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-toa-do-c.js
const box = (s) => [s.left, s.top, s.width, s.height].map((v) => Math.round(v * 10) / 10);
const out = [];
await PowerPoint.run(async (ctx) => {
  const layouts = ctx.presentation.slideMasters.getItemAt(0).layouts;
  layouts.load("items/id,items/name");
  await ctx.sync();
  for (const l of layouts.items) {
    const sl = ctx.presentation.slides;
    sl.load("items/id");
    await ctx.sync();
    const before = sl.items.map((x) => x.id);
    ctx.presentation.slides.add({ layoutId: l.id });
    await ctx.sync();
    sl.load("items/id");
    await ctx.sync();
    const id = sl.items.map((x) => x.id).find((x) => !before.includes(x));
    const s = ctx.presentation.slides.getItem(id);
    s.shapes.load("items/name,items/type,items/left,items/top,items/width,items/height");
    await ctx.sync();
    out.push({ boCuc: l.name, hinh: s.shapes.items.map((x) => `${x.name} [${x.type}] ${box(x).join(", ")}`) });
    s.delete();
    await ctx.sync();
  }
});
return out;
