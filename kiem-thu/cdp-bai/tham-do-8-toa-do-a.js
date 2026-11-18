// Thăm dò sau 4.7.1, bước A: khổ slide và vị trí ô tiêu đề của bố cục "Chỉ tiêu đề" trong bài đang mở (như pickChartLayout của ppt_add_chart).
// In JSON { size, layout } để dựng tệp slide biểu đồ bằng đúng buildChartPptx (bước B: tham-do-8-toa-do-b.js).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-toa-do-a.js
return await PowerPoint.run(async (ctx) => {
  const p = ctx.presentation;
  let size = { w: 720, h: 405 };
  try {
    p.load("slideWidth,slideHeight");
    await ctx.sync();
    if (p.slideWidth) size = { w: p.slideWidth, h: p.slideHeight };
  } catch {
    /* bản cũ không có khổ slide: giữ 16:9 */
  }
  const layouts = p.slideMasters.getItemAt(0).layouts;
  layouts.load("items/name,items/type");
  await ctx.sync();
  const lay = layouts.items.find((l) => String(l.type) === "TitleOnly") || layouts.items.find((l) => /title only|chỉ (có )?tiêu đề/i.test(l.name));
  if (!lay) return { size, layout: { name: "Title Only", type: "TitleOnly" } };
  lay.shapes.load("items/name,items/left,items/top,items/width,items/height");
  await ctx.sync();
  const t = lay.shapes.items.find((x) => /title|tiêu đề/i.test(x.name) && !/sub|phụ/i.test(x.name));
  return { size, layout: { name: lay.name, type: String(lay.type), title: t ? { left: t.left, top: t.top, width: t.width, height: t.height } : undefined } };
});
