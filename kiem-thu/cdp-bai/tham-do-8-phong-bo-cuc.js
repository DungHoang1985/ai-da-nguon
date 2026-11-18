// Thăm dò sau 4.7.1: phông của ô tiêu đề trên các bố cục (ppt_add_chart lấy phông này trả cho tiêu đề slide biểu đồ).
// Bài dựng bằng pptxgenjs (kiem-thu/ppt-thu.mjs) trượt bước "Đợt 7 · Chữ có ư…" cả trên bản 4.7.1: xem bố cục báo phông gì.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-phong-bo-cuc.js
const out = [];
await PowerPoint.run(async (ctx) => {
  const layouts = ctx.presentation.slideMasters.getItemAt(0).layouts;
  layouts.load("items/name");
  await ctx.sync();
  for (const l of layouts.items) {
    l.shapes.load("items/name,items/type");
    await ctx.sync();
    const hs = l.shapes.items.filter((x) => String(x.type) === "Placeholder");
    hs.forEach((x) => {
      x.placeholderFormat.load("type");
      x.textFrame.textRange.font.load("name");
      x.textFrame.textRange.load("text");
    });
    await ctx.sync();
    out.push({ boCuc: l.name, o: hs.map((x) => `${x.name} [${x.placeholderFormat.type}] phông=${JSON.stringify(x.textFrame.textRange.font.name)} chữ=${JSON.stringify(x.textFrame.textRange.text.slice(0, 20))}`) });
  }
  const m = ctx.presentation.slideMasters.getItemAt(0);
  m.shapes.load("items/name,items/type");
  await ctx.sync();
  m.shapes.items.forEach((x) => x.textFrame.textRange.font.load("name"));
  try {
    await ctx.sync();
    out.push({ master: m.shapes.items.map((x) => `${x.name} phông=${JSON.stringify(x.textFrame.textRange.font.name)}`) });
  } catch (e) {
    out.push({ master: "LOI " + e.message });
  }
});
return out;
