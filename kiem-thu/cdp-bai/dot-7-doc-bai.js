// Đọc lại bài PowerPoint bằng mã sau một bài Chạy thật: số slide, và với từng slide: bố cục, các hình (tên, loại, khung, chữ rút gọn).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-7-doc-bai.js
return await PowerPoint.run(async (ctx) => {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const out = [];
  for (let i = 0; i < slides.items.length; i++) {
    const s = slides.items[i];
    s.layout.load("name");
    s.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await ctx.sync();
    const hinh = [];
    for (const sh of s.shapes.items) {
      let chu = "";
      try {
        const tr = sh.textFrame.textRange;
        tr.load("text");
        await ctx.sync();
        chu = String(tr.text || "").replace(/\s+/g, " ").slice(0, 90);
      } catch (e) {
        chu = "";
      }
      hinh.push({ ten: sh.name, loai: String(sh.type), khung: [Math.round(sh.left), Math.round(sh.top), Math.round(sh.width), Math.round(sh.height)], ...(chu ? { chu } : {}) });
    }
    out.push({ slide: i + 1, boCuc: s.layout.name, hinh });
  }
  return { soSlide: slides.items.length, slides: out };
});
