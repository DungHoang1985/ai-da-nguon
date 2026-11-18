// Đọc phông của từng hình có chữ trên mọi slide bằng Office.js (font.name của cả khung và của từng đoạn), để đối chiếu với kết quả ppt_check_deck.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-7-doc-phong.js
return await PowerPoint.run(async (ctx) => {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const out = [];
  for (let i = 0; i < slides.items.length; i++) {
    const s = slides.items[i];
    s.layout.load("name");
    s.shapes.load("items/id,items/name,items/type");
    await ctx.sync();
    const hinh = [];
    for (const sh of s.shapes.items) {
      if (!["Placeholder", "TextBox", "GeometricShape"].includes(String(sh.type))) continue;
      try {
        const tr = sh.textFrame.textRange;
        tr.load("text");
        tr.font.load("name,size");
        await ctx.sync();
        const text = String(tr.text || "");
        if (!text.trim()) {
          hinh.push({ ten: sh.name, rong: true });
          continue;
        }
        const doan = [];
        let at = 0;
        for (const line of text.split(/\r\n|\r|\n|\v/)) {
          if (line.length) {
            const r = tr.getSubstring(at, line.length);
            r.font.load("name,size");
            await ctx.sync();
            doan.push({ chu: line.slice(0, 30), phong: r.font.name, co: r.font.size });
          }
          at += line.length + 1;
        }
        hinh.push({ ten: sh.name, phongCaKhung: tr.font.name, co: tr.font.size, doan });
      } catch (e) {
        hinh.push({ ten: sh.name, loi: String(e.code || e.message) });
      }
    }
    out.push({ slide: i + 1, boCuc: s.layout.name, hinh });
  }
  return out;
});
