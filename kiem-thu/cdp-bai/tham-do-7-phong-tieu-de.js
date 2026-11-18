// Thăm dò Đợt 7: vì sao tiêu đề của slide vừa thêm có lúc mang phông Times New Roman thay vì phông chủ đề.
// Thêm lần lượt các slide bố cục "Title and Content" bằng Office.js thô (đúng cách công cụ ppt_add_slide làm), đặt chữ tiêu đề theo
// nhiều cách, rồi đọc phông của từng ký tự. Xong thì xóa các slide thử.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/tham-do-7-phong-tieu-de.js
const CA = [
  { ten: "A không dấu", chu: "Tong quan doanh thu", cach: "text" },
  { ten: "B có ổ", chu: "Tổng quan doanh thu", cach: "text" },
  { ten: "C có ậ, ướ", chu: "Nhận xét xu hướng", cach: "text" },
  { ten: "D chỉ ư", chu: "Xu huong tư", cach: "text" },
  { ten: "E chỉ ậ", chu: "Nhan xet nhận", cach: "text" },
  { ten: "F có ý", chu: "Doanh thu theo quý", cach: "text" },
  { ten: "G lặp lại B", chu: "Tổng quan doanh thu", cach: "text" },
  { ten: "H ậ ướ, đặt phông chủ đề sau khi ghi", chu: "Nhận xét xu hướng", cach: "text+font" },
  { ten: "I ậ ướ, ghi bằng insertText", chu: "Nhận xét xu hướng", cach: "insert" },
];
const phongTheoKyTu = async (ctx, tr, text) => {
  const rs = [];
  for (let i = 0; i < text.length; i++) {
    const r = tr.getSubstring(i, 1);
    r.font.load("name");
    rs.push(r);
  }
  await ctx.sync();
  const runs = [];
  rs.forEach((r, i) => {
    const last = runs[runs.length - 1];
    if (last && last.phong === r.font.name) last.chu += text[i];
    else runs.push({ phong: r.font.name, chu: text[i] });
  });
  return runs.map((x) => `[${x.phong}] ${x.chu}`);
};
return await PowerPoint.run(async (ctx) => {
  const master = ctx.presentation.slideMasters.getItemAt(0);
  master.load("id");
  master.layouts.load("items/id,items/name");
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const base = slides.items.length;
  const layout = master.layouts.items.find((l) => /^title and content$/i.test(l.name)) || master.layouts.items[1];
  const out = [];
  for (const c of CA) {
    slides.add({ slideMasterId: master.id, layoutId: layout.id });
    await ctx.sync();
    slides.load("items/id");
    await ctx.sync();
    const slide = slides.items[slides.items.length - 1];
    slide.shapes.load("items/id,items/name,items/type");
    await ctx.sync();
    const title = slide.shapes.items.find((s) => /title/i.test(s.name));
    const tr = title.textFrame.textRange;
    // Phông của ô tiêu đề TRƯỚC khi ghi chữ
    tr.font.load("name");
    await ctx.sync();
    const truoc = tr.font.name;
    if (c.cach === "insert") {
      tr.text = "";
      await ctx.sync();
      title.textFrame.textRange.text = c.chu;
    } else tr.text = c.chu;
    await ctx.sync();
    if (c.cach === "text+font" && truoc) {
      title.textFrame.textRange.font.name = truoc;
      await ctx.sync();
    }
    const t2 = title.textFrame.textRange;
    t2.load("text");
    t2.font.load("name");
    await ctx.sync();
    out.push({ ca: c.ten, phongTruocKhiGhi: truoc, phongCaKhungSau: t2.font.name, theoKyTu: await phongTheoKyTu(ctx, t2, t2.text) });
  }
  // Dọn
  slides.load("items/id");
  await ctx.sync();
  for (let k = slides.items.length - 1; k >= base; k--) slides.items[k].delete();
  await ctx.sync();
  slides.load("items/id");
  await ctx.sync();
  return { boCuc: layout.name, soSlideTruoc: base, soSlideSau: slides.items.length, ketQua: out };
});
