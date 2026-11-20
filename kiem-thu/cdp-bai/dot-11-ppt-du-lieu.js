// Đợt 11: dựng lại bài trình chiếu thử về đúng trạng thái đầu trước mỗi bài (bài trước / sau so được với nhau mà không phải mở lại PowerPoint).
// Xóa mọi slide rồi thêm 4 slide theo bố cục của mẫu; slide 3 có một ô chữ cỡ 9 (lỗi cài sẵn cho bài kiểm cả bài), slide 4 có chữ mẫu sót.
// Chạy: node kiem-thu/cdp.mjs eval 9229/PowerPoint kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-11-ppt-du-lieu.js
return PowerPoint.run(async (ctx) => {
  const pres = ctx.presentation;
  pres.slides.load("items");
  const layouts = pres.slideMasters.getItemAt(0).layouts;
  layouts.load("items/name,items/id");
  await ctx.sync();
  for (const s of pres.slides.items) s.delete();
  await ctx.sync();
  const lay = (name) => (layouts.items.find((l) => l.name === name) || layouts.items[1]).id;
  const plan = [
    ["Title Slide", ["Dự án Khu đô thị Ánh Dương", "Báo cáo Hội đồng quản trị · Quý 4/2026"]],
    ["Title and Content", ["Tổng quan dự án", "Diện tích 42 ha tại Gia Lâm, Hà Nội\nQuy mô 1.850 căn hộ và 320 nhà liền kề\nTổng mức đầu tư 12.400 tỷ đồng"]],
    ["Title and Content", ["Tình hình bán hàng", "Đã bán 1.120 căn, đạt 52% kế hoạch\nGiá bình quân 58 triệu đồng/m²\nTỷ lệ hủy cọc 3%"]],
    ["Title and Content", ["Kiến nghị", "Click to add text"]],
  ];
  for (const [l] of plan) pres.slides.add({ layoutId: lay(l) });
  await ctx.sync();
  pres.slides.load("items");
  await ctx.sync();
  for (let i = 0; i < plan.length; i++) {
    const shapes = pres.slides.items[i].shapes;
    shapes.load("items/name");
    await ctx.sync();
    plan[i][1].forEach((t, j) => {
      if (shapes.items[j]) shapes.items[j].textFrame.textRange.text = t;
    });
  }
  const tb = pres.slides.items[2].shapes.addTextBox("Nguồn: Phòng Kinh doanh, số liệu đến 15/11/2026", { left: 60, top: 470, width: 400, height: 24 });
  tb.name = "Ghi chú nguồn";
  tb.textFrame.textRange.font.size = 9;
  await ctx.sync();
  const first = pres.slides.getItemAt(0);
  first.load("id");
  await ctx.sync();
  pres.setSelectedSlides([first.id]);
  await ctx.sync();
  return { ok: true, slides: plan.length, layouts: layouts.items.map((l) => l.name).slice(0, 12) };
});
