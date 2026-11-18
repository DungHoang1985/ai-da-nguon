// Thăm dò sau 4.7.1, bước B: vì sao ô tiêu đề của slide biểu đồ (ppt_add_chart) được báo ở toạ độ 0, 0, 0, 0.
// Chèn tệp slide biểu đồ dựng bằng đúng buildChartPptx (biến B64, tệp ghép trước tệp này), đọc vị trí ô tiêu đề theo nhiều cách,
// thử đặt lại vị trí bằng vị trí ô tiêu đề của bố cục, đọc lại, rồi xóa slide thử.
// Chạy: node kiem-thu/cdp.mjs eval 9229 <tệp b64.js> kiem-thu/cdp-bai/tham-do-8-toa-do-b.js
const out = {};
const box = (s) => [s.left, s.top, s.width, s.height].map((v) => Math.round(v * 10) / 10);
let freshId = "";
try {
  await PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides;
    sl.load("items/id");
    await ctx.sync();
    const before = sl.items.map((x) => x.id);
    ctx.presentation.insertSlidesFromBase64(B64, { formatting: "UseDestinationTheme", targetSlideId: before[before.length - 1] });
    await ctx.sync();
    sl.load("items/id");
    await ctx.sync();
    freshId = sl.items.map((x) => x.id).find((id) => !before.includes(id));
    const s = ctx.presentation.slides.getItem(freshId);
    s.layout.load("name");
    s.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    s.layout.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await ctx.sync();
    out.boCuc = s.layout.name;
    const ph = (list) => {
      const hs = list.filter((x) => String(x.type) === "Placeholder");
      hs.forEach((x) => x.placeholderFormat.load("type"));
      return hs;
    };
    const mine = ph(s.shapes.items);
    const model = ph(s.layout.shapes.items);
    await ctx.sync();
    out.hinhTrenSlide = s.shapes.items.map((x) => ({ ten: x.name, kieu: String(x.type), o: box(x), ph: mine.includes(x) ? String(x.placeholderFormat.type) : undefined }));
    out.oTieuDeBoCuc = model.filter((x) => /Title/.test(String(x.placeholderFormat.type))).map((x) => ({ ten: x.name, o: box(x) }));
    const t = mine.find((x) => /Title/.test(String(x.placeholderFormat.type)));
    const m = model.find((x) => /Title/.test(String(x.placeholderFormat.type)));
    // Đọc lại bằng getItem riêng (khác lần nạp theo danh sách?)
    if (t) {
      const again = s.shapes.getItem(t.id);
      again.load("left,top,width,height");
      await ctx.sync();
      out.docLaiTheoId = box(again);
    }
    // Thử đặt vị trí bằng vị trí ô tiêu đề của bố cục rồi đọc lại
    if (t && m) {
      t.left = m.left;
      t.top = m.top;
      t.width = m.width;
      t.height = m.height;
      await ctx.sync();
      const again = s.shapes.getItem(t.id);
      again.load("left,top,width,height");
      again.textFrame.textRange.load("text");
      await ctx.sync();
      out.sauKhiDat = box(again);
      out.chuTieuDe = again.textFrame.textRange.text;
    }
    // Ảnh slide để nhìn: có lệch so với trước khi đặt không (chỉ đo kích thước chuỗi, không in)
    try {
      const img = s.getImageAsBase64({ width: 640 });
      await ctx.sync();
      out.anhSauKhiDat = img.value.length;
      window.__aidnAnhThamDo8 = img.value;
    } catch (e) {
      out.anhSauKhiDat = "LOI " + e.message;
    }
  });
} catch (e) {
  out.LOI = (e && e.message) || String(e);
} finally {
  if (freshId && !window.__aidnGiuSlideThamDo8) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(freshId).delete();
      await ctx.sync();
    });
    out.daXoaSlideThu = true;
  }
}
return out;
