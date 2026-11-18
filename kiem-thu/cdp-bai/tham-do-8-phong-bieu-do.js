// Thăm dò sau 4.7.1: tiêu đề slide biểu đồ trên bài pptxgenjs (phông chủ đề Georgia) vẫn Times New Roman sau restoreFont — vì sao?
// Chèn tệp slide biểu đồ (biến B64, tệp ghép trước), đọc phông theo dòng, đặt phông Georgia như restoreFont, đọc lại cả dòng và từng ký tự.
// Chạy: node kiem-thu/cdp.mjs eval 9229 <tệp b64.js> kiem-thu/cdp-bai/tham-do-8-phong-bieu-do.js
const out = {};
let freshId = "";
const perChar = async (ctx, tr, text) => {
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
    if (last && last.p === r.font.name) last.c += text[i];
    else runs.push({ p: r.font.name, c: text[i] });
  });
  return runs.map((x) => `[${x.p}] ${x.c}`);
};
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
    s.shapes.load("items/id,items/name,items/type");
    await ctx.sync();
    const t = s.shapes.items.find((x) => String(x.type) === "Placeholder");
    const tr = t.textFrame.textRange;
    tr.load("text");
    tr.font.load("name");
    await ctx.sync();
    out.chu = tr.text;
    out.phongCaKhung = tr.font.name;
    out.truoc = await perChar(ctx, tr, tr.text);
    const line = tr.getSubstring(0, tr.text.length);
    line.font.name = "Georgia";
    await ctx.sync();
    out.sauDatCaDong = await perChar(ctx, tr, tr.text);
    for (let i = 0; i < tr.text.length; i++) tr.getSubstring(i, 1).font.name = "Georgia";
    await ctx.sync();
    out.sauDatTungKyTu = await perChar(ctx, tr, tr.text);
    tr.font.name = "Georgia";
    await ctx.sync();
    out.sauDatCaKhung = await perChar(ctx, tr, tr.text);
    tr.font.name = "Arial";
    await ctx.sync();
    out.sauDatArial = await perChar(ctx, tr, tr.text);
    const x = s.exportAsBase64();
    await ctx.sync();
    window.__aidnSlideThamDo8 = x.value;
    out.xuat = x.value.length;
  });
} catch (e) {
  out.LOI = (e && e.message) || String(e);
} finally {
  if (freshId) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(freshId).delete();
      await ctx.sync();
    });
    out.daXoaSlideThu = true;
  }
}
return out;
