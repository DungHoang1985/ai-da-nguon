// Thăm dò (b) sau 4.7.1: chữ có “ư” ghi vào ô chữ, hình, bảng (thêm mới và sửa ô) bằng Office.js — có thành Times New Roman không,
// và đặt lại phông có cứu được không. Ghi đúng cách các công cụ ppt_add_textbox, ppt_add_shape (cả sơ đồ quy trình), ppt_add_table, ppt_edit_table làm.
// Thêm một slide thử (bố cục đầu tiên có ô tiêu đề), chạy các ca, đọc phông theo từng ký tự, rồi xóa slide thử.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-chu-u.js
const CHU = "Nhận xét xu hướng tư";
const KHONG_U = "Tong quan doanh thu";
const out = { ca: {} };
let slideId = "";
const runs = async (ctx, tr) => {
  tr.load("text");
  await ctx.sync();
  const text = tr.text || "";
  const rs = [];
  for (let i = 0; i < text.length; i++) {
    const r = tr.getSubstring(i, 1);
    r.font.load("name");
    rs.push(r);
  }
  await ctx.sync();
  const o = [];
  rs.forEach((r, i) => {
    const last = o[o.length - 1];
    if (last && last.p === r.font.name) last.c += text[i];
    else o.push({ p: r.font.name, c: text[i] });
  });
  return o.map((x) => `[${x.p}] ${x.c}`).join(" | ");
};
const ca = async (ten, fn) => {
  try {
    out.ca[ten] = await PowerPoint.run(fn);
  } catch (e) {
    out.ca[ten] = { LOI: (e && e.message) || String(e) };
  }
};
try {
  await PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides;
    sl.load("items/id");
    await ctx.sync();
    const before = sl.items.map((x) => x.id);
    ctx.presentation.slides.add();
    await ctx.sync();
    sl.load("items/id");
    await ctx.sync();
    slideId = sl.items.map((x) => x.id).find((x) => !before.includes(x));
  });
  const S = (ctx) => ctx.presentation.slides.getItem(slideId);

  // Ô chữ: chữ đưa thẳng vào addTextBox (cách ppt_add_textbox làm)
  for (const [ten, chu] of [["1a ô chữ, có ư", CHU], ["1b ô chữ, không dấu", KHONG_U]]) {
    await ca(ten, async (ctx) => {
      const shp = S(ctx).shapes.addTextBox(chu, { left: 40, top: 40, width: 400, height: 50 });
      await ctx.sync();
      const truoc = await runs(ctx, shp.textFrame.textRange);
      const f = shp.textFrame.textRange.font;
      f.load("name");
      await ctx.sync();
      return { phongCaKhung: f.name, phong: truoc };
    });
  }
  // Ô chữ: đọc phông khung rỗng trước, ghi chữ, đặt lại đúng phông đó
  await ca("1c ô chữ rỗng → ghi chữ có ư → đặt lại phông cũ", async (ctx) => {
    const shp = S(ctx).shapes.addTextBox("", { left: 40, top: 100, width: 400, height: 50 });
    const f0 = shp.textFrame.textRange.font;
    f0.load("name");
    await ctx.sync();
    const cu = f0.name;
    shp.textFrame.textRange.text = CHU;
    await ctx.sync();
    const sauGhi = await runs(ctx, shp.textFrame.textRange);
    shp.textFrame.textRange.font.name = cu;
    await ctx.sync();
    return { phongKhungRong: cu, sauGhi, sauDatLai: await runs(ctx, shp.textFrame.textRange) };
  });
  // Hình (ppt_add_shape, sơ đồ quy trình): ghi chữ vào hình
  await ca("2 hình chữ nhật, ghi chữ có ư", async (ctx) => {
    const shp = S(ctx).shapes.addGeometricShape("RoundRectangle", { left: 40, top: 160, width: 300, height: 60 });
    const f0 = shp.textFrame.textRange.font;
    f0.load("name");
    await ctx.sync();
    const cu = f0.name;
    shp.textFrame.textRange.text = CHU;
    await ctx.sync();
    const sauGhi = await runs(ctx, shp.textFrame.textRange);
    shp.textFrame.textRange.font.name = cu;
    await ctx.sync();
    const sauDatLai = await runs(ctx, shp.textFrame.textRange);
    shp.textFrame.textRange.font.name = "Arial";
    await ctx.sync();
    return { phongHinhRong: cu, sauGhi, sauDatLai, datArial: await runs(ctx, shp.textFrame.textRange) };
  });
  // Bảng: thêm mới với giá trị có ư, rồi sửa một ô
  await ca("3 bảng thêm mới + sửa ô", async (ctx) => {
    const shp = S(ctx).shapes.addTable(2, 2, { values: [[CHU, KHONG_U], ["Quý 1", "120"]], left: 360, top: 160 });
    await ctx.sync();
    const t = shp.getTable();
    const cells = [[0, 0], [0, 1], [1, 0]].map(([r, c]) => {
      const x = t.getCellOrNullObject(r, c);
      x.font.load("name");
      x.load("text");
      return x;
    });
    await ctx.sync();
    const themMoi = cells.map((x) => `[${x.font.name}] ${x.text}`);
    const sua = t.getCellOrNullObject(1, 1);
    sua.text = "Tăng trưởng ư";
    await ctx.sync();
    sua.font.load("name");
    sua.load("text");
    await ctx.sync();
    const suaO = `[${sua.font.name}] ${sua.text}`;
    // đặt lại phông của ô bằng phông ô không dấu
    const ref = cells[1].font.name;
    sua.font.name = ref;
    cells[0].font.name = ref;
    await ctx.sync();
    sua.font.load("name");
    cells[0].font.load("name");
    await ctx.sync();
    return { themMoi, suaO, phongOKhongDau: ref, sauDatLai: [cells[0].font.name, sua.font.name] };
  });
  // Ảnh slide (để nhìn, không in)
  await ca("anh", async (ctx) => {
    const img = S(ctx).getImageAsBase64({ width: 800 });
    await ctx.sync();
    window.__aidnAnhChuU = img.value;
    return img.value.length;
  });
} catch (e) {
  out.LOI = (e && e.message) || String(e);
} finally {
  if (slideId) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(slideId).delete();
      await ctx.sync();
    });
    out.daXoaSlideThu = true;
  }
}
return out;
