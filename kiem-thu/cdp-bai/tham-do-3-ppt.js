// Thăm dò Đợt 3 (B3, B4): các API PowerPoint cho ảnh slide, màu chủ đề, bố cục và ô giữ chỗ, định dạng chữ theo đoạn, bảng, vùng chọn,
// đổi bố cục, chèn ảnh, ghi chú diễn giả, phông chủ đề, biểu đồ. Chạy trên bài thử (thêm một slide thử rồi xóa).
// Chạy: node kiem-thu/cdp.mjs eval 9229/Powerpoint kiem-thu/cdp-bai/tham-do-3-ppt.js
const sup = (v) => Office.context.requirements.isSetSupported("PowerPointApi", v);
const out = { sets: {}, kq: {} };
for (const v of ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"]) out.sets[v] = sup(v);
const step = async (name, fn) => {
  const t0 = performance.now();
  try {
    out.kq[name] = await fn();
  } catch (e) {
    out.kq[name] = { LOI: (e && e.message) || String(e), ma: e && e.code, o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
  out.kq[name] = { ms: Math.round(performance.now() - t0), ...(typeof out.kq[name] === "object" ? out.kq[name] : { v: out.kq[name] }) };
};
let slideId = "";

await step("0. dựng slide thử: tiêu đề, hộp chữ 3 đoạn, bảng 3×3", () =>
  PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.add();
    await ctx.sync();
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const sl = slides.items[slides.items.length - 1];
    slideId = sl.id;
    const tb = sl.shapes.addTextBox("Đoạn một của hộp chữ thử\nĐoạn hai có chữ cần tô đậm\nĐoạn ba", { left: 40, top: 150, width: 300, height: 120 });
    tb.name = "aidn-thu-hop";
    const t = sl.shapes.addTable(3, 3, { values: [["Hạng mục", "Kế hoạch", "Thực tế"], ["Móng", "100", "90"], ["Thân", "200", "150"]], left: 380, top: 150, width: 300, height: 120 });
    t.name = "aidn-thu-bang";
    await ctx.sync();
    sl.shapes.load("items/name,items/type,items/id");
    await ctx.sync();
    return { slides: slides.items.length, shapes: sl.shapes.items.map((s) => `${s.name}:${s.type}`) };
  })
);

await step("1. ảnh slide (getImageAsBase64, rộng 800)", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    const img = sl.getImageAsBase64({ width: 800 });
    await ctx.sync();
    const b = img.value;
    return { kyTu: b.length, dau: b.slice(0, 12), laPng: b.startsWith("iVBOR"), laJpeg: b.startsWith("/9j/") };
  })
);

await step("2. màu chủ đề (themeColorScheme.getThemeColor)", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    const names = ["Dark1", "Light1", "Dark2", "Light2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"];
    const rs = names.map((n) => sl.themeColorScheme.getThemeColor(n));
    await ctx.sync();
    return Object.fromEntries(names.map((n, i) => [n, rs[i].value]));
  })
);

await step("3. bố cục của mẫu và ô giữ chỗ", () =>
  PowerPoint.run(async (ctx) => {
    const masters = ctx.presentation.slideMasters;
    masters.load("items/name,items/id");
    await ctx.sync();
    const m = masters.items[0];
    m.layouts.load("items/name,items/id,items/type");
    await ctx.sync();
    const lays = m.layouts.items.slice(0, 12);
    lays.forEach((l) => l.shapes.load("items/name,items/type"));
    await ctx.sync();
    const ph = [];
    for (const l of lays.slice(0, 4)) {
      const got = l.shapes.items.filter((s) => s.type === "Placeholder");
      got.forEach((s) => s.placeholderFormat.load("type"));
      await ctx.sync();
      ph.push(`${l.name}: ${got.map((s) => s.placeholderFormat.type).join(", ")}`);
    }
    return { masters: masters.items.map((x) => x.name), layouts: m.layouts.items.map((l) => `${l.name}${l.type ? " (" + l.type + ")" : ""}`), oGiuCho: ph };
  })
);

await step("4. định dạng chữ theo từng đoạn và từng cụm (getSubstring, paragraphFormat)", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    sl.shapes.load("items/name");
    await ctx.sync();
    const tb = sl.shapes.items.find((s) => s.name === "aidn-thu-hop");
    const tr = tb.textFrame.textRange;
    tr.load("text");
    await ctx.sync();
    const text = tr.text;
    const at = text.indexOf("tô đậm");
    const sub = tr.getSubstring(at, "tô đậm".length);
    sub.font.bold = true;
    sub.font.color = "#C00000";
    sub.font.size = 20;
    const p2start = text.indexOf("Đoạn hai");
    const p2 = tr.getSubstring(p2start, text.indexOf("Đoạn ba") - p2start - 1);
    p2.paragraphFormat.horizontalAlignment = "Center";
    p2.paragraphFormat.bulletFormat.visible = true;
    await ctx.sync();
    sub.font.load("bold,color,size,name");
    p2.paragraphFormat.load("horizontalAlignment");
    p2.paragraphFormat.bulletFormat.load("visible");
    const p1 = tr.getSubstring(0, 8);
    p1.font.load("bold,size,name");
    await ctx.sync();
    return { xuongDong: JSON.stringify(text.match(/[\r\n\v]+/g)), cum: { bold: sub.font.bold, color: sub.font.color, size: sub.font.size }, doan2: { canh: p2.paragraphFormat.horizontalAlignment, bullet: p2.paragraphFormat.bulletFormat.visible }, doan1: { bold: p1.font.bold, size: p1.font.size, name: p1.font.name } };
  })
);

await step("5. bảng có sẵn: đọc, sửa một ô, thêm hàng", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    sl.shapes.load("items/name,items/type");
    await ctx.sync();
    const sh = sl.shapes.items.find((s) => s.name === "aidn-thu-bang");
    const t = sh.getTable();
    t.load("rowCount,columnCount,values");
    await ctx.sync();
    const truoc = t.values;
    const c = t.getCellOrNullObject(1, 2);
    c.text = "95";
    c.font.bold = true;
    c.fill.setSolidColor("#FFF2CC");
    await ctx.sync();
    let themHang = "";
    try {
      t.rows.add(null, 1);
      await ctx.sync();
      themHang = "được";
    } catch (e) {
      themHang = "LỖI " + e.message;
    }
    t.load("rowCount,values");
    await ctx.sync();
    return { loai: sh.type, truoc, sau: t.values, soHang: t.rowCount, themHang };
  })
);

await step("6. vùng chọn: hình và đoạn chữ đang chọn", () =>
  PowerPoint.run(async (ctx) => {
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load("items/name");
    const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
    tr.load("text,isNullObject");
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    await ctx.sync();
    return { hinh: shapes.items.map((s) => s.name), chu: tr.isNullObject ? null : tr.text, slide: slides.items.length };
  })
);

await step("6b. chọn hình bằng mã rồi đọc lại (setSelectedShapes)", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    sl.shapes.load("items/name,items/id");
    await ctx.sync();
    const tb = sl.shapes.items.find((s) => s.name === "aidn-thu-hop");
    ctx.presentation.setSelectedSlides([slideId]);
    await ctx.sync();
    sl.setSelectedShapes([tb.id]);
    await ctx.sync();
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load("items/name");
    await ctx.sync();
    return { hinh: shapes.items.map((s) => s.name) };
  })
);

await step("7. đổi bố cục slide (applyLayout)", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    sl.layout.load("name,id");
    const m = ctx.presentation.slideMasters.getItemAt(0);
    m.layouts.load("items/name,items/id");
    await ctx.sync();
    const truoc = sl.layout.name;
    const target = m.layouts.items.find((l) => /title only|chỉ tiêu đề/i.test(l.name)) || m.layouts.items.find((l) => l.name !== truoc);
    sl.applyLayout(target);
    await ctx.sync();
    sl.layout.load("name");
    sl.shapes.load("items/name,items/type");
    await ctx.sync();
    return { truoc, muon: target.name, sau: sl.layout.name, hinh: sl.shapes.items.map((s) => `${s.name}:${s.type}`) };
  })
);

await step("8. chèn ảnh: hình chữ nhật tô bằng ảnh (fill.setImage)", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    const png = "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAFklEQVR4nGP8z8Dwn4EIwESMolGF+AEAVjgCDkVgCwQAAAAASUVORK5CYII=";
    const r = sl.shapes.addGeometricShape("Rectangle", { left: 40, top: 300, width: 80, height: 80 });
    r.name = "aidn-thu-anh";
    r.fill.setImage(png);
    r.lineFormat.visible = false;
    await ctx.sync();
    r.fill.load("type");
    await ctx.sync();
    return { fill: r.fill.type };
  })
);

await step("9. ghi chú diễn giả, phông chủ đề: có API không", async () => {
  const proto = Object.getOwnPropertyNames(PowerPoint.Slide.prototype);
  return { thuocTinhSlideCoNotes: proto.filter((k) => /note/i.test(k)), coThemeFont: Object.getOwnPropertyNames(PowerPoint.SlideMaster.prototype).filter((k) => /font|theme/i.test(k)), biểuĐồ: Object.getOwnPropertyNames(PowerPoint.Shape.prototype).filter((k) => /chart/i.test(k)) };
});

await step("10. xuất slide (exportAsBase64) để đọc phông chủ đề, ghi chú, biểu đồ từ gói tệp", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    const b = sl.exportAsBase64();
    await ctx.sync();
    return { kyTu: b.value.length, dau: b.value.slice(0, 8) };
  })
);

await step("11. chữ tràn khung: đọc autoSize, kích thước hình và chữ", () =>
  PowerPoint.run(async (ctx) => {
    const sl = ctx.presentation.slides.getItem(slideId);
    const tb = sl.shapes.addTextBox("Chữ rất dài cố ý để tràn khỏi khung. ".repeat(12), { left: 400, top: 320, width: 200, height: 40 });
    tb.name = "aidn-thu-tran";
    tb.textFrame.autoSizeSetting = "AutoSizeNone";
    tb.textFrame.wordWrap = true;
    await ctx.sync();
    tb.load("left,top,width,height");
    tb.textFrame.load("autoSizeSetting,wordWrap,leftMargin,rightMargin,topMargin,bottomMargin,hasText");
    tb.textFrame.textRange.font.load("size,name");
    await ctx.sync();
    return { hinh: { w: tb.width, h: tb.height }, khung: { autoSize: tb.textFrame.autoSizeSetting, wrap: tb.textFrame.wordWrap, le: [tb.textFrame.leftMargin, tb.textFrame.topMargin] }, co: tb.textFrame.textRange.font.size, phong: tb.textFrame.textRange.font.name };
  })
);

await step("99. xóa slide thử", () =>
  PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.getItem(slideId).delete();
    await ctx.sync();
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    return { slides: slides.items.length };
  })
);
return out;
