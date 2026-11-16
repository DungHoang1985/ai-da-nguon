// Thăm dò Đợt 3 (A1): các API Word cần để dựng văn bản hành chính đúng thể thức, chạy trên văn bản thử TRỐNG.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Word kiem-thu/cdp-bai/tham-do-3-vbhc.js
const sup = (set, v) => Office.context.requirements.isSetSupported(set, v);
const out = { sets: {}, kq: {} };
for (const v of ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]) out.sets["WordApi " + v] = sup("WordApi", v);
for (const v of ["1.1", "1.2", "1.3", "1.4", "1.5"]) out.sets["WordApiDesktop " + v] = sup("WordApiDesktop", v);
const mm = (x) => (x * 72) / 25.4;
const step = async (name, fn) => {
  try {
    out.kq[name] = await fn();
  } catch (e) {
    out.kq[name] = { LOI: (e && e.message) || String(e), ma: e && e.code, o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
};

await step("1. ghi lề và khổ giấy (pageSetup)", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.sections.getFirst().pageSetup;
    ps.load("topMargin,leftMargin,rightMargin,bottomMargin,paperSize,pageWidth,differentFirstPageHeaderFooter,headerDistance");
    await ctx.sync();
    const truoc = { top: ps.topMargin, left: ps.leftMargin, right: ps.rightMargin, bottom: ps.bottomMargin, paper: ps.paperSize, w: ps.pageWidth, first: ps.differentFirstPageHeaderFooter };
    ps.paperSize = "A4";
    ps.topMargin = mm(20);
    ps.bottomMargin = mm(20);
    ps.leftMargin = mm(30);
    ps.rightMargin = mm(20);
    ps.differentFirstPageHeaderFooter = true;
    await ctx.sync();
    ps.load("topMargin,leftMargin,rightMargin,bottomMargin,paperSize,pageWidth,differentFirstPageHeaderFooter");
    await ctx.sync();
    return { truoc, sau: { top: ps.topMargin, left: ps.leftMargin, right: ps.rightMargin, bottom: ps.bottomMargin, paper: ps.paperSize, w: ps.pageWidth, first: ps.differentFirstPageHeaderFooter } };
  })
);

await step("2. số trang ở đầu trang (trường PAGE), canh giữa", () =>
  Word.run(async (ctx) => {
    const h = ctx.document.sections.getFirst().getHeader("Primary");
    const p = h.paragraphs.getFirst();
    p.alignment = "Centered";
    const f = p.getRange("Start").insertField("Start", "Page");
    f.load("code,type");
    await ctx.sync();
    h.load("text");
    await ctx.sync();
    return { code: f.code, type: f.type, text: h.text };
  })
);

await step("3. bảng 2 cột không viền, khi đang theo dõi thay đổi", () =>
  Word.run(async (ctx) => {
    ctx.document.changeTrackingMode = "TrackAll";
    await ctx.sync();
    const t = ctx.document.body.paragraphs.getLast().insertTable(2, 2, "After", [
      ["TÊN CƠ QUAN CHỦ QUẢN", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"],
      ["Số:      /CV-ABC", "Hà Nội, ngày      tháng      năm 2026"],
    ]);
    t.getBorder("All").type = "None";
    t.load("rowCount,width,styleBuiltIn");
    await ctx.sync();
    const c00 = t.getCell(0, 0);
    const c01 = t.getCell(0, 1);
    c00.columnWidth = mm(60);
    c01.columnWidth = mm(100);
    c00.horizontalAlignment = "Centered";
    c01.horizontalAlignment = "Centered";
    await ctx.sync();
    // Thêm đoạn thứ hai trong ô và định dạng từng đoạn
    const p1 = c00.body.paragraphs.getFirst();
    p1.font.size = 13;
    p1.font.name = "Times New Roman";
    const p2 = p1.insertParagraph("TÊN CƠ QUAN BAN HÀNH", "After");
    p2.font.bold = true;
    p2.font.size = 13;
    const q1 = c01.body.paragraphs.getFirst();
    q1.font.bold = true;
    q1.font.size = 13;
    const q2 = q1.insertParagraph("Độc lập - Tự do - Hạnh phúc", "After");
    q2.font.bold = true;
    q2.font.size = 14;
    await ctx.sync();
    c00.body.paragraphs.load("items/text,items/alignment");
    c00.load("columnWidth");
    c01.load("columnWidth");
    t.load("width");
    await ctx.sync();
    return { hang: t.rowCount, rong: t.width, cot0: c00.columnWidth, cot1: c01.columnWidth, doanO00: c00.body.paragraphs.items.map((x) => [x.text, x.alignment]) };
  })
);

await step("4. đường kẻ = viền dưới của một đoạn, thu ngắn bằng thụt lề", () =>
  Word.run(async (ctx) => {
    const t = ctx.document.body.tables.getFirst();
    const c = t.getCell(0, 1);
    const last = c.body.paragraphs.getLast();
    const line = last.insertParagraph("", "After");
    line.font.size = 2;
    line.leftIndent = mm(25);
    line.rightIndent = mm(25);
    const bs = line.borders;
    bs.load("items");
    await ctx.sync();
    const n = bs.items.length;
    let cach = "";
    try {
      const b = bs.getByLocation ? bs.getByLocation("Bottom") : null;
      if (b) {
        b.isVisible = true;
        b.lineStyle = "Single";
        b.lineWidth = "Pt075";
        await ctx.sync();
        cach = "getByLocation(Bottom)";
      }
    } catch (e) {
      cach = "getByLocation lỗi: " + e.message;
    }
    return { soVien: n, cach, coGetByLocation: typeof bs.getByLocation };
  })
);

await step("4b. viền dưới qua items[]", () =>
  Word.run(async (ctx) => {
    const t = ctx.document.body.tables.getFirst();
    const c = t.getCell(0, 0);
    const line = c.body.paragraphs.getLast().insertParagraph("", "After");
    line.font.size = 2;
    const bs = line.borders;
    bs.load("items/isVisible,items/lineStyle");
    await ctx.sync();
    const truoc = bs.items.map((b) => [b.isVisible, b.lineStyle]);
    // thử từng viền: bật viền thứ k rồi đọc lại xem là viền nào
    const kq = [];
    for (let k = 0; k < bs.items.length; k++) {
      try {
        bs.items[k].lineStyle = "Single";
        bs.items[k].lineWidth = "Pt075";
        await ctx.sync();
        kq.push(k + ": được");
        if (k !== 2) {
          bs.items[k].lineStyle = "None";
          await ctx.sync();
        }
      } catch (e) {
        kq.push(k + ": " + e.message);
      }
    }
    return { truoc, kq };
  })
);

await step("5. thụt đầu dòng, giãn dòng, cách đoạn, giữ với đoạn sau", () =>
  Word.run(async (ctx) => {
    const p = ctx.document.body.insertParagraph("Nội dung thử của công văn, canh đều hai lề, thụt đầu dòng một xăng-ti-mét, cách đoạn sáu point, giãn dòng một phẩy hai.", "End");
    p.font.name = "Times New Roman";
    p.font.size = 14;
    p.font.bold = false;
    p.alignment = "Justified";
    p.firstLineIndent = mm(10);
    p.leftIndent = 0;
    p.rightIndent = 0;
    p.spaceAfter = 6;
    p.spaceBefore = 0;
    p.lineSpacing = 14 * 1.2;
    await ctx.sync();
    p.load("firstLineIndent,spaceAfter,lineSpacing,alignment,lineUnitAfter");
    await ctx.sync();
    return { firstLineIndent: p.firstLineIndent, spaceAfter: p.spaceAfter, lineSpacing: p.lineSpacing, alignment: p.alignment };
  })
);

await step("6. sửa đổi được theo dõi sinh ra và trạng thái", () =>
  Word.run(async (ctx) => {
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type");
    ctx.document.body.paragraphs.load("items/text");
    ctx.document.body.tables.load("items/rowCount");
    await ctx.sync();
    return { suaDoi: tc.items.length, loai: [...new Set(tc.items.map((x) => x.type))], soDoan: ctx.document.body.paragraphs.items.length, soBang: ctx.document.body.tables.items.length };
  })
);
return out;
