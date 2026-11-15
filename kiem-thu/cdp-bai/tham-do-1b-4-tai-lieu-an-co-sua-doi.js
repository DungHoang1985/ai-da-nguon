// Thăm dò Đợt 1B (bước 2a, phần 3): tài liệu ẩn (createDocument) khi tệp kia đang có sửa đổi được theo dõi và chú thích:
// đọc bản hiện tại được không, có đọc được chú thích, đầu/chân trang của tệp kia không.
// Ghép trước một tệp định nghĩa B64 (bản "sach": có 1 cặp sửa đổi của Đối tác ở 8.3, 3 chú thích cuối trang).
const out = {};
const tryIt = async (name, fn) => {
  try {
    out[name] = await fn();
  } catch (e) {
    out[name] = { LOI: e.message, ma: e.code, o: e.debugInfo && e.debugInfo.errorLocation };
  }
};
await tryIt("1. text + getText + getReviewedText trên đoạn của tài liệu ẩn", () =>
  Word.run(async (ctx) => {
    const doc = ctx.application.createDocument(B64);
    const ps = doc.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const k = ps.items.findIndex((p) => /đề nghị cải tạo trong vòng/.test(p.text));
    const n = ps.items.findIndex((p) => /^“Tiền Thuê” là/.test(p.text));
    const res = { doan83: JSON.stringify(ps.items[k].text.slice(0, 70)), doan14: JSON.stringify(ps.items[n].text.slice(-25)) };
    const a = ps.items[k].getText();
    const b = ps.items[k].getReviewedText("Current");
    const c = ps.items[k].getReviewedText("Original");
    const d = ps.items[n].getText();
    try {
      await ctx.sync();
      res.getText = JSON.stringify(a.value.slice(0, 70));
      res.current = JSON.stringify(b.value.slice(0, 70));
      res.original = JSON.stringify(c.value.slice(0, 70));
      res.getText14 = JSON.stringify(d.value.slice(-25));
    } catch (e) {
      res.LOI_getText = e.message;
    }
    return res;
  })
);
await tryIt("2. getTrackedChanges của thân tài liệu ẩn", () =>
  Word.run(async (ctx) => {
    const doc = ctx.application.createDocument(B64);
    const tc = doc.body.getTrackedChanges();
    tc.load("items/type,items/text");
    await ctx.sync();
    return tc.items.map((t) => `${t.type}: ${t.text}`);
  })
);
await tryIt("3. chú thích, đầu/chân trang của tài liệu ẩn", () =>
  Word.run(async (ctx) => {
    const doc = ctx.application.createDocument(B64);
    const fn = doc.body.footnotes;
    fn.load("items/body/text");
    const secs = doc.sections;
    secs.load("items");
    await ctx.sync();
    const h = secs.items[0].getHeader("Primary");
    h.load("text");
    await ctx.sync();
    return { chuThich: fn.items.map((x) => x.body.text.slice(0, 40)), dauTrang: h.text };
  })
);
await tryIt("4. tài liệu đang mở: getText của đoạn có mốc chú thích", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const n = ps.items.findIndex((p) => /^“Tiền Thuê” là/.test(p.text));
    const a = ps.items[n].getText();
    const b = ps.items[n].getReviewedText("Current");
    const c = ps.items[n].getText({ IncludeHiddenText: false, IncludeTextMarkedAsDeleted: true });
    await ctx.sync();
    return { text: JSON.stringify(ps.items[n].text.slice(-20)), getText: JSON.stringify(a.value.slice(-20)), current: JSON.stringify(b.value.slice(-20)), coChuDaXoa: JSON.stringify(c.value.slice(-20)) };
  })
);
return out;
