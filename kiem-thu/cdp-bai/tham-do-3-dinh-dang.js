// Thăm dò Đợt 1A (bước 3): chữ đậm giữa câu sau khi sửa mức từ; thay cả đoạn trên đoạn đang có sửa đổi của đối tác;
// Document.undo; getText trên mọi đoạn trong một lần sync (tốc độ).
// Chạy trên văn bản thăm dò còn nguyên (mở lại tệp): node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-3-dinh-dang.js
const out = {};
const who = (a) => (a === "Đối tác" ? a : "(tài khoản Office đang dùng)");
const test = async (name, fn) => {
  try {
    out[name] = await fn();
  } catch (e) {
    out[name] = { LOI: (e && e.message) || String(e), o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
};
const find = async (ctx, start) => {
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  await ctx.sync();
  const i = ps.items.findIndex((p) => p.text.startsWith(start));
  return { ps, i, p: ps.items[i] };
};

await test("a. chữ đậm sau khi đổi '30' thứ hai thành '45'", () =>
  Word.run(async (ctx) => {
    const { p } = await find(ctx, "Bên B thanh toán trong 30 ngày");
    ctx.document.changeTrackingMode = "TrackAll";
    const h = p.search("30", { matchCase: true });
    h.load("items");
    await ctx.sync();
    h.items[1].insertText("45", "Replace");
    await ctx.sync();
    const parts = ["45", "ngày thì", "trong", "phạt"].map((s) => {
      const r = p.search(s, { matchCase: true });
      r.load("items/font/bold,items/text");
      return r;
    });
    const tc = p.getTrackedChanges();
    tc.load("items/type,items/text");
    await ctx.sync();
    ctx.document.changeTrackingMode = "Off";
    await ctx.sync();
    return {
      dam: Object.fromEntries(parts.map((r) => [r.items[0] ? r.items[0].text : "?", r.items.map((x) => x.font.bold)])),
      suaDoi: tc.items.map((t) => `${t.type} ${JSON.stringify(t.text)}`),
    };
  })
);

await test("b. thay cả đoạn trên đoạn đang có sửa đổi của đối tác (TrackAll)", () =>
  Word.run(async (ctx) => {
    const { p } = await find(ctx, "Thời hạn thuê là 24 tháng12 tháng");
    ctx.document.changeTrackingMode = "TrackAll";
    await ctx.sync();
    p.insertText("Thời hạn thuê là 36 tháng kể từ ngày bàn giao mặt bằng.", "Replace");
    await ctx.sync();
    const cur = p.getText();
    const orig = p.getReviewedText("Original");
    const tc = p.getTrackedChanges();
    tc.load("items/type,items/text,items/author");
    p.load("text");
    await ctx.sync();
    ctx.document.changeTrackingMode = "Off";
    await ctx.sync();
    return { banHienTai: cur.value, banGoc: orig.value, ".text": p.text, suaDoi: tc.items.map((t) => `${t.type} ${JSON.stringify(t.text)} · ${who(t.author)}`) };
  })
);

await test("c. getText() cho mọi đoạn trong một lần sync", () =>
  Word.run(async (ctx) => {
    const t0 = performance.now();
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const r = ps.items.map((p) => p.getText());
    await ctx.sync();
    const khac = ps.items.map((p, i) => (p.text !== r[i].value.replace(/\r$/, "") ? i : -1)).filter((i) => i >= 0);
    return { soDoan: ps.items.length, giay: Math.round(performance.now() - t0) / 1000, doanKhacText: khac };
  })
);

await test("d. Document.undo sau một lần sửa bằng add-in", () =>
  Word.run(async (ctx) => {
    const { p } = await find(ctx, "HẾT");
    p.insertText("HẾT (sửa thử để hoàn tác)", "Replace");
    await ctx.sync();
    const r = ctx.document.undo(1);
    await ctx.sync();
    p.load("text");
    await ctx.sync();
    return { undoTraVe: r.value, sauUndo: p.text };
  })
);
return out;
