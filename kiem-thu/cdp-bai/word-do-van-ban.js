// Đo văn bản đang mở trong Word: số đoạn, số trang, số sửa đổi và nhận xét, thời gian của các lối đọc mà công cụ Word dùng.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/word-do-van-ban.js
const out = {};
const t = async (name, fn) => {
  const t0 = performance.now();
  try {
    out[name] = await fn();
  } catch (e) {
    out[name] = "LỖI " + e.message + " @ " + (e.debugInfo && e.debugInfo.errorLocation);
  }
  out[name + " (giây)"] = Math.round(performance.now() - t0) / 1000;
};
await t("đoạn + nhãn số + mã danh sách, một lần sync", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text,items/style,items/outlineLevel,items/tableNestingLevel,items/uniqueLocalId,items/listItemOrNullObject/listString,items/listItemOrNullObject/level,items/listOrNullObject/id");
    await ctx.sync();
    return { soDoan: ps.items.length, kyTu: ps.items.reduce((n, p) => n + p.text.length, 0), coSo: ps.items.filter((p) => !p.listItemOrNullObject.isNullObject).length };
  })
);
await t("getText() cho mọi đoạn", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items");
    await ctx.sync();
    const r = ps.items.map((p) => p.getText());
    await ctx.sync();
    return r.length;
  })
);
await t("sửa đổi và nhận xét", () =>
  Word.run(async (ctx) => {
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type,items/author");
    const cs = ctx.document.body.getComments();
    cs.load("items/authorName");
    await ctx.sync();
    const by = {};
    tc.items.forEach((x) => (by[x.type] = (by[x.type] || 0) + 1));
    return { suaDoi: tc.items.length, theoLoai: by, tacGiaSuaDoi: [...new Set(tc.items.map((x) => (x.author === "Đối tác" ? x.author : "(tài khoản Office)")))], nhanXet: cs.items.length };
  })
);
await t("số trang (WordApiDesktop)", () =>
  Word.run(async (ctx) => {
    const pages = ctx.document.activeWindow.activePane.pages;
    pages.load("items");
    await ctx.sync();
    return pages.items.length;
  })
);
return out;
