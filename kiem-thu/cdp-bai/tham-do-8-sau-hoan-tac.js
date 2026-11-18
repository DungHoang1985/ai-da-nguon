// Thăm dò Đợt 8: sau khi Hoàn tác (từ chối sửa đổi) một lần điền, ô điền mẫu từ chối ghi ở lần điền sau. Đọc trạng thái từng ô
// (khóa, sửa đổi chờ duyệt bên trong, OOXML rút gọn) và thử ghi lại từng loại, báo lỗi đúng chỗ. Xong thì từ chối mọi sửa đổi vừa tạo.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-sau-hoan-tac.js
const TYPES = ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"];
const out = { o: [], thu: {} };
await Word.run(async (ctx) => {
  const cc = ctx.document.getContentControls({ types: TYPES });
  cc.load("items/id,items/tag,items/type,items/text,items/cannotEdit,items/cannotDelete");
  await ctx.sync();
  const tcs = cc.items.map((c) => {
    const t = c.getRange("Whole").getTrackedChanges();
    t.load("items/type,items/text");
    return t;
  });
  const xs = cc.items.map((c) => c.getRange("Whole").getOoxml());
  await ctx.sync();
  cc.items.forEach((c, k) => {
    const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(xs[k].value)?.[0] || "";
    out.o.push({
      tag: c.tag,
      id: c.id,
      type: c.type,
      text: c.text,
      khoa: c.cannotEdit,
      suaDoiTrong: tcs[k].items.map((x) => `${x.type}:${x.text}`),
      sdtPr: (/<w:sdtPr>[\s\S]*?<\/w:sdtPr>/.exec(sdt)?.[0] || "").replace(/<w:listItem[^>]*\/>/g, "").slice(0, 400),
      noiDung: (/<w:sdtContent>[\s\S]*?<\/w:sdtContent>/.exec(sdt)?.[0] || "").replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "").slice(0, 300),
    });
  });
});
const thu = async (ten, fn) => {
  try {
    out.thu[ten] = await Word.run(async (ctx) => {
      ctx.document.changeTrackingMode = "TrackAll";
      await ctx.sync();
      await fn(ctx);
      return "được";
    });
  } catch (e) {
    out.thu[ten] = `LOI ${e.message} ${(e.debugInfo && e.debugInfo.errorLocation) || ""}`;
  }
};
const get = (ctx, tag) => ctx.document.getContentControls({ types: TYPES }).getByTag(tag).getFirst();
await thu("ben_a insertText", async (ctx) => {
  get(ctx, "ben_a").insertText("Công ty Thử", "Replace");
  await ctx.sync();
});
await thu("ben_a getRange(Content).insertText", async (ctx) => {
  get(ctx, "ben_a").getRange("Content").insertText("Công ty Thử", "Replace");
  await ctx.sync();
});
await thu("loai_ts getRange(Content).insertText", async (ctx) => {
  get(ctx, "loai_ts").getRange("Content").insertText("Căn hộ chung cư", "Replace");
  await ctx.sync();
});
await thu("la_to_chuc isChecked", async (ctx) => {
  const c = get(ctx, "la_to_chuc");
  c.checkboxContentControl.isChecked = true;
  await ctx.sync();
  c.checkboxContentControl.load("isChecked");
  c.load("text");
  await ctx.sync();
  out.thu.la_to_chuc_sau = `${c.checkboxContentControl.isChecked} "${c.text}"`;
});
await Word.run(async (ctx) => {
  const t = ctx.document.body.getTrackedChanges();
  t.load("items");
  await ctx.sync();
  out.suaDoiTruocKhiDon = t.items.length;
  ctx.document.changeTrackingMode = "Off";
  await ctx.sync();
});
return out;
