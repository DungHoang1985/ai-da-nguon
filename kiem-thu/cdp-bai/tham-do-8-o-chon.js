// Thăm dò Đợt 8, bước 2: điền ô chọn (DropDownList), ô chọn hoặc nhập (ComboBox), ô ngày (DatePicker) — insertText trên ô báo
// GeneralException khi bật Track Changes (tham-do-8-dien-mau.js). Thử: tắt theo dõi; ghi vào vùng nội dung của ô; đọc danh sách mục
// qua API khác; trường hợp trọn dãy chấm "……" tìm đúng cả dãy. Xong thì trả mọi thứ (từ chối sửa đổi / gõ lại chữ cũ).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-o-chon.js
const out = { ca: {} };
const TYPES = ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"];
const step = async (ten, fn) => {
  try {
    out.ca[ten] = await Word.run(fn);
  } catch (e) {
    out.ca[ten] = { LOI: (e && e.message) || String(e), ma: e && e.code, o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
};
const cc = (ctx, tag) => ctx.document.getContentControls({ types: TYPES }).getByTag(tag).getFirst();
const textOf = async (ctx, tag) => {
  const c = cc(ctx, tag);
  c.load("text");
  await ctx.sync();
  return c.text;
};
const mode = (m) =>
  step("chế độ " + m, async (ctx) => {
    ctx.document.changeTrackingMode = m;
    await ctx.sync();
    return { ok: true };
  });

await step("A thuộc tính con của ô chọn", async (ctx) => {
  const c = cc(ctx, "thanh_toan");
  c.load();
  await ctx.sync();
  const j = c.toJSON();
  return { khoa: Object.keys(j), dd: j.dropDownListContentControl, cb: j.comboBoxContentControl, dp: j.datePickerContentControl };
});
await mode("TrackAll");
for (const [tag, val] of [["thanh_toan", "Chuyển khoản"], ["loai_ts", "Căn hộ chung cư"], ["ngay_ky", "18/11/2026"]]) {
  await step(`B ${tag}: getRange("Content").insertText, có theo dõi`, async (ctx) => {
    cc(ctx, tag).getRange("Content").insertText(val, "Replace");
    await ctx.sync();
    return { sau: await textOf(ctx, tag) };
  });
  await step(`C ${tag}: getRange("Whole") search chữ gợi ý rồi thay, có theo dõi`, async (ctx) => {
    const r = cc(ctx, tag).getRange("Content");
    r.load("text");
    await ctx.sync();
    return { chuTrongVung: r.text };
  });
}
await mode("Off");
for (const [tag, val] of [["thanh_toan", "Chuyển khoản"], ["loai_ts", "Căn hộ chung cư"], ["ngay_ky", "18/11/2026"]]) {
  await step(`D ${tag}: insertText Replace, TẮT theo dõi`, async (ctx) => {
    cc(ctx, tag).insertText(val, "Replace");
    await ctx.sync();
    return { sau: await textOf(ctx, tag) };
  });
}
await step("E dãy chấm: tìm đúng trọn dãy từ chữ của đoạn", async (ctx) => {
  ctx.document.changeTrackingMode = "TrackAll";
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  await ctx.sync();
  const para = ps.items.find((x) => x.text.startsWith("Họ và tên người đại diện"));
  const m = /(?:\.{3,}|…+)+/u.exec(para.text);
  const hits = para.search(m[0]);
  hits.load("items/text");
  await ctx.sync();
  hits.items[0].insertText("Nguyễn Văn Thử", "Replace");
  await ctx.sync();
  para.load("text");
  await ctx.sync();
  return { day: m[0].length, timThay: hits.items.length, sau: para.text };
});
await step("F trả lại: từ chối sửa đổi có theo dõi, tắt theo dõi", async (ctx) => {
  const t = ctx.document.body.getTrackedChanges();
  t.load("items");
  await ctx.sync();
  const n = t.items.length;
  t.rejectAll();
  ctx.document.changeTrackingMode = "Off";
  await ctx.sync();
  const all = ctx.document.getContentControls({ types: TYPES });
  all.load("items/tag,items/text");
  await ctx.sync();
  return { tuChoi: n, o: all.items.map((c) => `${c.tag} "${c.text}"`) };
});
return out;
