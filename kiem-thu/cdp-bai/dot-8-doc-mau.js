// Đợt 8: đọc lại mẫu sau khi AI điền (bằng Office.js thô, không qua công cụ của add-in): mọi ô điền mẫu, các đoạn có chỗ trống
// theo bản hiện tại (không gồm chữ đã xóa chờ duyệt), bảng, số sửa đổi chờ duyệt.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/dot-8-doc-mau.js
return await Word.run(async (ctx) => {
  const cc = ctx.document.getContentControls({ types: ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"] });
  cc.load("items/tag,items/type,items/text");
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  const tc = ctx.document.body.getTrackedChanges();
  tc.load("items/type");
  const t = ctx.document.body.tables.getFirst();
  t.load("values");
  await ctx.sync();
  const cur = ps.items.map((p) => p.getText());
  await ctx.sync();
  const want = /^(Họ và tên|Số CCCD|Số tiền|Thời hạn|Hà Nội)/;
  return {
    o: cc.items.map((c) => `${c.tag} [${c.type}] "${c.text}"`),
    doan: cur.map((r) => r.value).filter((x) => want.test(x)),
    bang: t.values,
    suaDoi: tc.items.length,
  };
});
