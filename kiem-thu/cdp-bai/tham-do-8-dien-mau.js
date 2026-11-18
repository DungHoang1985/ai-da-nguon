// Thăm dò Đợt 8 (điền mẫu): Office.js điền được những loại ô điền mẫu nào, có ghi Track Changes không, Hoàn tác bằng từ chối sửa đổi
// có trả ô về như cũ (cả chữ gợi ý) không, và điền chỗ trống "……" trong đoạn thường. Chạy trên mẫu `kiem-thu/mau-thu.mjs`.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-dien-mau.js   (tệp mẫu đóng không lưu sau khi thử)
const out = { sets: {}, ca: {} };
for (const v of ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"]) out.sets[v] = Office.context.requirements.isSetSupported("WordApi", v);
const TYPES = ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"];
const step = async (ten, fn) => {
  const t0 = performance.now();
  try {
    out.ca[ten] = await Word.run(fn);
  } catch (e) {
    out.ca[ten] = { LOI: (e && e.message) || String(e), ma: e && e.code, o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
  if (out.ca[ten] && typeof out.ca[ten] === "object") out.ca[ten].ms = Math.round(performance.now() - t0);
};
const cc = (ctx, tag) => ctx.document.getContentControls({ types: TYPES }).getByTag(tag).getFirst();
const revCount = async (ctx) => {
  const t = ctx.document.body.getTrackedChanges();
  t.load("items");
  await ctx.sync();
  return t.items.length;
};
const doc = async (ctx, tag) => {
  const c = cc(ctx, tag);
  c.load("text,type,showingPlaceholderText");
  await ctx.sync();
  return { text: c.text, type: c.type, goiY: c.showingPlaceholderText };
};

await step("0 liệt kê ô", async (ctx) => {
  const all = ctx.document.getContentControls({ types: TYPES });
  all.load("items/tag,items/title,items/type,items/text,items/showingPlaceholderText,items/cannotEdit,items/cannotDelete,items/placeholderText");
  await ctx.sync();
  return { o: all.items.map((c) => `${c.tag} [${c.type}] "${c.text}" goiY=${c.showingPlaceholderText} khoaND=${c.cannotEdit} khoaXoa=${c.cannotDelete} ph="${c.placeholderText}"`) };
});
await step("1 bật Track Changes", async (ctx) => {
  ctx.document.load("changeTrackingMode");
  await ctx.sync();
  const cu = ctx.document.changeTrackingMode;
  ctx.document.changeTrackingMode = "TrackAll";
  await ctx.sync();
  return { cu, suaDoi: await revCount(ctx) };
});
await step("2 ô chữ thường (PlainText) insertText Replace", async (ctx) => {
  const c = cc(ctx, "ben_a");
  c.insertText("Công ty TNHH Thử Nghiệm Hưng Thịnh", "Replace");
  await ctx.sync();
  return { sau: await doc(ctx, "ben_a"), suaDoi: await revCount(ctx) };
});
await step("3 ô chữ định dạng (RichText) insertText Replace", async (ctx) => {
  const c = cc(ctx, "dia_chi_a");
  c.insertText("Số 1 phố Thử, phường Hoàn Kiếm, Hà Nội", "Replace");
  await ctx.sync();
  return { sau: await doc(ctx, "dia_chi_a"), suaDoi: await revCount(ctx) };
});
await step("4 ô đánh dấu (CheckBox) isChecked", async (ctx) => {
  const c = cc(ctx, "la_to_chuc");
  c.load("checkboxContentControl");
  await ctx.sync();
  c.checkboxContentControl.isChecked = true;
  await ctx.sync();
  c.checkboxContentControl.load("isChecked");
  await ctx.sync();
  return { isChecked: c.checkboxContentControl.isChecked, sau: await doc(ctx, "la_to_chuc"), suaDoi: await revCount(ctx) };
});
await step("5 ô chọn (DropDownList): danh sách + chọn mục", async (ctx) => {
  const c = cc(ctx, "thanh_toan");
  const list = c.dropDownListContentControl.getListItemsOrNullObject();
  list.load("items/displayText,items/value");
  await ctx.sync();
  const muc = list.isNullObject ? [] : list.items.map((x) => x.displayText);
  if (!list.isNullObject) {
    list.items.find((x) => x.displayText === "Chuyển khoản").select();
    await ctx.sync();
  }
  return { muc, sau: await doc(ctx, "thanh_toan"), suaDoi: await revCount(ctx) };
});
await step("5b ô chọn (DropDownList): insertText chữ ngoài danh sách", async (ctx) => {
  cc(ctx, "thanh_toan").insertText("Trả góp", "Replace");
  await ctx.sync();
  return { sau: await doc(ctx, "thanh_toan") };
});
await step("6 ô chọn hoặc nhập (ComboBox): insertText chữ tự do", async (ctx) => {
  cc(ctx, "loai_ts").insertText("Căn hộ chung cư", "Replace");
  await ctx.sync();
  return { sau: await doc(ctx, "loai_ts"), suaDoi: await revCount(ctx) };
});
await step("7 ô ngày (DatePicker): insertText", async (ctx) => {
  cc(ctx, "ngay_ky").insertText("18/11/2026", "Replace");
  await ctx.sync();
  return { sau: await doc(ctx, "ngay_ky"), suaDoi: await revCount(ctx) };
});
await step("8 ô khóa nội dung: insertText", async (ctx) => {
  cc(ctx, "so_hd").insertText("HĐĐC-SỬA", "Replace");
  await ctx.sync();
  return { sau: await doc(ctx, "so_hd") };
});
await step("9 chỗ trống '……' trong đoạn: search rồi thay", async (ctx) => {
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  await ctx.sync();
  const para = ps.items.find((x) => x.text.startsWith("Họ và tên người đại diện"));
  const hits = para.search("……………………………………", { matchCase: false });
  hits.load("items/text");
  await ctx.sync();
  if (hits.items.length) hits.items[0].insertText("Nguyễn Văn Thử", "Replace");
  await ctx.sync();
  para.load("text");
  await ctx.sync();
  return { timThay: hits.items.length, sau: para.text, suaDoi: await revCount(ctx) };
});
await step("9b '[●]' hai chỗ trong một đoạn: thay chỗ thứ hai", async (ctx) => {
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  await ctx.sync();
  const para = ps.items.find((x) => x.text.startsWith("Số tiền đặt cọc"));
  const hits = para.search("[●]");
  hits.load("items/text");
  await ctx.sync();
  if (hits.items.length > 1) hits.items[1].insertText("năm trăm triệu", "Replace");
  await ctx.sync();
  para.load("text");
  await ctx.sync();
  return { timThay: hits.items.length, sau: para.text };
});
await step("10 ô bảng trống: điền", async (ctx) => {
  const t = ctx.document.body.tables.getFirst();
  t.load("values");
  await ctx.sync();
  const c = t.getCell(1, 1);
  c.body.insertText("85,5", "Start");
  await ctx.sync();
  t.load("values");
  await ctx.sync();
  return { sau: t.values, suaDoi: await revCount(ctx) };
});
await step("11 Hoàn tác: từ chối mọi sửa đổi đã thêm, đọc lại các ô", async (ctx) => {
  const t = ctx.document.body.getTrackedChanges();
  t.load("items/type,items/text,items/author");
  await ctx.sync();
  const n = t.items.length;
  t.rejectAll();
  await ctx.sync();
  const all = ctx.document.getContentControls({ types: TYPES });
  all.load("items/tag,items/type,items/text,items/showingPlaceholderText");
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  await ctx.sync();
  const cb = cc(ctx, "la_to_chuc");
  cb.checkboxContentControl.load("isChecked");
  await ctx.sync().catch(() => undefined);
  return {
    daTuChoi: n,
    cacLoai: [...new Set(t.items.map((x) => x.type))],
    o: all.items.map((c) => `${c.tag} "${c.text}" goiY=${c.showingPlaceholderText}`),
    oDanhDauSauHoanTac: cb.checkboxContentControl.isChecked,
    doan: ps.items.map((x) => x.text).filter((x) => /Họ và tên|Số tiền/.test(x)),
  };
});
await step("12 trả chế độ theo dõi", async (ctx) => {
  ctx.document.changeTrackingMode = "Off";
  await ctx.sync();
  return { ok: true };
});
return out;
