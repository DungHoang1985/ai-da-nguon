// Đọc lại văn bản Word đang mở: các đoạn, bảng, số sửa đổi được theo dõi, số nhận xét, chế độ theo dõi.
return await Word.run(async (ctx) => {
  const body = ctx.document.body;
  const ps = body.paragraphs;
  ps.load("items/text");
  const tables = body.tables;
  tables.load("items/values");
  const tc = body.getTrackedChanges();
  tc.load("items/type,items/text");
  const cs = body.getComments();
  cs.load("items/content,items/resolved");
  ctx.document.load("changeTrackingMode");
  await ctx.sync();
  return {
    doan: ps.items.map((p) => p.text).filter((t) => t.trim()),
    bang: tables.items.map((t) => t.values),
    suaDoiTheoDoi: tc.items.map((x) => `${x.type}: ${String(x.text).slice(0, 60)}`),
    nhanXet: cs.items.map((c) => `${c.resolved ? "[đã xử lý] " : ""}${String(c.content).slice(0, 80)}`),
    cheDoTheoDoi: ctx.document.changeTrackingMode,
  };
});
