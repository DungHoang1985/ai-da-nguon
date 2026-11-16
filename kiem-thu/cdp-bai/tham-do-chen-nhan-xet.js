// Thăm dò (16/11/2026): vì sao Range.insertComment có lúc bị AccessDenied. Thử chèn một nhận xét vào đoạn đầu, ghi lỗi đầy đủ,
// rồi xóa nhận xét thử nếu chèn được. Chỉ chạy trên tệp thử. Tên tác giả không được in ra.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/tham-do-chen-nhan-xet.js
const out = {};
const tryIt = async (name, fn) => {
  try {
    out[name] = await Word.run(fn);
  } catch (e) {
    out[name] = { loi: e.code || e.name, thongDiep: String(e.message || "").slice(0, 160), viTri: e.debugInfo?.errorLocation, cau: e.debugInfo?.statement, trong: e.debugInfo?.innerError?.code };
  }
};
await tryIt("trangThai", async (ctx) => {
  const d = ctx.document;
  d.load("changeTrackingMode,saved");
  const cs = d.body.getComments();
  cs.load("items/id");
  const sel = d.getSelection();
  sel.load("isEmpty");
  await ctx.sync();
  let prot = "không đọc được";
  try {
    const p = d.properties;
    p.load("security");
    await ctx.sync();
    prot = p.security;
  } catch (e) {
    prot = "lỗi " + (e.code || e.name);
  }
  return { cheDoTheoDoi: d.changeTrackingMode, daLuu: d.saved, soNhanXet: cs.items.length, vungChonRong: sel.isEmpty, baoMat: prot };
});
await tryIt("chenVaoDoanDau", async (ctx) => {
  const p = ctx.document.body.paragraphs.getFirst();
  const c = p.getRange("Whole").insertComment("AIDN thăm dò chèn nhận xét");
  c.load("id");
  await ctx.sync();
  const id = c.id;
  c.delete();
  await ctx.sync();
  return { chenDuoc: true, daXoa: true, coId: !!id };
});
await tryIt("chenVaoVungChon", async (ctx) => {
  const c = ctx.document.getSelection().insertComment("AIDN thăm dò chèn nhận xét 2");
  c.load("id");
  await ctx.sync();
  c.delete();
  await ctx.sync();
  return { chenDuoc: true, daXoa: true };
});
await tryIt("traLoiNhanXetCoSan", async (ctx) => {
  const cs = ctx.document.body.getComments();
  cs.load("items/id");
  await ctx.sync();
  if (!cs.items.length) return { khongCoNhanXet: true };
  const r = cs.items[0].reply("AIDN thăm dò trả lời");
  r.load("id");
  await ctx.sync();
  r.delete();
  await ctx.sync();
  return { traLoiDuoc: true, daXoa: true };
});
return out;
