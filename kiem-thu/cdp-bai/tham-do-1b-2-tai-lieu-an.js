// Thăm dò Đợt 1B (bước 2a, phần 1): đọc một tệp .docx KHÁC mà không mở nó ra — Application.createDocument(base64) trả tài liệu ẩn.
// Ghép trước tệp này một tệp định nghĩa `const B64 = "<base64 của tệp .docx>"`.
// Chạy: node kiem-thu/cdp.mjs eval 9229 <b64.js> kiem-thu/cdp-bai/tham-do-1b-2-tai-lieu-an.js
const out = {};
const t0 = performance.now();
try {
  out.taiLieuAn = await Word.run(async (ctx) => {
    const doc = ctx.application.createDocument(B64);
    const ps = doc.body.paragraphs;
    ps.load("items/text,items/style,items/outlineLevel,items/tableNestingLevel,items/listItemOrNullObject/listString,items/listItemOrNullObject/level,items/listOrNullObject/id");
    await ctx.sync();
    const rows = ps.items.map((p, i) => `[${i}] ${p.listItemOrNullObject.isNullObject ? "" : p.listItemOrNullObject.listString + " "}${p.text.slice(0, 50)}`);
    return { soDoan: ps.items.length, dau: rows.slice(7, 20), cuoi: rows.slice(-6) };
  });
} catch (e) {
  out.taiLieuAn = { LOI: e.message, ma: e.code, o: e.debugInfo && e.debugInfo.errorLocation };
}
out.giay = Math.round(performance.now() - t0) / 1000;
// Tài liệu đang mở có bị đụng không
out.dangMo = await Word.run(async (ctx) => {
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  const tc = ctx.document.body.getTrackedChanges();
  tc.load("items/type");
  await ctx.sync();
  return { soDoan: ps.items.length, soSuaDoi: tc.items.length };
});
return out;
