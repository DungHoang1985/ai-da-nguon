// Thăm dò Đợt 8, bước 3: ô chọn (DropDownList) không nhận insertText. Đọc danh sách mục từ OOXML của ô; chọn một mục bằng cách
// ghi lại OOXML của ô (đổi chữ trong w:sdtContent, bỏ w:showingPlcHdr) — có ghi Track Changes không, từ chối có trả lại không.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-o-chon-ooxml.js
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
let xml = "";
await step("A đọc OOXML của ô chọn", async (ctx) => {
  const r = cc(ctx, "thanh_toan").getRange("Whole").getOoxml();
  await ctx.sync();
  xml = r.value;
  const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(xml)?.[0] || "";
  const muc = [...sdt.matchAll(/<w:listItem\b[^>]*w:displayText="([^"]*)"/g)].map((m) => m[1]);
  return { doDai: xml.length, muc, coSdt: !!sdt };
});
await step("B chọn 'Chuyển khoản' bằng insertOoxml (có theo dõi)", async (ctx) => {
  ctx.document.changeTrackingMode = "TrackAll";
  const sdtNew = xml.replace(/<w:sdt>([\s\S]*?)<\/w:sdt>/, (all) =>
    all.replace(/<w:showingPlcHdr\/>/, "").replace(/(<w:sdtContent>[\s\S]*?<w:t[^>]*>)[^<]*(<\/w:t>)/, "$1Chuyển khoản$2")
  );
  cc(ctx, "thanh_toan").getRange("Whole").insertOoxml(sdtNew, "Replace");
  await ctx.sync();
  const all = ctx.document.getContentControls({ types: TYPES });
  all.load("items/tag,items/type,items/text");
  const t = ctx.document.body.getTrackedChanges();
  t.load("items/type,items/text");
  await ctx.sync();
  return { o: all.items.filter((c) => /thanh_toan/.test(c.tag)).map((c) => `${c.tag} [${c.type}] "${c.text}"`), soO: all.items.length, suaDoi: t.items.map((x) => `${x.type}: ${x.text}`) };
});
await step("C từ chối sửa đổi, đọc lại", async (ctx) => {
  const t = ctx.document.body.getTrackedChanges();
  t.load("items");
  await ctx.sync();
  const n = t.items.length;
  t.rejectAll();
  ctx.document.changeTrackingMode = "Off";
  await ctx.sync();
  const all = ctx.document.getContentControls({ types: TYPES });
  all.load("items/tag,items/type,items/text");
  await ctx.sync();
  return { tuChoi: n, o: all.items.map((c) => `${c.tag} [${c.type}] "${c.text}"`) };
});
return out;
