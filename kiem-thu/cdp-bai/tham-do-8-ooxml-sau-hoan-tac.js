// Thăm dò Đợt 8: OOXML của ô điền mẫu sau một vòng "ghi cả ô bằng insertOoxml (có theo dõi) → từ chối sửa đổi". Dựng một ô chữ ở cuối
// văn bản, ghi đè cả ô, từ chối, rồi in phần w:sdt của OOXML trước và sau (rút gọn rPr). Xong thì xóa đoạn thử.
// Dọn trước khi chạy lại: xóa các đoạn "Đoạn thăm dò OOXML" còn sót (lần chạy hỏng).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-8-ooxml-sau-hoan-tac.js
const TAG = "aidn_tham_do_ooxml";
const out = {};
const sdtOf = (xml) => {
  const s = xml.indexOf("<w:sdt>") >= 0 ? xml.indexOf("<w:sdt>") : xml.indexOf("<w:sdt ");
  const e = xml.lastIndexOf("</w:sdt>");
  return s < 0 ? "(không có w:sdt) " + xml.slice(xml.indexOf("<w:body>"), xml.indexOf("<w:body>") + 600) : xml.slice(s, e + 8).replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "");
};
const cc = (ctx) => ctx.document.getContentControls({ types: ["RichText", "PlainText"] }).getByTag(TAG).getFirst();
try {
await Word.run(async (ctx) => {
  ctx.document.changeTrackingMode = "Off";
  const p = ctx.document.body.insertParagraph("Đoạn thăm dò OOXML: X", "End");
  out.buoc = 1;
  await ctx.sync();
  const x = p.search("X", { matchCase: true });
  x.load("items");
  out.buoc = 2;
  await ctx.sync();
  const c = x.items[0].insertContentControl();
  c.tag = TAG;
  c.title = "Thăm dò";
  c.load("id");
  out.buoc = 3;
  await ctx.sync();
  const myId = c.id;
  const o = c.getRange("Whole").getOoxml();
  out.buoc = 4;
  await ctx.sync();
  out.truoc = sdtOf(o.value);
  // ghi cả ô
  ctx.document.changeTrackingMode = "TrackAll";
  const nx = o.value.replace(/<w:showingPlcHdr\/>/g, "").replace(/(<w:sdtContent>)[\s\S]*?(<\/w:sdtContent>)/, '$1<w:r><w:t xml:space="preserve">Giá trị</w:t></w:r>$2');
  c.getRange("Whole").insertOoxml(nx, "Replace");
  out.buoc = 5;
  await ctx.sync();
  const all = ctx.document.getContentControls({ types: ["RichText", "PlainText"] }).getByTag(TAG);
  all.load("items/text");
  out.buoc = 6;
  await ctx.sync();
  out.sauGhi = all.items.map((y) => y.text);
  const t = ctx.document.body.getTrackedChanges();
  t.load("items");
  out.buoc = 7;
  await ctx.sync();
  t.rejectAll();
  ctx.document.changeTrackingMode = "Off";
  out.buoc = 8;
  await ctx.sync();
  const all2 = ctx.document.getContentControls({ types: ["RichText", "PlainText"] }).getByTag(TAG);
  all2.load("items/text,items/id");
  out.buoc = 9;
  await ctx.sync();
  out.sauHoanTac = all2.items.map((y) => `${y.id} "${y.text}"`);
  const mine = all2.items.find((y) => y.id === myId) || all2.items[0];
  const o2 = mine.getRange("Whole").getOoxml();
  const o3 = mine.getRange("Content").getOoxml();
  out.buoc = 10;
  await ctx.sync();
  out.ooxmlWhole = sdtOf(o2.value);
  out.ooxmlContent = sdtOf(o3.value);
  out.coSdtTrongWhole = /<w:sdt[\s>]/.test(o2.value);
  // dọn
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  out.buoc = 11;
  await ctx.sync();
  ps.items.filter((q) => q.text.startsWith("Đoạn thăm dò OOXML")).forEach((q) => q.delete());
  out.buoc = 12;
  await ctx.sync();
});
} catch (e) {
  out.LOI = e.message + " " + ((e.debugInfo && e.debugInfo.errorLocation) || "");
}
return out;
