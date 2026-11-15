// Thăm dò Đợt 1B (bước 2a, phần 2): Document.compareFromBase64 (WordApiDesktop 1.2).
// Ghép trước: tệp định nghĩa B64 và tệp định nghĩa `const TARGET = "CompareTargetNew" | "CompareTargetCurrent"`.
// Chạy: node kiem-thu/cdp.mjs eval 9229 <b64.js> <target.js> kiem-thu/cdp-bai/tham-do-1b-3-so-sanh.js
const snap = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type,items/author,items/text");
    ctx.document.load("changeTrackingMode,saved");
    await ctx.sync();
    let h = 0;
    for (const p of ps.items) for (let i = 0; i < p.text.length; i++) h = (Math.imul(h, 31) + p.text.charCodeAt(i)) | 0;
    const byAuthor = {};
    tc.items.forEach((t) => {
      const a = t.author === "Đối tác" || t.author === "So sánh AIDN" ? t.author : "(tài khoản Office)";
      byAuthor[a] = (byAuthor[a] || 0) + 1;
    });
    return { soDoan: ps.items.length, bam: h, soSuaDoi: tc.items.length, theoTacGia: byAuthor, cheDo: ctx.document.changeTrackingMode, daLuu: ctx.document.saved, vai: tc.items.slice(0, 12).map((t) => `${t.type}: ${t.text.slice(0, 40)}`) };
  });
const out = { target: TARGET, truoc: await snap() };
const t0 = performance.now();
try {
  await Word.run(async (ctx) => {
    ctx.document.compareFromBase64(B64, {
      compareTarget: TARGET,
      authorName: "So sánh AIDN",
      addToRecentFiles: false,
      detectFormatChanges: false,
      ignoreAllComparisonWarnings: true,
    });
    await ctx.sync();
  });
  out.goi = "không lỗi";
} catch (e) {
  out.goi = { LOI: e.message, ma: e.code, o: e.debugInfo && e.debugInfo.errorLocation };
}
out.giay = Math.round(performance.now() - t0) / 1000;
await new Promise((r) => setTimeout(r, 1500));
try {
  out.sau = await snap();
} catch (e) {
  out.sau = { LOI: e.message };
}
return out;
