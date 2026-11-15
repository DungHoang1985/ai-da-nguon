// Bấm địa chỉ ô (Excel) hoặc "slide N" (PowerPoint) trong câu trả lời rồi xem ứng dụng có chuyển tới đúng chỗ không.
// Chạy trong khung của Excel hoặc PowerPoint: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-lib.js kiem-thu/cdp-bai/excel-ppt-bam-dia-chi.js
const isExcel = typeof Excel !== "undefined" && Office.context.host === Office.HostType.Excel;
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
if (isExcel) {
  const r = await send("Ghi số 5 vào ô D7 của sheet đang mở, rồi trả lời một câu có nhắc địa chỉ ô theo dạng Sheet!D7.", { tail: 500 });
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getActiveWorksheet().getRange("A1").select();
    await ctx.sync();
  });
  const link = [...document.querySelectorAll('[role="link"]')].find((x) => /D7/.test(x.innerText));
  if (!link) return { loi: "Không thấy liên kết tới ô D7", daHoi: r.daHoi, khung: r.khung };
  link.click();
  await sleep(1500);
  const sau = await Excel.run(async (ctx) => {
    const a = ctx.workbook.getSelectedRange();
    a.load("address,values");
    await ctx.sync();
    return { diaChi: a.address, giaTri: a.values[0][0] };
  });
  return { daHoi: r.daHoi, lienKet: link.innerText, vungChonSauKhiBam: sau, traLoi: r.khung.split("\n").slice(-7, -3).join(" / ") };
}
const r = await send("Không dùng công cụ ghi. Bài này có mấy slide? Trả lời một câu có nhắc “slide 1”.", { tail: 500 });
const link = [...document.querySelectorAll('[role="link"]')].find((x) => /slide\s*1\b/i.test(x.innerText));
if (!link) return { loi: "Không thấy liên kết slide 1", khung: r.khung };
link.click();
await sleep(1500);
const sau = await PowerPoint.run(async (ctx) => {
  const sel = ctx.presentation.getSelectedSlides();
  sel.load("items/id");
  const all = ctx.presentation.slides;
  all.load("items/id");
  await ctx.sync();
  return { slideDangChon: sel.items.map((s) => all.items.findIndex((x) => x.id === s.id) + 1), tongSlide: all.items.length };
});
return { lienKet: link.innerText, sauKhiBam: sau, traLoi: r.khung.split("\n").slice(-6, -3).join(" / ") };
