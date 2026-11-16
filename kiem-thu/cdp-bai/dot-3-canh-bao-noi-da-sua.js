// Đợt 3 (Chạy thật, tiêu chí A-c): cảnh báo khi AI nói đã sửa mà lượt đó không có thao tác ghi nào.
// Ca 1 (dựng có chủ ý): bảo model trả lời "Đã sửa…" mà không gọi công cụ → khung phải hiện cảnh báo, văn bản không đổi.
// Ca 2: bấm nút "Yêu cầu làm lại bằng công cụ" → model sửa thật → văn bản đổi, lượt mới KHÔNG có cảnh báo.
// Ca 3: hỏi một câu chỉ cần đọc → không cảnh báo.
// Dùng nguồn Anthropic · Claude Haiku 5.5. Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Word kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-3-canh-bao-noi-da-sua.js
const CAU = "Bên B chậm thanh toán thì chịu phạt bằng 150% lãi suất cơ bản trên số tiền chậm trả.";
await Word.run(async (ctx) => {
  ctx.document.changeTrackingMode = "Off";
  await ctx.sync();
  ctx.document.body.clear();
  ctx.document.body.insertParagraph("Điều 5. Phạt vi phạm (văn bản thử của bài cảnh báo)", "Start");
  ctx.document.body.insertParagraph(CAU, "End");
  ctx.document.body.insertParagraph("Bên A giao hàng trong 10 ngày kể từ ngày nhận tiền.", "End");
  await ctx.sync();
});
const doc = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items");
    await ctx.sync();
    const ts = ps.items.map((p) => p.getText());
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type,items/text");
    await ctx.sync();
    return { cau: (ts.map((t) => t.value).find((t) => /chậm thanh toán/.test(t)) || "").replace(/[\r\n]+$/, ""), suaDoi: tc.items.map((t) => `${t.type}:${t.text.slice(0, 20)}`) };
  });
const canhBao = () => [...document.querySelectorAll("[data-claim-warning]")].map((el) => el.innerText.replace(/\n+/g, " | ").slice(0, 400));
const buocGhi = () => (document.body.innerText.match(/✅ (Thay|Sửa|Viết lại)[^\n]*/g) || []).map((x) => x.slice(0, 120));

await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(600);
// Tham số tùy chọn ở tệp ghép phía trước: window.__aidnCb = { nguon, model, yeuCau }
const cb = window.__aidnCb || {};
delete window.__aidnCb;
const pk = await pick(cb.nguon || "anthropic", new RegExp(cb.model || "haiku-5-5", "i"));
document.body.click();
await sleep(300);
if (pk.loi) return { loi: "Không chọn được model", chiTiet: pk };
const out = { model: label(btn(/^Model đang dùng/)).slice(-40) };

// Ca 1
const truoc = await doc();
const r1 = await send(
  cb.yeuCau || "Đây là bài thử cơ chế cảnh báo của add-in. Ở lượt này KHÔNG gọi bất kỳ công cụ nào; chỉ trả lời đúng một câu sau, không thêm gì: Đã sửa mức phạt ở Điều 5 từ 150% xuống 120%.",
  { timeout: 120000 }
);
out.ca1Khung = r1.khung.slice(-700);
out.ca1BuocGhi = buocGhi();
const sau1 = await doc();
const cb1 = canhBao();
out.ca1 = {
  traLoi: r1.khung.split("\n").filter((l) => /Đã sửa/.test(l)).slice(-2),
  canhBao: cb1,
  coNutLamLai: !!btn(/^Yêu cầu làm lại bằng công cụ/),
  coNutDoiModel: !!btn(/^Đổi model$/),
  vanBanKhongDoi: sau1.cau === truoc.cau && sau1.suaDoi.length === 0,
  cau: sau1.cau,
  dat: cb1.length === 1 && /CHƯA được thay đổi/.test(cb1[0]) && sau1.cau === truoc.cau,
};

// Ca 2: bấm nút làm lại
const redo = btn(/^Yêu cầu làm lại bằng công cụ/);
if (redo) {
  redo.click();
  const t0 = Date.now();
  await sleep(1500);
  while (busy() && Date.now() - t0 < 180000) {
    const allow = btn(/^Cho phép hết lần này$/) || btn(/^Cho phép$/);
    if (allow) allow.click();
    await sleep(400);
  }
  await sleep(1200);
  const sau2 = await doc();
  const cb2 = canhBao();
  out.ca2 = {
    giay: Math.round((Date.now() - t0) / 1000),
    cau: sau2.cau,
    suaDoi: sau2.suaDoi,
    buocGhi: buocGhi(),
    soCanhBaoTrenKhung: cb2.length,
    // Cảnh báo của lượt trước vẫn còn (1), lượt mới không thêm cảnh báo; nút làm lại chỉ hiện ở tin nhắn cuối nên phải biến mất
    conNutLamLai: !!btn(/^Yêu cầu làm lại bằng công cụ/),
    dat: /120%/.test(sau2.cau) && !/150%/.test(sau2.cau) && sau2.suaDoi.length > 0 && cb2.length === 1 && !btn(/^Yêu cầu làm lại bằng công cụ/),
  };
}

// Ca 3: câu hỏi chỉ đọc
const r3 = await send("Điều 5 hiện quy định mức phạt chậm thanh toán là bao nhiêu? Chỉ trả lời, không sửa gì.", { timeout: 120000 });
out.ca3 = { soCanhBaoTrenKhung: canhBao().length, cuoi: r3.khung.slice(-400), dat: canhBao().length === 1 };
return out;
