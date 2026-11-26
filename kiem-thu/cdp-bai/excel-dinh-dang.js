// 4.13.2 (Chạy thật): lệnh định dạng Excel trả chữ hiển thị + cảnh báo ô số có thể hiện "####". Dựng bảng có cột số lớn ở cột
// rộng mặc định, giao việc định dạng số (KHÔNG nhắc độ rộng cột) rồi xem: AI có thấy cảnh báo và tự nới cột không, sổ tính sau cùng
// có còn ô bị "####" không (ước theo cùng hệ số với src/core/exceldisplay.ts, đã đo bằng tự giãn cột). Không đọc kho lưu của khung.
// Đổi chế độ / nguồn / model: window.__aidnEx = { cheDo: "paid" | "free", nguon: "anthropic", model: "<regex>" }.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js [tệp tham số] kiem-thu/cdp-bai/excel-dinh-dang.js
const cfg = window.__aidnEx || {};
delete window.__aidnEx;
await toChat();
const want = cfg.cheDo || "paid";
const card = document.querySelector('[data-aidn="chon-che-do"]');
if (card) card.querySelector(`[data-mode="${want}"]`).click();
else {
  if (!document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
  await sleep(700);
  const m = [...document.querySelectorAll('[data-aidn="che-do"] button')].find((b) => (want === "free" ? /Miễn phí/ : /Trả phí/).test(b.innerText));
  m?.click();
  await sleep(800);
  if (document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
}
await sleep(2000);
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const pk = await pick(cfg.nguon || "anthropic", new RegExp(cfg.model || "claude-sonnet-5-5$", "i"));
if (pk.loi) {
  document.body.click();
  return { cheDo: want, loi: pk.loi, coSan: pk.coSan };
}
document.body.click();
await sleep(300);

const SHEET = "AIDN_Thu_DinhDang";
await Excel.run(async (ctx) => {
  const old = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
  await ctx.sync();
  if (!old.isNullObject) old.delete();
  const ws = ctx.workbook.worksheets.add(SHEET);
  ws.getRange("A1:B6").values = [
    ["Hạng mục", "Giá trị (đồng)"],
    ["Tiền sử dụng đất", 1234567890.5],
    ["Phí thẩm định", 45000000],
    ["Lệ phí trước bạ", 98765432.25],
    ["Chi phí bồi thường", 23456789012],
    ["Tổng", "=SUM(B2:B5)"],
  ];
  ws.activate();
  await ctx.sync();
});

// Cùng hệ số với estimateWidth (src/core/exceldisplay.ts)
const est = (t, size = 11) => {
  let em = 0.49;
  for (const ch of String(t)) em += /[0-9]/.test(ch) ? 0.49 : /[.,\s]/.test(ch) ? 0.26 : /[/:]/.test(ch) ? 0.38 : ch === "%" ? 0.68 : 0.55;
  return em * size;
};
const docCot = () =>
  Excel.run(async (ctx) => {
    const r = ctx.workbook.worksheets.getItem(SHEET).getRange("B2:B6");
    r.load("text,values,numberFormat");
    r.format.load("columnWidth");
    r.format.font.load("size");
    await ctx.sync();
    const w = r.format.columnWidth;
    const size = typeof r.format.font.size === "number" ? r.format.font.size : 11;
    const hep = r.text.map((row, i) => (typeof r.values[i][0] === "number" && est(row[0], size) > w + 1 ? `B${i + 2}` : null)).filter(Boolean);
    return { chu: r.text.map((x) => x[0]), dinhDang: r.numberFormat[0][0], rongCot: w, oHep: hep };
  });
const truoc = await docCot();
const YEU_CAU = "Định dạng cột Giá trị (B2:B6) thành số có dấu phân cách hàng nghìn và 2 chữ số thập phân.";
const r = await send(YEU_CAU, { tail: 2500, timeout: 300000 });
const sau = await docCot();
const txt = document.body.innerText;
const buoc = (txt.match(/[✅❌⚠️] [^\n]{0,200}/g) || []).slice(-12);
const lanGoi = (txt.match(/(\d+) lần gọi AI/g) || []).slice(-1);
await Excel.run(async (ctx) => {
  ctx.workbook.worksheets.getItem(SHEET).delete();
  await ctx.sync();
});
return {
  cheDo: want,
  model: pk.daChon,
  giay: r.giay,
  daHoi: r.daHoi,
  buoc,
  lanGoi,
  truoc,
  sau,
  conONgan: sau.oHep.length > 0,
  traLoi: r.khung.slice(-1500),
};
