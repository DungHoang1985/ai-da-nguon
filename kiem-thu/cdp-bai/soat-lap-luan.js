// 4.13.0 (Chạy thật): công cụ soát lập luận word_explain_edits. Dựng hợp đồng thử có Điều 5 (phạt 12%), giao việc sửa điều khoản
// phạt có một căn cứ luật bên ngoài (KHÔNG nêu tên công cụ trong yêu cầu) rồi xem: AI có tự gọi soát lập luận trước khi ghi không,
// người soát trả gì, câu trả lời có nói rõ căn cứ luật chưa kiểm được không, văn bản sau sửa. Không đọc kho lưu của khung.
// Đổi chế độ / nguồn / model: window.__aidnSl = { cheDo: "paid" | "free", nguon: "openrouter", model: "<regex>" }; model: "?" chỉ liệt kê.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/soat-lap-luan.js
const cfg = window.__aidnSl || {};
delete window.__aidnSl;
await toChat();
// Chọn chế độ (như dot-9-chon.js)
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
const pk = await pick(cfg.nguon || "openrouter", new RegExp(cfg.model === "?" ? "^$" : cfg.model || "claude-sonnet", "i"));
if (pk.loi) {
  document.body.click();
  return { cheDo: want, loi: pk.loi, coSan: (pk.coSan || []).filter((t) => /claude|nemotron|free/i.test(t)) };
}
document.body.click();
await sleep(300);

await Word.run(async (ctx) => {
  const b = ctx.document.body;
  b.clear();
  const ps = [
    "HỢP ĐỒNG MUA BÁN THIẾT BỊ (văn bản thử soát lập luận)",
    "Điều 1. Đối tượng hợp đồng",
    "Bên A bán cho Bên B 10 máy phát điện theo phụ lục kèm theo.",
    "Điều 5. Phạt vi phạm",
    "Bên vi phạm nghĩa vụ hợp đồng chịu phạt 12% giá trị phần nghĩa vụ hợp đồng bị vi phạm.",
    "Điều 6. Bồi thường thiệt hại",
    "Bên vi phạm bồi thường toàn bộ thiệt hại thực tế phát sinh cho bên bị vi phạm.",
  ];
  ps.forEach((t) => b.insertParagraph(t, "End"));
  await ctx.sync();
  const all = b.paragraphs;
  all.load("items/text");
  await ctx.sync();
  if (all.items[0].text === "") all.items[0].delete();
  await ctx.sync();
});
const vanBan = () =>
  Word.run(async (ctx) => {
    const b = ctx.document.body;
    const tc = b.getTrackedChanges();
    tc.load("items/type,items/text");
    const ps = b.paragraphs;
    ps.load("items");
    await ctx.sync();
    const cur = ps.items.map((p) => p.getText());
    await ctx.sync();
    return { doan: cur.map((x) => x.value), suaDoi: tc.items.map((t) => `${t.type}: ${t.text}`) };
  });
const truoc = await vanBan();
const YEU_CAU =
  "Sửa điều khoản phạt vi phạm (Điều 5) của hợp đồng này: giảm mức phạt xuống 8% và thêm một câu nói rõ tổng mức phạt không vượt quá 8% giá trị phần nghĩa vụ bị vi phạm theo giới hạn của pháp luật thương mại.";
const r = await send(YEU_CAU, { tail: 3000, timeout: 420000 });
const sau = await vanBan();
const txt = document.body.innerText;
const buoc = (txt.match(/[✅❌⚠️] [^\n]{0,180}/g) || []).slice(-14);
const soat = buoc.filter((l) => /Soát lập luận/.test(l));
const lanGoi = (txt.match(/(\d+) lần gọi AI/g) || []).slice(-2);
// Phần trả lời cuối của AI: sau bước cuối cùng
const traLoi = r.khung.slice(-1800);
await Word.run(async (ctx) => {
  ctx.document.body.clear();
  await ctx.sync();
});
return {
  cheDo: want,
  model: pk.daChon,
  giay: r.giay,
  daHoi: r.daHoi,
  goiSoat: soat.length > 0,
  soat,
  buoc,
  lanGoi,
  nhacCanDoiChieu: /đối chiếu văn bản gốc|chưa kiểm được|không kiểm được|chưa được xác nhận/i.test(traLoi),
  truoc: truoc.doan,
  sau: sau.doan,
  suaDoi: sau.suaDoi,
  traLoi,
};
