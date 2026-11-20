// Đợt 11: chạy một bài trước / sau trong Excel hoặc PowerPoint trên hội thoại mới. Biến (đặt ở tệp ghép trước): CHE_DO ("paid" | "free"),
// NGUON, MODEL_RE, YEU_CAU. Trả thêm danh sách bước công cụ đã chạy (dòng ✅ / ❌ trên khung) để đối chiếu.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-lib.js <tệp biến> kiem-thu/cdp-bai/dot-11-chay.js
await toChat();
if (document.querySelector('select[aria-label="Nhà cung cấp"]')) {
  btn(/^Model đang dùng/).click();
  await sleep(500);
}
btn(/^Hội thoại mới/)?.click();
await sleep(800);
// Chế độ
const card = document.querySelector('[data-aidn="chon-che-do"]');
if (card) card.querySelector(`[data-mode="${CHE_DO}"]`).click();
else {
  btn(/^Model đang dùng/).click();
  await sleep(700);
  [...document.querySelectorAll('[data-aidn="che-do"] button')].find((b) => (CHE_DO === "free" ? /Miễn phí/ : /Trả phí/).test(b.innerText)).click();
  await sleep(500);
  if (document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
}
await sleep(1500);
const chon = await pick(NGUON, MODEL_RE);
const r = await send(YEU_CAU, { timeout: 400000, tail: 4000 });
const phienBan = (document.body.innerText.match(/\bv4\.\d+(\.\d+)?/) || [])[0] || null;
return {
  phienBan,
  cheDo: CHE_DO,
  model: chon,
  ...r,
  buoc: r.khung.split("\n").filter((l) => /^[✅❌]/.test(l)),
  doiNguon: r.khung.match(/↪[^\n]+/g) || [],
  lanGoiAI: (r.khung.match(/(\d+) lần gọi AI/) || [])[1] || null,
  token: (r.khung.match(/Token: [^\n]+/) || [])[0] || null,
};
