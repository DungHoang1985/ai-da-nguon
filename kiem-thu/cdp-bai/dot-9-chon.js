// Đợt 9: bấm chọn chế độ (biến CHE_DO = "free" | "paid") ở thẻ chọn chế độ hoặc ở bảng chọn model
const want = typeof CHE_DO === "string" ? CHE_DO : "paid";
await toChat();
const card = document.querySelector('[data-aidn="chon-che-do"]');
if (card) card.querySelector(`[data-mode="${want}"]`).click();
else {
  btn(/^Model đang dùng/).click();
  await sleep(700);
  [...document.querySelectorAll('[data-aidn="che-do"] button')].find((b) => (want === "free" ? /Miễn phí/ : /Trả phí/).test(b.innerText)).click();
  await sleep(500);
  if (document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
}
await sleep(2500);
return { daChon: want };
