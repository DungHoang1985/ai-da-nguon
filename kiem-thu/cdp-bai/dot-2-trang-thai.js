// Đợt 2: đọc trạng thái khung và thẻ OpenRouter trong Cài đặt. KHÔNG đọc kho lưu của khung; chỉ đọc chữ đang hiện.
// Số tiền và số lượt trong thẻ hạn mức được che (thay chữ số bằng #) trước khi trả về: tệp bằng chứng không chứa số dư.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-trang-thai.js
const che = (s) => s.replace(/\d+(?:[.,]\d+)*/g, "#");
const version = (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0];
const model = btn(/^Model đang dùng/) ? label(btn(/^Model đang dùng/)).slice(0, 90) : "(màn chào hoặc đang ở Cài đặt)";
if (!document.querySelector('[data-aidn-card="signin-openrouter"]')) {
  const gear = btn(/^Cài đặt/);
  if (gear) {
    gear.click();
    await sleep(1200);
  }
}
const card = document.querySelector('[data-aidn-card="signin-openrouter"]');
if (!card) return { loi: "Không thấy thẻ OpenRouter", khung: tail(500) };
card.scrollIntoView();
// Chờ thẻ hạn mức tải xong (tối đa 15 giây)
for (let i = 0; i < 30 && /Bấm “Làm mới” để xem hạn mức|^$/.test((card.querySelector('[data-aidn="han-muc"]') || {}).innerText || ""); i++) await sleep(500);
const hanMuc = card.querySelector('[data-aidn="han-muc"]');
const res = {
  phienBan: version,
  model,
  theOpenRouter: che(card.innerText.replace(/\n+/g, " | ")).slice(0, 1200),
  coTheHanMuc: !!hanMuc,
  soDongHanMuc: hanMuc ? hanMuc.innerText.split("\n").filter((x) => x.trim()).length : 0,
  nutTrongThe: [...card.querySelectorAll("button")].map((b) => b.innerText.trim()),
  hopThoaiHoTro: typeof Office !== "undefined" && typeof Office.context?.ui?.displayDialogAsync === "function",
  moTrinhDuyetHoTro: typeof Office !== "undefined" && typeof Office.context?.ui?.openBrowserWindow === "function",
};
await toChat();
return res;
