// Đợt 2: bấm "Đăng nhập lại" trong thẻ OpenRouter để mở hộp thoại đăng nhập của Office, trả về dòng trạng thái của thẻ.
// Khóa đang lưu KHÔNG bị đổi: chỉ khi đổi mã thành công thì khóa mới thay khóa cũ. Bài này dừng trước bước đăng nhập.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-mo-hop-thoai.js
// Sau đó: node kiem-thu/cdp.mjs list 9229   (xem hộp thoại đang ở trang nào, tham số gì)
if (!document.querySelector('[data-aidn-card="signin-openrouter"]')) {
  btn(/^Cài đặt/)?.click();
  await sleep(1200);
}
const card = document.querySelector('[data-aidn-card="signin-openrouter"]');
if (!card) return { loi: "Không thấy thẻ OpenRouter", khung: tail(500) };
card.scrollIntoView();
const inCard = (re) => [...card.querySelectorAll("button")].find((b) => re.test(b.innerText.trim()));
let go = inCard(/^Đăng nhập (OpenRouter|lại)$/);
if (!go) {
  inCard(/^Đăng nhập lại \(đổi khóa\)$/)?.click();
  await sleep(600);
  go = inCard(/^Đăng nhập (OpenRouter|lại)$/);
}
if (!go) return { loi: "Không thấy nút đăng nhập", the: card.innerText.replace(/\n+/g, " | ").slice(0, 500) };
for (let i = 0; i < 20 && go.disabled; i++) await sleep(300);
go.click();
await sleep(5000);
const status = [...card.querySelectorAll('[role="status"]')].map((el) => el.innerText).join(" | ");
return { nut: go.innerText.trim(), trangThaiThe: status, nutTrongThe: [...card.querySelectorAll("button")].map((b) => b.innerText.trim()) };
