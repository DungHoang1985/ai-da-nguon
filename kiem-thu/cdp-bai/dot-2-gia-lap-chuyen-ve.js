// Đợt 2: chạy TRONG hộp thoại đăng nhập đang ở trang của nhà cung cấp. Giả lập bước nhà cung cấp chuyển về trang của add-in
// kèm một MÃ GIẢ (không đăng nhập gì cả), để kiểm đường về: trang dang-nhap.html gửi mã về khung → khung đổi mã → máy chủ từ chối mã giả.
// state lấy từ chính địa chỉ của hộp thoại và không in ra. Đặt window.__aidnSaiState = true trước khi chạy để thử ca state không khớp.
// Chạy: node kiem-thu/cdp.mjs eval "9229#openrouter.ai" kiem-thu/cdp-bai/dot-2-gia-lap-chuyen-ve.js
const here = new URL(location.href);
const back = here.searchParams.get("redirect_url");
const src = back ? new URL(back, here.origin) : here;
const cb = src.searchParams.get("callback_url");
const st = src.searchParams.get("state");
if (!cb || !st) return { loi: "Hộp thoại không mang callback_url / state", duongDan: here.pathname };
const sai = !!window.__aidnSaiState;
const target = `${cb}?code=ma-gia-0123456789abcdef&state=${encodeURIComponent(sai ? "state-sai-0123456789abc" : st)}`;
setTimeout(() => {
  location.href = target;
}, 300);
return { chuyenVe: cb, maGia: true, stateDung: !sai };
