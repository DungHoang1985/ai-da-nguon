// Đợt 1A: tiêu đề các hội thoại do bài Chạy thật tạo ra, và nguồn phải trả lại. Ghép trước don-dep-sau-khi-thu.js:
//   node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1a-cau-hoi.js kiem-thu/cdp-bai/word-1a-don-dep.js kiem-thu/cdp-bai/don-dep-sau-khi-thu.js
const escRe = (s) => "^" + s.slice(0, 40).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
window.__aidnTieuDeThem = [
  ...(window.__aidnCauHoi || []).map((c) => escRe(c[0])),
  "^Tóm tắt toàn bộ các sửa đổi được theo dõi",
  "^Viết lại Điều 4\\.2",
  "^Thêm một khoản mới ngay sau Điều 8\\.2",
];
window.__aidnNguonCu = "openrouter";
