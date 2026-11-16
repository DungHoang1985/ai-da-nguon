// Đợt 2: đặt sẵn các biến cho bài dọn dẹp chung (ghép TRƯỚC don-dep-sau-khi-thu.js) và trả model OpenRouter về model anh dùng trước khi thử.
// Nguồn trả về: window.__aidnNguonVe (mặc định "anthropic" — nguồn anh đang chọn lúc bắt đầu buổi ngồi máy 16/11/2026).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-don-dep.js kiem-thu/cdp-bai/don-dep-sau-khi-thu.js
window.__aidnTieuDeThem = ["^Ở câu về chậm hoàn trả", "^Không dùng công cụ\\. Trả lời đúng một câu: tín dụng", "^Không dùng công cụ\\. Ghi nhớ mã thử của phiên này"];
window.__aidnNguonCu = window.__aidnNguonVe || "anthropic";
await toChat();
const modelOrVe = window.__aidnModelOrVe || "baidu/ernie-5.5-flash-preview:free";
const pk = await pick("openrouter", new RegExp("^" + modelOrVe.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "$"));
document.body.click();
await sleep(400);
window.__aidnTraModelOr = pk.daChon || pk.loi;
