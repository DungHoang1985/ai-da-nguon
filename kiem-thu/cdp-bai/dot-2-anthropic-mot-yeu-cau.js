// Đợt 2, tiêu chí 4: gửi MỘT yêu cầu ngắn bằng Claude (key Anthropic đang lưu) để anh xem số dư tín dụng ở trang Thanh toán của Console giảm.
// Nhớ nguồn và model Anthropic đang chọn, dùng Claude Haiku 5.5 cho rẻ, xong trả lại đúng nguồn và model cũ. Không đọc kho lưu.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-anthropic-mot-yeu-cau.js
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
const openPicker = async () => {
  btn(/^Model đang dùng/).click();
  await sleep(800);
  return document.querySelector('select[aria-label="Nhà cung cấp"]');
};
const selected = () => {
  const o = document.querySelector('[role="listbox"][aria-label="Model"] [role="option"][aria-selected="true"]');
  return o ? o.getAttribute("title") : null;
};
let sel = await openPicker();
const nguonCu = sel.value;
if (sel.value !== "anthropic") {
  setNative(sel, "anthropic");
  await sleep(2000);
}
for (let i = 0; i < 20 && !document.querySelector('[role="option"]'); i++) await sleep(500);
const modelCu = selected();
document.body.click();
await sleep(300);
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const out = { nguonCu, modelAnthropicCu: modelCu };
out.chon = await pick("anthropic", /haiku-5-5/);
document.body.click();
await sleep(300);
window.__aidnTieuDeThem = [...(window.__aidnTieuDeThem || []), "^Không dùng công cụ\\. Trả lời đúng một câu: tín dụng"];
const q = "Không dùng công cụ. Trả lời đúng một câu: tín dụng API đã dùng được.";
const r = await send(q, { tail: 500 });
const all = document.body.innerText;
const ans = all.slice(all.lastIndexOf(q) + q.length);
out.giay = r.giay;
out.traLoi = ans.split("\n").filter((l) => l.trim() && !/^(Chèn vào chỗ con trỏ|Sao chép|Hỏi trước|Anthropic · )/.test(l.trim())).join(" | ").slice(0, 300);
out.coLoi = /hết tiền|Key |từ chối|Lỗi/.test(out.traLoi);
// Trả lại model Anthropic cũ rồi nguồn cũ
if (modelCu) {
  const back = await pick("anthropic", new RegExp("^" + esc(modelCu) + "$"));
  document.body.click();
  await sleep(300);
  out.traLaiModelAnthropic = back.daChon || back.loi;
}
if (nguonCu !== "anthropic") {
  sel = await openPicker();
  setNative(sel, nguonCu);
  await sleep(2000);
  document.body.click();
  await sleep(300);
}
out.nguonSauCung = label(btn(/^Model đang dùng/)).slice(0, 90);
return out;
