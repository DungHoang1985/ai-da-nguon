// Chọn nguồn và model cho các bài Chạy thật, trả về nguồn và model ĐANG chọn trước đó (để trả lại khi dọn).
// Tham số ở tệp ghép phía trước: window.__aidnChon = { nguon: "anthropic", model: "haiku-5-5" }  (model: biểu thức khớp tên / mã model;
// thêm chinhXac: true để khớp đúng cả chuỗi — dùng khi trả lại model cũ).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-7-chon-model.js
const c = window.__aidnChon || {};
delete window.__aidnChon;
if (!c.nguon || !c.model) return { loi: "Thiếu window.__aidnChon = { nguon, model }" };
await toChat();
// Nút "Model đang dùng" là nút bật / tắt: chỉ bấm khi bảng chưa mở
if (!document.querySelector('select[aria-label="Nhà cung cấp"]')) {
  btn(/^Model đang dùng/).click();
  await sleep(900);
}
const s = document.querySelector('select[aria-label="Nhà cung cấp"]');
if (!s) return { loi: "Không mở được bảng chọn model", khung: tail(400) };
const nguonCu = s.value;
const optCu = document.querySelector('[role="listbox"][aria-label="Model"] [role="option"][aria-selected="true"]');
const modelCu = optCu ? optCu.getAttribute("title") : null;
if (s.value !== c.nguon) {
  setNative(s, c.nguon);
  await sleep(1800);
}
for (let i = 0; i < 20 && !document.querySelector('[role="option"]'); i++) await sleep(500);
const opts = [...document.querySelectorAll('[role="listbox"][aria-label="Model"] [role="option"]')];
const esc = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const re = c.chinhXac ? new RegExp("^" + esc(c.model) + "$") : new RegExp(c.model, "i");
const hit = opts.find((o) => re.test(o.getAttribute("title") || ""));
if (!hit) return { loi: "Không thấy model khớp", coSan: opts.map((o) => o.getAttribute("title")).slice(0, 40), nguonCu, modelCu };
hit.click();
await sleep(700);
const nhan = label(btn(/^Model đang dùng/)).slice(0, 140);
await toChat();
return { nguonCu, modelCu, daChon: nhan };
