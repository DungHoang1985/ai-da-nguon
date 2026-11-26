// Chạy thật đường tệp đính kèm (4.14.x, B7 phần 1 — prepareAttachments): đính kèm một tệp chữ nhỏ bằng ô chọn tệp của khung
// (DataTransfer, không cần hộp thoại Windows), hỏi một điều chỉ có trong tệp, đọc câu trả lời và thông báo. Không đọc kho lưu của khung.
// Tham số: window.__aidnDk = { cheDo: "paid" | "free", nguon: "anthropic", model: "<regex>" }.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js [tệp tham số] kiem-thu/cdp-bai/dinh-kem-hoi.js
const cfg = window.__aidnDk || {};
delete window.__aidnDk;
await toChat();
const want = cfg.cheDo || "paid";
if (!document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
await sleep(700);
[...document.querySelectorAll('[data-aidn="che-do"] button')].find((b) => (want === "free" ? /Miễn phí/ : /Trả phí/).test(b.innerText))?.click();
await sleep(800);
if (document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
await sleep(1500);
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const pk = await pick(cfg.nguon || "anthropic", new RegExp(cfg.model || "claude-haiku-5-5$", "i"));
document.body.click();
if (pk.loi) return { loi: pk.loi, coSan: pk.coSan };
const input = [...document.querySelectorAll('input[type="file"]')][0];
if (!input) return { loi: "Không thấy ô chọn tệp" };
const dt = new DataTransfer();
dt.items.add(new File(["Biên bản họp ngày 20/11/2026 giữa Bên A và Bên B.\nBên A đồng ý giao hàng chậm nhất ngày 15/12/2026.\nPhạt chậm giao 0,05% mỗi ngày.\n"], "bien-ban-thu.txt", { type: "text/plain" }));
input.files = dt.files;
input.dispatchEvent(new Event("change", { bubbles: true }));
for (let i = 0; i < 20 && !document.body.innerText.includes("bien-ban-thu.txt"); i++) await sleep(300);
const coChip = document.body.innerText.includes("bien-ban-thu.txt");
const r = await send("Theo tệp đính kèm, Bên A phải giao hàng chậm nhất ngày nào? Trả lời một câu, không sửa tài liệu.", { tail: 1500, timeout: 240000 });
const txt = r.khung;
return {
  model: pk.daChon,
  coChipTruocKhiGui: coChip,
  giay: r.giay,
  daHoi: r.daHoi,
  traLoiDungNgay: /15\/12\/2026|15 tháng 12/.test(txt),
  tepTrongTinNhan: /📎 bien-ban-thu\.txt/.test(document.body.innerText),
  lanGoi: (document.body.innerText.match(/\d+ lần gọi AI/g) || []).slice(-1),
  cuoiKhung: txt.slice(-700),
};
