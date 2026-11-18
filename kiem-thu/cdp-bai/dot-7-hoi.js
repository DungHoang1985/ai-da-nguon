// Bài Chạy thật chung: gửi một yêu cầu (với nguồn và model đang chọn), chờ xong, trả các bước công cụ, số lần gọi AI và phần cuối khung.
// Tham số ở tệp ghép phía trước: window.__aidnBai = { yeuCau, giuHoiThoai?: true (không bấm "Hội thoại mới"), traLoi?: "allow" | "all" | "deny", tail?: số ký tự }
// Tệp đính kèm: đặt trước bằng `node kiem-thu/cdp.mjs file <cổng> "input[type=file]" <tệp>` rồi chạy bài này với giuHoiThoai: true.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-7-hoi.js
const bai = window.__aidnBai || {};
delete window.__aidnBai;
if (!bai.yeuCau) return { loi: "Thiếu window.__aidnBai.yeuCau" };
await toChat();
if (!bai.giuHoiThoai) {
  btn(/^Hội thoại mới/)?.click();
  await sleep(500);
}
// Chờ tệp đính kèm đọc xong (nếu có)
for (let i = 0; i < 40 && /Đang đọc tệp/.test(document.body.innerText); i++) await sleep(500);
const dinhKem = [...document.querySelectorAll("span")].map((e) => e.getAttribute("title") && /^[🖼📑📄]/.test(e.innerText) ? e.innerText.replace(/\s+/g, " ").trim().slice(0, 80) : "").filter(Boolean);
const r = await send(bai.yeuCau, { approve: () => bai.traLoi || "allow", timeout: bai.timeout || 240000, tail: bai.tail || 2400 });
const khung = (r.khung || "").replace(/\n{2,}/g, "\n");
const calls = [...document.querySelectorAll("[data-calls]")];
return {
  phienBan: (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0],
  model: label(btn(/^Model đang dùng/) || document.body).slice(-60),
  cheDoThu: !!document.querySelector('[data-test-mode="1"]'),
  dinhKem,
  giay: r.giay,
  daHoi: r.daHoi,
  lanGoiAI: calls.length ? Number(calls[calls.length - 1].getAttribute("data-calls")) : 1,
  buoc: (khung.match(/^[✅❌] ?.*$/gm) || []).map((l) => l.slice(0, 200)),
  theLamTiep: document.querySelector('[data-paused="1"]')?.innerText || null,
  khung,
};
