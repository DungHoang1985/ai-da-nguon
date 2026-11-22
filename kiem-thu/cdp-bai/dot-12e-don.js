// Dọn sau bài thử dot-12e-bo-nho.js chạy với window.__e2GiuLai: xóa mọi mục "E2-thu" qua giao diện Cài đặt → Bộ nhớ.
if (!document.querySelector("[data-memory-card]")) {
  btn(/^Cài đặt/)?.click();
  await sleep(1200);
}
for (let i = 0; i < 10; i++) {
  const el = [...document.querySelectorAll("[data-memory-item]")].find((x) => x.innerText.startsWith("E2-thu"));
  if (!el) break;
  el.querySelector("[data-memory-delete]").click();
  await sleep(400);
}
const con = [...document.querySelectorAll("[data-memory-item]")].length;
await toChat();
btn(/^Hội thoại mới/)?.click();
delete window.__e2GiuLai;
return { soMucConLai: con };
