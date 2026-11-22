// Mở Cài đặt, cuộn tới mục Bộ nhớ (để chụp)
if (!document.querySelector("[data-memory-card]")) {
  btn(/^Cài đặt/)?.click();
  await sleep(1200);
}
document.querySelector("[data-memory-card]").scrollIntoView();
return { ok: true };
