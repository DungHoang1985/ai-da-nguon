// Đợt 1B: bật / tắt hai tùy chọn mới rồi trả về trạng thái. Đặt trước: window.__aidnCongTac = { nghiKy: true|false, soatLapLuan: true|false }
// (bỏ trống khóa nào thì không đụng tới khóa đó). Bài thử phải trả cả hai về tắt khi xong.
const want = window.__aidnCongTac || {};
const out = {};
await toChat();
if (want.nghiKy !== undefined) {
  btn(/^Thêm: tệp/).click();
  await sleep(600);
  const item = [...document.querySelectorAll('[role="menuitemcheckbox"]')].find((x) => /Suy nghĩ kỹ/.test(x.innerText));
  if ((item.getAttribute("aria-checked") === "true") !== want.nghiKy) item.click();
  await sleep(400);
  document.body.click();
  await sleep(300);
  out.nghiKy = [...document.querySelectorAll("span")].some((x) => /Nghĩ kỹ/.test(x.innerText || ""));
}
if (want.soatLapLuan !== undefined) {
  btn(/Cài đặt/).click();
  await sleep(1200);
  const sw = [...document.querySelectorAll('input[role="switch"]')].find((x) => /soát lập luận/i.test((x.closest("div") && x.closest("div").innerText) || x.getAttribute("aria-label") || ""));
  if (!sw) return { loi: "Không thấy công tắc Soát lập luận", khung: tail(500) };
  if (sw.checked !== want.soatLapLuan) sw.click();
  await sleep(500);
  out.soatLapLuan = sw.checked;
  await toChat();
}
return out;
