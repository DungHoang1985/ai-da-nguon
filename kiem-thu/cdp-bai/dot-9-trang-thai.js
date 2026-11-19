// Đợt 9: đọc trạng thái khung — phiên bản, có thẻ chọn chế độ không, chế độ trên nhãn model, các nguồn trong bảng chọn (chỉ mã nguồn, không đọc kho lưu)
await sleep(3000);
const out = { phienBan: (document.body.innerText.match(/v\d+\.\d+\.\d+/) || [""])[0] };
const card = document.querySelector('[data-aidn="chon-che-do"]');
out.theChonCheDo = !!card;
if (card) out.chuThe = card.innerText.slice(0, 900);
else {
  const m = btn(/^Model đang dùng/);
  out.nhanModel = m ? label(m) : null;
  if (m) {
    m.click();
    await sleep(800);
    out.nguon = [...document.querySelectorAll('select[aria-label="Nhà cung cấp"] option')].map((o) => o.value);
    out.nutCheDo = [...document.querySelectorAll('[data-aidn="che-do"] button')].map((b) => `${b.innerText}${b.getAttribute("aria-pressed") === "true" ? " (đang bật)" : ""}`);
    out.model = [...document.querySelectorAll('[role="listbox"][aria-label="Model"] [role="option"]')].map((o) => o.getAttribute("title")).slice(0, 30);
    if (document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click();
    await sleep(300);
  }
}
return out;
