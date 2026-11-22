// Đợt 12 phần E — E3 Chạy thật: model thật (nguồn đang chọn) theo bộ nhớ; AI tự đề xuất nhớ khi người dùng nói "từ giờ…".
// Mục thử có tiền tố "E3-thu" và bị xóa lại qua giao diện ở cuối bài. Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-12e-chay-that.js
const out = { buoc: [] };
const ok = (ten, dat, chiTiet) => out.buoc.push({ ten, dat: dat === null ? null : !!dat, ...(chiTiet !== undefined ? { chiTiet } : {}) });
// Đúng câu anh đặt ở tiêu chí E3 (không tiền tố thử: tiền tố làm model coi là chữ thử); dọn theo đúng chữ này
const A = "Trả lời kết luận trước, dùng bảng Tiêu chí · Kết quả.";
const openSettings = async () => {
  for (let i = 0; i < 50 && !document.querySelector("[data-memory-card]"); i++) {
    if (i % 5 === 0) btn(/^Cài đặt/)?.click();
    await sleep(300);
  }
  return document.querySelector("[data-memory-card]");
};
// Khung vừa tải lại: chờ nút chọn model
for (let i = 0; i < 60 && !btn(/^Model đang dùng/); i++) await sleep(300);
out.nguon = btn(/^Model đang dùng/) ? label(btn(/^Model đang dùng/)) : "chưa thấy nút chọn model";
try {
  const card = await openSettings();
  setNative(card.querySelector("[data-memory-input]"), A);
  await sleep(300);
  card.querySelector("[data-memory-add]").click();
  await sleep(500);
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  await sleep(500);
  const r1 = await send("Hợp đồng thuê văn phòng 3 năm giữa hai doanh nghiệp có bắt buộc công chứng không? Không cần đọc tài liệu.", { timeout: 240000, tail: 2500 });
  const tables = [...document.querySelectorAll("table")];
  const last = tables[tables.length - 1];
  const head = last ? [...last.querySelectorAll("th")].map((x) => x.innerText.trim()) : [];
  ok("Lượt 1 (hội thoại mới): khung báo gửi kèm bộ nhớ", !!document.querySelector("[data-mem-used]"), document.querySelector("[data-mem-used]")?.innerText);
  ok("Câu trả lời có bảng cột Tiêu chí | Kết quả", head.some((h) => /Tiêu chí/i.test(h)) && head.some((h) => /Kết quả/i.test(h)), head.join(" | "));
  out.traLoi1 = r1.khung;

  btn(/^Hội thoại mới/)?.click();
  await sleep(500);
  const r2 = await send("Từ giờ, công văn của tôi luôn ký thừa lệnh Giám đốc, chức danh người ký là Phó Giám đốc. Ghi nhớ giúp tôi.", { timeout: 240000, tail: 1500 });
  const prop = document.querySelector("[data-mem-proposal]");
  ok("AI tự đề xuất nhớ (thẻ AI muốn nhớ)", !!prop, prop ? prop.innerText.replace(/\n+/g, " | ") : r2.khung);
  if (prop && prop.querySelector("[data-mem-drop]")) {
    prop.querySelector("[data-mem-drop]").click();
    await sleep(400);
    ok("Bấm Bỏ: không lưu", document.querySelector('[data-mem-proposal="dropped"]') !== null);
  }
  out.traLoi2 = r2.khung;
} catch (e) {
  out.loi = String((e && e.stack) || e);
}
const first = (x) => x.innerText.split(String.fromCharCode(10))[0];
await openSettings();
for (let i = 0; i < 10; i++) {
  const el = [...document.querySelectorAll("[data-memory-item]")].find((x) => first(x) === A);
  if (!el) break;
  el.querySelector("[data-memory-delete]").click();
  await sleep(400);
}
ok("Dọn: không còn mục thử", ![...document.querySelectorAll("[data-memory-item]")].some((x) => first(x) === A));
out.soMucConLai = document.querySelectorAll("[data-memory-item]").length;
await toChat();
return out;
