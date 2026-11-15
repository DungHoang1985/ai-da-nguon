// Bài 4 (Đợt 0B): đính kèm ảnh, nhập kỹ năng qua giao diện, tìm trên web. Chạy theo ba pha vì việc đặt tệp phải làm từ ngoài
// (node kiem-thu/cdp.mjs file …) xen giữa. Đặt window.__aidnPha trước khi chạy: "anh" | "kynang-mo" | "kynang" | "web".
const pha = window.__aidnPha;
const out = { pha };

if (pha === "anh") {
  // Trước đó: cdp.mjs file 9229 'input[type=file][accept*="image/png"]' assets/icon-128.png
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  for (let i = 0; i < 20 && !/🖼/.test(document.body.innerText); i++) await sleep(300);
  out.theTep = (document.body.innerText.match(/🖼[^\n]*/) || ["(không thấy thẻ ảnh)"])[0];
  const r = await send("Không dùng công cụ. Mô tả ngắn ảnh đính kèm trong hai câu: hình gì, màu chủ đạo nào?", { tail: 700 });
  out.traLoi = r.khung;
} else if (pha === "kynang-mo") {
  btn(/^Cài đặt/).click();
  await sleep(1200);
  out.coONhapTep = !!document.querySelector('input[type=file][accept*=".skill"]');
} else if (pha === "kynang") {
  // Trước đó: cdp.mjs file 9229 'input[type=file][accept*=".skill"]' <tệp SKILL.md thử>
  await sleep(1500);
  out.baoCaoNhap = (document.body.innerText.match(/Đã (nhập|thêm)[^\n]*/) || ["(không thấy báo cáo)"])[0];
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  // Gõ "/" phải thấy kỹ năng trong danh sách
  const ta = document.querySelector("textarea");
  setNative(ta, "/aidn-kiem");
  await sleep(600);
  out.hienTrongMenu = [...document.querySelectorAll('[role="option"]')].map((x) => x.innerText.split("\n")[0]).slice(0, 5);
  setNative(ta, "");
  const r = await send("/aidn-kiem-thu-ky-nang chạy đi", { tail: 500 });
  out.traLoi = r.khung;
  // Dọn: xóa kỹ năng thử
  btn(/^Cài đặt/).click();
  await sleep(1200);
  if (!btn(/^Ẩn danh sách$/)) btn(/^Xem danh sách$/)?.click();
  await sleep(800);
  const row = [...document.querySelectorAll("span")].find((x) => x.innerText.trim() === "/aidn-kiem-thu-ky-nang");
  const del = row && [...row.closest("div").parentElement.querySelectorAll("button")].find((b) => /^Xóa$/.test(b.innerText.trim()));
  if (del) {
    del.click();
    await sleep(1200);
  }
  out.daXoaKyNangThu = !!del && !/\/aidn-kiem-thu-ky-nang/.test(document.body.innerText.replace(/Đã xóa kỹ năng[^\n]*/, ""));
  await toChat();
} else if (pha === "web") {
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  const dangBat = () => /🌐 Web/.test(document.body.innerText);
  const batTruoc = dangBat();
  const toggle = async () => {
    btn(/^Thêm: tệp/).click();
    await sleep(600);
    const item = [...document.querySelectorAll('[role="menuitemcheckbox"]')].find((x) => /Tìm trên web/.test(x.innerText));
    if (!item) return false;
    item.click();
    await sleep(500);
    document.body.click();
    await sleep(300);
    return true;
  };
  if (!batTruoc) out.coMucTimWeb = await toggle();
  out.daBatWeb = dangBat();
  const r = await send("Tìm trên web rồi trả lời một câu, kèm tên trang nguồn: phiên bản Node.js LTS mới nhất hiện nay là số mấy?", { tail: 900, timeout: 180000 });
  out.traLoi = r.khung;
  if (!batTruoc && dangBat()) await toggle(); // trả lại trạng thái cũ của người dùng
  out.traLaiTrangThaiCu = dangBat() === batTruoc;
}
return out;
