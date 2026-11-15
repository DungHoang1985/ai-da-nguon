// Dọn sau khi thử: xóa các hội thoại do bài thử tạo ra (theo tiêu đề), và nếu đặt window.__aidnNguonCu thì chọn lại nhà cung cấp đó
// (model của từng nguồn add-in tự nhớ). Chạy trong khung của từng ứng dụng.
const TITLES = [
  /^Trong bảng của văn bản, sửa số tiền cọc/,
  /^Không dùng công cụ\. Một năm có mấy quý/,
  /^Không dùng công cụ\. Mô tả ngắn ảnh/,
  /^\/aidn-kiem-thu-ky-nang/,
  /^Tìm trên web rồi trả lời một câu/,
  /^Việc của phiên này: kiểm văn bản thử AIDN-DAI/,
  /^Ghi số 5 vào ô D7/,
  /^Không dùng công cụ ghi\. Bài này có mấy slide/,
  // Thêm tiêu đề của bài thử khác: đặt window.__aidnTieuDeThem = ["^chuỗi mẫu", …] ở một tệp ghép phía trước
  ...(window.__aidnTieuDeThem || []).map((s) => new RegExp(s)),
];
// Hội thoại "Xử lý các nhận xét…" trùng tên câu gợi ý có sẵn: chỉ xóa bản mới nhất (bản do bài thử tạo và vừa mở lại)
const ONCE = /^(▶ )?Xử lý các nhận xét trong văn bản/;
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(400);
btn(/^Lịch sử hội thoại/).click();
await sleep(1200);
const rows = () => [...document.querySelectorAll('[title="Mở lại"]')];
const count = () => Number((/Lịch sử trong \S+ \((\d+)\)/.exec(document.body.innerText) || [])[1]);
const truoc = count();
const daXoa = [];
let onceDone = !window.__aidnXoaXuLyNhanXet;
for (let guard = 0; guard < 60; guard++) {
  const r = rows().find((x) => {
    const t = x.innerText.replace(/^▶ /, "");
    return TITLES.some((re) => re.test(t)) || (!onceDone && ONCE.test(t));
  });
  if (!r) break;
  if (ONCE.test(r.innerText.replace(/^▶ /, ""))) onceDone = true;
  daXoa.push(r.innerText.split("\n")[0].slice(0, 50));
  [...r.parentElement.querySelectorAll("button")].find((b) => /Xóa/.test(b.innerText)).click();
  await sleep(1000);
}
const sau = count();
btn(/^Đóng$/)?.click();
await sleep(300);
let nguon = null;
if (window.__aidnNguonCu) {
  btn(/^Model đang dùng/).click();
  await sleep(700);
  const sel = document.querySelector('select[aria-label="Nhà cung cấp"]');
  if (sel.value !== window.__aidnNguonCu) {
    setNative(sel, window.__aidnNguonCu);
    await sleep(2500);
  }
  document.body.click();
  await sleep(500);
  nguon = label(btn(/^Model đang dùng/));
}
return { lichSuTruoc: truoc, lichSuSau: sau, daXoa, modelSauKhiTraLai: nguon };
