// Đợt 10: mở hội thoại mới trước mỗi bài của bộ bài trước / sau (để mỗi bài bắt đầu từ lịch sử rỗng, số token so được với nhau).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-10-moi.js
await toChat();
if (document.querySelector('select[aria-label="Nhà cung cấp"]')) {
  btn(/^Model đang dùng/).click();
  await sleep(500);
}
btn(/^Hội thoại mới/)?.click();
await sleep(800);
return { phienBan: (document.body.innerText.match(/\bv4\.\d+(\.\d+)?/) || [])[0] || null };
