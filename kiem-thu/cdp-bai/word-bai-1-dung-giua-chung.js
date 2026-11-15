// Bài 1 (Đợt 0B): bấm Dừng giữa chuỗi sửa rồi hỏi "anh vừa làm gì". Lượt hai cũng dùng để xem số token đọc từ bộ nhớ đệm.
// Chạy sau word-dung-van-ban-thu.js, đã chọn model.
btn(/^Hội thoại mới/)?.click();
await sleep(500);
const done = () => (document.body.innerText.match(/✅ (Thay|Sửa đoạn)/g) || []).length;
const luot1 = await send(
  "Sửa lần lượt từng điều, mỗi điều một lần gọi công cụ riêng, làm xong điều này mới sang điều kia: " +
    "Điều 1 đổi “12 tháng” thành “24 tháng”; Điều 2 đổi “20 triệu” thành “25 triệu”; Điều 3 đổi “2 tháng” thành “3 tháng”; " +
    "Điều 4 đổi “ngày 05” thành “ngày 10”; Điều 5 đổi “hết thời hạn” thành “hết thời hạn hoặc theo thỏa thuận của hai bên”.",
  { stopWhen: () => done() >= 2 }
);
const soBuocXong = done();
const luot2 = await send("Anh vừa làm gì? Liệt kê điều nào đã sửa xong, điều nào chưa sửa. Không sửa thêm gì.");
const token = (document.body.innerText.match(/Token: [^\n]+/g) || []).slice(-2);
return { luot1, soBuocXongKhiDung: soBuocXong, luot2, dongToken: token };
