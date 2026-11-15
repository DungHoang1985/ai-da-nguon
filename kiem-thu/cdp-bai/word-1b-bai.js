// Đợt 1B: các bài Chạy thật (giao việc cho model thật trong Word). Ghép trước word-1b-hoi.js cùng một tệp chọn bài:
//   echo 'window.__aidnChon = "soat";' > chon.js
//   node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1b-bai.js chon.js kiem-thu/cdp-bai/word-1b-hoi.js
// Bài cần tệp đính kèm: trước đó chạy  node kiem-thu/cdp.mjs file 9229 'input[type=file]' <tệp>
// Tệp thử: node kiem-thu/soat-thu.mjs <sach|loi|v2|cong-van> <tệp>, node kiem-thu/pdf-quet-thu.mjs <tệp.pdf>
window.__aidnBai1B = {
  // Tiêu chí 1 — mở bản "loi" (9 lỗi cài sẵn) hoặc bản "sach" (không được báo nhầm)
  soat: "Soát nhất quán văn bản này trước khi trình ký. Liệt kê đầy đủ từng điểm cần sửa, mỗi điểm một dòng, kèm số điều và số đoạn; cuối cùng ghi tổng số điểm. Chỉ soát, không sửa gì.",
  // Tiêu chí 2 — mở bản "sach"
  chuThich: "Điều 4.3 quy định gì về Phí Dịch Vụ, và có chú thích cuối trang nào kèm theo không? Chân trang và đầu trang của văn bản ghi gì?",
  // Tiêu chí 3 — mở bản "sach", đính kèm bản "v2"
  soSanh: "So văn bản đang mở với tệp đính kèm (phiên bản khác của hợp đồng). Liệt kê đầy đủ mọi chỗ khác nhau, mỗi chỗ một dòng kèm số điều; cuối cùng ghi tổng số chỗ khác. Không sửa gì trong văn bản.",
  // Tiêu chí 4 — đính kèm PDF quét 3 trang
  pdf: "Không dùng công cụ. Tệp đính kèm là bản quét. Tóm tắt tệp trong 3 câu, rồi trả lời đúng ba số liệu: (1) Giá trị hợp đồng; (2) Diện tích sàn bàn giao; (3) Số tiền giữ lại bảo hành.",
  // Tiêu chí 5 (bổ sung Chạy thật) — mở công văn lề sai
  theThuc: "Kiểm thể thức trang của công văn này theo chuẩn văn bản hành chính: lề, khổ giấy, phông và cỡ chữ. Nêu rõ từng chỗ lệch chuẩn. Không sửa gì.",
  // Suy nghĩ kỹ + giữ khối suy nghĩ qua nhiều bước công cụ — mở bản "sach"
  nghiKy: "Đọc Điều 6 và Điều 7 rồi cho biết: nếu Bên Thuê chậm trả Tiền Thuê 50 ngày thì Bên Cho Thuê có những quyền gì theo hợp đồng này? Dẫn số điều cho từng quyền.",
  // Soát lập luận trước khi ghi — mở bản "sach", bật tùy chọn trong Cài đặt
  soatLapLuan: "Sửa Điều 9.2: đổi thời hạn khắc phục từ 07 ngày thành 05 ngày, giữ nguyên phần còn lại của câu.",
};
