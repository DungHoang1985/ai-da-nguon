// Dựng tệp .docx "thông tin Bên B" (bịa để thử) làm tệp đính kèm cho bài Chạy thật C6 của Đợt 8 (điền mẫu từ tệp đính kèm).
//   node kiem-thu/thong-tin-thu.mjs <tệp ra .docx>
// Mọi thông tin là bịa (số CCCD không có thật), không có thông tin của người thật.
import { buildDocx, run } from "./docx-thu.mjs";

const [, , out] = process.argv;
if (!out) {
  console.error("Thiếu tệp ra");
  process.exit(1);
}
const p = (t, b = false) => `<w:p>${run(t, { b })}</w:p>`;
const body = [
  p("THÔNG TIN BÊN ĐẶT CỌC (BÊN B) — DỮ LIỆU THỬ", true),
  p("Tên tổ chức: Công ty Cổ phần Đầu tư Thử Nghiệm Minh An"),
  p("Người đại diện theo pháp luật: Ông Lê Văn Thử, chức vụ Tổng Giám đốc"),
  p("Căn cước công dân số: 001099012345, cấp ngày 15/03/2021 tại Cục Cảnh sát quản lý hành chính về trật tự xã hội"),
  p("Địa chỉ trụ sở: Số 9 đường Thử Nghiệm, phường Cầu Giấy, Hà Nội"),
].join("");
await buildDocx(out, { body });
console.log(`Đã tạo ${out}`);
