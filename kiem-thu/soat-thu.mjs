// Tệp thử cho Đợt 1B (Word: soát và tra). Nội dung bịa để thử, không có thông tin thật.
//
//   node kiem-thu/soat-thu.mjs sach <tệp ra>        hợp đồng thuê văn phòng 14 điều + phụ lục gõ tay, KHÔNG có lỗi cài sẵn; có chú thích cuối trang,
//                                                   chú thích cuối văn bản, đầu trang, chân trang, bookmark, trường dẫn chiếu, ô điền mẫu, hộp văn bản,
//                                                   một sửa đổi được theo dõi và một nhận xét của "Đối tác"
//   node kiem-thu/soat-thu.mjs loi <tệp ra>         như trên nhưng cài sẵn 9 lỗi: 3 thuật ngữ lệch, 2 dẫn chiếu hỏng, 1 chỗ nhảy số,
//                                                   1 số tiền lệch chữ, 2 chỗ trống (danh sách ở LOI_CAI_SAN)
//   node kiem-thu/soat-thu.mjs v2 <tệp ra>          "phiên bản khác" của bản sạch, lệch đúng 10 chỗ (danh sách ở KHAC_BIET_V2), không có sửa đổi được theo dõi
//   node kiem-thu/soat-thu.mjs cong-van <tệp ra>    công văn có lề sai chuẩn (lề trái 20 mm, lề trên 15 mm), phông Arial 12 lẫn trong văn bản
//   node kiem-thu/soat-thu.mjs cong-van-dung <tệp ra>  công văn đúng thể thức (A4, lề 20/20/30/15 mm, Times New Roman 14)
//   node kiem-thu/soat-thu.mjs json <tệp ra.json>   các đoạn của ba bản hợp đồng dưới dạng dữ liệu (cho kiểm thử máy: test/unit/dot1b.test.ts)
//
// Tên tệp thử phải bắt đầu bằng "aidn-thu-" hoặc "tham-do-" để kiem-thu/mo-office.ps1 nhận là tệp thử.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { para, run, del, ins, commented, field, footnoteRef, endnoteRef, bookmark, sdt, textbox, runFont, buildDocx } from "./docx-thu.mjs";

// Mỗi điều: [tiêu đề, [khoản…]]; khoản là chuỗi, hoặc [chuỗi, [điểm…]], hoặc { free: chuỗi } = đoạn chữ thường không đánh số.
// Mốc trong chuỗi: {-cũ-}{+mới+} sửa đổi của "Đối tác" · [[fn:n]] chú thích cuối trang · [[en:n]] chú thích cuối văn bản ·
// [[bm:Tên|chữ]] bookmark · [[ref:Tên|kết quả]] trường REF · [[sdt:tag|tiêu đề|chữ]] ô điền mẫu · [[c:chữ được neo|nhận xét]] nhận xét
const DIEU = [
  ["Định nghĩa", [
    "“Tòa Nhà” là tòa nhà văn phòng Hồng Lĩnh Tower tại số 18 đường Trần Phú, do Bên Cho Thuê làm chủ sở hữu.",
    "“Diện Tích Thuê” là phần diện tích 420 m² tại tầng 9 của Tòa Nhà, có vị trí và ranh giới mô tả tại Phụ lục 01.",
    "“Ngày Bàn Giao” là ngày hai bên ký biên bản bàn giao Diện Tích Thuê theo Điều 5.",
    "“Tiền Thuê” là khoản tiền Bên Thuê trả hằng tháng cho việc sử dụng Diện Tích Thuê, xác định theo Điều 4.[[fn:1]]",
    "“Phí Dịch Vụ” là khoản phí quản lý vận hành Tòa Nhà mà Bên Thuê trả theo Điều 4.3.",
    "“Tiền Đặt Cọc” là khoản tiền bảo đảm thực hiện hợp đồng quy định tại Điều 6.",
    "“Ngày Làm Việc” là các ngày từ thứ Hai đến thứ Sáu, trừ ngày nghỉ lễ, tết theo quy định của pháp luật.",
  ]],
  ["Đối tượng và mục đích thuê", [
    "Bên Cho Thuê đồng ý cho thuê và Bên Thuê đồng ý thuê Diện Tích Thuê để làm văn phòng làm việc.",
    "Diện Tích Thuê được đo theo tim tường; chênh lệch dưới 1% so với số đo nêu tại Phụ lục 01 không làm thay đổi Tiền Thuê.",
    "Bên Thuê không được dùng Diện Tích Thuê vào mục đích khác, như kho hàng, lưu trú, lớp học v.v... nếu chưa có chấp thuận bằng văn bản của Bên Cho Thuê.",
  ]],
  ["Thời hạn thuê", [
    "Thời hạn thuê là 05 (năm) năm kể từ Ngày Bàn Giao.",
    "Bên Thuê được ưu tiên gia hạn nếu gửi đề nghị bằng văn bản chậm nhất 06 (sáu) tháng trước ngày hết hạn.",
    "Hết thời hạn thuê mà không gia hạn thì hợp đồng chấm dứt và Bên Thuê hoàn trả Diện Tích Thuê theo Điều 12.",
  ]],
  ["Tiền thuê và phí dịch vụ", [
    "[[bm:GiaThue|Tiền Thuê là 120.000.000 đồng (Một trăm hai mươi triệu đồng) mỗi tháng]], chưa gồm thuế giá trị gia tăng.",
    "Tiền Thuê được giữ cố định trong 02 (hai) năm đầu; từ năm thứ ba, mỗi năm tăng 5% so với năm liền trước.",
    "Phí Dịch Vụ là 85.000 đồng cho mỗi mét vuông mỗi tháng và được điều chỉnh theo thông báo của ban quản lý Tòa Nhà, nhưng mỗi lần tăng không quá 10%.[[fn:2]]",
  ]],
  ["Bàn giao", [
    "Bên Cho Thuê bàn giao Diện Tích Thuê cho Bên Thuê chậm nhất vào ngày 01 tháng 02 năm 2027.",
    "Tại Ngày Bàn Giao, hai bên lập biên bản ghi nhận hiện trạng, chỉ số công tơ điện và danh mục trang thiết bị kèm theo.",
    "Nếu Bên Cho Thuê chậm bàn giao quá 30 (ba mươi) ngày, Bên Thuê được miễn Tiền Thuê tương ứng với gấp đôi số ngày chậm.",
  ]],
  ["Đặt cọc", [
    "Bên Thuê chuyển Tiền Đặt Cọc bằng 03 (ba) tháng Tiền Thuê, tương đương 360.000.000 đồng (Ba trăm sáu mươi triệu đồng), trong vòng 10 (mười) Ngày Làm Việc kể từ ngày ký hợp đồng.",
    "Tiền Đặt Cọc không tính lãi và được hoàn trả trong vòng 30 ngày kể từ ngày Bên Thuê hoàn thành mọi nghĩa vụ khi chấm dứt hợp đồng.",
    ["Bên Cho Thuê được khấu trừ Tiền Đặt Cọc trong các trường hợp sau:", [
      "Bên Thuê chậm trả Tiền Thuê hoặc Phí Dịch Vụ quá thời hạn quy định tại khoản 3 Điều 7;",
      "Bên Thuê gây hư hỏng cho Tòa Nhà mà không khắc phục;",
      "Bên Thuê đơn phương chấm dứt hợp đồng trái với Điều 11.",
    ]],
  ]],
  ["Thanh toán", [
    "[[bm:_Ref7001|Tiền Thuê và Phí Dịch Vụ]] được thanh toán theo quý, trong vòng 10 ngày đầu của mỗi quý, bằng chuyển khoản vào tài khoản của Bên Cho Thuê số [[sdt:so_tai_khoan|Số tài khoản|0011 2233 4455]] tại Ngân hàng Thương mại Cổ phần Đông Hải.",
    "Bên Cho Thuê xuất hóa đơn cho từng kỳ thanh toán trước ngày đến hạn ít nhất 07 ngày.",
    "Nếu chậm thanh toán, Bên Thuê trả lãi chậm trả 0,05% mỗi ngày trên số tiền chậm trả; chậm quá 45 ngày là vi phạm nghiêm trọng theo Điều 11.2.",
    "Lãi chậm trả nêu tại khoản 3 Điều này không vượt quá mức lãi suất tối đa theo Điều 357 của Bộ luật Dân sự.",
  ]],
  ["Sử dụng và cải tạo", [
    "Bên Thuê sử dụng Diện Tích Thuê đúng công năng và tuân thủ nội quy Tòa Nhà.",
    "Mọi cải tạo làm thay đổi kết cấu, hệ thống điện hoặc phòng cháy chữa cháy phải được Bên Cho Thuê chấp thuận bằng văn bản trước khi thi công.",
    "Bên Cho Thuê trả lời đề nghị cải tạo trong vòng {-15-}{+10+} Ngày Làm Việc; quá thời hạn này mà không trả lời thì coi như đồng ý.",
  ]],
  ["Bảo trì và bảo hiểm", [
    "Bên Cho Thuê chịu trách nhiệm sửa chữa kết cấu, mặt ngoài và hệ thống kỹ thuật chung của Tòa Nhà.",
    "Hư hỏng thuộc trách nhiệm của Bên Cho Thuê phải được khắc phục trong vòng 07 ngày kể từ ngày nhận thông báo; trường hợp khẩn cấp là 24 giờ.",
    "Bên Thuê mua bảo hiểm cháy nổ cho tài sản của mình với mức trách nhiệm không thấp hơn 5.000.000.000 đồng (Năm tỷ đồng).[[fn:3]]",
  ]],
  ["Quyền và nghĩa vụ của các bên", [
    "Bên Cho Thuê bảo đảm quyền sử dụng ổn định của Bên Thuê trong suốt thời hạn thuê nêu tại Điều 3.",
    "Bên Thuê trả Tiền Thuê, Phí Dịch Vụ đầy đủ, đúng hạn theo Điều [[ref:_Ref7001|7.1]].",
    "Mỗi bên chịu trách nhiệm về hành vi của nhân viên, khách và nhà thầu của mình trong Tòa Nhà.[[en:1]]",
  ]],
  ["Chấm dứt hợp đồng", [
    "Hợp đồng chấm dứt khi hết thời hạn thuê mà không gia hạn theo Điều 3.2, hoặc theo thỏa thuận bằng văn bản của hai bên.",
    "Bên Cho Thuê được đơn phương chấm dứt hợp đồng nếu Bên Thuê chậm thanh toán quá thời hạn nêu tại Điều 7.3 mà không khắc phục trong 15 ngày kể từ ngày nhận thông báo.",
    "Bên Thuê được chấm dứt hợp đồng trước hạn sau năm thứ hai, với điều kiện báo trước 06 tháng và trả khoản phí chấm dứt bằng 02 tháng Tiền Thuê.",
    "Bên vi phạm chịu phạt 8% giá trị phần nghĩa vụ bị vi phạm và bồi thường thiệt hại thực tế.",
  ]],
  ["Hoàn trả mặt bằng", [
    "Khi chấm dứt hợp đồng, Bên Thuê hoàn trả Diện Tích Thuê trong vòng 15 ngày theo hiện trạng lúc nhận bàn giao, trừ hao mòn tự nhiên.",
    "Chậm hoàn trả, Bên Thuê trả tiền sử dụng bằng [[c:150%|Đề nghị giảm xuống 120%.]] Tiền Thuê tính theo ngày cho thời gian chậm.",
  ]],
  ["Bất khả kháng và thông báo", [
    "Bên gặp sự kiện bất khả kháng thông báo cho bên kia trong vòng 07 ngày kể từ ngày sự kiện xảy ra, kèm tài liệu chứng minh.",
    "Mọi thông báo lập bằng văn bản, gửi tới địa chỉ nêu tại phần đầu hợp đồng; bên thay đổi địa chỉ báo cho bên kia trong vòng 05 Ngày Làm Việc.",
    { free: "Điều 3 của hợp đồng này và Điều 12 vẫn được áp dụng trong thời gian xảy ra sự kiện bất khả kháng." },
  ]],
  ["Giải quyết tranh chấp và hiệu lực", [
    "Tranh chấp được giải quyết bằng thương lượng trong vòng 30 ngày; không thành thì đưa ra Tòa án có thẩm quyền.",
    "Hợp đồng có hiệu lực từ ngày ký và được lập thành 04 (bốn) bản có giá trị như nhau, mỗi bên giữ 02 (hai) bản.",
    "Các phụ lục là phần không tách rời của hợp đồng.",
  ]],
];

const MO_DAU = [
  "Hôm nay, ngày 10 tháng 11 năm 2026, tại trụ sở Bên Cho Thuê, các bên gồm:",
];
// Phụ lục đánh số GÕ TAY
const PHU_LUC = [
  "1. Mô tả mặt bằng",
  "1.1 Diện Tích Thuê nằm tại tầng 9, phía Đông Nam của Tòa Nhà, có 02 mặt thoáng.",
  "1.2 Trần thạch cao, sàn nâng kỹ thuật và hệ thống điều hòa trung tâm đã được hoàn thiện.",
  "2. Trang thiết bị bàn giao",
  "2.1 Hệ thống chiếu sáng gồm 96 bộ đèn âm trần.",
  "2.2 Hệ thống báo cháy và chữa cháy tự động theo thiết kế đã được thẩm duyệt.",
  "2.3 Bốn (04) tủ điện tầng và 01 tủ mạng trung tâm.",
  "3. Giờ hoạt động",
  "3.1 Tòa Nhà mở cửa từ 6 giờ đến 22 giờ các Ngày Làm Việc; ngoài giờ này Bên Thuê đăng ký trước với ban quản lý.",
];

export const CHU_THICH = [
  "Tiền Thuê chưa gồm thuế giá trị gia tăng 10% và được thanh toán bằng đồng Việt Nam.",
  "Mức Phí Dịch Vụ năm 2027 do ban quản lý Tòa Nhà thông báo ngày 01/12/2026.",
  "Giấy chứng nhận bảo hiểm phải được gửi cho Bên Cho Thuê trước Ngày Bàn Giao 05 ngày.",
];
export const CHU_THICH_CUOI = ["Nội quy Tòa Nhà ban hành kèm theo hợp đồng này là bản cập nhật tháng 10 năm 2026."];
export const DAU_TRANG = "Hợp đồng thuê văn phòng số 12/2026/HĐT-VP";
export const CHAN_TRANG = "Bản dự thảo lần 3 ngày 10/11/2026 — Lưu hành nội bộ";
export const HOP_VAN_BAN = "DỰ THẢO LẦN 3 — chưa có giá trị pháp lý";

/** 9 lỗi cài sẵn của bản "loi". Khóa: "điều.khoản", "điều.khoản.điểm" (a = 1), "mo-dau.n", "phu-luc.n", "free.điều" */
export const LOI_CAI_SAN = [
  { loai: "thuật ngữ lệch", o: "5.2", tu: "Tại Ngày Bàn Giao", thanh: "Tại ngày bàn giao" },
  { loai: "thuật ngữ lệch", o: "6.3.1", tu: "chậm trả Tiền Thuê", thanh: "chậm trả Tiền thuê" },
  { loai: "thuật ngữ lệch", o: "10.2", tu: "Phí Dịch Vụ", thanh: "Phí Quản Lý" },
  { loai: "dẫn chiếu hỏng", o: "free.13", tu: "và Điều 12 vẫn", thanh: "và Điều 19 vẫn" },
  { loai: "dẫn chiếu hỏng", o: "11.2", tu: "nêu tại Điều 7.3", thanh: "nêu tại khoản 6 Điều 7" },
  { loai: "nhảy số", o: "phu-luc.7", tu: "2.3 Bốn", thanh: "2.4 Bốn" },
  { loai: "số tiền lệch chữ", o: "4.1", tu: "(Một trăm hai mươi triệu đồng)", thanh: "(Một trăm hai mươi lăm triệu đồng)" },
  { loai: "chỗ trống", o: "mo-dau.1", tu: "ngày 10 tháng 11 năm 2026", thanh: "ngày ... tháng ... năm 2026" },
  { loai: "chỗ trống", o: "7.1", tu: "[[sdt:so_tai_khoan|Số tài khoản|0011 2233 4455]]", thanh: "………………" },
];

/** 10 chỗ khác của "phiên bản khác" so với bản sạch (theo bản hiện tại của bản sạch). kieu: sua | them-khoan | xoa-khoan */
export const KHAC_BIET_V2 = [
  { dieu: "3.1", o: "3.1", tu: "05 (năm) năm", thanh: "07 (bảy) năm" },
  { dieu: "4.1", o: "4.1", tu: "120.000.000 đồng (Một trăm hai mươi triệu đồng)", thanh: "135.000.000 đồng (Một trăm ba mươi lăm triệu đồng)" },
  { dieu: "4.2", o: "4.2", tu: "tăng 5%", thanh: "tăng 7%" },
  { dieu: "5.1", o: "5.1", tu: "ngày 01 tháng 02 năm 2027", thanh: "ngày 01 tháng 03 năm 2027" },
  { dieu: "6.2", o: "6.2", tu: "trong vòng 30 ngày", thanh: "trong vòng 45 ngày" },
  { dieu: "7.3", o: "7.3", tu: "0,05%", thanh: "0,08%" },
  { dieu: "8.2", o: "8.2", tu: "trước khi thi công.", thanh: "trước khi thi công. Chi phí thẩm tra hồ sơ cải tạo do Bên Thuê chịu." },
  { dieu: "9.4 (khoản mới sau 9.3)", o: "9.3", kieu: "them-khoan", thanh: "Bên Cho Thuê mua bảo hiểm mọi rủi ro tài sản cho phần kết cấu của Tòa Nhà." },
  { dieu: "11.3 (bị bỏ)", o: "11.3", kieu: "xoa-khoan" },
  { dieu: "Phụ lục 01, mục 2.1", o: "phu-luc.5", tu: "96 bộ đèn", thanh: "108 bộ đèn" },
];

/** Chuỗi có mốc → XML các run. tracked = false: lấy bản hiện tại (bỏ chữ đã xóa, giữ chữ thêm), bỏ nhận xét */
function toRuns(s, ctx) {
  let out = "";
  const re = /\{-(.*?)-\}|\{\+(.*?)\+\}|\[\[(fn|en|bm|ref|sdt|c):(.*?)\]\]/g;
  let pos = 0;
  let m;
  while ((m = re.exec(s))) {
    if (m.index > pos) out += run(s.slice(pos, m.index));
    pos = m.index + m[0].length;
    if (m[1] !== undefined) out += ctx.tracked ? del(m[1]) : "";
    else if (m[2] !== undefined) out += ctx.tracked ? ins(m[2]) : run(m[2]);
    else {
      const a = m[4].split("|");
      if (m[3] === "fn") out += ctx.extras ? footnoteRef(Number(a[0])) : "";
      else if (m[3] === "en") out += ctx.extras ? endnoteRef(Number(a[0])) : "";
      else if (m[3] === "bm") out += ctx.extras ? bookmark(a[0], run(a[1])) : run(a[1]);
      else if (m[3] === "ref") out += ctx.extras ? field(`REF ${a[0]} \\r \\h`, a[1]) : run(a[1]);
      else if (m[3] === "sdt") out += ctx.extras ? sdt(a[0], a[1], a[2]) : run(a[2]);
      else if (m[3] === "c") {
        if (ctx.tracked) {
          const id = ctx.comments.length;
          ctx.comments.push({ text: a[1], author: "Đối tác" });
          out += commented(id, run(a[0]));
        } else out += run(a[0]);
      }
    }
  }
  if (pos < s.length) out += run(s.slice(pos));
  return out;
}
/** Chữ của bản hiện tại (không mốc) — cho kiểm thử máy */
export const plain = (s) =>
  s
    .replace(/\{-(.*?)-\}/g, "")
    .replace(/\{\+(.*?)\+\}/g, "$1")
    .replace(/\[\[(?:fn|en):.*?\]\]/g, "")
    .replace(/\[\[(?:bm|ref):[^|\]]*\|(.*?)\]\]/g, "$1")
    .replace(/\[\[sdt:[^|]*\|[^|]*\|(.*?)\]\]/g, "$1")
    .replace(/\[\[c:([^|]*)\|.*?\]\]/g, "$1");

/** Danh sách đoạn của một bản: [{ text (có mốc), style, key }] theo thứ tự trong văn bản */
export function doan(variant = "sach") {
  const swaps = new Map();
  const add = (o, tu, thanh) => {
    if (!swaps.has(o)) swaps.set(o, []);
    swaps.get(o).push([tu, thanh]);
  };
  if (variant === "loi" || variant === "tham-do") LOI_CAI_SAN.forEach((e) => add(e.o, e.tu, e.thanh));
  if (variant === "v2") KHAC_BIET_V2.filter((e) => !e.kieu).forEach((e) => add(e.o, e.tu, e.thanh));
  const used = new Set();
  const fix = (key, text) => {
    let t = text;
    for (const [tu, thanh] of swaps.get(key) || []) {
      // bản v2 so trên bản hiện tại: tìm trên chữ đã bỏ mốc sửa đổi
      const base = variant === "v2" ? t.replace(/\{-(.*?)-\}/g, "").replace(/\{\+(.*?)\+\}/g, "$1") : t;
      if (!base.includes(tu)) throw new Error(`${variant} ${key}: không thấy “${tu}”`);
      t = base.replace(tu, thanh);
      used.add(key + "|" + tu);
    }
    return t;
  };
  const out = [];
  const push = (text, style, key) => out.push({ text: fix(key, text), style, key });
  push("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Title0", "tieu-de.1");
  push("HỢP ĐỒNG THUÊ VĂN PHÒNG", "Title0", "tieu-de.2");
  push("Số: 12/2026/HĐT-VP (bản thử của AI Đa Nguồn, nội dung không có thật)", "", "tieu-de.3");
  MO_DAU.forEach((t, k) => push(t, "", `mo-dau.${k + 1}`));
  push("BÊN CHO THUÊ: Công ty Cổ phần Bất động sản Hồng Lĩnh, sau đây gọi là “Bên Cho Thuê”.", "Party", "ben.1");
  push("BÊN THUÊ: Công ty TNHH Tư vấn Thiết kế Lam Giang, sau đây gọi là “Bên Thuê”.", "Party", "ben.2");
  push("Hai bên thống nhất ký hợp đồng thuê văn phòng với các điều khoản sau:", "", "mo-dau.9");
  const xoa = variant === "v2" ? new Set(KHAC_BIET_V2.filter((e) => e.kieu === "xoa-khoan").map((e) => e.o)) : new Set();
  const them = variant === "v2" ? new Map(KHAC_BIET_V2.filter((e) => e.kieu === "them-khoan").map((e) => [e.o, e.thanh])) : new Map();
  DIEU.forEach(([title, clauses], a) => {
    push(title, "Heading1", `${a + 1}`);
    let k = 0;
    clauses.forEach((cl) => {
      if (cl && cl.free) return push(cl.free, "Khoan", `free.${a + 1}`);
      k++;
      const key = `${a + 1}.${k}`;
      const [text, points] = Array.isArray(cl) ? cl : [cl, null];
      if (!xoa.has(key)) {
        push(text, "Heading2", key);
        (points || []).forEach((pt, j) => push(pt, "Heading3", `${key}.${j + 1}`));
      }
      if (them.has(key)) push(them.get(key), "Heading2", `${key}+`);
    });
  });
  push("PHỤ LỤC 01. MÔ TẢ DIỆN TÍCH THUÊ VÀ TRANG THIẾT BỊ", "TitlePL", "phu-luc.0");
  PHU_LUC.forEach((t, k) => push(t, "Khoan", `phu-luc.${k + 1}`));
  push("", "", "ky.0");
  push("ĐẠI DIỆN BÊN CHO THUÊ                                        ĐẠI DIỆN BÊN THUÊ", "Title0", "ky.1");
  push("(Ký, ghi rõ họ tên và đóng dấu)", "", "ky.2");
  push("Ông [[sdt:dai_dien_a|Người đại diện bên cho thuê|Trần Văn Minh]]                                                  Bà [[sdt:dai_dien_b|Người đại diện bên thuê|Lê Thu Hà]]", "", "ky.3");
  if (variant === "tham-do") push("Ô điền mẫu còn trống: [[sdt:o_trong|Ô chưa điền|]]", "", "ky.4");
  for (const [o, list] of swaps) for (const [tu] of list) if (!used.has(o + "|" + tu)) throw new Error(`${variant}: chưa áp được thay đổi ở ${o}`);
  return out;
}

export function hopDongSoat(variant = "sach") {
  const lists = [
    [`<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>`],
    [
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:suff w:val="space"/><w:lvlText w:val="Điều %1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>`,
      `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="680" w:hanging="680"/></w:pPr></w:lvl>`,
      `<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:pStyle w:val="Heading3"/><w:lvlText w:val="%3)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="400"/></w:pPr></w:lvl>`,
    ],
  ];
  const ctx = { tracked: variant !== "v2", extras: variant !== "v2", comments: [] };
  const body = doan(variant)
    .map((d) => {
      let xml = d.text ? toRuns(d.text, ctx) : "";
      if (d.style === "Title0" || d.style === "TitlePL") xml = xml.replace(/<w:r><w:t /g, '<w:r><w:rPr><w:b/></w:rPr><w:t ');
      if (d.key === "tieu-de.2" && ctx.extras) xml += textbox(HOP_VAN_BAN);
      if (d.style === "Party") xml = xml.replace(/<w:r><w:t xml:space="preserve">(BÊN [^:]+: )/, '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">$1</w:t></w:r><w:r><w:t xml:space="preserve">');
      const style = /^Heading|^Khoan$/.test(d.style) ? d.style : undefined;
      return para(xml, style ? { style } : {});
    })
    .join("");
  const spec = { body, lists, headingNumId: 2, comments: ctx.comments, page: { top: 20, bottom: 20, left: 30, right: 15 }, font: { name: "Times New Roman", pt: 13 } };
  if (ctx.extras) {
    spec.footnotes = CHU_THICH;
    spec.endnotes = CHU_THICH_CUOI;
    spec.header = DAU_TRANG;
    spec.footer = para(run(CHAN_TRANG + " — Trang ") + field("PAGE", "1"));
  }
  return spec;
}

// ───────────────────────── Công văn ─────────────────────────
export function congVan(dung = false) {
  const F = "Times New Roman";
  const p = (xml, jc) => `<w:p>${jc ? `<w:pPr><w:jc w:val="${jc}"/></w:pPr>` : ""}${xml}</w:p>`;
  const body = [
    p(runFont("CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN HỒNG LĨNH", F, 13, { b: true })),
    p(runFont("Số: 215/HL-PC", F, 13)),
    p(runFont("V/v đề nghị hướng dẫn thủ tục điều chỉnh chủ trương đầu tư", F, 12)),
    p(runFont("Hồng Lĩnh, ngày 10 tháng 11 năm 2026", F, 14, { i: true }), "right"),
    p(runFont("Kính gửi: Sở Kế hoạch và Đầu tư tỉnh (bản thử của AI Đa Nguồn, nội dung không có thật)", F, 14), "center"),
    p(runFont("Công ty Cổ phần Bất động sản Hồng Lĩnh là nhà đầu tư dự án Khu nhà ở Hồng Lĩnh Riverside, đã được chấp thuận chủ trương đầu tư năm 2024.", F, 14), "both"),
    p(runFont("Trong quá trình triển khai, Công ty có nhu cầu điều chỉnh tiến độ thực hiện dự án do vướng mắc trong công tác bồi thường, giải phóng mặt bằng tại phần diện tích 2,4 ha phía Tây dự án.", F, 14), "both"),
    // Đoạn lệch phông và cỡ (chỉ ở bản sai)
    p(runFont("Công ty kính đề nghị Quý Sở hướng dẫn thành phần hồ sơ, trình tự và thời hạn giải quyết thủ tục điều chỉnh chủ trương đầu tư trong trường hợp nêu trên.", dung ? F : "Arial", dung ? 14 : 12), "both"),
    p(runFont("Trân trọng cảm ơn.", F, 14), "both"),
    p(runFont("Nơi nhận: Như trên; Lưu: VT, PC.", F, 11)),
    p(runFont("TỔNG GIÁM ĐỐC", F, 14, { b: true }), "right"),
  ].join("");
  return {
    body,
    page: dung ? { top: 20, bottom: 20, left: 30, right: 15 } : { top: 15, bottom: 20, left: 20, right: 15 },
    font: { name: F, pt: 14 },
  };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [, , kind, out] = process.argv;
  if (!out) {
    console.log("Cách dùng: node kiem-thu/soat-thu.mjs <sach|loi|v2|cong-van|cong-van-dung|json> <tệp ra>");
    process.exit(1);
  }
  if (kind === "json") {
    const data = {};
    for (const v of ["sach", "loi", "v2"]) data[v] = doan(v).map((d) => ({ text: plain(d.text), style: d.style, key: d.key }));
    fs.writeFileSync(out, JSON.stringify({ doan: data, loi: LOI_CAI_SAN, v2: KHAC_BIET_V2 }, null, 1));
  } else if (kind === "sach" || kind === "loi" || kind === "v2" || kind === "tham-do") await buildDocx(out, hopDongSoat(kind));
  else if (kind === "cong-van") await buildDocx(out, congVan(false));
  else if (kind === "cong-van-dung") await buildDocx(out, congVan(true));
  else throw new Error("Loại phải là sach, loi, v2, cong-van, cong-van-dung hoặc json");
  console.log(`Đã tạo ${out}`);
}
