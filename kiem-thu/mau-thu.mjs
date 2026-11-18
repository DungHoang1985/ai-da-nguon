// Dựng mẫu Word thử cho Đợt 8 (điền mẫu): đủ 6 loại ô điền mẫu (content control) và các kiểu chỗ trống thường gặp
// trong mẫu hợp đồng, tờ khai Việt Nam ("……", "____", "[●]", "ngày   tháng   năm"), cùng một bảng thông tin để điền.
//   node kiem-thu/mau-thu.mjs <tệp ra .docx>
// Nội dung bịa để thử, không có thông tin thật. Tệp tự mở khung add-in (cùng mẫu với `cdp.mjs make word`).
import { buildDocx, esc, run } from "./docx-thu.mjs";

const [, , out] = process.argv;
if (!out) {
  console.error("Thiếu tệp ra");
  process.exit(1);
}

let id = 8000;
const pr = (tag, title, extra = "") => `<w:sdtPr><w:alias w:val="${esc(title)}"/><w:tag w:val="${esc(tag)}"/><w:id w:val="${id++}"/>${extra}`;
/** Ô chữ thường (PlainText) hoặc chữ định dạng (RichText); rỗng = đang hiện chữ gợi ý */
const ccText = (tag, title, text, { rich = false, lock = "" } = {}) =>
  `<w:sdt>${pr(tag, title, (lock ? `<w:lock w:val="${lock}"/>` : "") + (text ? "" : "<w:showingPlcHdr/>") + (rich ? "" : "<w:text/>"))}</w:sdtPr>` +
  `<w:sdtContent>${run(text || "Bấm để nhập")}</w:sdtContent></w:sdt>`;
const ccCheck = (tag, title, checked) =>
  `<w:sdt>${pr(tag, title, `<w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`)}</w:sdtPr>` +
  `<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:t>${checked ? "☒" : "☐"}</w:t></w:r></w:sdtContent></w:sdt>`;
const items = (list) => list.map((x) => `<w:listItem w:displayText="${esc(x)}" w:value="${esc(x)}"/>`).join("");
const ccDrop = (tag, title, list) =>
  `<w:sdt>${pr(tag, title, `<w:showingPlcHdr/><w:dropDownList>${items(list)}</w:dropDownList>`)}</w:sdtPr><w:sdtContent>${run("Chọn một mục")}</w:sdtContent></w:sdt>`;
const ccCombo = (tag, title, list) =>
  `<w:sdt>${pr(tag, title, `<w:showingPlcHdr/><w:comboBox>${items(list)}</w:comboBox>`)}</w:sdtPr><w:sdtContent>${run("Chọn hoặc nhập")}</w:sdtContent></w:sdt>`;
const ccDate = (tag, title) =>
  `<w:sdt>${pr(tag, title, `<w:showingPlcHdr/><w:date><w:dateFormat w:val="dd/MM/yyyy"/><w:lid w:val="vi-VN"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>`)}</w:sdtPr><w:sdtContent>${run("Chọn ngày")}</w:sdtContent></w:sdt>`;

const p = (inner, { center = false, b = false } = {}) => `<w:p>${center ? '<w:pPr><w:jc w:val="center"/></w:pPr>' : ""}${typeof inner === "string" && !inner.startsWith("<") ? run(inner, { b }) : inner}</w:p>`;
const cell = (inner, w = 3000) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${p(inner)}</w:tc>`;
const table = (rows) =>
  `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("")}</w:tblBorders></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>` +
  rows.map(([a, b]) => `<w:tr>${cell(a)}${cell(b, 6000)}</w:tr>`).join("") +
  `</w:tbl>`;

const body = [
  p("HỢP ĐỒNG ĐẶT CỌC (MẪU THỬ)", { center: true, b: true }),
  p(run("Số hợp đồng: ") + ccText("so_hd", "Số hợp đồng", "HĐĐC-THU-01", { lock: "sdtContentLocked" })),
  p("PHẦN 1. Ô ĐIỀN MẪU", { b: true }),
  p(run("Bên nhận đặt cọc (Bên A): ") + ccText("ben_a", "Tên Bên A", "")),
  p(run("Địa chỉ trụ sở Bên A: ") + ccText("dia_chi_a", "Địa chỉ Bên A", "", { rich: true })),
  p(run("Bên đặt cọc là: ") + ccCheck("la_ca_nhan", "Cá nhân", false) + run(" Cá nhân   ") + ccCheck("la_to_chuc", "Tổ chức", false) + run(" Tổ chức")),
  p(run("Hình thức thanh toán: ") + ccDrop("thanh_toan", "Hình thức thanh toán", ["Chuyển khoản", "Tiền mặt", "Kết hợp"])),
  p(run("Loại tài sản: ") + ccCombo("loai_ts", "Loại tài sản", ["Căn hộ chung cư", "Nhà ở riêng lẻ", "Đất nền"])),
  p(run("Ngày ký hợp đồng: ") + ccDate("ngay_ky", "Ngày ký")),
  p("PHẦN 2. CHỖ TRỐNG GÕ TAY", { b: true }),
  p("Họ và tên người đại diện Bên B: ……………………………………"),
  p("Số CCCD: ____________________ cấp ngày ………… tại …………………"),
  p("Số tiền đặt cọc: [●] đồng (bằng chữ: [●] đồng)."),
  p("Thời hạn đặt cọc: …… ngày kể từ ngày ký."),
  p("Hà Nội, ngày   tháng   năm 2026"),
  p("PHẦN 3. BẢNG THÔNG TIN", { b: true }),
  table([
    [run("Thông tin", { b: true }), run("Nội dung", { b: true })],
    [run("Diện tích (m²)"), run("")],
    [run("Giá bán"), run("")],
    [run("Mã căn"), run("")],
  ]),
  p("ĐẠI DIỆN BÊN A                                   ĐẠI DIỆN BÊN B"),
  p("(Ký, ghi rõ họ tên)                               (Ký, ghi rõ họ tên)"),
].join("");

await buildDocx(out, { body });
console.log(`Đã tạo ${out}: 9 ô điền mẫu (6 loại, 1 ô khóa nội dung), 6 dòng chỗ trống gõ tay, 1 bảng 3 dòng trống`);
