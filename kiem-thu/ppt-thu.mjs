// Dựng bài PowerPoint thử cho Đợt 3: 10 slide theo một mẫu có phông chủ đề và bố cục riêng, tự mở khung add-in,
// cài sẵn 5 lỗi lệch mẫu (2 phông lạ, 1 slide thiếu tiêu đề, 2 chỗ chữ dưới 12 pt) và 1 slide chữ cố ý tràn khung.
//   node kiem-thu/ppt-thu.mjs <tệp ra .pptx>          bài cài lỗi
//   node kiem-thu/ppt-thu.mjs <tệp ra .pptx> sach     bài sạch (không cài lỗi, không slide tràn)
// Lỗi cài: slide 3 hộp "Ghi chú tiến độ" phông Comic Sans MS · slide 5 hộp "Nguồn số liệu" phông Courier New · slide 7 không có tiêu đề ·
// slide 8 hộp "Chú thích" cỡ 9 pt · slide 10 hộp "Liên hệ" cỡ 10 pt. Slide 6: hộp "Rủi ro" chữ dài trong khung cố định (tràn).
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const [, , out, mode] = process.argv;
if (!out) {
  console.error("Thiếu tệp ra");
  process.exit(1);
}
const sach = mode === "sach";
const PptxGenJS = (await import("pptxgenjs")).default;
const JSZip = (await import("jszip")).default;

const pptx = new PptxGenJS();
pptx.defineLayout({ name: "AIDN169", width: 10, height: 5.625 });
pptx.layout = "AIDN169";
pptx.theme = { headFontFace: "Georgia", bodyFontFace: "Verdana" };
const NAVY = "1F3864";
const TEAL = "2E75B6";
pptx.defineSlideMaster({
  title: "Tiêu đề và nội dung",
  background: { color: "FFFFFF" },
  objects: [
    { rect: { x: 0, y: 0, w: 10, h: 0.18, fill: { color: NAVY } } },
    { rect: { x: 0, y: 5.45, w: 10, h: 0.175, fill: { color: TEAL } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.35, w: 9, h: 0.8, fontFace: "Georgia", fontSize: 28, bold: true, color: NAVY }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1.3, w: 9, h: 3.2, fontFace: "Verdana", fontSize: 18, color: "262626" }, text: "" } },
  ],
});
pptx.defineSlideMaster({
  title: "Chỉ nội dung",
  background: { color: "FFFFFF" },
  objects: [
    { rect: { x: 0, y: 0, w: 10, h: 0.18, fill: { color: NAVY } } },
    { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 0.6, w: 9, h: 4.2, fontFace: "Verdana", fontSize: 18, color: "262626" }, text: "" } },
  ],
});

const DATA = [
  ["Tổng quan dự án Khu nhà ở Sông Hồng", ["Quy mô 12,4 ha, 3 khối nhà, 1.150 căn", "Tổng mức đầu tư 3.200 tỷ đồng", "Khởi công tháng 3/2026"]],
  ["Pháp lý dự án", ["Chấp thuận chủ trương đầu tư: đã có", "Quy hoạch chi tiết 1/500: đã phê duyệt", "Giấy phép xây dựng khối A: đã cấp"]],
  ["Tiến độ thi công", ["San nền: hoàn thành", "Hạ tầng kỹ thuật: 62% (kế hoạch 80%)", "Móng khối A: 55% (kế hoạch 60%)"]],
  ["Tiến độ giải ngân", ["Đã giải ngân 1.120 tỷ đồng", "Kế hoạch quý IV: 380 tỷ đồng", "Nguồn: vốn chủ sở hữu và tín dụng"]],
  ["Kinh doanh", ["Đã ký 412 hợp đồng đặt cọc", "Giá bán bình quân 48 triệu đồng/m²", "Dự kiến mở bán đợt 2 tháng 1/2027"]],
  ["Rủi ro chính", ["Chậm bàn giao mặt bằng đợt 2", "Thiếu nhân công thi công phần thân", "Biến động giá vật liệu"]],
  ["", ["Slide này cố ý không có tiêu đề", "Nội dung về kế hoạch quý tới", "Ba mốc cần phê duyệt"]],
  ["Kế hoạch quý IV/2026", ["Hoàn thành hạ tầng kỹ thuật giai đoạn 1", "Cất nóc khối A", "Nộp hồ sơ nghiệm thu phòng cháy"]],
  ["Kiến nghị", ["Gia hạn tiến độ hạ tầng 06 tháng", "Bổ sung nhà thầu phụ phần thân", "Phê duyệt chi phí phát sinh 4,2 tỷ đồng"]],
  ["Liên hệ và bước tiếp theo", ["Họp Hội đồng quản trị ngày 28/11", "Đầu mối: Ban Pháp chế", "Tài liệu kèm: báo cáo tiến độ tháng 11"]],
];
DATA.forEach(([title, bullets], i) => {
  const n = i + 1;
  const noTitle = !sach && n === 7;
  const s = pptx.addSlide({ masterName: noTitle ? "Chỉ nội dung" : "Tiêu đề và nội dung" });
  if (!noTitle) s.addText(title || `Kế hoạch quý tới`, { placeholder: "title" });
  s.addText(
    bullets.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })),
    { placeholder: "body" }
  );
  const box = (name, text, o) => s.addText(text, { x: 5.6, y: 4.55, w: 3.9, h: 0.6, fontFace: "Verdana", fontSize: 14, color: "404040", objectName: name, ...o });
  if (sach) return;
  if (n === 3) box("Ghi chú tiến độ", "Số liệu tính đến ngày 15/11/2026", { fontFace: "Comic Sans MS" });
  if (n === 5) box("Nguồn số liệu", "Nguồn: phòng Kinh doanh", { fontFace: "Courier New" });
  if (n === 8) box("Chú thích", "Các mốc có thể thay đổi theo thời tiết", { fontSize: 9 });
  if (n === 10) box("Liên hệ", "Điện thoại đầu mối: xem danh bạ nội bộ", { fontSize: 10 });
  if (n === 6)
    s.addText("Chi tiết rủi ro: việc bàn giao mặt bằng đợt 2 chậm so với cam kết kéo theo chậm hạ tầng kỹ thuật; nhà thầu thiếu nhân công ở phần thân khối A; giá thép và xi măng tăng làm đội chi phí; thủ tục điều chỉnh giấy phép kéo dài hơn dự kiến; cần phương án dự phòng cho từng rủi ro và người chịu trách nhiệm theo dõi hằng tuần.", {
      x: 5.6,
      y: 3.6,
      w: 3.9,
      h: 0.7,
      fontFace: "Verdana",
      fontSize: 16,
      color: "262626",
      objectName: "Rủi ro",
      fit: "none",
      valign: "top",
    });
});

const buf = await pptx.write({ outputType: "nodebuffer" });
// Gắn phần "tự mở khung add-in" lấy từ mẫu của office-addin-dev-settings (cùng cách với cdp.mjs make)
const zip = await JSZip.loadAsync(buf);
const tpl = await JSZip.loadAsync(fs.readFileSync(path.join(root, "node_modules/office-addin-dev-settings/templates/PowerPointPresentationWithTaskPane.pptx")));
const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
const id = /<Id>([^<]+)<\/Id>/.exec(manifest)[1];
const version = /<Version>([^<]+)<\/Version>/.exec(manifest)[1];
for (const f of ["ppt/webextensions/taskpanes.xml", "ppt/webextensions/_rels/taskpanes.xml.rels"]) zip.file(f, await tpl.file(f).async("nodebuffer"));
const ext = (await tpl.file("ppt/webextensions/webextension.xml").async("string")).replace(/00000000-0000-0000-0000-000000000000/g, id).replace(/1\.0\.0\.0/g, version);
zip.file("ppt/webextensions/webextension.xml", ext);
let ct = await zip.file("[Content_Types].xml").async("string");
ct = ct.replace(
  "</Types>",
  '<Override PartName="/ppt/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/><Override PartName="/ppt/webextensions/webextension.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>'
);
zip.file("[Content_Types].xml", ct);
let rels = await zip.file("_rels/.rels").async("string");
rels = rels.replace("</Relationships>", '<Relationship Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="/ppt/webextensions/taskpanes.xml" Id="rAidnTaskpanes"/></Relationships>');
zip.file("_rels/.rels", rels);
fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }));
console.log(`Đã tạo ${out}: 10 slide${sach ? " (bài sạch)" : ", cài 5 lỗi lệch mẫu (slide 3, 5, 7, 8, 10) và 1 slide chữ tràn (slide 6)"}; add-in ${id}`);
