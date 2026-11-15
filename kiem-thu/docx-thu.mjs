// Dựng tệp .docx thử cho Word từ mẫu "tài liệu tự mở khung add-in" (cùng mẫu với `cdp.mjs make word`),
// có kiểm soát từng chi tiết mà Office.js không tạo được: đánh số nhiều cấp gắn kiểu tiêu đề, trường đánh số,
// sửa đổi được theo dõi và nhận xét của một tác giả khác ("Đối tác").
//
//   node kiem-thu/docx-thu.mjs tham-do <tệp ra>      văn bản thăm dò Đợt 1A: 4 kiểu đánh số + đoạn có sửa đổi
//   node kiem-thu/docx-thu.mjs hop-dong <tệp ra>     hợp đồng thử ~50 trang, 3 cấp, số tự động lẫn số gõ tay, 30 sửa đổi + 5 nhận xét
//
// Tệp tạo ra chỉ chứa nội dung bịa để thử, không có thông tin thật.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");

export const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

let revId = 100;
const DATE = "2026-11-10T09:00:00Z";
/** Chữ thường; opts.b = đậm, opts.i = nghiêng */
export const run = (text, opts = {}) =>
  `<w:r>${opts.b || opts.i ? `<w:rPr>${opts.b ? "<w:b/>" : ""}${opts.i ? "<w:i/>" : ""}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
/** Chữ bị xóa có theo dõi */
export const del = (text, author = "Đối tác") =>
  `<w:del w:id="${revId++}" w:author="${esc(author)}" w:date="${DATE}"><w:r><w:delText xml:space="preserve">${esc(text)}</w:delText></w:r></w:del>`;
/** Chữ được thêm có theo dõi */
export const ins = (text, author = "Đối tác") =>
  `<w:ins w:id="${revId++}" w:author="${esc(author)}" w:date="${DATE}"><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:ins>`;
/** Trường (SEQ, LISTNUM, AUTONUM…) kèm kết quả đã tính sẵn */
export const field = (code, result = "") =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${esc(code)} </w:instrText></w:r>` +
  (result ? `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${esc(result)}</w:t></w:r>` : "") +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
/** Mốc nhận xét: bọc `inner` trong vùng neo của nhận xét số id */
export const commented = (id, inner) =>
  `<w:commentRangeStart w:id="${id}"/>${inner}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;

/** Một đoạn. content: chuỗi (chữ thường) hoặc XML các run. opts: style, num:[numId, ilvl], markDeleted, markInserted */
export function para(content, opts = {}) {
  const inner = /^</.test(content) ? content : content ? run(content) : "";
  const ppr = [
    opts.style ? `<w:pStyle w:val="${opts.style}"/>` : "",
    opts.num ? `<w:numPr><w:ilvl w:val="${opts.num[1]}"/><w:numId w:val="${opts.num[0]}"/></w:numPr>` : "",
    opts.markDeleted ? `<w:rPr><w:del w:id="${revId++}" w:author="${esc(opts.markDeleted)}" w:date="${DATE}"/></w:rPr>` : "",
    opts.markInserted ? `<w:rPr><w:ins w:id="${revId++}" w:author="${esc(opts.markInserted)}" w:date="${DATE}"/></w:rPr>` : "",
  ].join("");
  return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${inner}</w:p>`;
}

export function table(rows) {
  const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para(t)}</w:tc>`;
  const border = ["top", "left", "bottom", "right", "insideH", "insideV"].map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${border}</w:tblBorders></w:tblPr>${rows
    .map((r) => `<w:tr>${r.map(cell).join("")}</w:tr>`)
    .join("")}</w:tbl>`;
}

const heading = (id, name, level, size, numId) =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:spacing w:before="${level === 0 ? 240 : 120}" w:after="60"/>${
    numId ? `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>` : ""
  }<w:outlineLvl w:val="${level}"/></w:pPr><w:rPr>${level < 2 ? "<w:b/>" : ""}<w:sz w:val="${size}"/></w:rPr></w:style>`;

/** Kiểu thêm vào styles.xml. headingNumId: danh sách nhiều cấp gắn vào Heading 1–3 (0 = không gắn) */
const extraStyles = (headingNumId) =>
  heading("Heading1", "heading 1", 0, 28, headingNumId) +
  heading("Heading2", "heading 2", 1, 24, headingNumId) +
  heading("Heading3", "heading 3", 2, 24, headingNumId) +
  `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Khoan"><w:name w:val="Khoản hợp đồng"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80"/><w:jc w:val="both"/></w:pPr></w:style>`;

/** Một mức của danh sách. fmt: decimal | lowerLetter | lowerRoman; text: "Điều %1." … ; pStyle: kiểu gắn vào mức này */
const lvl = (i, fmt, text, pStyle, left = 720 * (i + 1)) =>
  `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>${pStyle ? `<w:pStyle w:val="${pStyle}"/>` : ""}<w:lvlText w:val="${esc(text)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${pStyle ? left : 360}"/></w:pPr></w:lvl>`;

/** numbering.xml: mỗi phần tử là mảng các mức; numId = thứ tự trong mảng + 1 */
function numberingXml(lists) {
  const abs = lists.map((levels, k) => `<w:abstractNum w:abstractNumId="${k}"><w:multiLevelType w:val="multilevel"/>${levels.join("")}</w:abstractNum>`).join("");
  const nums = lists.map((_, k) => `<w:num w:numId="${k + 1}"><w:abstractNumId w:val="${k}"/></w:num>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abs}${nums}</w:numbering>`;
}

function commentsXml(comments) {
  const items = comments
    .map(
      (c, id) =>
        `<w:comment w:id="${id}" w:author="${esc(c.author || "Đối tác")}" w:date="${DATE}" w:initials="DT"><w:p><w:r><w:t xml:space="preserve">${esc(c.text)}</w:t></w:r></w:p></w:comment>`
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${items}</w:comments>`;
}

/** Ghi tệp .docx. spec: { body: XML các đoạn, lists: mảng danh sách, headingNumId, comments: [{text, author}] } */
export async function buildDocx(out, spec) {
  const JSZip = (await import("jszip")).default;
  const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
  const id = /<Id>([^<]+)<\/Id>/.exec(manifest)[1];
  const version = /<Version>([^<]+)<\/Version>/.exec(manifest)[1];
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(root, "node_modules/office-addin-dev-settings/templates/WordDocumentWithTaskPane.docx")));
  const get = (f) => zip.file(f).async("string");
  const ext = "word/webextensions/webextension.xml";
  zip.file(ext, (await get(ext)).replace(/00000000-0000-0000-0000-000000000000/g, id).replace(/1\.0\.0\.0/g, version));

  const doc = await get("word/document.xml");
  const sect = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(doc)[0];
  zip.file("word/document.xml", doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${spec.body}${sect}</w:body>`));
  zip.file("word/styles.xml", (await get("word/styles.xml")).replace("</w:styles>", extraStyles(spec.headingNumId || 0) + "</w:styles>"));

  let rels = await get("word/_rels/document.xml.rels");
  let types = await get("[Content_Types].xml");
  const addPart = (file, relType, contentType, relId) => {
    rels = rels.replace("</Relationships>", `<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${relType}" Target="/word/${file}" Id="${relId}" /></Relationships>`);
    types = types.replace("</Types>", `<Override PartName="/word/${file}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${contentType}+xml" /></Types>`);
  };
  if (spec.lists?.length) {
    zip.file("word/numbering.xml", numberingXml(spec.lists));
    addPart("numbering.xml", "numbering", "numbering", "RaidnNum");
  }
  if (spec.comments?.length) {
    zip.file("word/comments.xml", commentsXml(spec.comments));
    addPart("comments.xml", "comments", "comments", "RaidnCmt");
  }
  zip.file("word/_rels/document.xml.rels", rels);
  zip.file("[Content_Types].xml", types);
  fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }));
  return out;
}

// ───────────────────────── Văn bản thăm dò Đợt 1A ─────────────────────────
export function thamDo() {
  const lists = [
    // 1: danh sách thường "1."
    [lvl(0, "decimal", "%1.")],
    // 2: nhiều cấp GẮN kiểu tiêu đề: "Điều 1." / "1.1." / "a)"
    [lvl(0, "decimal", "Điều %1.", "Heading1", 0), lvl(1, "decimal", "%1.%2.", "Heading2", 0), lvl(2, "lowerLetter", "%3)", "Heading3", 360)],
    // 3: nhiều cấp KHÔNG gắn kiểu (numPr nằm trên đoạn): "1." / "1.1" / "(a)"
    [lvl(0, "decimal", "%1."), lvl(1, "decimal", "%1.%2"), lvl(2, "lowerLetter", "(%3)")],
  ];
  const body = [
    para("VĂN BẢN THĂM DÒ ĐÁNH SỐ (AIDN)"),
    para("[A] Danh sách thường"),
    para("Mục thường thứ nhất", { style: "ListParagraph", num: [1, 0] }),
    para("Mục thường thứ hai", { style: "ListParagraph", num: [1, 0] }),
    para("Mục thường thứ ba", { style: "ListParagraph", num: [1, 0] }),
    para("[B] Nhiều cấp gắn kiểu tiêu đề"),
    para("Phạm vi hợp đồng", { style: "Heading1" }),
    para("Bên A giao cho Bên B thực hiện công việc nêu tại Phụ lục 01.", { style: "Heading2" }),
    para("Bên B cam kết hoàn thành đúng tiến độ.", { style: "Heading2" }),
    para("tiến độ tổng thể theo Phụ lục 02;", { style: "Heading3" }),
    para("tiến độ chi tiết do hai bên thống nhất.", { style: "Heading3" }),
    para("Giá và thanh toán", { style: "Heading1" }),
    para("Giá trị hợp đồng là 5.000.000.000 đồng.", { style: "Heading2" }),
    para("Thanh toán làm ba đợt.", { style: "Heading2" }),
    para("[B2] Nhiều cấp không gắn kiểu"),
    para("Quyền của Bên A", { num: [3, 0] }),
    para("Được kiểm tra tiến độ.", { num: [3, 1] }),
    para("Được yêu cầu sửa chữa sai sót.", { num: [3, 1] }),
    para("bằng văn bản;", { num: [3, 2] }),
    para("Nghĩa vụ của Bên A", { num: [3, 0] }),
    para("[C] Số gõ tay"),
    para("Điều 7. Thời hạn"),
    para("7.1 Thời hạn thuê là 24 tháng."),
    para("7.2. Gia hạn theo thỏa thuận."),
    para("a) bằng phụ lục hợp đồng;"),
    para("[D] Trường đánh số"),
    para(run("Điều ") + field("SEQ Dieu \\* ARABIC", "9") + run(". Bảo mật")),
    para(run("Điều ") + field("SEQ Dieu \\* ARABIC", "10") + run(". Bất khả kháng")),
    para(field("LISTNUM NumberDefault \\l 1") + run(" Khoản dùng trường LISTNUM")),
    para(field("AUTONUM") + run(" Khoản dùng trường AUTONUM")),
    para("[E] Đoạn có sửa đổi của đối tác"),
    para(run("Thời hạn thuê là ") + del("24 tháng") + ins("12 tháng") + run(" kể từ ngày ký.")),
    para(del("Đoạn này đối tác đã xóa toàn bộ."), { markDeleted: "Đối tác" }),
    para(ins("Đoạn này đối tác thêm mới."), { markInserted: "Đối tác" }),
    para(
      run("Bên B thanh toán trong 30 ngày; chậm quá ") +
        run("30 ngày", { b: true }) +
        run(" thì chịu phạt 30 triệu đồng và lãi 0,05% mỗi ngày trên số tiền chậm trả.")
    ),
    para("[F] Bảng"),
    table([
      ["Hạng mục", "Số tiền"],
      ["Đợt 1", "1.500.000.000"],
    ]),
    para("HẾT"),
  ].join("");
  return { body, lists, headingNumId: 2 };
}

// ───────────────────────── chạy từ dòng lệnh ─────────────────────────
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [, , kind, out] = process.argv;
  if (!out) {
    console.log("Cách dùng: node kiem-thu/docx-thu.mjs <tham-do|hop-dong> <tệp ra>");
    process.exit(1);
  }
  let spec;
  if (kind === "tham-do") spec = thamDo();
  else if (kind === "hop-dong") spec = (await import("./hop-dong-thu.mjs")).hopDong();
  else throw new Error("Loại phải là tham-do hoặc hop-dong");
  await buildDocx(out, spec);
  console.log(`Đã tạo ${out}`);
}
