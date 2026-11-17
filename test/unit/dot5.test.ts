/** Đợt 5, phần A: đọc tệp PDF đính kèm theo trang — chọn trang, phân loại, đường đi của PDF quét, nhắc riêng tư, công cụ read_attachment_pages */
import * as assert from "assert";
import {
  MAX_SCAN_PAGES_PER_CALL,
  SCAN_RULES,
  WHOLE_PDF_MAX_PAGES,
  NO_VISION_NOTE,
  attachmentPrivacyNote,
  canSeePages,
  formatPages,
  isScanPage,
  parsePageSpec,
  planPages,
  routeScanPdf,
} from "../../src/core/pdfPages";
import { Attachment, PDF_MIME, withAttachments } from "../../src/core/attachments";
import { clearAttachedFiles, findAttachedFile, listAttachedFiles, registerAttachedFile } from "../../src/core/attachedFiles";
import { PdfReader, createCommonTools } from "../../src/office/common";

describe("Đợt 5 — chọn trang và phân loại trang", () => {
  it("đọc được khoảng trang, danh sách, bỏ trùng, cắt theo số trang của tệp", () => {
    assert.deepStrictEqual(parsePageSpec("1-3, 5", 10).pages, [1, 2, 3, 5]);
    assert.deepStrictEqual(parsePageSpec("5,1-2,2", 10).pages, [1, 2, 5]);
    assert.deepStrictEqual(parsePageSpec("8-20", 10).pages, [8, 9, 10]);
    assert.deepStrictEqual(parsePageSpec("3", 3).pages, [3]);
    assert.deepStrictEqual(parsePageSpec("2–4", 9).pages, [2, 3, 4], "gạch nối dài cũng nhận");
  });
  it("lời gọi sai thì trả câu lỗi, không trả trang", () => {
    for (const bad of ["", "abc", "0", "4-2", "1..3"]) {
      const r = parsePageSpec(bad, 10);
      assert.ok(r.error, `"${bad}" phải báo lỗi`);
      assert.deepStrictEqual(r.pages, []);
    }
    assert.match(parsePageSpec("12", 10).error || "", /chỉ có 10 trang/);
  });
  it("ghi dãy trang gọn lại", () => {
    assert.strictEqual(formatPages([1, 2, 3, 5]), "1-3, 5");
    assert.strictEqual(formatPages([7, 4]), "4, 7");
    assert.strictEqual(formatPages([]), "");
  });
  it("trang dưới 300 ký tự lớp chữ là trang quét (dòng chữ ký số không làm trang quét thành trang chữ)", () => {
    assert.strictEqual(isScanPage(0), true);
    assert.strictEqual(isScanPage(120), true);
    assert.strictEqual(isScanPage(300), false);
    assert.strictEqual(isScanPage(2500), false);
  });
  it("mỗi lần đưa tối đa 4 trang quét, phần còn lại để lần gọi sau, không nhảy cóc", () => {
    const allScan = planPages([1, 2, 3, 4, 5, 6], () => true);
    assert.deepStrictEqual(allScan.now, [1, 2, 3, 4]);
    assert.deepStrictEqual(allScan.later, [5, 6]);
    // Trang 1-2 chữ, 3-7 quét, 8 chữ: dừng ở trang quét thứ 5, không đưa trang 8 trước trang 7
    const mixed = planPages([1, 2, 3, 4, 5, 6, 7, 8], (n) => n >= 3 && n <= 7);
    assert.deepStrictEqual(mixed.now, [1, 2, 3, 4, 5, 6]);
    assert.deepStrictEqual(mixed.later, [7, 8]);
    assert.strictEqual(MAX_SCAN_PAGES_PER_CALL, 4, "vòng lặp tác tử chỉ chuyển 4 ảnh một lần gọi");
  });
});

describe("Đợt 5 — đường đi của PDF quét lúc gửi", () => {
  const base = { pdfCapable: true, vision: true, tools: true };
  it("tệp ngắn, nguồn đọc được nguyên tệp: gửi nguyên tệp như trước", () => {
    assert.strictEqual(routeScanPdf({ ...base, pages: 3 }), "whole");
    assert.strictEqual(routeScanPdf({ ...base, pages: WHOLE_PDF_MAX_PAGES }), "whole");
  });
  it("tệp dài: AI đọc theo trang bằng công cụ", () => {
    assert.strictEqual(routeScanPdf({ ...base, pages: WHOLE_PDF_MAX_PAGES + 1 }), "pages");
    assert.strictEqual(routeScanPdf({ ...base, pages: 80 }), "pages");
  });
  it("nguồn không nhận PDF nhưng model nhận ảnh và có công cụ: đọc theo trang, kể cả tệp ngắn", () => {
    assert.strictEqual(routeScanPdf({ pages: 2, pdfCapable: false, vision: true, tools: true }), "pages");
  });
  it("không có công cụ: nguồn đọc được thì gửi nguyên tệp dù dài; nguồn không đọc được thì chịu", () => {
    assert.strictEqual(routeScanPdf({ pages: 40, pdfCapable: true, vision: true, tools: false }), "whole");
    assert.strictEqual(routeScanPdf({ pages: 2, pdfCapable: false, vision: true, tools: false }), "none");
  });
  it("nguồn không nhận PDF mà danh mục không ghi rõ model nhận ảnh (Cerebras…): không gửi ảnh trang, kẻo bị từ chối cả lượt", () => {
    assert.strictEqual(routeScanPdf({ pages: 7, pdfCapable: false, vision: undefined, tools: true }), "none");
    assert.strictEqual(canSeePages({ pdfCapable: false, vision: undefined }), false);
    assert.strictEqual(canSeePages({ pdfCapable: false, vision: true }), true);
    assert.strictEqual(canSeePages({ pdfCapable: true, vision: undefined }), true, "nguồn đọc được PDF: model không rõ vẫn coi là xem được, như trước");
    assert.strictEqual(canSeePages({ pdfCapable: true, vision: false }), false);
  });
  it("model không nhận ảnh: không đọc được, dù có công cụ", () => {
    assert.strictEqual(routeScanPdf({ pages: 2, pdfCapable: true, vision: false, tools: true }), "none");
    assert.strictEqual(routeScanPdf({ pages: 2, pdfCapable: false, vision: false, tools: true }), "none");
  });
});

describe("Đợt 5 — nhắc riêng tư khi gửi tệp đính kèm", () => {
  it("OpenRouter: model miễn phí nhắc về huấn luyện, model trả tiền nhắc về nơi trung chuyển", () => {
    assert.match(attachmentPrivacyNote({ providerId: "openrouter", providerName: "OpenRouter", modelId: "x/y:free" }), /miễn phí/);
    const paid = attachmentPrivacyNote({ providerId: "openrouter", providerName: "OpenRouter", modelId: "anthropic/claude-sonnet-5.5" });
    assert.match(paid, /trung chuyển/);
    assert.doesNotMatch(paid, /miễn phí/);
  });
  it("nguồn gửi thẳng, model trả tiền: không nhắc (không báo thừa)", () => {
    assert.strictEqual(attachmentPrivacyNote({ providerId: "anthropic", providerName: "Anthropic", modelId: "claude-sonnet-5-5" }), "");
    assert.strictEqual(attachmentPrivacyNote({ providerId: "gemini", providerName: "Google Gemini", modelId: "gemini-3.8-flash" }), "");
  });
  it("model được danh mục đánh dấu miễn phí: nhắc", () => {
    assert.match(attachmentPrivacyNote({ providerId: "khac", providerName: "Nguồn X", modelId: "m", free: true }), /Nguồn X.*miễn phí/);
  });
});

describe("Đợt 5 — ghép tệp vào câu hỏi", () => {
  const scan: Attachment = { name: "quet.pdf", kind: "file", mime: PDF_MIME, data: "AAAA", pages: 12, note: "PDF quét, 12 trang, gửi nguyên tệp" };
  it("PDF quét đi đường công cụ: nói rõ nội dung CHƯA gửi và cách đọc; không nói 'gửi kèm tin nhắn này'", () => {
    const t = withAttachments("Hỏi", [scan], { byPages: ["quet.pdf"] });
    assert.match(t, /CHƯA được gửi kèm/);
    assert.match(t, /read_attachment_pages với file = "quet\.pdf"/);
    assert.doesNotMatch(t, /gửi kèm tin nhắn này để bạn đọc trực tiếp/);
  });
  it("PDF quét gửi nguyên tệp: giữ câu cũ", () => {
    const t = withAttachments("Hỏi", [scan]);
    assert.match(t, /gửi kèm tin nhắn này để bạn đọc trực tiếp/);
    assert.doesNotMatch(t, /read_attachment_pages/);
  });
  it("PDF lẫn trang quét: chỉ nhắc công cụ khi lượt này có công cụ và tệp còn giữ bản gốc", () => {
    const mixed: Attachment = { name: "lan.pdf", kind: "text", mime: PDF_MIME, data: "AAAA", pages: 9, text: "[Trang 1]\nchữ", note: "2 trang không có lớp chữ", scanPages: [3, 7] };
    assert.match(withAttachments("Hỏi", [mixed], { pageTool: true }), /read_attachment_pages với file = "lan\.pdf"/);
    assert.doesNotMatch(withAttachments("Hỏi", [mixed]), /read_attachment_pages/);
    assert.doesNotMatch(withAttachments("Hỏi", [{ ...mixed, data: undefined }], { pageTool: true }), /read_attachment_pages/);
  });
});

describe("Đợt 5 — công cụ read_attachment_pages", () => {
  /** Tệp giả 10 trang: trang 1-2 có chữ, 3-9 quét, 10 có chữ */
  const scanPage = (n: number) => n >= 3 && n <= 9;
  const rendered: number[] = [];
  const pdf: PdfReader = {
    pdfPageCount: async () => 10,
    readPdfPages: async (_n, _d, pages) => pages.map((p) => ({ page: p, chars: scanPage(p) ? 0 : 900, scan: scanPage(p), text: scanPage(p) ? "" : `Chữ trang ${p}` })),
    renderPdfPage: async (_n, _d, p) => {
      rendered.push(p);
      return { mime: "image/jpeg", data: `IMG${p}` };
    },
  };
  const tools = () => createCommonTools({ host: "Word", skills: [], shareContext: false, pdf });
  beforeEach(() => {
    clearAttachedFiles();
    rendered.length = 0;
  });
  after(() => clearAttachedFiles());

  it("công cụ luôn có trong bộ công cụ chung (bộ công cụ không đổi theo lượt)", () => {
    assert.ok(tools().specs.some((s) => s.name === "read_attachment_pages"));
    assert.ok(createCommonTools({ host: "Excel", skills: [], shareContext: false }).specs.some((s) => s.name === "read_attachment_pages"));
  });
  it("chưa đính kèm PDF: báo rõ, không bịa", async () => {
    await assert.rejects(() => tools().execute("read_attachment_pages", { pages: "1" }), /Chưa có tệp PDF nào đính kèm/);
    registerAttachedFile("a.docx", "X");
    await assert.rejects(() => tools().execute("read_attachment_pages", { pages: "1" }), /Chưa có tệp PDF nào đính kèm/, "tệp .docx không bị coi là PDF");
  });
  it("trang có chữ ra chữ; trang quét ra ảnh kèm lời dặn; ghi rõ trang đã đưa và trang còn lại", async () => {
    registerAttachedFile("hd.pdf", "DATA");
    const r = await tools().execute("read_attachment_pages", { pages: "1-10" });
    assert.strictEqual(r.total_pages, 10);
    assert.strictEqual(r.pages_returned, "1-6");
    assert.strictEqual(r.pages_not_returned, "7-10");
    assert.match(r.next, /pages = "7-10"/);
    assert.deepStrictEqual(r.pages.map((p: any) => p.kind), ["text", "text", "scan", "scan", "scan", "scan"]);
    assert.strictEqual(r.pages[0].text, "Chữ trang 1");
    assert.deepStrictEqual(rendered, [3, 4, 5, 6], "chỉ dựng ảnh các trang quét được đưa");
    assert.strictEqual(r.__images.length, 4);
    assert.match(r.__images[0].label, /trang 3 của hd\.pdf/);
    assert.strictEqual(r.rules, SCAN_RULES);
    assert.match(r.rules, /\[\?\]/);
    assert.match(r.rules, /số trang/);
  });
  it("model không chắc xem được ảnh: trang quét không dựng, không gửi; kết quả dặn không đoán; trang chữ vẫn đọc", async () => {
    registerAttachedFile("hd.pdf", "DATA");
    const t = createCommonTools({ host: "Word", skills: [], shareContext: false, pdf, seesPages: false });
    const r = await t.execute("read_attachment_pages", { pages: "2-4" });
    assert.deepStrictEqual(rendered, []);
    assert.strictEqual(r.__images, undefined);
    assert.strictEqual(r.image_note, NO_VISION_NOTE);
    assert.match(r.image_note, /Không được đoán/);
    assert.strictEqual(r.pages[0].text, "Chữ trang 2");
    assert.strictEqual(r.pages[1].kind, "scan");
  });
  it("chỉ trang có chữ: không ảnh, không lời dặn về ảnh", async () => {
    registerAttachedFile("hd.pdf", "DATA");
    const r = await tools().execute("read_attachment_pages", { file: "hd.pdf", pages: "1-2, 10" });
    assert.strictEqual(r.pages_returned, "1-2, 10");
    assert.strictEqual(r.__images, undefined);
    assert.strictEqual(r.rules, undefined);
    assert.strictEqual(r.pages_not_returned, undefined);
  });
  it("trang ngoài tệp, tên tệp sai: câu lỗi rõ", async () => {
    registerAttachedFile("hd.pdf", "DATA");
    await assert.rejects(() => tools().execute("read_attachment_pages", { pages: "11" }), /chỉ có 10 trang/);
    await assert.rejects(() => tools().execute("read_attachment_pages", { file: "khac.pdf", pages: "1" }), /Đang có: hd\.pdf/);
  });
});

describe("Đợt 5 — kho tệp đính kèm lọc theo loại", () => {
  beforeEach(() => clearAttachedFiles());
  after(() => clearAttachedFiles());
  it("so sánh phiên bản chỉ thấy .docx, đọc theo trang chỉ thấy .pdf", () => {
    registerAttachedFile("v1.docx", "A");
    registerAttachedFile("quet.PDF", "B");
    assert.strictEqual(findAttachedFile(undefined, ".docx")?.data, "A", "tệp gần nhất là PDF nhưng công cụ so sánh vẫn lấy đúng .docx");
    assert.strictEqual(findAttachedFile(undefined, ".pdf")?.data, "B");
    assert.deepStrictEqual(listAttachedFiles(".docx"), ["v1.docx"]);
    assert.deepStrictEqual(listAttachedFiles(".pdf"), ["quet.PDF"]);
    assert.strictEqual(findAttachedFile("quet", ".pdf")?.data, "B");
  });
});

import { PENDING_NOTE, packParagraphs } from "../../src/office/util";
import { REASON_TOOLS, WORD_SPECS, reasonOf } from "../../src/office/wordTools";
import { describeCall } from "../../src/office/tools";

describe("Đợt 5 — đoạn đang có sửa đổi chờ duyệt: đưa cả chữ gốc và chữ sau sửa", () => {
  const rows = [
    { text: "Điều 6. Thanh toán", style: "Heading 1" },
    { text: "Lãi chậm trả tính sau 90 ngày.", style: "Normal", changed: true, before: "Lãi chậm trả tính sau 45 ngày.", changed_by: "Đối tác" },
    { text: "Đoạn chỉ đổi định dạng.", style: "Normal", changed: true, before: "Đoạn chỉ đổi định dạng." },
    { text: "Đoạn không đổi.", style: "Normal" },
  ];
  it("đoạn có sửa đổi kèm before và changed_by; kết quả kèm câu nhắc nêu cả hai bản", () => {
    const r: any = packParagraphs(rows, 0, 4);
    assert.strictEqual(r.paragraphs[1].before, "Lãi chậm trả tính sau 45 ngày.");
    assert.strictEqual(r.paragraphs[1].changed_by, "Đối tác");
    assert.strictEqual(r.pending_note, PENDING_NOTE);
    assert.match(r.pending_note, /chờ duyệt/);
  });
  it("chữ gốc trùng chữ hiện tại (chỉ đổi định dạng) thì không lặp before; đoạn không đổi không có gì thêm", () => {
    const r: any = packParagraphs(rows, 0, 4);
    assert.strictEqual(r.paragraphs[2].changed, true);
    assert.strictEqual("before" in r.paragraphs[2], false);
    assert.strictEqual("before" in r.paragraphs[3], false);
    assert.strictEqual("changed" in r.paragraphs[3], false);
  });
  it("không đoạn nào có chữ gốc khác thì không kèm câu nhắc (không báo thừa)", () => {
    const r: any = packParagraphs([rows[0], rows[2], rows[3]], 0, 3);
    assert.strictEqual(r.pending_note, undefined);
  });
});

describe("Đợt 5 — lý do kèm mỗi chỗ sửa trong Word", () => {
  it("các công cụ sửa nội dung Word có tham số reason; công cụ đọc thì không", () => {
    for (const n of REASON_TOOLS) {
      const sp = WORD_SPECS.find((s) => s.name === n);
      assert.ok(sp, `thiếu công cụ ${n}`);
      assert.strictEqual((sp!.parameters as any).properties.reason?.type, "string", n);
      assert.ok(!((sp!.parameters as any).required || []).includes("reason"), "reason không bắt buộc (model cũ không nêu vẫn sửa được)");
    }
    assert.strictEqual((WORD_SPECS.find((s) => s.name === "word_read_paragraphs")!.parameters as any).properties.reason, undefined);
  });
  it("nhãn bước kèm lý do; không nêu lý do thì nhãn như cũ", () => {
    const base = describeCall({ id: "1", name: "word_set_paragraph_text", args: { index: 7 } });
    assert.strictEqual(base, "Sửa [¶7]");
    assert.strictEqual(describeCall({ id: "1", name: "word_set_paragraph_text", args: { index: 7, reason: "  Thống nhất\nthuật ngữ Bên A " } }), "Sửa [¶7] — lý do: Thống nhất thuật ngữ Bên A");
  });
  it("lý do dài bị cắt một dòng; công cụ không thuộc nhóm sửa Word thì bỏ qua reason", () => {
    assert.ok(reasonOf("word_replace_text", { reason: "x".repeat(400) }).length <= 160);
    assert.strictEqual(reasonOf("excel_write_range", { reason: "abc" }), "");
    assert.strictEqual(describeCall({ id: "1", name: "word_read_paragraphs", args: { reason: "abc" } }).includes("lý do"), false);
  });
});
