/** Mức 3 tối ưu theo bản PowerPoint chính thức (27/11/2026): việc 10 storyboard + văn phong, việc 11 biểu đồ theo đặc tả */
import * as assert from "assert";
import JSZip from "jszip";
import { lintFontSize, lintSlideText, lintTitles, styleWarningsFor, withStyleWarnings } from "../../src/office/pptStyle";
import { buildChartPptx, chartSpec, MUTED_HEX } from "../../src/office/pptAdvanced";

const kinds = (ws: { kind: string }[]) => ws.map((w) => w.kind).sort();

describe("Mức 3 · việc 10 — soát văn phong chữ trên slide", () => {
  it("bắt biểu tượng cảm xúc, khuôn 'không phải… mà là…', lời trao đổi, gạch ngang dày, từ nhấn rỗng", () => {
    assert.deepStrictEqual(kinds(lintSlideText("Doanh thu tăng 🚀")), ["emoji"]);
    assert.deepStrictEqual(kinds(lintSlideText("Đây không phải chi phí, mà là khoản đầu tư")), ["contrast_frame"]);
    assert.deepStrictEqual(kinds(lintSlideText("Bản rút gọn theo góp ý của anh")), ["meta"]);
    assert.deepStrictEqual(kinds(lintSlideText("A — B — C — D")), ["dashes"]);
    assert.deepStrictEqual(kinds(lintSlideText("AI đã rút gọn phần này")), ["meta"]);
    assert.deepStrictEqual(kinds(lintSlideText("Giải pháp đột phá cho dự án")), ["hype"]);
  });

  it("chữ bình thường của slide pháp lý không bị bắt nhầm", () => {
    for (const t of [
      "Tiến độ pháp lý: đã có chủ trương đầu tư, chờ giao đất",
      "Không phải nộp tiền sử dụng đất bổ sung", // "không phải" nhưng không có "mà là"
      "Thi công móng — đạt 60%",
      "Vướng mắc: chưa có quyết định giao đất; kế hoạch tháng 12",
      // Phản biện mức 3: đại từ "ai", ® ™, khoảng năm / quý, "dưới đây là" trong thân bài
      "Người ai đã ký hợp đồng phải chịu trách nhiệm; ai vừa nộp hồ sơ thì chờ",
      "Vinhomes® Grand Park™ — giai đoạn 2",
      "Lộ trình 2020–2025: Q1–Q3 hoàn tất, 10–12 tầng",
      "Các điều kiện dưới đây là bắt buộc",
    ])
      assert.deepStrictEqual(lintSlideText(t), [], t);
  });

  it("cỡ chữ: thân dưới 14 pt cảnh báo; chú thích nguồn được tới 10 pt", () => {
    assert.deepStrictEqual(kinds(lintFontSize(12, "Doanh thu quý 3")), ["small_font"]);
    assert.deepStrictEqual(lintFontSize(14, "Doanh thu quý 3"), []);
    assert.deepStrictEqual(lintFontSize(10, "Nguồn: Báo cáo tài chính 2026"), []);
    assert.deepStrictEqual(kinds(lintFontSize(8, "Nguồn: BCTC")), ["small_font"]);
    assert.deepStrictEqual(lintFontSize(undefined, "x"), []);
  });

  it("tiêu đề: quá 15 từ, chấm cuối, trùng nhau", () => {
    const long = "Một hai ba bốn năm sáu bảy tám chín mười mười một mười hai mười ba mười bốn mười lăm mười sáu";
    assert.deepStrictEqual(kinds(lintTitles([long, "Doanh thu tăng 32%.", "Kết luận", "kết luận"])), ["title_dup", "title_long", "title_period"]);
    assert.deepStrictEqual(lintTitles(["Doanh thu tăng 32% năm 2026", "Chi phí giảm 5%", ""]), []);
  });

  it("ppt_add_slides từ 3 slide thiếu key_message thì nhắc; có thì không", () => {
    const slides = [{ title: "A tăng 5%" }, { title: "B giảm 2%" }, { title: "C giữ 10%" }];
    assert.ok(kinds(styleWarningsFor("ppt_add_slides", { slides })).includes("no_key_message"));
    assert.ok(!kinds(styleWarningsFor("ppt_add_slides", { slides, key_message: "Dự án đúng tiến độ" })).includes("no_key_message"));
    assert.ok(!kinds(styleWarningsFor("ppt_add_slides", { slides: slides.slice(0, 2) })).includes("no_key_message"));
  });

  it("gắn style_warnings vào kết quả; không có cảnh báo thì kết quả giữ nguyên; tối đa 8 dòng", () => {
    const r = { ok: true, message: "x" };
    assert.strictEqual(withStyleWarnings("ppt_add_textbox", { text: "Doanh thu 5 tỷ", font_size: 16 }, r), r);
    const w = withStyleWarnings("ppt_add_textbox", { text: "Tuyệt vời 🎉", font_size: 9 }, r) as any;
    assert.strictEqual(w.ok, true);
    assert.strictEqual(w.style_warnings.length, 3);
    const many = { slides: Array.from({ length: 10 }, () => ({ title: "Trùng.", body: "🚀 đột phá" })) };
    assert.ok((withStyleWarnings("ppt_add_slides", many, r) as any).style_warnings.length <= 8);
    // Lệnh đọc không bị gắn
    assert.strictEqual(withStyleWarnings("ppt_get_overview", { text: "🚀" }, r), r);
    // Kết quả lỗi, mảng: không gắn
    const bad = { ok: false, message: "lỗi" };
    assert.strictEqual(withStyleWarnings("ppt_add_textbox", { text: "🚀" }, bad), bad);
    const arr: any = [1];
    assert.strictEqual(withStyleWarnings("ppt_add_textbox", { text: "🚀" }, arr), arr);
  });
});

describe("Mức 3 · việc 11 — biểu đồ theo đặc tả", () => {
  const base = { chart_type: "column", categories: ["Q1", "Q2", "Q3"], series: [{ name: "Doanh thu", values: [10, 30, 20] }], title: "Q2 cao nhất: 30 tỷ" };

  it("điểm nhấn: cột nhấn màu Accent1 (hex), còn lại xám; không phân biệt hoa thường", () => {
    const s = chartSpec({ ...base, highlight: ["q2"] }, "#1F6FB2");
    assert.deepStrictEqual(s.pointColors, [MUTED_HEX, "1F6FB2", MUTED_HEX]);
    assert.deepStrictEqual(s.warnings, []);
  });

  it("điểm nhấn bị bỏ qua khi nhiều chuỗi, dạng đường, hoặc không trúng danh mục nào — và báo lý do", () => {
    const two = chartSpec({ ...base, series: [...base.series, { name: "Chi phí", values: [1, 2, 3] }], highlight: ["Q2"] });
    assert.strictEqual(two.pointColors, undefined);
    assert.match(two.warnings.join(" "), /một chuỗi/);
    assert.match(chartSpec({ ...base, chart_type: "line", highlight: ["Q2"] }).warnings.join(" "), /một chuỗi/);
    assert.match(chartSpec({ ...base, highlight: ["Q9"] }).warnings.join(" "), /không có danh mục/);
  });

  it("màu Accent1 hỏng thì dùng màu mặc định hợp lệ (pptxgenjs chỉ nhận hex 6 ký tự)", () => {
    assert.deepStrictEqual(chartSpec({ ...base, highlight: ["Q1"] }, "accent1").pointColors, ["4472C4", MUTED_HEX, MUTED_HEX]);
  });

  it("cảnh báo đặc tả: tròn quá 5 phần, tròn nhiều chuỗi, đơn vị trên biểu đồ tròn, tiêu đề không có số", () => {
    const pie = { chart_type: "pie", categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "x", values: [1, 2, 3, 4, 5, 6] }, { name: "y", values: [1] }], unit: "%" };
    const w = chartSpec(pie).warnings.join(" | ");
    assert.match(w, /6 phần/);
    assert.match(w, /chuỗi đầu/);
    assert.match(w, /không có trục/);
    assert.strictEqual(chartSpec(pie).unit, "");
    // Tiêu đề không có số: không cảnh báo (chỉ "có số liệu khi có")
    assert.deepStrictEqual(chartSpec({ ...base, title: "Doanh thu theo quý" }).warnings, []);
  });

  it("dựng tệp: có trục đơn vị, màu từng cột theo điểm nhấn, chú thích nguồn 10 pt", async () => {
    const b64 = await buildChartPptx({ ...base, unit: "tỷ đồng", source: "BCTC 2026", highlight: ["Q2"] }, { w: 960, h: 540 }, undefined, "1F6FB2");
    const zip = await JSZip.loadAsync(b64, { base64: true });
    const chartFile = Object.keys(zip.files).find((f) => /charts\/chart\d+\.xml$/.test(f))!;
    const chart = await zip.file(chartFile)!.async("string");
    assert.ok(chart.includes("tỷ đồng"), "thiếu tiêu đề trục đơn vị");
    assert.ok(chart.includes('val="1F6FB2"'), "thiếu màu nhấn");
    assert.ok(chart.includes(`val="${MUTED_HEX}"`), "thiếu màu xám");
    assert.ok(!/val="accent\d"/.test(chart.split("<c:dPt>").slice(1).join("")), "màu từng cột không được là tên màu chủ đề");
    const slide = await zip.file("ppt/slides/slide1.xml")!.async("string");
    assert.ok(slide.includes("Nguồn: BCTC 2026"), "thiếu chú thích nguồn");
    assert.ok(/sz="1000"/.test(slide), "chú thích nguồn phải 10 pt");
  });

  it("không đặc tả thêm thì tệp dựng như cũ (màu chủ đề, không trục đơn vị, không nguồn)", async () => {
    const b64 = await buildChartPptx(base, { w: 960, h: 540 });
    const zip = await JSZip.loadAsync(b64, { base64: true });
    const chartFile = Object.keys(zip.files).find((f) => /charts\/chart\d+\.xml$/.test(f))!;
    const chart = await zip.file(chartFile)!.async("string");
    assert.ok(!chart.includes(MUTED_HEX));
    const slide = await zip.file("ppt/slides/slide1.xml")!.async("string");
    assert.ok(!slide.includes("Nguồn:"));
  });
});
