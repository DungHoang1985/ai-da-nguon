/** Đợt 3 — phần B: PowerPoint (kiểm bài theo mẫu bằng mã, công cụ mới, ảnh slide gửi cho model). */
import * as assert from "assert";
import { DeckSlide, checkDeck, deckBrief, dominantFonts, estimateTextHeight } from "../../src/core/deckcheck";
import { PPT_DESIGN_SPECS, splitParagraphs } from "../../src/office/pptDesign";
import { pptTools } from "../../src/office/pptTools";
import { describeCall } from "../../src/office/tools";
import { runAgent } from "../../src/core/agent/loop";

describe("Đợt 3 · B3 · Kiểm bài trình chiếu theo mẫu bằng mã", () => {
  const P = (text: string, font = "Calibri", size: number | null = 18) => ({ text, font, size });
  const title = (t: string) => ({
    name: "Title 1",
    type: "Placeholder",
    isTitle: true,
    box: [40, 30, 640, 60] as [number, number, number, number],
    paragraphs: t ? [P(t, "Calibri Light", 32)] : [],
    autoSize: "AutoSizeTextToFitShape",
    wordWrap: true,
  });
  const box = (name: string, paras: any[], b: [number, number, number, number] = [40, 120, 300, 200], autoSize = "AutoSizeShapeToFitText") => ({ name, type: "TextBox", box: b, paragraphs: paras, autoSize, wordWrap: true });
  const O = { slideW: 720, slideH: 405, themeFonts: ["Calibri Light", "Calibri"] };

  it("Bài 10 slide cài 5 lỗi lệch mẫu: tìm đủ 5 (2 phông lạ, 1 slide thiếu tiêu đề, 2 chỗ chữ dưới 12 pt), không báo thêm", () => {
    const slides: DeckSlide[] = Array.from({ length: 10 }, (_, i) => ({ slide: i + 1, layout: "Title and Content", shapes: [title(`Tiêu đề ${i + 1}`), box("Nội dung", [P("Ý một"), P("Ý hai")])] }));
    slides[1].shapes.push(box("Hộp A", [P("Chữ phông lạ", "Comic Sans MS")], [380, 120, 300, 60]));
    slides[4].shapes.push(box("Hộp B", [P("Chữ thường"), P("Chữ phông lạ khác", "Courier New")], [380, 120, 300, 60]));
    slides[6].shapes[0] = title("");
    slides[7].shapes.push(box("Chú thích", [P("Nguồn: báo cáo quý", "Calibri", 9)], [380, 330, 300, 30]));
    slides[9].shapes.push(box("Ghi chú", [P("Ghi chú nhỏ", "Calibri", 11)], [380, 330, 300, 30]));
    const r = checkDeck(slides, O);
    assert.deepStrictEqual(r.counts, { font: 2, missing_title: 1, small_text: 2 });
    assert.deepStrictEqual(
      r.issues.map((i) => [i.kind, i.slide]),
      [
        ["font", 2],
        ["font", 5],
        ["missing_title", 7],
        ["small_text", 8],
        ["small_text", 10],
      ]
    );
    assert.match(r.issues[0].message, /Comic Sans MS.*mẫu dùng Calibri Light, Calibri/);
    assert.match(deckBrief(r), /5 điểm cần xem: 2 phông lệch mẫu, 1 slide thiếu tiêu đề, 2 chữ quá nhỏ/);
    const clean = checkDeck(
      slides.filter((s) => ![2, 5, 7, 8, 10].includes(s.slide)),
      O
    );
    assert.strictEqual(clean.issues.length, 0);
    assert.match(deckBrief(clean), /không thấy điểm lệch mẫu/);
  });

  it("Chữ tràn khung (ước lượng), hình ra ngoài slide, khung chữ chồng nhau, chữ mẫu còn sót", () => {
    const long = "Chữ rất dài cố ý để tràn khỏi khung. ".repeat(10);
    const s: DeckSlide = {
      slide: 1,
      shapes: [
        title("Có tiêu đề"),
        box("Tràn", [P(long, "Calibri", 18)], [40, 120, 220, 50], "AutoSizeNone"),
        box("Tự giãn", [P(long, "Calibri", 18)], [300, 120, 220, 50], "AutoSizeShapeToFitText"),
        box("Ngoài", [P("x")], [650, 380, 200, 60]),
        box("Mẫu", [P("Click to add text")], [40, 300, 100, 30]),
      ],
    };
    const r = checkDeck([s], O);
    const kinds = r.issues.map((i) => `${i.kind}:${i.shape}`);
    assert.ok(kinds.includes("overflow:Tràn") && !kinds.includes("overflow:Tự giãn"), kinds.join(","));
    assert.ok(kinds.includes("off_slide:Ngoài") && kinds.includes("placeholder_text:Mẫu"));
    const many = [P("Dòng một của khung"), P("Dòng hai của khung"), P("Dòng ba của khung"), P("Dòng bốn của khung")];
    const two: DeckSlide = { slide: 2, shapes: [title("T"), { ...box("A", many, [40, 120, 300, 100]), n: 2 }, { ...box("B", many, [100, 150, 300, 100]), n: 3 }] };
    const ov = checkDeck([two], O).issues.find((i) => i.kind === "overlap")!;
    assert.match(ov.message, /“A” \(hình số 2\) và “B” \(hình số 3\)/);
    // Khung nội dung cao nhưng ít chữ, nằm trùm lên một hộp chữ ở dưới: chữ không đè nhau thì không báo (ca gặp khi Chạy thật)
    const roomy: DeckSlide = { slide: 3, shapes: [title("T"), box("Nội dung", [P("Ý một"), P("Ý hai"), P("Ý ba")], [36, 94, 648, 230]), box("Rủi ro", [P("Ghi chú ngắn", "Calibri", 14)], [403, 259, 281, 50])] };
    assert.ok(!checkDeck([roomy], O).issues.some((i) => i.kind === "overlap"));
    const fixed: DeckSlide = { slide: 1, shapes: [title("T"), box("Tràn", [P(long, "Calibri", 10)], [40, 120, 560, 110], "AutoSizeNone")] };
    assert.ok(!checkDeck([fixed], O).issues.some((i) => i.kind === "overflow"));
    assert.ok(estimateTextHeight([P("a".repeat(100), "Calibri", 20)], 200) > 100);
  });

  it("Không biết phông chủ đề: lấy phông dùng nhiều nhất làm chuẩn; đoạn lẫn nhiều phông cũng báo", () => {
    const slides: DeckSlide[] = [{ slide: 1, shapes: [title("T"), box("N", [P("Rất nhiều chữ dùng Arial ở đây", "Arial"), P("Một dòng Tahoma", "Tahoma"), { text: "Lẫn phông", font: null, size: 18 }])] }];
    slides[0].shapes[0].paragraphs = [P("T", "Arial", 30)];
    assert.deepStrictEqual(dominantFonts(slides), ["Arial"]);
    const r = checkDeck(slides, { slideW: 720, slideH: 405 });
    assert.deepStrictEqual(r.standardFonts, ["Arial"]);
    assert.strictEqual(r.counts.font, 2);
  });

  it("Tách đoạn của khung chữ PowerPoint theo dấu ngắt đoạn và giữ đúng vị trí", () => {
    assert.deepStrictEqual(splitParagraphs("Một\nHai\rBa\vBốn"), [
      { text: "Một", start: 0 },
      { text: "Hai", start: 4 },
      { text: "Ba", start: 8 },
      { text: "Bốn", start: 11 },
    ]);
    assert.deepStrictEqual(splitParagraphs("A\r\nB"), [
      { text: "A", start: 0 },
      { text: "B", start: 3 },
    ]);
  });

  it("Bảy công cụ PowerPoint mới: có trong bộ công cụ, có nhãn tiếng Việt, lời nhắc hệ thống nhắc tới; công cụ ghi đúng là công cụ ghi", () => {
    assert.strictEqual(PPT_DESIGN_SPECS.length, 7);
    for (const sp of PPT_DESIGN_SPECS) {
      assert.ok(
        pptTools.specs.some((x) => x.name === sp.name),
        sp.name
      );
      assert.notStrictEqual(describeCall({ id: "x", name: sp.name, args: { slide: 1 } }), sp.name);
      assert.ok(pptTools.systemPrompt().includes(sp.name), `lời nhắc thiếu ${sp.name}`);
    }
    assert.deepStrictEqual(
      PPT_DESIGN_SPECS.filter((s) => s.write).map((s) => s.name),
      ["ppt_format_text", "ppt_edit_table", "ppt_set_layout"]
    );
  });
});

describe("Đợt 3 · B3 · Ảnh do công cụ trả về được gửi cho model", () => {
  const sse = (events: any[]) => new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });
  const run = async (vision: boolean | undefined) => {
    let step = 0;
    const bodies: any[] = [];
    const fetchFn: any = async (_u: string, init: any) => {
      bodies.push(JSON.parse(init.body));
      step++;
      return step === 1
        ? sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "ppt_get_slide_image", arguments: JSON.stringify({ slide: 2 }) } }] } }] }])
        : sse([{ choices: [{ delta: { content: "Slide 2 ổn." } }] }]);
    };
    const tools: any = {
      host: "PowerPoint",
      specs: [{ name: "ppt_get_slide_image", description: "x", parameters: { type: "object", properties: {} } }],
      execute: async () => ({ ok: true, slide: 2, measured: [], __images: [{ mime: "image/png", data: "iVBORw0KGgo=", label: "Ảnh slide 2" }] }),
      overview: async () => ({}),
      systemPrompt: () => "x",
    };
    const msgs = await runAgent({
      provider: { id: "openrouter", name: "OpenRouter", kind: "openai", baseUrl: "https://x.invalid/v1" } as any,
      key: "k",
      model: "m",
      tools,
      system: "x",
      messages: [{ role: "user", content: "kiểm slide 2" }],
      log: [],
      onEvent: () => undefined,
      vision,
      fetchFn,
      retryDelaysMs: [],
    });
    return { msgs, bodies };
  };

  it("Model nhận ảnh: kết quả công cụ không chứa dữ liệu ảnh; ảnh đi trong một tin nhắn riêng ngay sau, ghi rõ là dữ liệu", async () => {
    const { msgs, bodies } = await run(true);
    const tool = msgs.find((m) => m.role === "tool") as any;
    assert.ok(!tool.content.includes("iVBOR") && !tool.content.includes("__images"));
    const after = msgs[msgs.indexOf(tool) + 1] as any;
    assert.strictEqual(after.role, "user");
    assert.deepStrictEqual(after.images, [{ mime: "image/png", data: "iVBORw0KGgo=" }]);
    assert.match(after.content, /Ảnh slide 2.*không phải yêu cầu mới của người dùng/);
    assert.ok(JSON.stringify(bodies[1]).includes("iVBORw0KGgo="));
  });

  it("Model không nhận ảnh: không gửi ảnh, kết quả công cụ nói rõ lý do", async () => {
    const { msgs, bodies } = await run(false);
    const tool = msgs.find((m) => m.role === "tool") as any;
    assert.match(tool.content, /không nhận ảnh/);
    assert.ok(!msgs.some((m: any) => m.images?.length));
    assert.ok(!JSON.stringify(bodies[1]).includes("iVBORw0KGgo="));
  });
});
