/** B7 phần 1 + B5 (nhóm 2 rà soát kiến trúc, bước 4a, 26/11/2026): phần thuần tách khỏi Chat.tsx — tệp đính kèm, gợi ý nhóm công cụ,
 *  lệnh ghi chéo — và chụp tài liệu chỉ khi có khung khác mở. Các quy tắc Đợt 5, 10–12 trước đây chỉ Chạy thật mới kiểm được. */
import * as assert from "assert";
import { groupHintsFor, handleIncoming, prepareAttachments, AttachmentsInput } from "../../src/core/turnPrep";
import { getProvider } from "../../src/core/providers";
import { DOCX_MIME, MAX_FILES_TOTAL_B64, PDF_MIME } from "../../src/core/attachments";
import { SHARE_MAX_AGE_MS, SHARE_PREFIX, clearSnapshot, publishSnapshot, readSnapshot, shouldShare } from "../../src/core/share";

const base = (o: Partial<AttachmentsInput> = {}): AttachmentsInput => ({
  content: "Hỏi",
  atts: [],
  host: "Word",
  provider: getProvider("anthropic")!,
  modelId: "claude-sonnet-5-5",
  vision: true,
  useAgent: true,
  agentOn: true,
  historyFilesB64: 0,
  ...o,
});
const scan = (name: string, pages = 2, data = "JVBERi0x") => ({ name, kind: "file" as const, mime: PDF_MIME, data, pages });

describe("B7 · prepareAttachments (tệp đính kèm của một lượt)", () => {
  it("không có tệp (lượt Làm tiếp): nội dung giữ nguyên, không thông báo, không đăng ký gì", () => {
    const r: any = prepareAttachments(base());
    assert.strictEqual(r.stop, false);
    assert.strictEqual(r.content, "Hỏi");
    assert.deepStrictEqual([r.images, r.pdfs, r.register, r.notice], [[], [], [], undefined]);
  });

  it("PDF quét ngắn gửi nguyên tệp và được đăng ký; nguồn đọc được PDF thì không báo lỗi", () => {
    const r: any = prepareAttachments(base({ atts: [scan("a.pdf")] }));
    assert.strictEqual(r.pdfs.length, 1);
    assert.deepStrictEqual(r.register.map((f: any) => f.name), ["a.pdf"]);
    assert.match(r.content, /Tệp đính kèm: a\.pdf/);
    assert.ok(!r.notice || !r.notice.err);
  });

  it("quá trần tổng dung lượng tệp trong hội thoại: dừng, báo lỗi, không gửi", () => {
    const r: any = prepareAttachments(base({ atts: [scan("b.pdf", 2, "x".repeat(10))], historyFilesB64: MAX_FILES_TOTAL_B64 }));
    assert.strictEqual(r.stop, true);
    assert.ok(r.notice.err && /quá lớn để gửi tiếp/.test(r.notice.text));
  });

  it("nguồn không đọc được PDF, tắt tác tử: báo lỗi nêu tên tệp, chỉ đường đổi nguồn", () => {
    const r: any = prepareAttachments(base({ provider: getProvider("groq")!, useAgent: false, agentOn: false, atts: [scan("c.pdf")] }));
    assert.ok(r.notice?.err && /c\.pdf/.test(r.notice.text) && /Anthropic hoặc Gemini/.test(r.notice.text));
  });

  it("gửi qua OpenRouter: nhắc riêng tư (không chặn)", () => {
    const r: any = prepareAttachments(base({ provider: getProvider("openrouter")!, modelId: "x/y:free", atts: [{ name: "t.txt", kind: "text", text: "abc" }] }));
    assert.ok(r.notice && !r.notice.err && /OpenRouter/.test(r.notice.text));
  });

  it(".docx được đăng ký bản gốc; PowerPoint + tác tử: ảnh đăng ký và nêu tên trong tin nhắn", () => {
    const d: any = prepareAttachments(base({ atts: [{ name: "hd.docx", kind: "text", text: "x", mime: DOCX_MIME, data: "UEsD" }] }));
    assert.deepStrictEqual(d.register.map((f: any) => f.name), ["hd.docx"]);
    const p: any = prepareAttachments(base({ host: "PowerPoint", atts: [{ name: "logo.png", kind: "image", mime: "image/png", data: "iVBO" }] }));
    assert.deepStrictEqual(p.register.map((f: any) => f.name), ["logo.png"]);
    assert.strictEqual(p.images.length, 1);
    assert.match(p.content, /Ảnh đính kèm: “logo\.png”/);
    const w: any = prepareAttachments(base({ host: "Word", atts: [{ name: "logo.png", kind: "image", mime: "image/png", data: "iVBO" }] }));
    assert.deepStrictEqual(w.register, []);
  });
});

describe("B7 · groupHintsFor", () => {
  it("Word + .docx → soat; PowerPoint + ảnh lượt này hoặc còn trong phiên → hinh_anh; Làm tiếp (không tệp) → rỗng", () => {
    assert.deepStrictEqual(groupHintsFor({ host: "Word", atts: [{ name: "a.docx", kind: "text", mime: DOCX_MIME }], images: 0, sessionImages: 0 }), ["soat"]);
    assert.deepStrictEqual(groupHintsFor({ host: "PowerPoint", atts: [], images: 1, sessionImages: 0 }), ["hinh_anh"]);
    assert.deepStrictEqual(groupHintsFor({ host: "PowerPoint", atts: [], images: 0, sessionImages: 2 }), ["hinh_anh"]);
    assert.deepStrictEqual(groupHintsFor({ host: "Word", atts: [], images: 3, sessionImages: 3 }), []);
    assert.deepStrictEqual(groupHintsFor({ host: "Excel", atts: [{ name: "a.docx", kind: "text", mime: DOCX_MIME }], images: 0, sessionImages: 0 }), []);
  });
});

describe("B7 · handleIncoming (một lệnh ghi chéo)", () => {
  const cmd = { id: "c1", from: "Word", payload: { op: "x" } };
  const o = (exec: any) => ({ selfApp: "Excel", linked: ["Word"], what: "ghi A1", exec, errText: (e: any) => String(e?.message || e) });
  it("khung tắt kết nối với nguồn gửi: từ chối, không ghi, không có mục", async () => {
    let ran = false;
    const r = await handleIncoming({ ...cmd, from: "PowerPoint" }, o(async () => ((ran = true), "ok")));
    assert.strictEqual(ran, false);
    assert.deepStrictEqual(r.reply, { ok: false, error: "Khung AI Đa Nguồn ở Excel đang tắt kết nối với PowerPoint." });
    assert.strictEqual(r.item, undefined);
  });
  it("ghi được: trả lời ok, mục có Hoàn tác", async () => {
    const r = await handleIncoming(cmd, o(async (_p: any, log: any[]) => (log.push({ label: "x", undo: async () => {} }), "Đã ghi")));
    assert.deepStrictEqual(r.reply, { ok: true, message: "Đã ghi" });
    assert.strictEqual(r.item?.undo?.length, 1);
    assert.match(r.item!.incoming, /từ khung AI Đa Nguồn ở Word: ghi A1/);
  });
  it("ghi dở rồi lỗi: gỡ phần đã làm theo thứ tự ngược, báo lỗi, không có mục", async () => {
    const undone: number[] = [];
    const r = await handleIncoming(
      cmd,
      o(async (_p: any, log: any[]) => {
        log.push({ label: "1", undo: async () => undone.push(1) }, { label: "2", undo: async () => undone.push(2) });
        throw new Error("hỏng");
      })
    );
    assert.deepStrictEqual(undone, [2, 1]);
    assert.deepStrictEqual(r.reply, { ok: false, error: "Excel không ghi được: hỏng" });
    assert.strictEqual(r.item, undefined);
  });
});

describe("B5 · chụp tài liệu cho đọc chéo", () => {
  const mem = () => {
    const m = new Map<string, string>();
    return { getItem: (k: string) => m.get(k) ?? null, setItem: (k: string, v: string) => void m.set(k, v), removeItem: (k: string) => void m.delete(k), m };
  };
  it("chỉ chụp khi có khung ở ứng dụng khác đang mở (và bật chia sẻ, có công cụ chụp, không bận)", () => {
    const ok = { shareContext: true, canSnapshot: true, busy: false, links: [{ open: true }] };
    assert.strictEqual(shouldShare(ok), true);
    assert.strictEqual(shouldShare({ ...ok, links: [{ open: false }] }), false);
    assert.strictEqual(shouldShare({ ...ok, links: [] }), false);
    assert.strictEqual(shouldShare({ ...ok, busy: true }), false);
    assert.strictEqual(shouldShare({ ...ok, shareContext: false }), false);
    assert.strictEqual(shouldShare({ ...ok, canSnapshot: false }), false);
  });
  it("bản quá hạn bị gỡ khi đọc; đóng khung gỡ bản của mình", () => {
    const s = mem();
    publishSnapshot("Word", { a: 1 }, "t", s);
    assert.ok(readSnapshot("Word", Date.now(), s));
    assert.strictEqual(readSnapshot("Word", Date.now() + SHARE_MAX_AGE_MS + 1000, s), null);
    assert.strictEqual(s.m.has(SHARE_PREFIX + "Word"), false);
    publishSnapshot("Excel", { b: 2 }, "t", s);
    clearSnapshot("Excel", s);
    assert.strictEqual(s.m.has(SHARE_PREFIX + "Excel"), false);
  });
});
