/* global Word */
/** Hai công cụ Word của Đợt 3 cho văn bản hành chính: dựng trọn một văn bản đúng thể thức theo bản vẽ của core/adminDoc.ts,
 *  và đặt thể thức trang (khổ giấy, lề, số trang). Các API dùng ở đây đã thăm dò trong Word 16/11/2026
 *  (kiem-thu/cdp-bai/tham-do-3-vbhc.js): ghi pageSetup, trường PAGE ở đầu trang, bảng không viền, viền dưới của đoạn (items[2]). */
import { ToolSpec } from "../core/agent/types";
import { ADMIN_PAGE, AdminDocPlan, DOC_TYPES, Para, adminDocBrief, planAdminDocument } from "../core/adminDoc";
import { PT_PER_MM, ptToMm } from "../core/pageformat";

const mm = (v: number) => v * PT_PER_MM;
const FONT = "Times New Roman";

export const WORD_ADMIN_SPECS: ToolSpec[] = [
  {
    name: "word_create_admin_document",
    description:
      "Build a complete Vietnamese administrative document (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản, kế hoạch) with the correct formal layout in one call: national title and motto, issuing organisation, number and symbol, place and date, document title and subject, 'Kính gửi', body, signature block and 'Nơi nhận', all in Times New Roman with the right sizes, alignment and rule lines. On an empty document it also sets A4 paper, standard margins and the page number. ALWAYS use this tool to draft such a document instead of inserting the parts by hand. Leave unknown facts (number, date, signer) out: the tool leaves blanks and reports them. Do not invent them.",
    parameters: {
      type: "object",
      properties: {
        doc_type: { type: "string", enum: DOC_TYPES as unknown as string[], description: "cong_van has no title line and shows the subject as 'V/v …' under the number." },
        org: { type: "string", description: "Issuing organisation, full official name." },
        parent_org: { type: "string", description: "Direct parent organisation, if any." },
        number: { type: "string", description: "Document number. Normally omit: the records clerk assigns it." },
        symbol: { type: "string", description: "Symbol after the number, e.g. 'TTr-ABC' (type abbreviation + organisation) or 'ABC-PC' for a công văn." },
        place: { type: "string", description: "Place name, e.g. 'Hà Nội'." },
        date: { type: "string", description: "e.g. 'ngày 05 tháng 01 năm 2026'. Omit to leave blanks." },
        title: { type: "string", description: "Document title in capitals; only needed for doc_type 'khac' or to override the default." },
        subject: { type: "string", description: "Subject line (trích yếu). For công văn: the text after 'V/v'." },
        to: { type: "array", items: { type: "string" }, description: "'Kính gửi' addressees." },
        body: {
          type: "string",
          description:
            "Body text, one paragraph per line. Line markers: '# ' centred bold heading (e.g. '# QUYẾT ĐỊNH:'), '## ' bold heading (e.g. '## Điều 1. …', '## I. …', '## 1. …'), a whole line wrapped in _underscores_ is italic (legal bases 'Căn cứ …;'). No other markdown.",
        },
        signer_authority: { type: "string", description: "Signing authority line if any, e.g. 'TM. HỘI ĐỒNG QUẢN TRỊ', 'KT. GIÁM ĐỐC'." },
        signer_title: { type: "string", description: "Signer's position, e.g. 'GIÁM ĐỐC', 'PHÓ GIÁM ĐỐC'." },
        signer_name: { type: "string", description: "Signer's full name without academic titles." },
        left_signer_title: { type: "string", description: "Second signer on the left (biên bản: 'THƯ KÝ')." },
        left_signer_name: { type: "string" },
        recipients: { type: "array", items: { type: "string" }, description: "'Nơi nhận' lines. 'Như trên' and 'Lưu: VT, …' are added when missing." },
        font_size: { type: "number", enum: [13, 14], description: "Body size (default 14)." },
        page_setup: { type: "boolean", description: "Set A4, margins and page number for the whole document. Default: true on an empty document, false otherwise." },
        position: { type: "string", enum: ["end", "after_paragraph"], description: "Default end." },
        paragraph_index: { type: "number" },
        expect: { type: "string", description: "With paragraph_index: first ~30 characters of that paragraph." },
      },
      required: ["doc_type", "org", "body"],
    },
    write: true,
  },
  {
    name: "word_set_page_format",
    description:
      "Set paper size and margins for the whole document (not a tracked change; undoable from the pane). Default: the Vietnamese administrative standard — A4 portrait, top 20 mm, bottom 20 mm, left 30 mm, right 20 mm. Use after word_check_page_format reports deviating margins or paper. page_number=true also puts the page number at the top centre, hidden on page 1.",
    parameters: {
      type: "object",
      properties: {
        top_mm: { type: "number" },
        bottom_mm: { type: "number" },
        left_mm: { type: "number" },
        right_mm: { type: "number" },
        a4: { type: "boolean", description: "Set A4 portrait (default true)." },
        page_number: { type: "boolean", description: "Add the page number at the top centre, hidden on page 1 (default false)." },
      },
    },
    write: true,
  },
];

type UndoStep = () => Promise<void>;
interface Deps {
  supports: (v: string, set?: string) => boolean;
  tracked: <T>(ctx: Word.RequestContext, fn: () => Promise<T>) => Promise<T>;
  resolvePara: (ctx: Word.RequestContext, index: any, expect?: string) => Promise<{ p: Word.Paragraph; index: number }>;
}

interface PageOptions {
  topMm: number;
  bottomMm: number;
  leftMm: number;
  rightMm: number;
  a4: boolean;
  pageNumber: boolean;
}

const num = (v: any, d: number, lo: number, hi: number) => {
  const n = Number(v);
  return Number.isFinite(n) && n >= lo && n <= hi ? n : d;
};

export function createWordAdminTools(d: Deps) {
  const names = new Set(WORD_ADMIN_SPECS.map((s) => s.name));

  /** Chạy với Theo dõi thay đổi tắt (thể thức trang và đầu trang không phải nội dung để duyệt), xong trả lại chế độ cũ */
  async function untracked<T>(ctx: Word.RequestContext, fn: () => Promise<T>): Promise<T> {
    if (!d.supports("1.4")) return fn();
    ctx.document.load("changeTrackingMode");
    await ctx.sync();
    const mode = ctx.document.changeTrackingMode;
    if (mode !== "Off") {
      ctx.document.changeTrackingMode = "Off" as any;
      await ctx.sync();
    }
    try {
      return await fn();
    } finally {
      if (mode !== "Off") {
        ctx.document.changeTrackingMode = mode;
        await ctx.sync();
      }
    }
  }

  /** Đặt khổ giấy, lề, số trang; trả mô tả và bước hoàn tác */
  async function applyPage(ctx: Word.RequestContext, o: PageOptions): Promise<{ note: string; undo: UndoStep | null; changed: boolean }> {
    if (!d.supports("1.3", "WordApiDesktop")) {
      return { note: "Word trên máy này không cho add-in đặt lề và khổ giấy (cần WordApiDesktop 1.3): hãy đặt tay ở thẻ Layout.", undo: null, changed: false };
    }
    return untracked(ctx, async () => {
      const secs = ctx.document.sections;
      secs.load("items");
      await ctx.sync();
      const setups = secs.items.slice(0, 30).map((s) => s.pageSetup);
      setups.forEach((p) => p.load("topMargin,bottomMargin,leftMargin,rightMargin,paperSize,orientation,pageWidth,pageHeight,differentFirstPageHeaderFooter"));
      await ctx.sync();
      const before = setups.map((p) => ({
        top: p.topMargin,
        bottom: p.bottomMargin,
        left: p.leftMargin,
        right: p.rightMargin,
        w: p.pageWidth,
        h: p.pageHeight,
        orientation: String(p.orientation),
        first: p.differentFirstPageHeaderFooter,
      }));
      for (const p of setups) {
        if (o.a4) {
          p.orientation = "Portrait" as any;
          p.paperSize = "A4" as any;
        }
        p.topMargin = mm(o.topMm);
        p.bottomMargin = mm(o.bottomMm);
        p.leftMargin = mm(o.leftMm);
        p.rightMargin = mm(o.rightMm);
      }
      await ctx.sync();
      let numbered = false;
      let numberNote = "";
      if (o.pageNumber) {
        try {
          const first = secs.items[0];
          const h = first.getHeader("Primary" as any);
          h.load("text");
          await ctx.sync();
          if (h.text.replace(/[\s\u0001-\u001f]/g, "")) numberNote = " Đầu trang đang có nội dung nên không chèn số trang.";
          else {
            const hp = h.paragraphs.getFirst();
            hp.alignment = "Centered" as any;
            hp.getRange("Start").insertField("Start" as any, "Page" as any);
            hp.font.name = FONT;
            hp.font.size = 13;
            first.pageSetup.differentFirstPageHeaderFooter = true;
            await ctx.sync();
            numbered = true;
          }
        } catch (e: any) {
          numberNote = ` Không chèn được số trang (${e?.message || e}).`;
        }
      }
      const b0 = before[0];
      const same =
        !!b0 &&
        before.every(
          (b) =>
            Math.abs(b.top - mm(o.topMm)) < 0.5 &&
            Math.abs(b.bottom - mm(o.bottomMm)) < 0.5 &&
            Math.abs(b.left - mm(o.leftMm)) < 0.5 &&
            Math.abs(b.right - mm(o.rightMm)) < 0.5 &&
            (!o.a4 || (Math.abs(b.w - 595.3) < 1.5 && b.orientation === "Portrait"))
        );
      const undo: UndoStep = () =>
        Word.run(async (c2) =>
          untracked(c2, async () => {
            const ss = c2.document.sections;
            ss.load("items");
            await c2.sync();
            if (numbered) {
              ss.items[0].getHeader("Primary" as any).clear();
              await c2.sync();
            }
            ss.items.slice(0, before.length).forEach((s, k) => {
              const b = before[k];
              const p = s.pageSetup;
              if (o.a4) {
                p.orientation = b.orientation as any;
                p.pageWidth = b.w;
                p.pageHeight = b.h;
              }
              p.topMargin = b.top;
              p.bottomMargin = b.bottom;
              p.leftMargin = b.left;
              p.rightMargin = b.right;
              if (numbered && k === 0) p.differentFirstPageHeaderFooter = b.first;
            });
            await c2.sync();
          })
        );
      const was = b0 ? `trước đó: trên ${ptToMm(b0.top)}, dưới ${ptToMm(b0.bottom)}, trái ${ptToMm(b0.left)}, phải ${ptToMm(b0.right)} mm` : "";
      return {
        note: `${o.a4 ? "A4 dọc, " : ""}lề trên ${o.topMm}, dưới ${o.bottomMm}, trái ${o.leftMm}, phải ${o.rightMm} mm${same ? " (đã đúng từ trước)" : was ? ` (${was})` : ""}${numbered ? "; số trang ở giữa đầu trang, ẩn ở trang 1" : ""}.${numberNote}`,
        undo: same && !numbered ? null : undo,
        changed: !same || numbered,
      };
    });
  }

  const pageOptions = (a: any, pageNumber: boolean): PageOptions => ({
    topMm: num(a.top_mm, ADMIN_PAGE.topMm, 5, 80),
    bottomMm: num(a.bottom_mm, ADMIN_PAGE.bottomMm, 5, 80),
    leftMm: num(a.left_mm, ADMIN_PAGE.leftMm, 5, 80),
    rightMm: num(a.right_mm, ADMIN_PAGE.rightMm, 5, 80),
    a4: a.a4 !== false,
    pageNumber,
  });

  /** Văn bản trống = không đoạn nào có chữ và không có bảng */
  async function isEmptyDoc(ctx: Word.RequestContext): Promise<{ empty: boolean; paras: number }> {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    const ts = ctx.document.body.tables;
    ts.load("items");
    await ctx.sync();
    return { empty: !ts.items.length && ps.items.every((p) => !p.text.trim()), paras: ps.items.filter((p) => p.text.trim()).length };
  }

  function fmt(p: Word.Paragraph, x: Para, widthMm: number, rules: { p: Word.Paragraph }[]) {
    p.font.name = FONT;
    p.font.size = x.size;
    p.font.bold = !!x.bold;
    p.font.italic = !!x.italic;
    p.font.underline = "None" as any;
    p.font.color = "#000000";
    p.alignment = x.align as any;
    p.firstLineIndent = x.indent ? mm(10) : 0;
    p.spaceAfter = x.after || 0;
    p.spaceBefore = x.before || 0;
    p.lineSpacing = 12 * (x.line || 1);
    if (x.ruleMm) {
      const side = Math.max(0, (widthMm - x.ruleMm) / 2);
      p.leftIndent = mm(side);
      p.rightIndent = mm(side);
      rules.push({ p });
    } else {
      p.leftIndent = 0;
      p.rightIndent = 0;
    }
  }

  async function build(ctx: Word.RequestContext, plan: AdminDocPlan, anchor: Word.Paragraph, where: "Before" | "After", textWidthMm: number) {
    const rules: { p: Word.Paragraph }[] = [];
    const plain: Word.Paragraph[] = [];
    let prev: Word.Paragraph | Word.Table | null = null;
    for (const b of plan.blocks) {
      if (b.kind === "table") {
        const values = b.rows.map((r) => [r[0].text, r[1].text]);
        const t: Word.Table = prev ? prev.insertTable(values.length, 2, "After", values) : anchor.insertTable(values.length, 2, where, values);
        t.headerRowCount = 0;
        t.getBorder("All" as any).type = "None" as any;
        t.alignment = "Centered" as any;
        t.setCellPadding("Left" as any, 0);
        t.setCellPadding("Right" as any, 0);
        t.getCell(0, 0).columnWidth = mm(b.widths[0]);
        t.getCell(0, 1).columnWidth = mm(b.widths[1]);
        b.rows.forEach((r, ri) =>
          r.forEach((cell, ci) => {
            fmt(t.getCell(ri, ci).body.paragraphs.getFirst(), cell, b.widths[ci], rules);
          })
        );
        prev = t;
        await ctx.sync();
      } else {
        const p: Word.Paragraph = prev ? prev.insertParagraph(b.text, "After") : anchor.insertParagraph(b.text, where);
        p.styleBuiltIn = "Normal" as any;
        fmt(p, b, textWidthMm, rules);
        plain.push(p);
        prev = p;
      }
    }
    await ctx.sync();
    // Đoạn chèn sau một mục danh sách thừa hưởng danh sách của nó: gỡ ra
    plain.forEach((p) => p.load("isListItem"));
    await ctx.sync();
    const inList = plain.filter((p) => p.isListItem);
    if (inList.length) {
      inList.forEach((p) => p.detachFromList());
      await ctx.sync();
    }
    // Đường kẻ = viền dưới của một đoạn rỗng (items: 0 trên, 1 trái, 2 dưới, 3 phải — đã đo)
    if (rules.length && d.supports("1.3", "WordApiDesktop")) {
      rules.forEach((r) => r.p.borders.load("items"));
      await ctx.sync();
      for (const r of rules) {
        const bottom = r.p.borders.items[2];
        if (!bottom) continue;
        bottom.lineStyle = "Single" as any;
        bottom.lineWidth = "Pt075" as any;
      }
      await ctx.sync();
    }
    return { rules: rules.length };
  }

  async function createDocument(a: any, extra: UndoStep[]) {
    const plan = planAdminDocument({ ...a, year: new Date().getFullYear() });
    return Word.run(async (ctx) => {
      const state = await isEmptyDoc(ctx);
      const wantPage = a.page_setup === undefined || a.page_setup === null ? state.empty : !!a.page_setup;
      let pageNote = "";
      if (wantPage) {
        const pg = await applyPage(ctx, { ...pageOptions({}, true) });
        pageNote = ` Thể thức trang: ${pg.note}`;
        if (pg.undo) extra.push(pg.undo);
      }
      // Bề rộng vùng trình bày thật của trang (để canh đường kẻ dưới trích yếu)
      let width = 160;
      if (d.supports("1.3", "WordApiDesktop")) {
        const ps = ctx.document.sections.getFirst().pageSetup;
        ps.load("pageWidth,leftMargin,rightMargin");
        await ctx.sync();
        width = ptToMm(ps.pageWidth - ps.leftMargin - ps.rightMargin);
      }
      const res = await d.tracked(ctx, async () => {
        let anchor: Word.Paragraph;
        let where: "Before" | "After" = "After";
        if (a.position === "after_paragraph") anchor = (await d.resolvePara(ctx, a.paragraph_index, a.expect)).p;
        else if (state.empty) {
          anchor = ctx.document.body.paragraphs.getFirst();
          where = "Before";
        } else anchor = ctx.document.body.paragraphs.getLast();
        return build(ctx, plan, anchor, where, width);
      });
      const narrow = width < 150 ? ` Lưu ý: vùng trình bày của trang chỉ rộng ${width} mm (chuẩn 160 mm), các bảng thể thức có thể tràn lề — đặt lại lề bằng word_set_page_format.` : "";
      return {
        ok: true,
        brief: adminDocBrief(plan),
        message: `Đã dựng văn bản đúng thể thức (ghi có theo dõi): phần đầu, ${plan.type === "cong_van" ? "V/v" : "tên loại và trích yếu"}, nội dung ${plan.blocks.filter((b) => b.kind === "para").length} đoạn, phần ký và nơi nhận; ${res.rules} đường kẻ.${pageNote}${narrow}`,
        blanks: plan.blanks,
        ...(plan.notes.length ? { notes: plan.notes } : {}),
        next: "Báo người dùng các chỗ còn trống (blanks) cần điền trước khi ký. Muốn sửa câu chữ thì dùng word_set_paragraph_text trên đoạn tương ứng; không dựng lại cả văn bản.",
      };
    });
  }

  async function setPage(a: any, extra: UndoStep[]) {
    return Word.run(async (ctx) => {
      const pg = await applyPage(ctx, pageOptions(a || {}, a?.page_number === true));
      if (pg.undo) extra.push(pg.undo);
      return { ok: true, changed: pg.changed, brief: pg.changed ? "đã đặt lại thể thức trang" : "thể thức trang đã đúng từ trước", message: `Thể thức trang: ${pg.note} Thay đổi này không ghi thành sửa đổi được theo dõi; hoàn tác bằng nút Hoàn tác của khung.` };
    });
  }

  return {
    has: (name: string) => names.has(name),
    run: (name: string, a: any, extra: UndoStep[]) => (name === "word_set_page_format" ? setPage(a, extra) : createDocument(a || {}, extra)),
    /** Cảnh báo rủi ro: đổi thể thức trang của một văn bản đang có nội dung khi dựng văn bản mới vào đó */
    risk: {
      word_create_admin_document: (a: any) =>
        Word.run(async (ctx) => {
          if (a?.page_setup !== true) return null;
          const s = await isEmptyDoc(ctx);
          return s.empty ? null : `Văn bản đang có ${s.paras} đoạn nội dung: thao tác này đổi khổ giấy, lề và đầu trang của CẢ văn bản, rồi chèn văn bản mới vào.`;
        }),
    } as Record<string, (a: any) => Promise<string | null> | string | null>,
  };
}
