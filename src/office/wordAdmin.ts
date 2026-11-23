/* global Word */
/** Hai công cụ Word của Đợt 3 cho văn bản hành chính: dựng trọn một văn bản đúng thể thức theo bản vẽ của core/adminDoc.ts,
 *  và đặt thể thức trang (khổ giấy, lề, số trang). Các API dùng ở đây đã thăm dò trong Word 16/11/2026
 *  (kiem-thu/cdp-bai/tham-do-3-vbhc.js): ghi pageSetup, trường PAGE ở đầu trang, bảng không viền, viền dưới của đoạn (items[2]). */
import { ToolSpec } from "../core/agent/types";
import {
  ADMIN_PAGE,
  AdminDocPlan,
  DOC_TYPES,
  Para,
  adminDocBrief,
  planAdminDocument,
} from "../core/adminDoc";
import { PT_PER_MM, ptToMm } from "../core/pageformat";

const mm = (v: number) => v * PT_PER_MM;
const FONT = "Times New Roman";

export const WORD_ADMIN_SPECS: ToolSpec[] = [
  {
    name: "word_create_admin_document",
    description:
      "Build a complete Vietnamese administrative document (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản, kế hoạch) or a contract (hop_dong) in the correct formal layout in ONE call: national motto, organisation, number/symbol, place/date, title, subject, Kính gửi, body, signature block, Nơi nhận. On an empty document also sets A4, margins, page number. ALWAYS use it instead of inserting parts by hand. body = content only, never the header, title, number, date or signature lines. Omit unknown facts (number, date, signer): blanks are left and reported; never invent.",
    parameters: {
      type: "object",
      properties: {
        doc_type: {
          type: "string",
          enum: DOC_TYPES as unknown as string[],
          description:
            "cong_van: no title line, subject shown as 'V/v …'. hop_dong: contract, two-party signatures.",
        },
        org: {
          type: "string",
          description:
            "Issuing organisation, full official name. hop_dong: optional (top-left, opposite the motto).",
        },
        parent_org: {
          type: "string",
          description:
            "Organisation directly above the issuing one, if any.",
        },
        number: {
          type: "string",
          description: "Normally omit.",
        },
        symbol: {
          type: "string",
          description:
            "e.g. 'TTr-ABC', 'ABC-PC'.",
        },
        place: { type: "string", description: "Place name, e.g. 'Hà Nội'." },
        date: {
          type: "string",
          description: "e.g. 'ngày 05 tháng 01 năm 2026'. Omit = blanks.",
        },
        title: {
          type: "string",
          description:
            "For 'khac'; contract name for hop_dong.",
        },
        subject: {
          type: "string",
          description: "Trích yếu (công văn: text after 'V/v').",
        },
        to: { type: "array", items: { type: "string" }, description: "'Kính gửi' addressees." },
        body: {
          type: "string",
          description:
            "One paragraph per line. '# ' centred bold ('# QUYẾT ĐỊNH:'), '## ' bold heading ('## Điều 1. …', '## I. …'), a whole line in _underscores_ italic ('Căn cứ …;'). No other markdown.",
        },
        signer_authority: {
          type: "string",
          description:
            "e.g. 'TM. HỘI ĐỒNG QUẢN TRỊ', 'KT. GIÁM ĐỐC'. hop_dong: default 'ĐẠI DIỆN BÊN B'.",
        },
        signer_title: {
          type: "string",
          description: "e.g. 'GIÁM ĐỐC'.",
        },
        signer_name: { type: "string" },
        left_signer_title: {
          type: "string",
          description: "Left signer (biên bản: 'THƯ KÝ'; hop_dong: Bên A).",
        },
        left_signer_name: { type: "string" },
        left_signer_authority: { type: "string", description: "hop_dong: default 'ĐẠI DIỆN BÊN A'." },
        recipients: {
          type: "array",
          items: { type: "string" },
          description: "Nơi nhận lines ('Như trên', 'Lưu: VT' auto-added).",
        },
        font_size: { type: "number", enum: [13, 14], description: "Body size (default 14)." },
        page_setup: {
          type: "boolean",
          description:
            "Set A4, margins, page number. Default: true only on an empty document.",
        },
        position: { type: "string", enum: ["end", "after_paragraph"], description: "Default end." },
        paragraph_index: { type: "number" },
        expect: {
          type: "string",
          description: "First ~30 characters of that paragraph.",
        },
      },
      required: ["doc_type", "body"],
    },
    write: true,
  },
  {
    name: "word_set_page_format",
    description:
      "Set paper and margins for the whole document (not tracked; undoable). Default: administrative standard, A4 portrait, top 20, bottom 20, left 30, right 20 mm. Use after word_check_page_format reports deviations.",
    parameters: {
      type: "object",
      properties: {
        top_mm: { type: "number" },
        bottom_mm: { type: "number" },
        left_mm: { type: "number" },
        right_mm: { type: "number" },
        a4: { type: "boolean", description: "A4 portrait. Default true." },
        page_number: {
          type: "boolean",
          description: "Page number at top centre, hidden on page 1. Default false.",
        },
      },
    },
    write: true,
  },
];

type UndoStep = () => Promise<void>;
interface Deps {
  supports: (v: string, set?: string) => boolean;
  tracked: <T>(ctx: Word.RequestContext, fn: () => Promise<T>) => Promise<T>;
  resolvePara: (
    ctx: Word.RequestContext,
    index: any,
    expect?: string
  ) => Promise<{ p: Word.Paragraph; index: number }>;
}

interface PageOptions {
  topMm: number;
  bottomMm: number;
  leftMm: number;
  rightMm: number;
  a4: boolean;
  pageNumber: boolean;
}

/** Model yếu hay truyền "true", 1, "yes" thay cho true */
export const asBool = (v: any): boolean | undefined => {
  if (v === undefined || v === null || v === "") return undefined;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  const t = String(v).trim().toLowerCase();
  if (["true", "1", "yes", "y", "có", "co"].includes(t)) return true;
  if (["false", "0", "no", "n", "không", "khong"].includes(t)) return false;
  return undefined;
};

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
  async function applyPage(
    ctx: Word.RequestContext,
    o: PageOptions
  ): Promise<{ note: string; undo: UndoStep | null; changed: boolean }> {
    if (!d.supports("1.3", "WordApiDesktop")) {
      return {
        note: "Word trên máy này không cho add-in đặt lề và khổ giấy (cần WordApiDesktop 1.3): hãy đặt tay ở thẻ Layout.",
        undo: null,
        changed: false,
      };
    }
    return untracked(ctx, async () => {
      const secs = ctx.document.sections;
      secs.load("items");
      await ctx.sync();
      const setups = secs.items.slice(0, 30).map((s) => s.pageSetup);
      setups.forEach((p) =>
        p.load(
          "topMargin,bottomMargin,leftMargin,rightMargin,paperSize,orientation,pageWidth,pageHeight,differentFirstPageHeaderFooter"
        )
      );
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
          // Đầu trang chỉ có ảnh hoặc hình (logo) thì chữ rỗng: phải đếm cả ảnh, hình, nếu không số trang và "trang đầu khác" làm mất logo
          const pics = h.inlinePictures;
          pics.load("items");
          const floats = d.supports("1.2", "WordApiDesktop") ? h.shapes : null;
          floats?.load("items");
          await ctx.sync();
          if (h.text.replace(/[\s\u0001-\u001f]/g, "") || pics.items.length || floats?.items.length)
            numberNote = " Đầu trang đang có nội dung nên không chèn số trang.";
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
      const was = b0
        ? `trước đó: trên ${ptToMm(b0.top)}, dưới ${ptToMm(b0.bottom)}, trái ${ptToMm(b0.left)}, phải ${ptToMm(b0.right)} mm`
        : "";
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
    a4: asBool(a.a4) !== false,
    pageNumber,
  });

  /** Văn bản trống = không đoạn nào có chữ và không có bảng */
  async function isEmptyDoc(ctx: Word.RequestContext): Promise<{ empty: boolean; paras: number }> {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    const ts = ctx.document.body.tables;
    ts.load("items");
    // Văn bản chỉ có ảnh, ô điền mẫu hay hình thì không phải "trống"
    const pics = ctx.document.body.inlinePictures;
    pics.load("items");
    const ccs = ctx.document.body.contentControls;
    ccs.load("items");
    const floats = d.supports("1.2", "WordApiDesktop") ? ctx.document.body.shapes : null;
    floats?.load("items");
    await ctx.sync();
    return {
      empty:
        !ts.items.length &&
        !pics.items.length &&
        !ccs.items.length &&
        !floats?.items.length &&
        ps.items.every((p) => !p.text.trim()),
      paras: ps.items.filter((p) => p.text.trim()).length,
    };
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
      p.leftIndent = x.leftMm ? mm(x.leftMm) : 0;
      p.rightIndent = 0;
    }
  }

  async function build(
    ctx: Word.RequestContext,
    plan: AdminDocPlan,
    anchor: Word.Paragraph,
    where: "Before" | "After",
    textWidthMm: number,
    bag: { plain: Word.Paragraph[]; tables: Word.Table[] }
  ) {
    const rules: { p: Word.Paragraph }[] = [];
    // Ghi vào `bag` ngay khi chèn: nếu lỗi giữa chừng, bên gọi gỡ được phần đã chèn dở
    const plain = bag.plain;
    const tables = bag.tables;
    /** Định dạng làm SAU, khi Theo dõi thay đổi đã tắt: phần chèn là sửa đổi được theo dõi, còn định dạng của chính phần mới chèn
     *  không ghi thành một loạt sửa đổi "Formatted" riêng (xem ảnh thử 16/11/2026: mỗi đoạn một ô chú thích ở lề, rất rối khi duyệt) */
    const later: (() => void)[] = [];
    // Đoạn đứng ngay sau chỗ chèn (trước khi chèn): để nhận ra đoạn rỗng Word tự sinh sau bảng cuối (đo 23/11/2026 với phần ký
    // của hợp đồng: Word thêm một đoạn "Normal" rỗng sau bảng, không thuộc phần đã chèn nên Hoàn tác để sót)
    let nextAt: Word.Paragraph | null = null;
    if (d.supports("1.6")) {
      nextAt = where === "Before" ? anchor : anchor.getNextOrNullObject();
      nextAt.load("uniqueLocalId");
      await ctx.sync();
    }
    const nextId = nextAt && !nextAt.isNullObject ? nextAt.uniqueLocalId : "";
    let prev: Word.Paragraph | Word.Table | null = null;
    for (const b of plan.blocks) {
      if (b.kind === "table") {
        // Ô nhiều đoạn: đoạn sẵn có của ô mang chữ của đoạn CUỐI, các đoạn trước chèn "Before" nó (đã đo: "After" rơi sang ô kế)
        const values = b.rows.map((r) => [r[0][r[0].length - 1].text, r[1][r[1].length - 1].text]);
        const t: Word.Table = prev
          ? prev.insertTable(values.length, 2, "After", values)
          : anchor.insertTable(values.length, 2, where, values);
        later.push(() => {
          t.headerRowCount = 0;
          t.getBorder("All" as any).type = "None" as any;
          t.alignment = "Centered" as any;
          t.setCellPadding("Left" as any, 0);
          t.setCellPadding("Right" as any, 0);
          t.getCell(0, 0).columnWidth = mm(b.widths[0]);
          t.getCell(0, 1).columnWidth = mm(b.widths[1]);
        });
        b.rows.forEach((r, ri) =>
          r.forEach((cell, ci) => {
            let cur = t.getCell(ri, ci).body.paragraphs.getFirst();
            const last = cur;
            later.push(() => fmt(last, cell[cell.length - 1], b.widths[ci], rules));
            for (let k = cell.length - 2; k >= 0; k--) {
              const np = cur.insertParagraph(cell[k].text, "Before");
              later.push(() => fmt(np, cell[k], b.widths[ci], rules));
              cur = np;
            }
          })
        );
        prev = t;
        tables.push(t);
        await ctx.sync();
      } else {
        const p: Word.Paragraph = prev
          ? prev.insertParagraph(b.text, "After")
          : anchor.insertParagraph(b.text, where);
        later.push(() => {
          p.styleBuiltIn = "Normal" as any;
          fmt(p, b, textWidthMm, rules);
        });
        plain.push(p);
        prev = p;
      }
    }
    await ctx.sync();
    return { finish };

    async function finish() {
      later.forEach((f) => f());
      await ctx.sync();
      // Bảng là khối cuối: đoạn rỗng Word tự sinh ngay sau nó (không phải đoạn vốn có của văn bản) thì bỏ
      const lastBlock = plan.blocks[plan.blocks.length - 1];
      if (nextId && lastBlock?.kind === "table" && tables.length) {
        const extra = tables[tables.length - 1].getParagraphAfterOrNullObject();
        extra.load("text,uniqueLocalId,tableNestingLevel");
        await ctx.sync();
        if (!extra.isNullObject && extra.uniqueLocalId !== nextId && !extra.text.trim() && extra.tableNestingLevel === 0) {
          extra.delete();
          await ctx.sync();
        }
      }
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
      // Định danh của thứ đã chèn — cho cách hoàn tác thay thế khi Word gộp phần chèn vào một sửa đổi có từ trước
      const made: { id: string; head: string; table: boolean }[] = [];
      if (d.supports("1.6")) {
        const marks: { p: Word.Paragraph; table: boolean }[] = [];
        // Theo đúng thứ tự trong văn bản: khi gỡ (từ cuối lên) không đoạn rỗng nào còn đứng ngay trước một bảng
        // (đã đo: xóa đoạn rỗng đứng ngay trước bảng thì Word giữ lại đoạn đó)
        let k = 0;
        let j = 0;
        for (const b of plan.blocks) {
          if (b.kind === "table")
            marks.push({ p: tables[k++].getCell(0, 0).body.paragraphs.getFirst(), table: true });
          else marks.push({ p: plain[j++], table: false });
        }
        marks.forEach((m) => m.p.load("uniqueLocalId,text"));
        await ctx.sync();
        marks.forEach((m) =>
          made.push({ id: m.p.uniqueLocalId, head: m.p.text.slice(0, 20), table: m.table })
        );
      }
      return { rules: rules.length, made };
    }
  }

  /** Gỡ hẳn các bảng và đoạn công cụ đã chèn (tìm theo định danh, đối chiếu chữ đầu); thứ không còn thì bỏ qua */
  const removeMade =
    (made: { id: string; head: string; table: boolean }[]): UndoStep =>
    async () => {
      for (const m of made.slice().reverse()) {
        await Word.run(async (ctx) =>
          untracked(ctx, async () => {
            try {
              const p = ctx.document.getParagraphByUniqueLocalId(m.id);
              p.load("text");
              await ctx.sync();
              if (m.head ? !p.text.includes(m.head) : !!p.text.trim()) return;
              if (m.table) p.parentTable.delete();
              else p.delete();
              await ctx.sync();
            } catch {
              /* đã không còn (sửa đổi đã bị từ chối) */
            }
          })
        );
      }
    };

  async function createDocument(
    a: any,
    extra: UndoStep[],
    inverse: UndoStep[],
    touched: string[],
    after: UndoStep[]
  ) {
    const plan = planAdminDocument({ ...a, year: new Date().getFullYear() });
    return Word.run(async (ctx) => {
      const state = await isEmptyDoc(ctx);
      const wantPage = asBool(a.page_setup) ?? state.empty;
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
      const bag = { plain: [] as Word.Paragraph[], tables: [] as Word.Table[] };
      const res = await (async () => {
        const built = await d.tracked(ctx, async () => {
          let anchor: Word.Paragraph;
          let where: "Before" | "After" = "After";
          if (a.position === "after_paragraph")
            anchor = (await d.resolvePara(ctx, a.paragraph_index, a.expect)).p;
          else if (state.empty) {
            anchor = ctx.document.body.paragraphs.getFirst();
            where = "Before";
          } else anchor = ctx.document.body.paragraphs.getLast();
          // Đã đo: chèn bảng ngay SAU một đoạn rỗng thì bảng nuốt đoạn đó và thứ tự bị đảo → gặp đoạn neo rỗng thì chèn TRƯỚC nó
          if (where === "After") {
            anchor.load("text");
            await ctx.sync();
            if (!anchor.text.trim()) where = "Before";
          }
          return build(ctx, plan, anchor, where, width, bag);
        });
        const res = await untracked(ctx, () => built.finish());
        return res;
      })().catch(async (e) => {
        // Lỗi giữa chừng: gỡ phần đã chèn dở (nếu không, model thử lại sẽ thành văn bản chèn đôi mà Hoàn tác không gỡ được)
        await untracked(ctx, async () => {
          bag.tables.forEach((t) => t.delete());
          await ctx.sync();
          bag.plain
            .slice()
            .reverse()
            .forEach((p) => p.delete());
          await ctx.sync();
        }).catch(() => undefined);
        throw e;
      });
      if (res.made.length) {
        // Hoàn tác = gỡ thẳng theo định danh (đứng trước bước trả lại thể thức trang trong danh sách, nên chạy sau nó)
        extra.push(removeMade(res.made));
        void inverse;
        void after;
        res.made.forEach((m) => touched.push(m.id));
      }
      const narrow =
        width < 150
          ? ` Lưu ý: vùng trình bày của trang chỉ rộng ${width} mm (chuẩn 160 mm), các bảng thể thức có thể tràn lề — đặt lại lề bằng word_set_page_format.`
          : "";
      return {
        ok: true,
        brief: adminDocBrief(plan),
        message: `Đã dựng văn bản đúng thể thức (ghi có theo dõi): phần đầu, ${plan.type === "cong_van" ? "V/v" : plan.type === "hop_dong" ? "tên hợp đồng" : "tên loại và trích yếu"}, nội dung ${plan.blocks.filter((b) => b.kind === "para").length} đoạn, ${plan.type === "hop_dong" ? "phần ký hai bên" : "phần ký và nơi nhận"}; ${res.rules} đường kẻ.${pageNote}${narrow}`,
        blanks: plan.blanks,
        ...(plan.notes.length ? { notes: plan.notes } : {}),
        ...(plan.foreign?.length
          ? {
              foreign_text: plan.foreign,
              fix_now: "Nội dung vừa dựng lẫn chữ không phải tiếng Việt (foreign_text). Trừ khi đó là tên riêng người dùng đưa, sửa NGAY từng cụm thành tiếng Việt bằng word_replace_text (find = đúng cụm chữ lạ) trước khi trả lời.",
            }
          : {}),
        next: "Báo người dùng các chỗ còn trống (blanks) cần điền trước khi ký. Muốn sửa câu chữ thì dùng word_set_paragraph_text trên đoạn tương ứng; không dựng lại cả văn bản. Phần đầu và phần ký là bảng không viền hai cột (các đoạn có trường cell khi đọc): bên trái và bên phải nằm ngang hàng nhau là đúng thể thức; bớt hay sửa một dòng ở đó thì chỉ xóa hoặc sửa đúng đoạn đó, không xóa rồi chèn lại để đổi chỗ.",
      };
    });
  }

  async function setPage(a: any, extra: UndoStep[]) {
    return Word.run(async (ctx) => {
      const pg = await applyPage(ctx, pageOptions(a || {}, asBool(a?.page_number) === true));
      if (pg.undo) extra.push(pg.undo);
      return {
        ok: true,
        changed: pg.changed,
        brief: pg.changed ? "đã đặt lại thể thức trang" : "thể thức trang đã đúng từ trước",
        message: `Thể thức trang: ${pg.note} Thay đổi này không ghi thành sửa đổi được theo dõi; hoàn tác bằng nút Hoàn tác của khung.`,
      };
    });
  }

  return {
    has: (name: string) => names.has(name),
    run: (
      name: string,
      a: any,
      u: { extra: UndoStep[]; inverse: UndoStep[]; touched: string[]; after?: UndoStep[] }
    ) =>
      name === "word_set_page_format"
        ? setPage(a, u.extra)
        : createDocument(a || {}, u.extra, u.inverse, u.touched, u.after || []),
    /** Cảnh báo rủi ro: đổi thể thức trang của một văn bản đang có nội dung khi dựng văn bản mới vào đó */
    risk: {
      // Đổi khổ giấy của văn bản nhiều phần hoặc có trang ngang: xoay trang, dồn lại bố cục — hỏi trước
      word_set_page_format: (a: any) =>
        Word.run(async (ctx) => {
          if (!d.supports("1.3", "WordApiDesktop")) return null;
          const secs = ctx.document.sections;
          secs.load("items");
          await ctx.sync();
          const setups = secs.items.slice(0, 30).map((s) => s.pageSetup);
          setups.forEach((p) => p.load("orientation"));
          await ctx.sync();
          const land = setups.filter((p) => String(p.orientation) === "Landscape").length;
          const s = await isEmptyDoc(ctx);
          if (s.empty) return null;
          const a4 = asBool(a?.a4) !== false;
          if (secs.items.length > 1 || (land && a4))
            return `Đặt lại lề${a4 ? " và khổ giấy A4 dọc" : ""} cho CẢ văn bản (${secs.items.length} phần${land ? `, trong đó ${land} phần đang là trang ngang${a4 ? " sẽ bị xoay dọc" : ""}` : ""}). Thay đổi này không hiện trong thẻ Review; gỡ bằng nút Hoàn tác của khung.`;
          return null;
        }),
      word_create_admin_document: (a: any) =>
        Word.run(async (ctx) => {
          if (asBool(a?.page_setup) !== true) return null;
          const s = await isEmptyDoc(ctx);
          return s.empty
            ? null
            : `Văn bản đang có ${s.paras} đoạn nội dung: thao tác này đổi khổ giấy, lề và đầu trang của CẢ văn bản, rồi chèn văn bản mới vào.`;
        }),
    } as Record<string, (a: any) => Promise<string | null> | string | null>,
  };
}
