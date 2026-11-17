/* global Word */
/** Đợt 1B — công cụ "soát và tra" của Word: soát nhất quán bằng mã, đọc phần ngoài thân văn bản (chú thích, đầu và chân trang,
 *  bookmark, trường, ô điền mẫu, hộp văn bản), kiểm thể thức trang, so sánh với phiên bản khác. Cả bốn đều CHỈ ĐỌC.
 *  Căn cứ kỹ thuật: bảng thăm dò trong tasks/dot-1b.md (đo trong Word ngày 15/11/2026). */
import { ToolSpec } from "../core/agent/types";
import { OutlineItem, ParaRow, buildOutline, clauseAt, clauseLabel } from "../core/outline";
import { ALL_CHECKS, IssueKind, checkConsistency } from "../core/consistency";
import { PageSetupInfo, ParaFont, VAN_BAN_HANH_CHINH, checkPageFormat, customStandard } from "../core/pageformat";
import { compareVersions } from "../core/compare";
import { findAttachedFile, listAttachedFiles } from "../core/attachedFiles";
import { RESULT_JSON_BUDGET, takeWithinBudget, truncate } from "./util";

export interface ReviewRow extends ParaRow {
  id?: string;
}

export interface ReviewDeps {
  supports(version: string, set?: string): boolean;
  /** Mọi đoạn của văn bản đang mở, kèm số do Word tự đánh */
  loadRows(ctx: Word.RequestContext): Promise<{ rows: ReviewRow[] }>;
  /** Mọi đoạn của một thân văn bản khác (tài liệu ẩn dựng từ tệp đính kèm) */
  loadRowsOf(ctx: Word.RequestContext, body: Word.Body): Promise<{ rows: ReviewRow[] }>;
}

const PARTS = ["footnotes", "endnotes", "headers_footers", "bookmarks", "fields", "content_controls", "text_boxes"] as const;
type Part = (typeof PARTS)[number];
const CC_TYPES = ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"];
/** Tên tác giả của các sửa đổi trong bản so sánh do Word dựng */
export const COMPARE_AUTHOR = "So sánh (AI Đa Nguồn)";

export const WORD_REVIEW_SPECS: ToolSpec[] = [
  {
    name: "word_check_consistency",
    description:
      "Run the built-in consistency check on the whole document (done by code, instant, no reading needed): numbering gaps or duplicates; cross-references to clauses that do not exist; defined terms written inconsistently or capitalised terms that are never defined; a party called by several names; blanks left unfilled; amounts where the figure and the words disagree. Returns a list of points to look at, each with paragraph index, clause number and an excerpt. Use it FIRST when asked to check, proofread or review a contract before signing; then read only the flagged places. The checks are pattern-based: confirm each point before reporting it as an error.",
    parameters: {
      type: "object",
      properties: {
        checks: {
          type: "array",
          items: { type: "string", enum: ALL_CHECKS },
          description: "Which checks to run. Default: all (numbering, reference, term, party, blank, amount).",
        },
        start: { type: "number", description: "Skip the first N points (paging, from next_start). Default 0." },
      },
    },
  },
  {
    name: "word_read_parts",
    description:
      "Read the parts of the document that are NOT in the body paragraphs: footnotes and endnotes (with the paragraph and clause they belong to), headers and footers, bookmarks, fields such as cross-references (code and current result), fill-in fields (content controls, with empty ones flagged) and text boxes. Call it when the user asks about any of these, and when reviewing a whole document.",
    parameters: {
      type: "object",
      properties: {
        parts: { type: "array", items: { type: "string", enum: PARTS as unknown as string[] }, description: "Which parts to read. Default: all." },
        include_hidden_bookmarks: { type: "boolean", description: "Also list hidden bookmarks (names starting with _, created by Word for cross-references). Default false." },
        start: { type: "number", description: "Skip the first N items of each requested list (paging: when a list comes back with truncated=true, call again with parts set to that one part and start=next_start)." },
      },
    },
  },
  {
    name: "word_check_page_format",
    description:
      "Check the page layout against a formatting standard: paper size and orientation, the four margins, fonts and font sizes. Default standard: Vietnamese administrative documents such as official letters (công văn) and submissions (tờ trình): A4 portrait, top and bottom margins 20–25 mm, left 30–35 mm, right 15–20 mm, Times New Roman, body size 13–14. Pass your own numbers to check against another standard. Returns the measured values and exactly which items deviate.",
    parameters: {
      type: "object",
      properties: {
        top_mm: { type: "number", description: "Required top margin in mm (overrides the default standard)." },
        bottom_mm: { type: "number" },
        left_mm: { type: "number" },
        right_mm: { type: "number" },
        font: { type: "string", description: "Required font name (overrides the default standard)." },
        font_size: { type: "number", description: "Required body font size in points (overrides the default standard)." },
      },
    },
  },
  {
    name: "word_compare_versions",
    description:
      "Compare the open document with another version of it: a .docx file the user attached in this chat. Returns every difference grouped by clause number (changed paragraphs with the exact words that differ, paragraphs only in the open document, paragraphs only in the other version, and clauses whose number shifted). The open document is NOT modified. Set open_redline=true only when the user wants to see Word's own redline: it opens the comparison in a NEW Word window and leaves the open document untouched.",
    parameters: {
      type: "object",
      properties: {
        file: { type: "string", description: "Name of the attached .docx to compare with. Default: the most recently attached .docx." },
        open_redline: { type: "boolean", description: "Also open Word's redline of the two versions in a new window. Default false." },
        start: { type: "number", description: "Skip the first N differences (paging, from next_start). Default 0." },
      },
    },
  },
];

const KIND_VI: Record<IssueKind, string> = {
  numbering: "đánh số",
  reference: "dẫn chiếu",
  term: "thuật ngữ",
  party: "tên các bên",
  blank: "chỗ trống",
  amount: "số và chữ",
};

/** Bỏ ký tự mốc mà Word đặt ở đầu chữ của chú thích */
const noteText = (s: any) =>
  String(s ?? "")
    .replace(/[\u0001\u0002]/g, "")
    .replace(/\r/g, "\n")
    .trim();

export function createWordReviewTools(deps: ReviewDeps) {
  const { supports } = deps;

  /** Chỉ số đoạn và số điều của một loạt đoạn neo (theo định danh đoạn; không có thì so chữ) */
  const locator = (rows: ReviewRow[], outline: OutlineItem[]) => {
    const byId = new Map<string, number>();
    rows.forEach((r) => r.id && !byId.has(r.id) && byId.set(r.id, r.i));
    return (p: Word.Paragraph): { i?: number; clause?: string } => {
      let i = -1;
      try {
        if (supports("1.6") && p.uniqueLocalId && byId.has(p.uniqueLocalId)) i = byId.get(p.uniqueLocalId) as number;
      } catch {
        /* chưa nạp định danh */
      }
      if (i < 0) {
        const t = noteText(p.text);
        if (t) i = rows.findIndex((r) => r.text === t);
      }
      if (i < 0) return {};
      const c = clauseLabel(outline, clauseAt(outline, i));
      return { i, ...(c ? { clause: c } : {}) };
    };
  };
  const anchorProps = () => (supports("1.6") ? "text,uniqueLocalId" : "text");

  async function checkConsistencyTool(a: any) {
    return Word.run(async (ctx) => {
      const { rows } = await deps.loadRows(ctx);
      const outline = buildOutline(rows);
      const asked: string[] = Array.isArray(a.checks) ? a.checks.map(String) : [];
      const checks = ALL_CHECKS.filter((k) => !asked.length || asked.includes(k));
      if (!checks.length) throw new Error(`checks không hợp lệ. Dùng: ${ALL_CHECKS.join(", ")}.`);
      const res = checkConsistency(rows, { checks, outline });
      const issues = res.issues.map((x) => ({ kind: x.kind, i: x.p, ...(x.clause ? { clause: x.clause } : {}), message: x.message, ...(x.excerpt ? { excerpt: x.excerpt } : {}), ...(x.also?.length ? { also: x.also } : {}) }));
      // Ô điền mẫu còn trống cũng là chỗ trống chưa điền
      if (checks.includes("blank") && supports("1.5")) {
        try {
          const cc = ctx.document.getContentControls({ types: CC_TYPES as any });
          cc.load("items/title,items/tag,items/text,items/type");
          await ctx.sync();
          const empty = cc.items.filter((c) => c.type !== "CheckBox" && !String(c.text || "").trim());
          if (empty.length) {
            const ps = empty.map((c) => {
              const p = c.paragraphs.getFirstOrNullObject();
              p.load(anchorProps());
              return p;
            });
            await ctx.sync();
            const at = locator(rows, outline);
            empty.forEach((c, k) => {
              const w = ps[k].isNullObject ? {} : at(ps[k]);
              issues.push({ kind: "blank", i: w.i ?? -1, ...(w.clause ? { clause: w.clause } : {}), message: `Ô điền mẫu “${c.title || c.tag || "(không tên)"}” chưa điền.` });
            });
          }
        } catch {
          /* không đọc được ô điền mẫu: bỏ qua phần này */
        }
      }
      const byKind: Record<string, number> = {};
      issues.forEach((x) => (byKind[KIND_VI[x.kind as IssueKind]] = (byKind[KIND_VI[x.kind as IssueKind]] || 0) + 1));
      const from = Math.max(0, Math.floor(Number(a.start)) || 0);
      const { picked, truncated } = takeWithinBudget(issues.slice(from), RESULT_JSON_BUDGET - 1800);
      const next = from + picked.length;
      return {
        ok: true,
        brief: issues.length
          ? `${issues.length} điểm cần xem: ${Object.entries(byKind)
              .map(([k, n]) => `${k} ${n}`)
              .join(", ")}`
          : "không thấy điểm nào",
        total: issues.length,
        by_kind: byKind,
        checks_run: checks,
        ...(next < issues.length ? { next_start: next } : {}),
        ...(truncated ? { truncated: true } : {}),
        defined_terms: res.terms.slice(0, 40),
        parties: res.parties,
        note: issues.length
          ? "Đây là các điểm CẦN XEM do mã dò theo mẫu, chưa chắc đều là lỗi: đọc chỗ được chỉ rồi mới kết luận. i = số đoạn; kind = loại."
          : "Không thấy điểm nào theo các phép soát đã chạy. Phép soát theo mẫu không thay cho việc đọc nội dung.",
        issues: picked,
      };
    });
  }

  async function readParts(a: any) {
    const asked: string[] = Array.isArray(a.parts) ? a.parts.map(String) : [];
    const want = (p: Part) => !asked.length || asked.includes(p);
    if (asked.length && !asked.some((p) => (PARTS as readonly string[]).includes(p))) throw new Error(`parts không hợp lệ. Dùng: ${PARTS.join(", ")}.`);
    const out: any = { ok: true };
    const fail = (e: any) => ({ unavailable: `Word trên máy này không đọc được phần này (${truncate(e?.message || String(e), 120)}).` });
    let rows: ReviewRow[] = [];
    let outline: OutlineItem[] = [];
    await Word.run(async (ctx) => {
      rows = (await deps.loadRows(ctx)).rows;
      outline = buildOutline(rows);
    });
    const at = locator(rows, outline);

    for (const kind of ["footnotes", "endnotes"] as const) {
      if (!want(kind)) continue;
      if (!supports("1.5")) {
        out[kind] = { unavailable: "Cần WordApi 1.5 để đọc chú thích." };
        continue;
      }
      try {
        out[kind] = await Word.run(async (ctx) => {
          const notes = kind === "footnotes" ? ctx.document.body.footnotes : ctx.document.body.endnotes;
          notes.load("items/body/text");
          await ctx.sync();
          const ps = notes.items.map((n) => {
            const p = n.reference.paragraphs.getFirstOrNullObject();
            p.load(anchorProps());
            return p;
          });
          await ctx.sync();
          return notes.items.map((n, k) => ({ n: k + 1, ...(ps[k].isNullObject ? {} : at(ps[k])), text: truncate(noteText(n.body.text), 1500) }));
        });
      } catch (e) {
        out[kind] = fail(e);
      }
    }

    if (want("headers_footers")) {
      try {
        out.headers_footers = await Word.run(async (ctx) => {
          const secs = ctx.document.sections;
          secs.load("items");
          await ctx.sync();
          const TYPES: [Word.HeaderFooterType | "Primary" | "FirstPage" | "EvenPages", string][] = [
            ["Primary", "các trang"],
            ["FirstPage", "trang đầu"],
            ["EvenPages", "trang chẵn"],
          ];
          const list = secs.items.slice(0, 12).flatMap((s, k) =>
            TYPES.map(([t, label]) => {
              const h = s.getHeader(t as any);
              const f = s.getFooter(t as any);
              h.load("text");
              f.load("text");
              if (supports("1.5")) {
                h.fields.load("items/type");
                f.fields.load("items/type");
              }
              return { section: k + 1, pages: label, h, f };
            })
          );
          await ctx.sync();
          const seen = new Set<string>();
          const res: any[] = [];
          for (const x of list) {
            const header = noteText(x.h.text);
            const footer = noteText(x.f.text);
            if (!header && !footer) continue;
            const key = `${x.pages}|${header}|${footer}`;
            if (seen.has(key)) continue;
            seen.add(key);
            const auto = supports("1.5") ? Array.from(new Set([...x.h.fields.items, ...x.f.fields.items].map((fd) => String(fd.type)))) : [];
            res.push({
              section: x.section,
              pages: x.pages,
              ...(header ? { header: truncate(header, 1200) } : {}),
              ...(footer ? { footer: truncate(footer, 1200) } : {}),
              // Trường tự động trong đầu / chân trang (Page = số trang): con số đọc được là giá trị của một trang, Word tự đổi theo trang
              ...(auto.length ? { auto_fields: auto } : {}),
            });
          }
          return { sections: secs.items.length, items: res, ...(res.length ? {} : { note: "Văn bản không có đầu trang hay chân trang." }) };
        });
      } catch (e) {
        out.headers_footers = fail(e);
      }
    }

    if (want("bookmarks")) {
      if (!supports("1.4")) out.bookmarks = { unavailable: "Cần WordApi 1.4 để đọc bookmark." };
      else {
        try {
          out.bookmarks = await Word.run(async (ctx) => {
            const whole = ctx.document.body.getRange("Whole");
            const all = whole.getBookmarks(true, false);
            await ctx.sync();
            const names = (all.value || []).slice().sort();
            const hidden = names.filter((n) => n.startsWith("_"));
            const shown = (a.include_hidden_bookmarks ? names : names.filter((n) => !n.startsWith("_"))).slice(0, 400);
            const rs = shown.map((n) => {
              const r = ctx.document.getBookmarkRangeOrNullObject(n);
              r.load("text");
              const p = r.paragraphs.getFirstOrNullObject();
              p.load(anchorProps());
              return { r, p };
            });
            await ctx.sync();
            return {
              total: names.length,
              hidden: hidden.length,
              items: shown.map((n, k) => ({ name: n, ...(rs[k].r.isNullObject || rs[k].p.isNullObject ? {} : at(rs[k].p)), text: rs[k].r.isNullObject ? "" : truncate(noteText(rs[k].r.text), 200) })),
              ...(hidden.length && !a.include_hidden_bookmarks ? { note: `${hidden.length} bookmark ẩn (Word tạo cho dẫn chiếu chéo) chưa liệt kê; cần thì gọi lại với include_hidden_bookmarks=true.` } : {}),
            };
          });
        } catch (e) {
          out.bookmarks = fail(e);
        }
      }
    }

    if (want("fields")) {
      if (!supports("1.4")) out.fields = { unavailable: "Cần WordApi 1.4 để đọc trường." };
      else {
        try {
          out.fields = await Word.run(async (ctx) => {
            const fs = ctx.document.body.fields;
            fs.load(supports("1.5") ? "items/code,items/type,items/result/text" : "items/code,items/result/text");
            await ctx.sync();
            const items = fs.items.slice(0, 600);
            const ps = items.map((f) => {
              const p = f.result.paragraphs.getFirstOrNullObject();
              p.load(anchorProps());
              return p;
            });
            await ctx.sync();
            return {
              total: fs.items.length,
              items: items.map((f, k) => {
                const result = noteText(f.result.text);
                return {
                  ...(supports("1.5") ? { type: String(f.type) } : {}),
                  code: truncate(noteText(f.code), 120),
                  result: truncate(result, 160),
                  ...(ps[k].isNullObject ? {} : at(ps[k])),
                  ...(/^(Error!|Lỗi!)/.test(result) ? { broken: true } : {}),
                };
              }),
            };
          });
        } catch (e) {
          out.fields = fail(e);
        }
      }
    }

    if (want("content_controls")) {
      if (!supports("1.5")) out.content_controls = { unavailable: "Cần WordApi 1.5 để đọc ô điền mẫu." };
      else {
        try {
          out.content_controls = await Word.run(async (ctx) => {
            // Thăm dò 15/11/2026: document.contentControls chỉ trả ô kiểu RichText; phải nêu rõ các loại
            const cc = ctx.document.getContentControls({ types: CC_TYPES as any });
            cc.load("items/title,items/tag,items/text,items/type");
            await ctx.sync();
            const items = cc.items.slice(0, 600);
            const ps = items.map((c) => {
              const p = c.paragraphs.getFirstOrNullObject();
              p.load(anchorProps());
              return p;
            });
            await ctx.sync();
            return {
              total: cc.items.length,
              empty: cc.items.filter((c) => c.type !== "CheckBox" && !String(c.text || "").trim()).length,
              items: items.map((c, k) => {
                const text = noteText(c.text);
                return { title: c.title || "", ...(c.tag ? { tag: c.tag } : {}), type: String(c.type), text: truncate(text, 300), ...(c.type !== "CheckBox" && !text ? { empty: true } : {}), ...(ps[k].isNullObject ? {} : at(ps[k])) };
              }),
            };
          });
        } catch (e) {
          out.content_controls = fail(e);
        }
      }
    }

    if (want("text_boxes")) {
      if (!supports("1.2", "WordApiDesktop")) out.text_boxes = { unavailable: "Cần WordApiDesktop 1.2 để đọc hộp văn bản." };
      else {
        try {
          out.text_boxes = await Word.run(async (ctx) => {
            const sh = ctx.document.body.shapes;
            sh.load("items/id,items/name,items/type");
            await ctx.sync();
            const boxes = sh.items.filter((s) => s.type === "TextBox").slice(0, 200);
            boxes.forEach((s) => s.body.load("text"));
            await ctx.sync();
            return {
              total: boxes.length,
              other_shapes: sh.items.length - sh.items.filter((s) => s.type === "TextBox").length,
              items: boxes.map((s) => ({ name: s.name, text: truncate(noteText(s.body.text), 1200) })),
            };
          });
        } catch (e) {
          out.text_boxes = fail(e);
        }
      }
    }

    // Một dòng cho khung chat: đọc được những gì
    const n = (v: any) => (Array.isArray(v) ? v.length : typeof v?.total === "number" ? v.total : Array.isArray(v?.items) ? v.items.length : -1);
    const LABEL: [Part, string][] = [
      ["footnotes", "chú thích cuối trang"],
      ["endnotes", "chú thích cuối văn bản"],
      ["headers_footers", "đầu và chân trang"],
      ["bookmarks", "bookmark"],
      ["fields", "trường"],
      ["content_controls", "ô điền mẫu"],
      ["text_boxes", "hộp văn bản"],
    ];
    out.brief = LABEL.filter(([k]) => out[k] !== undefined)
      .map(([k, label]) => (out[k]?.unavailable ? `${label}: không đọc được` : `${k === "headers_footers" ? out[k].items.length : n(out[k])} ${label}`))
      .join(", ");

    // Phân trang: bỏ N mục đầu của mỗi danh sách (dùng khi lần gọi trước trả truncated cho một phần)
    const skip = Math.max(0, Math.floor(Number(a.start)) || 0);
    if (skip) {
      for (const [k, v] of Object.entries(out)) {
        if (Array.isArray(v)) out[k] = { total: v.length, start: skip, items: v.slice(skip) };
        else if (Array.isArray((v as any)?.items) && k !== "headers_footers") out[k] = { ...(v as any), start: skip, items: (v as any).items.slice(skip) };
      }
    }
    // Kết quả quá dài: rút gọn danh sách dài nhất cho tới khi vừa ngân sách
    for (let guard = 0; guard < 40 && JSON.stringify(out).length > RESULT_JSON_BUDGET; guard++) {
      let longest = "";
      let size = 0;
      for (const [k, v] of Object.entries(out)) {
        const n = Array.isArray(v) ? v.length : Array.isArray((v as any)?.items) ? (v as any).items.length : 0;
        if (n > size) {
          size = n;
          longest = k;
        }
      }
      if (!longest || size <= 3) break;
      const keep = Math.max(3, Math.floor(size * 0.7));
      const hint = `Danh sách dài: gọi lại với parts = ["${longest}"] và start = next_start để đọc tiếp.`;
      if (Array.isArray(out[longest])) out[longest] = { shown: keep, total: size, truncated: true, next_start: skip + keep, note: hint, items: out[longest].slice(0, keep) };
      else out[longest] = { ...out[longest], shown: keep, truncated: true, next_start: skip + keep, note: hint, items: out[longest].items.slice(0, keep) };
    }
    return out;
  }

  async function checkPageFormatTool(a: any) {
    const std = [a.top_mm, a.bottom_mm, a.left_mm, a.right_mm, a.font, a.font_size].some((v) => v !== undefined && v !== null && v !== "") ? customStandard(VAN_BAN_HANH_CHINH, a) : VAN_BAN_HANH_CHINH;
    return Word.run(async (ctx) => {
      const sections: PageSetupInfo[] = [];
      let marginNote = "";
      if (supports("1.3", "WordApiDesktop")) {
        const secs = ctx.document.sections;
        secs.load("items");
        await ctx.sync();
        const setups = secs.items.slice(0, 30).map((s) => {
          s.pageSetup.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight,paperSize,orientation");
          return s.pageSetup;
        });
        await ctx.sync();
        setups.forEach((p, k) =>
          sections.push({ section: k + 1, topMargin: p.topMargin, bottomMargin: p.bottomMargin, leftMargin: p.leftMargin, rightMargin: p.rightMargin, pageWidth: p.pageWidth, pageHeight: p.pageHeight, paperSize: String(p.paperSize), orientation: String(p.orientation) })
        );
      } else marginNote = "Word trên máy này không cho add-in đọc lề và khổ giấy (cần WordApiDesktop 1.3): chỉ kiểm được phông và cỡ chữ.";
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text,items/font/name,items/font/size");
      await ctx.sync();
      const fonts: ParaFont[] = ps.items.map((p, i) => ({ i, name: String(p.font.name || ""), size: typeof p.font.size === "number" && p.font.size > 0 ? p.font.size : null, chars: noteText(p.text).length }));
      const rep = checkPageFormat(sections, fonts, std);
      const line = (x: (typeof rep.issues)[number]) => `${x.label}${x.section ? ` (phần ${x.section})` : ""}: ${x.actual} — chuẩn ${x.expected}${x.p !== undefined ? ` [¶${x.p}]` : ""}`;
      return {
        ok: true,
        brief: rep.issues.length ? `${rep.issues.length} điểm lệch chuẩn: ${rep.issues.map((x) => `${x.label.toLowerCase()} ${x.actual}`).join(", ")}` : marginNote ? "không đọc được lề" : "đúng chuẩn",
        standard: rep.standard,
        conforms: rep.ok && !marginNote,
        problems: rep.issues.length,
        message: rep.issues.length ? `Có ${rep.issues.length} điểm lệch chuẩn: ${rep.issues.map(line).join("; ")}.` : marginNote || "Thể thức trang đúng chuẩn: khổ giấy, bốn lề, phông và cỡ chữ.",
        ...(marginNote ? { note: marginNote } : {}),
        issues: rep.issues.map((x) => ({ what: x.what, label: x.label, actual: x.actual, expected: x.expected, ...(x.section ? { section: x.section } : {}), ...(x.p !== undefined ? { i: x.p } : {}), ...(x.also?.length ? { also: x.also } : {}) })),
        sections: rep.sections,
        fonts: rep.fonts,
      };
    });
  }

  async function compareTool(a: any) {
    const f = findAttachedFile(a.file, ".docx");
    if (!f) {
      const have = listAttachedFiles(".docx");
      throw new Error(
        have.length
          ? `Không thấy tệp đính kèm tên “${truncate(a.file, 80)}”. Các tệp .docx đã đính kèm trong phiên này: ${have.map((n) => `“${n}”`).join(", ")}.`
          : "Chưa có tệp .docx nào được đính kèm trong phiên làm việc này (tải lại khung add-in thì tệp đã đính kèm không còn). Nhờ người dùng bấm dấu + cạnh ô nhập → Đính kèm tệp, chọn phiên bản .docx cần so, rồi gửi lại yêu cầu."
      );
    }
    if (!supports("1.3", "WordApiHiddenDocument")) {
      throw new Error("Word trên máy này không cho add-in đọc một tệp .docx khác (cần WordApiHiddenDocument 1.3). Nội dung tệp đính kèm đã nằm trong tin nhắn: hãy so theo chữ ở đó và nói rõ đây là so thủ công.");
    }
    return Word.run(async (ctx) => {
      const cur = await deps.loadRows(ctx);
      // Tài liệu ẩn: Word tự đọc tệp kia (đủ số điều do Word đánh, bản hiện tại nếu tệp có sửa đổi) mà không mở ra và không đụng
      // tới văn bản đang mở (thăm dò 15/11/2026)
      const doc = ctx.application.createDocument(f.data);
      const other = await deps.loadRowsOf(ctx, doc.body);
      const outlineCurrent = buildOutline(cur.rows);
      // from / to (không công bố cho model): chỉ so một khúc của văn bản đang mở — dùng cho bước Tự kiểm tra trong vùng thử
      const lo = Number.isFinite(Number(a.from)) && a.from !== undefined && a.from !== null ? Number(a.from) : -Infinity;
      const hi = Number.isFinite(Number(a.to)) && a.to !== undefined && a.to !== null ? Number(a.to) : Infinity;
      const res = compareVersions(
        cur.rows.filter((r) => r.i >= lo && r.i <= hi),
        other.rows,
        { outlineCurrent }
      );
      let redline: string | undefined;
      if (a.open_redline === true) {
        if (!supports("1.2", "WordApiDesktop")) redline = "Word trên máy này không dựng được bản so sánh qua add-in (cần WordApiDesktop 1.2).";
        else {
          try {
            // Đích là tài liệu MỚI: ghi vào tài liệu đang mở sẽ làm mất các sửa đổi đang chờ duyệt (thăm dò 15/11/2026)
            ctx.document.compareFromBase64(f.data, { compareTarget: "CompareTargetNew", authorName: COMPARE_AUTHOR, addToRecentFiles: false, detectFormatChanges: false, ignoreAllComparisonWarnings: true } as any);
            await ctx.sync();
            redline = "Đã mở bản so sánh (redline) của Word trong một cửa sổ mới; văn bản đang mở không đổi.";
          } catch (e: any) {
            redline = `Không mở được bản so sánh của Word: ${truncate(e?.message || String(e), 160)}`;
          }
        }
      }
      const TYPE_VI = { changed: "sửa", only_current: "chỉ có ở bản đang mở", only_other: "chỉ có ở phiên bản khác" } as const;
      const list = res.diffs.map((d) => ({
        type: d.type,
        clause: d.clause || "(ngoài điều khoản)",
        ...(d.clause_other ? { clause_other: d.clause_other } : {}),
        ...(d.p !== undefined ? { i: d.p } : {}),
        ...(d.type === "changed" ? { changes: d.changes } : {}),
        ...(d.type === "only_current" ? { current: truncate(d.current, 300) } : {}),
        ...(d.type === "only_other" ? { other: truncate(d.other, 300) } : {}),
      }));
      const from = Math.max(0, Math.floor(Number(a.start)) || 0);
      const { picked, truncated } = takeWithinBudget(list.slice(from), RESULT_JSON_BUDGET - 1500);
      const next = from + picked.length;
      const count = (t: keyof typeof TYPE_VI) => res.diffs.filter((d) => d.type === t).length;
      return {
        ok: true,
        brief: res.identical ? "hai phiên bản giống hệt nhau" : `${res.diffs.length} chỗ khác (${count("changed")} sửa, ${count("only_current")} chỉ ở bản đang mở, ${count("only_other")} chỉ ở phiên bản khác); văn bản đang mở không đổi`,
        file: f.name,
        open_document_modified: false,
        paragraphs_current: res.paragraphs_current,
        paragraphs_other: res.paragraphs_other,
        differences: res.diffs.length,
        summary: res.identical ? "Hai phiên bản giống hệt nhau về chữ." : `${res.diffs.length} chỗ khác: ${count("changed")} đoạn bị sửa, ${count("only_current")} đoạn chỉ có ở bản đang mở, ${count("only_other")} đoạn chỉ có ở phiên bản khác.`,
        ...(next < list.length ? { next_start: next } : {}),
        ...(truncated ? { truncated: true } : {}),
        ...(redline ? { redline } : {}),
        legend:
          "type: changed = đoạn có ở cả hai bản nhưng chữ khác (changes: “chữ ở bản đang mở” → “chữ ở phiên bản khác”); only_current = đoạn chỉ có ở bản đang mở; only_other = đoạn chỉ có ở phiên bản khác (clause = điều khoản đứng ngay trước chỗ đó trong bản đang mở, clause_other = số của nó trong phiên bản khác). i = số đoạn trong bản đang mở.",
        ...(res.renumbered.length ? { renumbered: res.renumbered.slice(0, 40).map((x) => ({ current: x.current, other: x.other, i: x.p })), renumbered_note: "Các điều khoản giữ nguyên chữ nhưng mang số khác ở hai bản (do thêm hoặc bớt khoản)." } : {}),
        diffs: picked,
      };
    });
  }

  return {
    specs: WORD_REVIEW_SPECS,
    has: (name: string) => WORD_REVIEW_SPECS.some((s) => s.name === name),
    run(name: string, a: any): Promise<any> {
      if (name === "word_check_consistency") return checkConsistencyTool(a || {});
      if (name === "word_read_parts") return readParts(a || {});
      if (name === "word_check_page_format") return checkPageFormatTool(a || {});
      if (name === "word_compare_versions") return compareTool(a || {});
      throw new Error(`Không có công cụ ${name}`);
    },
  };
}
