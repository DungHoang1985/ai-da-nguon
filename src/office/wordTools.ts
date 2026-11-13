/* global Word, Office */
/** Công cụ để AI thao tác trực tiếp trên văn bản Word. Sửa đổi được ghi bằng Theo dõi thay đổi (Track Changes). */
import { HostToolSet, ToolSpec, UndoEntry } from "../core/agent/types";
import { diffMiddle, escapeWordSearch, normalize2D, parseMarkdownLine, truncate } from "./util";

const MAX_READ_CHARS = 30000;

const STYLE_ENUM = ["Normal", "Heading1", "Heading2", "Heading3", "Title", "ListBullet", "ListNumber", "Quote"];

export const WORD_SPECS: ToolSpec[] = [
  {
    name: "word_get_overview",
    description:
      "Get a document overview: paragraph count, headings with their paragraph index, first paragraphs, current selection, track-changes mode, comment count.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "word_read_paragraphs",
    description: "Read paragraphs by index range (0-based, end exclusive) with their style. Or read the current selection.",
    parameters: {
      type: "object",
      properties: {
        start: { type: "number", description: "First paragraph index (default 0)." },
        end: { type: "number", description: "End index, exclusive (default start+80)." },
        selection_only: { type: "boolean", description: "Read only the current selection." },
      },
    },
  },
  {
    name: "word_replace_text",
    description:
      "Find text in the document and replace it (tracked change). Best for small edits of words or phrases. The search text must be at most 255 characters.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string" },
        replace: { type: "string" },
        match_case: { type: "boolean" },
        whole_word: { type: "boolean" },
        max_replacements: { type: "number", description: "Default 50." },
      },
      required: ["find", "replace"],
    },
    write: true,
  },
  {
    name: "word_set_paragraph_text",
    description:
      "Rewrite one paragraph with new text (tracked change). Only the changed part is replaced, so redlines stay small. Paragraph formatting is kept.",
    parameters: {
      type: "object",
      properties: { index: { type: "number", description: "Paragraph index (0-based)." }, text: { type: "string" } },
      required: ["index", "text"],
    },
    write: true,
  },
  {
    name: "word_insert_paragraphs",
    description:
      "Insert one or more paragraphs (tracked). Each line becomes a paragraph; lines starting with '# ', '## ', '### ' become headings, '- ' bullets, '1. ' numbered items.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string" },
        position: {
          type: "string",
          enum: ["end", "start", "after_paragraph", "before_paragraph", "after_selection"],
          description: "Where to insert (default end).",
        },
        paragraph_index: { type: "number", description: "Needed for after_paragraph / before_paragraph." },
        style: { type: "string", enum: STYLE_ENUM, description: "Force one style for all lines (optional)." },
      },
      required: ["text"],
    },
    write: true,
  },
  {
    name: "word_replace_selection",
    description: "Replace the user's current selection (or insert at the cursor) with text (tracked).",
    parameters: { type: "object", properties: { text: { type: "string" } }, required: ["text"] },
    write: true,
  },
  {
    name: "word_format_text",
    description:
      "Format text: either every occurrence of `find`, a paragraph by index, or the current selection. Font options apply to text; alignment and style apply to whole paragraphs.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string" },
        paragraph_index: { type: "number" },
        use_selection: { type: "boolean" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_size: { type: "number" },
        font_name: { type: "string" },
        font_color: { type: "string", description: "Hex colour like #C00000." },
        highlight_color: { type: "string", description: "Hex colour like #FFFF00, or 'none'." },
        alignment: { type: "string", enum: ["Left", "Centered", "Right", "Justified"] },
        style: { type: "string", enum: STYLE_ENUM },
      },
    },
    write: true,
  },
  {
    name: "word_add_comment",
    description:
      "Add a review comment anchored to text (first occurrence of `find`), to a paragraph, or to the selection. Use for review notes and legal risk remarks.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string", description: "Comment content." },
        find: { type: "string", description: "Anchor text (max 255 characters)." },
        paragraph_index: { type: "number" },
        use_selection: { type: "boolean" },
      },
      required: ["text"],
    },
    write: true,
  },
  {
    name: "word_insert_table",
    description: "Insert a table (tracked) after a paragraph, after the selection, or at the end. First row is the header.",
    parameters: {
      type: "object",
      properties: {
        values: { type: "array", items: { type: "array", items: { type: "string" } } },
        position: { type: "string", enum: ["end", "after_paragraph", "after_selection"] },
        paragraph_index: { type: "number" },
      },
      required: ["values"],
    },
    write: true,
  },
  {
    name: "word_delete_paragraphs",
    description: "Delete paragraphs from start to end index inclusive (tracked deletion). Only when the user asks to remove content.",
    parameters: {
      type: "object",
      properties: { start: { type: "number" }, end: { type: "number" } },
      required: ["start"],
    },
    write: true,
  },
];

export interface WordSettings {
  trackChanges: boolean;
}

function supports(v: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("WordApi", v);
  } catch {
    return false;
  }
}

function mapStyle(s?: string): string {
  const k = String(s || "")
    .replace(/\s+/g, "")
    .toLowerCase();
  const found = STYLE_ENUM.find((x) => x.toLowerCase() === k);
  if (found) return found;
  if (/^h(eading)?1$/.test(k)) return "Heading1";
  if (/^h(eading)?2$/.test(k)) return "Heading2";
  if (/^h(eading)?3$/.test(k)) return "Heading3";
  if (/bullet/.test(k)) return "ListBullet";
  if (/number/.test(k)) return "ListNumber";
  return "Normal";
}

async function loadParagraphs(ctx: Word.RequestContext, props = "items/text") {
  const paras = ctx.document.body.paragraphs;
  paras.load(props);
  await ctx.sync();
  return paras;
}

async function paragraphAt(ctx: Word.RequestContext, index: number) {
  const paras = await loadParagraphs(ctx);
  const i = Number(index);
  if (!Number.isInteger(i) || i < 0 || i >= paras.items.length) {
    throw new Error(`Không có đoạn số ${index}. Văn bản có ${paras.items.length} đoạn (đánh số từ 0).`);
  }
  return paras.items[i];
}

export function createWordTools(getSettings: () => WordSettings): HostToolSet {
  /** Bật Theo dõi thay đổi trong lúc AI sửa, xong trả lại chế độ cũ. */
  async function tracked<T>(ctx: Word.RequestContext, fn: () => Promise<T>): Promise<T> {
    if (!getSettings().trackChanges || !supports("1.4")) return fn();
    const doc = ctx.document;
    doc.load("changeTrackingMode");
    await ctx.sync();
    const prev = doc.changeTrackingMode;
    if (prev !== "TrackAll") {
      doc.changeTrackingMode = "TrackAll" as any;
      await ctx.sync();
    }
    try {
      return await fn();
    } finally {
      if (prev !== "TrackAll") {
        doc.changeTrackingMode = prev;
        await ctx.sync();
      }
    }
  }

  async function overview() {
    return Word.run(async (ctx) => {
      const paras = await loadParagraphs(ctx, "items/text,items/style,items/styleBuiltIn");
      const sel = ctx.document.getSelection();
      sel.load("text");
      let mode: string | null = null;
      let comments: number | null = null;
      if (supports("1.4")) {
        ctx.document.load("changeTrackingMode");
        const cs = ctx.document.body.getComments();
        cs.load("items");
        await ctx.sync();
        mode = ctx.document.changeTrackingMode;
        comments = cs.items.length;
      } else {
        await ctx.sync();
      }
      const items = paras.items;
      const headings = items
        .map((p, i) => ({ index: i, style: p.styleBuiltIn !== "Other" ? p.styleBuiltIn : p.style, text: truncate(p.text, 120) }))
        .filter((h) => /heading|title|tiêu đề/i.test(String(h.style)) && h.text.trim())
        .slice(0, 80);
      return {
        paragraphs: items.length,
        characters: items.reduce((n, p) => n + p.text.length, 0),
        headings,
        first_paragraphs: items.slice(0, 6).map((p, i) => ({ index: i, style: p.style, text: truncate(p.text, 200) })),
        selection: truncate(sel.text, 1500),
        track_changes: mode,
        comments,
      };
    });
  }

  async function targetRanges(ctx: Word.RequestContext, a: any): Promise<{ ranges: Word.Range[]; paras: Word.Paragraph[] }> {
    if (a.find) {
      if (String(a.find).length > 255) throw new Error("Chuỗi tìm quá dài (tối đa 255 ký tự).");
      const res = ctx.document.body.search(escapeWordSearch(a.find), { matchCase: true });
      res.load("items");
      await ctx.sync();
      if (!res.items.length) throw new Error(`Không tìm thấy "${truncate(a.find, 80)}".`);
      const ps = res.items.map((r) => r.paragraphs.getFirst());
      return { ranges: res.items.slice(0, 50), paras: ps.slice(0, 50) };
    }
    if (a.paragraph_index !== undefined && a.paragraph_index !== null) {
      const p = await paragraphAt(ctx, a.paragraph_index);
      return { ranges: [p.getRange("Content")], paras: [p] };
    }
    const sel = ctx.document.getSelection();
    const ps = sel.paragraphs;
    ps.load("items");
    await ctx.sync();
    return { ranges: [sel], paras: ps.items };
  }

  async function execute(name: string, a: any, _ctx: { log: UndoEntry[] }): Promise<any> {
    switch (name) {
      case "word_get_overview":
        return overview();

      case "word_read_paragraphs":
        return Word.run(async (ctx) => {
          if (a.selection_only) {
            const sel = ctx.document.getSelection();
            sel.load("text");
            const ps = sel.paragraphs;
            ps.load("items/text,items/style");
            await ctx.sync();
            return { ok: true, selection: truncate(sel.text, MAX_READ_CHARS), paragraphs: ps.items.length };
          }
          const paras = await loadParagraphs(ctx, "items/text,items/style");
          const start = Math.max(0, Number(a.start) || 0);
          const end = Math.min(paras.items.length, Number(a.end) || start + 80);
          const out: { i: number; style: string; text: string }[] = [];
          let chars = 0;
          for (let i = start; i < end; i++) {
            const t = paras.items[i].text;
            if (chars + t.length > MAX_READ_CHARS && out.length) break;
            chars += t.length;
            out.push({ i, style: paras.items[i].style, text: t });
          }
          const last = out.length ? out[out.length - 1].i : start - 1;
          return {
            ok: true,
            total_paragraphs: paras.items.length,
            paragraphs: out,
            ...(last + 1 < paras.items.length ? { next_start: last + 1 } : {}),
          };
        });

      case "word_replace_text":
        return Word.run(async (ctx) => {
          const find = String(a.find || "");
          if (!find) throw new Error("Thiếu chuỗi cần tìm.");
          if (find.length > 255) throw new Error("Chuỗi tìm quá dài (>255 ký tự). Dùng word_set_paragraph_text để sửa cả đoạn.");
          return tracked(ctx, async () => {
            const res = ctx.document.body.search(escapeWordSearch(find), {
              matchCase: !!a.match_case,
              matchWholeWord: !!a.whole_word,
            });
            res.load("items");
            await ctx.sync();
            const n = Math.min(res.items.length, Number(a.max_replacements) || 50);
            for (let i = 0; i < n; i++) res.items[i].insertText(String(a.replace ?? ""), "Replace");
            await ctx.sync();
            return {
              ok: n > 0,
              found: res.items.length,
              replaced: n,
              message: n ? `Đã thay ${n} chỗ` : `Không tìm thấy "${truncate(find, 80)}"`,
            };
          });
        });

      case "word_set_paragraph_text":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const p = await paragraphAt(ctx, a.index);
            const oldText = p.text;
            const newText = String(a.text ?? "");
            const d = diffMiddle(oldText, newText);
            if (!d.removed && !d.added) return { ok: true, message: "Nội dung không đổi" };
            let mode = "cả đoạn";
            if (d.removed && d.removed.length <= 255) {
              const hits = p.search(escapeWordSearch(d.removed), { matchCase: true });
              hits.load("items");
              await ctx.sync();
              if (hits.items.length === 1) {
                if (d.added) hits.items[0].insertText(d.added, "Replace");
                else hits.items[0].delete();
                mode = "phần thay đổi";
              }
            } else if (!d.removed && d.added) {
              if (!d.prefix) {
                p.insertText(d.added, "Start");
                mode = "phần thêm";
              } else {
                const anchor = d.prefix.slice(-Math.min(d.prefix.length, 120));
                const hits = p.search(escapeWordSearch(anchor), { matchCase: true });
                hits.load("items");
                await ctx.sync();
                if (hits.items.length === 1) {
                  hits.items[0].insertText(d.added, "After");
                  mode = "phần thêm";
                }
              }
            }
            if (mode === "cả đoạn") p.insertText(newText, "Replace");
            await ctx.sync();
            return { ok: true, message: `Đã sửa đoạn ${a.index} (${mode})` };
          })
        );

      case "word_insert_paragraphs":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const lines = String(a.text ?? "")
              .split(/\r?\n/)
              .filter((l) => l.trim() !== "");
            if (!lines.length) throw new Error("Không có nội dung để chèn.");
            const pos = a.position || "end";
            let anchor: Word.Paragraph | null = null;
            let loc: "After" | "Before" = "After";
            if (pos === "after_paragraph" || pos === "before_paragraph") {
              anchor = await paragraphAt(ctx, a.paragraph_index);
              loc = pos === "before_paragraph" ? "Before" : "After";
            } else if (pos === "after_selection") {
              anchor = ctx.document.getSelection().paragraphs.getLast();
            }
            let prev: Word.Paragraph | null = null;
            for (const raw of lines) {
              const parsed = a.style ? { text: raw, style: mapStyle(a.style) } : parseMarkdownLine(raw);
              let np: Word.Paragraph;
              if (prev) np = prev.insertParagraph(parsed.text, "After");
              else if (anchor) np = anchor.insertParagraph(parsed.text, loc);
              else np = ctx.document.body.insertParagraph(parsed.text, pos === "start" ? "Start" : "End");
              np.styleBuiltIn = parsed.style as any;
              prev = np;
            }
            await ctx.sync();
            return { ok: true, message: `Đã chèn ${lines.length} đoạn (${pos})` };
          })
        );

      case "word_replace_selection":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const sel = ctx.document.getSelection();
            sel.insertText(String(a.text ?? ""), "Replace");
            await ctx.sync();
            return { ok: true, message: "Đã thay nội dung đang chọn" };
          })
        );

      case "word_format_text":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const { ranges, paras } = await targetRanges(ctx, a);
            for (const r of ranges) {
              if (a.bold !== undefined) r.font.bold = !!a.bold;
              if (a.italic !== undefined) r.font.italic = !!a.italic;
              if (a.underline !== undefined) r.font.underline = (a.underline ? "Single" : "None") as any;
              if (a.font_size) r.font.size = Number(a.font_size);
              if (a.font_name) r.font.name = a.font_name;
              if (a.font_color) r.font.color = a.font_color;
              if (a.highlight_color) r.font.highlightColor = /^none$/i.test(a.highlight_color) ? (null as any) : a.highlight_color;
            }
            for (const p of paras) {
              if (a.alignment) p.alignment = a.alignment;
              if (a.style) p.styleBuiltIn = mapStyle(a.style) as any;
            }
            await ctx.sync();
            return { ok: true, message: `Đã định dạng ${ranges.length} vùng` };
          })
        );

      case "word_add_comment":
        return Word.run(async (ctx) => {
          if (!supports("1.4")) throw new Error("Word trên máy này chưa hỗ trợ thêm nhận xét qua add-in.");
          const { ranges } = await targetRanges(ctx, a);
          const c = ranges[0].insertComment(String(a.text || ""));
          c.load("id");
          await ctx.sync();
          return { ok: true, message: "Đã thêm nhận xét" };
        });

      case "word_insert_table":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const vals = normalize2D(a.values).map((r) => r.map((v) => String(v)));
            if (!vals.length) throw new Error("values rỗng.");
            let anchor: Word.Paragraph;
            if (a.position === "after_paragraph") anchor = await paragraphAt(ctx, a.paragraph_index);
            else if (a.position === "after_selection") anchor = ctx.document.getSelection().paragraphs.getLast();
            else anchor = ctx.document.body.paragraphs.getLast();
            const t = anchor.insertTable(vals.length, vals[0].length, "After", vals);
            t.headerRowCount = 1;
            t.styleBuiltIn = "GridTable4_Accent1" as any;
            await ctx.sync();
            return { ok: true, message: `Đã chèn bảng ${vals.length}×${vals[0].length}` };
          })
        );

      case "word_delete_paragraphs":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const paras = await loadParagraphs(ctx);
            const s = Number(a.start);
            const e = a.end === undefined || a.end === null ? s : Number(a.end);
            if (!(s >= 0 && e >= s && e < paras.items.length)) throw new Error(`Chỉ số đoạn không hợp lệ (có ${paras.items.length} đoạn).`);
            for (let i = e; i >= s; i--) paras.items[i].delete();
            await ctx.sync();
            return { ok: true, message: `Đã xóa đoạn ${s}${e > s ? `–${e}` : ""}` };
          })
        );
    }
    throw new Error(`Không có công cụ ${name}`);
  }

  return {
    host: "Word",
    specs: WORD_SPECS,
    execute,
    overview,
    systemPrompt: (ov) =>
      [
        "Bạn là trợ lý AI được tích hợp trong Microsoft Word, có thể thao tác trực tiếp trên văn bản đang mở bằng các công cụ (word_*).",
        "Nguyên tắc:",
        "- Khi người dùng muốn sửa, viết, chèn, định dạng, rà soát… hãy GỌI CÔNG CỤ để làm thật trong văn bản, không chỉ đưa gợi ý.",
        "- Đoạn văn đánh số từ 0. Đọc nội dung (word_read_paragraphs) trước khi sửa.",
        "- Sửa nhỏ (từ, cụm từ, số liệu) dùng word_replace_text; viết lại cả đoạn dùng word_set_paragraph_text. Sửa đổi được ghi bằng Theo dõi thay đổi để người dùng duyệt.",
        "- Rà soát văn bản/hợp đồng: ghi nhận xét đúng chỗ bằng word_add_comment; chỉ sửa trực tiếp khi được yêu cầu.",
        "- Giữ nguyên văn phong, định dạng; không tự bịa số hiệu văn bản, điều khoản hay căn cứ pháp lý.",
        "- Xong việc: trả lời ngắn gọn bằng tiếng Việt: đã sửa/chèn gì, ở đoạn nào.",
        "",
        "Trạng thái hiện tại của văn bản (JSON):",
        JSON.stringify(ov).slice(0, 6000),
      ].join("\n"),
  };
}
