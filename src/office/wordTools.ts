/* global Word, Office */
/** Công cụ để AI thao tác trực tiếp trên văn bản Word. Sửa đổi được ghi bằng Theo dõi thay đổi (Track Changes). */
import { HostToolSet, ToolSpec, UndoEntry } from "../core/agent/types";
import {
  RESULT_JSON_BUDGET,
  diffMiddle,
  escapeWordSearch,
  normalize2D,
  overviewForPrompt,
  packParagraphs,
  paragraphIndex,
  parseMarkdownLine,
  takeWithinBudget,
  truncate,
} from "./util";

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
    description:
      "Read paragraphs by index range (0-based, end exclusive) with their style. Or read the current selection. Long results stop early: when the result has next_start, call again with start=next_start to continue; compare with total_paragraphs to be sure you read everything.",
    parameters: {
      type: "object",
      properties: {
        start: { type: "number", description: "First paragraph index (default 0)." },
        end: { type: "number", description: "End index, exclusive (default start+200; the result may stop earlier, see next_start)." },
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
  {
    name: "word_list_comments",
    description:
      "List comments with their anchored text, author, paragraph index, replies and resolved state. If the result has next_start, call again with start=next_start for the rest.",
    parameters: {
      type: "object",
      properties: {
        include_resolved: { type: "boolean", description: "Default false." },
        start: { type: "number", description: "Skip the first N comments (paging). Default 0." },
      },
    },
  },
  {
    name: "word_reply_comment",
    description: "Reply to a comment thread (e.g. explain what was changed).",
    parameters: { type: "object", properties: { id: { type: "string" }, text: { type: "string" } }, required: ["id", "text"] },
    write: true,
  },
  {
    name: "word_resolve_comment",
    description: "Mark a comment as resolved (or reopen it).",
    parameters: { type: "object", properties: { id: { type: "string" }, resolved: { type: "boolean", description: "Default true." } }, required: ["id"] },
    write: true,
  },
  {
    name: "word_list_revisions",
    description:
      "List tracked changes (redlines) in the document: index, type (Added/Deleted/Formatted…), author, date, text, paragraph index. If the result has next_start, call again with start=next_start for the rest.",
    parameters: { type: "object", properties: { start: { type: "number", description: "First revision index (paging). Default 0." } } },
  },
  {
    name: "word_review_revision",
    description: "Accept or reject tracked changes: by index list from word_list_revisions, or all at once.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["accept", "reject"] },
        indexes: { type: "array", items: { type: "number" } },
        all: { type: "boolean" },
      },
      required: ["action"],
    },
    write: true,
  },
  {
    name: "word_read_tables",
    description: "Read existing tables: index, size and cell values. If the result has next_start, call again with start=next_start for the remaining tables.",
    parameters: {
      type: "object",
      properties: {
        max_rows: { type: "number", description: "Default 40 rows per table." },
        start: { type: "number", description: "First table index (paging). Default 0." },
      },
    },
  },
  {
    name: "word_fill_table",
    description: "Fill cells of an existing table in place (tracked), starting at a row/column; extra rows are appended.",
    parameters: {
      type: "object",
      properties: {
        table_index: { type: "number", description: "0-based, from word_read_tables." },
        values: { type: "array", items: { type: "array", items: { type: "string" } } },
        start_row: { type: "number", description: "Default 1 (skip header row)." },
        start_col: { type: "number", description: "Default 0." },
      },
      required: ["table_index", "values"],
    },
    write: true,
  },
  {
    name: "word_go_to",
    description: "Select a paragraph so the user sees it.",
    parameters: { type: "object", properties: { paragraph_index: { type: "number" } }, required: ["paragraph_index"] },
  },
];

/** Từ 3.6 không còn công tắc tắt Track Changes: mọi sửa đổi của AI trong Word luôn ghi có theo dõi (anh duyệt 15/11/2026). */
export type WordSettings = unknown;

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
  void getSettings;
  /** Luôn bật Theo dõi thay đổi trong lúc AI sửa, xong trả lại chế độ cũ. Word không có API theo dõi (dưới WordApi 1.4) → từ chối ghi. */
  async function tracked<T>(ctx: Word.RequestContext, fn: () => Promise<T>): Promise<T> {
    if (!supports("1.4")) {
      throw new Error(
        "Word trên máy này chưa hỗ trợ ghi có theo dõi qua add-in (cần WordApi 1.4) nên AI không sửa trực tiếp, để mọi sửa đổi luôn duyệt được. Hãy cập nhật Office; tạm thời đề xuất nội dung để người dùng tự chèn."
      );
    }
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
          // Ngân sách tính theo độ dài JSON thật; next_start và truncated nằm đầu kết quả (không bị cắt mất)
          const start = Math.max(0, Number(a.start) || 0);
          return packParagraphs(
            paras.items.map((p) => ({ text: p.text, style: p.style })),
            start,
            Number(a.end) || start + 200
          );
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
            let prevKind: string | null = null;
            for (const raw of lines) {
              const parsed = a.style ? { text: raw, style: mapStyle(a.style) } : parseMarkdownLine(raw);
              let np: Word.Paragraph;
              if (prev) np = prev.insertParagraph(parsed.text, "After");
              else if (anchor) np = anchor.insertParagraph(parsed.text, loc);
              else np = ctx.document.body.insertParagraph(parsed.text, pos === "start" ? "Start" : "End");
              const kind = parsed.style === "ListBullet" || parsed.style === "ListNumber" ? parsed.style : "other";
              // Đoạn mới chèn sau một mục danh sách sẽ tự thuộc danh sách đó
              let inList: boolean;
              if (prevKind === null) {
                np.load("isListItem");
                await ctx.sync();
                inList = np.isListItem;
              } else inList = prevKind !== "other";
              if (kind === "other") {
                if (inList) np.detachFromList();
                np.styleBuiltIn = parsed.style as any;
              } else if (prevKind !== kind) {
                if (inList) np.detachFromList();
                const list = np.startNewList();
                if (kind === "ListBullet") list.setLevelBullet(0, "Solid" as any);
                else list.setLevelNumbering(0, "Arabic" as any, [0, "."]);
              }
              prev = np;
              prevKind = kind;
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
            const style = a.style ? mapStyle(a.style) : "";
            const isList = style === "ListBullet" || style === "ListNumber";
            for (const p of paras) {
              if (a.alignment) p.alignment = a.alignment;
              if (style && !isList) p.styleBuiltIn = style as any;
            }
            if (isList) {
              paras.forEach((p) => p.load("isListItem"));
              await ctx.sync();
              let listId: number | undefined;
              for (const p of paras) {
                if (p.isListItem) continue;
                if (listId === undefined) {
                  const list = p.startNewList();
                  if (style === "ListBullet") list.setLevelBullet(0, "Solid" as any);
                  else list.setLevelNumbering(0, "Arabic" as any, [0, "."]);
                  list.load("id");
                  await ctx.sync();
                  listId = list.id;
                } else p.attachToList(listId, 0);
              }
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
      case "word_list_comments":
        return Word.run(async (ctx) => {
          if (!supports("1.4")) throw new Error("Word trên máy này chưa hỗ trợ đọc nhận xét qua add-in.");
          const withIds = supports("1.6");
          const paras = await loadParagraphs(ctx, withIds ? "items/text,items/uniqueLocalId" : "items/text");
          const cs = ctx.document.body.getComments();
          cs.load("items/id,items/content,items/authorName,items/creationDate,items/resolved");
          await ctx.sync();
          const all = cs.items.filter((c) => a.include_resolved || !c.resolved);
          const from = Math.max(0, Number(a.start) || 0);
          const items = all.slice(from, from + 100);
          const anchors = items.map((c) => {
            const r = c.getRange();
            r.load("text");
            const p = r.paragraphs.getFirst();
            p.load(withIds ? "text,uniqueLocalId" : "text");
            const reps = c.replies;
            reps.load("items/content,items/authorName");
            return { r, p, reps };
          });
          await ctx.sync();
          const texts = paras.items.map((p) => p.text);
          // Số đoạn theo định danh đoạn (đúng cả khi có đoạn trùng nội dung); Word cũ thì so chữ
          const ids = withIds ? paras.items.map((p) => p.uniqueLocalId) : null;
          const rows = items.map((c, i) => ({
            id: c.id,
            author: c.authorName,
            text: truncate(c.content, 2000),
            anchor: truncate(anchors[i].r.text, 300),
            paragraph: paragraphIndex(withIds ? anchors[i].p.uniqueLocalId : undefined, anchors[i].p.text, ids, texts),
            resolved: c.resolved,
            replies: anchors[i].reps.items.map((x) => ({ author: x.authorName, text: truncate(x.content, 1000) })),
          }));
          const { picked, truncated } = takeWithinBudget(rows, RESULT_JSON_BUDGET - 400);
          const next = from + picked.length;
          return {
            ok: true,
            total: all.length,
            shown_from: from,
            shown: picked.length,
            ...(next < all.length ? { next_start: next } : {}),
            ...(truncated ? { truncated: true } : {}),
            comments: picked,
          };
        });

      case "word_reply_comment":
      case "word_resolve_comment":
        return Word.run(async (ctx) => {
          if (!supports("1.4")) throw new Error("Word trên máy này chưa hỗ trợ nhận xét qua add-in.");
          const cs = ctx.document.body.getComments();
          cs.load("items/id");
          await ctx.sync();
          const c = cs.items.find((x) => x.id === String(a.id));
          if (!c) throw new Error(`Không có nhận xét id ${a.id}.`);
          if (name === "word_reply_comment") {
            c.reply(String(a.text || ""));
            await ctx.sync();
            return { ok: true, message: "Đã trả lời nhận xét" };
          }
          c.resolved = a.resolved !== false;
          await ctx.sync();
          return { ok: true, message: a.resolved === false ? "Đã mở lại nhận xét" : "Đã đánh dấu đã xử lý" };
        });

      case "word_list_revisions":
        return Word.run(async (ctx) => {
          if (!supports("1.6")) throw new Error("Word trên máy này chưa hỗ trợ đọc sửa đổi được theo dõi qua add-in.");
          const paras = await loadParagraphs(ctx, "items/text,items/uniqueLocalId");
          const tc = ctx.document.body.getTrackedChanges();
          tc.load("items/type,items/author,items/date,items/text");
          await ctx.sync();
          const from = Math.max(0, Number(a.start) || 0);
          const items = tc.items.slice(from, from + 200);
          const ps = items.map((t) => {
            const p = t.getRange().paragraphs.getFirst();
            p.load("text,uniqueLocalId");
            return p;
          });
          await ctx.sync();
          const texts = paras.items.map((p) => p.text);
          const ids = paras.items.map((p) => p.uniqueLocalId);
          const rows = items.map((t, k) => ({
            index: from + k,
            type: t.type,
            author: t.author,
            date: t.date ? new Date(t.date as any).toLocaleString("vi-VN") : "",
            text: truncate(t.text, 300),
            paragraph: paragraphIndex(ps[k].uniqueLocalId, ps[k].text, ids, texts),
          }));
          const { picked, truncated } = takeWithinBudget(rows, RESULT_JSON_BUDGET - 400);
          const next = from + picked.length;
          return {
            ok: true,
            total: tc.items.length,
            shown_from: from,
            shown: picked.length,
            ...(next < tc.items.length ? { next_start: next } : {}),
            ...(truncated ? { truncated: true } : {}),
            revisions: picked,
          };
        });

      case "word_review_revision":
        return Word.run(async (ctx) => {
          if (!supports("1.6")) throw new Error("Word trên máy này chưa hỗ trợ duyệt sửa đổi qua add-in.");
          const tc = ctx.document.body.getTrackedChanges();
          tc.load("items");
          await ctx.sync();
          const accept = a.action === "accept";
          if (a.all) {
            const n = tc.items.length;
            if (accept) tc.acceptAll();
            else tc.rejectAll();
            await ctx.sync();
            return { ok: true, message: `Đã ${accept ? "chấp nhận" : "từ chối"} tất cả ${n} sửa đổi` };
          }
          const idx = Array.from(new Set<number>((a.indexes || []).map((x: any) => Number(x))))
            .filter((i) => i >= 0 && i < tc.items.length)
            .sort((x, y) => y - x);
          if (!idx.length) throw new Error("Cần indexes (từ word_list_revisions) hoặc all=true.");
          for (const i of idx) {
            if (accept) tc.items[i].accept();
            else tc.items[i].reject();
          }
          await ctx.sync();
          return { ok: true, message: `Đã ${accept ? "chấp nhận" : "từ chối"} ${idx.length} sửa đổi` };
        });

      case "word_read_tables":
        return Word.run(async (ctx) => {
          const tables = ctx.document.body.tables;
          tables.load("items/rowCount,items/values");
          await ctx.sync();
          const maxRows = Number(a.max_rows) || 40;
          const from = Math.max(0, Number(a.start) || 0);
          const rows = tables.items.slice(from, from + 30).map((t, k) => ({
            index: from + k,
            rows: t.rowCount,
            columns: t.values[0]?.length || 0,
            ...(t.rowCount > maxRows ? { rows_shown: maxRows } : {}),
            values: t.values.slice(0, maxRows).map((r) => r.map((v) => truncate(v, 200))),
          }));
          const { picked, truncated } = takeWithinBudget(rows, RESULT_JSON_BUDGET - 400, (x, room) => {
            // Bảng đầu quá lớn: bớt hàng cho vừa
            let vals = x.values;
            while (vals.length > 1 && JSON.stringify(vals).length > room) vals = vals.slice(0, Math.ceil(vals.length / 2));
            return { ...x, rows_shown: vals.length, values: vals };
          });
          const next = from + picked.length;
          return {
            ok: true,
            count: tables.items.length,
            ...(next < tables.items.length ? { next_start: next } : {}),
            ...(truncated ? { truncated: true } : {}),
            tables: picked,
          };
        });

      case "word_fill_table":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const tables = ctx.document.body.tables;
            tables.load("items/rowCount,items/values");
            await ctx.sync();
            const t = tables.items[Number(a.table_index)];
            if (!t) throw new Error(`Không có bảng số ${a.table_index} (có ${tables.items.length} bảng).`);
            const vals = normalize2D(a.values).map((r) => r.map((v) => String(v)));
            const r0 = a.start_row === undefined ? 1 : Number(a.start_row);
            const c0 = Number(a.start_col) || 0;
            const cols = t.values[0]?.length || 0;
            let written = 0;
            const extra: string[][] = [];
            vals.forEach((row, i) => {
              const r = r0 + i;
              if (r < t.rowCount) {
                row.forEach((v, j) => {
                  if (c0 + j < cols) {
                    t.getCell(r, c0 + j).value = v;
                    written++;
                  }
                });
              } else {
                const full = Array(cols).fill("");
                row.forEach((v, j) => c0 + j < cols && (full[c0 + j] = v));
                extra.push(full);
              }
            });
            if (extra.length) t.addRows("End", extra.length, extra);
            await ctx.sync();
            return { ok: true, message: `Đã điền ${written} ô${extra.length ? `, thêm ${extra.length} hàng` : ""} vào bảng ${a.table_index}` };
          })
        );

      case "word_go_to":
        return Word.run(async (ctx) => {
          const p = await paragraphAt(ctx, a.paragraph_index);
          p.select();
          await ctx.sync();
          return { ok: true, message: `Đã chuyển tới đoạn ${a.paragraph_index}` };
        });

    }
    throw new Error(`Không có công cụ ${name}`);
  }

  async function snapshot() {
    return Word.run(async (ctx) => {
      const paras = await loadParagraphs(ctx, "items/text");
      let chars = 0;
      const out: string[] = [];
      for (let i = 0; i < paras.items.length && chars < 40000; i++) {
        const t = paras.items[i].text;
        if (!t.trim()) continue;
        out.push(`[¶${i}] ${t}`);
        chars += t.length;
      }
      return { title: "Văn bản Word đang mở", data: { paragraphs: paras.items.length, text: out.join("\n") } };
    });
  }

  const RISK: Record<string, (a: any) => string | null> = {
    word_delete_paragraphs: (a) => `Xóa đoạn ${a.start}${a.end !== undefined && a.end !== a.start ? `–${a.end}` : ""} (ghi bằng Track Changes nếu đang bật).`,
    word_review_revision: (a) => (a.all ? `${a.action === "accept" ? "Chấp nhận" : "Từ chối"} TẤT CẢ sửa đổi đang theo dõi trong văn bản.` : null),
  };

  return {
    host: "Word",
    snapshot,
    specs: WORD_SPECS.map((sp) => (RISK[sp.name] ? { ...sp, risky: async (a: any) => RISK[sp.name](a) } : sp)),
    execute,
    overview,
    systemPrompt: (ov) =>
      [
        "Bạn là trợ lý AI được tích hợp trong Microsoft Word, có thể thao tác trực tiếp trên văn bản đang mở bằng các công cụ (word_*).",
        "Nguyên tắc:",
        "- Khi người dùng muốn sửa, viết, chèn, định dạng, rà soát… hãy GỌI CÔNG CỤ để làm thật trong văn bản, không chỉ đưa gợi ý.",
        "- Đoạn văn đánh số từ 0. Đọc nội dung (word_read_paragraphs) trước khi sửa. Kết quả có next_start là chưa đọc hết: gọi tiếp với start = next_start cho tới khi đã đọc đủ total_paragraphs.",
        "- Sửa nhỏ (từ, cụm từ, số liệu) dùng word_replace_text; viết lại cả đoạn dùng word_set_paragraph_text. Sửa đổi được ghi bằng Theo dõi thay đổi để người dùng duyệt.",
        "- Rà soát văn bản/hợp đồng: ghi nhận xét đúng chỗ bằng word_add_comment; chỉ sửa trực tiếp khi được yêu cầu.",
        "- Giữ nguyên văn phong, định dạng; không tự bịa số hiệu văn bản, điều khoản hay căn cứ pháp lý.",
        "- Xong việc: trả lời ngắn gọn bằng tiếng Việt: đã sửa/chèn gì, ở đoạn nào.",
        "",
        "Trạng thái hiện tại của văn bản (JSON):",
        overviewForPrompt(ov),
      ].join("\n"),
  };
}
