/* global Word, Office */
/** Công cụ để AI thao tác trực tiếp trên văn bản Word. Sửa đổi được ghi bằng Theo dõi thay đổi (Track Changes).
 *  Bản 4.0 (Đợt 1A): mục lục điều khoản, tìm và đọc theo điều, chống sửa nhầm đoạn, sửa mức từ, đọc đúng khi văn bản đang có
 *  sửa đổi, chèn điều khoản kế thừa đánh số, Hoàn tác. Căn cứ kỹ thuật: bảng thăm dò trong tasks/dot/dot-1a/dot-1a.md. */
import { HostToolSet, ToolSpec, UndoEntry, Verifier } from "../core/agent/types";
import { officeErrorHint } from "../core/errors";
import { diffUserChanges, snapshotOf } from "../core/userChanges";
import { EXPLAIN_TOOL, VERIFIER_SYSTEM, assembleResult, buildVerifierInput, normalizeEditsInfo, parseVerdicts } from "../core/explainEdits";
import {
  RESULT_JSON_BUDGET,
  escapeWordSearch,
  isMarkdownNumberedList,
  normalize2D,
  planTableFill,
  overviewForPrompt,
  packParagraphs,
  PENDING_NOTE,
  paragraphIndex,
  parseMarkdownLine,
  replaceLimit,
  searchRows,
  searchRowsByWords,
  takeWithinBudget,
  truncate,
} from "./util";
import { OutlineItem, ParaRow, buildOutline, clauseAt, clauseName, findClause, outlineByParagraph } from "../core/outline";
import { WORD_REVIEW_SPECS, createWordReviewTools } from "./wordReview";
import { createWordReadTools } from "./wordRead";
import { createWordEditTools } from "./wordEdit";
import { createWordCommentTools } from "./wordComments";
import { WORD_ADMIN_SPECS, createWordAdminTools } from "./wordAdmin";
import { WORD_FORM_SPECS, createWordFormTools } from "./wordForm";
import { WORD_SPECIAL_LINE } from "../core/toolgroups";
import { newTextCheck } from "../core/textcheck";
/** Lệnh ghi chữ mới mà kết quả cần kiểm từ lạ (tham số text) */
const NEW_TEXT_TOOLS = new Set(["word_set_paragraph_text", "word_insert_paragraphs", "word_insert_clause"]);
import {
  ChangeSnap,
  beforeExcerpt,
  EditPlan,
  Locator,
  Wanted,
  changeKey,
  diffChanges,
  locateParagraph,
  pickWanted,
  planParagraphEdit,
  resolveAuthor,
  snapChange,
} from "../core/worddiff";
import { STYLE_ENUM, MAX_READ_CHARS, MANY_REPLACEMENTS, BIG_SELECTION_CHARS, BIG_SELECTION_PARAS, WHOLE_PARAGRAPH_RATIO, WHOLE_PARAGRAPH_HUNKS, tracked, targetRanges, indexOfIds, restoreText, removeParas, supports, mapStyle, NOTE_MARK, noMarks, withNoteSpans, strip, UNTRUSTED_NOTE, curOf, currentTableValues, loadRows, loadRowsOf, markCells, tableSpanProblem, withNums, resolvePara, missingExpect, pendingInfo, movedNote, liveHits, applyPlan, changeSnaps, loadChanges, rejectWanted, NOT_FOUND_PENDING, markPendingHits, annotatePending } from "./wordCore";
import type { Row, Loaded, PendingRow, UndoStep, Undo } from "./wordCore";
/** Hằng số dời sang wordCore.ts (B13): tái xuất để nơi khác vẫn nhập từ wordTools */
export { MANY_REPLACEMENTS, BIG_SELECTION_CHARS, BIG_SELECTION_PARAS, WHOLE_PARAGRAPH_RATIO, WHOLE_PARAGRAPH_HUNKS };


/** Đợt 10: viết gọn — lời nhắc hệ thống (mục SỬA ĐÚNG ĐOẠN) đã giải thích expect một lần cho mọi công cụ */
const EXPECT_DESC = "First ~30 characters of the paragraph exactly as read (checked before writing).";

export const WORD_SPECS: ToolSpec[] = [
  {
    name: "word_get_overview",
    description:
      "Document overview: paragraph count, top-level outline, first paragraphs, selection, track-changes mode, pending changes and comments count.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "word_get_outline",
    description:
      "Clause outline: every numbered clause (Word-generated and typed numbers) with paragraph index `i`, full number `num` ('Điều 5', '5.3', '5.3(a)'), depth `lvl`, title, last paragraph `end`. Call first on contracts and long documents, then read only the clauses you need.",
    parameters: {
      type: "object",
      properties: {
        max_level: { type: "number", description: "0 = top level only. Default: all." },
        start: { type: "number", description: "From next_start." },
      },
    },
  },
  {
    name: "word_find",
    description:
      "Search the current text; returns matching paragraphs with clause number and excerpt. Case-insensitive; falls back to ignoring accents, then to paragraphs with most of the words (best first), so one call with key words is usually enough. Ignores tracked-deleted text.",
    parameters: {
      type: "object",
      properties: {
        query: { type: "string", description: "Word or short phrase to find." },
        match_case: { type: "boolean" },
        whole_word: { type: "boolean" },
        max: { type: "number", description: "Default 30." },
      },
      required: ["query"],
    },
  },
  {
    name: "word_read_clause",
    description:
      "Read one whole clause with its sub-clauses by number: 'Điều 5', '5.3', '5.3(a)', 'khoản 2 Điều 5', 'Chương II'.",
    parameters: {
      type: "object",
      properties: { number: { type: "string" } },
      required: ["number"],
    },
  },
  {
    name: "word_read_paragraphs",
    description:
      "Read paragraphs by index range (0-based, end exclusive) with style and clause number `num`, or the selection. Text is the CURRENT version. `changed` = pending tracked changes; then `before` is the text before them and `changed_by` the author: when you rely on such a paragraph, state both versions and that the change is pending. If next_start is present, call again from it (compare with total_paragraphs).",
    parameters: {
      type: "object",
      properties: {
        start: { type: "number", description: "Default 0." },
        end: { type: "number", description: "Exclusive. Default start+200." },
        selection_only: { type: "boolean" },
        view: {
          type: "string",
          enum: ["current", "original"],
          description: "original = before pending tracked changes. Default current.",
        },
      },
    },
  },
  {
    name: "word_replace_text",
    description:
      "Find and replace text (tracked), for small edits of words, phrases, figures (max 255 chars). If it occurs more than once nothing is replaced and you get the places: call again with paragraph_index (and occurrence), or all=true.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string" },
        replace: { type: "string" },
        paragraph_index: { type: "number" },
        expect: { type: "string", description: "With paragraph_index: " + EXPECT_DESC },
        occurrence: { type: "number", description: "1 = first." },
        all: { type: "boolean" },
        match_case: { type: "boolean" },
        whole_word: { type: "boolean" },
        max_replacements: { type: "number", description: "With all=true. Default 50." },
      },
      required: ["find", "replace"],
    },
    write: true,
    reviewable: true,
  },
  {
    name: "word_set_paragraph_text",
    description:
      "Rewrite one paragraph (tracked): give its FULL new text, without a Word-generated clause number, no line breaks. Only differing words change; formatting is kept.",
    parameters: {
      type: "object",
      properties: {
        index: { type: "number" },
        text: { type: "string" },
        expect: { type: "string", description: EXPECT_DESC },
      },
      required: ["index", "text", "expect"],
    },
    write: true,
  },
  {
    name: "word_insert_paragraphs",
    description:
      "Insert free content (tracked), one paragraph per line: '# ', '## ', '### ' headings, '- ' bullets, lines '1. ', '2. ' a numbered list. For clauses continuing the document's numbering use word_insert_clause.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string" },
        position: {
          type: "string",
          enum: ["end", "start", "after_paragraph", "before_paragraph", "after_selection"],
          description: "Default end.",
        },
        paragraph_index: { type: "number" },
        expect: { type: "string", description: "With paragraph_index: " + EXPECT_DESC },
        style: { type: "string", enum: STYLE_ENUM, description: "One style for all lines." },
      },
      required: ["text"],
    },
    write: true,
  },
  {
    name: "word_insert_clause",
    description:
      "Insert clause(s) after an existing clause (and its sub-clauses), inheriting its style and auto numbering (tracked): after 8.2 the new one becomes 8.3. One clause per line; do not type the number when Word numbers the document. Returns the new numbers.",
    parameters: {
      type: "object",
      properties: {
        after_index: { type: "number", description: "First paragraph of the anchor clause." },
        expect: { type: "string", description: EXPECT_DESC },
        text: { type: "string" },
        level: {
          type: "string",
          enum: ["same", "deeper", "shallower"],
          description: "same (default) = sibling; deeper = its first sub-clause; shallower = one level up.",
        },
      },
      required: ["after_index", "text", "expect"],
    },
    write: true,
  },
  {
    name: "word_change_level",
    description: "Move an auto-numbered clause one level deeper or shallower in Word's numbering (e.g. 8.3 → 8.2(a)).",
    parameters: {
      type: "object",
      properties: {
        index: { type: "number" },
        expect: { type: "string", description: EXPECT_DESC },
        direction: { type: "string", enum: ["deeper", "shallower"] },
      },
      required: ["index", "direction", "expect"],
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
      "Format text. Target one of: every occurrence of `find` (optionally inside paragraph_index), a whole paragraph_index, or use_selection. Alignment, indent, spacing and style apply to whole paragraphs.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string" },
        paragraph_index: { type: "number" },
        expect: { type: "string", description: "With paragraph_index: " + EXPECT_DESC },
        use_selection: { type: "boolean" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_size: { type: "number" },
        font_name: { type: "string" },
        font_color: { type: "string", description: "Hex, e.g. #C00000." },
        highlight_color: { type: "string", description: "Hex or 'none'." },
        alignment: { type: "string", enum: ["Left", "Centered", "Right", "Justified"] },
        first_line_indent_mm: { type: "number", description: "Administrative documents: 10. 0 removes." },
        space_after_pt: { type: "number", description: "Administrative documents: 6." },
        line_spacing: { type: "number", description: "In lines: 1, 1.2, 1.5." },
        style: { type: "string", enum: STYLE_ENUM },
      },
    },
    write: true,
  },
  {
    name: "word_add_comment",
    description:
      "Add a review comment (review notes, legal risks). Target one of: first occurrence of `find` (optionally inside paragraph_index), a whole paragraph_index, or use_selection.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string" },
        find: { type: "string", description: "Anchor text, max 255 chars." },
        paragraph_index: { type: "number" },
        expect: { type: "string", description: "With paragraph_index: " + EXPECT_DESC },
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
        expect: { type: "string", description: "With paragraph_index: " + EXPECT_DESC },
      },
      required: ["values"],
    },
    write: true,
  },
  {
    name: "word_delete_paragraphs",
    description:
      "Delete paragraphs start..end inclusive (tracked; the user confirms first). Only when the user asks to remove content. Never across table cells or a table edge.",
    parameters: {
      type: "object",
      properties: {
        start: { type: "number" },
        end: { type: "number" },
        expect: { type: "string", description: "For `start`: " + EXPECT_DESC },
      },
      required: ["start", "expect"],
    },
    write: true,
  },
  {
    name: "word_list_comments",
    description:
      "List comments: anchored text, author, paragraph index, clause number, replies, resolved state. Page with next_start.",
    parameters: {
      type: "object",
      properties: {
        include_resolved: { type: "boolean", description: "Default false." },
        start: { type: "number", description: "From next_start." },
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
      "List pending tracked changes by paragraph: clause number, sentence before/after, and each change (index, type, author, date, text). `total` = number of changes, `authors` = count per author. Call again from next_start until you have them all.",
    parameters: {
      type: "object",
      properties: {
        start: { type: "number", description: "From next_start." },
        author: { type: "string", description: "Only this author (name or part)." },
      },
    },
  },
  {
    name: "word_review_revision",
    description:
      "Accept or reject tracked changes by indexes from word_list_revisions (with `texts` to verify), or all (optionally one author's). Not undoable by the add-in.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["accept", "reject"] },
        indexes: { type: "array", items: { type: "number" } },
        texts: {
          type: "array",
          items: { type: "string" },
          description: "Per index: start of that change's text as listed.",
        },
        all: { type: "boolean" },
        author: { type: "string", description: "With all=true: only this author's changes." },
      },
      required: ["action"],
    },
    write: true,
  },
  {
    name: EXPLAIN_TOOL,
    description:
      "Before writing self-drafted legal wording (clause, definition, obligation, deadline, amount, deletion, accept/reject redline): give each planned edit's claims about the CURRENT document. A fresh reviewer without this chat checks them against the document, redlines, user instructions, loaded skills. Never ground on 'the user asked'. Skip mechanical edits.",
    parameters: {
      type: "object",
      properties: {
        edits: {
          type: "array",
          items: {
            type: "object",
            properties: {
              target: { type: "string", description: "e.g. Điều 5.2, [¶12]" },
              proposed_text: { type: "string" },
              claims: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    claim: { type: "string" },
                    basis: { type: "string", enum: ["document_text", "tracked_change", "skill_instruction", "user_instruction", "external_law", "model_judgment"] },
                    source: { type: "string", description: "clause, skill or law" },
                  },
                  required: ["claim", "basis"],
                },
              },
            },
            required: ["target", "proposed_text", "claims"],
          },
        },
      },
      required: ["edits"],
    },
  },
  {
    name: "word_read_tables",
    description: "Read tables: index, size, cell values. Page with next_start.",
    parameters: {
      type: "object",
      properties: {
        max_rows: { type: "number", description: "Per table. Default 40." },
        start: { type: "number", description: "From next_start." },
      },
    },
  },
  {
    name: "word_fill_table",
    description: "Fill cells of an existing table in place (tracked), starting at a row/column; extra rows are appended.",
    parameters: {
      type: "object",
      properties: {
        table_index: { type: "number", description: "From word_read_tables." },
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
  // Đợt 1B: soát nhất quán, đọc phần ngoài thân văn bản, thể thức trang, so sánh phiên bản (wordReview.ts)
  ...WORD_REVIEW_SPECS,
  // Đợt 3: dựng văn bản hành chính đúng thể thức, đặt thể thức trang (wordAdmin.ts)
  ...WORD_ADMIN_SPECS,
  // Đợt 8: điền mẫu — liệt kê chỗ cần điền, điền hết trong một lệnh (wordForm.ts)
  ...WORD_FORM_SPECS,
];

/** Công cụ sửa nội dung Word nhận thêm `reason`: một câu lý do ngắn, hiện cạnh bước trên khung, trong hộp hỏi và ở mục Hoàn tác
 *  (như thẻ đề xuất kèm lý do của add-in Claude chính thức). Lý do chỉ để hiển thị, không ghi vào tài liệu. */
export const REASON_TOOLS = new Set(["word_replace_text", "word_set_paragraph_text", "word_insert_paragraphs", "word_insert_clause", "word_replace_selection", "word_delete_paragraphs", "word_fill_table"]);
for (const sp of WORD_SPECS) {
  if (REASON_TOOLS.has(sp.name)) {
    (sp.parameters as any).properties.reason = {
      type: "string",
      description: "Always give: why, one short sentence in the user's language (shown to the user).",
    };
  }
}

/** Lý do AI nêu cho một lần sửa, gọn một dòng (rỗng nếu không nêu) */
export function reasonOf(name: string, a: any): string {
  if (!REASON_TOOLS.has(name)) return "";
  const r = String(a?.reason ?? "").replace(/\s+/g, " ").trim();
  return r.length > 160 ? r.slice(0, 159) + "…" : r;
}
/** Công cụ ghi chữ vào điều khoản: khi người dùng bật "Soát lập luận trước khi sửa", mỗi lần gọi được soát trước khi ghi */
const REVIEWABLE = ["word_set_paragraph_text", "word_insert_clause", "word_replace_selection"];
WORD_SPECS.forEach((sp) => {
  if (REVIEWABLE.includes(sp.name)) sp.reviewable = true;
});
/** Cổng soát lập luận (4.13.1, anh chọn 25/11/2026): các lệnh ghi câu chữ vào thân văn bản. Chế độ Miễn phí công cụ soát nằm nhóm "soat". */
const EXPLAIN_GATED = ["word_replace_text", "word_set_paragraph_text", "word_insert_clause", "word_replace_selection", "word_insert_paragraphs"];
WORD_SPECS.forEach((sp) => {
  if (EXPLAIN_GATED.includes(sp.name)) sp.needsExplain = { tool: EXPLAIN_TOOL, hint: 'gọi load_tools nhóm "soat" trước' };
});

/** Từ 3.6 không còn công tắc tắt Track Changes: mọi sửa đổi của AI trong Word luôn ghi có theo dõi (anh duyệt 15/11/2026). */
export type WordSettings = unknown;


export function createWordTools(getSettings: () => WordSettings): HostToolSet {
  void getSettings;
  /** Chữ các đoạn lúc AI xong lượt trước (định danh → chữ) — chỉ trong bộ nhớ của khung (4.12.1) */
  let lastSnap: Map<string, string> | null = null;
  async function afterTurn() {
    if (!supports("1.6")) return;
    lastSnap = await Word.run(async (ctx) => snapshotOf((await loadRows(ctx)).rows));
  }

  async function overview() {
    return Word.run(async (ctx) => {
      const { rows, paras, footnotes, endnotes } = await loadRows(ctx, true);
      const userChanges = supports("1.6") ? diffUserChanges(lastSnap, rows) : null;
      const outline = buildOutline(rows);
      const sel = ctx.document.getSelection();
      sel.load("text");
      let mode: string | null = null;
      let comments: number | null = null;
      let changes: number | null = null;
      if (supports("1.4")) {
        ctx.document.load("changeTrackingMode");
        const cs = ctx.document.body.getComments();
        cs.load("items");
        const tc = supports("1.6") ? ctx.document.body.getTrackedChanges() : null;
        tc?.load("items/type");
        await ctx.sync();
        mode = ctx.document.changeTrackingMode;
        comments = cs.items.length;
        changes = tc ? tc.items.length : null;
      } else {
        await ctx.sync();
      }
      // Mục lục rút gọn: cấp cao nhất trước; văn bản ít mục thì lấy thêm cấp hai
      let top = outline.filter((o) => o.level === 0);
      if (top.length < 12) top = outline.filter((o) => o.level <= 1);
      const headings = top.slice(0, 80).map((o) => ({ index: o.i, num: o.num, text: truncate(o.title, 80) }));
      const live = rows.filter((r) => !r.deleted);
      return {
        paragraphs: paras.items.length,
        characters: live.reduce((n, r) => n + r.text.length, 0),
        clauses: outline.length,
        headings,
        first_paragraphs: live.slice(0, 6).map((r) => ({ index: r.i, style: r.style, text: truncate(r.text, 200) })),
        ...(userChanges ? { user_changes: userChanges } : {}),
        selection: truncate(sel.text, 1500),
        track_changes: mode,
        tracked_changes: changes,
        comments,
        ...(footnotes ? { footnotes } : {}),
        ...(endnotes ? { endnotes } : {}),
      };
    });
  }

  const review = createWordReviewTools({
    supports,
    loadRows: (ctx) => loadRows(ctx, true),
    loadRowsOf: (ctx, body) => loadRowsOf(ctx, body, true, true),
  });

  const admin = createWordAdminTools({ supports, tracked, resolvePara });
  const form = createWordFormTools({ supports, tracked, loadRows: (ctx) => loadRows(ctx), tableValues: (ctx, t, v) => currentTableValues(ctx, t, v) });

  const read = createWordReadTools({ overview });
  const edit = createWordEditTools({ overview });
  const comments = createWordCommentTools({ overview });

  async function run(name: string, a: any, u: Undo): Promise<any> {
    if (review.has(name)) return review.run(name, a);
    if (admin.has(name)) return admin.run(name, a, u);
    if (form.has(name)) return form.run(name, a);
    if (read.has(name)) return read.run(name, a, u);
    if (edit.has(name)) return edit.run(name, a, u);
    if (comments.has(name)) return comments.run(name, a, u);
    throw new Error(`Không có công cụ ${name}`);
  }

  const WRITE = new Set(WORD_SPECS.filter((sp) => sp.write).map((sp) => sp.name));
  const UNDO_LABEL: Record<string, string> = {
    word_add_comment: "Gỡ nhận xét vừa thêm",
    word_reply_comment: "Gỡ câu trả lời nhận xét",
    word_resolve_comment: "Trả lại trạng thái nhận xét",
    word_set_page_format: "Trả lại khổ giấy và lề cũ",
    // Gỡ thẳng các bảng và đoạn đã dựng theo định danh (không qua từ chối từng sửa đổi: phần chèn cạnh bảng để lại đoạn rỗng — đã đo)
    word_create_admin_document: "Gỡ văn bản hành chính vừa dựng",
  };

  /** Chạy công cụ và ghi cách hoàn tác: sửa đổi được theo dõi nào xuất hiện thêm sau công cụ này là của AI — hoàn tác = từ chối đúng
   *  các sửa đổi đó (ngày của sửa đổi chỉ chính xác tới phút nên không lọc theo thời gian); nhận xét thì gỡ theo id. */
  /** Sửa đổi mới trùng khóa với một sửa đổi có từ trước (cùng tác giả, cùng phút, cùng chữ — ví dụ người dùng vừa tự sửa y hệt ở đoạn khác):
   *  xác định sửa đổi nào là của AI theo đoạn mà công cụ đã đụng. */
  async function wantedOf(added: string[], after: ChangeSnap[], touched: string[]): Promise<Wanted[]> {
    const wants: Wanted[] = added.map((key) => ({ key }));
    const count = (list: string[], k: string) => list.filter((x) => x === k).length;
    const afterKeys = after.map((x) => x.key);
    const dup = new Set(added.filter((k) => count(afterKeys, k) > count(added, k)));
    if (!dup.size || !touched.length) return wants;
    const paras = await Word.run(async (ctx) => (await loadChanges(ctx, dup)).cur).catch(() => []);
    const mine = new Set(touched);
    for (const k of dup) {
      const inTouched = paras.filter((c) => c.key === k && c.para && mine.has(c.para)).map((c) => c.para as string);
      wants.filter((w) => w.key === k).forEach((w, n) => (w.para = inTouched[n]));
    }
    return wants;
  }

  /** Các công cụ ghi chạy lần lượt: công cụ bị hết hạn chờ vẫn chạy ngầm trong Office, công cụ kế không được chen vào giữa
   *  (hai lần chụp danh sách sửa đổi của Hoàn tác sẽ lẫn nhau). */
  let queue: Promise<unknown> = Promise.resolve();
  /** Soát lập luận (4.13.0): đọc văn bản kèm số điều và danh sách sửa đổi chờ duyệt, nhờ người soát ngữ cảnh mới kiểm từng căn cứ.
   *  Phần đọc Word đi qua hàng đợi ghi (không đọc giữa lúc một lệnh ghi đang chạy); lần gọi người soát thì ngoài hàng đợi. */
  async function explainEdits(a: any, verifier?: Verifier) {
    const { edits, dropped } = normalizeEditsInfo(a?.edits);
    if (!edits.length || edits.every((e) => !e.claims.length)) throw new Error("Thiếu edits: mỗi sửa đổi định làm cần target, proposed_text và ít nhất một căn cứ (claims).");
    const read = queue.then(() => Word.run(async (ctx) => {
      const { rows } = await loadRows(ctx, true);
      const nums = withNums(rows, buildOutline(rows)) as { text: string; num?: string; deleted?: boolean }[];
      const doc = rows
        .map((r, k) => (r.deleted ? "" : `[¶${r.i}]${nums[k]?.num ? ` (${nums[k].num})` : ""} ${r.text}`))
        .filter(Boolean)
        .join("\n");
      // Word chưa có API đọc sửa đổi thì nói rõ, để người soát trả could_not_verify chứ không phải "không có sửa đổi đó"
      let tc = "(không đọc được: Word trên máy chưa hỗ trợ đọc sửa đổi — căn cứ loại tracked_change trả could_not_verify)";
      if (supports("1.6")) {
        const list = ctx.document.body.getTrackedChanges();
        list.load("items/type,items/text,items/author");
        await ctx.sync();
        tc = list.items
          .slice(0, 200)
          .map((t, k) => `${k}. ${t.type} · ${t.author}: “${truncate(t.text, 200)}”`)
          .join("\n");
        if (list.items.length > 200) tc += `\n[DANH SÁCH BỊ CẮT: còn ${list.items.length - 200} sửa đổi không đưa — không thấy thì trả could_not_verify]`;
      }
      return { document: doc, tracked: tc };
    }));
    queue = read.catch(() => undefined);
    const { document, tracked } = await read;
    const input = buildVerifierInput({ edits, document, tracked, instructions: verifier?.instructions, skills: verifier?.skills });
    if (!input.order.length) return assembleResult(edits, [], [], { dropped });
    if (!verifier) throw new Error("Lượt này không có người soát (add-in không gọi được lượt model riêng): chưa kiểm được căn cứ nào.");
    const text = await verifier.ask(VERIFIER_SYSTEM, input.text);
    return assembleResult(edits, input.order, parseVerdicts(text, input.order.length, input.sources), { dropped, cut: input.cut });
  }

  function execute(name: string, a: any, ctx: { log: UndoEntry[]; verifier?: Verifier }): Promise<any> {
    if (name === EXPLAIN_TOOL) return explainEdits(a, ctx.verifier);
    const job = queue.then(() => executeNow(name, a, ctx));
    queue = job.catch(() => undefined);
    return job;
  }

  async function executeNow(name: string, a: any, ctx: { log: UndoEntry[] }): Promise<any> {
    const u: Undo = { extra: [], inverse: [], fallback: [], touched: [], after: [] };
    if (!WRITE.has(name) || !supports("1.6")) return run(name, a, u);
    const diffable = name !== "word_review_revision" && !UNDO_LABEL[name];
    const before = diffable ? await changeSnaps().catch(() => null) : null;
    try {
      const r = await run(name, a, u);
      // Chữ mới ghi lẫn từ lạ ("DeniedĐến nay", "Collaboratively配合" — vòng thử 3, 24/11/2026): báo check_now để model sửa ngay
      const written = name === "word_replace_text" ? a?.replace : NEW_TEXT_TOOLS.has(name) ? a?.text : undefined;
      if (typeof written === "string" && written && r && typeof r === "object" && r.ok !== false) {
        const c = newTextCheck(written, String(name === "word_replace_text" ? a?.find ?? "" : a?.expect ?? ""));
        if (c) return { ...r, check_now: [c] };
      }
      return r;
    } finally {
      let d = { added: [] as string[], merged: [] as string[], removed: 0 };
      let wants: Wanted[] = [];
      if (before) {
        const after = await changeSnaps().catch(() => before);
        d = diffChanges(before, after);
        wants = await wantedOf(d.added, after, u.touched);
      }
      const tangled = d.merged.length > 0 || d.removed > 0;
      const untracked = !wants.length && !tangled && u.fallback.length > 0;
      const blind = diffable && !before;
      if (wants.length || tangled || u.extra.length || untracked || blind || u.after?.length) {
        ctx.log.push({
          label: (UNDO_LABEL[name] || `Hoàn tác ${name.replace(/^word_/, "")} (${d.added.length + d.merged.length} sửa đổi)`) + (reasonOf(name, a) ? ` — lý do: ${reasonOf(name, a)}` : ""),
          undo: async () => {
            if (blind) throw new Error("không chụp được danh sách sửa đổi trước khi sửa nên không gỡ tự động được; xem thẻ Review của Word");
            for (const x of u.extra.slice().reverse()) await x();
            let missing = 0;
            if (wants.length) missing = wants.length - (await rejectWanted(wants));
            if (tangled) {
              if (!u.inverse.length) {
                throw new Error("một phần thay đổi đã nhập vào sửa đổi có từ trước nên không hoàn tác riêng được; xem thẻ Review của Word");
              }
              for (const x of u.inverse.slice().reverse()) await x();
            } else if (untracked) {
              for (const x of u.fallback.slice().reverse()) await x();
            }
            for (const x of (u.after || []).slice().reverse()) await x();
            if (missing > 0 && !tangled) {
              throw new Error(`${missing} sửa đổi của bước này không còn nguyên (đã được chấp nhận, từ chối hoặc sửa tiếp) nên không gỡ được`);
            }
          },
        });
      }
    }
  }

  async function snapshot() {
    return Word.run(async (ctx) => {
      const { rows } = await loadRows(ctx, true);
      const byPara = outlineByParagraph(buildOutline(rows));
      let chars = 0;
      const out: string[] = [];
      for (let i = 0; i < rows.length && chars < 40000; i++) {
        const t = rows[i].text;
        if (!t.trim()) continue;
        const o = byPara.get(i);
        // Số do Word tự đánh không nằm trong chữ của đoạn: ghi kèm để ứng dụng khác đọc được
        out.push(`[¶${i}] ${o && o.source === "list" && o.label ? o.label + " " : ""}${t}`);
        chars += t.length;
      }
      return { title: "Văn bản Word đang mở", data: { paragraphs: rows.length, text: out.join("\n") } };
    });
  }

  /** Chữ hiện tại của chỗ mà một lệnh ghi sắp sửa, kèm số điều và vài đoạn cùng điều khoản — cho lượt soát lập luận đối chiếu */
  async function reviewContext(name: string, a: any): Promise<string> {
    return Word.run(async (ctx) => {
      if (name === "word_replace_selection") {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        return `Vùng đang chọn (sẽ bị thay):\n${truncate(sel.text, 6000)}`;
      }
      const { rows } = await loadRows(ctx, true);
      const outline = buildOutline(rows);
      const byPara = outlineByParagraph(outline);
      const around = (i: number, mark: string) => {
        const c = clauseAt(outline, i);
        const from = Math.max(c ? c.i : i, i - 3);
        const to = Math.min(c ? c.end : i, i + 3);
        const lines: string[] = [];
        for (let k = from; k <= to; k++) {
          const r = rows[k];
          if (!r || r.deleted || !r.text.trim()) continue;
          const num = byPara.get(k)?.num;
          lines.push(`[¶${k}]${k === i ? ` (${mark})` : ""} ${num ? num + " " : ""}${truncate(r.text, 2500)}`);
        }
        return `${c ? `Thuộc ${clauseName(c)}` : "Ngoài mọi điều khoản"}:\n${lines.join("\n")}`;
      };
      const pick = (index: any, expect: any): number => {
        const loc = locateParagraph(
          rows.map((r) => r.text),
          index,
          expect
        );
        return loc.ok === false ? Math.floor(Number(index)) : loc.index;
      };
      if (name === "word_set_paragraph_text") return around(pick(a.index, a.expect), "đoạn sắp được viết lại");
      if (name === "word_insert_clause") return around(pick(a.after_index, a.expect), "điều khoản mới sẽ chèn sau đoạn này");
      if (name === "word_replace_text") {
        if (a.paragraph_index !== undefined && a.paragraph_index !== null) return around(pick(a.paragraph_index, a.expect), "đoạn có cụm sắp thay");
        const hits = searchRows(rows, String(a.find || ""), { matchCase: !!a.match_case, wholeWord: !!a.whole_word }).hits.slice(0, 4);
        return hits.length ? hits.map((h) => around(h.i, "đoạn có cụm sắp thay")).join("\n\n") : "";
      }
      return "";
    });
  }

  /** Cảnh báo rủi ro: trả về lý do (luôn hỏi người dùng) hoặc null (không rủi ro). */
  const RISK: Record<string, (a: any) => string | null | Promise<string | null>> = {
    ...admin.risk,
    ...form.risk,
    // Xóa đoạn: hộp xác nhận hiện trích nội dung sắp xóa
    word_delete_paragraphs: (a) =>
      Word.run(async (ctx) => {
        const { rows, paras } = await loadRows(ctx);
        const loc = locateParagraph(
          rows.map((r) => r.text),
          a.start,
          a.expect
        );
        if (loc.ok === false) return null; // công cụ sẽ tự từ chối
        const span = a.end === undefined || a.end === null ? 0 : Math.max(0, Number(a.end) - Number(a.start));
        const s = loc.index;
        const e = Math.min(rows.length - 1, s + span);
        // Vùng vắt mép bảng: công cụ sẽ từ chối, không hỏi người dùng trước
        if (await tableSpanProblem(ctx, paras, s, e)) return null;
        const shown = rows.slice(s, Math.min(e + 1, s + 5)).map((r) => `[¶${r.i}] “${truncate(r.text, 80)}”`);
        return `Xóa ${e - s + 1} đoạn (ghi bằng Theo dõi thay đổi): ${shown.join(" · ")}${e - s + 1 > 5 ? ` · … và ${e - s + 1 - 5} đoạn nữa` : ""}`;
      }),
    word_review_revision: (a) => {
      const verb = a.action === "accept" ? "Chấp nhận" : "Từ chối";
      if (!a.all) {
        const n = Array.isArray(a.indexes) ? new Set(a.indexes.map((x: any) => Number(x))).size : 0;
        return n > MANY_REPLACEMENTS ? `${verb} ${n} sửa đổi cùng lúc (không hoàn tác được bằng add-in).` : null;
      }
      return Word.run(async (ctx) => {
        const tc = ctx.document.body.getTrackedChanges();
        tc.load("items/author");
        await ctx.sync();
        if (!String(a.author || "").trim()) return `${verb} TẤT CẢ ${tc.items.length} sửa đổi đang theo dõi trong văn bản (không hoàn tác được bằng add-in).`;
        const who = resolveAuthor(
          tc.items.map((t) => String(t.author)),
          a.author
        );
        if (typeof who !== "string") return null; // không đúng một tác giả: công cụ sẽ tự từ chối
        const n = tc.items.filter((t) => String(t.author) === who).length;
        return `${verb} TẤT CẢ ${n} sửa đổi đang theo dõi của “${who}” (không hoàn tác được bằng add-in).`;
      });
    },
    // Thay nhiều chỗ một lúc, hoặc xóa chữ ở nhiều chỗ
    word_replace_text: (a) => {
      const find = String(a.find || "");
      if (!find || find.length > 255) return null; // công cụ sẽ tự báo lỗi tham số
      if (a.all !== true) return null; // không có all thì nhiều nhất thay một chỗ
      return Word.run(async (ctx) => {
        const hasIndex = a.paragraph_index !== undefined && a.paragraph_index !== null;
        let scope: Word.Paragraph | Word.Body = ctx.document.body;
        if (hasIndex) scope = (await resolvePara(ctx, a.paragraph_index, a.expect, false)).p;
        const res = scope.search(escapeWordSearch(find), { matchCase: !!a.match_case, matchWholeWord: !!a.whole_word });
        res.load("items");
        await ctx.sync();
        const n = Math.min(res.items.length, replaceLimit(a.max_replacements));
        const to = String(a.replace ?? "");
        if (n > MANY_REPLACEMENTS) return `Thay ${n} chỗ cùng lúc trong văn bản: “${truncate(find, 40)}” → “${truncate(to, 40)}”.`;
        if (n > 1 && !to) return `Xóa “${truncate(find, 40)}” ở ${n} chỗ trong văn bản.`;
        return null;
      });
    },
    // Xóa hết chữ của một đoạn
    word_set_paragraph_text: (a) => (String(a.text ?? "").trim() ? null : `Xóa toàn bộ chữ của đoạn ${a.index} (“${truncate(a.expect, 60)}”).`),
    // Thay vùng chọn lớn
    word_replace_selection: () =>
      Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        const ps = sel.paragraphs;
        ps.load("items");
        await ctx.sync();
        const chars = sel.text.length;
        const n = ps.items.length;
        return chars > BIG_SELECTION_CHARS || n >= BIG_SELECTION_PARAS
          ? `Thay toàn bộ vùng đang chọn (${n} đoạn, ${chars} ký tự; bắt đầu bằng “${truncate(sel.text, 60)}”) bằng nội dung mới.`
          : null;
      }),
    // Ghi đè ô bảng đang có dữ liệu
    word_fill_table: (a) =>
      Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load("items/rowCount,items/values");
        await ctx.sync();
        const t = tables.items[Number(a.table_index)];
        if (!t) return null; // công cụ sẽ tự báo không có bảng
        const over = planTableFill(await currentTableValues(ctx, t, t.values), a.values, a.start_row, a.start_col).overwrites;
        if (!over.length) return null;
        const x = over[0];
        return `Ghi đè ${over.length} ô đang có dữ liệu trong bảng ${a.table_index} (ví dụ hàng ${x.row}, cột ${x.col}: “${truncate(x.old, 30)}” → “${truncate(x.value, 30)}”).`;
      }),
  };
  /** Không kiểm tra được mức rủi ro thì coi là rủi ro: hỏi người dùng thay vì lặng lẽ cho qua. */
  const riskOf = async (name: string, a: any): Promise<string | null> => {
    try {
      return await RISK[name](a || {});
    } catch (e: any) {
      return `Không kiểm tra trước được mức rủi ro của thao tác này (${truncate(e?.message || String(e), 120)}).`;
    }
  };

  return {
    host: "Word",
    snapshot,
    reviewContext,
    afterTurn,
    specs: WORD_SPECS.map((sp) => (RISK[sp.name] ? { ...sp, risky: (a: any) => riskOf(sp.name, a) } : sp)),
    execute,
    overview,
    systemPrompt: () => WORD_SYSTEM_PROMPT,
    stateText: (ov) =>
      `Trạng thái hiện tại của văn bản (JSON; headings = mục lục điều khoản cấp cao nhất, clauses = tổng số điều khoản; track_changes là chế độ người dùng đang đặt — sửa đổi của bạn LUÔN được ghi có theo dõi dù chế độ là Off):\n${overviewForPrompt(ov)}`,
  };
}

/** Lời nhắc hệ thống CỐ ĐỊNH của Word (không chứa trạng thái tài liệu, để bộ nhớ đệm lời nhắc có tác dụng) */
export const WORD_SYSTEM_PROMPT = [
  "Bạn là trợ lý AI được tích hợp trong Microsoft Word, có thể thao tác trực tiếp trên văn bản đang mở bằng các công cụ (word_*).",
  "Nguyên tắc:",
  "- Khi người dùng muốn sửa, viết, chèn, định dạng, rà soát… hãy GỌI CÔNG CỤ để làm thật trong văn bản, không chỉ đưa gợi ý.",
  "- ĐỌC ĐÚNG CHỖ: với hợp đồng và văn bản dài, xem mục lục điều khoản trước (word_get_outline; phần headings của trạng thái văn bản đã là mục lục cấp cao nhất), rồi đọc đúng điều cần bằng word_read_clause (theo số điều) hoặc tìm theo nội dung bằng word_find. Chỉ đọc tuần tự toàn văn (word_read_paragraphs) khi việc đòi hỏi rà soát toàn bộ; kết quả có next_start là chưa đọc hết.",
  "- CÂU HỎI TRA CỨU (điều nào nói gì, quy định về … nằm ở đâu): thường chỉ cần một hoặc hai lần gọi công cụ. Đã đọc được điều khoản trả lời đúng câu hỏi thì trả lời ngay, không đọc lan sang điều khác để kiểm thêm, trừ khi câu hỏi cần đối chiếu nhiều điều.",
  "- SỐ ĐIỀU: trường num là số điều khoản đầy đủ (ví dụ 'Điều 5', '5.3', '5.3(a)'), kể cả khi Word tự đánh số và số không nằm trong chữ của đoạn. Với số gõ tay, số đã nằm sẵn ở đầu chữ của đoạn nên num chỉ lặp lại số đó: đó không phải lỗi lặp số trong văn bản. Khi trả lời, dẫn theo số điều kèm số đoạn, ví dụ: Điều 5.3 [¶57]. Đoạn văn đánh số từ 0.",
  "- SỬA ĐÚNG ĐOẠN: chỉ số đoạn dời đi sau mỗi lần chèn hoặc xóa. Mọi công cụ sửa theo đoạn cần expect = khoảng 30 ký tự đầu của đoạn đúng như đã đọc; công cụ từ chối nếu không khớp. Không đoán chỉ số: đọc lại khi bị từ chối.",
  "- SỬA NHỎ: đổi một từ, cụm từ, con số trong một đoạn → word_replace_text kèm paragraph_index (và occurrence nếu cụm đó lặp lại). Viết lại câu → word_set_paragraph_text với toàn bộ câu mới của đoạn (không gõ số điều nếu Word tự đánh số); công cụ chỉ đổi đúng các từ khác đi và giữ định dạng. Thay một cụm ở mọi nơi → word_replace_text với all=true.",
  "- THÊM ĐIỀU KHOẢN: dùng word_insert_clause để khoản mới nối đúng số và cùng kiểu với khoản bên cạnh; word_insert_paragraphs chỉ dùng cho nội dung tự do. Sau khi thêm hoặc xóa điều khoản, kiểm lại các dẫn chiếu tới số điều đã đổi (word_find).",
  "- VĂN BẢN ĐANG CÓ SỬA ĐỔI (Track Changes): nội dung đọc được là bản hiện tại (đã tính các sửa đổi chờ duyệt); đoạn có changed=true đang có sửa đổi chưa duyệt, deleted=true là đoạn đã bị xóa. Muốn biết ai sửa gì: word_list_revisions (đọc hết theo next_start, đối chiếu total). Chỉ chấp nhận hoặc từ chối sửa đổi khi người dùng ra lệnh.",
  "- Chú thích của một đoạn đi kèm đoạn đó trong kết quả đọc (trường notes).",
  WORD_SPECIAL_LINE,
  "- Mọi sửa đổi của bạn được ghi bằng Theo dõi thay đổi để người dùng duyệt, và người dùng có nút Hoàn tác cho cả lượt.",
  "- Rà soát văn bản/hợp đồng: ghi nhận xét đúng chỗ bằng word_add_comment; chỉ sửa trực tiếp khi được yêu cầu.",
  "- Giữ nguyên văn phong, định dạng; không tự bịa số hiệu văn bản, điều khoản hay căn cứ pháp lý.",
  "- Xong việc: trả lời ngắn gọn bằng tiếng Việt: đã sửa/chèn gì, ở điều nào, đoạn nào.",
].join("\n");
