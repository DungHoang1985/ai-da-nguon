/* global Word, Office */
/** Công cụ để AI thao tác trực tiếp trên văn bản Word. Sửa đổi được ghi bằng Theo dõi thay đổi (Track Changes).
 *  Bản 4.0 (Đợt 1A): mục lục điều khoản, tìm và đọc theo điều, chống sửa nhầm đoạn, sửa mức từ, đọc đúng khi văn bản đang có
 *  sửa đổi, chèn điều khoản kế thừa đánh số, Hoàn tác. Căn cứ kỹ thuật: bảng thăm dò trong tasks/dot-1a.md. */
import { HostToolSet, ToolSpec, UndoEntry } from "../core/agent/types";
import { officeErrorHint } from "../core/errors";
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
import { WORD_ADMIN_SPECS, createWordAdminTools } from "./wordAdmin";
import { WORD_FORM_SPECS, createWordFormTools } from "./wordForm";
import { WORD_SPECIAL_LINE } from "../core/toolgroups";
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

const MAX_READ_CHARS = 30000;
/** Thay quá số chỗ này trong một lần gọi thì hỏi người dùng trước */
export const MANY_REPLACEMENTS = 5;
/** Vùng chọn từ mức này trở lên mà bị thay cả khối thì hỏi người dùng trước */
export const BIG_SELECTION_CHARS = 1000;
export const BIG_SELECTION_PARAS = 3;
/** Đoạn không có sửa đổi chờ duyệt, bị viết lại quá tỷ lệ này ở nhiều chỗ: thay cả đoạn cho dễ đọc */
export const WHOLE_PARAGRAPH_RATIO = 0.6;
export const WHOLE_PARAGRAPH_HUNKS = 4;

const STYLE_ENUM = ["Normal", "Heading1", "Heading2", "Heading3", "Title", "ListBullet", "ListNumber", "Quote"];
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
      "Delete paragraphs start..end inclusive (tracked; the user confirms first). Only when the user asks to remove content.",
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

/** Từ 3.6 không còn công tắc tắt Track Changes: mọi sửa đổi của AI trong Word luôn ghi có theo dõi (anh duyệt 15/11/2026). */
export type WordSettings = unknown;

function supports(v: string, set = "WordApi"): boolean {
  try {
    return Office.context.requirements.isSetSupported(set, v);
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

/** Mốc chú thích cuối trang / cuối văn bản nằm trong paragraph.text dưới dạng ký tự U+0002; getText() thì không trả ký tự này
 *  (thăm dò 15/11/2026, tasks/dot-1b.md). Chữ đưa cho AI không kèm mốc; khi sửa đoạn, mốc được coi là chữ "không sống" (như chữ
 *  đã xóa có theo dõi) để lần sửa mức từ không đụng tới, và đoạn có mốc không bao giờ bị thay cả đoạn (sẽ mất chú thích). */
const NOTE_MARK = "\u0002";
const noMarks = (s: string) => (s.includes(NOTE_MARK) ? s.split(NOTE_MARK).join("") : s);
const withNoteSpans = (M: string, spans?: [number, number][]): [number, number][] | undefined => {
  if (!M.includes(NOTE_MARK)) return spans;
  const out: [number, number][] = spans ? spans.slice() : [];
  for (let k = M.indexOf(NOTE_MARK); k >= 0; k = M.indexOf(NOTE_MARK, k + 1)) out.push([k, k + 1]);
  return out;
};

/** Bỏ dấu hết đoạn / hết ô bảng mà getText, getReviewedText trả kèm ở cuối */
const strip = (s: any) => String(s ?? "").replace(/[\r\n\u0007\u000b\u000c]+$/, "");

/** Bản hiện tại của một đoạn từ getText / getReviewedText: bỏ dấu hết đoạn; đoạn nằm trong ô bảng còn kèm một dấu tab ở cuối
 *  (thăm dò 15/11/2026: "12 tháng\t") — bỏ nếu chuỗi đối chiếu `ref` không kết thúc bằng tab */
const curOf = (value: any, ref: string) => {
  const t = strip(value);
  return t.endsWith("\t") && !ref.endsWith("\t") ? t.slice(0, -1) : t;
};

/** Giá trị HIỆN TẠI của các ô một bảng. table.values lẫn cả chữ đã bị xóa có theo dõi ("1224 tháng" thay vì "12 tháng" — thăm dò
 *  15/11/2026), nên khi văn bản đang có sửa đổi phải đọc từng ô theo bản hiện tại. Bảng có ô gộp đọc không được thì trả values gốc. */
async function currentTableValues(ctx: Word.RequestContext, t: Word.Table, values: string[][], maxRows = Infinity): Promise<string[][]> {
  if (!supports("1.6")) return values;
  const tc = ctx.document.body.getTrackedChanges();
  tc.load("items/type");
  await ctx.sync();
  if (!tc.items.length) return values;
  try {
    const rows = Math.min(values.length, maxRows);
    const res = values.slice(0, rows).map((row, r) => row.map((_, c) => t.getCell(r, c).body.getReviewedText("Current")));
    await ctx.sync();
    return values.map((row, r) =>
      r < rows
        ? row.map((_, c) =>
            String(res[r][c].value ?? "")
              .replace(/[\r\n\t\u0007]+$/, "")
              .replace(/\r/g, "\n")
          )
        : row
    );
  } catch {
    return values;
  }
}

/** Một đoạn: text = bản hiện tại (không gồm chữ đã xóa có theo dõi); markup = chuỗi Word dùng để tìm kiếm (gồm cả chữ đã xóa) */
interface Row extends ParaRow {
  markup: string;
  id?: string;
  /** Chú thích cuối trang / cuối văn bản neo vào đoạn này */
  notes?: string[];
}

interface Loaded {
  paras: Word.ParagraphCollection;
  rows: Row[];
  /** Văn bản đang có sửa đổi chờ duyệt */
  hasChanges: boolean;
  /** Số chú thích cuối trang và cuối văn bản (chỉ có khi nạp kèm mục lục) */
  footnotes?: number;
  endnotes?: number;
}

/** Nạp mọi đoạn. withList = nạp thêm nhãn số do Word tự đánh, cấp tiêu đề, bảng (cho mục lục điều khoản).
 *  Thăm dò 15/11/2026: paragraph.text lẫn chữ đã xóa có theo dõi → khi văn bản có sửa đổi, đọc bản hiện tại bằng getText() từng đoạn. */
function loadRows(ctx: Word.RequestContext, withList = false): Promise<Loaded> {
  return loadRowsOf(ctx, ctx.document.body, withList);
}

/** Như loadRows cho một thân văn bản bất kỳ. foreign = thân của tài liệu ẩn dựng từ tệp khác (so sánh phiên bản): không nạp
 *  định danh đoạn và chú thích. */
async function loadRowsOf(ctx: Word.RequestContext, body: Word.Body, withList = false, foreign = false): Promise<Loaded> {
  const paras = body.paragraphs;
  const props = ["items/text", "items/style"];
  const list = withList && supports("1.3");
  if (list) {
    props.push(
      "items/outlineLevel",
      "items/tableNestingLevel",
      "items/listItemOrNullObject/listString",
      "items/listItemOrNullObject/level",
      "items/listOrNullObject/id"
    );
  }
  const ids = supports("1.6") && !foreign;
  if (ids) props.push("items/uniqueLocalId");
  paras.load(props.join(","));
  let tc: Word.TrackedChangeCollection | null = null;
  if (supports("1.6")) {
    tc = body.getTrackedChanges();
    tc.load("items/type");
  }
  // Chú thích cuối trang, cuối văn bản: nạp cùng lượt để mọi kết quả đọc đều kèm chú thích của đoạn
  const withNotes = withList && !foreign && supports("1.5");
  const fn = withNotes ? body.footnotes : null;
  const en = withNotes ? body.endnotes : null;
  fn?.load("items/body/text");
  en?.load("items/body/text");
  await ctx.sync();
  const hasChanges = !!tc && tc.items.length > 0;
  let current: string[] | null = null;
  if (hasChanges && supports("1.4")) {
    const res = paras.items.map((p) => (supports("1.7") ? p.getText() : p.getReviewedText("Current")));
    await ctx.sync();
    current = res.map((r, i) => curOf(r.value, paras.items[i].text));
  }
  const rows: Row[] = paras.items.map((p, i) => {
    const markup = p.text;
    const text = noMarks(current ? current[i] : markup);
    const row: Row = { i, text, markup, style: p.style };
    if (current && !text && noMarks(markup).trim()) row.deleted = true;
    if (ids) row.id = p.uniqueLocalId;
    if (list) {
      row.outlineLevel = Number(p.outlineLevel) || 10;
      row.inTable = p.tableNestingLevel > 0;
      const li = p.listItemOrNullObject;
      if (!li.isNullObject) {
        row.listString = li.listString;
        row.listLevel = li.level;
        if (!p.listOrNullObject.isNullObject) row.listId = p.listOrNullObject.id;
      }
    }
    return row;
  });
  const notes = [...(fn?.items || []).map((n) => ({ n, label: "Chú thích cuối trang" })), ...(en?.items || []).map((n) => ({ n, label: "Chú thích cuối văn bản" }))];
  if (notes.length && ids) {
    try {
      const anchors = notes.map((x) => {
        const p = x.n.reference.paragraphs.getFirstOrNullObject();
        p.load("uniqueLocalId");
        return p;
      });
      await ctx.sync();
      const byId = new Map<string, Row>();
      rows.forEach((r) => r.id && !byId.has(r.id) && byId.set(r.id, r));
      notes.forEach((x, k) => {
        const row = anchors[k].isNullObject ? undefined : byId.get(anchors[k].uniqueLocalId);
        const t = noMarks(String(x.n.body.text ?? "")).replace(/\r/g, " ").trim();
        if (row && t) (row.notes = row.notes || []).push(`${x.label}: ${truncate(t, 600)}`);
      });
    } catch {
      /* không neo được chú thích vào đoạn: vẫn đọc được bằng word_read_parts */
    }
  }
  return { paras, rows, hasChanges, ...(withNotes ? { footnotes: fn?.items.length || 0, endnotes: en?.items.length || 0 } : {}) };
}

/** Số điều khoản của từng đoạn mở đầu điều khoản (để mọi kết quả đọc đều kèm số điều) */
function withNums(rows: Row[], outline: OutlineItem[]) {
  const byPara = outlineByParagraph(outline);
  return rows.map((r) => ({ text: r.text, style: String(r.style ?? ""), num: byPara.get(r.i)?.num || undefined, deleted: r.deleted, notes: r.notes }));
}

/** Định vị đoạn theo chỉ số + mấy chữ đầu AI đã đọc; lệch mà không chắc thì từ chối (không sửa nhầm đoạn) */
async function resolvePara(ctx: Word.RequestContext, index: any, expect?: string, needExpect = true) {
  if (needExpect && !String(expect ?? "").trim()) throw new Error(missingExpect(index));
  const loaded = await loadRows(ctx);
  const loc = locateParagraph(
    loaded.rows.map((r) => r.text),
    index,
    expect
  );
  if (loc.ok === false) {
    const cand = loc.candidates.map((k) => `[¶${k}] “${truncate(loaded.rows[k].text, 60)}”`).join("; ");
    throw new Error(loc.reason + (cand ? ` Ứng viên: ${cand}` : ""));
  }
  return { ...loaded, p: loaded.paras.items[loc.index], row: loaded.rows[loc.index], index: loc.index, moved: loc.moved };
}

/** Công cụ ghi chỉ đoạn bằng số mà không kèm expect: từ chối (chỉ số dời sau mỗi lần chèn, xóa — soát mã độc lập 15/11/2026, mục T4) */
const missingExpect = (index: any) =>
  `Thiếu expect: khi chỉ một đoạn bằng số (đoạn ${index}) phải kèm expect = khoảng 30 ký tự đầu của đoạn đó đúng như đã đọc, để công cụ không sửa nhầm đoạn khi chỉ số đã dời.`;

/** Đoạn có sửa đổi chờ duyệt không, và các khúc đã bị xóa nằm ở đâu trong chuỗi Word dùng để tìm (vị trí thật, cần WordApiDesktop 1.4) */
async function pendingInfo(ctx: Word.RequestContext, p: Word.Paragraph): Promise<{ count: number; spans?: [number, number][] }> {
  const tc = p.getTrackedChanges();
  tc.load("items/type");
  await ctx.sync();
  const dels = tc.items.filter((t) => t.type === "Deleted");
  let spans: [number, number][] | undefined;
  if (dels.length && supports("1.4", "WordApiDesktop")) {
    const base = p.getRange("Content") as any;
    base.load("start");
    const rs = dels.map((t) => {
      const x = t.getRange() as any;
      x.load("start,end");
      return x;
    });
    await ctx.sync();
    spans = rs.map((x) => [x.start - base.start, x.end - base.start] as [number, number]);
  }
  return { count: tc.items.length, spans };
}

const movedNote = (asked: any, r: { index: number; moved: boolean }) =>
  r.moved ? ` (đoạn này nay là [¶${r.index}], không còn là ${asked}: chỉ số đã dời do chèn hoặc xóa)` : "";

/** Chỉ giữ kết quả tìm nằm trọn trong chữ đang sống (search của Word tìm cả trong chữ đã xóa có theo dõi) */
async function liveHits(ctx: Word.RequestContext, hits: Word.Range[], hasChanges: boolean): Promise<Word.Range[]> {
  if (!hasChanges || !hits.length || !supports("1.4")) return hits;
  const cur = hits.map((h) => h.getReviewedText("Current"));
  await ctx.sync();
  return hits.filter((h, k) => strip(cur[k].value) === strip(h.text));
}

/** Áp kế hoạch sửa mức từ vào một đoạn (hoặc một vùng chữ tìm thấy): tìm mọi chỗ trong một lần sync, sửa trong một lần sync, từ phải sang trái */
async function applyPlan(ctx: Word.RequestContext, p: Word.Paragraph | Word.Range, plan: EditPlan): Promise<void> {
  const locs: Locator[] = [];
  for (const op of plan.ops) {
    if (op.kind === "insertStart") continue;
    locs.push(op.from);
    if (op.kind === "replace" && op.to) locs.push(op.to);
  }
  const found = locs.map((l) => {
    const h = p.search(escapeWordSearch(l.search), { matchCase: true });
    h.load("items");
    return h;
  });
  await ctx.sync();
  const picked = new Map<Locator, Word.Range>();
  const off: number[] = [];
  locs.forEach((l, k) => {
    if (found[k].items.length === l.total) picked.set(l, found[k].items[l.occurrence]);
    else off.push(k);
  });
  if (off.length) {
    // Word đếm khác cách tính (ký tự đặc biệt): đối chiếu bằng vị trí thật của từng kết quả (WordApiDesktop 1.4)
    if (!supports("1.4", "WordApiDesktop")) {
      const l = locs[off[0]];
      throw new Error(`Không định vị chắc chắn được “${truncate(l.search, 40)}” trong đoạn (Word tìm ra ${found[off[0]].items.length} chỗ, tính ra ${l.total}).`);
    }
    const base = p.getRange("Content") as any;
    base.load("start");
    off.forEach((k) => found[k].items.forEach((r) => (r as any).load("start")));
    await ctx.sync();
    for (const k of off) {
      const l = locs[k];
      const hit = found[k].items.find((r) => (r as any).start === base.start + l.at);
      if (!hit) throw new Error(`Không định vị chắc chắn được “${truncate(l.search, 40)}” trong đoạn.`);
      picked.set(l, hit);
    }
  }
  for (const op of plan.ops.slice().reverse()) {
    if (op.kind === "insertStart") {
      p.insertText(op.text, "Start");
      continue;
    }
    const from = picked.get(op.from) as Word.Range;
    if (op.kind === "replace") {
      const r = op.to ? from.expandTo(picked.get(op.to) as Word.Range) : from;
      if (op.text) r.insertText(op.text, "Replace");
      else r.delete();
    } else from.insertText(op.text, op.kind === "insertAfter" ? "After" : "Before");
  }
  await ctx.sync();
}

/** Danh sách sửa đổi đang theo dõi (kèm khóa nhận diện), theo thứ tự trong văn bản */
async function changeSnaps(): Promise<ChangeSnap[]> {
  return Word.run(async (ctx) => {
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type,items/author,items/date,items/text");
    await ctx.sync();
    return tc.items.map((t) => snapChange(t as any));
  });
}

/** Nạp danh sách sửa đổi hiện có; sửa đổi nào có khóa thuộc `paraKeys` thì nạp thêm định danh đoạn chứa nó */
async function loadChanges(ctx: Word.RequestContext, paraKeys: Set<string>) {
  const tc = ctx.document.body.getTrackedChanges();
  tc.load("items/type,items/author,items/date,items/text");
  await ctx.sync();
  const cur: { key: string; para?: string }[] = tc.items.map((t) => ({ key: changeKey(t as any) }));
  const need = cur.map((c, i) => (paraKeys.has(c.key) ? i : -1)).filter((i) => i >= 0);
  if (need.length) {
    const ps = need.map((i) => {
      const p = tc.items[i].getRange().paragraphs.getFirst();
      p.load("uniqueLocalId");
      return p;
    });
    await ctx.sync();
    need.forEach((i, k) => (cur[i].para = ps[k].uniqueLocalId));
  }
  return { tc, cur };
}

/** Hoàn tác: từ chối đúng các sửa đổi đã ghi (theo khóa, và theo đoạn khi biết). Trả số sửa đổi đã từ chối. */
async function rejectWanted(wants: Wanted[]): Promise<number> {
  const paraKeys = new Set(wants.filter((w) => w.para).map((w) => w.key));
  try {
    return await Word.run(async (ctx) => {
      const { tc, cur } = await loadChanges(ctx, paraKeys);
      const idx = pickWanted(cur, wants);
      idx.forEach((i) => tc.items[i].reject());
      await ctx.sync();
      return idx.length;
    });
  } catch {
    // Từ chối cả loạt lỗi (một sửa đổi nằm trong sửa đổi khác): làm từng cái một, mỗi lần tính lại từ trạng thái hiện tại
    let done = 0;
    const left = wants.slice();
    for (let guard = 0; guard < 500 && left.length; guard++) {
      const more = await Word.run(async (ctx) => {
        const { tc, cur } = await loadChanges(ctx, paraKeys);
        const idx = pickWanted(cur, left);
        if (!idx.length) return false;
        const hit = cur[idx[0]];
        tc.items[idx[0]].reject();
        await ctx.sync();
        let k = left.findIndex((w) => w.key === hit.key && w.para === hit.para);
        if (k < 0) k = left.findIndex((w) => w.key === hit.key);
        left.splice(k, 1);
        return true;
      });
      if (!more) break;
      done++;
    }
    return done;
  }
}

/** Thêm vào câu "không tìm thấy" khi văn bản có sửa đổi chờ duyệt: tìm kiếm chỉ chạy trên chữ HIỆN TẠI */
const NOT_FOUND_PENDING =
  " LƯU Ý: văn bản đang có sửa đổi chờ duyệt và tìm kiếm chỉ chạy trên chữ hiện tại (sau sửa). Cụm này có thể nằm trong chữ GỐC đã bị sửa hoặc xóa: xem các sửa đổi bằng word_list_revisions, hoặc đọc bản gốc bằng word_read_paragraphs với view = \"original\". Chưa kiểm hai nơi đó thì không được kết luận là văn bản không có nội dung này.";

/** Kết quả tìm kiếm nằm trong đoạn đang có sửa đổi chờ duyệt: gắn changed, chữ gốc của cả đoạn (before, cắt gọn) và người sửa,
 *  để AI nêu được cả hai bản mà không cần gọi thêm công cụ. Trả về số kết quả được gắn. */
async function markPendingHits(ctx: Word.RequestContext, paras: Word.ParagraphCollection, hits: { i: number }[], current?: (i: number) => string): Promise<number> {
  if (!supports("1.6") || !hits.length) return 0;
  const idx = [...new Set(hits.map((h) => h.i))].filter((i) => paras.items[i]);
  const tcs = idx.map((i) => {
    const t = paras.items[i].getTrackedChanges();
    t.load("items/type,items/author");
    return t;
  });
  await ctx.sync();
  const hot = idx.map((i, k) => ({ i, k })).filter((x) => tcs[x.k].items.length > 0);
  if (!hot.length) return 0;
  const orig = hot.map((x) => paras.items[x.i].getReviewedText("Original"));
  await ctx.sync();
  const info = new Map(hot.map((x, j) => [x.i, { before: beforeExcerpt(noMarks(strip(orig[j].value)), current ? current(x.i) : "", 400), by: [...new Set(tcs[x.k].items.map((t) => String(t.author || "").trim()).filter(Boolean))].slice(0, 3).join(", ") }]));
  let n = 0;
  for (const h of hits as any[]) {
    const x = info.get(h.i);
    if (!x) continue;
    h.changed = true;
    h.before = x.before;
    if (x.by) h.changed_by = x.by;
    n++;
  }
  return n;
}

type PendingRow = { text: string; style: string; num?: string; changed?: boolean; deleted?: boolean; notes?: string[]; before?: string; changed_by?: string };

/** Đánh dấu các đoạn đang có sửa đổi chờ duyệt trong cửa sổ [start, end): changed, chữ gốc trước sửa (before), người sửa (changed_by).
 *  original = người gọi muốn xem bản gốc: text thành chữ gốc, không kèm before. Cần WordApi 1.6; máy cũ thì chỉ biết đoạn có chữ bị xóa. */
async function annotatePending(ctx: Word.RequestContext, loaded: { rows: any[]; paras: Word.ParagraphCollection }, all: PendingRow[], start: number, end: number, original = false) {
  const { rows, paras } = loaded;
  const s0 = Math.max(0, start);
  const e0 = Math.min(rows.length, end);
  if (!supports("1.6")) {
    for (let i = s0; i < e0; i++) if (rows[i].deleted || noMarks(rows[i].markup) !== rows[i].text) all[i].changed = true;
    return;
  }
  const win = paras.items.slice(s0, e0);
  const tcs = win.map((p) => {
    const t = p.getTrackedChanges();
    t.load("items/type,items/author");
    return t;
  });
  await ctx.sync();
  const hit: number[] = [];
  win.forEach((_, k) => {
    const row = rows[s0 + k];
    if (tcs[k].items.length || row.deleted || noMarks(row.markup) !== row.text) hit.push(k);
  });
  if (!hit.length) return;
  const orig = hit.map((k) => win[k].getReviewedText("Original"));
  await ctx.sync();
  hit.forEach((k, j) => {
    const row = all[s0 + k];
    const before = noMarks(strip(orig[j].value));
    row.changed = true;
    if (original) {
      row.text = before;
      row.deleted = false;
      return;
    }
    row.before = before;
    const by = [...new Set(tcs[k].items.map((t) => String(t.author || "").trim()).filter(Boolean))];
    if (by.length) row.changed_by = by.slice(0, 3).join(", ");
  });
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
      const { rows, paras, footnotes, endnotes } = await loadRows(ctx, true);
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
        selection: truncate(sel.text, 1500),
        track_changes: mode,
        tracked_changes: changes,
        comments,
        ...(footnotes ? { footnotes } : {}),
        ...(endnotes ? { endnotes } : {}),
      };
    });
  }

  /** Vùng đích của định dạng / nhận xét: đúng một trong ba cách chỉ (find, paragraph_index, use_selection) */
  async function targetRanges(ctx: Word.RequestContext, a: any): Promise<{ ranges: Word.Range[]; paras: Word.Paragraph[] }> {
    const hasIndex = a.paragraph_index !== undefined && a.paragraph_index !== null;
    if (a.find) {
      if (String(a.find).length > 255) throw new Error("Chuỗi tìm quá dài (tối đa 255 ký tự).");
      let scope: Word.Paragraph | Word.Body = ctx.document.body;
      let hasChanges = true;
      if (hasIndex) {
        const r = await resolvePara(ctx, a.paragraph_index, a.expect);
        scope = r.p;
        hasChanges = r.hasChanges;
      }
      const res = scope.search(escapeWordSearch(a.find), { matchCase: true });
      res.load("items/text");
      await ctx.sync();
      const live = await liveHits(ctx, res.items.slice(0, 200), hasChanges);
      if (!live.length) throw new Error(`Không tìm thấy "${truncate(a.find, 80)}"${hasIndex ? ` trong đoạn ${a.paragraph_index}` : ""}.`);
      const ps = live.map((r) => r.paragraphs.getFirst());
      return { ranges: live.slice(0, 50), paras: ps.slice(0, 50) };
    }
    if (hasIndex) {
      const r = await resolvePara(ctx, a.paragraph_index, a.expect);
      return { ranges: [r.p.getRange("Content")], paras: [r.p] };
    }
    if (a.use_selection !== true) {
      throw new Error("Phải chỉ rõ đúng một vị trí: find (chuỗi chữ), paragraph_index (số đoạn) hoặc use_selection=true (vùng người dùng đang chọn).");
    }
    const sel = ctx.document.getSelection();
    const ps = sel.paragraphs;
    ps.load("items");
    await ctx.sync();
    return { ranges: [sel], paras: ps.items };
  }

  /** Chỉ số đoạn của một loạt đoạn (theo định danh đoạn) */
  const indexOfIds = (rows: Row[]) => {
    const m = new Map<string, number>();
    rows.forEach((r) => r.id && !m.has(r.id) && m.set(r.id, r.i));
    return m;
  };

  type UndoStep = () => Promise<void>;
  /** extra = việc hoàn tác luôn phải làm (gỡ nhận xét…); inverse = cách hoàn tác thay thế, chỉ dùng khi sửa đổi của công cụ
   *  đã nhập vào một sửa đổi có từ trước nên không từ chối riêng được */
  interface Undo {
    extra: UndoStep[];
    inverse: UndoStep[];
    /** Dùng khi công cụ không sinh sửa đổi được theo dõi nào (ví dụ đổi cấp đánh số) */
    fallback: UndoStep[];
    /** Việc dọn luôn làm SAU khi đã từ chối các sửa đổi (ví dụ gỡ đoạn rỗng cạnh bảng mà Word để lại khi từ chối phần chèn) */
    after?: UndoStep[];
    /** Định danh các đoạn công cụ đã đụng: để phân biệt sửa đổi của AI với sửa đổi trùng hệt người dùng tự làm ở đoạn khác */
    touched: string[];
  }

  /** Đưa chữ của một đoạn (tìm theo định danh) về câu cũ bằng sửa mức từ */
  const restoreText = (id: string, prev: string): UndoStep => () =>
    Word.run(async (ctx) =>
      tracked(ctx, async () => {
        const p = ctx.document.getParagraphByUniqueLocalId(id);
        p.load("text");
        const cur = supports("1.7") ? p.getText() : p.getReviewedText("Current");
        await ctx.sync();
        const C = noMarks(curOf(cur.value, p.text));
        if (C === prev) return;
        const info = await pendingInfo(ctx, p);
        const plan = planParagraphEdit(p.text, C, prev, withNoteSpans(p.text, info.spans));
        if (typeof plan === "string") throw new Error(`không đưa được đoạn “${truncate(C, 40)}” về câu cũ: ${plan}`);
        await applyPlan(ctx, p, plan);
        const now = supports("1.7") ? p.getText() : p.getReviewedText("Current");
        await ctx.sync();
        if (noMarks(curOf(now.value, prev)) !== prev) throw new Error(`đoạn “${truncate(prev, 40)}” chưa về đúng câu cũ; xem thẻ Review của Word`);
      })
    );

  /** Gỡ hẳn các đoạn AI đã chèn (tìm theo định danh); đoạn không còn thì bỏ qua. Trước khi gỡ đối chiếu chữ đầu đoạn với chữ AI đã
   *  chèn: định danh đoạn trỏ sang đoạn khác thì không xóa. */
  const removeParas = (ids: string[], texts: string[]): UndoStep => async () => {
    let skipped = 0;
    for (let k = ids.length - 1; k >= 0; k--) {
      await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        const mode = ctx.document.changeTrackingMode;
        try {
          const p = ctx.document.getParagraphByUniqueLocalId(ids[k]);
          p.load("text");
          await ctx.sync();
          const head = String(texts[k] ?? "").slice(0, 20);
          if (head && !p.text.includes(head)) {
            skipped++;
            return;
          }
          ctx.document.changeTrackingMode = "Off" as any;
          p.delete();
          await ctx.sync();
        } catch {
          /* đoạn đã không còn */
        } finally {
          ctx.document.changeTrackingMode = mode;
          await ctx.sync();
        }
      });
    }
    if (skipped) throw new Error(`${skipped} đoạn đã chèn nay mang nội dung khác nên không gỡ; xem thẻ Review của Word`);
  };

  const review = createWordReviewTools({
    supports,
    loadRows: (ctx) => loadRows(ctx, true),
    loadRowsOf: (ctx, body) => loadRowsOf(ctx, body, true, true),
  });

  const admin = createWordAdminTools({ supports, tracked, resolvePara });
  const form = createWordFormTools({ supports, tracked, loadRows: (ctx) => loadRows(ctx), tableValues: (ctx, t, v) => currentTableValues(ctx, t, v) });

  async function run(name: string, a: any, u: Undo): Promise<any> {
    const extra = u.extra;
    if (review.has(name)) return review.run(name, a);
    if (admin.has(name)) return admin.run(name, a, u);
    if (form.has(name)) return form.run(name, a);
    switch (name) {
      case "word_get_overview":
        return overview();

      case "word_get_outline":
        return Word.run(async (ctx) => {
          const { rows } = await loadRows(ctx, true);
          const all = buildOutline(rows);
          const maxLevel = a.max_level === undefined || a.max_level === null ? 99 : Math.max(0, Number(a.max_level) || 0);
          const items = all.filter((o) => o.level <= maxLevel);
          const from = Math.max(0, Number(a.start) || 0);
          const list = items.slice(from).map((o) => ({
            i: o.i,
            num: o.num,
            lvl: o.level,
            title: truncate(o.title, o.level === 0 ? 90 : 60),
            end: o.end,
          }));
          const { picked, truncated } = takeWithinBudget(list, RESULT_JSON_BUDGET - 600);
          const next = from + picked.length;
          const perLevel: Record<string, number> = {};
          all.forEach((o) => (perLevel[o.level] = (perLevel[o.level] || 0) + 1));
          return {
            ok: true,
            total_paragraphs: rows.length,
            total_items: items.length,
            items_per_level: perLevel,
            ...(next < items.length ? { next_start: next } : {}),
            ...(truncated ? { truncated: true, note: "Mục lục dài: gọi lại với start = next_start, hoặc max_level = 0 để chỉ xem cấp cao nhất." } : {}),
            ...(all.length ? {} : { note: "Văn bản không có điều khoản đánh số hay tiêu đề nào: đọc bằng word_read_paragraphs hoặc tìm bằng word_find." }),
            items: picked,
          };
        });

      case "word_find":
        return Word.run(async (ctx) => {
          const q = String(a.query || "").trim();
          if (!q) throw new Error("Thiếu chuỗi cần tìm.");
          // Tìm trên bản hiện tại của từng đoạn (không dùng search của Word: nó tìm cả trong chữ đã xóa có theo dõi
          // và không thấy cụm bị chữ đã xóa chen giữa — thăm dò 15/11/2026)
          const { rows, paras, hasChanges } = await loadRows(ctx, true);
          const outline = buildOutline(rows);
          const max = Math.min(100, Math.max(1, Math.floor(Number(a.max)) || 30));
          let res = searchRows(rows, q, { matchCase: !!a.match_case, wholeWord: !!a.whole_word });
          let loose = false;
          if (!res.total && !a.match_case) {
            res = searchRows(rows, q, { wholeWord: !!a.whole_word, ignoreAccents: true });
            loose = res.total > 0;
          }
          const hits = res.hits.map((h) => {
            const t = rows[h.i].text;
            const s = Math.max(0, h.at - 70);
            const excerpt = (s > 0 ? "…" : "") + t.slice(s, h.at) + "«" + t.slice(h.at, h.at + h.len) + "»" + truncate(t.slice(h.at + h.len), 110);
            return { i: h.i, clause: clauseName(clauseAt(outline, h.i)), ...(h.n > 1 ? { n: h.n } : {}), text: excerpt };
          });
          if (!res.total) {
            // Không đoạn nào chứa nguyên cụm: tìm theo từng từ, đoạn nào có nhiều từ của câu hỏi nhất thì đứng trước
            const w = searchRowsByWords(rows, q);
            if (w.hits.length) {
              const list = w.hits.slice(0, Math.min(max, 12)).map((h) => ({
                i: h.i,
                clause: clauseName(clauseAt(outline, h.i)),
                words: `${h.matched}/${w.words.length}`,
                text: truncate(h.sentence, 260),
              }));
              if (hasChanges) await markPendingHits(ctx, paras, list, (i) => rows[i].text);
              const { picked: some } = takeWithinBudget(list, RESULT_JSON_BUDGET - 600);
              const pendingSome = some.filter((h: any) => h.changed).length;
              return {
                ok: true,
                found: 0,
                mode: "by_words",
                note: `Không đoạn nào chứa nguyên cụm "${truncate(q, 60)}". Dưới đây là các đoạn chứa nhiều từ của cụm đó nhất (words = số từ có mặt), xếp từ khớp nhất.`,
                paragraphs: w.hits.length,
                shown: some.length,
                ...(pendingSome ? { pending_note: PENDING_NOTE } : {}),
                hits: some,
              };
            }
          }
          const cand = hits.slice(0, max);
          if (hasChanges) await markPendingHits(ctx, paras, cand, (i) => rows[i].text);
          const { picked } = takeWithinBudget(cand, RESULT_JSON_BUDGET - 600);
          const pendingHits = picked.filter((h: any) => h.changed).length;
          return {
            ok: true,
            found: res.total,
            paragraphs: hits.length,
            shown: picked.length,
            ...(picked.length < hits.length ? { note: `Còn ${hits.length - picked.length} đoạn khớp chưa hiện: thu hẹp chuỗi tìm hoặc tăng max.` } : {}),
            ...(loose ? { ignoring_accents: true } : {}),
            ...(pendingHits ? { pending_note: PENDING_NOTE } : {}),
            ...(res.total
              ? {}
              : {
                  message:
                    `Không tìm thấy "${truncate(q, 80)}". Tìm kiếm theo đúng chữ: thử từ khóa khác hoặc ngắn hơn, hoặc xem mục lục bằng word_get_outline.` +
                    (hasChanges ? NOT_FOUND_PENDING : ""),
                }),
            hits: picked,
          };
        });

      case "word_read_clause":
        return Word.run(async (ctx) => {
          const { rows, paras, hasChanges } = await loadRows(ctx, true);
          const outline = buildOutline(rows);
          const matches = findClause(outline, String(a.number || ""));
          if (!matches.length) {
            const top = outline.filter((o) => o.level === 0 && o.num).map((o) => o.num);
            return {
              ok: false,
              message: `Không có điều khoản "${truncate(a.number, 40)}" trong văn bản.`,
              top_level: top.slice(0, 80),
              hint: "Xem word_get_outline để lấy đúng số, hoặc word_find để tìm theo nội dung.",
            };
          }
          if (matches.length > 1) {
            return {
              ok: false,
              message: `Có ${matches.length} điều khoản mang số "${truncate(a.number, 40)}" (văn bản đánh số lại theo từng phần). Chọn một rồi đọc bằng word_read_paragraphs(start=i, end=end+1).`,
              candidates: matches.slice(0, 12).map((o) => ({
                i: o.i,
                end: o.end,
                num: o.num,
                title: truncate(o.title, 60),
                within: o.parent >= 0 ? clauseName(outline[o.parent]) : "",
              })),
            };
          }
          const it = matches[0];
          const all = withNums(rows, outline) as PendingRow[];
          if (hasChanges) await annotatePending(ctx, { rows, paras }, all, it.i, Math.min(it.end + 1, it.i + 300));
          const r: any = packParagraphs(all, it.i, it.end + 1);
          // next_start chỉ có nghĩa khi điều khoản chưa đọc hết
          if (r.read_to >= it.end) delete r.next_start;
          return { ...r, clause: it.num || it.title, title: it.title, from: it.i, to: it.end };
        });

      case "word_read_paragraphs":
        return Word.run(async (ctx) => {
          if (a.selection_only) {
            const sel = ctx.document.getSelection();
            sel.load("text");
            const ps = sel.paragraphs;
            ps.load("items/text,items/style");
            const cur = supports("1.4") ? sel.getReviewedText(a.view === "original" ? "Original" : "Current") : null;
            await ctx.sync();
            return { ok: true, selection: truncate(cur ? strip(cur.value) : sel.text, MAX_READ_CHARS), paragraphs: ps.items.length };
          }
          const { rows, paras, hasChanges } = await loadRows(ctx, true);
          const outline = buildOutline(rows);
          const all = withNums(rows, outline) as { text: string; style: string; num?: string; changed?: boolean; deleted?: boolean; notes?: string[] }[];
          // Ngân sách tính theo độ dài JSON thật; next_start và truncated nằm đầu kết quả (không bị cắt mất)
          const start = Math.max(0, Number(a.start) || 0);
          // Không nạp sửa đổi cho quá 300 đoạn một lần: kết quả cũng chỉ trả vài chục tới vài trăm đoạn theo ngân sách
          const end = Math.min(rows.length, Number(a.end) || start + 200, start + 300);
          if (hasChanges) await annotatePending(ctx, { rows, paras }, all, start, end, a.view === "original");
          const r = packParagraphs(all, start, end);
          return a.view === "original" ? { ...r, view: "original" } : r;
        });

      case "word_replace_text":
        return Word.run(async (ctx) => {
          const find = String(a.find || "");
          if (!find) throw new Error("Thiếu chuỗi cần tìm.");
          if (escapeWordSearch(find).length > 255) throw new Error("Chuỗi tìm quá dài (>255 ký tự). Dùng word_set_paragraph_text để sửa cả đoạn.");
          const replace = String(a.replace ?? "");
          if (/[\r\n]/.test(replace)) throw new Error("Chuỗi thay không được có xuống dòng: dùng word_insert_paragraphs để thêm đoạn.");
          return tracked(ctx, async () => {
            const hasIndex = a.paragraph_index !== undefined && a.paragraph_index !== null;
            let scope: Word.Paragraph | Word.Body = ctx.document.body;
            let where = "";
            let loaded: Loaded;
            if (hasIndex) {
              const r = await resolvePara(ctx, a.paragraph_index, a.expect);
              scope = r.p;
              loaded = r;
              where = ` trong đoạn ${r.index}${movedNote(a.paragraph_index, r)}`;
            } else loaded = await loadRows(ctx);
            const res = scope.search(escapeWordSearch(find), { matchCase: !!a.match_case, matchWholeWord: !!a.whole_word });
            res.load("items/text");
            await ctx.sync();
            const live = await liveHits(ctx, res.items.slice(0, 400), loaded.hasChanges);
            if (!live.length) return { ok: false, found: 0, replaced: 0, message: `Không tìm thấy "${truncate(find, 80)}"${where}` };
            let targets: Word.Range[];
            const occ = Math.floor(Number(a.occurrence));
            if (Number.isFinite(occ) && occ >= 1) {
              if (occ > live.length) throw new Error(`Chỉ có ${live.length} chỗ khớp${where}, không có chỗ thứ ${occ}.`);
              targets = [live[occ - 1]];
            } else if (live.length === 1) targets = live;
            else if (a.all === true) targets = live.slice(0, replaceLimit(a.max_replacements));
            else {
              // Nhiều chỗ khớp mà chưa nói rõ: không thay, trả danh sách để AI chọn
              const ps = live.slice(0, 15).map((h) => {
                const p = h.paragraphs.getFirst();
                p.load(supports("1.6") ? "text,uniqueLocalId" : "text");
                return p;
              });
              await ctx.sync();
              const ids = indexOfIds(loaded.rows);
              const texts = loaded.rows.map((r) => r.markup);
              const count = new Map<number, number>();
              const places = ps.map((p) => {
                const id = supports("1.6") ? p.uniqueLocalId : undefined;
                const i = id !== undefined && ids.has(id) ? (ids.get(id) as number) : texts.indexOf(p.text);
                const n = (count.get(i) || 0) + 1;
                count.set(i, n);
                return { paragraph_index: i, occurrence: n, text: i >= 0 ? truncate(loaded.rows[i].text, 110) : "" };
              });
              return {
                ok: false,
                found: live.length,
                replaced: 0,
                message: `Có ${live.length} chỗ khớp "${truncate(find, 60)}"${where}: CHƯA thay gì. Gọi lại với paragraph_index + occurrence (số thứ tự trong đoạn đó) để thay đúng một chỗ, hoặc all=true để thay tất cả.`,
                places,
              };
            }
            if (supports("1.6")) {
              const tps = targets.map((t) => {
                const p = t.paragraphs.getFirst();
                p.load("uniqueLocalId");
                return p;
              });
              await ctx.sync();
              const byId = new Map(loaded.rows.map((r) => [r.id as string, r.text] as [string, string]));
              new Set(tps.map((p) => p.uniqueLocalId)).forEach((id) => {
                u.touched.push(id);
                if (byId.has(id)) u.inverse.push(restoreText(id, byId.get(id) as string));
              });
            }
            // Cụm cũ và cụm mới có phần chung: chỉ đổi đúng các từ khác đi (dấu sửa đổi nhỏ nhất, giữ định dạng); khác hẳn thì thay cả cụm
            const same = targets.filter((t) => t.text === replace).length;
            if (same === targets.length) return { ok: true, found: live.length, replaced: 0, message: `Chữ ở đó đã đúng là "${truncate(replace, 60)}": không cần thay${where}` };
            for (const t of targets) {
              if (t.text === replace) continue;
              const plan = replace ? planParagraphEdit(t.text, t.text, replace) : "xóa";
              if (typeof plan !== "string" && plan.hunks >= 1 && plan.hunks <= 3 && plan.ratio < 0.8) await applyPlan(ctx, t, plan);
              else if (replace) t.insertText(replace, "Replace");
              else t.delete();
            }
            await ctx.sync();
            return { ok: true, found: live.length, replaced: targets.length - same, message: `Đã thay ${targets.length - same} chỗ${where}` };
          });
        });

      case "word_set_paragraph_text":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const raw = String(a.text ?? "").replace(/[\r\n]+$/, "");
            if (/[\r\n]/.test(raw)) {
              throw new Error("text có xuống dòng: mỗi lần chỉ viết lại MỘT đoạn. Thêm đoạn mới bằng word_insert_paragraphs hoặc word_insert_clause.");
            }
            const r = await resolvePara(ctx, a.index, a.expect);
            const note = movedNote(a.index, r);
            if (r.row.deleted) throw new Error(`Đoạn ${r.index} đã bị xóa (sửa đổi đang chờ duyệt): không sửa được. Muốn dùng lại phải từ chối sửa đổi xóa đó.`);
            const M = r.row.markup;
            const C = r.row.text;
            if (C === raw) return { ok: true, index: r.index, changes: 0, message: `Nội dung đoạn ${r.index} không đổi${note}` };
            // Đoạn có mốc chú thích: chỉ sửa mức từ (thay cả đoạn sẽ làm mất chú thích)
            const hasNotes = M.includes(NOTE_MARK);
            // Đoạn đang có sửa đổi chờ duyệt? (và các khúc đã xóa nằm ở đâu trong chuỗi tìm kiếm)
            let pending = noMarks(M) !== C;
            let spans: [number, number][] | undefined;
            if (r.hasChanges && supports("1.6")) {
              const info = await pendingInfo(ctx, r.p);
              pending = pending || info.count > 0;
              spans = info.spans;
            }
            const plan = planParagraphEdit(M, C, raw, withNoteSpans(M, spans));
            if (r.row.id) {
              u.touched.push(r.row.id);
              u.inverse.push(restoreText(r.row.id, C));
            }
            let mode: string;
            let changes: number;
            if (typeof plan === "string") {
              if (pending) {
                throw new Error(
                  `${plan} Đoạn ${r.index} đang có sửa đổi chờ duyệt nên không thay cả đoạn (sẽ làm hỏng dấu vết sửa đổi). Hãy sửa từng cụm bằng word_replace_text với paragraph_index=${r.index}.`
                );
              }
              if (hasNotes) {
                throw new Error(`${plan} Đoạn ${r.index} có chú thích cuối trang nên không thay cả đoạn (sẽ làm mất chú thích). Hãy sửa từng cụm bằng word_replace_text với paragraph_index=${r.index}.`);
              }
              r.p.insertText(raw, "Replace");
              await ctx.sync();
              mode = "thay cả đoạn";
              changes = 1;
            } else if (!pending && !hasNotes && plan.ratio > WHOLE_PARAGRAPH_RATIO && plan.hunks >= WHOLE_PARAGRAPH_HUNKS) {
              r.p.insertText(raw, "Replace");
              await ctx.sync();
              mode = "thay cả đoạn vì gần như viết lại toàn bộ";
              changes = 1;
            } else {
              await applyPlan(ctx, r.p, plan);
              mode = `${plan.hunks} chỗ, sửa đúng phần khác`;
              changes = plan.hunks;
            }
            // Đọc lại để chắc câu sau khi sửa đúng là câu được yêu cầu
            let now = raw;
            if (supports("1.4")) {
              const cur = supports("1.7") ? r.p.getText() : r.p.getReviewedText("Current");
              await ctx.sync();
              now = noMarks(curOf(cur.value, raw));
            }
            if (now !== raw) {
              return {
                ok: false,
                index: r.index,
                changes,
                message: `Đã ghi sửa đổi vào đoạn ${r.index} nhưng câu hiện tại KHÁC câu yêu cầu. Đọc lại đoạn rồi sửa tiếp.`,
                current_text: truncate(now, 1500),
              };
            }
            return { ok: true, index: r.index, changes, message: `Đã sửa đoạn ${r.index} (${mode})${note}` };
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
            let note = "";
            /** Đoạn neo là chữ thường (không phải mục danh sách, không phải tiêu đề): đoạn mới giữ kiểu của nó */
            let keepAnchorStyle = false;
            let anchorStyle = "";
            if (pos === "after_paragraph" || pos === "before_paragraph") {
              const r = await resolvePara(ctx, a.paragraph_index, a.expect);
              anchor = r.p;
              note = movedNote(a.paragraph_index, r);
              loc = pos === "before_paragraph" ? "Before" : "After";
              if (supports("1.3")) {
                anchor.load("isListItem,outlineLevel,style");
                await ctx.sync();
                keepAnchorStyle = !anchor.isListItem && !(Number(anchor.outlineLevel) >= 1 && Number(anchor.outlineLevel) <= 9);
                anchorStyle = anchor.style;
              }
            } else if (pos === "after_selection") {
              anchor = ctx.document.getSelection().paragraphs.getLast();
            }
            const numbered = isMarkdownNumberedList(lines);
            let prev: Word.Paragraph | null = null;
            let prevKind: string | null = null;
            let first: Word.Paragraph | null = null;
            const news: Word.Paragraph[] = [];
            const newTexts: string[] = [];
            for (const raw of lines) {
              const parsed = a.style ? { text: raw, style: mapStyle(a.style) } : parseMarkdownLine(raw, numbered);
              let np: Word.Paragraph;
              if (prev) np = prev.insertParagraph(parsed.text, "After");
              else if (anchor) np = anchor.insertParagraph(parsed.text, loc);
              else np = ctx.document.body.insertParagraph(parsed.text, pos === "start" ? "Start" : "End");
              if (!first) first = np;
              news.push(np);
              newTexts.push(parsed.text);
              const kind = parsed.style === "ListBullet" || parsed.style === "ListNumber" ? parsed.style : "other";
              if (kind === "other" || prevKind !== kind) {
                // Đoạn mới thừa hưởng kiểu và danh sách của đoạn đứng trước: đặt kiểu trước, rồi gỡ khỏi danh sách nếu còn dính.
                // (Thử 15/11/2026: đoạn thừa hưởng kiểu tiêu đề có gắn đánh số thì startNewList lỗi GeneralException.)
                const plain = kind !== "other" || parsed.style === "Normal";
                if (kind === "other" && plain && keepAnchorStyle && !a.style) np.style = anchorStyle;
                else np.styleBuiltIn = (kind === "other" ? parsed.style : "Normal") as any;
                if (plain) {
                  np.load("isListItem");
                  await ctx.sync();
                  if (np.isListItem) {
                    np.detachFromList();
                    await ctx.sync();
                  }
                }
                if (kind !== "other") {
                  const list = np.startNewList();
                  if (kind === "ListBullet") list.setLevelBullet(0, "Solid" as any);
                  else list.setLevelNumbering(0, "Arabic" as any, [0, "."]);
                }
              }
              prev = np;
              prevKind = kind;
            }
            let firstIndex: number | undefined;
            let total: number | undefined;
            if (first && supports("1.6")) {
              news.forEach((p) => p.load("uniqueLocalId"));
              const ps = ctx.document.body.paragraphs;
              ps.load("items/uniqueLocalId");
              await ctx.sync();
              u.touched.push(...news.map((p) => p.uniqueLocalId));
              u.inverse.push(removeParas(news.map((p) => p.uniqueLocalId), newTexts));
              total = ps.items.length;
              const k = ps.items.findIndex((p) => p.uniqueLocalId === (first as Word.Paragraph).uniqueLocalId);
              if (k >= 0) firstIndex = k;
            } else await ctx.sync();
            return {
              ok: true,
              ...(firstIndex !== undefined ? { first_index: firstIndex } : {}),
              count: lines.length,
              ...(total !== undefined ? { total_paragraphs: total } : {}),
              message: `Đã chèn ${lines.length} đoạn (${pos})${firstIndex !== undefined ? ` tại [¶${firstIndex}]` : ""}${note}. Chỉ số các đoạn phía sau đã dời.`,
            };
          })
        );

      case "word_insert_clause":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            if (!supports("1.6")) throw new Error("Word trên máy này chưa hỗ trợ chèn điều khoản kế thừa đánh số qua add-in (cần WordApi 1.6).");
            const lines = String(a.text ?? "")
              .split(/\r?\n/)
              .map((l) => l.trim())
              .filter(Boolean);
            if (!lines.length) throw new Error("Không có nội dung để chèn.");
            const loaded = await loadRows(ctx, true);
            const loc = locateParagraph(
              loaded.rows.map((r) => r.text),
              a.after_index,
              a.expect
            );
            if (loc.ok === false) throw new Error(loc.reason);
            if (!String(a.expect ?? "").trim()) throw new Error(missingExpect(a.after_index));
            const outline = buildOutline(loaded.rows);
            const item = outline.find((o) => o.i === loc.index) || null;
            const anchorRow = loaded.rows[loc.index];
            const anchor = loaded.paras.items[loc.index];
            const deeper = a.level === "deeper";
            // Cùng cấp / lên cấp: chèn sau toàn bộ điều khoản (gồm khoản con); xuống cấp: chèn ngay sau đoạn mở đầu
            let lastIdx = loc.index;
            if (item && !deeper) {
              // Có điều khoản cùng cấp hoặc cao hơn đứng sau: phạm vi của điều neo kết thúc ngay trước nó. Không có (điều cuối của văn bản
              // hay của phụ lục): phạm vi kéo tới hết văn bản, gồm cả phần chữ ký — khi đó chỉ tính tới khoản con cuối cùng.
              const bounded = outline.some((o) => o.i > item.i && o.level <= item.level);
              if (bounded) lastIdx = item.end;
              else outline.forEach((o) => o.i > item.i && o.i <= item.end && (lastIdx = Math.max(lastIdx, o.i)));
            }
            while (lastIdx > loc.index && (loaded.rows[lastIdx].inTable || loaded.rows[lastIdx].deleted || !loaded.rows[lastIdx].markup.trim())) lastIdx--;
            const last = loaded.paras.items[lastIdx];
            const isList = !!anchorRow.listString && anchorRow.listId !== undefined;
            const base = anchorRow.listLevel || 0;
            const level = Math.max(0, base + (deeper ? 1 : a.level === "shallower" ? -1 : 0));
            let prev = last;
            const made: Word.Paragraph[] = [];
            for (const line of lines) {
              const np = prev.insertParagraph(line, "After");
              made.push(np);
              prev = np;
            }
            if (lastIdx !== loc.index) {
              // Đoạn cuối của điều khoản có thể là khoản con hoặc chữ thường: đưa đoạn mới về kiểu của điều khoản neo
              made.forEach((np) => (np.style = anchorRow.style as string));
              await ctx.sync();
            }
            made.forEach((np) => np.load("isListItem"));
            await ctx.sync();
            if (isList) {
              // Kiểu không tự kéo theo đánh số (danh sách đặt trên đoạn): gắn vào đúng danh sách của điều khoản neo
              const loose = made.filter((np) => !np.isListItem);
              if (loose.length) {
                loose.forEach((np) => np.attachToList(anchorRow.listId as number, level));
                await ctx.sync();
              }
            } else {
              made.filter((np) => np.isListItem).forEach((np) => np.detachFromList());
            }
            if (isList) {
              // Đoạn mới có thể thừa hưởng danh sách KHÁC của đoạn đứng trước (khoản con nằm ở danh sách riêng): đưa về đúng danh sách của điều neo
              made.forEach((np) => np.load("listOrNullObject/id"));
              await ctx.sync();
              const wrong = made.filter((np) => np.listOrNullObject.isNullObject || np.listOrNullObject.id !== anchorRow.listId);
              if (wrong.length) {
                wrong.forEach((np) => np.attachToList(anchorRow.listId as number, level));
                await ctx.sync();
              }
              const lis = made.map((np) => {
                const li = np.listItemOrNullObject;
                li.load("level");
                return li;
              });
              await ctx.sync();
              lis.forEach((li) => {
                if (!li.isNullObject && li.level !== level) li.level = level;
              });
            }
            made.forEach((np) => np.load("uniqueLocalId"));
            await ctx.sync();
            u.touched.push(...made.map((np) => np.uniqueLocalId));
            u.inverse.push(removeParas(made.map((np) => np.uniqueLocalId), lines));
            // Đọc lại số mới
            const after = await loadRows(ctx, true);
            const out2 = buildOutline(after.rows);
            const ids = indexOfIds(after.rows);
            const news = made.map((np) => {
              const i = ids.get(np.uniqueLocalId) as number;
              const o = out2.find((x) => x.i === i);
              return { i, num: o ? o.num : "", text: truncate(after.rows[i]?.text, 80) };
            });
            const lastNew = news[news.length - 1];
            const nextSib = out2.find((o) => o.i > lastNew.i && o.source === "list");
            const numbered = news.every((n) => n.num);
            return {
              ok: true,
              inserted: news,
              ...(nextSib ? { next_clause: { i: nextSib.i, num: nextSib.num, text: truncate(nextSib.title, 60) } } : {}),
              total_paragraphs: after.rows.length,
              message: numbered
                ? `Đã chèn ${news.length} điều khoản: ${news.map((n) => n.num).join(", ")}; các điều khoản phía sau tự đánh số lại${loc.moved ? ` (đoạn neo nay là [¶${loc.index}])` : ""}. Chỉ số các đoạn phía sau đã dời.`
                : `Đã chèn ${news.length} đoạn sau [¶${lastIdx}], cùng kiểu với điều khoản neo. Văn bản đánh số gõ tay ở chỗ này: số không tự nhảy, phải tự gõ số cho khoản mới và sửa số các khoản phía sau.`,
            };
          })
        );

      case "word_change_level":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            if (!supports("1.3")) throw new Error("Word trên máy này chưa hỗ trợ đổi cấp đánh số qua add-in.");
            const r = await resolvePara(ctx, a.index, a.expect);
            const li = r.p.listItemOrNullObject;
            li.load("level,listString");
            await ctx.sync();
            if (li.isNullObject) throw new Error(`Đoạn ${r.index} không thuộc danh sách đánh số tự động của Word (số gõ tay thì sửa bằng word_set_paragraph_text).`);
            const before = li.listString;
            const oldLevel = li.level;
            const level = oldLevel + (a.direction === "deeper" ? 1 : -1);
            if (level < 0) throw new Error("Đoạn này đang ở cấp cao nhất.");
            if (level > 8) throw new Error("Đoạn này đang ở cấp sâu nhất.");
            li.level = level;
            await ctx.sync();
            if (r.row.id) {
              const id = r.row.id;
              u.touched.push(id);
              // Đổi cấp trên đoạn thường sinh một sửa đổi được theo dõi (hoàn tác bằng từ chối); trên đoạn đang là phần chèn chờ duyệt
              // thì nhập vào sửa đổi đó, hoặc không sinh gì: khi ấy đặt lại cấp cũ
              const setBack: UndoStep = () =>
                Word.run(async (c2) => {
                  const back = c2.document.getParagraphByUniqueLocalId(id).listItemOrNullObject;
                  back.load("level");
                  await c2.sync();
                  if (!back.isNullObject && back.level !== oldLevel) back.level = oldLevel;
                  await c2.sync();
                });
              u.fallback.push(setBack);
              u.inverse.push(setBack);
            }
            const after = await loadRows(ctx, true);
            const o = buildOutline(after.rows).find((x) => x.i === r.index);
            return { ok: true, index: r.index, num: o?.num || "", message: `Đã đổi cấp đoạn ${r.index}: ${before} → ${o?.num || "(không số)"}${movedNote(a.index, r)}` };
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
            // Tham số số học: chỉ nhận số thật (hoặc chuỗi chữ số) trong khoảng hợp lệ; chuỗi rỗng, false, "6pt", -1 đều bị bỏ và được báo lại
            const skipped: string[] = [];
            const numIn = (key: string, lo: number, hi: number): number | null => {
              const v = a[key];
              if (v === undefined || v === null) return null;
              const n = typeof v === "number" ? v : typeof v === "string" && /^\d+(\.\d+)?$/.test(v.trim()) ? Number(v) : NaN;
              if (!(n >= lo && n <= hi)) {
                skipped.push(key);
                return null;
              }
              return n;
            };
            const indent = numIn("first_line_indent_mm", 0, 100);
            const spaceAfter = numIn("space_after_pt", 0, 200);
            const lineSp = numIn("line_spacing", 0.5, 5);
            for (const p of paras) {
              if (a.alignment) p.alignment = a.alignment;
              if (indent !== null) p.firstLineIndent = (indent * 72) / 25.4;
              if (spaceAfter !== null) p.spaceAfter = spaceAfter;
              if (lineSp !== null) p.lineSpacing = 12 * lineSp;
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
            return { ok: true, message: `Đã định dạng ${ranges.length} vùng${skipped.length ? `. Bỏ qua tham số không hợp lệ: ${skipped.join(", ")} (phải là số trong khoảng cho phép)` : ""}` };
          })
        );

      case "word_add_comment":
        return Word.run(async (ctx) => {
          if (!supports("1.4")) throw new Error("Word trên máy này chưa hỗ trợ thêm nhận xét qua add-in.");
          const text = String(a.text || "").trim();
          if (!text) throw new Error("Nhận xét rỗng: cần nội dung trong text.");
          const { ranges } = await targetRanges(ctx, a);
          const c = ranges[0].insertComment(text);
          c.load("id");
          try {
            await ctx.sync();
          } catch (e: any) {
            const hint = officeErrorHint("word_add_comment", e);
            if (hint) throw new Error(hint);
            throw e;
          }
          const id = c.id;
          extra.push(() =>
            Word.run(async (c2) => {
              const cs = c2.document.body.getComments();
              cs.load("items/id");
              await c2.sync();
              cs.items.find((x) => x.id === id)?.delete();
              await c2.sync();
            })
          );
          return { ok: true, id, message: "Đã thêm nhận xét" };
        });

      case "word_insert_table":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const vals = normalize2D(a.values).map((r) => r.map((v) => String(v)));
            if (!vals.length) throw new Error("values rỗng.");
            let anchor: Word.Paragraph;
            let note = "";
            if (a.position === "after_paragraph") {
              const r = await resolvePara(ctx, a.paragraph_index, a.expect);
              anchor = r.p;
              note = movedNote(a.paragraph_index, r);
            } else if (a.position === "after_selection") anchor = ctx.document.getSelection().paragraphs.getLast();
            else anchor = ctx.document.body.paragraphs.getLast();
            const t = anchor.insertTable(vals.length, vals[0].length, "After", vals);
            t.headerRowCount = 1;
            t.styleBuiltIn = "GridTable4_Accent1" as any;
            await ctx.sync();
            return { ok: true, message: `Đã chèn bảng ${vals.length}×${vals[0].length}${note}. Chỉ số các đoạn phía sau đã dời (mỗi ô bảng là một đoạn).` };
          })
        );

      case "word_delete_paragraphs":
        return Word.run(async (ctx) =>
          tracked(ctx, async () => {
            const r = await resolvePara(ctx, a.start, a.expect);
            const span = a.end === undefined || a.end === null ? 0 : Number(a.end) - Number(a.start);
            if (!(span >= 0)) throw new Error("end phải lớn hơn hoặc bằng start.");
            const s = r.index;
            const e = s + span;
            if (e >= r.rows.length) throw new Error(`Chỉ số đoạn không hợp lệ (có ${r.rows.length} đoạn).`);
            for (let i = s; i <= e; i++) if (r.rows[i].id) u.touched.push(r.rows[i].id as string);
            for (let i = e; i >= s; i--) r.paras.items[i].delete();
            await ctx.sync();
            return {
              ok: true,
              deleted: e - s + 1,
              message: `Đã xóa đoạn ${s}${e > s ? `–${e}` : ""} (ghi có theo dõi)${movedNote(a.start, r)}. Đoạn bị xóa vẫn giữ chỗ trong chỉ số cho tới khi sửa đổi được chấp nhận.`,
            };
          })
        );

      case "word_list_comments":
        return Word.run(async (ctx) => {
          if (!supports("1.4")) throw new Error("Word trên máy này chưa hỗ trợ đọc nhận xét qua add-in.");
          const withIds = supports("1.6");
          const { rows } = await loadRows(ctx, true);
          const outline = buildOutline(rows);
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
          const texts = rows.map((r) => r.markup);
          // Số đoạn theo định danh đoạn (đúng cả khi có đoạn trùng nội dung); Word cũ thì so chữ
          const ids = withIds ? rows.map((r) => r.id as string) : null;
          const list = items.map((c, i) => {
            const para = paragraphIndex(withIds ? anchors[i].p.uniqueLocalId : undefined, anchors[i].p.text, ids, texts);
            return {
              id: c.id,
              author: c.authorName,
              text: truncate(c.content, 2000),
              anchor: truncate(anchors[i].r.text, 300),
              paragraph: para,
              clause: para >= 0 ? clauseName(clauseAt(outline, para)) : "",
              resolved: c.resolved,
              replies: anchors[i].reps.items.map((x) => ({ author: x.authorName, text: truncate(x.content, 1000) })),
            };
          });
          const { picked, truncated } = takeWithinBudget(list, RESULT_JSON_BUDGET - 400, (c, room) => {
            // Một nhận xét quá dài (nhiều câu trả lời): bớt câu trả lời cho vừa, để lần gọi sau vẫn đi tiếp được
            let reps = c.replies;
            while (reps.length > 1 && JSON.stringify(reps).length > room - 3000) reps = reps.slice(0, Math.ceil(reps.length / 2));
            return { ...c, replies: reps.map((x) => ({ ...x, text: truncate(x.text, 300) })), replies_total: c.replies.length };
          });
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
          cs.load("items/id,items/resolved");
          await ctx.sync();
          const id = String(a.id);
          const c = cs.items.find((x) => x.id === id);
          if (!c) throw new Error(`Không có nhận xét id ${a.id}.`);
          const find = async (c2: Word.RequestContext) => {
            const all = c2.document.body.getComments();
            all.load("items/id");
            await c2.sync();
            return all.items.find((x) => x.id === id);
          };
          if (name === "word_reply_comment") {
            const text = String(a.text || "").trim();
            if (!text) throw new Error("Câu trả lời rỗng: cần nội dung trong text.");
            const rep = c.reply(text);
            rep.load("id");
            await ctx.sync();
            const repId = rep.id;
            extra.push(() =>
              Word.run(async (c2) => {
                const cm = await find(c2);
                if (!cm) return;
                const reps = cm.replies;
                reps.load("items/id");
                await c2.sync();
                reps.items.find((x) => x.id === repId)?.delete();
                await c2.sync();
              })
            );
            return { ok: true, message: "Đã trả lời nhận xét" };
          }
          const prev = c.resolved;
          c.resolved = a.resolved !== false;
          await ctx.sync();
          extra.push(() =>
            Word.run(async (c2) => {
              const cm = await find(c2);
              if (!cm) return;
              cm.resolved = prev;
              await c2.sync();
            })
          );
          return { ok: true, message: a.resolved === false ? "Đã mở lại nhận xét" : "Đã đánh dấu đã xử lý" };
        });

      case "word_list_revisions":
        return Word.run(async (ctx) => {
          if (!supports("1.6")) throw new Error("Word trên máy này chưa hỗ trợ đọc sửa đổi được theo dõi qua add-in.");
          const { rows, paras } = await loadRows(ctx, true);
          const outline = buildOutline(rows);
          const tc = ctx.document.body.getTrackedChanges();
          tc.load("items/type,items/author,items/date,items/text");
          await ctx.sync();
          const authors: Record<string, number> = {};
          tc.items.forEach((t) => (authors[t.author] = (authors[t.author] || 0) + 1));
          const want = String(a.author || "")
            .trim()
            .toLowerCase();
          const match = (t: Word.TrackedChange) => !want || String(t.author).toLowerCase().includes(want);
          const from = Math.max(0, Number(a.start) || 0);
          // Cửa sổ: tối đa 150 sửa đổi khớp bộ lọc, tính từ `from`
          const win: number[] = [];
          for (let k = from; k < tc.items.length && win.length < 150; k++) if (match(tc.items[k])) win.push(k);
          const ps = win.map((k) => {
            const p = tc.items[k].getRange().paragraphs.getFirst();
            p.load("uniqueLocalId");
            return p;
          });
          await ctx.sync();
          const ids = indexOfIds(rows);
          const paraOf = win.map((_, j) => (ids.has(ps[j].uniqueLocalId) ? (ids.get(ps[j].uniqueLocalId) as number) : -1));
          // Câu trước khi sửa của từng đoạn có sửa đổi
          const uniq = Array.from(new Set(paraOf.filter((i) => i >= 0)));
          const orig = new Map<number, OfficeExtension.ClientResult<string>>();
          uniq.forEach((i) => orig.set(i, paras.items[i].getReviewedText("Original")));
          await ctx.sync();
          type Group = { para: number; clause: string; before: string; after: string; changes: any[] };
          const groups: Group[] = [];
          win.forEach((k, j) => {
            const t = tc.items[k];
            const para = paraOf[j];
            let g = groups[groups.length - 1];
            if (!g || g.para !== para) {
              g = {
                para,
                clause: para >= 0 ? clauseName(clauseAt(outline, para)) : "",
                before: para >= 0 ? truncate(strip(orig.get(para)?.value), 260) : "",
                after: para >= 0 ? truncate(rows[para].text, 260) : "",
                changes: [],
              };
              groups.push(g);
            }
            g.changes.push({
              index: k,
              type: t.type,
              author: t.author,
              date: t.date ? new Date(t.date as any).toLocaleString("vi-VN") : "",
              text: truncate(strip(t.text), 300),
            });
          });
          const { picked, truncated } = takeWithinBudget(groups, RESULT_JSON_BUDGET - 800, (g, room) => {
            let ch = g.changes;
            while (ch.length > 1 && JSON.stringify(ch).length > room - 700) ch = ch.slice(0, Math.ceil(ch.length / 2));
            return { ...g, changes: ch };
          });
          const shown = picked.reduce((n, g) => n + g.changes.length, 0);
          const lastShown = shown ? picked[picked.length - 1].changes[picked[picked.length - 1].changes.length - 1].index : from - 1;
          let more = false;
          for (let k = lastShown + 1; k < tc.items.length && !more; k++) more = match(tc.items[k]);
          return {
            ok: true,
            total: tc.items.length,
            authors,
            ...(want ? { filtered_by_author: a.author } : {}),
            shown_from: from,
            shown,
            ...(more ? { next_start: lastShown + 1 } : {}),
            ...(truncated || more ? { truncated: true, note: "Chưa hết: gọi lại với start = next_start." } : {}),
            groups: picked,
          };
        });

      case "word_review_revision":
        return Word.run(async (ctx) => {
          if (!supports("1.6")) throw new Error("Word trên máy này chưa hỗ trợ duyệt sửa đổi qua add-in.");
          const tc = ctx.document.body.getTrackedChanges();
          tc.load("items/author,items/text,items/type");
          await ctx.sync();
          const accept = a.action === "accept";
          const verb = accept ? "chấp nhận" : "từ chối";
          const want = String(a.author || "")
            .trim()
            .toLowerCase();
          if (a.all) {
            if (want) {
              const who = resolveAuthor(
                tc.items.map((t) => String(t.author)),
                want
              );
              if (typeof who !== "string") {
                throw new Error(
                  who.length
                    ? `"${a.author}" khớp nhiều tác giả: ${who.map((x) => `“${x}”`).join(", ")}. Gọi lại với author là đúng một tên trong số đó.`
                    : `Không có sửa đổi nào của "${a.author}". Tác giả trong văn bản: ${Array.from(new Set(tc.items.map((t) => `“${t.author}”`))).join(", ") || "(không có sửa đổi)"}.`
                );
              }
              const mine = tc.items.filter((t) => String(t.author) === who);
              for (const t of mine.slice().reverse()) {
                if (accept) t.accept();
                else t.reject();
              }
              await ctx.sync();
              return { ok: true, reviewed: mine.length, author: who, message: `Đã ${verb} ${mine.length} sửa đổi của “${who}”. Thao tác này add-in không hoàn tác được.` };
            }
            const n = tc.items.length;
            if (accept) tc.acceptAll();
            else tc.rejectAll();
            await ctx.sync();
            return { ok: true, reviewed: n, message: `Đã ${verb} tất cả ${n} sửa đổi. Thao tác này add-in không hoàn tác được.` };
          }
          const raw: number[] = (a.indexes || []).map((x: any) => Number(x));
          const texts: string[] = Array.isArray(a.texts) ? a.texts.map((x: any) => String(x ?? "")) : [];
          if (!raw.length) throw new Error("Cần indexes (từ word_list_revisions) hoặc all=true.");
          if (texts.length !== raw.length || texts.some((x) => !x.trim())) {
            throw new Error("Thiếu texts: với mỗi số trong indexes phải kèm phần đầu chữ của sửa đổi đó (đúng như word_list_revisions đã liệt kê) để công cụ đối chiếu trước khi duyệt. Duyệt là thao tác không hoàn tác được.");
          }
          const squash = (s: string) => strip(s).replace(/\s+/g, " ").replace(/…$/, "").trim();
          raw.forEach((i, k) => {
            if (!(Number.isInteger(i) && i >= 0 && i < tc.items.length)) throw new Error(`Không có sửa đổi số ${i} (có ${tc.items.length}). Đọc lại word_list_revisions.`);
            const exp = squash(texts[k] || "");
            if (exp && !squash(tc.items[i].text).startsWith(exp.slice(0, 60))) {
              throw new Error(
                `Sửa đổi số ${i} hiện là ${tc.items[i].type} “${truncate(strip(tc.items[i].text), 60)}”, không phải “${truncate(exp, 60)}”: danh sách đã đổi (mỗi lần AI sửa lại sinh thêm sửa đổi). Đọc lại word_list_revisions rồi duyệt.`
              );
            }
          });
          const idx = Array.from(new Set(raw)).sort((x, y) => y - x);
          for (const i of idx) {
            if (accept) tc.items[i].accept();
            else tc.items[i].reject();
          }
          await ctx.sync();
          return { ok: true, reviewed: idx.length, message: `Đã ${verb} ${idx.length} sửa đổi. Số thứ tự các sửa đổi còn lại đã đổi: đọc lại danh sách trước khi duyệt tiếp.` };
        });

      case "word_read_tables":
        return Word.run(async (ctx) => {
          const tables = ctx.document.body.tables;
          tables.load("items/rowCount,items/values");
          await ctx.sync();
          const maxRows = Number(a.max_rows) || 40;
          const from = Math.max(0, Number(a.start) || 0);
          const page = tables.items.slice(from, from + 30);
          // Ô bảng đọc theo bản hiện tại (không lẫn chữ đã xóa có theo dõi)
          const vals: string[][][] = [];
          for (const t of page) vals.push(await currentTableValues(ctx, t, t.values, maxRows));
          const list = page.map((t, k) => ({
            index: from + k,
            rows: t.rowCount,
            columns: t.values[0]?.length || 0,
            ...(t.rowCount > maxRows ? { rows_shown: maxRows } : {}),
            values: vals[k].slice(0, maxRows).map((r) => r.map((v) => truncate(v, 200))),
          }));
          const { picked, truncated } = takeWithinBudget(list, RESULT_JSON_BUDGET - 400, (x, room) => {
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
            const plan = planTableFill(await currentTableValues(ctx, t, t.values), a.values, a.start_row, a.start_col);
            plan.writes.forEach((x) => (t.getCell(x.row, x.col).value = x.value));
            const written = plan.writes.length;
            const more = plan.extraRows;
            if (more.length) t.addRows("End", more.length, more);
            await ctx.sync();
            return { ok: true, message: `Đã điền ${written} ô${more.length ? `, thêm ${more.length} hàng` : ""} vào bảng ${a.table_index}` };
          })
        );

      case "word_go_to":
        return Word.run(async (ctx) => {
          const r = await resolvePara(ctx, a.paragraph_index, undefined, false);
          r.p.select();
          await ctx.sync();
          return { ok: true, message: `Đã chuyển tới đoạn ${a.paragraph_index}` };
        });
    }
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
  function execute(name: string, a: any, ctx: { log: UndoEntry[] }): Promise<any> {
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
      return await run(name, a, u);
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
        const { rows } = await loadRows(ctx);
        const loc = locateParagraph(
          rows.map((r) => r.text),
          a.start,
          a.expect
        );
        if (loc.ok === false) return null; // công cụ sẽ tự từ chối
        const span = a.end === undefined || a.end === null ? 0 : Math.max(0, Number(a.end) - Number(a.start));
        const s = loc.index;
        const e = Math.min(rows.length - 1, s + span);
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
