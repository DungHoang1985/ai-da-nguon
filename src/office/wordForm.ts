/** Đợt 8 — điền mẫu trong Word: liệt kê mọi chỗ cần điền và điền hết trong một lệnh (có Track Changes, một mục Hoàn tác).
 *  Ba loại chỗ điền:
 *   - ô điền mẫu (content control), mã c<id>: ô chữ, ô chữ định dạng, ô đánh dấu, ô chọn, ô chọn hoặc nhập, ô ngày;
 *   - chỗ trống gõ tay trong đoạn ("……", "____", "[●]", "ngày   tháng   năm"), mã b<đoạn>.<thứ tự trong đoạn>;
 *   - ô bảng trống có ô nhãn bên trái, mã t<bảng>.<hàng>.<cột>.
 *  Đo 18/11/2026 (`tasks/dot-8.md`): ghi chữ vào ô điền mẫu rồi Hoàn tác (từ chối sửa đổi) làm Word khóa ô với mọi lệnh ghi sau →
 *  MỌI loại ô được ghi bằng cách dựng lại OOXML của cả ô và insertOoxml (Hoàn tác trả nguyên ô cũ). Chỗ trống tìm đúng trọn chuỗi
 *  và đúng lần xuất hiện, bỏ các kết quả tìm nằm trong chữ đã xóa chờ duyệt. Soát mã độc lập: `tasks/dot-8-bang-chung/soat-ma-dot-8.md`. */
import { ToolSpec } from "../core/agent/types";
import { ParaRow } from "../core/outline";
import {
  checkboxValue,
  controlRefusal,
  countOccurrences,
  dateFormatFromOoxml,
  fillControlOoxml,
  findBlanks,
  formatDate,
  listItemsFromOoxml,
  matchListItem,
  parseDate,
} from "../core/formfill";
import { checkConsistency } from "../core/consistency";
import { buildOutline } from "../core/outline";
import { RESULT_JSON_BUDGET, takeWithinBudget, truncate } from "./util";

const CC_TYPES = ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"];
/** Word chỉ tìm chuỗi tới 255 ký tự */
const MAX_SEARCH = 255;
/** Ô chọn dài quá thì danh sách mục chỉ đưa bấy nhiêu mục đầu */
const MAX_ITEMS = 40;

export const WORD_FORM_SPECS: ToolSpec[] = [
  {
    name: "word_list_form_fields",
    description:
      "List every spot to fill in a form/template, each with an id and the label written before it: content controls (id c<n>; type text/checkbox/dropdown with its allowed items/combobox/date with its format; locked ones are marked and cannot be filled), typed blanks inside paragraphs such as '……', '____', '[●]', 'ngày   tháng   năm' (id b<paragraph>.<k>; 'blank' is the exact blank text that will be replaced), and empty table cells with a label cell on their left (id t<table>.<row>.<col>). Call this first, then fill everything with ONE word_fill_form call.",
    parameters: {
      type: "object",
      properties: {
        start: { type: "integer", description: "Continue from this position when the previous result had next_start." },
      },
    },
  },
  {
    name: "word_fill_form",
    description:
      "Fill many spots listed by word_list_form_fields in ONE call (tracked changes; one Undo for the whole call). fields: [{id, value, label}] — label is REQUIRED and must be the label you got from the list for that id (the tool refuses a spot whose label does not match, so values never land in the wrong place). Checkbox: value 'true' or 'false'. Dropdown: exactly one of its items. Date: dd/MM/yyyy (written in the control's own format). Typed blank: value replaces the whole blank text, e.g. blank 'ngày   tháng   năm' → 'ngày 18 tháng 11 năm'. Ids change after filling: call word_list_form_fields again before another fill. Only fill spots you have real information for; never invent ID numbers, dates or amounts — leave them out and tell the user which spots are still empty.",
    parameters: {
      type: "object",
      properties: {
        fields: {
          type: "array",
          items: {
            type: "object",
            properties: {
              id: { type: "string" },
              value: { type: "string" },
              label: { type: "string", description: "The label from word_list_form_fields for this id." },
            },
            required: ["id", "value", "label"],
          },
        },
      },
      required: ["fields"],
    },
    write: true,
  },
];

interface Loaded {
  paras: Word.ParagraphCollection;
  rows: (ParaRow & { markup?: string })[];
}
interface Deps {
  supports: (v: string, set?: string) => boolean;
  tracked: <T>(ctx: Word.RequestContext, fn: () => Promise<T>) => Promise<T>;
  loadRows: (ctx: Word.RequestContext) => Promise<Loaded>;
  /** Chữ các ô của bảng theo bản hiện tại (không gồm chữ đã xóa chờ duyệt) */
  tableValues: (ctx: Word.RequestContext, t: Word.Table, values: string[][]) => Promise<string[][]>;
}

/** Ô điền mẫu còn sống. Ghi đè cả ô (OOXML) để lại ô cũ "chờ xóa" cạnh ô mới, cùng tên; đọc cả hai thì ô cũ trông như ô còn trống
 *  (Chạy thật 18/11/2026). Chỉ bỏ ô mà MỌI sửa đổi bên trong là xóa VÀ còn một ô sống khác cùng tên, cùng tag — ô người dùng chỉ xóa
 *  bớt chữ bên trong vẫn được giữ (soát mã Đợt 8, #18). */
export async function liveContentControls(ctx: Word.RequestContext, supports: (v: string) => boolean, props: string): Promise<Word.ContentControl[]> {
  const cc = ctx.document.getContentControls({ types: CC_TYPES as any });
  cc.load(props.includes("items/tag") && props.includes("items/title") ? props : `${props},items/tag,items/title`);
  await ctx.sync();
  if (!supports("1.6") || !cc.items.length) return cc.items;
  const tcs = cc.items.map((c) => {
    const t = c.getRange("Whole").getTrackedChanges();
    t.load("items/type");
    return t;
  });
  try {
    await ctx.sync();
  } catch {
    return cc.items;
  }
  const gone = cc.items.map((_, k) => {
    const types = tcs[k].items.map((x) => String(x.type));
    return types.length > 0 && types.every((t) => t === "Deleted");
  });
  const key = (c: Word.ContentControl) => `${c.tag}\u0000${c.title}`;
  return cc.items.filter((c, k) => !(gone[k] && cc.items.some((o, j) => j !== k && !gone[j] && key(o) === key(c))));
}

export interface FormField {
  id: string;
  kind: "control" | "blank" | "cell";
  /** Loại ô điền mẫu (PlainText…) */
  type?: string;
  label: string;
  value?: string;
  empty: boolean;
  locked?: string;
  items?: string[];
  date_format?: string;
  blank?: string;
  after?: string;
  i?: number;
}

/** Đọc một mã chỗ điền */
export function parseFieldId(id: unknown): { kind: "control"; cc: number } | { kind: "blank"; i: number; k: number } | { kind: "cell"; t: number; r: number; c: number } | null {
  const s = String(id ?? "").trim();
  // Mã ô điền mẫu do Word cấp có thể âm (đo 18/11/2026: c-672413777)
  let m = /^c(-?\d+)$/i.exec(s);
  if (m) return { kind: "control", cc: Number(m[1]) };
  if ((m = /^b(\d+)\.(\d+)$/i.exec(s))) return { kind: "blank", i: Number(m[1]), k: Number(m[2]) };
  if ((m = /^t(\d+)\.(\d+)\.(\d+)$/i.exec(s))) return { kind: "cell", t: Number(m[1]), r: Number(m[2]), c: Number(m[3]) };
  return null;
}

const norm = (a: unknown) =>
  String(a ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
const same = (a: unknown, b: unknown) => norm(a) === norm(b);

/** Nhãn của ô bảng: ô có chữ gần nhất bên trái, kèm tên cột (hàng đầu) nếu có (soát mã Đợt 8, #20) */
export function cellLabel(v: string[][], r: number, c: number): string | null {
  const row = v[r] || [];
  let left = "";
  for (let k = c - 1; k >= 0 && !left; k--) left = String(row[k] ?? "").trim();
  if (!left) return null;
  const head = r > 0 ? String(v[0]?.[c] ?? "").trim() : "";
  return head && head !== left ? `${left} — ${head}` : left;
}

/** Ô bảng trống có nhãn bên trái (không lấy cột đầu — đó là cột nhãn, soát mã Đợt 8, #21) */
export function emptyCells(v: string[][]): { r: number; c: number; label: string }[] {
  const out: { r: number; c: number; label: string }[] = [];
  v.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (c === 0 || String(cell ?? "").trim()) return;
      const label = cellLabel(v, r, c);
      if (label) out.push({ r, c, label });
    })
  );
  return out;
}

export function createWordFormTools(d: Deps) {
  const names = new Set(WORD_FORM_SPECS.map((s) => s.name));

  /** Ô điền mẫu trong thân văn bản: nhãn (tên ô, không có thì chữ đứng trước ô trong đoạn — soát mã Đợt 8, #2), trạng thái ô đánh dấu,
   *  OOXML cho ô chọn / ô ngày (danh sách mục, định dạng ngày). */
  async function loadControls(ctx: Word.RequestContext) {
    if (!d.supports("1.5")) return [];
    const items = await liveContentControls(ctx, d.supports, "items/id,items/tag,items/title,items/type,items/text,items/cannotEdit,items/cannotDelete");
    const paras = items.map((c) => {
      const p = c.paragraphs.getFirstOrNullObject();
      p.load("text");
      return p;
    });
    const xmls = items.map((c) => (/DropDownList|ComboBox|DatePicker/.test(String(c.type)) ? c.getRange("Whole").getOoxml() : null));
    await ctx.sync();
    // Chữ đứng trước ô trong đoạn của nó
    let before: (string | null)[] = items.map(() => null);
    try {
      const rs = items.map((c, k) => {
        if (paras[k].isNullObject) return null;
        const r = paras[k].getRange("Start").expandTo(c.getRange("Start"));
        r.load("text");
        return r;
      });
      await ctx.sync();
      before = rs.map((r) => (r ? r.text : null));
    } catch {
      /* không đo được: dùng tên ô / tag */
    }
    // Trạng thái ô đánh dấu đọc từ Word (ký hiệu có thể là phông Wingdings — soát mã Đợt 8, #11)
    const checked: (boolean | null)[] = items.map(() => null);
    const boxes = items.map((c, k) => ({ c, k })).filter((x) => String(x.c.type) === "CheckBox");
    if (boxes.length) {
      try {
        boxes.forEach((x) => x.c.checkboxContentControl.load("isChecked"));
        await ctx.sync();
        boxes.forEach((x) => (checked[x.k] = !!x.c.checkboxContentControl.isChecked));
      } catch {
        boxes.forEach((x) => (checked[x.k] = /[☒☑✓✔]/.test(String(x.c.text ?? ""))));
      }
    }
    return items.map((c, k) => {
      const xml = xmls[k]?.value || "";
      const prev = String(before[k] ?? "")
        .replace(/\s+/g, " ")
        .trim();
      const label = String(c.title || "").trim() || truncate(prev, 80).trim() || String(c.tag || "").trim() || `ô ${c.type}`;
      return {
        c,
        label,
        para: paras[k].isNullObject ? "" : paras[k].text,
        checked: checked[k],
        items: xml ? listItemsFromOoxml(xml) : [],
        dateFormat: xml ? dateFormatFromOoxml(xml) : null,
        locked: c.cannotEdit ? "khóa nội dung" : c.cannotDelete ? "khóa không cho xóa ô (add-in phải thay cả ô nên không điền được)" : "",
      };
    });
  }

  /** Các kết quả tìm `str` trong đoạn, chỉ giữ chỗ là chữ đang có (Word tìm cả trong chữ đã xóa chờ duyệt — soát mã Đợt 8, #3) */
  async function liveHits(ctx: Word.RequestContext, para: Word.Paragraph, str: string): Promise<Word.Range[]> {
    const hits = para.search(str, { matchCase: true });
    hits.load("items");
    await ctx.sync();
    if (!hits.items.length || !d.supports("1.4")) return hits.items;
    const cur = hits.items.map((h) => h.getReviewedText("Current"));
    await ctx.sync();
    return hits.items.filter((_, k) => cur[k].value === str);
  }

  /** Vùng của chỗ trống thứ k trong đoạn i; lỗi (câu tiếng Việt) nếu không chắc đúng chỗ */
  async function locateBlank(ctx: Word.RequestContext, loaded: Loaded, i: number, k: number): Promise<{ range: Word.Range; found: ReturnType<typeof findBlanks>[number] }> {
    const row = loaded.rows[i];
    const found = row && !row.deleted ? findBlanks(row.text)[k] : undefined;
    if (!found) throw new Error("chỗ trống không còn ở đó (văn bản đã đổi): gọi lại word_list_form_fields");
    if (found.str.length > MAX_SEARCH) throw new Error("chỗ trống dài quá 255 ký tự, Word không tìm được: sửa bằng word_replace_text");
    const live = await liveHits(ctx, loaded.paras.items[i], found.str);
    if (live.length !== countOccurrences(row.text, found.str)) {
      throw new Error("không định vị chắc được chỗ trống này (đoạn có sửa đổi chờ duyệt đan vào chỗ trống): duyệt sửa đổi trước hoặc sửa bằng word_replace_text");
    }
    const range = live[found.occurrence];
    // Chỗ trống nằm trong chữ gợi ý của một ô điền mẫu: điền qua mã c… của ô (soát mã Đợt 8, #4)
    if (d.supports("1.3")) {
      const pc = range.parentContentControlOrNullObject;
      pc.load("id");
      await ctx.sync();
      if (!pc.isNullObject) throw new Error(`chỗ trống nằm trong ô điền mẫu c${pc.id}: điền bằng mã c${pc.id}`);
    }
    return { range, found };
  }

  async function listFields(): Promise<FormField[]> {
    return Word.run(async (ctx) => {
      const out: FormField[] = [];
      for (const x of await loadControls(ctx)) {
        const c = x.c;
        const type = String(c.type);
        const text = String(c.text ?? "");
        out.push({
          id: `c${c.id}`,
          kind: "control",
          type,
          label: x.label,
          value: type === "CheckBox" ? String(!!x.checked) : text,
          empty: type === "CheckBox" ? false : !text.trim(),
          ...(x.locked ? { locked: x.locked } : {}),
          ...(x.items.length ? { items: x.items.slice(0, MAX_ITEMS) } : {}),
          ...(type === "DatePicker" && x.dateFormat ? { date_format: x.dateFormat } : {}),
        });
      }
      const loaded = await d.loadRows(ctx);
      const withBlanks = loaded.rows.filter((r) => !r.deleted && findBlanks(r.text).length);
      // Đoạn có ô điền mẫu: bỏ các chỗ trống nằm trong ô (đã liệt kê theo mã c… của ô)
      const ccIn = withBlanks.map((r) => {
        const cs = loaded.paras.items[r.i].contentControls;
        cs.load("items");
        return cs;
      });
      if (withBlanks.length) await ctx.sync();
      for (let n = 0; n < withBlanks.length; n++) {
        const r = withBlanks[n];
        const bs = findBlanks(r.text);
        for (let k = 0; k < bs.length; k++) {
          if (ccIn[n].items.length) {
            try {
              await locateBlank(ctx, loaded, r.i, k);
            } catch (e: any) {
              if (/nằm trong ô điền mẫu/.test(String(e?.message))) continue;
            }
          }
          out.push({ id: `b${r.i}.${k}`, kind: "blank", label: bs[k].label, blank: bs[k].str, after: bs[k].after, empty: true, i: r.i });
        }
      }
      const tables = ctx.document.body.tables;
      tables.load("items/values");
      await ctx.sync();
      for (let ti = 0; ti < tables.items.length; ti++) {
        const t = tables.items[ti];
        const v = await d.tableValues(ctx, t, t.values || []);
        for (const x of emptyCells(v)) out.push({ id: `t${ti}.${x.r}.${x.c}`, kind: "cell", label: x.label, empty: true });
      }
      return out;
    });
  }

  async function list(a: any) {
    const all = await listFields();
    const from = Math.max(0, Math.floor(Number(a.start)) || 0);
    const { picked, truncated } = takeWithinBudget(all.slice(from), RESULT_JSON_BUDGET - 1500);
    const next = from + picked.length;
    const n = (k: FormField["kind"]) => all.filter((f) => f.kind === k).length;
    const empty = all.filter((f) => f.empty).length;
    return {
      ok: true,
      brief: all.length ? `${all.length} chỗ điền (${n("control")} ô điền mẫu, ${n("blank")} chỗ trống gõ tay, ${n("cell")} ô bảng), ${empty} chỗ còn trống` : "không thấy chỗ nào cần điền",
      total: all.length,
      fields: picked,
      ...(truncated || next < all.length ? { next_start: next } : {}),
      ...(all.length ? {} : { note: "Không thấy ô điền mẫu, chỗ trống (……, ____, [●]) hay ô bảng trống có ô nhãn. Nếu mẫu dùng cách khác, đọc văn bản rồi sửa bằng các công cụ sửa đoạn." }),
    };
  }

  /** Lỗi Office.js thành câu ngắn */
  const why = (e: any) => truncate(`${e?.message || e}${e?.code && !String(e?.message).includes(e.code) ? ` [${e.code}]` : ""}`, 200);

  async function fill(a: any) {
    const fields: { id: string; value: string; label?: string }[] = (Array.isArray(a.fields) ? a.fields : [])
      .filter((f: any) => f && typeof f === "object" && String(f.id ?? "").trim())
      .map((f: any) => ({ id: String(f.id).trim(), value: String(f.value ?? ""), label: f.label === undefined || f.label === null ? undefined : String(f.label) }));
    if (!fields.length) throw new Error("Thiếu fields: danh sách { id, value, label } lấy từ word_list_form_fields.");
    const filled: string[] = [];
    const skipped: { id: string; reason: string }[] = [];
    const seen = new Set<string>();
    const blanks: { f: (typeof fields)[number]; i: number; k: number }[] = [];
    const cells: { f: (typeof fields)[number]; t: number; r: number; c: number }[] = [];
    const ccs: { f: (typeof fields)[number]; cc: number }[] = [];
    for (const f of fields) {
      const key = f.id.toLowerCase();
      if (seen.has(key)) {
        skipped.push({ id: f.id, reason: "trùng mã trong cùng lệnh" });
        continue;
      }
      seen.add(key);
      const p = parseFieldId(f.id);
      if (!p) skipped.push({ id: f.id, reason: "mã không hợp lệ (dùng mã từ word_list_form_fields)" });
      else if (f.label === undefined || !f.label.trim()) skipped.push({ id: f.id, reason: "thiếu label: đưa đúng nhãn lấy từ word_list_form_fields" });
      else if (p.kind === "blank") blanks.push({ f, i: p.i, k: p.k });
      else if (p.kind === "cell") cells.push({ f, t: p.t, r: p.r, c: p.c });
      else ccs.push({ f, cc: p.cc });
    }
    await Word.run(async (ctx) =>
      d.tracked(ctx, async () => {
        // Thứ tự: chỗ trống → ô bảng → ô điền mẫu. Thay một ô chứa cả đoạn để lại đoạn cũ "chờ xóa" làm lệch số đoạn của các chỗ
        // trống phía sau (soát mã Đợt 8, #1); chỗ trống ghi trước thì số đoạn còn đúng như lúc liệt kê.
        if (blanks.length) {
          const loaded = await d.loadRows(ctx);
          // Trong một đoạn: điền từ chỗ trống cuối lên
          blanks.sort((x, y) => x.i - y.i || y.k - x.k);
          for (const b of blanks) {
            try {
              const { range, found } = await locateBlank(ctx, loaded, b.i, b.k);
              if (!same(b.f.label, found.label)) throw new Error(`nhãn không khớp (chỗ này đứng sau “${truncate(found.label, 60)}”): gọi lại word_list_form_fields`);
              range.insertText(b.f.value, "Replace");
              await ctx.sync();
              filled.push(b.f.id);
            } catch (e: any) {
              skipped.push({ id: b.f.id, reason: why(e) });
            }
          }
        }

        if (cells.length) {
          const tables = ctx.document.body.tables;
          tables.load("items/values");
          await ctx.sync();
          const cur = new Map<number, string[][]>();
          for (const x of cells) {
            try {
              const t = tables.items[x.t];
              if (!t) throw new Error("không còn bảng này (văn bản đã đổi): gọi lại word_list_form_fields");
              if (!cur.has(x.t)) cur.set(x.t, await d.tableValues(ctx, t, t.values || []));
              const v = cur.get(x.t)!;
              if (x.c === 0) throw new Error("cột đầu là cột nhãn, không điền");
              if (!v[x.r] || x.c >= v[x.r].length) throw new Error("không còn ô bảng này (văn bản đã đổi): gọi lại word_list_form_fields");
              const label = cellLabel(v, x.r, x.c);
              if (!label || !(same(x.f.label, label) || same(String(x.f.label).split(" — ")[0], label.split(" — ")[0]))) {
                throw new Error(`nhãn không khớp (ô này có nhãn “${truncate(label || "(không có)", 60)}”)`);
              }
              const cell = t.getCell(x.r, x.c);
              if (String(v[x.r][x.c] ?? "").trim()) cell.body.insertText(x.f.value, "Replace");
              else cell.body.insertText(x.f.value, "Start");
              await ctx.sync();
              filled.push(x.f.id);
            } catch (e: any) {
              skipped.push({ id: x.f.id, reason: why(e) });
            }
          }
        }

        if (ccs.length) {
          const controls = await loadControls(ctx);
          for (const { f, cc } of ccs) {
            try {
              const x = controls.find((y) => y.c.id === cc);
              if (!x) throw new Error("không còn ô điền mẫu này (văn bản đã đổi): gọi lại word_list_form_fields");
              const c = x.c;
              if (!(same(f.label, x.label) || same(f.label, c.title) || same(f.label, c.tag))) throw new Error(`nhãn không khớp (ô này là “${truncate(x.label, 60)}”)`);
              if (x.locked) throw new Error(`ô bị ${x.locked}: Word không cho sửa`);
              const type = String(c.type);
              // Mọi loại ô: dựng lại OOXML của cả ô rồi ghi đè cả ô (có theo dõi). Ghi chữ vào ô rồi Hoàn tác làm Word khóa ô
              // với mọi lệnh ghi sau (đo 18/11/2026) — thay cả ô thì Hoàn tác trả nguyên ô cũ.
              let want: { text?: string; checked?: boolean; date?: { d: number; m: number; y: number } };
              if (type === "CheckBox") {
                const v = checkboxValue(f.value);
                if (v === null) throw new Error(`ô đánh dấu cần true hoặc false, nhận “${truncate(f.value, 30)}”`);
                want = { checked: v };
              } else if (type === "DropDownList") {
                const item = matchListItem(x.items, f.value);
                if (!item) throw new Error(`không có trong danh sách của ô: ${x.items.map((s) => `“${s}”`).join(", ") || "(không đọc được danh sách)"}`);
                want = { text: item };
              } else if (type === "DatePicker") {
                const dt = parseDate(f.value);
                if (!dt) throw new Error(`ngày không đọc được “${truncate(f.value, 30)}” (dùng dd/MM/yyyy)`);
                want = { text: formatDate(dt, x.dateFormat), date: dt };
              } else if (type === "ComboBox") {
                want = { text: matchListItem(x.items, f.value) ?? f.value };
              } else {
                want = { text: f.value };
              }
              const ox = c.getRange("Whole").getOoxml();
              await ctx.sync();
              const refuse = controlRefusal(ox.value);
              if (refuse) throw new Error(refuse);
              const next = fillControlOoxml(ox.value, want);
              if (!next) throw new Error("không dựng được nội dung mới của ô (cấu trúc ô lạ): sửa ô này bằng tay");
              c.getRange("Whole").insertOoxml(next, "Replace");
              await ctx.sync();
              filled.push(f.id);
            } catch (e: any) {
              skipped.push({ id: f.id, reason: why(e) });
            }
          }
        }
      })
    );
    // Soát số tiền bằng số – bằng chữ ngay sau khi điền (khỏi một lượt gọi công cụ soát riêng)
    let amount: { i: number; message: string }[] = [];
    if (filled.length) {
      try {
        amount = await Word.run(async (ctx) => {
          const { rows } = await d.loadRows(ctx);
          return checkConsistency(rows, { checks: ["amount"], outline: buildOutline(rows) })
            .issues.slice(0, 5)
            .map((x) => ({ i: x.p, message: x.message }));
        });
      } catch {
        amount = [];
      }
    }
    return {
      ok: filled.length > 0,
      brief: `đã điền ${filled.length}/${fields.length} chỗ${skipped.length ? `, bỏ qua ${skipped.length}` : ""}${amount.length ? `, ${amount.length} chỗ số lệch chữ` : ""}`,
      filled,
      ...(skipped.length ? { skipped } : {}),
      amount_check: amount.length ? amount : "số tiền bằng số và bằng chữ trong văn bản khớp nhau (hoặc không có)",
      note: "Mọi chỗ đã điền được ghi bằng Theo dõi thay đổi. Mã chỗ điền đã đổi sau lệnh này: muốn điền tiếp thì gọi lại word_list_form_fields. Báo cho người dùng những chỗ bỏ qua và những chỗ còn trống vì thiếu thông tin.",
    };
  }

  /** Hỏi trước khi ghi đè ô đã có chữ (ô điền mẫu, ô bảng) hoặc bỏ đánh dấu ô đang đánh dấu (soát mã Đợt 8, #12).
   *  Chỗ trống và ô còn trống: không hỏi. */
  async function risk(a: any): Promise<string | null> {
    const fields: any[] = Array.isArray(a?.fields) ? a.fields : [];
    const ids = fields.map((f) => ({ f, p: parseFieldId(f?.id) })).filter((x) => x.p && x.p.kind !== "blank");
    if (!ids.length) return null;
    return Word.run(async (ctx) => {
      const over: string[] = [];
      const controls = ids.some((x) => x.p!.kind === "control") ? await loadControls(ctx) : [];
      const tables = ctx.document.body.tables;
      tables.load("items/values");
      await ctx.sync();
      const cur = new Map<number, string[][]>();
      for (const { f, p } of ids) {
        if (p!.kind === "control") {
          const x = controls.find((y) => y.c.id === (p as any).cc);
          if (!x) continue;
          if (String(x.c.type) === "CheckBox") {
            if (x.checked && checkboxValue(f.value) === false) over.push(`bỏ đánh dấu “${truncate(x.label, 30)}”`);
            continue;
          }
          const old = String(x.c.text ?? "").trim();
          if (old && !same(old, f.value)) over.push(`${truncate(x.label, 30)}: “${truncate(old, 30)}” → “${truncate(f.value, 30)}”`);
        } else if (p!.kind === "cell") {
          const q = p as { t: number; r: number; c: number };
          const t = tables.items[q.t];
          if (!t) continue;
          if (!cur.has(q.t)) cur.set(q.t, await d.tableValues(ctx, t, t.values || []));
          const old = String(cur.get(q.t)?.[q.r]?.[q.c] ?? "").trim();
          if (old && !same(old, f.value)) over.push(`ô bảng ${f.id}: “${truncate(old, 30)}” → “${truncate(f.value, 30)}”`);
        }
      }
      return over.length ? `Ghi đè ${over.length} chỗ đã có chữ: ${over.slice(0, 4).join(" · ")}${over.length > 4 ? ` · … và ${over.length - 4} chỗ nữa` : ""}.` : null;
    });
  }

  return {
    has: (name: string) => names.has(name),
    run: (name: string, a: any) => (name === "word_list_form_fields" ? list(a || {}) : fill(a || {})),
    risk: { word_fill_form: risk },
  };
}
