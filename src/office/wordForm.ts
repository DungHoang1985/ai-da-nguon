/** Đợt 8 — điền mẫu trong Word: liệt kê mọi chỗ cần điền và điền hết trong một lệnh (có Track Changes, một mục Hoàn tác).
 *  Ba loại chỗ điền:
 *   - ô điền mẫu (content control), mã c<id>: ô chữ, ô chữ định dạng, ô đánh dấu, ô chọn, ô chọn hoặc nhập, ô ngày;
 *   - chỗ trống gõ tay trong đoạn ("……", "____", "[●]", "ngày   tháng   năm"), mã b<đoạn>.<thứ tự trong đoạn>;
 *   - ô bảng trống đứng cạnh ô nhãn, mã t<bảng>.<hàng>.<cột>.
 *  Cách ghi từng loại theo thăm dò 18/11/2026 (`tasks/dot-8.md`): ô chọn hoặc nhập và ô ngày ghi vào vùng nội dung của ô
 *  (insertText trên ô báo GeneralException); ô chọn không nhận chữ → ghi lại OOXML của ô với mục đã chọn; chỗ trống tìm đúng
 *  trọn chuỗi và đúng lần xuất hiện. Từ chối các sửa đổi lệnh này thêm vào thì văn bản trở lại y cũ (đã đo cho mọi loại). */
import { ToolSpec } from "../core/agent/types";
import { ParaRow } from "../core/outline";
import {
  checkboxValue,
  countOccurrences,
  dateFormatFromOoxml,
  findBlanks,
  formatDate,
  listItemsFromOoxml,
  matchListItem,
  parseDate,
  fillControlOoxml,
} from "../core/formfill";
import { checkConsistency } from "../core/consistency";
import { buildOutline } from "../core/outline";
import { RESULT_JSON_BUDGET, takeWithinBudget, truncate } from "./util";

const CC_TYPES = ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"];
const TYPE_VI: Record<string, string> = {
  PlainText: "ô chữ",
  RichText: "ô chữ",
  CheckBox: "ô đánh dấu",
  DropDownList: "ô chọn",
  ComboBox: "ô chọn hoặc nhập",
  DatePicker: "ô ngày",
};
/** Word chỉ tìm chuỗi tới 255 ký tự */
const MAX_SEARCH = 255;
/** Ô chọn dài quá thì danh sách mục chỉ đưa bấy nhiêu mục đầu */
const MAX_ITEMS = 40;

export const WORD_FORM_SPECS: ToolSpec[] = [
  {
    name: "word_list_form_fields",
    description:
      "List every spot to fill in a form/template, each with an id and the label written before it: content controls (id c<n>; type text/checkbox/dropdown with its allowed items/combobox/date with its format; locked ones are marked), typed blanks inside paragraphs such as '……', '____', '[●]', 'ngày   tháng   năm' (id b<paragraph>.<k>; 'blank' is the exact blank text that will be replaced), and empty table cells next to a label cell (id t<table>.<row>.<col>). Call this first, then fill everything with ONE word_fill_form call.",
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
      "Fill many spots listed by word_list_form_fields in ONE call (tracked changes; one Undo for the whole call). fields: [{id, value, label}] — label = the label you got from the list (the tool refuses a spot whose label changed). Checkbox: value 'true' or 'false'. Dropdown: exactly one of its items. Date: dd/MM/yyyy (written in the control's own format). Typed blank: value replaces the whole blank text, e.g. blank 'ngày   tháng   năm' → 'ngày 18 tháng 11 năm'. Only fill spots you have real information for; never invent ID numbers, dates or amounts — leave them out and tell the user which spots are still empty.",
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
            required: ["id", "value"],
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
}

/** Ô điền mẫu còn sống: bỏ các ô đang chờ xóa có theo dõi. Ghi lại một ô chọn (OOXML) để lại ô cũ "chờ xóa" cạnh ô mới, cùng tên;
 *  đọc cả hai thì ô cũ trông như ô còn trống (đo 18/11/2026, Chạy thật: AI phải đọc lại đoạn để kiểm). */
export async function liveContentControls(ctx: Word.RequestContext, supports: (v: string) => boolean, props: string): Promise<Word.ContentControl[]> {
  const cc = ctx.document.getContentControls({ types: CC_TYPES as any });
  cc.load(props);
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
  return cc.items.filter((_, k) => {
    const types = tcs[k].items.map((x) => String(x.type));
    return !(types.length && types.every((t) => t === "Deleted"));
  });
}

export interface FormField {
  id: string;
  kind: "control" | "blank" | "cell";
  /** Loại ô điền mẫu (PlainText…) */
  type?: string;
  label: string;
  value?: string;
  empty: boolean;
  locked?: boolean;
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

const same = (a: unknown, b: unknown) =>
  String(a ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase() ===
  String(b ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

export function createWordFormTools(d: Deps) {
  const names = new Set(WORD_FORM_SPECS.map((s) => s.name));

  /** Ô điền mẫu trong thân văn bản, kèm OOXML cho ô chọn / ô ngày (để đọc danh sách mục và định dạng ngày) */
  async function loadControls(ctx: Word.RequestContext) {
    if (!d.supports("1.5")) return [];
    const items = await liveContentControls(ctx, d.supports, "items/id,items/tag,items/title,items/type,items/text,items/cannotEdit");
    const cc = { items };
    const paras = cc.items.map((c) => {
      const p = c.paragraphs.getFirstOrNullObject();
      p.load("text");
      return p;
    });
    const xmls = cc.items.map((c) => (/DropDownList|ComboBox|DatePicker/.test(String(c.type)) ? c.getRange("Whole").getOoxml() : null));
    await ctx.sync();
    return cc.items.map((c, k) => {
      const xml = xmls[k]?.value || "";
      return { c, para: paras[k].isNullObject ? "" : paras[k].text, items: xml ? listItemsFromOoxml(xml) : [], dateFormat: xml ? dateFormatFromOoxml(xml) : null };
    });
  }

  async function listFields(): Promise<FormField[]> {
    return Word.run(async (ctx) => {
      const out: FormField[] = [];
      for (const x of await loadControls(ctx)) {
        const c = x.c;
        const type = String(c.type);
        const text = String(c.text ?? "");
        // Nhãn: tên ô (alias); không có thì chữ của đoạn đứng trước ô
        const before = x.para && text ? x.para.slice(0, Math.max(0, x.para.indexOf(text))) : x.para;
        const label = String(c.title || "").trim() || truncate(before.replace(/\s+/g, " ").trim(), 80) || String(c.tag || "");
        out.push({
          id: `c${c.id}`,
          kind: "control",
          type,
          label,
          value: type === "CheckBox" ? String(/[☒☑✓✔]/.test(text)) : text,
          empty: type === "CheckBox" ? false : !text.trim(),
          ...(c.cannotEdit ? { locked: true } : {}),
          ...(x.items.length ? { items: x.items.slice(0, MAX_ITEMS) } : {}),
          ...(type === "DatePicker" && x.dateFormat ? { date_format: x.dateFormat } : {}),
        });
      }
      const { rows } = await d.loadRows(ctx);
      for (const r of rows) {
        if (r.deleted) continue;
        findBlanks(r.text).forEach((b, k) => out.push({ id: `b${r.i}.${k}`, kind: "blank", label: b.label, blank: b.str, after: b.after, empty: true, i: r.i }));
      }
      const tables = ctx.document.body.tables;
      tables.load("items/values,items/rowCount");
      await ctx.sync();
      tables.items.forEach((t, ti) => {
        const v = t.values || [];
        v.forEach((row, ri) =>
          row.forEach((cell, ci) => {
            if (ci === 0 || String(cell ?? "").trim()) return;
            const left = String(row[ci - 1] ?? "").trim();
            if (!left) return;
            const head = ri > 0 ? String(v[0]?.[ci] ?? "").trim() : "";
            out.push({ id: `t${ti}.${ri}.${ci}`, kind: "cell", label: head ? `${left} — ${head}` : left, empty: true });
          })
        );
      });
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
      ...(all.length ? {} : { note: "Không thấy ô điền mẫu, chỗ trống (……, ____, [●]) hay ô bảng trống cạnh ô nhãn. Nếu mẫu dùng cách khác, đọc văn bản rồi sửa bằng các công cụ sửa đoạn." }),
    };
  }

  /** Lỗi Office.js thành câu ngắn */
  const why = (e: any) => truncate(`${e?.message || e}${e?.code && !String(e?.message).includes(e.code) ? ` [${e.code}]` : ""}`, 160);

  async function fill(a: any) {
    const fields: { id: string; value: string; label?: string }[] = (Array.isArray(a.fields) ? a.fields : [])
      .filter((f: any) => f && typeof f === "object" && String(f.id ?? "").trim())
      .map((f: any) => ({ id: String(f.id).trim(), value: String(f.value ?? ""), label: f.label === undefined ? undefined : String(f.label) }));
    if (!fields.length) throw new Error("Thiếu fields: danh sách { id, value } lấy từ word_list_form_fields.");
    const filled: string[] = [];
    const skipped: { id: string; reason: string }[] = [];
    const seen = new Set<string>();
    await Word.run(async (ctx) =>
      d.tracked(ctx, async () => {
        const controls = fields.some((f) => /^c/i.test(f.id)) ? await loadControls(ctx) : [];
        const blanks: { f: (typeof fields)[number]; i: number; k: number }[] = [];
        const cells: { f: (typeof fields)[number]; t: number; r: number; c: number }[] = [];
        for (const f of fields) {
          const key = f.id.toLowerCase();
          if (seen.has(key)) {
            skipped.push({ id: f.id, reason: "trùng mã trong cùng lệnh" });
            continue;
          }
          seen.add(key);
          const p = parseFieldId(f.id);
          if (!p) {
            skipped.push({ id: f.id, reason: "mã không hợp lệ (dùng mã từ word_list_form_fields)" });
            continue;
          }
          if (p.kind === "blank") blanks.push({ f, i: p.i, k: p.k });
          else if (p.kind === "cell") cells.push({ f, t: p.t, r: p.r, c: p.c });
          else {
            const x = controls.find((y) => y.c.id === p.cc);
            if (!x) {
              skipped.push({ id: f.id, reason: "không còn ô điền mẫu này (văn bản đã đổi): gọi lại word_list_form_fields" });
              continue;
            }
            const c = x.c;
            const type = String(c.type);
            if (f.label !== undefined && c.title && !same(f.label, c.title) && !same(f.label, c.tag)) {
              skipped.push({ id: f.id, reason: `nhãn không khớp (ô này là “${c.title}”)` });
              continue;
            }
            if (c.cannotEdit) {
              skipped.push({ id: f.id, reason: "ô bị khóa nội dung, Word không cho sửa" });
              continue;
            }
            try {
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

        if (blanks.length) {
          const { paras, rows } = await d.loadRows(ctx);
          // Trong một đoạn: điền từ chỗ trống cuối lên — chữ đã xóa có theo dõi của chỗ sau không làm lệch thứ tự tìm của chỗ trước
          blanks.sort((x, y) => x.i - y.i || y.k - x.k);
          for (const b of blanks) {
            const row = rows[b.i];
            const found = row && !row.deleted ? findBlanks(row.text)[b.k] : undefined;
            if (!found) {
              skipped.push({ id: b.f.id, reason: "chỗ trống không còn ở đó (văn bản đã đổi): gọi lại word_list_form_fields" });
              continue;
            }
            if (b.f.label !== undefined && !same(b.f.label, found.label)) {
              skipped.push({ id: b.f.id, reason: `nhãn không khớp (chỗ này đứng sau “${truncate(found.label, 60)}”): gọi lại word_list_form_fields` });
              continue;
            }
            if (found.str.length > MAX_SEARCH) {
              skipped.push({ id: b.f.id, reason: "chỗ trống dài quá 255 ký tự, Word không tìm được: sửa bằng word_replace_text" });
              continue;
            }
            try {
              const hits = paras.items[b.i].search(found.str, { matchCase: true });
              hits.load("items");
              await ctx.sync();
              // Word tìm cả trong chữ đã xóa có theo dõi: số kết quả phải khớp số lần trong bản hiện tại thì mới chắc đúng chỗ
              if (hits.items.length !== countOccurrences(row.text, found.str)) {
                throw new Error("đoạn đang có sửa đổi chờ duyệt ở chỗ trống này nên không chắc đúng chỗ: duyệt sửa đổi trước hoặc sửa bằng word_replace_text");
              }
              hits.items[found.occurrence].insertText(b.f.value, "Replace");
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
          for (const x of cells) {
            const t = tables.items[x.t];
            const row = t?.values?.[x.r];
            if (!t || !row || x.c >= row.length) {
              skipped.push({ id: x.f.id, reason: "không còn ô bảng này (văn bản đã đổi): gọi lại word_list_form_fields" });
              continue;
            }
            const left = String(row[x.c - 1] ?? "").trim();
            if (x.f.label !== undefined && !same(String(x.f.label).split(" — ")[0], left)) {
              skipped.push({ id: x.f.id, reason: `nhãn không khớp (ô bên trái là “${truncate(left, 60)}”)` });
              continue;
            }
            try {
              const cell = t.getCell(x.r, x.c);
              if (String(row[x.c] ?? "").trim()) cell.body.insertText(x.f.value, "Replace");
              else cell.body.insertText(x.f.value, "Start");
              await ctx.sync();
              filled.push(x.f.id);
            } catch (e: any) {
              skipped.push({ id: x.f.id, reason: why(e) });
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
      note: "Mọi chỗ đã điền được ghi bằng Theo dõi thay đổi. Báo cho người dùng những chỗ bỏ qua và những chỗ còn trống vì thiếu thông tin.",
    };
  }

  /** Hỏi trước khi ghi đè ô đã có chữ (ô điền mẫu, ô bảng). Chỗ trống và ô còn trống: không hỏi. */
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
      for (const { f, p } of ids) {
        if (p!.kind === "control") {
          const x = controls.find((y) => y.c.id === (p as any).cc);
          if (!x || String(x.c.type) === "CheckBox") continue;
          const old = String(x.c.text ?? "").trim();
          if (old && !same(old, f.value)) over.push(`${x.c.title || f.id}: “${truncate(old, 30)}” → “${truncate(f.value, 30)}”`);
        } else if (p!.kind === "cell") {
          const q = p as { t: number; r: number; c: number };
          const old = String(tables.items[q.t]?.values?.[q.r]?.[q.c] ?? "").trim();
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
    typeLabel: (t: string) => TYPE_VI[t] || t,
  };
}
