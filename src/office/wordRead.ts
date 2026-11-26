/* global Word, Office */
/** Công cụ Word — đọc: tổng quan, mục lục, tìm, đọc điều / đoạn / bảng, chuyển tới đoạn. Dời nguyên văn từ createWordTools (B13 phần 2, nhóm 2 rà soát kiến trúc, 26/11/2026);
 *  lớp Hoàn tác, hàng đợi ghi và câu hỏi rủi ro ở lại wordTools.ts. Quy tắc: docs/quy-tac/word.md. */
import { buildOutline, clauseAt, clauseName, findClause } from "../core/outline";
import { RESULT_JSON_BUDGET, packParagraphs, PENDING_NOTE, searchRows, searchRowsByWords, takeWithinBudget, truncate } from "./util";
import { MAX_READ_CHARS, supports, strip, currentTableValues, loadRows, markCells, withNums, resolvePara, NOT_FOUND_PENDING, markPendingHits, annotatePending } from "./wordCore";
import type { PendingRow, Undo } from "./wordCore";

const NAMES = new Set(["word_get_overview", "word_get_outline", "word_find", "word_read_clause", "word_read_paragraphs", "word_read_tables", "word_go_to"]);

export function createWordReadTools(deps: { overview: () => Promise<any> }) {
  const { overview } = deps;
  async function run(name: string, a: any, u: Undo): Promise<any> {
    const extra = u.extra;
    void extra;
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
          const all = withNums(rows, outline) as { text: string; style: string; num?: string; changed?: boolean; deleted?: boolean; notes?: string[]; cell?: string }[];
          // Ngân sách tính theo độ dài JSON thật; next_start và truncated nằm đầu kết quả (không bị cắt mất)
          const start = Math.max(0, Number(a.start) || 0);
          // Không nạp sửa đổi cho quá 300 đoạn một lần: kết quả cũng chỉ trả vài chục tới vài trăm đoạn theo ngân sách
          const end = Math.min(rows.length, Number(a.end) || start + 200, start + 300);
          await markCells(ctx, rows, paras, all, start, end);
          if (hasChanges) await annotatePending(ctx, { rows, paras }, all, start, end, a.view === "original");
          const r = packParagraphs(all, start, end);
          return a.view === "original" ? { ...r, view: "original" } : r;
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
  return { has: (n: string) => NAMES.has(n), run };
}
