/* global Word, Office */
/** Công cụ Word — nhận xét và sửa đổi chờ duyệt: thêm / liệt kê / trả lời / đánh dấu nhận xét, liệt kê / duyệt sửa đổi. Dời nguyên văn từ createWordTools (B13 phần 2, nhóm 2 rà soát kiến trúc, 26/11/2026);
 *  lớp Hoàn tác, hàng đợi ghi và câu hỏi rủi ro ở lại wordTools.ts. Quy tắc: docs/quy-tac/word.md. */
import { officeErrorHint } from "../core/errors";
import { buildOutline, clauseAt, clauseName } from "../core/outline";
import { resolveAuthor } from "../core/worddiff";
import { RESULT_JSON_BUDGET, paragraphIndex, takeWithinBudget, truncate } from "./util";
import { targetRanges, indexOfIds, supports, strip, UNTRUSTED_NOTE, loadRows } from "./wordCore";

import type { Undo } from "./wordCore";
const NAMES = new Set(["word_add_comment", "word_list_comments", "word_reply_comment", "word_resolve_comment", "word_list_revisions", "word_review_revision"]);

export function createWordCommentTools(deps: { overview: () => Promise<any> }) {
  void deps;
  async function run(name: string, a: any, u: Undo): Promise<any> {
    const extra = u.extra;
    void extra;
    switch (name) {
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
            untrusted: UNTRUSTED_NOTE("nhận xét"),
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
            untrusted: UNTRUSTED_NOTE("sửa đổi"),
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

    }
    throw new Error(`Không có công cụ ${name}`);
  }
  return { has: (n: string) => NAMES.has(n), run };
}
