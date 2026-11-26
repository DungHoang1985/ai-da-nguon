/* global Word, Office */
/** Công cụ Word — sửa: thay chữ, sửa đoạn, chèn đoạn / điều khoản, đổi cấp, thay vùng chọn, định dạng, chèn / điền bảng, xóa đoạn. Dời nguyên văn từ createWordTools (B13 phần 2, nhóm 2 rà soát kiến trúc, 26/11/2026);
 *  lớp Hoàn tác, hàng đợi ghi và câu hỏi rủi ro ở lại wordTools.ts. Quy tắc: docs/quy-tac/word.md. */
import { buildOutline } from "../core/outline";
import { locateParagraph, planParagraphEdit } from "../core/worddiff";
import { escapeWordSearch, isMarkdownNumberedList, normalize2D, planTableFill, parseMarkdownLine, replaceLimit, truncate } from "./util";
import { WHOLE_PARAGRAPH_RATIO, WHOLE_PARAGRAPH_HUNKS, tracked, targetRanges, indexOfIds, restoreText, removeParas, supports, mapStyle, NOTE_MARK, noMarks, withNoteSpans, curOf, currentTableValues, loadRows, tableSpanProblem, resolvePara, missingExpect, pendingInfo, movedNote, liveHits, applyPlan } from "./wordCore";
import type { Loaded, UndoStep, Undo } from "./wordCore";

const NAMES = new Set(["word_replace_text", "word_set_paragraph_text", "word_insert_paragraphs", "word_insert_clause", "word_change_level", "word_replace_selection", "word_format_text", "word_insert_table", "word_delete_paragraphs", "word_fill_table"]);

export function createWordEditTools(deps: { overview: () => Promise<any> }) {
  void deps;
  async function run(name: string, a: any, u: Undo): Promise<any> {
    const extra = u.extra;
    void extra;
    switch (name) {
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
            // Mức 2 (4.12.0): trả kèm chữ HIỆN TẠI của các đoạn vừa sửa (tối đa 5) để AI tự đối chiếu, khỏi gọi thêm lệnh đọc
            let now: string[] | undefined;
            if (supports("1.4")) {
              const cur = targets.slice(0, 5).map((t) => {
                const p = t.paragraphs.getFirst();
                return supports("1.7") ? p.getText() : p.getReviewedText("Current");
              });
              await ctx.sync();
              now = [...new Set(cur.map((c) => truncate(noMarks(curOf(c.value, "")), 300)))];
            }
            return {
              ok: true,
              found: live.length,
              replaced: targets.length - same,
              message: `Đã thay ${targets.length - same} chỗ${where}`,
              ...(now ? { paragraphs_now: now } : {}),
            };
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
            return {
              ok: true,
              index: r.index,
              changes,
              message: `Đã sửa đoạn ${r.index} (${mode})${note}${supports("1.4") ? "; đã đọc lại: câu hiện tại đúng câu yêu cầu" : ""}`,
            };
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
            // Điều khoản chèn theo kiểu của điều neo: ký hiệu markdown ("## Điều 8.", "**…**") bỏ đi, không ghi thành chữ
            // (thử 23/11/2026: Cohere chèn "## Điều 8. Hồ sơ pháp lý bàn giao" vào hợp đồng)
            const lines = String(a.text ?? "")
              .split(/\r?\n/)
              .map((l) => l.trim().replace(/^#{1,4}\s+/, "").replace(/\*\*(.*?)\*\*/g, "$1"))
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
            const bad = await tableSpanProblem(ctx, r.paras, s, e);
            if (bad) throw new Error(bad);
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

    }
    throw new Error(`Không có công cụ ${name}`);
  }
  return { has: (n: string) => NAMES.has(n), run };
}
