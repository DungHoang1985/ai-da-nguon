/* global Word, Office */
/** Tự kiểm tra Word (bản 4.0): chạy thật từng công cụ trong một VÙNG CÔ LẬP rồi dọn sạch.
 *  Vùng thử nằm giữa hai đoạn mốc, chèn ngay trước đoạn cuối của văn bản; mọi bước chỉ đụng tới vùng này.
 *  Bước cuối so lại phần văn bản ngoài vùng thử (số đoạn, chữ, kiểu, số sửa đổi, số nhận xét, chế độ theo dõi): phải y nguyên. */
import { HostToolSet, UndoEntry } from "../core/agent/types";

interface Kit {
  tools: HostToolSet;
  step: (name: string, fn: () => Promise<string | void>) => Promise<void>;
  run: (name: string, args?: any) => Promise<any>;
  expect: (cond: any, msg: string) => void;
  riskOf: (name: string, args: any) => Promise<string | null>;
  log: UndoEntry[];
  Skip: new (m?: string) => Error;
}

const sup = (v: string, set = "WordApi") => Office.context.requirements.isSetSupported(set, v);

export async function runWordSelfTest({ tools, step, run, expect, riskOf, log, Skip }: Kit): Promise<void> {
  const code = "AIDNKT" + Date.now().toString(36).slice(-5).toUpperCase();
  const START = `${code} VÙNG THỬ CỦA AI ĐA NGUỒN – tự xóa khi kiểm tra xong`;
  const END = `${code} HẾT VÙNG THỬ`;
  const T = {
    art1: `Thanh toán ${code}`,
    c11: `Bên B thanh toán trong 30 ngày; chậm quá 30 ngày thì chịu phạt 30 triệu đồng và lãi 0,05% mỗi ngày ${code}.`,
    c12: `Bên A xuất hóa đơn trong 05 ngày ${code}.`,
    c12a: `bằng chuyển khoản ${code};`,
    c13: `Bên A giao hàng tại kho ${code}.`,
    art2: `Bảo hành ${code}`,
    c21: `Thời hạn bảo hành ${code} là 12 tháng.`,
    typed: `Điều 97. Điều khoản gõ tay ${code}`,
    typed1: `97.1 Thời hạn thuê ${code} là 24 tháng kể từ ngày ký.`,
    plain: `Hợp đồng số ${code}-01/2026 ký ngày 01/01/2026.`,
  };
  /** Trạng thái phần văn bản ngoài vùng thử, chụp trước khi dựng vùng thử */
  let outside = { paras: 0, text: "", styles: "", changes: -1, comments: -1, mode: "" };
  let built = false;

  /** Ảnh chụp phần ngoài vùng thử: mọi đoạn trừ các đoạn từ mốc đầu tới mốc cuối */
  const snapshotOutside = () =>
    Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text,items/style");
      ctx.document.load("changeTrackingMode");
      const tc = sup("1.6") ? ctx.document.body.getTrackedChanges() : null;
      tc?.load("items/type");
      const cs = sup("1.4") ? ctx.document.body.getComments() : null;
      cs?.load("items/id");
      await ctx.sync();
      const s = ps.items.findIndex((p) => p.text === START);
      const e = ps.items.findIndex((p) => p.text === END);
      const out = s >= 0 && e >= s ? ps.items.filter((_, i) => i < s || i > e) : ps.items;
      return {
        paras: out.length,
        text: out.map((p) => p.text).join("\n"),
        styles: out.map((p) => p.style).join("|"),
        changes: tc ? tc.items.length : -1,
        comments: cs ? cs.items.length : -1,
        mode: String(ctx.document.changeTrackingMode),
        regionChanges: 0,
      };
    });

  /** Chỉ số đoạn đầu tiên trong vùng thử có chữ chứa `part` (-1 nếu không có) */
  const at = (part: string) =>
    Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text");
      await ctx.sync();
      const s = ps.items.findIndex((p) => p.text === START);
      return ps.items.findIndex((p, i) => i >= s && s >= 0 && p.text.includes(part));
    });
  const need = async (part: string) => {
    const i = await at(part);
    expect(i >= 0, `Không thấy đoạn thử “${part.slice(0, 40)}”`);
    return i;
  };
  /** Chạy mã Office.js với Theo dõi thay đổi tắt, xong trả lại chế độ cũ */
  const untracked = <R>(fn: (ctx: Word.RequestContext) => Promise<R>) =>
    Word.run(async (ctx) => {
      ctx.document.load("changeTrackingMode");
      await ctx.sync();
      const mode = ctx.document.changeTrackingMode;
      if (sup("1.4")) ctx.document.changeTrackingMode = "Off" as any;
      try {
        return await fn(ctx);
      } finally {
        if (sup("1.4")) {
          ctx.document.changeTrackingMode = mode;
          await ctx.sync();
        }
      }
    });
  /** Thông tin thật của một đoạn: bản hiện tại, bản gốc, các sửa đổi trong đoạn */
  const inspect = (index: number) =>
    Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text,items/style");
      await ctx.sync();
      const p = ps.items[index];
      const cur = sup("1.7") ? p.getText() : p.getReviewedText("Current");
      const orig = p.getReviewedText("Original");
      const tc = p.getTrackedChanges();
      tc.load("items/type,items/text");
      await ctx.sync();
      const clean = (s: string) => s.replace(/[\r\n\u0007]+$/, "");
      return { current: clean(cur.value), original: clean(orig.value), style: p.style, changes: tc.items.map((t) => `${t.type}:${clean(t.text)}`) };
    });
  const countAll = () =>
    Word.run(async (ctx) => {
      const tc = ctx.document.body.getTrackedChanges();
      tc.load("items/type");
      const cs = ctx.document.body.getComments();
      cs.load("items/id");
      await ctx.sync();
      return { changes: tc.items.length, comments: cs.items.length };
    });
  /** Mục lục của riêng vùng thử (đọc hết các trang) */
  const regionOutline = async () => {
    const s = await need(START);
    const e = await need(END);
    const items: any[] = [];
    let start = 0;
    for (let guard = 0; guard < 60; guard++) {
      const r = await run("word_get_outline", { start });
      items.push(...r.items);
      if (r.next_start === undefined) break;
      start = r.next_start;
    }
    return items.filter((o) => o.i > s && o.i < e);
  };

  await step("Bộ API Word máy hỗ trợ", async () => {
    const api = ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"].filter((v) => sup(v));
    const desk = ["1.1", "1.2", "1.3", "1.4", "1.5"].filter((v) => sup(v, "WordApiDesktop"));
    if (!sup("1.6")) throw new Skip(`WordApi ${api[api.length - 1] || "< 1.3"}: thiếu 1.6 nên các bước sửa có theo dõi, định danh đoạn và Hoàn tác không chạy được`);
    return `WordApi ${api[0]}–${api[api.length - 1]} · WordApiDesktop ${desk.length ? `${desk[0]}–${desk[desk.length - 1]}` : "không có"}`;
  });
  if (!sup("1.6")) return;

  await step("Vùng cô lập: dựng vùng thử trước đoạn cuối, chụp lại phần văn bản còn lại", async () => {
    outside = await snapshotOutside();
    await untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const last = ps.items[ps.items.length - 1];
      const texts = [START, T.art1, T.c11, T.c12, T.c12a, T.c13, T.art2, T.c21, T.typed, T.typed1, T.plain, END];
      const made: Word.Paragraph[] = [];
      texts.forEach((t, k) => made.push(k === 0 ? last.insertParagraph(t, "Before") : made[k - 1].insertParagraph(t, "After")));
      made.forEach((p) => p.load("isListItem"));
      await ctx.sync();
      // Đoạn mới thừa hưởng định dạng của đoạn cuối: đưa hết về chữ thường, không danh sách
      made.forEach((p) => {
        if (p.isListItem) p.detachFromList();
        p.styleBuiltIn = "Normal" as any;
      });
      await ctx.sync();
      // Danh sách nhiều cấp do Word tự đánh: "Điều 1." / "1.1" / "(a)"
      const list = made[1].startNewList();
      list.setLevelNumbering(0, "Arabic" as any, ["Điều ", 0, "."]);
      list.setLevelNumbering(1, "Arabic" as any, [0, ".", 1]);
      list.setLevelNumbering(2, "LowerLetter" as any, ["(", 2, ")"]);
      list.load("id");
      await ctx.sync();
      const levels: [number, number][] = [
        [2, 1],
        [3, 1],
        [4, 2],
        [5, 1],
        [6, 0],
        [7, 1],
      ];
      levels.forEach(([k, lv]) => made[k].attachToList(list.id, lv));
      await ctx.sync();
      // Chữ đậm giữa câu: "30 ngày" thứ hai
      const hits = made[2].search("30 ngày", { matchCase: true });
      hits.load("items");
      await ctx.sync();
      hits.items[1].font.bold = true;
      await ctx.sync();
    });
    // Một sửa đổi chờ duyệt có sẵn (giả lập bản đối tác gửi lại): "24" → "12"
    await Word.run(async (ctx) => {
      ctx.document.load("changeTrackingMode");
      await ctx.sync();
      const mode = ctx.document.changeTrackingMode;
      ctx.document.changeTrackingMode = "TrackAll" as any;
      const hit = ctx.document.body.search(`${code} là 24`, { matchCase: true });
      hit.load("items");
      await ctx.sync();
      const n = hit.items[0].search("24", { matchCase: true });
      n.load("items");
      await ctx.sync();
      n.items[0].insertText("12", "Replace");
      await ctx.sync();
      ctx.document.changeTrackingMode = mode;
      await ctx.sync();
    });
    built = true;
    const after = await snapshotOutside();
    expect(after.paras === outside.paras && after.text === outside.text, "Dựng vùng thử làm đổi phần văn bản bên ngoài");
    return `Văn bản có ${outside.paras} đoạn, ${outside.changes} sửa đổi, ${outside.comments} nhận xét ngoài vùng thử; vùng thử 12 đoạn đặt trước đoạn cuối`;
  });
  if (!built) return;

  await step("Xem tổng quan văn bản", async () => {
    const ov = await run("word_get_overview");
    expect(ov.paragraphs === outside.paras + 12, `Tổng quan báo ${ov.paragraphs} đoạn, đúng phải là ${outside.paras + 12}`);
    expect(ov.clauses >= 9, `Tổng quan báo ${ov.clauses} điều khoản`);
    expect(ov.tracked_changes === outside.changes + 2, `Số sửa đổi chờ duyệt: ${ov.tracked_changes}`);
    return `${ov.paragraphs} đoạn, ${ov.clauses} điều khoản, ${ov.tracked_changes} sửa đổi chờ duyệt, Track Changes: ${ov.track_changes ?? "không rõ"}`;
  });

  await step("Mục lục điều khoản: số Word tự đánh (3 cấp) và số gõ tay", async () => {
    const items = await regionOutline();
    const got = items.map((o) => o.num).join(" | ");
    const want = ["Điều 1", "1.1", "1.2", "1.2(a)", "1.3", "Điều 2", "2.1", "Điều 97", "97.1"].join(" | ");
    expect(got === want, `Mục lục vùng thử: ${got} (đúng phải là ${want})`);
    const a1 = items[0];
    expect(a1.end === (await need(T.c13)), `Điều 1 phải kéo tới khoản 1.3, công cụ báo end = ${a1.end}`);
    expect(String(a1.title).includes("Thanh toán"), `Tiêu đề Điều 1: ${a1.title}`);
    return `${items.length} điều khoản: ${got}`;
  });

  await step("Đọc trọn một điều theo số (số gõ tay và cách gọi 'khoản … Điều …')", async () => {
    const r = await run("word_read_clause", { number: "Điều 97" });
    expect(r.ok && r.from === (await need(T.typed)) && r.paragraphs?.[0]?.num === "Điều 97", `Đọc Điều 97: ${JSON.stringify(r).slice(0, 200)}`);
    expect(r.paragraphs[1].num === "97.1", `Đoạn thứ hai phải là 97.1: ${r.paragraphs[1].num}`);
    const r2 = await run("word_read_clause", { number: "khoản 1 Điều 97" });
    expect(r2.ok && r2.paragraphs?.[0]?.num === "97.1" && r2.from === (await need(`97.1 Thời hạn thuê ${code}`)), `Đọc khoản 1 Điều 97: ${JSON.stringify(r2).slice(0, 200)}`);
    const auto = await run("word_read_clause", { number: "1.2(a)" });
    const cand = auto.ok ? [{ i: auto.from }] : auto.candidates || [];
    expect(cand.some((c: any) => c.i === 0 + (r.from - 4)), `Đọc 1.2(a) (số Word tự đánh): ${JSON.stringify(auto).slice(0, 200)}`);
    const none = await run("word_read_clause", { number: "Điều 987" });
    expect(none.ok === false && Array.isArray(none.top_level), "Điều không có mà không báo");
    return `Điều 97 (gõ tay): từ [¶${r.from}]; khoản 1 Điều 97: [¶${r2.from}]; 1.2(a) (Word tự đánh): [¶${r.from - 4}]; Điều 987: báo không có`;
  });

  await step("Đọc đoạn: kèm số điều; bản hiện tại không lẫn chữ đã xóa; đọc được bản gốc", async () => {
    const s = await need(START);
    const r = await run("word_read_paragraphs", { start: s, end: s + 12 });
    const c11 = r.paragraphs.find((p: any) => p.text === T.c11);
    expect(c11 && c11.num === "1.1", `Khoản 1.1: ${JSON.stringify(c11)}`);
    const t1 = r.paragraphs.find((p: any) => p.num === "97.1");
    expect(t1 && t1.text.includes("là 12 tháng") && !t1.text.includes("24"), `Bản hiện tại của 97.1: ${t1?.text}`);
    expect(t1.changed === true, "Đoạn đang có sửa đổi chờ duyệt mà không có cờ changed");
    expect(!c11.changed, "Đoạn sạch mà có cờ changed");
    const o = await run("word_read_paragraphs", { start: t1.i, end: t1.i + 1, view: "original" });
    expect(o.paragraphs[0].text.includes("là 24 tháng") && !o.paragraphs[0].text.includes("12"), `Bản gốc của 97.1: ${o.paragraphs[0].text}`);
    return `1.1 có số điều; 97.1 hiện tại “…${t1.text.slice(-32)}”, bản gốc “…${o.paragraphs[0].text.slice(-28)}”`;
  });

  await step("Tìm trong văn bản: trả đoạn kèm số điều; bỏ qua chữ đã xóa; tìm được cụm bị chữ đã xóa chen giữa", async () => {
    const r = await run("word_find", { query: `bảo hành ${code}`, match_case: true });
    expect(r.found === 1 && r.hits[0].i === (await need(T.c21)) && r.hits[0].clause === "2.1", `Tìm “bảo hành ${code}”: ${JSON.stringify(r).slice(0, 240)}`);
    const dead = await run("word_find", { query: `${code} là 24 tháng` });
    expect(dead.found === 0, `Cụm chỉ còn trong chữ đã xóa mà vẫn tìm ra ${dead.found}`);
    const across = await run("word_find", { query: `thuê ${code} là 12 tháng` });
    expect(across.found === 1 && across.hits[0].clause === "97.1", `Cụm bị chữ đã xóa chen giữa: ${JSON.stringify(across).slice(0, 200)}`);
    const noAccent = await run("word_find", { query: `thoi han bao hanh ${code}` });
    expect(noAccent.found === 1, `Tìm không dấu: ${noAccent.found}`);
    return `“bảo hành ${code}” → [¶${r.hits[0].i}] thuộc 2.1; cụm đã xóa: 0; cụm bị chen giữa: tìm thấy ở 97.1; gõ không dấu vẫn tìm ra`;
  });

  await step("Sửa mức từ: hai chỗ cách xa trong một đoạn → đúng hai cặp xóa/thêm, chữ đậm còn nguyên", async () => {
    const i = await need(T.c11);
    const want = T.c11.replace("quá 30 ngày", "quá 45 ngày").replace("0,05%", "0,03%");
    const res = await run("word_set_paragraph_text", { index: i, expect: T.c11.slice(0, 30), text: want });
    expect(res.ok && res.changes === 2, `Kết quả: ${JSON.stringify(res)}`);
    const info = await inspect(i);
    expect(info.current === want, `Câu sau sửa: ${info.current}`);
    const got = info.changes.slice().sort().join(" · ");
    const exp = ["Added:0,03%", "Added:45", "Deleted:0,05%", "Deleted:30"].sort().join(" · ");
    expect(got === exp, `Sửa đổi trong đoạn: ${got} (đúng phải là ${exp})`);
    const bold = await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const a = ps.items[i].search("45", { matchCase: true });
      const b = ps.items[i].search("triệu", { matchCase: true });
      a.load("items/font/bold");
      b.load("items/font/bold");
      await ctx.sync();
      return { new45: a.items[0]?.font.bold, plain: b.items[0]?.font.bold };
    });
    expect(bold.new45 === true && bold.plain === false, `Chữ đậm: “45” = ${bold.new45}, “triệu” = ${bold.plain}`);
    return `${res.message}: ${got}; “45” vẫn đậm, chữ hai bên vẫn thường`;
  });

  await step("Sửa trên đoạn đang có sửa đổi chờ duyệt: bản gốc và sửa đổi cũ còn nguyên", async () => {
    const i = await need(`97.1 Thời hạn thuê ${code}`);
    const want = `97.1 Thời hạn thuê ${code} là 18 tháng kể từ ngày bàn giao.`;
    const res = await run("word_set_paragraph_text", { index: i, expect: "97.1 Thời hạn thuê", text: want });
    const info = await inspect(i);
    expect(res.ok && info.current === want, `Câu sau sửa: ${info.current} · ${JSON.stringify(res)}`);
    expect(info.original === T.typed1, `Bản gốc bị đổi: ${info.original}`);
    expect(info.changes.includes("Deleted:24"), `Sửa đổi cũ (xóa “24”) không còn: ${info.changes.join(" · ")}`);
    return `Hiện tại “…${want.slice(-36)}”; bản gốc y nguyên; ${info.changes.length} sửa đổi trong đoạn`;
  });

  await step("Chống sửa nhầm đoạn: chèn hai đoạn phía trên rồi sửa theo số đoạn cũ → tự định vị lại; sai nội dung → từ chối", async () => {
    const old = await need(T.c21);
    const s = await need(START);
    const ins = await run("word_insert_paragraphs", { position: "after_paragraph", paragraph_index: s, expect: code, text: `Đoạn chèn thêm 1 ${code}\nĐoạn chèn thêm 2 ${code}` });
    expect(ins.first_index === s + 1 && ins.count === 2, `Kết quả chèn: ${JSON.stringify(ins)}`);
    const want = `Thời hạn bảo hành ${code} là 24 tháng.`;
    const res = await run("word_set_paragraph_text", { index: old, expect: "Thời hạn bảo hành", text: want });
    expect(res.ok && res.index === old + 2, `Sửa theo số đoạn cũ ${old}: ${JSON.stringify(res)}`);
    const moved = await inspect(old + 2);
    const stale = await inspect(old);
    expect(moved.current === want, `Đoạn đúng sau sửa: ${moved.current}`);
    expect(!stale.current.includes("24 tháng"), `Đoạn ở số cũ bị sửa nhầm: ${stale.current}`);
    let refused = "";
    try {
      await run("word_set_paragraph_text", { index: old, expect: "Nội dung này không có trong văn bản", text: "x" });
    } catch (e: any) {
      refused = e.message;
    }
    expect(/Đọc lại/.test(refused), `Sai nội dung mà không từ chối: ${refused}`);
    return `Số đoạn cũ ${old} → công cụ tự sửa đúng [¶${old + 2}]; nội dung không khớp → từ chối: “${refused.slice(0, 70)}…”`;
  });

  await step("Thay thế có phạm vi: nhiều chỗ khớp thì không thay; chỉ rõ đoạn + lần thứ mấy thì thay đúng một chỗ", async () => {
    const i = await need(T.c11.slice(0, 20));
    const before = await inspect(i);
    const many = await run("word_replace_text", { find: "30", replace: "60", paragraph_index: i, expect: "Bên B thanh toán" });
    expect(many.ok === false && many.replaced === 0 && many.found === 2, `Hai chỗ “30” còn sống (chỗ thứ ba đã thành 45): ${JSON.stringify(many).slice(0, 220)}`);
    expect((await inspect(i)).current === before.current, "Chưa chỉ rõ mà đã thay");
    const one = await run("word_replace_text", { find: "30", replace: "60", paragraph_index: i, expect: "Bên B thanh toán", occurrence: 2 });
    const after = await inspect(i);
    expect(one.replaced === 1 && after.current.includes("phạt 60 triệu") && after.current.includes("trong 30 ngày"), `Sau khi thay lần thứ 2: ${after.current}`);
    const uniq = await run("word_replace_text", { find: `${code}-01/2026`, replace: `${code}-02/2026` });
    expect(uniq.replaced === 1 && (await at(`${code}-02/2026`)) >= 0, `Thay chỗ duy nhất: ${JSON.stringify(uniq)}`);
    return `2 chỗ khớp: không thay, trả ${many.places?.length} vị trí; lần thứ 2 trong đoạn: “30 triệu” → “60 triệu”; chỗ duy nhất trong văn bản: thay ngay`;
  });

  await step("Tắt Track Changes rồi AI sửa: vẫn ghi có theo dõi, xong trả lại chế độ cũ", async () => {
    const i = await need(T.art2);
    const before = await Word.run(async (ctx) => {
      ctx.document.load("changeTrackingMode");
      await ctx.sync();
      const prev = ctx.document.changeTrackingMode;
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await ctx.sync();
      return { prev, n: (await countAll()).changes };
    });
    let after = { n: 0, mode: "" };
    let replaced = 0;
    try {
      replaced = (await run("word_replace_text", { find: "Bảo hành", replace: "Bảo hành sản phẩm", paragraph_index: i, expect: "Bảo hành" })).replaced;
    } finally {
      after = await Word.run(async (ctx) => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        const mode = String(ctx.document.changeTrackingMode);
        ctx.document.changeTrackingMode = before.prev;
        await ctx.sync();
        return { n: (await countAll()).changes, mode };
      });
    }
    expect(replaced === 1, `Thay ${replaced} chỗ (phải là 1)`);
    expect(after.n > before.n, `Sửa đổi được theo dõi: ${before.n} → ${after.n}`);
    expect(after.mode === "Off", `Sau khi AI sửa, chế độ là ${after.mode} (phải trả lại Off như người dùng đặt)`);
    return `Track Changes đang tắt: sửa đổi vẫn được ghi (${before.n} → ${after.n}), trả lại chế độ Off`;
  });

  await step("Chèn nội dung tự do: tiêu đề, gạch đầu dòng; dòng '5. …' đứng lẻ giữ nguyên số", async () => {
    const i = await need(`${code}-02/2026`);
    const res = await run("word_insert_paragraphs", {
      position: "after_paragraph",
      paragraph_index: i,
      expect: "Hợp đồng số",
      text: `## Mục kiểm tra ${code}\n5. Thanh toán bổ sung ${code}\n- Ý thứ nhất ${code}`,
    });
    expect(res.first_index === i + 1, `first_index = ${res.first_index}, đúng phải là ${i + 1}`);
    const r = await run("word_read_paragraphs", { start: i + 1, end: i + 4 });
    const [h, n5, b] = r.paragraphs;
    expect(/2|heading|tiêu đề/i.test(String(h.style)) && h.text === `Mục kiểm tra ${code}`, `Tiêu đề: ${JSON.stringify(h)}`);
    expect(n5.text === `5. Thanh toán bổ sung ${code}`, `Dòng số gõ tay: ${n5.text}`);
    const isBullet = await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/isListItem");
      await ctx.sync();
      return ps.items[b.i].isListItem;
    });
    expect(isBullet && b.text === `Ý thứ nhất ${code}`, `Dòng gạch đầu dòng: ${JSON.stringify(b)}`);
    return `Tiêu đề kiểu "${h.style}"; “5. Thanh toán bổ sung” giữ nguyên số; gạch đầu dòng là danh sách thật`;
  });

  await step("Thêm một khoản sau 1.2: khoản mới mang số 1.3, khoản cũ thành 1.4, cùng kiểu; thêm khoản con", async () => {
    const i12 = await need(T.c12);
    const before = await inspect(i12);
    const res = await run("word_insert_clause", { after_index: i12, expect: "Bên A xuất hóa đơn", text: `Khoản mới chèn ${code}.` });
    expect(res.ok && res.inserted?.[0]?.num === "1.3", `Khoản mới: ${JSON.stringify(res).slice(0, 240)}`);
    expect(res.inserted[0].i === (await need(T.c12a)) + 1, "Khoản mới phải nằm sau toàn bộ khoản 1.2 (sau điểm a)");
    const items = await regionOutline();
    const old13 = items.find((o) => String(o.title).includes("Bên A giao hàng tại kho"));
    expect(old13?.num === "1.4", `Khoản 1.3 cũ nay là ${old13?.num}`);
    const made = await inspect(res.inserted[0].i);
    expect(made.style === before.style, `Kiểu khoản mới “${made.style}” khác khoản bên cạnh “${before.style}”`);
    const sub = await run("word_insert_clause", { after_index: await need(T.c21.slice(0, 20)), expect: "Thời hạn bảo hành", text: `điểm con mới ${code};`, level: "deeper" });
    expect(sub.inserted?.[0]?.num === "2.1(a)", `Khoản con: ${JSON.stringify(sub.inserted)}`);
    return `Sau 1.2: khoản mới = ${res.inserted[0].num}, khoản cũ thành ${old13.num}, cùng kiểu “${made.style}”; khoản con của 2.1 = ${sub.inserted[0].num}`;
  });

  await step("Đổi cấp điều khoản", async () => {
    const i = await need(`điểm con mới ${code}`);
    const mark = log.length;
    const res = await run("word_change_level", { index: i, expect: "điểm con mới", direction: "shallower" });
    expect(res.ok && res.num === "2.2", `Sau khi lên một cấp: ${JSON.stringify(res)}`);
    // Hoàn tác việc đổi cấp rồi làm lại
    const entries = log.slice(mark);
    expect(entries.length === 1, `Đổi cấp ghi ${entries.length} mục hoàn tác (phải là 1)`);
    for (const x of entries.reverse()) await x.undo();
    log.length = mark;
    const back = (await regionOutline()).find((o) => o.i === i);
    expect(back?.num === "2.1(a)", `Hoàn tác đổi cấp: số điều là ${back?.num} (phải về 2.1(a))`);
    const again = await run("word_change_level", { index: i, expect: "điểm con mới", direction: "shallower" });
    expect(again.num === "2.2", `Đổi cấp lần hai: ${JSON.stringify(again)}`);
    return `${res.message}; hoàn tác thì về 2.1(a); đổi lại thành 2.2`;
  });

  await step("Định dạng chữ: phải chỉ rõ vị trí; đọc lại thấy đậm và đổi màu", async () => {
    let refused = "";
    try {
      await run("word_format_text", { bold: true });
    } catch (e: any) {
      refused = e.message;
    }
    expect(/đúng một vị trí/.test(refused), `Không chỉ vị trí mà không từ chối: ${refused}`);
    const i = await need(`Mục kiểm tra ${code}`);
    let noExpect = "";
    try {
      await run("word_format_text", { find: `kiểm tra ${code}`, paragraph_index: i, bold: true });
    } catch (e: any) {
      noExpect = e.message;
    }
    expect(/Thiếu expect/.test(noExpect), `Chỉ đoạn bằng số mà thiếu expect vẫn không từ chối: ${noExpect}`);
    await run("word_format_text", { find: `kiểm tra ${code}`, paragraph_index: i, expect: "Mục kiểm tra", bold: true, font_color: "#C00000" });
    const f = await Word.run(async (ctx) => {
      const hit = ctx.document.body.search(`kiểm tra ${code}`, { matchCase: true });
      hit.load("items/font/bold,items/font/color");
      await ctx.sync();
      return hit.items[0].font;
    });
    expect(f.bold === true && /C00000/i.test(String(f.color)), `Định dạng đọc lại: đậm ${f.bold}, màu ${f.color}`);
    return `“kiểm tra ${code}”: đậm, màu ${f.color}`;
  });

  await step("Thay nội dung đang chọn", async () => {
    await Word.run(async (ctx) => {
      const hit = ctx.document.body.search(`Đoạn chèn thêm 1 ${code}`, { matchCase: true });
      hit.load("items");
      await ctx.sync();
      hit.items[0].select();
      await ctx.sync();
    });
    await run("word_replace_selection", { text: `Đoạn đã thay vùng chọn ${code}` });
    const i = await need(`Đoạn đã thay vùng chọn ${code}`);
    expect((await inspect(i)).current === `Đoạn đã thay vùng chọn ${code}`, "Vùng chọn không được thay đúng");
    return `Đã thay vùng chọn ở [¶${i}]`;
  });

  let commentId = "";
  await step("Thêm nhận xét: phải chỉ rõ vị trí và có nội dung", async () => {
    let refused = "";
    try {
      await run("word_add_comment", { text: "Nhận xét không có vị trí" });
    } catch (e: any) {
      refused = e.message;
    }
    expect(/đúng một vị trí/.test(refused), `Không chỉ vị trí mà không từ chối: ${refused}`);
    const res = await run("word_add_comment", { find: `Mục kiểm tra ${code}`, text: "Nhận xét thử của AI Đa Nguồn" });
    commentId = res.id;
    const n = (await countAll()).comments;
    expect(n === outside.comments + 1, `Số nhận xét: ${n}, đúng phải là ${outside.comments + 1}`);
    return `${res.message}; thiếu vị trí → từ chối`;
  });

  await step("Đọc nhận xét kèm đoạn được neo và số điều", async () => {
    const r = await run("word_list_comments", {});
    const c = r.comments.find((x: any) => x.id === commentId);
    expect(c && String(c.anchor).includes(code), `Không thấy nhận xét thử: ${JSON.stringify(r.comments).slice(0, 200)}`);
    expect(c.paragraph === (await need(`Mục kiểm tra ${code}`)), `Nhận xét báo đoạn ${c.paragraph}`);
    return `Neo vào "${c.anchor}" ở [¶${c.paragraph}]${c.clause ? `, thuộc ${c.clause}` : ""}`;
  });

  await step("Trả lời và đánh dấu đã xử lý nhận xét", async () => {
    if (!commentId) throw new Skip("Không có nhận xét để trả lời");
    await run("word_reply_comment", { id: commentId, text: "Đã sửa theo nhận xét" });
    await run("word_resolve_comment", { id: commentId });
    const r = await run("word_list_comments", { include_resolved: true });
    const c = r.comments.find((x: any) => x.id === commentId);
    expect(c && c.replies?.length >= 1 && c.resolved === true, `Trạng thái: ${JSON.stringify(c)}`);
    return "Đã trả lời và đánh dấu đã xử lý";
  });

  await step("Chèn bảng; cảnh báo rủi ro khi ghi đè ô có dữ liệu; đọc và điền bảng", async () => {
    const i = await need(`Ý thứ nhất ${code}`);
    await run("word_insert_table", {
      position: "after_paragraph",
      paragraph_index: i,
      expect: "Ý thứ nhất",
      values: [
        [`Mục ${code}`, "Giá trị"],
        ["A", "1"],
      ],
    });
    const r = await run("word_read_tables", {});
    const t0 = r.tables.find((t: any) => t.values?.[0]?.[0] === `Mục ${code}`);
    expect(t0 && t0.values[1][0] === "A", `Không thấy bảng thử trong ${r.count} bảng`);
    const warn = await riskOf("word_fill_table", { table_index: t0.index, values: [["A", "10"]], start_row: 1 });
    expect(warn && /Ghi đè 1 ô/.test(warn), `Ghi đè ô "1" bằng "10" mà lời cảnh báo là: ${warn}`);
    const same = await riskOf("word_fill_table", { table_index: t0.index, values: [["A", "1"]], start_row: 1 });
    expect(same === null, `Ghi lại đúng giá trị cũ mà vẫn cảnh báo: ${same}`);
    const add = await riskOf("word_fill_table", { table_index: t0.index, values: [["C", "3"]], start_row: t0.rows });
    expect(add === null, `Thêm hàng mới mà vẫn cảnh báo: ${add}`);
    await run("word_fill_table", { table_index: t0.index, values: [["A", "10"], ["B", "2"]], start_row: 1 });
    const t = (await run("word_read_tables", {})).tables.find((x: any) => x.values?.[0]?.[0] === `Mục ${code}`);
    expect(String(t.values[1][1]).includes("10") && String(t.values[2]?.[0]).includes("B"), `Bảng sau khi điền: ${JSON.stringify(t.values)}`);
    return `Bảng ${t0.index}: có hỏi “${warn}”; ghi trùng giá trị cũ và thêm hàng mới: không hỏi; đã sửa 1 ô, thêm 1 hàng`;
  });

  await step("Đọc sửa đổi: gộp theo đoạn, kèm số điều và câu trước/sau; lọc theo tác giả", async () => {
    const s = await need(START);
    const e = await need(END);
    const groups: any[] = [];
    let total = 0;
    let authors: Record<string, number> = {};
    let start = 0;
    let listed = 0;
    for (let guard = 0; guard < 80; guard++) {
      const r = await run("word_list_revisions", { start });
      total = r.total;
      authors = r.authors;
      listed += r.shown;
      groups.push(...r.groups);
      if (r.next_start === undefined) break;
      start = r.next_start;
    }
    expect(listed === total, `Đọc hết các trang được ${listed}/${total} sửa đổi`);
    const mine = groups.filter((g) => g.para > s && g.para < e);
    const g11 = mine.find((g) => String(g.before).includes("0,05%"));
    expect(g11 && g11.clause === "1.1" && String(g11.after).includes("0,03%"), `Nhóm của khoản 1.1: ${JSON.stringify(g11).slice(0, 260)}`);
    const who = Object.keys(authors)[0];
    const some = await run("word_list_revisions", { author: who });
    expect(some.shown > 0 && some.groups.every((g: any) => g.changes.every((c: any) => c.author === who)), "Lọc theo tác giả trả lẫn người khác");
    const none = await run("word_list_revisions", { author: `không-có-ai-${code}` });
    expect(none.shown === 0, `Tác giả không tồn tại mà vẫn có ${none.shown} sửa đổi`);
    return `${total} sửa đổi của ${Object.keys(authors).length} tác giả, đọc đủ ${listed}; khoản 1.1: “…0,05%…” → “…0,03%…”; lọc tác giả đúng`;
  });

  await step("Duyệt sửa đổi kèm đối chiếu chữ: danh sách lệch thì từ chối, khớp thì duyệt", async () => {
    const r = await run("word_list_revisions", {});
    const all = r.groups.flatMap((g: any) => g.changes.map((c: any) => ({ ...c, para: g.para })));
    const target = all.find((c: any) => c.type === "Added" && c.text === "0,03%");
    expect(target, "Không thấy sửa đổi thêm “0,03%” trong trang đầu");
    let refused = "";
    try {
      await run("word_review_revision", { action: "accept", indexes: [target.index], texts: ["chữ khác hẳn"] });
    } catch (e: any) {
      refused = e.message;
    }
    expect(/danh sách đã đổi/.test(refused), `Chữ đối chiếu sai mà không từ chối: ${refused}`);
    let noTexts = "";
    try {
      await run("word_review_revision", { action: "accept", indexes: [target.index] });
    } catch (e: any) {
      noTexts = e.message;
    }
    expect(/Thiếu texts/.test(noTexts), `Duyệt theo số mà không kèm chữ đối chiếu vẫn không từ chối: ${noTexts}`);
    let noAuthor = "";
    try {
      await run("word_review_revision", { action: "reject", all: true, author: `không-có-ai-${code}` });
    } catch (e: any) {
      noAuthor = e.message;
    }
    expect(/Không có sửa đổi nào của/.test(noAuthor), `Duyệt tất cả của một tác giả không tồn tại: ${noAuthor}`);
    const askMany = await riskOf("word_review_revision", { action: "accept", indexes: [0, 1, 2, 3, 4, 5], texts: ["", "", "", "", "", ""] });
    expect(askMany && /6 sửa đổi cùng lúc/.test(askMany), `Duyệt 6 sửa đổi một lúc mà lời cảnh báo là: ${askMany}`);
    const before = (await countAll()).changes;
    await run("word_review_revision", { action: "accept", indexes: [target.index], texts: ["0,03%"] });
    const after = (await countAll()).changes;
    expect(after === before - 1, `Còn ${after} sửa đổi, trước đó ${before}`);
    return `Chữ đối chiếu sai hoặc thiếu → từ chối; tác giả không có → từ chối; duyệt 6 sửa đổi một lúc → có hỏi; đúng → đã chấp nhận 1 (Added “0,03%”), còn ${after}`;
  });

  await step("Đọc văn bản dài: đọc hết theo next_start, không sót đoạn", async () => {
    const i = await need(END);
    await untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const line = `${code} đoạn dài ` + "nội dung kiểm tra đọc văn bản dài ".repeat(11);
      for (let k = 0; k < 90; k++) ps.items[i].insertParagraph(`${line}${k}`, Word.InsertLocation.before);
      await ctx.sync();
    });
    let start = await need(START);
    const from = start;
    let calls = 0;
    let read = 0;
    let total = 0;
    for (;;) {
      const r = await run("word_read_paragraphs", { start });
      calls++;
      total = r.total_paragraphs;
      read += r.paragraphs.length;
      const len = JSON.stringify(r).length;
      expect(len < 30000, `Kết quả ${len} ký tự, vượt mức cắt 30.000`);
      if (r.next_start === undefined || calls > 50) break;
      start = r.next_start;
    }
    expect(read === total - from, `Đã đọc ${read}/${total - from} đoạn`);
    expect(calls >= 2, `Chỉ ${calls} lượt đọc, chưa thử được việc đọc tiếp`);
    return `${total - from} đoạn từ đầu vùng thử, đọc ${calls} lượt theo next_start, không sót đoạn nào`;
  });

  await step("Nhận xét ở đoạn trùng nội dung: báo đúng số đoạn", async () => {
    const dup = `${code} đoạn trùng`;
    const i = await need(END);
    const target = await untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      ps.items[i].insertParagraph(dup, Word.InsertLocation.before);
      ps.items[i].insertParagraph(dup, Word.InsertLocation.before);
      await ctx.sync();
      const hits = ctx.document.body.search(dup, { matchCase: true });
      hits.load("items");
      await ctx.sync();
      hits.items[1].insertComment("Nhận xét ở đoạn trùng thứ hai");
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text");
      await ctx.sync();
      return paras.items.map((x, k) => (x.text === dup ? k : -1)).filter((k) => k >= 0)[1];
    });
    const r = await run("word_list_comments", {});
    const c = r.comments.find((x: any) => String(x.text).includes("đoạn trùng thứ hai"));
    expect(c, "Không thấy nhận xét thử");
    expect(c.paragraph === target, `Công cụ báo đoạn ${c.paragraph}, đúng phải là ${target}`);
    return `Nhận xét neo ở đoạn trùng thứ hai: báo đúng [¶${target}]`;
  });

  await step("Cảnh báo rủi ro: thay nhiều chỗ một lúc, thay vùng chọn lớn", async () => {
    const longLine = `${code} đoạn dài`;
    const many = await riskOf("word_replace_text", { find: longLine, replace: "x", all: true });
    expect(many && /Thay \d+ chỗ/.test(many), `Thay hàng chục chỗ mà lời cảnh báo là: ${many}`);
    const few = await riskOf("word_replace_text", { find: `${code} đoạn trùng`, replace: "x", all: true });
    expect(few === null, `Thay 2 chỗ mà vẫn cảnh báo: ${few}`);
    const one = await riskOf("word_replace_text", { find: longLine, replace: "x" });
    expect(one === null, `Không có all (nhiều nhất thay một chỗ) mà vẫn cảnh báo: ${one}`);
    // Chọn 6 đoạn dài liền nhau rồi hỏi mức rủi ro của việc thay vùng chọn
    await Word.run(async (ctx) => {
      const hits = ctx.document.body.search(longLine, { matchCase: true });
      hits.load("items");
      await ctx.sync();
      expect(hits.items.length > 5, `Chỉ thấy ${hits.items.length} đoạn dài`);
      hits.items[0].paragraphs.getFirst().getRange("Whole").expandTo(hits.items[5].paragraphs.getFirst().getRange("Whole")).select();
      await ctx.sync();
    });
    const big = await riskOf("word_replace_selection", { text: "x" });
    expect(big && /vùng đang chọn/.test(big), `Thay vùng chọn 6 đoạn mà lời cảnh báo là: ${big}`);
    // Con trỏ không chọn gì
    await Word.run(async (ctx) => {
      const hit = ctx.document.body.search(START, { matchCase: true });
      hit.load("items");
      await ctx.sync();
      hit.items[0].getRange("Start").select();
      await ctx.sync();
    });
    const none = await riskOf("word_replace_selection", { text: "x" });
    expect(none === null, `Không chọn gì mà vẫn cảnh báo: ${none}`);
    return `Có hỏi: “${many}” · “${String(big).slice(0, 90)}…” · thay 2 chỗ và chèn tại con trỏ: không hỏi`;
  });

  await step("Xóa đoạn: hộp xác nhận hiện trích nội dung; xóa có theo dõi", async () => {
    const i = await need(`${code} đoạn trùng`);
    const warn = await riskOf("word_delete_paragraphs", { start: i, end: i + 1, expect: `${code} đoạn trùng` });
    expect(warn && warn.includes(`“${code} đoạn trùng”`) && /Xóa 2 đoạn/.test(warn), `Lời cảnh báo xóa: ${warn}`);
    const before = (await countAll()).changes;
    const res = await run("word_delete_paragraphs", { start: i, end: i + 1, expect: `${code} đoạn trùng` });
    const after = (await countAll()).changes;
    expect(res.deleted === 2 && after > before, `Xóa: ${JSON.stringify(res)}; sửa đổi ${before} → ${after}`);
    const r = await run("word_read_paragraphs", { start: i, end: i + 2 });
    expect(r.paragraphs.every((p: any) => p.deleted === true && p.text === ""), `Đoạn đã xóa đọc lại: ${JSON.stringify(r.paragraphs)}`);
    return `Có hỏi: “${warn}”; đã xóa 2 đoạn có theo dõi, đọc lại thấy cờ deleted`;
  });

  await step("Hoàn tác: 3 thao tác ghi và 1 nhận xét rồi hoàn tác → văn bản về như trước", async () => {
    const state = async () => {
      const s = await need(START);
      const e = await need(END);
      const c = await countAll();
      const text = await Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items");
        await ctx.sync();
        const rs = ps.items.slice(s, e + 1).map((p) => (sup("1.7") ? p.getText() : p.getReviewedText("Current")));
        await ctx.sync();
        return rs.map((x) => x.value).join("\n");
      });
      return { ...c, paras: e - s + 1, text };
    };
    const before = await state();
    const mark = log.length;
    const i13 = await need(T.c13);
    await run("word_set_paragraph_text", { index: i13, expect: "Bên A giao hàng", text: T.c13.replace("tại kho", "tại công trình").replace("Bên A", "Bên Bán") });
    const i12 = await need(T.c12);
    await run("word_replace_text", { find: "05 ngày", replace: "07 ngày", paragraph_index: i12, expect: "Bên A xuất hóa đơn" });
    await run("word_insert_clause", { after_index: i12, expect: "Bên A xuất hóa đơn", text: `Khoản thêm để thử hoàn tác ${code}.` });
    await run("word_add_comment", { find: `Thanh toán ${code}`, text: "Nhận xét thử hoàn tác" });
    const mid = await state();
    const entries = log.slice(mark);
    expect(entries.length === 4, `Ghi được ${entries.length} mục hoàn tác (phải là 4)`);
    expect(mid.changes > before.changes && mid.comments === before.comments + 1 && mid.text !== before.text, "Bốn thao tác thử không để lại thay đổi");
    for (const u of entries.reverse()) await u.undo();
    log.length = mark;
    const after = await state();
    expect(after.changes === before.changes, `Sửa đổi chờ duyệt: trước ${before.changes}, sau hoàn tác ${after.changes}`);
    expect(after.comments === before.comments, `Nhận xét: trước ${before.comments}, sau hoàn tác ${after.comments}`);
    expect(after.paras === before.paras, `Số đoạn vùng thử: trước ${before.paras}, sau hoàn tác ${after.paras}`);
    expect(after.text === before.text, "Chữ của vùng thử sau hoàn tác khác trước");
    // Sửa bên trong một khoản AI đã chèn từ trước (đang là sửa đổi chờ duyệt) rồi hoàn tác: chỉ lùi lần sửa, khoản vẫn còn
    const iNew = await need(`Khoản mới chèn ${code}`);
    const mark2 = log.length;
    await run("word_set_paragraph_text", { index: iNew, expect: "Khoản mới chèn", text: `Khoản mới chèn ${code}, đã sửa lại câu chữ.` });
    expect((await inspect(iNew)).current.includes("đã sửa lại câu chữ"), "Sửa trong khoản đang chờ duyệt không ăn");
    for (const x of log.slice(mark2).reverse()) await x.undo();
    log.length = mark2;
    const after2 = await state();
    expect(after2.text === before.text && after2.changes === before.changes, `Hoàn tác lần sửa trong khoản chờ duyệt: sửa đổi ${after2.changes} (trước ${before.changes}), chữ ${after2.text === before.text ? "đúng" : "khác"}`);
    return `Sau 4 thao tác: ${mid.changes - before.changes} sửa đổi mới, 1 nhận xét mới; hoàn tác xong: sửa đổi ${after.changes}, nhận xét ${after.comments}, chữ y như trước (kể cả khoản chèn sát một khoản chờ duyệt, và lần sửa bên trong khoản chờ duyệt)`;
  });

  // ───────────── Đợt 1B: soát và tra ─────────────
  /** Chèn (không theo dõi) các đoạn chữ thường ngay trước mốc cuối của vùng thử */
  const addTemp = (texts: string[]) =>
    untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text");
      await ctx.sync();
      const end = ps.items.find((p) => p.text === END);
      expect(!!end, "Không thấy mốc cuối của vùng thử");
      const made = texts.map((t) => (end as Word.Paragraph).insertParagraph(t, "Before"));
      made.forEach((p) => p.load("isListItem"));
      await ctx.sync();
      made.forEach((p) => {
        if (p.isListItem) p.detachFromList();
        p.styleBuiltIn = "Normal" as any;
      });
      await ctx.sync();
    });
  /** Gỡ hẳn (không theo dõi) các đoạn của vùng thử có chứa một trong các chuỗi */
  const removeTemp = (parts: string[]) =>
    untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text");
      await ctx.sync();
      const s = ps.items.findIndex((p) => p.text === START);
      const e = ps.items.findIndex((p) => p.text === END);
      ps.items.forEach((p, i) => {
        if (i > s && i < e && parts.some((x) => p.text.includes(x))) p.delete();
      });
      await ctx.sync();
    });

  await step("Soát nhất quán bằng mã: nhảy số, dẫn chiếu hỏng, thuật ngữ lệch, chỗ trống, số lệch chữ", async () => {
    const low = code.toLowerCase();
    const P = {
      def: `“Ngày Thử ${code}” là ngày chạy bài kiểm tra soát nhất quán.`,
      bad: `Sau ngày thử ${low} các bên ký biên bản theo Điều 97.8, trả 1.500.000 đồng (Một triệu sáu trăm nghìn đồng) vào tài khoản số ……… SOAT${code}.`,
      jump: `97.3 Khoản thử nhảy số SOAT${code}.`,
    };
    await addTemp([P.def, P.bad, P.jump]);
    try {
      const all: any[] = [];
      let total = 0;
      for (let start = 0, guard = 0; guard < 40; guard++) {
        const r = await run("word_check_consistency", { start });
        total = r.total;
        all.push(...r.issues);
        if (r.next_start === undefined) break;
        start = r.next_start;
      }
      expect(all.length === total, `Đọc được ${all.length}/${total} điểm`);
      const iBad = await need(P.bad.slice(0, 30));
      const iJump = await need(P.jump);
      const has = (kind: string, i: number, re: RegExp) => all.some((x) => x.kind === kind && x.i === i && re.test(x.message));
      expect(has("numbering", iJump, /nhảy/), "Không báo chỗ nhảy số 97.1 → 97.3");
      expect(has("reference", iBad, /97\.8/), "Không báo dẫn chiếu hỏng tới Điều 97.8");
      expect(has("term", iBad, /Ngày Thử/), "Không báo thuật ngữ viết lệch “ngày thử …”");
      expect(has("blank", iBad, /trống/), "Không báo chỗ trống “………”");
      expect(has("amount", iBad, /1\.600\.000/), "Không báo số tiền lệch chữ");
      const only = await run("word_check_consistency", { checks: ["amount"] });
      expect(only.issues.every((x: any) => x.kind === "amount") && only.issues.some((x: any) => x.i === iBad), "Lọc theo loại soát không đúng");
      await removeTemp([P.def.slice(0, 20), "SOAT" + code]);
      const again = await run("word_check_consistency", {});
      const s = await need(START);
      const e = await need(END);
      const left = again.issues.filter((x: any) => x.i > s && x.i < e && /97\.8|Ngày Thử|1\.600\.000|97\.3/.test(x.message + (x.excerpt || "")));
      expect(left.length === 0, `Sau khi gỡ đoạn cài lỗi vẫn còn báo ${left.length} điểm`);
      return `Cài 5 lỗi trong vùng thử: báo đủ 5 (nhảy số, dẫn chiếu, thuật ngữ, chỗ trống, số lệch chữ) trong ${total} điểm của cả văn bản; gỡ đoạn cài lỗi thì hết báo`;
    } finally {
      await removeTemp([P.def.slice(0, 20), "SOAT" + code]);
    }
  });

  await step("Đọc phần ngoài thân văn bản: chú thích, bookmark, ô điền mẫu, đầu và chân trang; sửa đoạn có chú thích không làm mất chú thích", async () => {
    if (!sup("1.5")) throw new Skip("Cần WordApi 1.5 để tạo và đọc chú thích");
    const PX = `Đoạn thử phần ngoài thân NGOAI${code} có thời hạn 10 ngày.`;
    const NOTE = `Chú thích thử ${code}`;
    const BM = `AIDN_${code}`;
    const CC = `Ô thử ${code}`;
    await addTemp([PX]);
    const mark = log.length;
    try {
      await untracked(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text");
        await ctx.sync();
        const p = ps.items.find((x) => x.text.includes("NGOAI" + code)) as Word.Paragraph;
        p.getRange("End").insertFootnote(NOTE);
        const word = p.search("thời hạn", { matchCase: true });
        word.load("items");
        await ctx.sync();
        word.items[0].insertBookmark(BM);
        const days = p.search("10 ngày", { matchCase: true });
        days.load("items");
        await ctx.sync();
        const cc = days.items[0].insertContentControl();
        cc.title = CC;
        cc.tag = "aidn_thu";
        await ctx.sync();
      });
      const i = await need("NGOAI" + code);
      const r = await run("word_read_parts", {});
      const fn = (r.footnotes || []).find((x: any) => String(x.text).includes(NOTE));
      expect(!!fn, "Không đọc được chú thích cuối trang vừa tạo");
      expect(fn.i === i, `Chú thích báo đoạn ${fn.i}, đúng phải là ${i}`);
      expect((r.bookmarks?.items || []).some((x: any) => x.name === BM && x.i === i && /thời hạn/.test(x.text)), "Không đọc được bookmark vừa tạo");
      expect((r.content_controls?.items || []).some((x: any) => x.title === CC && /10 ngày/.test(x.text) && x.i === i), "Không đọc được ô điền mẫu vừa tạo");
      expect(r.headers_footers && !r.headers_footers.unavailable && Array.isArray(r.headers_footers.items), "Không đọc được đầu trang, chân trang");
      const only = await run("word_read_parts", { parts: ["footnotes"] });
      expect(only.footnotes && only.bookmarks === undefined, "Lọc theo phần không đúng");
      // Đọc đoạn: chú thích đi kèm đoạn, chữ của đoạn không lẫn ký tự mốc
      const rd = await run("word_read_paragraphs", { start: i, end: i + 1 });
      const row = rd.paragraphs[0];
      expect(row.text === PX, `Chữ của đoạn có chú thích đọc ra “${row.text}”`);
      expect((row.notes || []).some((n: string) => n.includes(NOTE)), "Kết quả đọc đoạn không kèm chú thích");
      const ov = await run("word_get_overview");
      expect(ov.footnotes >= 1, "Tổng quan không báo số chú thích");
      // Sửa một chữ trong đoạn có chú thích: chú thích phải còn nguyên
      const edited = await run("word_set_paragraph_text", { index: i, expect: "Đoạn thử phần ngoài thân", text: PX.replace("có thời hạn", "có kỳ hạn") });
      expect(edited.ok, `Sửa đoạn có chú thích: ${edited.message}`);
      const after = await run("word_read_parts", { parts: ["footnotes"] });
      expect((after.footnotes || []).some((x: any) => String(x.text).includes(NOTE) && x.i === i), "Sửa đoạn làm mất chú thích cuối trang");
      expect((await inspect(i)).current.replace(/\u0002/g, "") === PX.replace("có thời hạn", "có kỳ hạn"), "Câu sau khi sửa không đúng");
      // Đoạn lúc này vừa có mốc chú thích vừa có chữ đã xóa chờ duyệt ("thời" → "kỳ"): sửa tiếp một chữ khác vẫn phải đúng chỗ
      const again = await run("word_set_paragraph_text", { index: i, expect: "Đoạn thử phần ngoài thân", text: PX.replace("có thời hạn", "có kỳ hạn").replace("Đoạn thử", "Đoạn kiểm") });
      expect(again.ok, `Sửa lần hai trên đoạn có chú thích và sửa đổi chờ duyệt: ${again.message}`);
      expect((await inspect(i)).current.replace(/\u0002/g, "") === PX.replace("có thời hạn", "có kỳ hạn").replace("Đoạn thử", "Đoạn kiểm"), "Câu sau lần sửa thứ hai không đúng");
      const still = await run("word_read_parts", { parts: ["footnotes"] });
      expect((still.footnotes || []).some((x: any) => String(x.text).includes(NOTE) && x.i === i), "Lần sửa thứ hai làm mất chú thích cuối trang");
      for (const u of log.slice(mark).reverse()) await u.undo();
      log.length = mark;
      return `Đọc được chú thích (đúng đoạn ${i}), bookmark, ô điền mẫu, đầu và chân trang (${r.headers_footers.items.length} mục); đọc đoạn kèm chú thích; sửa hai lần trong đoạn có chú thích (lần hai khi đoạn đã có chữ xóa chờ duyệt) thì chú thích còn nguyên`;
    } finally {
      await untracked(async (ctx) => {
        if (sup("1.4")) ctx.document.deleteBookmark(BM);
        const notes = ctx.document.body.footnotes;
        notes.load("items/body/text");
        await ctx.sync();
        notes.items.filter((n) => String(n.body.text).includes(NOTE)).forEach((n) => n.delete());
        await ctx.sync();
      }).catch(() => undefined);
      await removeTemp(["NGOAI" + code]);
    }
  });

  await step("Đợt 8 · Điền mẫu: 5 loại ô điền mẫu và chỗ trống, điền hết trong một lệnh có theo dõi, hỏi khi ghi đè; hoàn tác về y cũ rồi điền lại được", async () => {
    if (!sup("1.6")) throw new Skip("Cần WordApi 1.6 (Hoàn tác theo danh sách sửa đổi)");
    const K = `DIENMAU${code}`;
    const P = {
      block: `Khối ${K}: KHOI`,
      cc: `Ô ${K}: OCHON`,
      untitled: `Tên viết tắt ${K}: VIETTAT`,
      name: `Họ và tên người đại diện ${K}: ……………………`,
      money: `Số tiền ${K}: [●] đồng (bằng chữ: [●] đồng).`,
      // Hai chỗ trống cùng chuỗi trong một đoạn, điền trong CÙNG một lệnh (Chạy thật 18/11 lộ lỗi đọc chữ của đoạn một lần đầu lệnh)
      fee: `Phí ${K}: [●] và [●].`,
    };
    const T = { ten: `Tên ${code}`, hop: `Đánh dấu ${code}`, tt: `Thanh toán ${code}`, loai: `Loại ${code}`, ngay: `Ngày ${code}` };
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
    const pr = (title: string, tag: string, extra: string) => `<w:sdtPr><w:alias w:val="${esc(title)}"/><w:tag w:val="${tag}"/>${extra}</w:sdtPr>`;
    const item = (x: string) => `<w:listItem w:displayText="${esc(x)}" w:value="${esc(x)}"/>`;
    const r = (t: string) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
    const sdts =
      `<w:sdt>${pr(T.ten, "aidn_dien_ten", "<w:showingPlcHdr/><w:text/>")}<w:sdtContent>${r("Bấm để nhập")}</w:sdtContent></w:sdt>${r(" ")}` +
      `<w:sdt>${pr(T.hop, "aidn_dien_hop", '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>')}<w:sdtContent>${r("☐")}</w:sdtContent></w:sdt>${r(" ")}` +
      `<w:sdt>${pr(T.tt, "aidn_dien_tt", `<w:showingPlcHdr/><w:dropDownList>${item("Chuyển khoản")}${item("Tiền mặt")}</w:dropDownList>`)}<w:sdtContent>${r("Chọn một mục")}</w:sdtContent></w:sdt>${r(" ")}` +
      `<w:sdt>${pr(T.loai, "aidn_dien_loai", `<w:showingPlcHdr/><w:comboBox>${item("Căn hộ chung cư")}${item("Đất nền")}</w:comboBox>`)}<w:sdtContent>${r("Chọn hoặc nhập")}</w:sdtContent></w:sdt>${r(" ")}` +
      `<w:sdt>${pr(T.ngay, "aidn_dien_ngay", '<w:showingPlcHdr/><w:date><w:dateFormat w:val="dd/MM/yyyy"/><w:lid w:val="vi-VN"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>')}<w:sdtContent>${r("Chọn ngày")}</w:sdtContent></w:sdt>`;
    const pkg =
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p>' +
      sdts +
      "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
    // Ô chứa cả đoạn (đặt TRƯỚC các chỗ trống: thay nó để lại đoạn cũ chờ xóa, không được làm lệch số đoạn của chỗ trống — soát mã #1)
    const blockPkg = pkg.replace(/<w:body><w:p>[\s\S]*<\/w:p><\/w:body>/, () => `<w:body><w:sdt>${pr(`Khối ${code}`, "aidn_dien_khoi", "<w:showingPlcHdr/>")}<w:sdtContent><w:p>${r("Bấm để nhập đoạn")}</w:p></w:sdtContent></w:sdt></w:body>`);
    // Ô không tên: nhãn lấy từ chữ đứng trước ô (soát mã #2)
    const untitledPkg = pkg.replace(/<w:body><w:p>[\s\S]*<\/w:p><\/w:body>/, () => `<w:body><w:p><w:sdt><w:sdtPr><w:tag w:val="aidn_dien_vt"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>${r("Nhập")}</w:sdtContent></w:sdt></w:p></w:body>`);
    await addTemp([P.block, P.cc, P.untitled, P.name, P.money, P.fee]);
    const mark = log.length;
    try {
      await untracked(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text");
        await ctx.sync();
        const pc = ps.items.find((x) => x.text === P.cc) as Word.Paragraph;
        const w = pc.search("OCHON", { matchCase: true });
        w.load("items");
        await ctx.sync();
        w.items[0].insertOoxml(pkg, "Replace");
        const pu = ps.items.find((x) => x.text === P.untitled) as Word.Paragraph;
        const u = pu.search("VIETTAT", { matchCase: true });
        u.load("items");
        await ctx.sync();
        u.items[0].insertOoxml(untitledPkg, "Replace");
        const pb = ps.items.find((x) => x.text === P.block) as Word.Paragraph;
        pb.getRange("Whole").insertOoxml(blockPkg, "Replace");
        await ctx.sync();
      });
      const iName = await need(`người đại diện ${K}`);
      const iMoney = await need(`Số tiền ${K}`);
      const listMine = async () => {
        const l = await run("word_list_form_fields", {});
        const f = l.fields || [];
        const by = (title: string) => f.find((x: any) => x.kind === "control" && x.label === title);
        const vt = f.find((x: any) => x.kind === "control" && /^Tên viết tắt DIENMAU/.test(x.label));
        const iFee = await need(`Phí ${K}`);
        return { l, name: f.find((x: any) => x.kind === "blank" && x.i === iName), money: f.filter((x: any) => x.kind === "blank" && x.i === iMoney), ten: by(T.ten), hop: by(T.hop), tt: by(T.tt), loai: by(T.loai), ngay: by(T.ngay), khoi: by(`Khối ${code}`), vt, fee: f.filter((x: any) => x.kind === "blank" && x.i === iFee) };
      };
      const vals = (m: any) => [m.ten?.value, m.hop?.value, m.tt?.value, m.loai?.value, m.ngay?.value, m.khoi?.value, m.vt?.value].join(" | ");
      const m0 = await listMine();
      expect(m0.ten && m0.hop && m0.tt && m0.loai && m0.ngay && m0.khoi && m0.vt, `Không thấy đủ 7 ô điền mẫu vừa tạo (có ô chứa đoạn, ô không tên): ${JSON.stringify(m0.l.fields?.filter((x: any) => x.kind === "control"))}`);
      expect(m0.tt.items?.join(",") === "Chuyển khoản,Tiền mặt" && m0.ngay.date_format === "dd/MM/yyyy", `Danh sách mục / định dạng ngày đọc sai: ${JSON.stringify([m0.tt, m0.ngay])}`);
      expect(m0.name?.blank === "……………………" && m0.money.length === 2, `Chỗ trống đọc sai: ${JSON.stringify([m0.name, m0.money])}`);
      const empty0 = vals(m0);
      /** Một vòng: điền mọi chỗ trong một lệnh rồi kiểm */
      const round = async (m: any, v: { ten: string; tt: string; loai: string; ngay: string; ngayDoc: string; name: string; so: string; chu: string; khoi: string; vt: string }) => {
        const fields = [
          { id: m.ten.id, value: v.ten, label: m.ten.label },
          { id: m.hop.id, value: "true", label: m.hop.label },
          { id: m.tt.id, value: v.tt, label: m.tt.label },
          { id: m.loai.id, value: v.loai, label: m.loai.label },
          { id: m.ngay.id, value: v.ngay, label: m.ngay.label },
          { id: m.name.id, value: v.name, label: m.name.label },
          { id: m.money[0].id, value: v.so, label: m.money[0].label },
          { id: m.khoi.id, value: v.khoi, label: m.khoi.label },
          { id: m.vt.id, value: v.vt, label: m.vt.label },
          { id: m.fee[0].id, value: "10 triệu", label: m.fee[0].label },
          { id: m.fee[1].id, value: "20 triệu", label: m.fee[1].label },
        ];
        expect((await riskOf("word_fill_form", { fields })) === null, "Điền chỗ còn trống mà vẫn hỏi");
        const before = log.length;
        const res = await run("word_fill_form", { fields });
        expect(res.filled?.length === fields.length, `Điền ${res.filled?.length}/${fields.length}: ${JSON.stringify(res.skipped)}`);
        expect(log.length === before + 1, `Một lệnh điền phải có đúng một mục Hoàn tác, thấy ${log.length - before}`);
        const m2 = await listMine();
        expect(vals(m2) === [v.ten, "true", v.tt, v.loai, v.ngayDoc, v.khoi, v.vt].join(" | "), `Ô điền mẫu sau khi điền: ${vals(m2)}`);
        // Số đoạn đọc lại SAU khi điền: thay ô chứa đoạn có thể để lại đoạn cũ chờ xóa phía trên
        const jName = await need(`người đại diện ${K}`);
        const jMoney = await need(`Số tiền ${K}`);
        expect((await inspect(jName)).current === `Họ và tên người đại diện ${K}: ${v.name}`, `Đoạn họ tên sau khi điền: “${(await inspect(jName)).current}”`);
        const jFee = await need(`Phí ${K}`);
        expect((await inspect(jFee)).current === `Phí ${K}: 10 triệu và 20 triệu.`, `Hai “[●]” cùng đoạn điền trong một lệnh: “${(await inspect(jFee)).current}”`);
        expect((await inspect(jMoney)).current === `Số tiền ${K}: ${v.so} đồng (bằng chữ: [●] đồng).`, `Đoạn số tiền sau khi điền: “${(await inspect(jMoney)).current}”`);
        // Lần điền thứ hai khi lần đầu còn chờ duyệt trong cùng đoạn (soát mã #3): chỗ “[●]” còn lại vẫn điền được
        const m2b = await listMine();
        const again = (m2b.l.fields || []).filter((x: any) => x.kind === "blank" && x.i === jMoney && x.blank === "[●]");
        expect(again.length === 1, `Sau lần điền đầu phải còn đúng 1 chỗ “[●]”: ${JSON.stringify(again)}`);
        const r2 = await run("word_fill_form", { fields: [{ id: again[0].id, value: v.chu, label: again[0].label }] });
        expect(r2.filled?.length === 1, `Lần điền thứ hai: ${JSON.stringify(r2.skipped)}`);
        expect((await inspect(jMoney)).current === `Số tiền ${K}: ${v.so} đồng (bằng chữ: ${v.chu} đồng).`, `Đoạn số tiền sau lần điền thứ hai: “${(await inspect(jMoney)).current}”`);
        return m2;
      };
      /** Hoàn tác mọi thao tác của bước rồi kiểm về y cũ */
      const undoAll = async () => {
        for (const u of log.slice(mark).reverse()) await u.undo();
        log.length = mark;
        const m3 = await listMine();
        expect(vals(m3) === empty0, `Ô điền mẫu sau hoàn tác: ${vals(m3)} (đúng phải là ${empty0})`);
        expect((await inspect(await need(`người đại diện ${K}`))).current === P.name && (await inspect(await need(`Số tiền ${K}`))).current === P.money && (await inspect(await need(`Phí ${K}`))).current === P.fee, "Chỗ trống sau hoàn tác không về như cũ");
        return m3;
      };
      const filled = await round(m0, { ten: "Công ty TNHH Thử Hưng Thịnh", tt: "Tiền mặt", loai: "Đất nền", ngay: "05/03/2027", ngayDoc: "05/03/2027", name: "Nguyễn Văn Thử", so: "500.000.000", chu: "năm trăm triệu", khoi: "Điều khoản thử điền vào ô chứa đoạn", vt: "HTTN" });
      // Ghi đè ô đã có chữ: hỏi trước. Mã sai, mục ngoài danh sách: bỏ qua có lý do
      const warn = await riskOf("word_fill_form", { fields: [{ id: filled.ten.id, value: "Công ty khác" }] });
      expect(warn && /Ghi đè 1 chỗ/.test(warn), `Ghi đè ô đã có chữ mà không hỏi: ${warn}`);
      const bad = await run("word_fill_form", { fields: [{ id: "b99999.0", value: "x", label: "x" }, { id: filled.tt.id, value: "Trả góp", label: filled.tt.label }, { id: filled.ten.id, value: "y" }] });
      expect(!bad.ok && bad.skipped?.length === 3 && bad.skipped.some((x: any) => /danh sách/.test(x.reason)) && bad.skipped.some((x: any) => /thiếu label/.test(x.reason)), `Mã sai / mục ngoài danh sách / thiếu nhãn phải bị bỏ qua: ${JSON.stringify(bad)}`);
      const m1 = await undoAll();
      // Điền lại sau khi hoàn tác (trước đây Word khóa ô sau khi hoàn tác) rồi hoàn tác lần nữa
      await round(m1, { ten: "Công ty Thử Lần Hai", tt: "Chuyển khoản", loai: "Nhà xưởng", ngay: "2027-12-31", ngayDoc: "31/12/2027", name: "Trần Thị Thử", so: "1.000.000", chu: "một triệu", khoi: "Ô chứa đoạn lần hai", vt: "TTLH" });
      await undoAll();
      return `${m0.l.brief}; 5 loại ô + ô chứa đoạn + ô không tên + “……” + “[●]” điền trong một lệnh, “[●]” còn lại điền ở lệnh thứ hai khi lệnh đầu còn chờ duyệt; ghi đè thì hỏi; mã sai, mục ngoài danh sách, thiếu nhãn bị bỏ qua; hoàn tác về y cũ, điền lại vòng hai rồi hoàn tác vẫn đúng`;
    } finally {
      for (const u of log.slice(mark).reverse()) await u.undo().catch(() => undefined);
      log.length = mark;
      await untracked(async (ctx) => {
        const all = ctx.document.getContentControls({ types: ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"] as any });
        all.load("items/tag");
        await ctx.sync();
        all.items.filter((c) => /^aidn_dien_/.test(c.tag)).forEach((c) => c.delete(false));
        await ctx.sync();
        // Ô chứa đoạn để lại đoạn rỗng / chữ gợi ý sau khi gỡ ô: gỡ các đoạn còn sót của bước
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text");
        await ctx.sync();
        ps.items.filter((p) => p.text.includes(K) || p.text === "Bấm để nhập đoạn").forEach((p) => p.delete());
        await ctx.sync();
      }).catch(() => undefined);
      await removeTemp([K]);
    }
  });

  await step("Thể thức trang: báo đúng lề nào lệch chuẩn (đối chiếu với số đo đọc trực tiếp)", async () => {
    if (!sup("1.3", "WordApiDesktop")) throw new Skip("Cần WordApiDesktop 1.3 để đọc lề và khổ giấy");
    const direct = await Word.run(async (ctx) => {
      const secs = ctx.document.sections;
      secs.load("items");
      await ctx.sync();
      const ps = secs.items.map((x) => {
        x.pageSetup.load("topMargin,bottomMargin,leftMargin,rightMargin");
        return x.pageSetup;
      });
      await ctx.sync();
      return ps.map((p) => ({ top: p.topMargin, bottom: p.bottomMargin, left: p.leftMargin, right: p.rightMargin }));
    });
    // Tính độc lập: lề (điểm) → mm, so với khoảng chuẩn của văn bản hành chính
    const mm = (pt: number) => Math.round((pt * 25.4) / 72 * 10) / 10;
    const STD: Record<string, [number, number]> = { top: [20, 25], bottom: [20, 25], left: [30, 35], right: [15, 20] };
    const want: string[] = [];
    direct.forEach((d, k) => {
      (Object.keys(STD) as (keyof typeof d)[]).forEach((side) => {
        const v = mm(d[side]);
        if (v < STD[side][0] - 0.6 || v > STD[side][1] + 0.6) want.push(`${side}${direct.length > 1 ? "#" + (k + 1) : ""}`);
      });
    });
    const r = await run("word_check_page_format", {});
    const got: string[] = r.issues.filter((x: any) => ["top", "bottom", "left", "right"].includes(x.what)).map((x: any) => `${x.what}${x.section ? "#" + x.section : ""}`);
    expect(got.slice().sort().join(",") === want.slice().sort().join(","), `Công cụ báo lề lệch: [${got.join(", ")}], đo trực tiếp ra: [${want.join(", ")}]`);
    const s0 = r.sections[0];
    expect(Math.abs(s0.left_mm - mm(direct[0].left)) < 0.11 && Math.abs(s0.top_mm - mm(direct[0].top)) < 0.11, "Số đo lề công cụ trả về lệch với số đo trực tiếp");
    // Lấy chính số đo hiện tại làm chuẩn thì không còn lề nào lệch; đòi lề trái lớn hơn 10 mm thì chỉ lề trái lệch
    const self = await run("word_check_page_format", { top_mm: s0.top_mm, bottom_mm: s0.bottom_mm, left_mm: s0.left_mm, right_mm: s0.right_mm });
    expect(!self.issues.some((x: any) => ["top", "bottom", "left", "right"].includes(x.what) && (!x.section || x.section === 1)), "Lấy số đo hiện tại làm chuẩn mà vẫn báo lệch lề");
    const off = await run("word_check_page_format", { top_mm: s0.top_mm, bottom_mm: s0.bottom_mm, left_mm: s0.left_mm + 10, right_mm: s0.right_mm });
    const offSides = off.issues.filter((x: any) => ["top", "bottom", "left", "right"].includes(x.what) && (!x.section || x.section === 1)).map((x: any) => x.what);
    expect(offSides.join(",") === "left", `Chuẩn đòi lề trái ${s0.left_mm + 10} mm: công cụ báo lệch [${offSides.join(", ")}], đúng phải là [left]`);
    const lech = r.issues.filter((x: any) => ["top", "bottom", "left", "right", "paper", "orientation"].includes(x.what)).map((x: any) => `${x.label} ${x.actual} (chuẩn ${x.expected})`);
    return `Giấy ${s0.paper} ${s0.orientation}; lề trên ${s0.top_mm}, dưới ${s0.bottom_mm}, trái ${s0.left_mm}, phải ${s0.right_mm} mm → ${lech.length ? "lệch chuẩn: " + lech.join("; ") : "bốn lề và khổ giấy đúng chuẩn"}; phông: ${r.fonts
      .slice(0, 3)
      .map((f: any) => `${f.name} ${f.size ?? "?"} (${f.share})`)
      .join(", ")}`;
  });

  /** Đoạn neo riêng của hai bước thể thức (chữ thường, kiểu Normal), đặt ngay sau mốc đầu vùng thử; gỡ khi xong */
  const NEO = `Đoạn neo thể thức ${code} của bước kiểm tra`;
  const addNeo = async () => {
    const s = await need(START);
    await untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const np = ps.items[s].insertParagraph(NEO, "After");
      np.styleBuiltIn = "Normal" as any;
      await ctx.sync();
    });
    return need(NEO);
  };
  const dropNeo = async () => {
    const i = await at(NEO);
    if (i < 0) return;
    await untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      ps.items[i].delete();
      await ctx.sync();
    });
  };

  await step("Văn bản hành chính: dựng công văn đúng thể thức trong vùng thử, đọc lại cỡ chữ, kiểu chữ, đường kẻ; hoàn tác thì hết", async () => {
    const zone = async () => {
      const s = await need(START);
      const e = await need(END);
      return Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text,items/alignment,items/firstLineIndent,items/font/name,items/font/size,items/font/bold,items/font/italic,items/tableNestingLevel");
        const ts = ctx.document.body.tables;
        ts.load("items");
        await ctx.sync();
        return { tables: ts.items.length, paras: ps.items.slice(s, e + 1).map((p) => ({ text: p.text, align: String(p.alignment), indent: p.firstLineIndent, font: p.font.name, size: p.font.size, bold: p.font.bold, italic: p.font.italic, inTable: p.tableNestingLevel > 0 })) };
      });
    };
    const anchor = await addNeo();
    const before = await zone();
    const c0 = await countAll();
    const mark = log.length;
    const body = [`Công ty nhận được văn bản đề nghị cung cấp hồ sơ ${code} của Quý Sở.`, "## 1. Nội dung báo cáo", `Tiến độ thực hiện ${code} đạt 60% khối lượng.`, "Trân trọng."].join("\n");
    const r = await run("word_create_admin_document", {
      doc_type: "cong_van",
      org: `Công ty cổ phần thử ${code}`,
      parent_org: "Tổng công ty thử",
      symbol: "THU-PC",
      place: "Hà Nội",
      subject: `báo cáo tiến độ ${code}`,
      to: ["Sở Xây dựng thử"],
      body,
      signer_title: "Giám đốc",
      page_setup: false,
      position: "after_paragraph",
      paragraph_index: anchor,
      expect: NEO.slice(0, 30),
    });
    expect(r.ok && Array.isArray(r.blanks), "Công cụ không trả danh sách chỗ trống");
    expect(r.blanks.some((b: string) => b.includes("số văn bản")) && r.blanks.some((b: string) => b.includes("họ tên người ký")), `Thiếu chỗ trống phải báo: ${r.blanks.join("; ")}`);
    const mid = await zone();
    expect(mid.tables === before.tables + 2, `Số bảng tăng ${mid.tables - before.tables} (phải là 2: phần đầu và phần ký)`);
    const find = (part: string) => mid.paras.find((p) => p.text.includes(part));
    const qh = find("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    const tn = find("Độc lập - Tự do - Hạnh phúc");
    const so = find("/THU-PC");
    const vv = find(`V/v báo cáo tiến độ ${code}`);
    const kg = find("Kính gửi: Sở Xây dựng thử");
    const nd = find(`Tiến độ thực hiện ${code}`);
    const ky = find("GIÁM ĐỐC");
    const nn = find("Nơi nhận:");
    const luu = find("- Lưu: VT");
    const org = find(`CÔNG TY CỔ PHẦN THỬ ${code}`);
    for (const [name, p] of [["quốc hiệu", qh], ["tiêu ngữ", tn], ["số và ký hiệu", so], ["V/v", vv], ["Kính gửi", kg], ["nội dung", nd], ["chức vụ ký", ky], ["Nơi nhận", nn], ["dòng Lưu", luu], ["tên cơ quan", org]] as [string, any][])
      expect(!!p, `Không thấy thành phần: ${name}`);
    expect(qh!.inTable && qh!.bold && qh!.size === 13 && qh!.align === "Centered" && qh!.font === "Times New Roman", `Quốc hiệu sai: ${JSON.stringify(qh)}`);
    expect(tn!.bold && tn!.size === 14, `Tiêu ngữ sai cỡ hoặc không đậm: ${tn!.size}`);
    expect(org!.bold && org!.size === 13, "Tên cơ quan ban hành phải đậm cỡ 13");
    expect(vv!.size === 13 && !vv!.bold, "V/v phải cỡ 12–13, đứng");
    expect(!nd!.inTable && nd!.size === 14 && nd!.align === "Justified" && Math.abs(nd!.indent - 28.35) < 0.5 && !nd!.bold, `Đoạn nội dung sai: ${JSON.stringify(nd)}`);
    expect(/\.\/\.$/.test(find("Trân trọng")!.text), "Thiếu dấu kết thúc ./.");
    expect(ky!.bold && ky!.inTable, "Chức vụ ký phải đậm, nằm trong bảng phần ký");
    expect(nn!.size === 12 && nn!.bold && nn!.italic && luu!.size === 11, "Nơi nhận phải cỡ 12 đậm nghiêng, danh sách cỡ 11");
    // Đường kẻ: đếm đoạn rỗng có viền dưới trong vùng thử
    let rules = -1;
    if (sup("1.3", "WordApiDesktop")) {
      const s = await need(START);
      const e = await need(END);
      rules = await Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text");
        await ctx.sync();
        const empties = ps.items.slice(s, e + 1).filter((p) => !p.text.trim());
        empties.forEach((p) => p.borders.load("items/lineStyle"));
        await ctx.sync();
        return empties.filter((p) => p.borders.items[2] && String(p.borders.items[2].lineStyle) === "Single").length;
      });
      expect(rules === 2, `Số đường kẻ (viền dưới) là ${rules}, phải là 2: dưới tên cơ quan và dưới tiêu ngữ`);
    }
    const c1 = await countAll();
    expect(c1.changes > c0.changes, "Dựng văn bản mà không sinh sửa đổi được theo dõi");
    const entries = log.slice(mark);
    expect(entries.length === 1, `Ghi ${entries.length} mục hoàn tác (phải là 1)`);
    for (const u of entries.reverse()) await u.undo();
    log.length = mark;
    const after = await zone();
    const c2 = await countAll();
    expect(after.tables === before.tables, `Sau hoàn tác còn ${after.tables - before.tables} bảng thừa`);
    expect(after.paras.map((p) => p.text).join("\n") === before.paras.map((p) => p.text).join("\n"), "Chữ của vùng thử sau hoàn tác khác trước");
    expect(c2.changes === c0.changes, `Sửa đổi chờ duyệt: trước ${c0.changes}, sau hoàn tác ${c2.changes}`);
    // Tờ trình không có cơ quan chủ quản: ba đường kẻ (tên cơ quan, tiêu ngữ, trích yếu), có tên loại và trích yếu
    let rules2 = -1;
    let where2 = "";
    const mark3 = log.length;
    await run("word_create_admin_document", {
      doc_type: "to_trinh",
      org: `Công ty thử ${code}`,
      symbol: "TTr-THU",
      subject: `Phê duyệt phương án thử ${code}`,
      to: ["Hội đồng quản trị", "Ban Kiểm soát"],
      body: `## I. Sự cần thiết\nNội dung trình ${code}.`,
      signer_authority: "KT. Giám đốc",
      signer_title: "Phó Giám đốc",
      signer_name: "TS. Nguyễn Văn Thử",
      page_setup: false,
      position: "after_paragraph",
      paragraph_index: await need(NEO),
      expect: NEO.slice(0, 30),
    });
    const tt = await zone();
    const has = (t: string) => tt.paras.find((p) => p.text === t);
    expect(!!has("TỜ TRÌNH") && has("TỜ TRÌNH")!.bold && has("TỜ TRÌNH")!.align === "Centered", "Thiếu tên loại TỜ TRÌNH in đậm canh giữa");
    expect(!!has(`Phê duyệt phương án thử ${code}`)?.bold, "Trích yếu phải đậm");
    expect(!!has("- Hội đồng quản trị;") && !!has("- Ban Kiểm soát.") && !!has("Kính gửi:"), "Kính gửi nhiều nơi: mỗi nơi một dòng, chấm phẩy rồi chấm");
    expect(!!has("KT. GIÁM ĐỐC") && !!has("PHÓ GIÁM ĐỐC") && !!has("Nguyễn Văn Thử"), "Phần ký: quyền hạn, chức vụ in hoa; họ tên bỏ học vị");
    expect(tt.tables === before.tables + 2, `Tờ trình: ${tt.tables - before.tables} bảng (phải là 2)`);
    // Thứ tự các thành phần đúng như bản vẽ (đã có lần bảng chèn sau đoạn rỗng làm đảo thứ tự)
    const order = ["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "TỜ TRÌNH", `Phê duyệt phương án thử ${code}`, "Kính gửi:", "- Hội đồng quản trị;", "I. Sự cần thiết", "Nơi nhận:", "Nguyễn Văn Thử"].map((t) => tt.paras.findIndex((p) => p.text === t));
    expect(order.every((v, k) => v >= 0 && (k === 0 || v > order[k - 1])), `Thứ tự thành phần sai: ${order.join(", ")}`);
    if (sup("1.3", "WordApiDesktop")) {
      const s = await need(START);
      const e = await need(END);
      rules2 = await Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text");
        await ctx.sync();
        const empties = ps.items.slice(s, e + 1).filter((p) => !p.text.trim());
        empties.forEach((p) => p.borders.load("items/lineStyle"));
        await ctx.sync();
        const hit = empties.filter((p) => p.borders.items[2] && String(p.borders.items[2].lineStyle) === "Single");
        where2 = ps.items.map((p, k) => (hit.includes(p) ? `${k - s}:${ps.items[k - 1]?.text.slice(0, 12)}|${ps.items[k + 1]?.text.slice(0, 12)}` : "")).filter(Boolean).join(" ; ");
        return hit.length;
      });
      expect(rules2 === 3, `Tờ trình: ${rules2} đường kẻ, phải là 3 (tên cơ quan, tiêu ngữ, trích yếu) — ở: ${where2}`);
    }
    for (const u of log.slice(mark3).reverse()) await u.undo();
    log.length = mark3;
    const after2 = await zone();
    const left = after2.paras.map((p) => p.text).filter((t, k) => t !== before.paras[k]?.text).slice(0, 4);
    expect(after2.tables === before.tables && after2.paras.length === before.paras.length, `Hoàn tác tờ trình chưa sạch: bảng ${before.tables} → ${after2.tables}, đoạn ${before.paras.length} → ${after2.paras.length}; còn lại: ${left.map((t) => `“${t.slice(0, 30)}”`).join(", ")}`);
    await dropNeo();
    return `Tờ trình: 2 bảng, ${rules2 < 0 ? "không đo được" : rules2} đường kẻ, hoàn tác sạch. Công văn: 2 bảng không viền, ${mid.paras.length - before.paras.length} đoạn, ${rules < 0 ? "không đo được đường kẻ" : rules + " đường kẻ"}, báo ${r.blanks.length} chỗ trống; hoàn tác sạch`;
  });

  await step("Hợp đồng: dựng hợp đồng có tên đơn vị (bảng hai cột), không thừa dòng tên, không Nơi nhận; đọc thấy ô bảng; không xóa vắt mép bảng; hoàn tác sạch", async () => {
    const snap = async () => {
      const s = await need(START);
      const e = await need(END);
      return Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text,items/font/bold,items/font/size,items/alignment,items/tableNestingLevel,items/spaceAfter,items/spaceBefore");
        const ts = ctx.document.body.tables;
        ts.load("items");
        await ctx.sync();
        return { s, tables: ts.items.length, paras: ps.items.slice(s, e + 1).map((p) => ({ text: p.text, bold: p.font.bold, size: p.font.size, sa: p.spaceAfter, sb: p.spaceBefore, align: String(p.alignment), inTable: p.tableNestingLevel > 0 })) };
      });
    };
    const anchor = await addNeo();
    const before = await snap();
    const mark = log.length;
    const TEN = `HỢP ĐỒNG THI CÔNG THỬ ${code}`;
    // body chép lại phần đầu và phần ký như model đã làm (Chạy thật 23/11/2026): công cụ phải tự bỏ
    const body = [TEN, "Số: …/2026/HĐ-THU", "_Căn cứ nhu cầu của hai bên._", "Hôm nay, ngày … tháng … năm 2026, chúng tôi gồm:", `## Điều 1. Nội dung công việc ${code}`, "Hợp đồng lập thành 04 bản.", "ĐẠI DIỆN BÊN A", "ĐẠI DIỆN BÊN B"].join("\n");
    const r = await run("word_create_admin_document", {
      doc_type: "hop_dong",
      org: `Công ty thử ${code}`,
      title: TEN,
      symbol: "HĐ-THU",
      body,
      left_signer_title: "Giám đốc",
      signer_title: "Giám đốc",
      page_setup: false,
      position: "after_paragraph",
      paragraph_index: anchor,
      expect: NEO.slice(0, 30),
    });
    expect(r.ok && r.blanks.includes("số hợp đồng") && (r.notes || []).some((n: string) => /đã bỏ/.test(n)), `Kết quả công cụ thiếu chỗ trống hoặc ghi chú bỏ dòng lặp: ${JSON.stringify(r).slice(0, 300)}`);
    const mid = await snap();
    expect(mid.tables === before.tables + 2, `Số bảng tăng ${mid.tables - before.tables} (phải là 2: phần đầu và phần ký)`);
    const texts = mid.paras.map((p) => p.text);
    expect(texts.filter((t) => t === TEN).length === 1, `Tên hợp đồng xuất hiện ${texts.filter((t) => t === TEN).length} lần (phải là 1)`);
    expect(!texts.some((t) => /^Nơi nhận|Kính gửi|\.\/\.$/.test(t)), "Hợp đồng không được có Nơi nhận, Kính gửi, dấu ./.");
    const org = mid.paras.find((p) => p.text === `CÔNG TY THỬ ${code}`);
    const qh = mid.paras.find((p) => p.text === "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    const ten = mid.paras.find((p) => p.text === TEN);
    expect(!!org?.inTable && !!qh?.inTable && !!ten && !ten.inTable && ten.bold && ten.align === "Centered", "Tên đơn vị, quốc hiệu phải trong bảng; tên hợp đồng đậm canh giữa ngoài bảng");
    expect(texts.includes("ĐẠI DIỆN BÊN A") && texts.includes("ĐẠI DIỆN BÊN B") && texts.filter((t) => t === "ĐẠI DIỆN BÊN A").length === 1, "Phần ký phải có đúng một ĐẠI DIỆN BÊN A và một BÊN B");
    // Đọc lại: đoạn trong ô có cell, tên đơn vị cột 1, quốc hiệu cột 2 cùng hàng
    const orgIdx = before.s + texts.indexOf(`CÔNG TY THỬ ${code}`);
    const rd = await run("word_read_paragraphs", { start: orgIdx, end: orgIdx + 12 });
    const cellOf = (t: string) => rd.paragraphs.find((p: any) => p.text === t)?.cell as string | undefined;
    const c1 = cellOf(`CÔNG TY THỬ ${code}`);
    const c2 = cellOf("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    expect(!!rd.table_note && /hàng 1, cột 1$/.test(c1 || "") && /hàng 1, cột 2$/.test(c2 || "") && c1!.split(",")[0] === c2!.split(",")[0], `Đọc đoạn không báo đúng ô bảng: ${c1} / ${c2}`);
    expect(!cellOf(TEN), "Tên hợp đồng nằm ngoài bảng mà vẫn bị gắn cell");
    // Xóa vắt mép bảng (đoạn cuối phần đầu + tên hợp đồng) phải bị từ chối, văn bản không đổi
    // Từ dòng "Số: …" (ô trái, hàng 2 của phần đầu) tới tên hợp đồng: vừa trong ô vừa ngoài bảng
    const tenIdx = before.s + texts.indexOf(TEN);
    const soK = texts.findIndex((t) => t.startsWith("Số:") && t.includes("HĐ-THU"));
    expect(soK >= 0 && mid.paras[soK].inTable, "Dòng Số của hợp đồng phải nằm trong bảng phần đầu");
    let refused = "";
    try {
      await run("word_delete_paragraphs", { start: before.s + soK, end: tenIdx, expect: texts[soK].slice(0, 30) });
    } catch (e: any) {
      refused = String(e?.message || e);
    }
    expect(/trong ô bảng vừa có đoạn ngoài bảng/.test(refused), `Xóa vắt mép bảng không bị chặn: ${refused || "đã xóa"}`);
    const same = await snap();
    expect(same.paras.length === mid.paras.length && same.tables === mid.tables, "Lệnh xóa bị chặn mà văn bản vẫn đổi");
    // Xóa trong bảng nhưng vắt từ ô trái (tên đơn vị) sang ô phải (quốc hiệu) cũng phải bị từ chối
    const orgK = texts.indexOf(`CÔNG TY THỬ ${code}`);
    const qhK = texts.indexOf("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    let refused2 = "";
    try {
      await run("word_delete_paragraphs", { start: before.s + orgK, end: before.s + qhK, expect: texts[orgK].slice(0, 30) });
    } catch (e: any) {
      refused2 = String(e?.message || e);
    }
    expect(/nhiều ô bảng/.test(refused2), `Xóa vắt hai ô của phần đầu không bị chặn: ${refused2 || "đã xóa"}`);
    for (const u of log.slice(mark).reverse()) await u.undo();
    log.length = mark;
    const after = await snap();
    expect(after.tables === before.tables && after.paras.map((p) => p.text).join("\n") === before.paras.map((p) => p.text).join("\n"), `Hoàn tác hợp đồng chưa sạch: bảng ${before.tables} → ${after.tables}, đoạn ${before.paras.length} → ${after.paras.length}; khác: ${after.paras.map((p, k) => (p.text !== before.paras[k]?.text ? `${k}:“${p.text.slice(0, 30)}”(${p.inTable ? "ô" : "ngoài"} cỡ ${p.size} đậm ${p.bold} ${p.align} trước ${p.sb} sau ${p.sa})` : "")).filter(Boolean).slice(0, 4).join(", ")}; lúc dựng: ${mid.paras.slice(0, 14).map((p, k) => `${k}:“${p.text.slice(0, 14)}”${p.inTable ? "ô" : ""}${p.size}`).join(" ")}; trước: ${before.paras.slice(0, 4).map((p) => `“${p.text.slice(0, 14)}”${p.size}`).join(" ")}; cuối: ${mid.paras.slice(texts.indexOf("ĐẠI DIỆN BÊN A") - 2, texts.indexOf("ĐẠI DIỆN BÊN A") + 19).map((p) => `“${p.text.slice(0, 12)}”${p.inTable ? "ô" : ""}${p.size}`).join(" ")}`);
    await dropNeo();
    // Chèn hợp đồng ngay trước một bảng có sẵn của người dùng: đoạn ngăn bảng ký với bảng đó phải giữ (hai bảng không dính nhau)
    const base = await snap();
    const k = base.paras.findIndex((p, i) => !p.inTable && p.text.trim().length >= 5 && base.paras[i + 1]?.inTable);
    let nearTable = "không có bảng trong vùng thử: bỏ qua ca chèn sát bảng";
    if (k >= 0) {
      const rowsOf = () =>
        Word.run(async (ctx) => {
          const ts = ctx.document.body.tables;
          ts.load("items/rowCount");
          await ctx.sync();
          return ts.items.map((t) => t.rowCount).join(",");
        });
      const rows0 = await rowsOf();
      const mark2 = log.length;
      await run("word_create_admin_document", { doc_type: "hop_dong", org: `Công ty sát bảng ${code}`, title: "Hợp đồng thử sát bảng", body: `## Điều 1. Nội dung ${code}`, page_setup: false, position: "after_paragraph", paragraph_index: base.s + k, expect: base.paras[k].text.slice(0, 30) });
      const near = await snap();
      const sigEnd = near.paras.map((p) => p.text).lastIndexOf("…");
      expect(near.tables === base.tables + 2, `Chèn sát bảng: số bảng tăng ${near.tables - base.tables} (phải là 2, không dính vào bảng có sẵn)`);
      expect(sigEnd >= 0 && !near.paras[sigEnd + 1]?.inTable, "Giữa bảng ký và bảng có sẵn phải còn đoạn ngăn cách");
      for (const u of log.slice(mark2).reverse()) await u.undo();
      log.length = mark2;
      const back = await snap();
      expect(back.tables === base.tables && back.paras.map((p) => p.text).join("\n") === base.paras.map((p) => p.text).join("\n") && (await rowsOf()) === rows0, `Hoàn tác ca chèn sát bảng chưa sạch: bảng ${base.tables} → ${back.tables}, đoạn ${base.paras.length} → ${back.paras.length}`);
      nearTable = "chèn sát bảng có sẵn: bảng không dính, hoàn tác sạch";
    }
    return `2 bảng, tên hợp đồng 1 dòng, ${r.blanks.length} chỗ trống, ô bảng: ${c1} / ${c2}; xóa vắt mép bảng và vắt hai ô bị chặn; hoàn tác sạch; ${nearTable}`;
  });

  await step("Đặt thể thức trang: đổi lề rồi hoàn tác về đúng lề cũ; định dạng thụt đầu dòng, cách đoạn, giãn dòng", async () => {
    if (!sup("1.3", "WordApiDesktop")) throw new Skip("Cần WordApiDesktop 1.3 để đặt lề và khổ giấy");
    const read = () =>
      Word.run(async (ctx) => {
        const p = ctx.document.sections.getFirst().pageSetup;
        p.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight");
        await ctx.sync();
        return { top: p.topMargin, bottom: p.bottomMargin, left: p.leftMargin, right: p.rightMargin, w: p.pageWidth, h: p.pageHeight };
      });
    const mm = (pt: number) => Math.round(((pt * 25.4) / 72) * 10) / 10;
    const b = await read();
    const mark = log.length;
    const want = { top_mm: Math.round(mm(b.top)) + 3, bottom_mm: Math.round(mm(b.bottom)) + 2, left_mm: Math.round(mm(b.left)) + 4, right_mm: Math.round(mm(b.right)) + 1, a4: false };
    let msg = "";
    try {
      const r = await run("word_set_page_format", want);
      expect(r.ok && r.changed === true, "Công cụ không báo đã đổi lề");
      const m = await read();
      expect(Math.abs(mm(m.top) - want.top_mm) < 0.2 && Math.abs(mm(m.left) - want.left_mm) < 0.2 && Math.abs(mm(m.right) - want.right_mm) < 0.2 && Math.abs(mm(m.bottom) - want.bottom_mm) < 0.2, `Lề sau khi đặt: ${JSON.stringify(m)}`);
      expect(Math.abs(m.w - b.w) < 0.5 && Math.abs(m.h - b.h) < 0.5, "a4=false mà khổ giấy bị đổi");
      const same = await run("word_set_page_format", want);
      expect(same.changed === false, "Đặt lại đúng lề đang có mà vẫn báo đã đổi");
      msg = `lề trái ${mm(b.left)} → ${mm(m.left)} mm`;
    } finally {
      for (const u of log.slice(mark).reverse()) await u.undo();
      log.length = mark;
    }
    const a = await read();
    expect(Math.abs(a.top - b.top) < 0.3 && Math.abs(a.left - b.left) < 0.3 && Math.abs(a.right - b.right) < 0.3 && Math.abs(a.bottom - b.bottom) < 0.3, `Hoàn tác không trả đúng lề cũ: ${JSON.stringify(a)} so với ${JSON.stringify(b)}`);
    // Định dạng đoạn theo chuẩn văn bản hành chính
    const i = await addNeo();
    const mark2 = log.length;
    await run("word_format_text", { paragraph_index: i, expect: NEO.slice(0, 30), first_line_indent_mm: 10, space_after_pt: 6, line_spacing: 1.5 });
    const f = await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/firstLineIndent,items/spaceAfter,items/lineSpacing");
      await ctx.sync();
      const p = ps.items[i];
      return { indent: p.firstLineIndent, after: p.spaceAfter, line: p.lineSpacing };
    });
    expect(Math.abs(f.indent - 28.35) < 0.5 && Math.abs(f.after - 6) < 0.1 && Math.abs(f.line - 18) < 0.1, `Định dạng đoạn đọc lại: ${JSON.stringify(f)}`);
    for (const u of log.slice(mark2).reverse()) await u.undo();
    log.length = mark2;
    await dropNeo();
    return `${msg}, hoàn tác về ${mm(a.left)} mm; thụt đầu dòng 10 mm, cách đoạn 6 pt, giãn dòng 1,5`;
  });

  await step("So sánh phiên bản: liệt kê đúng 3 chỗ khác kèm số điều; văn bản đang mở không đổi", async () => {
    if (!sup("1.3", "WordApiHiddenDocument")) throw new Skip("Cần WordApiHiddenDocument 1.3 để đọc tệp .docx khác");
    const KEEP = `Đoạn so sánh giữ nguyên SS${code}.`;
    const EDIT = `Đoạn so sánh sẽ được sửa SS${code}, ký ngày 01/01/2026.`;
    const DROP = `Điều khoản phạt chậm giao hàng SS${code} bị bỏ ở phiên bản khác.`;
    const ADDED = `Phụ lục bảng giá ${code} mới thêm vào.`;
    const name = `aidn-thu-so-sanh-${code}.docx`;
    const { registerAttachedFile, unregisterAttachedFile } = await import("../core/attachedFiles");
    await addTemp([KEEP, EDIT, DROP]);
    try {
      const s = await need(START);
      const e = await need(END);
      // Chữ hiện tại của vùng thử → dựng "phiên bản khác": sửa một đoạn, bỏ một đoạn, thêm một đoạn
      const cur: { i: number; text: string }[] = [];
      for (let start = s, guard = 0; start <= e && guard < 60; guard++) {
        const r = await run("word_read_paragraphs", { start, end: e + 1 });
        r.paragraphs.filter((p: any) => p.i <= e && !p.deleted && String(p.text).trim()).forEach((p: any) => cur.push({ i: p.i, text: p.text }));
        if (r.next_start === undefined || r.next_start > e) break;
        start = r.next_start;
      }
      const kEdit = cur.findIndex((p) => p.text === EDIT);
      const kDrop = cur.findIndex((p) => p.text === DROP);
      expect(kEdit >= 0 && kDrop >= 0 && cur[cur.length - 1].text === END, "Không đọc lại được các đoạn mẫu của bước so sánh");
      const other = cur.map((p) => p.text);
      other[kEdit] = other[kEdit].replace("01/01/2026", "02/02/2026");
      other.splice(other.length - 1, 0, ADDED);
      other.splice(kDrop, 1);
      const xml = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      const JSZip = (await import("jszip")).default;
      const zip = new JSZip();
      zip.file(
        "[Content_Types].xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
      );
      zip.file(
        "_rels/.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
      );
      zip.file(
        "word/document.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${other
          .map((t) => `<w:p><w:r><w:t xml:space="preserve">${xml(t)}</w:t></w:r></w:p>`)
          .join("")}</w:body></w:document>`
      );
      registerAttachedFile(name, await zip.generateAsync({ type: "base64" }));
      const before = await countAll();
      const r = await run("word_compare_versions", { file: name, from: s, to: e });
      const after = await countAll();
      expect(r.open_document_modified === false && after.changes === before.changes && after.comments === before.comments, "So sánh làm đổi văn bản đang mở");
      expect((await need(END)) === e && (await need(START)) === s, "So sánh làm đổi số đoạn của văn bản đang mở");
      const kinds = r.diffs
        .map((d: any) => d.type)
        .sort()
        .join(",");
      expect(
        r.differences === 3 && kinds === "changed,only_current,only_other",
        `Phải ra đúng 3 chỗ khác (sửa, chỉ ở bản đang mở, chỉ ở phiên bản khác); công cụ ra ${r.differences}: ${r.diffs.map((d: any) => `${d.type}@${d.i}`).join(", ")}`
      );
      const ch = r.diffs.find((d: any) => d.type === "changed");
      expect(ch.i === cur[kEdit].i && (ch.changes || []).some((c: string) => c.includes("01/01/2026") && c.includes("02/02/2026")), "Chỗ sửa không chỉ đúng đoạn hoặc đúng chữ đổi");
      const gone = r.diffs.find((d: any) => d.type === "only_current");
      expect(gone.i === cur[kDrop].i && String(gone.current).includes("phạt chậm giao hàng"), `Đoạn chỉ có ở bản đang mở: báo [¶${gone.i}], đúng phải là [¶${cur[kDrop].i}]`);
      const added = r.diffs.find((d: any) => d.type === "only_other");
      expect(String(added.other).includes(ADDED), "Không thấy đoạn chỉ có ở phiên bản khác");
      expect(typeof ch.clause === "string" && ch.clause.length > 0, "Chỗ khác không kèm số điều");
      let miss = "";
      try {
        await run("word_compare_versions", { file: "khong-co-tep-nay.docx" });
      } catch (err: any) {
        miss = err.message;
      }
      expect(/Không thấy tệp đính kèm/.test(miss), "Gọi với tên tệp không có mà không báo lỗi rõ");
      return `So ${r.paragraphs_current} đoạn của vùng thử với tệp dựng sẵn: đúng 3 chỗ khác, đúng loại và đúng chỗ — sửa ở [¶${ch.i}] (${ch.clause}): ${ch.changes[0]}; bỏ [¶${gone.i}]; thêm một đoạn sau [¶${added.i}]. Văn bản đang mở không đổi (${after.changes} sửa đổi, ${after.comments} nhận xét như trước)`;
    } finally {
      // Tệp thử không được nằm lại trong kho tệp đính kèm của phiên (người dùng hỏi "so với bản đính kèm" sẽ so nhầm)
      unregisterAttachedFile(name);
      await removeTemp(["SS" + code]);
    }
  });

  await step("Chuyển tới đoạn", async () => (await run("word_go_to", { paragraph_index: await need(T.art1) })).message);

  await step("Dọn vùng thử; phần văn bản ngoài vùng thử không đổi", async () => {
    await untracked(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text");
      await ctx.sync();
      const s = ps.items.findIndex((p) => p.text === START);
      const e = ps.items.findIndex((p) => p.text === END);
      expect(s >= 0 && e > s, "Không tìm thấy hai đoạn mốc của vùng thử");
      // Trọn các đoạn của vùng thử (kể cả dấu hết đoạn của đoạn mốc cuối): không nhập đoạn nào của người dùng vào
      const rng = ps.items[s].getRange("Whole").expandTo(ps.items[e].getRange("Whole"));
      const cs = rng.getComments();
      cs.load("items");
      await ctx.sync();
      cs.items.forEach((c) => c.delete());
      rng.delete();
      await ctx.sync();
    });
    const now = await snapshotOutside();
    const left = await at(code);
    expect(left < 0, "Vẫn còn nội dung thử");
    expect(now.paras === outside.paras, `Số đoạn: trước ${outside.paras}, sau ${now.paras}`);
    expect(now.text === outside.text, "Chữ của văn bản ngoài vùng thử đã đổi");
    expect(now.styles === outside.styles, "Kiểu đoạn của văn bản ngoài vùng thử đã đổi");
    expect(now.changes === outside.changes, `Sửa đổi chờ duyệt: trước ${outside.changes}, sau ${now.changes}`);
    expect(now.comments === outside.comments, `Nhận xét: trước ${outside.comments}, sau ${now.comments}`);
    expect(now.mode === outside.mode, `Chế độ theo dõi: trước ${outside.mode}, sau ${now.mode}`);
    return `Đã xóa vùng thử; ngoài vùng thử y nguyên: ${now.paras} đoạn, ${now.changes} sửa đổi, ${now.comments} nhận xét, chế độ ${now.mode}`;
  });
  void tools;
}
