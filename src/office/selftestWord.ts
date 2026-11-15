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
