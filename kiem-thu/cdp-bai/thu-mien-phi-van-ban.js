// Một bài thử "model miễn phí soạn / sửa văn bản Word" (vòng thử 23–24/11/2026). Ghép SAU cdp-lib.js và SAU tệp đặt biến BAI:
//   const BAI = { de, seed?: [đoạn…], provider: "openrouter", modelRe: "nemotron-3-ultra", nhac?: "…", buocToiDa?: 1 };
// Do kiem-thu/thu-mien-phi.mjs sinh và gọi. Trả: model đã chọn, từng lượt (giây, lần hỏi), đổi nguồn, chữ khung chat của bài,
// định dạng từng đoạn Word, sửa đổi chờ duyệt, tệp .docx (base64). Không đọc kho lưu của khung.
const NHAC = BAI.nhac || 'Những thông tin chưa có thì bạn tự giả định hợp lý hoặc để trống "…" cho tôi điền sau. Cứ soạn luôn vào tài liệu Word.';
const out = {};
for (let i = 0; i < 60 && !btn(/^Model đang dùng/); i++) await sleep(300);
// Khung còn đang chạy bài trước (lệnh eval trước bị ngắt nhưng mã trong khung vẫn chạy): bấm Dừng rồi chờ rảnh
if (busy()) {
  btn(/^Dừng$/)?.click();
  for (let i = 0; i < 60 && busy(); i++) await sleep(500);
  out.khungBan = "khung đang bận bài trước: đã bấm Dừng";
}
await toChat();
if (BAI.provider) {
  const p = await pick(BAI.provider, new RegExp(BAI.modelRe, "i"));
  out.model = p;
  if (p.loi) return out;
}
btn(/^Hội thoại mới/)?.click();
await sleep(600);

// Tài liệu trắng (hoặc chỉ có đoạn mồi): tắt theo dõi, chấp nhận sửa đổi cũ, xóa thân, ghi đoạn mồi
await Word.run(async (ctx) => {
  ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  const body = ctx.document.body;
  try {
    body.getTrackedChanges().acceptAll();
  } catch (e) {}
  body.clear();
  await ctx.sync();
  for (const t of BAI.seed || []) body.insertParagraph(t, "End");
  await ctx.sync();
  // Đoạn trống ban đầu nằm đầu tài liệu khi có đoạn mồi: xóa đi
  if ((BAI.seed || []).length) {
    const ps = body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    if (ps.items.length && ps.items[0].text === "") ps.items[0].delete();
    await ctx.sync();
  }
});

const wordChars = async () =>
  Word.run(async (ctx) => {
    const b = ctx.document.body;
    b.load("text");
    await ctx.sync();
    return b.text.replace(/\s+/g, "").length;
  });
const goc = await wordChars();
const t0 = Date.now();
const r1 = await send(BAI.de, { timeout: 420000, tail: 400 });
out.luot1 = { giay: r1.giay, daHoi: r1.daHoi, chuWord: await wordChars() };
if (r1.loi) out.luot1.loi = r1.loi;
// Lượt chạm trần số bước: bấm "Làm tiếp" một lần (như người dùng)
const tiep = async (ten) => {
  const b = btn(/^(Làm tiếp|Thử tiếp)$/);
  if (!b) return null;
  b.click();
  await sleep(600);
  const r = { giay: 0, daHoi: [] };
  const s = Date.now();
  while (busy() && Date.now() - s < 300000) {
    const allow = btn(/^Cho phép$/);
    if (allow) {
      r.daHoi.push(allow.parentElement.parentElement.innerText.replace(/\n+/g, " | ").slice(0, 200));
      (btn(/^Cho phép hết lần này$/) || allow).click();
    }
    await sleep(400);
  }
  r.giay = Math.round((Date.now() - s) / 1000);
  r.ten = ten;
  return r;
};
out.lamTiep = [];
const lt1 = await tiep("sau lượt 1");
if (lt1) out.lamTiep.push(lt1);
out.luot2 = null;
if (!BAI.seed && (await wordChars()) - goc < 200) {
  const r2 = await send(NHAC, { timeout: 420000, tail: 400 });
  out.luot2 = { giay: r2.giay, daHoi: r2.daHoi, chuWord: await wordChars() };
  const lt2 = await tiep("sau lượt 2");
  if (lt2) out.lamTiep.push(lt2);
}
out.tongGiay = Math.round((Date.now() - t0) / 1000);

// Chữ khung chat của bài (từ câu hỏi trở đi)
const all = document.body.innerText;
const k = all.indexOf(BAI.de.slice(0, 40));
out.chat = k >= 0 ? all.slice(k) : all.slice(-6000);
out.doiNguon = (out.chat.match(/↪[^\n]*/g) || []).slice(0, 20);
out.traLoiBoi = (out.chat.match(/Trả lời bởi[^\n]*/g) || []).slice(0, 5);
out.lanGoiAI = (out.chat.match(/\d+ lần gọi AI[^\n]*/g) || []).slice(0, 5);

// Định dạng từng đoạn và sửa đổi chờ duyệt
Object.assign(
  out,
  await Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text,items/alignment,items/tableNestingLevel,items/font/bold,items/font/italic,items/font/size,items/font/name");
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type,items/text");
    await ctx.sync();
    return {
      doanVan: ps.items.map((p) => ({ t: p.text, al: p.alignment, b: p.font.bold, i: p.font.italic, sz: p.font.size, f: p.font.name, tb: p.tableNestingLevel })),
      suaDoi: tc.items.map((x) => ({ loai: x.type, chu: x.text.slice(0, 200) })),
    };
  })
);
// Tệp .docx của tài liệu
out.docx = await new Promise((res) => {
  Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (r) => {
    if (r.status !== "succeeded") return res(null);
    const f = r.value;
    const parts = [];
    const next = (i) =>
      f.getSliceAsync(i, (s) => {
        if (s.status !== "succeeded") {
          f.closeAsync();
          return res(null);
        }
        parts.push(...s.value.data);
        if (i + 1 < f.sliceCount) next(i + 1);
        else {
          f.closeAsync();
          let bin = "";
          for (let j = 0; j < parts.length; j += 8192) bin += String.fromCharCode.apply(null, parts.slice(j, j + 8192));
          res(btoa(bin));
        }
      });
    next(0);
  });
});
return out;
