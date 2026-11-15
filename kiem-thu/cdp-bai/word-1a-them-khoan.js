// Đợt 1A, tiêu chí 4 (Chạy thật, bổ sung cho Tự kiểm tra): "Thêm một khoản sau 8.2" → khoản mới mang số 8.3, khoản cũ thành 8.4, cùng kiểu.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1a-them-khoan.js
const dieu8 = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text,items/style,items/listItemOrNullObject/listString");
    await ctx.sync();
    return ps.items
      .map((p, i) => ({ i, so: p.listItemOrNullObject.isNullObject ? "" : p.listItemOrNullObject.listString, kieu: p.style, chu: p.text.slice(0, 70) }))
      .filter((x) => /^(Điều 8\.|8\.\d)/.test(x.so));
  });
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const truoc = await dieu8();
const cau = "Bên Cho Thuê bàn giao kèm hồ sơ hoàn công và bản vẽ hiện trạng của Khu Nhà Xưởng.";
const q = `Thêm một khoản mới ngay sau Điều 8.2 với nội dung: “${cau}”`;
const r = await send(q, { timeout: 300000 });
const all = document.body.innerText;
const ans = all.slice(all.lastIndexOf(q) + q.length);
const sau = await dieu8();
const moi = sau.find((x) => x.chu.startsWith(cau.slice(0, 40)));
const cu = sau.find((x) => x.chu.startsWith("Nếu Bên Cho Thuê chậm bàn giao"));
return {
  yeuCau: q,
  cacBuoc: (ans.match(/^✅.*$/gm) || []).map((l) => l.slice(0, 80)),
  daHoi: r.daHoi,
  truoc,
  sau,
  khoanMoiMangSo: moi ? moi.so : "(không thấy)",
  khoanCuNayLa: cu ? cu.so : "(không thấy)",
  cungKieu: !!moi && !!cu && moi.kieu === cu.kieu,
  traLoi: ans
    .split("\n")
    .filter((l) => l.trim() && !/^(✅|⏳|Token:|Sửa đổi được ghi|Chèn vào chỗ con trỏ|Sao chép|Hỏi trước|Hoàn tác thay đổi|Anthropic · )/.test(l.trim()))
    .join(" ")
    .slice(0, 500),
};
