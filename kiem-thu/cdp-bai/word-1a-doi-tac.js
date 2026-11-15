// Đợt 1A, tiêu chí 6: hợp đồng thử có 30 sửa đổi và 5 nhận xét của "Đối tác".
//  Lượt 1: tóm tắt đủ 30 sửa đổi kèm số điều.  Lượt 2: xử lý cả 5 nhận xét bằng Track Changes và trả lời từng nhận xét.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1a-doi-tac.js
// Tên tài khoản Office (tác giả các sửa đổi của AI) được che trước khi in: kho công khai.
const answerOf = (q) => {
  const all = document.body.innerText;
  return all.slice(all.lastIndexOf(q) + q.length);
};
const clean = (a) =>
  a
    .split("\n")
    .filter((l) => l.trim() && !/^(⏳|Token:|Sửa đổi được ghi|Chèn vào chỗ con trỏ|Sao chép|Hỏi trước|Hoàn tác thay đổi|Anthropic · )/.test(l.trim()))
    .join("\n");
const state = () =>
  Word.run(async (ctx) => {
    const body = ctx.document.body;
    const tc = body.getTrackedChanges();
    tc.load("items/type,items/author,items/text");
    const cs = body.getComments();
    cs.load("items/id,items/content,items/authorName,items/resolved");
    const ps = body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const reps = cs.items.map((c) => {
      const r = c.replies;
      r.load("items/content,items/authorName");
      return r;
    });
    const want = ["Đơn giá thuê là", "Bên Thuê đặt cọc cho Bên Cho Thuê", "Bên bị vi phạm gửi thông báo", "Mỗi bên giữ bí mật", "Mọi tranh chấp trước hết"];
    const idx = want.map((w) => ps.items.findIndex((p) => p.text.startsWith(w)));
    const cur = idx.map((i) => (i >= 0 ? ps.items[i].getText() : null));
    await ctx.sync();
    const me = [...new Set(tc.items.map((t) => t.author).filter((a) => a !== "Đối tác"))];
    return {
      me,
      suaDoiCuaDoiTac: tc.items.filter((t) => t.author === "Đối tác").length,
      suaDoiCuaAI: tc.items.filter((t) => t.author !== "Đối tác").map((t) => `${t.type} ${JSON.stringify(t.text.slice(0, 50))}`),
      nhanXet: cs.items.map((c, k) => ({
        noiDung: c.content.slice(0, 70),
        traLoi: reps[k].items.map((r) => r.content.slice(0, 160)),
        daXuLy: c.resolved,
      })),
      nămDoan: idx.map((i, k) => (cur[k] ? `[¶${i}] ` + cur[k].value.replace(/[\r\n]+$/, "").slice(0, 150) : "(không thấy)")),
    };
  });
const mask = (s, me) => me.reduce((t, name) => (name ? t.split(name).join("(tài khoản Office)") : t), s);

btn(/^Hội thoại mới/)?.click();
await sleep(600);
const truoc = await state();
const q1 =
  "Tóm tắt toàn bộ các sửa đổi được theo dõi (Track Changes) của đối tác trong hợp đồng này: liệt kê đủ từng sửa đổi theo số điều, không bỏ sót, và nêu tổng số sửa đổi. Chỉ báo cáo, không chấp nhận hay từ chối gì.";
const r1 = await send(q1, { timeout: 420000 });
const a1 = clean(answerOf(q1));
const CAN_CO = ["2.3", "3.2", "4.3", "5.1", "6.1", "6.4", "7.4", "8.3", "9.2", "10.3", "11.3", "15.2", "16.2", "22.4", "23.1", "23.3", "24.3", "27.3"];
const thieu = CAN_CO.filter((n) => !new RegExp(`(^|[^0-9.])${n.replace(".", "\\.")}(?![0-9])`).test(a1));

const q2 =
  "Xử lý cả 5 nhận xét của đối tác trong văn bản: sửa theo đúng yêu cầu của từng nhận xét bằng Track Changes, rồi trả lời từng nhận xét là đã sửa gì. Không đánh dấu đã xử lý.";
const r2 = await send(q2, { timeout: 480000 });
const a2 = clean(answerOf(q2));
const sau = await state();
const me = sau.me;
const KY_VONG = ["420.000", "02 tháng", "trong vòng 05 Ngày Làm Việc", "05 năm", "20 ngày"];
return {
  luot1: {
    giay: r1.giay,
    cacBuoc: (a1.match(/^✅.*$/gm) || []).map((l) => l.slice(0, 80)),
    neuTongSo30: /\b30\b/.test(a1),
    soDieuCanCo: CAN_CO.length,
    soDieuThieu: thieu,
    traLoi: mask(a1.replace(/^✅.*$/gm, "").trim(), me).slice(0, 6000),
  },
  luot2: {
    giay: r2.giay,
    daHoi: r2.daHoi,
    cacBuoc: (a2.match(/^✅.*$/gm) || []).map((l) => l.slice(0, 90)),
    traLoi: mask(a2.replace(/^✅.*$/gm, "").trim(), me).slice(0, 3000),
  },
  truoc: { suaDoiCuaDoiTac: truoc.suaDoiCuaDoiTac, suaDoiCuaAI: truoc.suaDoiCuaAI.length, nhanXet: truoc.nhanXet.length },
  sau: {
    suaDoiCuaDoiTac: sau.suaDoiCuaDoiTac,
    suaDoiCuaAI: sau.suaDoiCuaAI,
    nhanXet: sau.nhanXet.map((c) => ({ ...c, traLoi: c.traLoi.map((t) => mask(t, me)) })),
    namDoanSauKhiSua: sau.nămDoan,
    dungYeuCau: sau.nămDoan.map((t, k) => t.includes(KY_VONG[k])),
    nhanXetCoTraLoi: sau.nhanXet.filter((c) => c.traLoi.length > 0).length,
  },
};
