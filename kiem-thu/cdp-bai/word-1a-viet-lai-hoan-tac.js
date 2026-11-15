// Đợt 1A, tiêu chí 3 và 7 (Chạy thật): "Viết lại Điều 4.2…" rồi bấm nút Hoàn tác thay đổi trong khung.
// Chạy sau word-1a-doi-tac.js (Điều 4.2 khi đó đang có sửa đổi chờ duyệt và một nhận xét).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1a-viet-lai-hoan-tac.js
const look = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    const all = ctx.document.body.getTrackedChanges();
    all.load("items/type");
    const cs = ctx.document.body.getComments();
    cs.load("items/id");
    await ctx.sync();
    const i = ps.items.findIndex((p) => p.text.startsWith("Đơn giá thuê là"));
    const p = ps.items[i];
    const cur = p.getText();
    const tc = p.getTrackedChanges();
    tc.load("items/type,items/text,items/author");
    const near = [ps.items[i - 1].getText(), ps.items[i + 1].getText()];
    await ctx.sync();
    return {
      doan: i,
      hienTai: cur.value.replace(/[\r\n]+$/, ""),
      suaDoiTrongDoan: tc.items.map((t) => `${t.type} ${JSON.stringify(t.text.slice(0, 70))}${t.author === "Đối tác" ? " (Đối tác)" : ""}`),
      haiDoanBenCanh: near.map((x) => x.value.slice(0, 60)),
      tongSuaDoi: all.items.length,
      tongNhanXet: cs.items.length,
    };
  });
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const truoc = await look();
const them = "Đơn giá này không áp dụng cho diện tích sân bãi.";
const mong = truoc.hienTai.replace("chưa gồm thuế giá trị gia tăng", "đã gồm thuế giá trị gia tăng").replace("ba năm đầu", "hai năm đầu") + " " + them;
const q = `Viết lại Điều 4.2: đổi “chưa gồm thuế giá trị gia tăng” thành “đã gồm thuế giá trị gia tăng”, đổi “ba năm đầu” thành “hai năm đầu”, và thêm vào cuối điều câu “${them}”. Giữ nguyên mọi chữ khác của điều này.`;
const r = await send(q, { timeout: 300000 });
const all = document.body.innerText;
const ans = all.slice(all.lastIndexOf(q) + q.length);
const sau = await look();
const moi = sau.suaDoiTrongDoan.filter((x) => !truoc.suaDoiTrongDoan.includes(x));
// Bấm nút Hoàn tác của lượt vừa rồi
const undoBtn = [...document.querySelectorAll("button")].filter((b) => /^Hoàn tác thay đổi/.test(b.innerText)).pop();
const nhanNut = undoBtn ? undoBtn.innerText : "(không có nút)";
undoBtn?.click();
for (let k = 0; k < 40 && !/Đã hoàn tác|Hoàn tác \d+\/\d+/.test(document.body.innerText.slice(-600)); k++) await sleep(500);
await sleep(1500);
const sauHoanTac = await look();
return {
  yeuCau: q,
  cacBuoc: (ans.match(/^✅.*$/gm) || []).map((l) => l.slice(0, 80)),
  daHoi: r.daHoi,
  truoc,
  sauKhiSua: {
    dungCauMongMuon: sau.hienTai === mong,
    hienTai: sau.hienTai.slice(0, 200) + " … " + sau.hienTai.slice(-120),
    suaDoiMoiTrongDoan: moi,
    haiDoanBenCanhKhongDoi: JSON.stringify(sau.haiDoanBenCanh) === JSON.stringify(truoc.haiDoanBenCanh),
    tongSuaDoi: sau.tongSuaDoi,
  },
  nutHoanTac: nhanNut,
  thongBaoSauKhiBam: (document.body.innerText.slice(-500).match(/[^\n]*(hoàn tác|Hoàn tác)[^\n]*/g) || []).slice(-2),
  sauHoanTac: {
    veNhuTruoc: sauHoanTac.hienTai === truoc.hienTai,
    suaDoiTrongDoanNhuTruoc: JSON.stringify(sauHoanTac.suaDoiTrongDoan) === JSON.stringify(truoc.suaDoiTrongDoan),
    tongSuaDoi: sauHoanTac.tongSuaDoi,
    tongNhanXet: sauHoanTac.tongNhanXet,
  },
};
