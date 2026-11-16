// Đợt 4 (việc tồn T-1, T-2): gửi một yêu cầu trong Word rồi đọc lại khung và số nhận xét của văn bản.
// Tham số ở tệp ghép phía trước: window.__aidnBai = { yeuCau, nguon?, model?, dungDoan?: "chữ của đoạn thử thêm vào cuối nếu văn bản chưa có" }
// Trả: model trước / lúc thử, phần cuối khung (mã tổ chức nếu lọt ra sẽ bị đếm, không in), số nhận xét trước / sau.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Word kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-4-word-gui.js
const bai = window.__aidnBai || {};
delete window.__aidnBai;
if (!bai.yeuCau) return { loi: "Thiếu window.__aidnBai.yeuCau" };
const soNhanXet = () =>
  Word.run(async (ctx) => {
    const cs = ctx.document.body.getComments();
    cs.load("items/id");
    await ctx.sync();
    return cs.items.length;
  }).catch((e) => "LỖI " + (e.code || e.message));
if (bai.dungDoan) {
  await Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    if (!ps.items.some((p) => p.text.includes(bai.dungDoan))) ctx.document.body.insertParagraph(bai.dungDoan, "End");
    await ctx.sync();
  });
}
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
const truoc = label(btn(/^Model đang dùng/) || document.body).slice(0, 120);
let chon = null;
if (bai.nguon && bai.model) chon = await pick(bai.nguon, new RegExp(bai.model, "i"));
if (chon?.loi) return { loi: chon.loi, coSan: chon.coSan, modelTruoc: truoc };
const n0 = await soNhanXet();
const r = await send(bai.yeuCau, { approve: () => "allow", timeout: 180000, tail: 2000 });
const khung = (r.khung || "").replace(/\n{2,}/g, "\n");
const all = document.body.innerText;
return {
  modelTruoc: truoc,
  modelThu: chon?.daChon || truoc,
  phienBan: (all.match(/v\d+\.\d+(\.\d+)?/) || [""])[0],
  giay: r.giay,
  nhanXetTruoc: n0,
  nhanXetSau: await soNhanXet(),
  // Kiểm lộ thông tin riêng mà không in ra: đếm chuỗi dạng mã tổ chức và dạng key trong toàn bộ chữ của khung
  soMaToChucTrenKhung: (all.match(/\borg[_-][A-Za-z0-9]{8,}/g) || []).length,
  coNutDoiNguon: !!btn(/^Đổi nguồn/),
  khung: khung.replace(/\borg[_-][A-Za-z0-9]{8,}/g, "org_<ĐÃ CHE KHI GHI BẰNG CHỨNG>"),
};
