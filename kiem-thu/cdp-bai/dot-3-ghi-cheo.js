// Đợt 3 (Chạy thật, tiêu chí B-2 và B-3): ghi chéo ứng dụng. Ba pha, chọn bằng window.__aidnPha ở tệp ghép phía trước:
//   "truoc"  (chạy ở khung ĐÍCH, ví dụ Word): dựng nội dung nền của tài liệu thử, trả trạng thái trước khi ghi
//   "gui"    (chạy ở khung GỬI, ví dụ Excel): chọn model, gửi yêu cầu (window.__aidnYeuCau), ghi lại câu hỏi xác nhận và trả lời;
//            window.__aidnTraLoi = "deny" thì bấm Bỏ qua ở câu hỏi ghi sang ứng dụng khác
//   "sau"    (chạy ở khung ĐÍCH): đọc tài liệu sau khi ghi, đọc mục "vừa ghi theo yêu cầu", bấm Hoàn tác, đọc lại
// Không đọc kho lưu của khung. Chạy: node kiem-thu/cdp.mjs eval 9229/<Ứng dụng> kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-3-ghi-cheo.js
const pha = window.__aidnPha;
const yeuCau = window.__aidnYeuCau;
const traLoi = window.__aidnTraLoi;
const model = window.__aidnModel || ["anthropic", "haiku-5-5"];
delete window.__aidnPha;
delete window.__aidnYeuCau;
delete window.__aidnTraLoi;
delete window.__aidnModel;
const host = Office.context.host;

const docWord = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items");
    await ctx.sync();
    const ts = ps.items.map((p) => p.getReviewedText("Current"));
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type");
    await ctx.sync();
    return { doan: ts.map((t) => t.value.replace(/[\r\n]+$/, "")).filter((t) => t.trim()), suaDoi: tc.items.length };
  });
const ketNoi = () => document.querySelector("[data-links]")?.innerText || null;

if (pha === "truoc") {
  if (host !== "Word") return { loi: "Pha này viết cho Word" };
  await Word.run(async (ctx) => {
    ctx.document.changeTrackingMode = "Off";
    await ctx.sync();
    ctx.document.body.clear();
    ctx.document.body.insertParagraph("BÁO CÁO TIẾN ĐỘ DỰ ÁN (văn bản thử ghi chéo)", "Start");
    ctx.document.body.insertParagraph("Phần mở đầu do người dùng viết, không được đụng tới.", "End");
    await ctx.sync();
  });
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  await sleep(400);
  return { ketNoi: ketNoi(), ...(await docWord()) };
}

if (pha === "gui") {
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  await sleep(500);
  const pk = await pick(model[0], new RegExp(model[1], "i"));
  document.body.click();
  await sleep(300);
  if (pk.loi) return { loi: "Không chọn được model", chiTiet: pk };
  const truoc = ketNoi();
  const r = await send(yeuCau, { timeout: 300000, tail: 2500, approve: (q) => (/Ghi sang/.test(q) ? traLoi || "allow" : "all") });
  return {
    model: label(btn(/^Model đang dùng/)).slice(-40),
    ketNoiTruoc: truoc,
    ketNoiSau: ketNoi(),
    giay: r.giay,
    daHoi: r.daHoi,
    hoiGhiCheo: r.daHoi.filter((q) => /Ghi sang/.test(q)).length,
    canhBaoNoiDaSua: !!document.querySelector("[data-claim-warning]"),
    cuoiKhung: r.khung.slice(-1600),
  };
}

if (pha === "sau") {
  if (host !== "Word") return { loi: "Pha này viết cho Word" };
  const sau = await docWord();
  const muc = [...document.querySelectorAll("[data-incoming]")].map((el) => el.innerText.slice(0, 300));
  const undo = btn(/^Hoàn tác thay đổi/);
  const out = { ketNoi: ketNoi(), sauKhiGhi: sau, mucVuaGhi: muc, coNutHoanTac: !!undo };
  if (undo && window.__aidnHoanTac !== false) {
    undo.click();
    await sleep(4000);
    out.sauHoanTac = await docWord();
    out.thongBao = tail(300);
  }
  delete window.__aidnHoanTac;
  return out;
}
return { loi: "Thiếu window.__aidnPha" };
