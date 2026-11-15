// Đợt 1B: gửi một bài Chạy thật (window.__aidnBai1B[window.__aidnChon]) trong một hội thoại mới rồi ghi lại: các bước công cụ,
// câu trả lời, dòng token, và trạng thái văn bản trước / sau (số đoạn, số sửa đổi, số nhận xét, mã băm chữ) để chứng minh văn bản
// có đổi hay không. Đặt window.__aidnGiuHoiThoai = true để hỏi tiếp trong hội thoại đang mở (không tạo hội thoại mới).
// Tên tác giả của sửa đổi không được in ra (chỉ đếm).
const q = window.__aidnBai1B[window.__aidnChon];
if (!q) return { loi: "Không có bài " + window.__aidnChon };
const state = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type");
    const cs = ctx.document.body.getComments();
    cs.load("items/id");
    await ctx.sync();
    let h = 0;
    for (const p of ps.items) for (let i = 0; i < p.text.length; i++) h = (Math.imul(h, 31) + p.text.charCodeAt(i)) | 0;
    return { soDoan: ps.items.length, soSuaDoi: tc.items.length, soNhanXet: cs.items.length, bamChu: h };
  });
const truoc = await state();
if (!window.__aidnGiuHoiThoai) {
  // Giữ tệp đã đặt sẵn vào ô đính kèm: nút "Hội thoại mới" không xóa tệp chờ gửi
  btn(/^Hội thoại mới/)?.click();
  await sleep(600);
}
const tepCho = [...document.querySelectorAll("span")].filter((x) => /^(📄|📑|🖼) /.test(x.innerText || "")).map((x) => x.innerText.replace(/\s*✕\s*$/, "").trim());
const r = await send(q, { timeout: 420000, tail: 400 });
const all = document.body.innerText;
const from = all.lastIndexOf(q.slice(0, 60));
const answer = all.slice(from + q.length);
const lines = answer.split("\n").map((l) => l.trim()).filter(Boolean);
const buoc = lines.filter((l) => /^(✅|❌)/.test(l)).map((l) => l.slice(0, 160));
const token = (answer.match(/Token: [^\n]+/g) || []).slice(-1)[0] || "";
const traLoi = lines.filter((l) => !/^(✅|❌|⏳|Token:|Sửa đổi được ghi|Gửi|Hội thoại mới|Hoàn tác thay đổi|📄|📑|🧠|🌐|Tự sửa|Hỏi trước)/.test(l)).join("\n");
const sau = await state();
const model = label(btn(/^Model đang dùng/));
return {
  bai: window.__aidnChon,
  phienBan: (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0],
  model,
  cauHoi: q,
  tepDinhKem: tepCho,
  giay: r.giay,
  daHoi: r.daHoi,
  soBuocCongCu: buoc.length,
  cacBuoc: buoc,
  token,
  vanBanTruoc: truoc,
  vanBanSau: sau,
  vanBanKhongDoi: JSON.stringify(truoc) === JSON.stringify(sau),
  traLoi,
};
