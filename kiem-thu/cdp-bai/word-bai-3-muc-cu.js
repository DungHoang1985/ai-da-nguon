// Bài 3 (Đợt 0B): các mục cũ chưa chạy thật — bấm địa chỉ đoạn, mở lại và xóa hội thoại, lưu Chỉ dẫn riêng (AI làm theo).
// Chạy ngay sau word-bai-2 (khung đang hiện câu trả lời có "[¶3]"), đã chọn model.
const out = {};
const selection = () =>
  Word.run(async (ctx) => {
    const s = ctx.document.getSelection();
    s.load("text");
    await ctx.sync();
    return s.text;
  });

// --- A. Bấm địa chỉ đoạn trong câu trả lời
await Word.run(async (ctx) => {
  ctx.document.body.getRange("Start").select();
  await ctx.sync();
});
const link = [...document.querySelectorAll('[role="link"]')].find((x) => /¶\s*3/.test(x.innerText));
if (!link) out.bamDiaChi = { loi: "Không thấy liên kết [¶3] trong khung" };
else {
  const truoc = await selection();
  link.click();
  await sleep(1500);
  out.bamDiaChi = { lienKet: link.innerText, vungChonTruoc: truoc.slice(0, 40), vungChonSau: (await selection()).slice(0, 60) };
}

// --- B. Mở lại hội thoại cũ, hỏi tiếp, rồi xóa một hội thoại
const rows = () => [...document.querySelectorAll('[title="Mở lại"]')];
const count = () => Number((/Lịch sử trong Word \((\d+)\)/.exec(document.body.innerText) || [])[1]);
btn(/^Lịch sử hội thoại/).click();
await sleep(1200);
const truocXoa = count();
const cu = rows().find((r) => /Xử lý các nhận xét/.test(r.innerText));
if (!cu) out.moLai = { loi: "Không thấy hội thoại 'Xử lý các nhận xét' trong lịch sử", soHoiThoai: truocXoa };
else {
  cu.click();
  await sleep(1500);
  const hienLai = /Tôi chưa sửa gì trong văn bản/.test(document.body.innerText);
  const hoiTiep = await send("Không dùng công cụ. Ở câu trả lời trước, anh nói nhận xét neo ở đoạn số mấy và ai viết? Trả lời một câu.", { tail: 500 });
  out.moLai = { soHoiThoai: truocXoa, hienLaiNoiDungCu: hienLai, hoiTiep: hoiTiep.khung.split("\n").slice(-7, -3).join(" / ") };
}
btn(/^Lịch sử hội thoại/).click();
await sleep(1200);
const xoa = rows().find((r) => /Sửa lần lượt từng điều/.test(r.innerText));
if (!xoa) out.xoa = { loi: "Không thấy hội thoại 'Sửa lần lượt từng điều' để xóa" };
else {
  const n0 = count();
  [...xoa.parentElement.querySelectorAll("button")].find((b) => /Xóa/.test(b.innerText)).click();
  await sleep(1500);
  out.xoa = { truoc: n0, sau: count(), conDong: rows().some((r) => /Sửa lần lượt từng điều/.test(r.innerText)) };
}
btn(/^Đóng$/)?.click();
await sleep(400);

// --- C. Chỉ dẫn riêng: thêm một dòng, lưu, hỏi AI, rồi trả lại nguyên văn chỉ dẫn cũ của người dùng
const MARK = "[AIDN-CHI-DAN-OK]";
const openSettings = async () => {
  btn(/^Cài đặt/).click();
  await sleep(1200);
  return [...document.querySelectorAll("textarea")].find((t) => /^Chỉ dẫn cho AI khi làm việc trong Word/.test(t.placeholder));
};
let ta = await openSettings();
if (!ta) out.chiDan = { loi: "Không thấy ô Chỉ dẫn riêng cho Word" };
else {
  if (window.__aidnChiDanGoc === undefined) window.__aidnChiDanGoc = ta.value; // giữ nguyên văn để trả lại; không in ra
  const goc = window.__aidnChiDanGoc;
  setNative(ta, `${goc}${goc ? "\n" : ""}Luôn kết thúc mọi câu trả lời bằng đúng một dòng riêng: ${MARK}`);
  await sleep(300);
  btn(/^Lưu chỉ dẫn$/).click();
  await sleep(800);
  btn(/^Hội thoại mới/)?.click();
  const r = await send("Không dùng công cụ. Một năm có mấy quý? Trả lời một câu.", { tail: 400 });
  const lamTheo = r.khung.includes(MARK);
  ta = await openSettings();
  setNative(ta, goc);
  await sleep(300);
  btn(/^Lưu chỉ dẫn$/).click();
  await sleep(800);
  ta = [...document.querySelectorAll("textarea")].find((t) => /^Chỉ dẫn cho AI khi làm việc trong Word/.test(t.placeholder));
  out.chiDan = { doDaiChiDanGoc: goc.length, aiLamTheo: lamTheo, traLoi: r.khung.split("\n").slice(-7, -3).join(" / "), daTraLaiNguyenVan: ta.value === goc };
  await toChat();
}
return out;
