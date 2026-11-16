// Đợt 3 (Chạy thật, tiêu chí A-a): giao AI soạn một văn bản hành chính trong Word rồi ĐỌC LẠI văn bản bằng mã để đối chiếu thể thức.
// Trước mỗi lần chạy: xóa trắng văn bản thử, trả lề về 25,4 mm và bỏ số trang (để chứng minh công cụ tự đặt thể thức trang).
// Tham số đặt ở một tệp ghép phía trước: window.__aidnBai = { nguon: "anthropic", model: "haiku-5-5", yeuCau: "…", loai: "cong_van" | "to_trinh" }
// (không có nguon thì dùng nguồn và model đang chọn). Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Word kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-3-soan-van-ban.js
const bai = window.__aidnBai || {};
delete window.__aidnBai;
if (!bai.yeuCau) return { loi: "Thiếu window.__aidnBai.yeuCau" };
const mm = (pt) => Math.round(((pt * 25.4) / 72) * 10) / 10;

const xoaTrang = () =>
  Word.run(async (ctx) => {
    ctx.document.changeTrackingMode = "Off";
    await ctx.sync();
    ctx.document.body.clear();
    const s = ctx.document.sections.getFirst();
    s.getHeader("Primary").clear();
    s.pageSetup.differentFirstPageHeaderFooter = false;
    s.pageSetup.paperSize = "Letter";
    s.pageSetup.topMargin = 72;
    s.pageSetup.bottomMargin = 72;
    s.pageSetup.leftMargin = 72;
    s.pageSetup.rightMargin = 72;
    await ctx.sync();
    const p = ctx.document.body.paragraphs.getFirst();
    p.styleBuiltIn = "Normal";
    p.font.name = "Calibri";
    p.font.size = 11;
    p.alignment = "Left";
    p.firstLineIndent = 0;
    await ctx.sync();
  });

const docLai = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text,items/alignment,items/firstLineIndent,items/spaceAfter,items/lineSpacing,items/font/name,items/font/size,items/font/bold,items/font/italic,items/font/color,items/tableNestingLevel");
    const ts = ctx.document.body.tables;
    ts.load("items");
    const s = ctx.document.sections.getFirst();
    s.pageSetup.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight,differentFirstPageHeaderFooter");
    const h = s.getHeader("Primary");
    h.load("text");
    const tc = ctx.document.body.getTrackedChanges();
    tc.load("items/type");
    await ctx.sync();
    const empties = ps.items.filter((p) => !p.text.trim());
    empties.forEach((p) => p.borders.load("items/lineStyle"));
    await ctx.sync();
    return {
      trang: { tren: mm(s.pageSetup.topMargin), duoi: mm(s.pageSetup.bottomMargin), trai: mm(s.pageSetup.leftMargin), phai: mm(s.pageSetup.rightMargin), rong: mm(s.pageSetup.pageWidth), cao: mm(s.pageSetup.pageHeight), soTrangDauTrang: h.text.trim(), anTrangDau: s.pageSetup.differentFirstPageHeaderFooter },
      bang: ts.items.length,
      duongKe: empties.filter((p) => p.borders.items[2] && String(p.borders.items[2].lineStyle) === "Single").length,
      suaDoi: tc.items.length,
      doan: ps.items.map((p) => ({ t: p.text, a: String(p.alignment), i: mm(p.firstLineIndent), sa: p.spaceAfter, f: p.font.name, s: p.font.size, b: p.font.bold, it: p.font.italic, c: p.font.color, tb: p.tableNestingLevel > 0 })),
    };
  });

/** Bảng đối chiếu thể thức, tính độc lập với công cụ của add-in */
function doiChieu(d, loai) {
  const P = d.doan.filter((x) => x.t.trim());
  const tim = (re) => P.find((x) => re.test(x.t));
  const kq = [];
  const them = (ten, dat, thay) => kq.push(`${dat ? "Đạt" : "KHÔNG"} · ${ten}${thay ? " — " + thay : ""}`);
  const qh = tim(/^CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM$/);
  const tn = tim(/^Độc lập - Tự do - Hạnh phúc$/);
  them("Quốc hiệu in hoa, đậm, cỡ 12–13, canh giữa", !!qh && qh.b && qh.s >= 12 && qh.s <= 13 && qh.a === "Centered", qh ? `cỡ ${qh.s}` : "không thấy");
  them("Tiêu ngữ đậm, cỡ 13–14, lớn hơn quốc hiệu 1 cỡ", !!tn && !!qh && tn.b && tn.s >= 13 && tn.s <= 14 && tn.s - qh.s === 1, tn ? `cỡ ${tn.s}` : "không thấy");
  const coQuan = P.filter((x) => x.tb && x.t === x.t.toLocaleUpperCase("vi-VN") && x.t.length > 8 && !/CỘNG HÒA/.test(x.t) && /[A-ZĐ]/.test(x.t)).slice(0, 2);
  them("Tên cơ quan ban hành in hoa, đậm, cỡ 12–13", coQuan.some((x) => x.b && x.s >= 12 && x.s <= 13), coQuan.map((x) => `“${x.t.slice(0, 50)}” ${x.b ? "đậm" : "thường"} ${x.s}`).join(" | "));
  const so = tim(/^Số:\s/);
  them("Số và ký hiệu, cỡ 13", !!so && so.s === 13, so ? `“${so.t}”` : "không thấy");
  const ngay = tim(/, ngày\s+\S*\s*tháng\s+\S*\s*năm/);
  them("Địa danh, ngày tháng: nghiêng, cỡ 13–14", !!ngay && ngay.it && ngay.s >= 13 && ngay.s <= 14, ngay ? `“${ngay.t}”` : "không thấy");
  if (loai === "cong_van") {
    const vv = tim(/^V\/v\s/);
    them("V/v dưới số ký hiệu, cỡ 12–13, đứng", !!vv && vv.tb && vv.s >= 12 && vv.s <= 13 && !vv.b && !vv.it, vv ? `“${vv.t.slice(0, 80)}”` : "không thấy");
  } else {
    const ten = tim(/^(TỜ TRÌNH|BÁO CÁO|QUYẾT ĐỊNH|THÔNG BÁO|GIẤY MỜI|BIÊN BẢN|KẾ HOẠCH)$/);
    const k = ten ? P.indexOf(ten) : -1;
    const ty = k >= 0 ? P[k + 1] : null;
    them("Tên loại in hoa, đậm, canh giữa, cỡ 13–14", !!ten && ten.b && ten.a === "Centered" && ten.s >= 13 && ten.s <= 14, ten ? ten.t : "không thấy");
    them("Trích yếu đậm, canh giữa, ngay dưới tên loại", !!ty && ty.b && ty.a === "Centered", ty ? `“${ty.t.slice(0, 80)}”` : "không thấy");
  }
  const kg = tim(/^Kính gửi:/);
  them("Kính gửi, cỡ 13–14", !!kg && kg.s >= 13 && kg.s <= 14, kg ? `“${kg.t.slice(0, 80)}”` : "không thấy");
  // Đoạn nội dung = đoạn dài ngoài bảng, không đậm (đề mục), không canh giữa (trích yếu, Kính gửi)
  const nd = P.filter((x) => !x.tb && x.t.length > 60 && !x.b && x.a !== "Centered");
  const ndDung = nd.filter((x) => x.s >= 13 && x.s <= 14 && x.a === "Justified" && x.i >= 9.5 && x.i <= 13 && x.sa >= 6 && !x.b);
  them("Nội dung: cỡ 13–14, canh đều, thụt đầu dòng 10–12,7 mm, cách đoạn ≥ 6 pt", nd.length > 0 && ndDung.length === nd.length, `${ndDung.length}/${nd.length} đoạn dài đúng`);
  them("Kết thúc nội dung bằng “./.”", P.some((x) => /\.\/\.$/.test(x.t)));
  const nn = tim(/^Nơi nhận:$/);
  them("“Nơi nhận:” cỡ 12, nghiêng, đậm", !!nn && nn.s === 12 && nn.b && nn.it);
  const luu = tim(/^- Lưu:/);
  them("Dòng “- Lưu: …” cỡ 11, kết bằng dấu chấm", !!luu && luu.s === 11 && /\.$/.test(luu.t), luu ? `“${luu.t}”` : "không thấy");
  const k2 = nn ? P.indexOf(nn) : -1;
  const ky = P.filter((x, i) => x.tb && i >= k2 && k2 >= 0 && x.b && x.t !== "Nơi nhận:");
  them("Phần ký: chức vụ in hoa đậm, họ tên đậm", ky.length >= 2 && ky[0].t === ky[0].t.toLocaleUpperCase("vi-VN"), ky.map((x) => x.t).join(" / "));
  them("Phông Times New Roman, chữ đen ở mọi đoạn", P.every((x) => x.f === "Times New Roman" && /^#?0{6}$/i.test(String(x.c))), [...new Set(P.map((x) => x.f))].join(", "));
  them("Đường kẻ dưới tên cơ quan, dưới tiêu ngữ" + (loai === "cong_van" ? "" : ", dưới trích yếu"), d.duongKe === (loai === "cong_van" ? 2 : 3), `${d.duongKe} đường kẻ`);
  const t = d.trang;
  them("Giấy A4, lề trên 20–25, dưới 20–25, trái 30–35, phải 15–20 mm", Math.abs(t.rong - 210) < 0.6 && Math.abs(t.cao - 297) < 0.6 && t.tren >= 19.9 && t.tren <= 25.1 && t.duoi >= 19.9 && t.duoi <= 25.1 && t.trai >= 29.9 && t.trai <= 35.1 && t.phai >= 14.9 && t.phai <= 20.1, `${t.rong}×${t.cao} mm; lề ${t.tren}/${t.duoi}/${t.trai}/${t.phai}`);
  them("Số trang ở đầu trang, ẩn trang đầu", !!t.soTrangDauTrang && t.anTrangDau, `đầu trang: “${t.soTrangDauTrang}”`);
  return kq;
}

await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(600);
if (bai.nguon) {
  const pk = await pick(bai.nguon, new RegExp(bai.model, "i"));
  document.body.click();
  await sleep(300);
  if (pk.loi) return { loi: "Không chọn được model", chiTiet: pk };
}
const model = label(btn(/^Model đang dùng/)).slice(0, 120);
await xoaTrang();
const truoc = await docLai();
const r = await send(bai.yeuCau, { timeout: 420000, tail: 3500 });
const sau = await docLai();
const khung = document.body.innerText;
const buoc = [...document.querySelectorAll("span, div")]
  .filter((el) => el.children.length === 0)
  .map((el) => (el.innerText || "").trim())
  .filter((t) => /^(Dựng văn bản hành chính|Kiểm thể thức trang|Đặt khổ giấy|Chèn|Định dạng|Đọc|Xem|Sửa|Thay|Dùng kỹ năng|Chèn bảng)/.test(t))
  .slice(0, 30);
const dc = doiChieu(sau, bai.loai || "cong_van");
return {
  model,
  yeuCau: bai.yeuCau,
  giay: r.giay,
  lanGoiAI: Number([...document.querySelectorAll("[data-calls]")].pop()?.getAttribute("data-calls") || 0),
  daHoi: r.daHoi,
  kyNangTuNap: (khung.match(/Kỹ năng[^\n]*soan-thao-van-ban-hanh-chinh[^\n]*/) || [""])[0].slice(0, 120) || /soan-thao-van-ban-hanh-chinh/.test(khung),
  canhBaoNoiDaSua: !!document.querySelector("[data-claim-warning]"),
  buocCongCu: buoc,
  truoc: { soDoanCoChu: truoc.doan.filter((x) => x.t.trim()).length, trang: truoc.trang },
  sau: { soDoanCoChu: sau.doan.filter((x) => x.t.trim()).length, bang: sau.bang, suaDoiDuocTheoDoi: sau.suaDoi, trang: sau.trang },
  doiChieu: dc,
  dat: `${dc.filter((x) => x.startsWith("Đạt")).length}/${dc.length}`,
  vanBan: sau.doan.filter((x) => x.t.trim()).map((x) => `${x.tb ? "▦ " : ""}${x.t.slice(0, 150)}`),
  cuoiKhung: r.khung.slice(-1500),
};
