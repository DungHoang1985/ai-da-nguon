// Đợt 2, tiêu chí 5 (Chạy thật): key không có số dư → thông báo tiếng Việt rõ lý do, hội thoại còn nguyên.
// Dùng tài khoản Gemini đang báo hết tiền trả trước. Trình tự:
//   1. Anthropic · Claude Haiku 5.5: nhờ ghi nhớ một mã thử (lượt thành công, để có hội thoại trước lượt lỗi).
//   2. Đổi sang Gemini, gửi một câu → nhà cung cấp từ chối vì hết tiền. Đếm số lần gọi (phải là 1: không tự thử lại).
//   3. Bấm nút "Đổi nguồn" cạnh thông báo lỗi (bảng chọn phải mở), chọn lại Anthropic.
//   4. Hỏi lại mã thử → câu trả lời có mã = hội thoại trước lượt lỗi còn nguyên và dùng tiếp được.
// Không đọc kho lưu. Bộ đếm chỉ ghi tên miền, đường dẫn và mã trạng thái (key Gemini nằm trong tiêu đề, không nằm trong địa chỉ).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-het-tien.js
const MA = "SEN-4207";
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
btn(/^Model đang dùng/).click();
await sleep(700);
if (!window.__aidnNguonCu) window.__aidnNguonCu = document.querySelector('select[aria-label="Nhà cung cấp"]').value;
document.body.click();
await sleep(300);
window.__aidnTieuDeThem = ["^Không dùng công cụ\\. Ghi nhớ mã thử của phiên này"];
const out = { nguonCu: window.__aidnNguonCu };

// 1. Lượt thành công
out.b1_chon = await pick("anthropic", /haiku-5-5/);
document.body.click();
await sleep(300);
if (out.b1_chon.loi) return out;
const r1 = await send(`Không dùng công cụ. Ghi nhớ mã thử của phiên này: ${MA}. Trả lời đúng hai chữ: đã nhớ.`, { tail: 400 });
out.b1_traLoi = r1.khung.replace(/\n+/g, " | ").slice(-260);

// 2. Gemini hết tiền
out.b2_chon = await pick("gemini", /^gemini-2\.5-flash$/);
document.body.click();
await sleep(300);
if (out.b2_chon.loi) return out;
const calls = [];
const realFetch = window.fetch;
window.fetch = async (...a) => {
  const u = new URL(String(a[0] && a[0].url ? a[0].url : a[0]), location.href);
  try {
    const r = await realFetch(...a);
    calls.push(`${u.hostname}${u.pathname.replace(/models\/[^:/]+/, "models/<model>")} -> ${r.status}`);
    return r;
  } catch (e) {
    calls.push(`${u.hostname}${u.pathname} -> lỗi mạng`);
    throw e;
  }
};
let r2;
const t0 = Date.now();
try {
  r2 = await send("Không dùng công cụ. Câu thử hết tiền: một tuần có mấy ngày?", { tail: 1200 });
} finally {
  window.fetch = realFetch;
}
out.b2_giay = Math.round((Date.now() - t0) / 1000);
out.b2_cacLanGoiGemini = calls.filter((c) => /generativelanguage/.test(c));
const errEl = [...document.querySelectorAll("span")].reverse().find((el) => el.children.length === 0 && /hết tiền|Hết hạn mức|Key |từ chối|Lỗi/.test(el.innerText || ""));
out.b2_thongBao = errEl ? errEl.innerText : "(không thấy thông báo lỗi) " + r2.khung.slice(-300);
out.b2_nutCanhLoi = [...document.querySelectorAll("button")].map((b) => b.innerText.trim()).filter((t) => /^(Đổi nguồn|Mở Cài đặt|Chép chi tiết lỗi)$/.test(t));
const text = document.body.innerText;
out.b2_cauHoiConHien = text.includes("Câu thử hết tiền: một tuần có mấy ngày?") && text.includes(`Ghi nhớ mã thử của phiên này: ${MA}`);

// 3. Nút "Đổi nguồn" mở bảng chọn
const swap = btn(/^Đổi nguồn$/);
if (swap) {
  swap.click();
  await sleep(800);
  out.b3_bangChonMo = !!document.querySelector('select[aria-label="Nhà cung cấp"]');
  document.body.click();
  await sleep(300);
} else out.b3_bangChonMo = "(không có nút Đổi nguồn)";
out.b3_chonLai = await pick("anthropic", /haiku-5-5/);
document.body.click();
await sleep(300);

// 4. Hội thoại còn nguyên
const r4 = await send("Không dùng công cụ. Mã thử tôi nhờ ghi nhớ ở đầu hội thoại này là gì? Chỉ ghi đúng mã đó.", { tail: 500 });
const answer = r4.khung.replace(/\n+/g, " | ");
out.b4_traLoi = answer.slice(-300);
out.b4_nhoDungMa = answer.split("Chỉ ghi đúng mã đó.").pop().includes(MA);
out.b4_soLuotTrongKhung = (document.body.innerText.match(/Không dùng công cụ\./g) || []).length;
return out;
