// Đợt 2 (Chạy thật, nửa sau của tiêu chí 1 và 2): với khóa OpenRouter đang có trên máy, danh sách model hiện và sửa được một đoạn Word.
// Dùng model OpenRouter đang chọn (không đổi model, không đổi khóa). Tệp thử: hợp đồng dựng bằng `node kiem-thu/soat-thu.mjs sach`.
// Yêu cầu: đổi "150%" thành "120%" ở câu về chậm hoàn trả. Kiểm bằng cách đọc lại văn bản và các sửa đổi được theo dõi.
// Không đọc kho lưu. Tên tác giả sửa đổi khác "Đối tác" được che.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-openrouter-sua-doan.js
const mask = (a) => (a === "Đối tác" ? a : "(người dùng)");
const doc = () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items");
    await ctx.sync();
    const texts = ps.items.map((p) => p.getText());
    const tcs = ctx.document.body.getTrackedChanges();
    tcs.load("items/type,items/text,items/author");
    await ctx.sync();
    const i = texts.findIndex((t) => /Chậm hoàn trả, Bên Thuê trả tiền sử dụng/.test(t.value));
    return {
      doan: i,
      cauHienTai: i >= 0 ? texts[i].value.replace(/[\r\t]+$/, "").slice(0, 160) : "(không thấy)",
      suaDoi: tcs.items.filter((t) => /1[25]0%|150|120/.test(t.text)).map((t) => `${t.type}: “${t.text.slice(0, 40)}” — ${mask(t.author)}`),
      tongSuaDoi: tcs.items.length,
    };
  });
await toChat();
window.__aidnTieuDeThem = [...(window.__aidnTieuDeThem || []), "^Ở câu về chậm hoàn trả"];
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const model = label(btn(/^Model đang dùng/)).slice(0, 110);
if (!/OpenRouter/.test(model)) return { loi: "Nguồn đang chọn không phải OpenRouter", model };
btn(/^Model đang dùng/).click();
await sleep(900);
const soModel = document.querySelectorAll('[role="listbox"][aria-label="Model"] [role="option"]').length;
document.body.click();
await sleep(300);
const truoc = await doc();
const q = "Ở câu về chậm hoàn trả (“Chậm hoàn trả, Bên Thuê trả tiền sử dụng bằng 150% Tiền Thuê…”), đổi “150%” thành “120%”. Chỉ sửa đúng chỗ đó.";
const r = await send(q, { timeout: 300000 });
const all = document.body.innerText;
const ans = all.slice(all.lastIndexOf(q) + q.length);
const sau = await doc();
return {
  model,
  soModelDangHien: soModel,
  yeuCau: q,
  giay: r.giay,
  daHoi: r.daHoi,
  cacBuoc: (ans.match(/^[✅❌].*$/gm) || []).map((l) => l.slice(0, 110)),
  truoc,
  sau,
  daDoi: /120%/.test(sau.cauHienTai) && !/150%/.test(sau.cauHienTai),
  coSuaDoiDuocTheoDoi: sau.tongSuaDoi > truoc.tongSuaDoi,
  traLoi: ans
    .split("\n")
    .filter((l) => l.trim() && !/^(✅|❌|⏳|Token:|Sửa đổi được ghi|Chèn vào chỗ con trỏ|Sao chép|Hỏi trước|Hoàn tác thay đổi|OpenRouter · )/.test(l.trim()))
    .join(" ")
    .slice(0, 400),
};
