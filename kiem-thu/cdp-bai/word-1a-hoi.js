// Đợt 1A, tiêu chí 2: hỏi từng câu trong một hội thoại mới, ghi lại số bước công cụ, câu trả lời và dòng token.
// Danh sách câu hỏi đặt ở window.__aidnCauHoi = [[câu hỏi, số điều đúng, cụm phải có], …] (tệp word-1a-cau-hoi.js ghép phía trước).
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1a-cau-hoi.js kiem-thu/cdp-bai/word-1a-hoi.js
const out = [];
const steps = () => (document.body.innerText.match(/^[✅❌⚠️] ?.*$/gm) || []).filter((l) => /^(✅|❌)/.test(l));
for (const [q, dieu, cum] of window.__aidnCauHoi) {
  btn(/^Hội thoại mới/)?.click();
  await sleep(600);
  const r = await send(q, { timeout: 240000, tail: 2500 });
  const all = document.body.innerText;
  const from = all.lastIndexOf(q);
  const answer = all.slice(from + q.length);
  const st = steps();
  const tok = (answer.match(/Token: [^\n]+/g) || []).slice(-1)[0] || "";
  // Phần trả lời: bỏ các dòng bước công cụ và dòng token
  const text = answer
    .split("\n")
    .filter((l) => l.trim() && !/^(✅|❌|⏳|Token:|Sửa đổi được ghi|Gửi|Hội thoại mới)/.test(l.trim()))
    .join(" ")
    .slice(0, 700);
  const norm = (s) => s.replace(/\s+/g, " ");
  out.push({
    cauHoi: q,
    dieuDung: dieu,
    soBuocCongCu: st.length,
    cacBuoc: st.map((l) => l.slice(0, 70)),
    neuDungSoDieu: norm(text).includes(dieu),
    coCumCanCo: norm(text).toLowerCase().includes(String(cum).toLowerCase()),
    giay: r.giay,
    traLoi: text,
    token: tok,
  });
}
const dung = out.filter((x) => x.neuDungSoDieu && x.coCumCanCo).length;
return { tongKet: `${dung}/${out.length} câu nêu đúng số điều và nội dung; số bước công cụ nhiều nhất: ${Math.max(...out.map((x) => x.soBuocCongCu))}`, ketQua: out };
