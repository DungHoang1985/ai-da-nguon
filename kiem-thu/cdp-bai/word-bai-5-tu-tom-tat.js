// Bài 5 (Đợt 0B): tự tóm tắt hội thoại dài. Dựng văn bản 100 đoạn dài (~40.000 ký tự), mỗi lượt bắt AI đọc lại toàn bộ
// bằng công cụ cho tới khi hội thoại vượt ngưỡng của model và khung báo "đã tự tóm tắt". Sau đó hỏi lại để xem AI còn nhớ việc.
await Word.run(async (ctx) => {
  ctx.document.load("changeTrackingMode");
  await ctx.sync();
  ctx.document.changeTrackingMode = "Off";
  const body = ctx.document.body;
  body.clear();
  const filler = "nội dung thử để kiểm việc tự tóm tắt hội thoại dài của AI Đa Nguồn, ".repeat(5);
  for (let k = 1; k <= 100; k++) body.insertParagraph(`AIDN-DAI đoạn ${k}: ${filler}mã cuối đoạn ${k * 7}.`, "End");
  await ctx.sync();
});
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
const NOTICE = /đã tự tóm tắt phần đầu/;
const turns = [];
let compactedAt = 0;
for (let i = 1; i <= 16; i++) {
  const r = await send(
    i === 1
      ? "Việc của phiên này: kiểm văn bản thử AIDN-DAI. Đọc toàn bộ văn bản bằng công cụ (đọc tiếp theo next_start cho tới hết) rồi cho biết: văn bản có bao nhiêu đoạn và “mã cuối đoạn” của đoạn 100 là số nào. Trả lời một câu."
      : `Lần ${i}: đọc lại toàn bộ văn bản bằng công cụ từ đầu (không dùng trí nhớ) rồi cho biết “mã cuối đoạn” của đoạn ${i * 5} là số nào. Trả lời một câu.`,
    { tail: 500, timeout: 200000 }
  );
  const lines = r.khung.split("\n");
  const tok = lines.filter((x) => /^Token:/.test(x)).pop() || "";
  const ans = lines[Math.max(0, lines.findIndex((x) => /^Token:/.test(x)) - 1)] || "";
  turns.push(`Lượt ${i} (${r.giay}s): ${ans.slice(0, 90)} | ${tok}`);
  if (NOTICE.test(document.body.innerText)) {
    compactedAt = i;
    break;
  }
  if (/Lỗi|Hết hạn mức|quá tải/.test(lines.slice(-6).join(" "))) {
    turns.push("DỪNG: có lỗi — " + lines.slice(-6).join(" / ").slice(0, 300));
    break;
  }
}
const thongBao = (document.body.innerText.match(/[^\n]*đã tự tóm tắt[^\n]*/) || [""])[0];
let sau = null;
if (compactedAt) {
  const r = await send("Không dùng công cụ. Việc của phiên này là gì, và “mã cuối đoạn” của đoạn 100 mà anh đã báo ở lượt đầu là số nào? Trả lời một câu.", { tail: 500 });
  sau = r.khung.split("\n").slice(-6, -3).join(" / ");
}
return { tomTatOLuot: compactedAt, thongBao, cacLuot: turns, hoiLaiSauTomTat: sau };
