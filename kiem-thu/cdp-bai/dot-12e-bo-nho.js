// Đợt 12 phần E — E2: bộ nhớ dài hạn trong khung add-in thật (Word), KHÔNG gọi model thật.
// Lời gọi model bị chặn ngay trong khung (fetch có thân JSON chứa messages / contents): bài thử chỉ đọc thân yêu cầu sắp gửi
// (không đọc tiêu đề, không in key), trả 400 giả — nhà cung cấp không nhận gì. Đề xuất nhớ của AI thử bằng một trả lời giả kiểu OpenAI
// (chỉ khi nguồn đang chọn nói kiểu OpenAI). Mục thử có tiền tố "E2-thu" và bị xóa lại qua giao diện ở cuối bài.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-12e-bo-nho.js
const out = { buoc: [] };
const ok = (ten, dat, chiTiet) => out.buoc.push({ ten, dat: !!dat, ...(chiTiet !== undefined ? { chiTiet } : {}) });
const A = "E2-thu: Trả lời kết luận trước, dùng bảng Tiêu chí · Kết quả.";
const B = "E2-thu: Số tiền ghi đơn vị triệu đồng.";
const P = "E2-thu: Công văn ký thừa lệnh Giám đốc.";

const openSettings = async () => {
  if (!document.querySelector("[data-memory-card]")) {
    // Khung vừa tải lại thì nút Cài đặt chưa có: chờ rồi bấm
    for (let i = 0; i < 50 && !document.querySelector("[data-memory-card]"); i++) {
      if (i % 5 === 0) btn(/^Cài đặt/)?.click();
      await sleep(300);
    }
  }
  return document.querySelector("[data-memory-card]");
};
const memTexts = () => [...document.querySelectorAll("[data-memory-item]")].map((el) => el.innerText.split("\n")[0]);
const addMem = async (text) => {
  const card = await openSettings();
  const ta = card.querySelector("[data-memory-input]");
  setNative(ta, text);
  await sleep(300);
  card.querySelector("[data-memory-add]").click();
  await sleep(500);
  return card.innerText.split("\n").slice(-1)[0];
};

// Chặn lời gọi model: ghi thân yêu cầu, trả lỗi 400 giả (hoặc lệnh gọi remember giả)
const sent = [];
let fakeRemember = false;
const realFetch = window.fetch;
window.fetch = async (input, init) => {
  const body = init && typeof init.body === "string" ? init.body : "";
  if (/"(messages|contents)"\s*:/.test(body)) {
    sent.push(body);
    if (fakeRemember && /"messages"/.test(body) && !/"role":"tool"/.test(body)) {
      const ev = (o) => `data: ${JSON.stringify(o)}\n\n`;
      const txt =
        ev({ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "remember", arguments: JSON.stringify({ text: P }) } }] } }] }) +
        ev({ choices: [{ delta: {}, finish_reason: "tool_calls" }] }) +
        "data: [DONE]\n\n";
      return new Response(txt, { status: 200, headers: { "content-type": "text/event-stream" } });
    }
    if (fakeRemember) {
      const ev = (o) => `data: ${JSON.stringify(o)}\n\n`;
      return new Response(ev({ choices: [{ delta: { content: "Em đã đề xuất nhớ." } }] }) + ev({ choices: [{ delta: {}, finish_reason: "stop" }] }) + "data: [DONE]\n\n", {
        status: 200,
        headers: { "content-type": "text/event-stream" },
      });
    }
    return new Response(JSON.stringify({ error: { message: "bai thu E2: chan loi goi model" } }), { status: 400, headers: { "content-type": "application/json" } });
  }
  return realFetch(input, init);
};
const lastSent = () => sent[sent.length - 1] || "";
/** Lời nhắc hệ thống trong thân yêu cầu (OpenAI: messages[0]; Anthropic: system; Gemini: systemInstruction) */
const systemOf = (raw) => {
  try {
    const j = JSON.parse(raw);
    if (j.system) return JSON.stringify(j.system);
    if (j.systemInstruction) return JSON.stringify(j.systemInstruction);
    return String(j.messages?.[0]?.content ?? "");
  } catch {
    return "";
  }
};
const sendAndSystem = async (text) => {
  const n = sent.length;
  await send(text, { timeout: 60000 });
  return sent.length > n ? systemOf(sent[n]) : null;
};

try {
  await openSettings();
  ok("Bản cục bộ có mục Bộ nhớ trong Cài đặt", !!document.querySelector("[data-memory-card]"));
  const before = memTexts();
  out.soMucTruoc = before.length;

  // 1. Thêm mục; mục giống CCCD bị chặn
  const r1 = await addMem(A);
  ok("Thêm mục A qua Cài đặt", memTexts().includes(A), r1);
  const r2 = await addMem("E2-thu: CCCD 001085012345");
  ok("Mục có số CCCD bị chặn, không lưu", /Không nhớ được/.test(r2) && !memTexts().some((t) => /CCCD/.test(t)), r2);

  // 2. Hội thoại mới → lời nhắc có mục A; lượt sau cùng hội thoại: lời nhắc giống hệt
  await toChat();
  btn(/^Hội thoại mới/)?.click();
  await sleep(500);
  out.nguon = label(btn(/^Model đang dùng/) || document.body);
  const s1 = await sendAndSystem("Xin chào (bài thử E2, lượt 1)");
  ok("Hội thoại mới: lời nhắc hệ thống có mục A", s1 && s1.includes(A.replace("·", "·")), s1 ? `dài ${s1.length}` : "không thấy yêu cầu gửi đi");
  ok("Dòng 'Gửi kèm … mục bộ nhớ' hiện trên khung", !!document.querySelector("[data-mem-used]"), document.querySelector("[data-mem-used]")?.innerText);

  // 3. Thêm mục B giữa hội thoại → lượt sau trong CÙNG hội thoại không đổi lời nhắc (đóng băng)
  await addMem(B);
  await toChat();
  const s2 = await sendAndSystem("Xin chào (bài thử E2, lượt 2)");
  ok("Cùng hội thoại sau khi thêm mục B: lời nhắc giống hệt lượt 1 (bộ nhớ đệm giữ nguyên)", s2 !== null && s2 === s1, s2 === s1 ? "giống hệt" : `khác (có B: ${s2?.includes(B)})`);

  // 4. Hội thoại mới → có cả B
  btn(/^Hội thoại mới/)?.click();
  await sleep(500);
  const s3 = await sendAndSystem("Xin chào (bài thử E2, lượt 3)");
  ok("Hội thoại mới tiếp theo: lời nhắc có cả A và B", !!s3 && s3.includes(A) && s3.includes(B));

  // 5. Đề xuất nhớ của AI (trả lời giả kiểu OpenAI): thẻ Nhớ / Bỏ; bấm Nhớ thì lưu
  if (/"messages"/.test(lastSent())) {
    fakeRemember = true;
    btn(/^Hội thoại mới/)?.click();
    await sleep(500);
    await send("Từ giờ công văn luôn ký thừa lệnh Giám đốc nhé (bài thử E2)", { timeout: 60000 });
    fakeRemember = false;
    const card = document.querySelector('[data-mem-proposal="open"]');
    ok("Thẻ 'AI muốn nhớ' hiện, chưa tự lưu", !!card && !memTexts().includes(P), card?.innerText.replace(/\n+/g, " | "));
    card?.querySelector("[data-mem-save]")?.click();
    await sleep(500);
    ok("Bấm Nhớ: thẻ đổi trạng thái đã nhớ", !!document.querySelector('[data-mem-proposal="saved"]'), document.querySelector('[data-mem-proposal="saved"]')?.innerText.replace(/\n+/g, " | "));
    await openSettings();
    ok("Mục AI đề xuất có trong Cài đặt, nguồn 'AI đề xuất'", [...document.querySelectorAll("[data-memory-item]")].some((el) => el.innerText.includes(P) && /AI đề xuất/.test(el.innerText)));
  } else {
    ok("Đề xuất nhớ của AI (nguồn đang chọn không nói kiểu OpenAI nên không giả được trả lời)", null);
  }
} catch (e) {
  out.loi = String(e && e.stack || e);
} finally {
  window.fetch = realFetch;
}
// window.__e2GiuLai = true trước khi chạy: dừng ở đây để chụp khung (thẻ đề xuất), rồi chạy dot-12e-don.js để dọn
if (window.__e2GiuLai) return out;
{
  // Dọn: xóa mọi mục "E2-thu" qua giao diện
  await openSettings();
  for (let i = 0; i < 10; i++) {
    const el = [...document.querySelectorAll("[data-memory-item]")].find((x) => x.innerText.startsWith("E2-thu"));
    if (!el) break;
    el.querySelector("[data-memory-delete]").click();
    await sleep(400);
  }
  ok("Xóa mục thử qua nút Xóa: bộ nhớ về đúng như trước bài thử", memTexts().length === out.soMucTruoc && !memTexts().some((t) => t.startsWith("E2-thu")), memTexts().length);
  await toChat();
  btn(/^Hội thoại mới/)?.click();
}
out.soLanChan = sent.length;
return out;
