// Đợt 1A: nhớ nguồn đang chọn (để trả lại khi dọn) rồi chọn Anthropic · Claude Haiku 5.5 cho các bài Chạy thật.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1a-chon-model.js
await toChat();
const version = (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0];
btn(/^Model đang dùng/).click();
await sleep(700);
const sel = document.querySelector('select[aria-label="Nhà cung cấp"]');
if (!window.__aidnNguonCu) window.__aidnNguonCu = sel.value;
document.body.click();
await sleep(300);
const r = await pick("anthropic", /haiku-5-5/);
document.body.click();
await sleep(300);
return { phienBan: version, nguonCu: window.__aidnNguonCu, ...r };
