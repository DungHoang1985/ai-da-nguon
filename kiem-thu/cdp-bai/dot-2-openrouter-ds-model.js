// Đợt 2: liệt kê các model OpenRouter đang hiện trong bảng chọn (theo bộ lọc hiện tại) và nhớ model đang chọn vào window.__aidnModelOrCu
// để bài sau trả lại. Không đổi gì. Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-openrouter-ds-model.js
await toChat();
btn(/^Model đang dùng/).click();
await sleep(900);
const sel = document.querySelector('select[aria-label="Nhà cung cấp"]');
const opts = [...document.querySelectorAll('[role="listbox"][aria-label="Model"] [role="option"]')];
const cur = opts.find((o) => o.getAttribute("aria-selected") === "true");
if (sel.value === "openrouter" && cur && !window.__aidnModelOrCu) window.__aidnModelOrCu = cur.getAttribute("title");
const res = {
  nguon: sel.value,
  dangChon: cur ? cur.getAttribute("title") : null,
  boLoc: [...document.querySelectorAll('[role="switch"], input[type="checkbox"]')].map((s) => `${s.parentElement?.innerText?.trim() || ""}: ${s.checked}`),
  model: opts.map((o) => o.getAttribute("title")),
};
document.body.click();
await sleep(300);
return res;
