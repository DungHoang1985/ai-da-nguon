// Đợt 7, A2 (Chạy thật, Excel): một việc cố ý dài hơn trần số bước → thẻ "Làm tiếp?" → bấm "Dừng ở đây" → thẻ mất, AI không chạy thêm.
// Bản sao của dot-7-lam-tiep.js, chỉ khác phần bấm nút.
// Kèm kiểm chế độ thử: số hội thoại trong lịch sử trước và sau bằng nhau (chỉ đọc ĐỘ DÀI của chỉ mục, không đọc nội dung kho).
// Chạy: node kiem-thu/cdp.mjs thu 9229 on   (tải lại khung để đầu khung hiện dòng "chế độ thử"; nhớ chạy lại serve nếu cần)
//       node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-7-lam-tiep.js
const SHEET = "AIDN_D7";
const soHoiThoai = () =>
  new Promise((resolve) => {
    try {
      const req = indexedDB.open("aidn", 1);
      req.onerror = () => resolve(-1);
      req.onsuccess = () => {
        try {
          const g = req.result.transaction("kv", "readonly").objectStore("kv").get("conv-index");
          g.onsuccess = () => resolve(Array.isArray(g.result) ? g.result.length : 0);
          g.onerror = () => resolve(-1);
        } catch (e) {
          resolve(-1);
        }
      };
    } catch (e) {
      resolve(-1);
    }
  });
const docCot = () =>
  Excel.run(async (ctx) => {
    const r = ctx.workbook.worksheets.getItem(SHEET).getRange("A1:A16");
    r.load("values");
    await ctx.sync();
    return r.values.map((x) => x[0]);
  });
const the = () => {
  const el = document.querySelector('[data-paused="1"]');
  return el ? el.innerText.replace(/\n+/g, " | ") : null;
};
const lanGoi = () => {
  const all = [...document.querySelectorAll("[data-calls]")];
  return all.length ? Number(all[all.length - 1].getAttribute("data-calls")) : null;
};

await toChat();
const cheDoThu = !!document.querySelector('[data-test-mode="1"]');
const truoc = await soHoiThoai();
btn(/^Hội thoại mới/)?.click();
await sleep(400);

// Nhớ nguồn và model đang chọn để trả lại
if (!document.querySelector('select[aria-label="Nhà cung cấp"]')) {
  btn(/^Model đang dùng/).click();
  await sleep(900);
}
const sel = document.querySelector('select[aria-label="Nhà cung cấp"]');
if (!sel) return { loi: "Không mở được bảng chọn model", khung: tail(500) };
const nguonCu = sel.value;
const nhanCu = label(btn(/^Model đang dùng/)).slice(0, 120);
const optCu = document.querySelector('[role="listbox"][aria-label="Model"] [role="option"][aria-selected="true"]');
const modelCu = optCu ? optCu.getAttribute("title") : null;
// Chọn nguồn và model thử ngay trong bảng đang mở (hàm pick sẽ bấm nút lần nữa và làm đóng bảng)
const chonTrongBang = async (providerId, modelRe) => {
  if (!document.querySelector('select[aria-label="Nhà cung cấp"]')) {
    btn(/^Model đang dùng/).click();
    await sleep(900);
  }
  const s = document.querySelector('select[aria-label="Nhà cung cấp"]');
  if (!s) return { loi: "Không mở được bảng chọn model" };
  if (s.value !== providerId) {
    setNative(s, providerId);
    await sleep(1800);
  }
  for (let i = 0; i < 20 && !document.querySelector('[role="option"]'); i++) await sleep(500);
  const opts = [...document.querySelectorAll('[role="listbox"][aria-label="Model"] [role="option"]')];
  const hit = opts.find((o) => modelRe.test(o.getAttribute("title") || ""));
  if (!hit) return { loi: "Không thấy model khớp", coSan: opts.map((o) => o.getAttribute("title")).slice(0, 40) };
  hit.click();
  await sleep(700);
  return { daChon: label(btn(/^Model đang dùng/)) };
};
const chon = await chonTrongBang("anthropic", /haiku-5-5/i);
if (chon.loi) return { loi: chon.loi, coSan: chon.coSan, nguonCu, nhanCu };
await toChat();

await Excel.run(async (ctx) => {
  const old = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
  await ctx.sync();
  if (!old.isNullObject) old.delete();
  ctx.workbook.worksheets.add(SHEET).activate();
  await ctx.sync();
});

const yeuCau =
  `Bài thử của add-in. Trên sheet ${SHEET}, ghi các số từ 1 tới 16 vào các ô A1 tới A16 (A1 = 1, A2 = 2, …). ` +
  "BẮT BUỘC làm theo cách sau để thử giới hạn số bước: mỗi lệnh gọi công cụ chỉ ghi ĐÚNG MỘT ô, và mỗi lần chỉ gọi MỘT lệnh, có kết quả rồi gọi ngay lệnh cho ô kế tiếp. Làm liên tục cho tới khi đủ cả 16 ô rồi mới trả lời; không dừng giữa chừng để hỏi hay báo cáo.";
const r1 = await send(yeuCau, { approve: () => "allow", timeout: 240000, tail: 900 });
const sau1 = { the: the(), lanGoi: lanGoi(), cot: await docCot(), giay: r1.giay };

let sau2 = null;
// Bấm "Dừng ở đây" trên thẻ: thẻ mất, không gọi AI thêm, cột giữ 12 ô, khung không còn bận, nút Hoàn tác còn
const dungO = [...document.querySelectorAll("button")].find((b) => /^Dừng ở đây$/.test(label(b)));
if (dungO) {
  dungO.click();
  await sleep(5000);
  sau2 = {
    the: the(),
    lanGoi: lanGoi(),
    cot: await docCot(),
    conBan: !![...document.querySelectorAll("button")].find((b) => /^Dừng$/.test(label(b))),
    nutHoanTac: label(btn(/^Hoàn tác/) || document.body).slice(0, 40),
  };
}
const khung = tail(700).replace(/\n{2,}/g, "\n");

// Dọn: xóa sheet thử, hội thoại mới, trả nguồn và model
await Excel.run(async (ctx) => {
  ctx.workbook.worksheets.getItem(SHEET).delete();
  await ctx.sync();
}).catch(() => undefined);
btn(/^Hội thoại mới/)?.click();
await sleep(400);
let traLai = null;
if (nguonCu && modelCu) {
  const esc = modelCu.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  traLai = await chonTrongBang(nguonCu, new RegExp("^" + esc + "$"));
} else if (nguonCu) {
  traLai = { loi: `Không đọc được model đang chọn lúc đầu (nhãn: ${nhanCu}); chưa trả lại` };
}
await toChat();
await sleep(1500);
return {
  phienBan: (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0],
  cheDoThuHienODauKhung: cheDoThu,
  hoiThoaiTruoc: truoc,
  hoiThoaiSau: await soHoiThoai(),
  modelThu: chon.daChon,
  sauLuot1: sau1,
  coNutDungODay: !!dungO,
  sauDungODay: sau2,
  khung,
  traLaiModel: traLai && (traLai.daChon || traLai.loi),
};
