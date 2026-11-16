// Thăm dò Đợt 3 (B1): Word, Excel, PowerPoint có dùng chung kho lưu của add-in không, và khung ở ứng dụng này có gửi được lệnh cho khung
// ở ứng dụng kia không (sự kiện storage, BroadcastChannel), khung ở ứng dụng nền có ghi được vào tài liệu không, hẹn giờ có bị bóp không.
// CHỈ ghi / đọc các khóa thử riêng bắt đầu bằng "aidn.thu.dot3." rồi xóa; không đọc khóa nào khác của kho.
// Ba pha, chọn bằng window.__aidnPha đặt ở tệp ghép phía trước:
//   "nghe"  (chạy ở khung NHẬN, ví dụ 9229/Excel): gắn bộ nghe; nhận lệnh thì ghi một ô (Excel) / một đoạn (Word) / một slide (PowerPoint) rồi trả lời
//   "gui"   (chạy ở khung GỬI): ghi khóa thử, gửi lệnh qua storage và BroadcastChannel, chờ trả lời, đo độ trễ
//   "don"   (chạy ở mọi khung): gỡ bộ nghe, xóa khóa thử, xóa thứ đã ghi thử
const pha = window.__aidnPha;
delete window.__aidnPha;
const P = "aidn.thu.dot3.";
const host = Office.context.host;

if (pha === "nghe") {
  const st = (window.__aidnThu = { storage: [], bc: [], ticks: [], viec: [], host });
  st.timer = setInterval(() => st.ticks.push(Date.now()), 1000);
  const lam = async (via, msg) => {
    const t1 = Date.now();
    let kq = "";
    try {
      if (host === "Excel")
        await Excel.run(async (ctx) => {
          let sh = ctx.workbook.worksheets.getItemOrNullObject("aidn-thu-kenh");
          await ctx.sync();
          if (sh.isNullObject) sh = ctx.workbook.worksheets.add("aidn-thu-kenh");
          const r = sh.getRange("A1");
          r.values = [[`ghi theo lệnh ${msg.id} qua ${via}`]];
          r.load("values");
          await ctx.sync();
          kq = r.values[0][0];
        });
      else if (host === "Word")
        await Word.run(async (ctx) => {
          const p = ctx.document.body.insertParagraph(`aidn-thu-kenh ghi theo lệnh ${msg.id} qua ${via}`, "End");
          p.load("text");
          await ctx.sync();
          kq = p.text;
        });
      else
        await PowerPoint.run(async (ctx) => {
          const sl = ctx.presentation.slides.getItemAt(0);
          const tb = sl.shapes.addTextBox(`aidn-thu-kenh ${msg.id} qua ${via}`, { left: 20, top: 20, width: 300, height: 30 });
          tb.name = "aidn-thu-kenh";
          tb.textFrame.textRange.load("text");
          await ctx.sync();
          kq = tb.textFrame.textRange.text;
        });
    } catch (e) {
      kq = "LỖI " + ((e && e.message) || e);
    }
    const rep = { id: msg.id, via, host, kq, nhanLuc: t1, xongLuc: Date.now(), an: document.hidden, tieuDiem: document.hasFocus() };
    st.viec.push(rep);
    try {
      localStorage.setItem(P + "tra." + msg.id + "." + via, JSON.stringify(rep));
    } catch (e) {
      /* bỏ qua */
    }
    st.chan && st.chan.postMessage({ tra: rep });
  };
  st.onStorage = (e) => {
    if (!e.key || !e.key.startsWith(P + "lenh.")) return;
    st.storage.push({ key: e.key, luc: Date.now() });
    if (e.newValue) lam("storage", JSON.parse(e.newValue));
  };
  window.addEventListener("storage", st.onStorage);
  try {
    st.chan = new BroadcastChannel(P + "kenh");
    st.chan.onmessage = (e) => {
      if (!e.data || !e.data.lenh) return;
      st.bc.push({ luc: Date.now() });
      lam("broadcast", e.data.lenh);
    };
  } catch (e) {
    st.bcLoi = String(e);
  }
  return { host, dangNghe: true, coBroadcastChannel: typeof BroadcastChannel !== "undefined", an: document.hidden, tieuDiem: document.hasFocus() };
}

if (pha === "gui") {
  const id = Math.random().toString(36).slice(2, 8);
  const out = { host, id, an: document.hidden, tieuDiem: document.hasFocus() };
  const tra = {};
  const onSt = (e) => {
    if (e.key && e.key.startsWith(P + "tra." + id) && e.newValue) {
      const r = JSON.parse(e.newValue);
      tra[r.host + "/" + r.via] = { ...r, veLuc: Date.now(), qua: "storage" };
    }
  };
  window.addEventListener("storage", onSt);
  const chan = new BroadcastChannel(P + "kenh");
  const traBc = {};
  chan.onmessage = (e) => {
    if (e.data && e.data.tra && e.data.tra.id === id) traBc[e.data.tra.host + "/" + e.data.tra.via] = Date.now();
  };
  const t0 = Date.now();
  localStorage.setItem(P + "lenh." + id, JSON.stringify({ id, t0, from: host }));
  await new Promise((r) => setTimeout(r, 4000));
  const t1 = Date.now();
  chan.postMessage({ lenh: { id: id + "b", t0: t1, from: host } });
  const onSt2 = (e) => {
    if (e.key && e.key.startsWith(P + "tra." + id + "b") && e.newValue) {
      const r = JSON.parse(e.newValue);
      tra[r.host + "/" + r.via] = { ...r, veLuc: Date.now(), qua: "storage" };
    }
  };
  window.addEventListener("storage", onSt2);
  await new Promise((r) => setTimeout(r, 4000));
  window.removeEventListener("storage", onSt);
  window.removeEventListener("storage", onSt2);
  chan.close();
  out.traLoi = Object.entries(tra).map(([k, r]) => ({ khung: k, ketQua: r.kq, treLenhMs: r.nhanLuc - (r.via === "storage" ? t0 : t1), treGhiMs: r.xongLuc - r.nhanLuc, treVeMs: r.veLuc - r.xongLuc, khungNhanAn: r.an, khungNhanCoTieuDiem: r.tieuDiem }));
  out.traLoiQuaBroadcast = Object.keys(traBc);
  // Đọc lại đúng khóa thử của mình (chứng minh kho dùng chung thì xem ở khung nhận: nó nhận được sự kiện là đã thấy khóa)
  out.khoaThuConTrongKho = Object.keys(localStorage).filter((k) => k.startsWith(P)).length;
  return out;
}

if (pha === "don") {
  const st = window.__aidnThu;
  const out = { host };
  if (st) {
    clearInterval(st.timer);
    window.removeEventListener("storage", st.onStorage);
    st.chan && st.chan.close();
    const d = st.ticks.slice(1).map((t, i) => t - st.ticks[i]);
    out.nhipHenGio = { soNhip: st.ticks.length, lonNhatMs: d.length ? Math.max(...d) : null, trungBinhMs: d.length ? Math.round(d.reduce((a, b) => a + b, 0) / d.length) : null };
    out.suKienStorage = st.storage.length;
    out.tinBroadcast = st.bc.length;
    out.viecDaLam = st.viec.map((v) => `${v.via}: ${v.kq} (${v.xongLuc - v.nhanLuc} ms; ẩn=${v.an}; tiêu điểm=${v.tieuDiem})`);
    delete window.__aidnThu;
  }
  const keys = Object.keys(localStorage).filter((k) => k.startsWith(P));
  keys.forEach((k) => localStorage.removeItem(k));
  out.daXoaKhoaThu = keys.length;
  try {
    if (host === "Excel")
      await Excel.run(async (ctx) => {
        const sh = ctx.workbook.worksheets.getItemOrNullObject("aidn-thu-kenh");
        await ctx.sync();
        if (!sh.isNullObject) sh.delete();
        await ctx.sync();
      });
    else if (host === "Word")
      await Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text");
        await ctx.sync();
        ps.items.filter((p) => p.text.startsWith("aidn-thu-kenh")).forEach((p) => p.delete());
        await ctx.sync();
      });
    else
      await PowerPoint.run(async (ctx) => {
        const shapes = ctx.presentation.slides.getItemAt(0).shapes;
        shapes.load("items/name");
        await ctx.sync();
        shapes.items.filter((s) => s.name === "aidn-thu-kenh").forEach((s) => s.delete());
        await ctx.sync();
      });
    out.daDonTaiLieu = true;
  } catch (e) {
    out.donTaiLieuLoi = (e && e.message) || String(e);
  }
  return out;
}
return { loi: "Thiếu window.__aidnPha" };
