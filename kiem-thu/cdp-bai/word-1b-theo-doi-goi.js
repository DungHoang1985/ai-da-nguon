// Đợt 1B: theo dõi các yêu cầu add-in gửi tới Anthropic trong khung (để chứng minh tham số suy nghĩ, tiêu đề beta, số khối suy nghĩ
// và số tệp PDF được gửi lại). CHỈ ghi các cờ và số đếm; KHÔNG ghi hay in tiêu đề chứa key, không ghi nội dung tin nhắn.
// Đặt window.__aidnXemGoi = true ở tệp ghép phía trước để đọc kết quả và gỡ theo dõi.
if (window.__aidnXemGoi) {
  const out = window.__aidnGoi || [];
  if (window.__aidnFetchGoc) {
    window.fetch = window.__aidnFetchGoc;
    delete window.__aidnFetchGoc;
  }
  delete window.__aidnGoi;
  return out;
}
if (!window.__aidnFetchGoc) {
  window.__aidnFetchGoc = window.fetch;
  window.__aidnGoi = [];
  window.fetch = async (u, init) => {
    const r = await window.__aidnFetchGoc(u, init);
    try {
      if (String(u).includes("api.anthropic.com/v1/messages")) {
        const b = JSON.parse(init.body);
        const kinds = b.messages.flatMap((m) => (Array.isArray(m.content) ? m.content.map((c) => c.type) : ["text"]));
        const n = (k) => kinds.filter((x) => x === k).length;
        window.__aidnGoi.push({
          trangThai: r.status,
          thinking: b.thinking || null,
          effort: (b.output_config && b.output_config.effort) || null,
          coTieuDeBeta: !!(init.headers && init.headers["anthropic-beta"]),
          soTinNhan: b.messages.length,
          khoiSuyNghiGuiLai: n("thinking") + n("redacted_thinking"),
          khoiTaiLieu: n("document"),
          soCongCu: (b.tools || []).length,
          coLoiNhacHeThong: !!b.system,
        });
      }
    } catch (e) {
      /* không ảnh hưởng yêu cầu thật */
    }
    return r;
  };
}
return "đang theo dõi";
