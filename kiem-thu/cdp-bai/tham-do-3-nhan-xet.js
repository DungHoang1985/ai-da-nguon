// Thăm dò Đợt 3 (B5): thao tác nào làm phiên Word từ chối chèn nhận xét mới (AccessDenied ở Range.insertComment)?
// Sau MỖI thao tác nghi ngờ, thử chèn rồi gỡ một nhận xét; ghi lại thao tác đầu tiên làm việc chèn bị từ chối.
// Chạy trên văn bản thử TRỐNG, phiên Word vừa mở: node kiem-thu/cdp.mjs eval 9229/Word kiem-thu/cdp-bai/tham-do-3-nhan-xet.js
const out = [];
const thuNhanXet = () =>
  Word.run(async (ctx) => {
    const p = ctx.document.body.paragraphs.getFirst();
    const c = p.getRange("Whole").insertComment("thử");
    c.load("id");
    await ctx.sync();
    c.delete();
    await ctx.sync();
    return "được";
  }).catch((e) => `LỖI ${e.code || ""} ${e.message}`.slice(0, 90));
const buoc = async (ten, fn) => {
  let loi = "";
  try {
    await fn();
  } catch (e) {
    loi = ` (thao tác lỗi: ${e.message.slice(0, 60)})`;
  }
  const kq = await thuNhanXet();
  out.push(`${ten}${loi} → chèn nhận xét: ${kq}`);
  return kq === "được";
};
const run = (fn) => Word.run(fn);
const mode = (m) =>
  run(async (ctx) => {
    ctx.document.changeTrackingMode = m;
    await ctx.sync();
  });

await buoc("0. Trạng thái ban đầu (thêm một đoạn chữ làm chỗ neo)", () =>
  run(async (ctx) => {
    ctx.document.changeTrackingMode = "Off";
    ctx.document.body.insertParagraph("Đoạn neo của bài thăm dò nhận xét.", "Start");
    await ctx.sync();
  })
);
const steps = [
  ["1. Bật theo dõi, chèn một đoạn, tắt theo dõi", async () => {
    await mode("TrackAll");
    await run(async (ctx) => { ctx.document.body.insertParagraph("Đoạn chèn có theo dõi.", "End"); await ctx.sync(); });
    await mode("Off");
  }],
  ["2. Chèn bảng 1×2 lúc đang theo dõi (After đoạn cuối)", async () => {
    await mode("TrackAll");
    await run(async (ctx) => { ctx.document.body.paragraphs.getLast().insertTable(1, 2, "After", [["A", "B"]]); await ctx.sync(); });
    await mode("Off");
  }],
  ["3. Xóa bảng đó khi đã tắt theo dõi (table.delete)", () =>
    run(async (ctx) => { const t = ctx.document.body.tables.getFirst(); t.delete(); await ctx.sync(); })],
  ["4. Chèn bảng lúc đang theo dõi rồi xóa bằng parentTable của đoạn lấy theo định danh, theo dõi tắt", async () => {
    await mode("TrackAll");
    const id = await run(async (ctx) => { const t = ctx.document.body.paragraphs.getLast().insertTable(1, 2, "After", [["C", "D"]]); const p = t.getCell(0, 0).body.paragraphs.getFirst(); p.load("uniqueLocalId"); await ctx.sync(); return p.uniqueLocalId; });
    await mode("Off");
    await run(async (ctx) => { const p = ctx.document.getParagraphByUniqueLocalId(id); p.parentTable.delete(); await ctx.sync(); });
  }],
  ["5. Chèn đoạn có theo dõi rồi xóa đoạn đó theo định danh khi tắt theo dõi", async () => {
    await mode("TrackAll");
    const id = await run(async (ctx) => { const p = ctx.document.body.insertParagraph("Đoạn sẽ bị xóa.", "End"); p.load("uniqueLocalId"); await ctx.sync(); return p.uniqueLocalId; });
    await mode("Off");
    await run(async (ctx) => { ctx.document.getParagraphByUniqueLocalId(id).delete(); await ctx.sync(); });
  }],
  ["6. Chèn bảng TRƯỚC đoạn đầu lúc đang theo dõi (ô nhiều đoạn chèn Before), để nguyên", async () => {
    await mode("TrackAll");
    await run(async (ctx) => { const t = ctx.document.body.paragraphs.getFirst().insertTable(1, 2, "Before", [["E", "F"]]); t.getCell(0, 0).body.paragraphs.getFirst().insertParagraph("E0", "Before"); await ctx.sync(); });
    await mode("Off");
  }],
  ["7. body.clear() khi tắt theo dõi (văn bản đang có sửa đổi chờ duyệt và bảng), rồi thêm lại đoạn neo", () =>
    run(async (ctx) => { ctx.document.body.clear(); await ctx.sync(); ctx.document.body.insertParagraph("Đoạn neo sau khi xóa trắng.", "Start"); await ctx.sync(); })],
  ["8. Chèn bảng sau một ĐOẠN RỖNG lúc đang theo dõi (ca bảng nuốt đoạn rỗng)", async () => {
    await mode("TrackAll");
    await run(async (ctx) => { const e = ctx.document.body.insertParagraph("", "End"); e.insertTable(1, 2, "After", [["G", "H"]]); await ctx.sync(); });
    await mode("Off");
  }],
  ["9. Từ chối mọi sửa đổi chờ duyệt (rejectAll)", () =>
    run(async (ctx) => { ctx.document.body.getTrackedChanges().rejectAll(); await ctx.sync(); })],
];
let dung = "";
for (const [ten, fn] of steps) {
  const ok = await buoc(ten, fn);
  if (!ok && !dung) {
    dung = ten;
    break;
  }
}
return { thaoTacDauTienGayLoi: dung || "không thao tác nào trong danh sách gây lỗi", nhatKy: out };
