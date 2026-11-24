// Chạy A/B định dạng kết quả đọc Word. A = bản phát hành 4.11.8 (JSON), B = bản cục bộ nhánh claude/ab-dinh-dang-gon (dòng gọn).
// node chay-ab.mjs <thư mục ra> [số lần mỗi ô, mặc định 2] [lọc bài]
import fs from "node:fs"; import path from "node:path"; import { spawn, spawnSync } from "node:child_process";
const W = "D:/ADD IN/.claude/worktrees/multi-agent-orchestration-eval-53eeac";
const HERE = path.dirname(new URL(import.meta.url).pathname).replace(/^\/([A-Z]:)/, "$1");
const OUT = process.argv[2]; const LAN = Number(process.argv[3] || 2); const LOC = process.argv[4] ? new RegExp(process.argv[4]) : null;
fs.mkdirSync(OUT, { recursive: true });
const { CAU_HOI } = await import("file:///" + W + "/kiem-thu/hop-dong-thu.mjs");
const log = (s) => { const l = `[${new Date().toLocaleTimeString("vi-VN")}] ${s}`; console.log(l); fs.appendFileSync(path.join(OUT, "nhat-ky.txt"), l + "\n"); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const cdp = (...a) => spawnSync("node", [W + "/kiem-thu/cdp.mjs", ...a], { cwd: W, encoding: "utf8", timeout: 700000 });
const evalJs = (files) => { const r = cdp("eval", "9229", ...files); try { return JSON.parse(r.stdout); } catch { return { LOI: (r.stdout + r.stderr).slice(-800) }; } };

const BAI = {
  H1: { tep: "aidn-thu-soat.docx", yc: "Chỉ đọc, không sửa gì. Trả lời ngắn, mỗi ý ghi số điều khoản và số đoạn dạng [¶số]: 1) Thời hạn thuê; 2) Đơn giá Phí Dịch Vụ; 3) Mức lãi chậm trả mỗi ngày; 4) Bên Thuê được chấm dứt hợp đồng trước hạn từ khi nào và phí chấm dứt bao nhiêu; 5) Mức trách nhiệm bảo hiểm cháy nổ tối thiểu Bên Thuê phải mua." },
  H2: { tep: "aidn-thu-hd50.docx", yc: "Chỉ đọc, không sửa gì. Trả lời lần lượt 10 câu, đánh số 1 đến 10, mỗi câu ghi rõ số điều khoản: " + CAU_HOI.map((q, k) => `${k + 1}) ${q[0]}`).join(" ") },
  R1: { tep: "aidn-thu-soat.docx", yc: "Rà soát Điều 3 và Điều 5 cho bên thuê, ghi nhận xét vào đúng chỗ có rủi ro. Không sửa chữ." },
  S1: { tep: "aidn-thu-soat.docx", yc: "Ở khoản 4.3, sửa mức tăng tối đa mỗi lần từ 10% thành 7%. Không sửa chỗ nào khác." },
  S2: { tep: "aidn-thu-hd50.docx", yc: "Ở điểm 18.2(b), đổi 72 giờ thành 48 giờ. Không sửa chỗ nào khác." },
};
async function moWord(tep) {
  const r = spawnSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", W + "/kiem-thu/mo-office.ps1", "-Tep", tep], { encoding: "utf8", timeout: 300000 });
  if (r.status !== 0) throw new Error("mo-office: " + (r.stdout + r.stderr).trim());
  for (let i = 0; i < 60; i++) { const l = cdp("list", "9229"); if (/taskpane\.html/.test(l.stdout)) return; await sleep(2000); }
  throw new Error("Không thấy khung add-in");
}
async function chayMot(bai, arm, lan) {
  const b = BAI[bai]; const ten = `${bai}-${arm}-${lan}`;
  const tep = path.join(OUT, `aidn-thu-${ten}.docx`); fs.copyFileSync(path.join(HERE, b.tep), tep);
  await moWord(tep); await sleep(8000);
  let sv = null, svLog = "";
  if (arm === "B") {
    sv = spawn("node", [W + "/kiem-thu/cdp.mjs", "serve", "9229", W + "/dist"], { cwd: W });
    sv.stdout.on("data", (d) => (svLog += d)); sv.stderr.on("data", (d) => (svLog += d));
    for (let i = 0; i < 60 && !/Đang trả bản cục bộ/.test(svLog); i++) await sleep(1000);
    await sleep(25000);
  }
  const dau = evalJs([HERE + "/trang-thai.js"]);
  const chon = path.join(OUT, "che-do-paid.js"); fs.writeFileSync(chon, "const CHE_DO = \"paid\";");
  const cd = evalJs([W + "/kiem-thu/cdp-lib.js", chon, W + "/kiem-thu/cdp-bai/dot-9-chon.js"]);
  const moi = evalJs([W + "/kiem-thu/cdp-lib.js", W + "/kiem-thu/cdp-bai/dot-10-moi.js"]);
  const vars = path.join(OUT, `bien-${ten}.js`);
  fs.writeFileSync(vars, `const NGUON = "openrouter"; const MODEL_RE = new RegExp(${JSON.stringify(process.env.MODEL_RE)}); const YEU_CAU = ${JSON.stringify(b.yc)};\n`);
  const t0 = Date.now();
  const kq = evalJs([W + "/kiem-thu/cdp-lib.js", vars, W + "/kiem-thu/cdp-bai/dot-9-chay.js"]);
  const sau = evalJs([HERE + "/trang-thai.js"]);
  if (sv) sv.kill();
    if (process.env.CUOI === `${bai}-${arm}-${lan}`) { const f = path.join(OUT, "che-do-free.js"); fs.writeFileSync(f, "const CHE_DO = \"free\";"); log("trả về Miễn phí: " + JSON.stringify(evalJs([W + "/kiem-thu/cdp-lib.js", f, W + "/kiem-thu/cdp-bai/dot-9-chon.js"]))); }
  const ghi = { bai, arm, lan, cheDo: cd, yeuCau: b.yc, giay: Math.round((Date.now() - t0) / 1000), phienBan: moi, banCucBo: arm === "B" ? svLog.slice(0, 400) : null, kq, dau, sau };
  fs.writeFileSync(path.join(OUT, `${ten}.json`), JSON.stringify(ghi, null, 1));
  const tk = (kq.khung || "").match(/(\d+) lần gọi AI · Token: vào ([\d.]+)/);
  log(`${ten}: ${kq.LOI ? "LỖI " + kq.LOI.slice(0, 200) : `${tk ? tk[1] + " lần gọi · vào " + tk[2] : "không thấy dòng token"} · ${ghi.giay}s · ${(kq.model?.daChon || kq.model?.loi || "").slice(-60)}`}`);
}
const ARMS = (process.env.ARMS || "A,B").split(",");
const ds = Object.keys(BAI).filter((k) => !LOC || LOC.test(k));
const CHI = process.argv[5] ? process.argv[5].split(",") : null;
for (let lan = 1; lan <= LAN; lan++) for (const bai of ds) for (const arm of (lan % 2 ? ["A", "B"] : ["B", "A"]).filter((a) => ARMS.includes(a))) {
  if (CHI && !CHI.includes(`${bai}-${arm}-${lan}`)) continue;
  try { await chayMot(bai, arm, lan); } catch (e) { log(`${bai}-${arm}-${lan}: DỪNG ${e.message}`); if (/DUNG/.test(e.message)) process.exit(2); }
}
spawnSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", W + "/kiem-thu/mo-office.ps1", "-Dong"], { encoding: "utf8" });
log("XONG");
