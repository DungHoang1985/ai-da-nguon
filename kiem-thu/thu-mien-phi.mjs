// Chạy loạt bài thử "model miễn phí soạn / sửa văn bản Word" qua cổng gỡ lỗi (Word thử phải đang mở, xem mo-office.ps1).
//   node kiem-thu/thu-mien-phi.mjs <tệp đề .json> <thư mục ra> <đề,đề|all> <model,model> [cổng=9229]
// Model: khóa trong MODELS dưới đây. Mỗi bài ghi <đề>__<model>.json (bỏ trường docx) và <đề>__<model>.docx.
// Đề: { "<khóa>": { "de": "…", "seed"?: [đoạn…], "nhac"?: "…" } }
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
export const MODELS = {
  "nemotron-ultra": ["openrouter", "nemotron-3-ultra-550b"],
  dots3note: ["openrouter", "dots-3-note"],
  "cohere-aplus": ["cohere", "command-a-plus"],
  "gemini-flash": ["gemini", "flash-latest"],
  "or-free": ["openrouter", "^openrouter/free$"],
};
const [, , deFile, outDir, deList, modelList, port = "9229"] = process.argv;
const DE = JSON.parse(fs.readFileSync(deFile, "utf8"));
fs.mkdirSync(outDir, { recursive: true });
const keys = deList === "all" ? Object.keys(DE) : deList.split(",");
for (const m of modelList.split(",")) {
  for (const key of keys) {
    const [provider, modelRe] = MODELS[m] || [];
    if (!provider) throw new Error("Không có model " + m);
    const bai = { ...DE[key], provider, modelRe };
    const pre = path.join(outDir, `.bai-${key}-${m}.js`);
    fs.writeFileSync(pre, `const BAI = ${JSON.stringify(bai)};\n`);
    const t = Date.now();
    let raw;
    try {
      raw = execFileSync("node", [path.join(here, "cdp.mjs"), "eval", port, path.join(here, "cdp-lib.js"), pre, path.join(here, "cdp-bai/thu-mien-phi-van-ban.js")], {
        maxBuffer: 200 * 1024 * 1024,
        encoding: "utf8",
        timeout: 660000,
      });
    } catch (e) {
      raw = JSON.stringify({ LOI: String(e.message).slice(0, 500), stdout: String(e.stdout || "").slice(0, 2000) });
    }
    fs.rmSync(pre, { force: true });
    let j;
    try {
      j = JSON.parse(raw);
    } catch {
      j = { LOI: "không đọc được kết quả", raw: raw.slice(0, 2000) };
    }
    const base = path.join(outDir, `${key}__${m}`);
    if (j.docx) fs.writeFileSync(base + ".docx", Buffer.from(j.docx, "base64"));
    delete j.docx;
    fs.writeFileSync(base + ".json", JSON.stringify(j, null, 1));
    const s = Math.round((Date.now() - t) / 1000);
    console.log(`${key} × ${m}: ${s}s · ${j.LOI ? "LỖI " + j.LOI.slice(0, 120) : `Word ${j.luot2?.chuWord ?? j.luot1?.chuWord} chữ · đổi nguồn ${j.doiNguon?.length || 0}`}`);
  }
}
