// Đo: trang https://dunghoang1985.github.io/ai-da-nguon/ (Edge --headless=new, hồ sơ tạm riêng) gọi điểm cuối http://127.0.0.1
// Máy chủ thử tự dựng (ghi lại yêu cầu dò trước) ở 3 chế độ + cổng language_server của Antigravity (chỉ dò, không token).
import http from "node:http";
import { spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const AGY_PORT = Number(process.argv[2] || 0);
const OUT = process.argv[3];
const EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe";
const PAGE = "https://dunghoang1985.github.io/ai-da-nguon/taskpane.html";
const log = [];

// mode qua đường dẫn: /cors-pna (CORS + Allow-Private-Network), /cors (chỉ CORS), /none (không tiêu đề nào)
const srv = http.createServer((req, res) => {
  const mode = req.url.split("/")[1];
  log.push({ mode, method: req.method, headers: Object.fromEntries(Object.entries(req.headers).filter(([k]) => /origin|access-control|sec-fetch|content-type/i.test(k))) });
  const h = {};
  if (mode !== "none") {
    h["Access-Control-Allow-Origin"] = req.headers.origin || "*";
    h["Access-Control-Allow-Headers"] = "authorization, content-type, anthropic-version, x-api-key";
    h["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS";
  }
  if (mode === "cors-pna") { h["Access-Control-Allow-Private-Network"] = "true"; h["Access-Control-Allow-Local-Network"] = "true"; }
  if (req.method === "OPTIONS") { res.writeHead(204, h); return res.end(); }
  res.writeHead(401, { ...h, "Content-Type": "application/json" });
  res.end('{"error":"thu-401"}');
});
await new Promise((r) => srv.listen(0, "127.0.0.1", r));
const PORT = srv.address().port;

const prof = fs.mkdtempSync(path.join(os.tmpdir(), "edge-do-"));
const DBG = 9300 + Math.floor(Math.random() * 500);
const edge = spawn(EDGE, ["--headless=new", `--user-data-dir=${prof}`, `--remote-debugging-port=${DBG}`, "--no-first-run", "--no-default-browser-check", "about:blank"], { windowsHide: true, stdio: "ignore" });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let wsUrl;
for (let i = 0; i < 40 && !wsUrl; i++) { await sleep(500); try { const t = await (await fetch(`http://127.0.0.1:${DBG}/json/list`)).json(); wsUrl = t.find((x) => x.type === "page")?.webSocketDebuggerUrl; } catch {} }
const ws = new WebSocket(wsUrl);
await new Promise((r) => (ws.onopen = r));
let id = 0; const pend = new Map(); const consoleMsgs = [];
ws.onmessage = (e) => { const m = JSON.parse(e.data); if (m.id && pend.has(m.id)) { pend.get(m.id)(m); pend.delete(m.id); } else if (m.method === "Log.entryAdded") consoleMsgs.push(m.params.entry.text); else if (m.method === "Runtime.consoleAPICalled") consoleMsgs.push(m.params.args.map((a) => a.value).join(" ")); };
const cmd = (method, params = {}) => new Promise((r) => { const i = ++id; pend.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
await cmd("Log.enable"); await cmd("Runtime.enable"); await cmd("Page.enable");
await cmd("Page.navigate", { url: PAGE });
await sleep(6000);
const loc = await cmd("Runtime.evaluate", { expression: "location.href + ' | secure=' + isSecureContext", returnByValue: true });

const targets = [
  ["Máy chủ thử: CORS + cho phép mạng riêng", `http://127.0.0.1:${PORT}/cors-pna/v1/messages`],
  ["Máy chủ thử: chỉ CORS", `http://127.0.0.1:${PORT}/cors/v1/messages`],
  ["Máy chủ thử: không tiêu đề cho phép", `http://127.0.0.1:${PORT}/none/v1/messages`],
  ["Máy chủ thử qua tên localhost: CORS + mạng riêng", `http://localhost:${PORT}/cors-pna/v1/messages`],
];
if (AGY_PORT) targets.push(["Antigravity language_server (cổng cục bộ, không token)", `http://127.0.0.1:${AGY_PORT}/`]);
const results = [];
for (const [ten, url] of targets) {
  const expr = `(async()=>{const t=Date.now();try{const r=await fetch(${JSON.stringify(url)},{method:"POST",headers:{"Content-Type":"application/json","Authorization":"Bearer khong-phai-key","anthropic-version":"2023-06-01"},body:"{}"});const b=(await r.text()).slice(0,160);return {ketQua:"doc-duoc",ma:r.status,ms:Date.now()-t,traLoi:b};}catch(e){return {ketQua:"bi-chan",ms:Date.now()-t,loi:String(e)};}})()`;
  const r = await cmd("Runtime.evaluate", { expression: expr, awaitPromise: true, returnByValue: true });
  results.push({ ten, url: url.replace(String(PORT), "<cong-thu>"), ...r.result.result.value });
}
const ver = await (await fetch(`http://127.0.0.1:${DBG}/json/version`)).json();
ws.close(); edge.kill(); srv.close();
await sleep(1000);
try { fs.rmSync(prof, { recursive: true, force: true }); } catch {}
const out = { ngay: new Date().toISOString(), trinhDuyet: ver.Browser, cheDo: "Edge --headless=new, hồ sơ tạm riêng, không cửa sổ", trang: loc.result.result.value, ketQua: results, yeuCauMayChuThuNhan: log, thongBaoTrinhDuyet: consoleMsgs.filter((t) => /private|local|network|CORS|blocked|mixed/i.test(t)).slice(0, 20) };
fs.writeFileSync(OUT, JSON.stringify(out, null, 1));
console.log(JSON.stringify(out, null, 1));
