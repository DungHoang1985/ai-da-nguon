// Điều khiển khung add-in AI Đa Nguồn đang chạy trong Office qua cổng gỡ lỗi của WebView2 (Chrome DevTools Protocol).
// Dùng khi không điều khiển được màn hình: mở Office bằng lệnh, khung tự mở theo tài liệu, rồi bấm nút / đọc kết quả bằng mã.
//
//   node kiem-thu/cdp.mjs make <word|excel|powerpoint> <tệp ra>      tạo tài liệu tự mở khung add-in (theo Id trong manifest.xml)
//   node kiem-thu/cdp.mjs eval <cổng> <tệp .js> [tệp .js…]           chạy đoạn mã trong khung, in kết quả JSON (các tệp ghép lại thành thân một hàm async;
//                                                                    thường là kiem-thu/cdp-lib.js rồi tới tệp bài thử)
//   node kiem-thu/cdp.mjs shot <cổng> <tệp .png>                     chụp khung add-in
//   node kiem-thu/cdp.mjs file <cổng> <bộ chọn CSS> <tệp> [tệp…]     đặt tệp vào ô chọn tệp (đính kèm, nhập kỹ năng)
// <cổng> có thể kèm tên ứng dụng khi mở nhiều ứng dụng: 9229/Word, 9229/Excel, 9229/PowerPoint
//
// Mở Office kèm cổng gỡ lỗi (PowerShell), khi KHÔNG còn ứng dụng Office nào đang chạy:
//   $env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=9229 --remote-allow-origins=*"
//   Start-Process winword.exe -ArgumentList '"<tệp ra>"'
// KHÔNG đọc hay in kho lưu của khung (localStorage, IndexedDB): ở đó có API key thật.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const [, , cmd, a1, a2] = process.argv;
/** eval nhận nhiều tệp: ghép lại theo thứ tự (tệp hàm dùng chung trước, tệp bài thử sau) */
const files = process.argv.slice(4);

const TEMPLATES = {
  word: ["WordDocumentWithTaskPane.docx", "word/webextensions/webextension.xml"],
  excel: ["ExcelWorkbookWithTaskPane.xlsx", "xl/webextensions/webextension.xml"],
  powerpoint: ["PowerPointPresentationWithTaskPane.pptx", "ppt/webextensions/webextension.xml"],
};

async function make(app, out) {
  const t = TEMPLATES[app];
  if (!t) throw new Error("Ứng dụng phải là word, excel hoặc powerpoint");
  const JSZip = (await import("jszip")).default;
  const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
  const id = /<Id>([^<]+)<\/Id>/.exec(manifest)[1];
  const version = /<Version>([^<]+)<\/Version>/.exec(manifest)[1];
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(root, "node_modules/office-addin-dev-settings/templates", t[0])));
  const xml = (await zip.file(t[1]).async("string")).replace(/00000000-0000-0000-0000-000000000000/g, id).replace(/1\.0\.0\.0/g, version);
  zip.file(t[1], xml);
  fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }));
  console.log(`Đã tạo ${out} (add-in ${id}, bản manifest ${version})`);
}

/** port có thể kèm tên ứng dụng khi nhiều ứng dụng Office cùng mở: "9229/Excel", "9229/PowerPoint", "9229/Word" */
async function connect(portAndHost) {
  const [port, host] = String(portAndHost).split("/");
  const list = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  const target = list.find((x) => x.type === "page" && /taskpane\.html/.test(x.url) && (!host || x.url.includes(`_host_Info=${host}`)));
  if (!target) throw new Error("Không thấy khung add-in. Các trang đang có: " + list.map((x) => `${x.type} ${x.url}`).join(" | "));
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((res, rej) => {
    ws.onopen = res;
    ws.onerror = () => rej(new Error("Không nối được WebSocket tới khung add-in"));
  });
  let seq = 0;
  const waiting = new Map();
  ws.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.id && waiting.has(m.id)) {
      const { res, rej } = waiting.get(m.id);
      waiting.delete(m.id);
      m.error ? rej(new Error(m.error.message)) : res(m.result);
    }
  };
  const send = (method, params = {}) =>
    new Promise((res, rej) => {
      const id = ++seq;
      waiting.set(id, { res, rej });
      ws.send(JSON.stringify({ id, method, params }));
    });
  return { send, close: () => ws.close(), url: target.url };
}

async function evalFile(port, list) {
  const body = list.map((f) => fs.readFileSync(f, "utf8")).join("\n");
  const c = await connect(port);
  try {
    const r = await c.send("Runtime.evaluate", {
      expression: `(async () => { ${body}\n})()`,
      awaitPromise: true,
      returnByValue: true,
      timeout: 600000,
    });
    if (r.exceptionDetails) {
      console.log(JSON.stringify({ LOI: r.exceptionDetails.exception?.description || r.exceptionDetails.text }, null, 1));
      process.exitCode = 1;
    } else console.log(typeof r.result.value === "string" ? r.result.value : JSON.stringify(r.result.value, null, 1));
  } finally {
    c.close();
  }
}

async function shot(port, out) {
  const c = await connect(port);
  try {
    const r = await c.send("Page.captureScreenshot", { format: "png" });
    fs.writeFileSync(out, Buffer.from(r.data, "base64"));
    console.log(`Đã chụp ${out}`);
  } finally {
    c.close();
  }
}

/** Đặt tệp cho một ô chọn tệp (input type=file) trong khung, thay cho hộp thoại chọn tệp của Windows */
async function setFile(port, selector, filePaths) {
  const c = await connect(port);
  try {
    const r = await c.send("Runtime.evaluate", { expression: `document.querySelector(${JSON.stringify(selector)})` });
    if (!r.result.objectId) throw new Error(`Không thấy phần tử ${selector}`);
    await c.send("DOM.setFileInputFiles", { files: filePaths.map((f) => path.resolve(f)), objectId: r.result.objectId });
    console.log(`Đã đặt ${filePaths.length} tệp vào ${selector}`);
  } finally {
    c.close();
  }
}

try {
  if (cmd === "make") await make(a1, a2);
  else if (cmd === "file") await setFile(a1, a2, process.argv.slice(5));
  else if (cmd === "eval") await evalFile(a1, files);
  else if (cmd === "shot") await shot(a1, a2);
  else console.log("Cách dùng: xem chú thích đầu tệp kiem-thu/cdp.mjs");
} catch (e) {
  console.error("LỖI:", e.message);
  process.exitCode = 1;
}
