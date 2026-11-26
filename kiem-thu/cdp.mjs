// Điều khiển khung add-in AI Đa Nguồn đang chạy trong Office qua cổng gỡ lỗi của WebView2 (Chrome DevTools Protocol).
// Dùng khi không điều khiển được màn hình: mở Office bằng lệnh, khung tự mở theo tài liệu, rồi bấm nút / đọc kết quả bằng mã.
//
//   node kiem-thu/cdp.mjs make <word|excel|powerpoint> <tệp ra>      tạo tài liệu tự mở khung add-in (theo Id trong manifest.xml)
//   node kiem-thu/cdp.mjs eval <cổng> <tệp .js> [tệp .js…]           chạy đoạn mã trong khung, in kết quả JSON (các tệp ghép lại thành thân một hàm async;
//                                                                    thường là kiem-thu/cdp-lib.js rồi tới tệp bài thử)
//   node kiem-thu/cdp.mjs shot <cổng> <tệp .png>                     chụp khung add-in
//   node kiem-thu/cdp.mjs file <cổng> <bộ chọn CSS> <tệp> [tệp…]     đặt tệp vào ô chọn tệp (đính kèm, nhập kỹ năng)
//   node kiem-thu/cdp.mjs reload <cổng>                              tải lại khung add-in, bỏ qua bộ nhớ tạm (như Ctrl+F5)
//   node kiem-thu/cdp.mjs thu <cổng> [on|off]                        bật / tắt CHẾ ĐỘ THỬ (hội thoại không lưu vào Lịch sử) rồi tải lại khung.
//                                                                    Lệnh eval trên khung add-in tự bật chế độ này; thêm --luu để eval KHÔNG tự bật.
//   node kiem-thu/cdp.mjs list <cổng>                               liệt kê các trang đang mở (khung add-in, hộp thoại đăng nhập); giá trị tham số nhạy cảm chỉ in độ dài
//   node kiem-thu/cdp.mjs nav <cổng>#<đoạn địa chỉ> <địa chỉ>         đưa một trang (ví dụ hộp thoại đăng nhập) tới địa chỉ khác
//   <cổng>#<đoạn địa chỉ> dùng được với eval, shot: chọn trang có địa chỉ chứa đoạn đó thay cho khung add-in (9229#dang-nhap, 9229#openrouter.ai)
//   node kiem-thu/cdp.mjs serve <cổng> [thư mục dist]                nạp BẢN CỤC BỘ vào khung: chặn mọi yêu cầu tới địa chỉ add-in và trả tệp
//                                                                    từ thư mục dist (sau npm run build), rồi tải lại khung. Chạy ở nền, giữ tới khi
//                                                                    thử xong; tắt tiến trình rồi tải lại khung là về bản đã phát hành. Không đẩy gì lên mạng.
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

/** Địa chỉ để in ra: giữ đường dẫn và tên tham số; giá trị chỉ in với tham số không nhạy cảm, còn lại in độ dài.
 *  Không bao giờ in mã đăng nhập (code), state hay mã thách thức ra màn hình hoặc tệp bằng chứng. */
const SHOW_PARAMS = new Set(["callback_url", "code_challenge_method", "key_label", "p"]);
function safeUrl(raw) {
  try {
    const u = new URL(raw);
    const qs = [...u.searchParams.entries()].map(([k, v]) => (SHOW_PARAMS.has(k) ? `${k}=${v}` : `${k}=<${v.length} ký tự>`));
    return u.origin + u.pathname + (qs.length ? "?" + qs.join("&") : "") + (u.hash ? "#<…>" : "");
  } catch {
    return String(raw).slice(0, 80);
  }
}

/** Liệt kê các trang đang mở trên cổng gỡ lỗi (khung add-in, hộp thoại đăng nhập…) */
async function listTargets(port) {
  const list = await (await fetch(`http://127.0.0.1:${String(port).split(/[/#]/)[0]}/json/list`)).json();
  console.log(JSON.stringify(list.filter((x) => x.type === "page" || x.type === "webview").map((x) => ({ loai: x.type, diaChi: safeUrl(x.url) })), null, 1));
}

/** port có thể kèm tên ứng dụng khi nhiều ứng dụng Office cùng mở: "9229/Excel", "9229/PowerPoint", "9229/Word";
 *  hoặc kèm một đoạn của địa chỉ để chọn trang khác khung add-in: "9229#dang-nhap", "9229#openrouter.ai" (hộp thoại đăng nhập) */
async function connect(portAndHost) {
  const [portHost, pick] = String(portAndHost).split("#");
  const [port, host] = portHost.split("/");
  const list = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  const target = pick
    ? list.find((x) => (x.type === "page" || x.type === "webview") && x.url.includes(pick))
    : list.find((x) => x.type === "page" && /taskpane\.html/.test(x.url) && (!host || x.url.includes(`_host_Info=${host}`)));
  if (!target) throw new Error((pick ? `Không thấy trang có "${pick}".` : "Không thấy khung add-in.") + " Các trang đang có: " + list.map((x) => `${x.type} ${safeUrl(x.url)}`).join(" | "));
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((res, rej) => {
    ws.onopen = res;
    ws.onerror = () => rej(new Error("Không nối được WebSocket tới khung add-in"));
  });
  let seq = 0;
  const waiting = new Map();
  const handlers = new Map();
  ws.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.method && handlers.has(m.method)) handlers.get(m.method)(m.params);
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
  return { send, on: (method, fn) => handlers.set(method, fn), close: () => ws.close(), closed: new Promise((r) => (ws.onclose = r)), url: target.url };
}

/** Khóa "chế độ thử" của add-in (src/core/conversations.ts, TEST_MODE_KEY): bật thì hội thoại không được lưu vào Lịch sử */
const TEST_MODE_KEY = "aidn-che-do-thu";
const TEST_MODE_ON = `try { sessionStorage.setItem(${JSON.stringify(TEST_MODE_KEY)}, "1"); } catch (e) {}\n`;

/** Bật / tắt chế độ thử rồi tải lại khung (để đầu khung hiện dòng "chế độ thử"). Cờ tự mất khi đóng Office. */
async function testMode(port, state) {
  const on = state !== "off";
  const c = await connect(port);
  try {
    await c.send("Runtime.evaluate", {
      expression: on ? TEST_MODE_ON : `try { sessionStorage.removeItem(${JSON.stringify(TEST_MODE_KEY)}); } catch (e) {}`,
    });
    await c.send("Page.reload", { ignoreCache: false });
    console.log(on ? "Đã BẬT chế độ thử: hội thoại không được lưu vào Lịch sử" : "Đã TẮT chế độ thử: hội thoại được lưu như thường");
  } finally {
    c.close();
  }
}

async function evalFile(port, list) {
  // Mọi bài thử chạy trong khung add-in đều tự bật chế độ thử (không lưu Lịch sử) — có hiệu lực ngay, không cần tải lại.
  // Bài cần kiểm chính việc lưu lịch sử thì tự gỡ cờ trong mã của bài, hoặc chạy `cdp.mjs thu <cổng> off` trước và thêm tham số --luu.
  const keep = process.argv.includes("--luu");
  list = list.filter((f) => f !== "--luu");
  const body = (String(port).includes("#") || keep ? "" : TEST_MODE_ON) + list.map((f) => fs.readFileSync(f, "utf8")).join("\n");
  const c = await connect(port);
  try {
    // Khung add-in phải vẽ xong mới chạy bài: 23/11 và 26/11/2026 bài trượt vì khung trắng (26/11 kẹt ở readyState "interactive"
    // tới khi tải lại). Chờ khung có chữ; trắng thì tự tải lại một lần (bỏ bộ nhớ tạm); vẫn trắng thì báo lỗi, không chạy bài.
    if (!String(port).includes("#")) {
      const ve = async (ms) => {
        for (const t0 = Date.now(); Date.now() - t0 < ms; await new Promise((r) => setTimeout(r, 1000))) {
          const r = await c.send("Runtime.evaluate", {
            expression: `document.readyState === "complete" && !!document.body && document.body.innerText.trim().length > 0`,
            returnByValue: true,
          }).catch(() => null);
          if (r?.result?.value === true) return true;
        }
        return false;
      };
      if (!(await ve(30000))) {
        console.error("Khung add-in còn trắng sau 30 giây: tải lại (bỏ bộ nhớ tạm) rồi chờ thêm.");
        await c.send("Page.reload", { ignoreCache: true });
        if (!(await ve(45000))) {
          console.log(JSON.stringify({ LOI: "Khung add-in trắng (chưa vẽ xong) cả sau khi tải lại: không chạy bài thử." }, null, 1));
          process.exitCode = 1;
          return;
        }
      }
    }
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

/** Tải lại khung add-in, bỏ qua bộ nhớ tạm (như Ctrl+F5): để lấy bản vừa phát hành */
async function reload(port) {
  const c = await connect(port);
  try {
    await c.send("Network.enable");
    await c.send("Network.setCacheDisabled", { cacheDisabled: true });
    await c.send("Page.reload", { ignoreCache: true });
    console.log("Đã tải lại khung, bỏ qua bộ nhớ tạm");
  } finally {
    c.close();
  }
}

/** Đưa một trang (chọn bằng "#<đoạn địa chỉ>") tới địa chỉ khác — dùng để giả lập bước nhà cung cấp chuyển về trang của add-in */
async function navigate(port, url) {
  const c = await connect(port);
  try {
    await c.send("Page.navigate", { url });
    console.log(`Đã chuyển trang tới ${safeUrl(url)}`);
  } finally {
    c.close();
  }
}

const MIME = { html: "text/html; charset=utf-8", js: "text/javascript; charset=utf-8", mjs: "text/javascript; charset=utf-8", css: "text/css; charset=utf-8", json: "application/json", map: "application/json", png: "image/png", svg: "image/svg+xml", ico: "image/x-icon", wasm: "application/wasm", woff2: "font/woff2", xml: "text/xml", txt: "text/plain; charset=utf-8" };

/** Nạp bản cục bộ: trả tệp trong thư mục dist cho mọi yêu cầu tới địa chỉ add-in (cùng nguồn gốc nên kho lưu, key vẫn như cũ) */
async function serve(port, dir) {
  const base = path.resolve(dir || path.join(root, "dist"));
  if (!fs.existsSync(path.join(base, "taskpane.html"))) throw new Error(`Không thấy ${path.join(base, "taskpane.html")}: chạy npm run build trước.`);
  const c = await connect(port);
  const prefix = c.url.slice(0, c.url.indexOf("taskpane.html"));
  let served = 0;
  c.on("Fetch.requestPaused", async (p) => {
    try {
      const u = new URL(p.request.url);
      const rel = decodeURIComponent((u.origin + u.pathname).slice(prefix.length)) || "taskpane.html";
      const file = path.join(base, rel);
      if (file.startsWith(base) && fs.existsSync(file) && fs.statSync(file).isFile()) {
        const ext = path.extname(file).slice(1).toLowerCase();
        await c.send("Fetch.fulfillRequest", {
          requestId: p.requestId,
          responseCode: 200,
          responseHeaders: [
            { name: "Content-Type", value: MIME[ext] || "application/octet-stream" },
            { name: "Cache-Control", value: "no-store" },
          ],
          body: fs.readFileSync(file).toString("base64"),
        });
        served++;
        console.log(`cục bộ  ${rel}`);
      } else {
        await c.send("Fetch.continueRequest", { requestId: p.requestId });
        console.log(`mạng    ${rel}`);
      }
    } catch (e) {
      console.error("LỖI khi trả tệp:", e.message);
    }
  });
  await c.send("Network.enable");
  await c.send("Network.setCacheDisabled", { cacheDisabled: true });
  await c.send("Fetch.enable", { patterns: [{ urlPattern: prefix + "*" }] });
  await c.send("Page.reload", { ignoreCache: true });
  // Trang mở sau đó (hộp thoại đăng nhập của Office là một trang riêng): tự gắn vào ngay khi nó được tạo, trước khi nó tải gì,
  // để yêu cầu đầu tiên của nó cũng nhận bản cục bộ (dang-nhap.html chưa có trên bản đã phát hành).
  try {
    const [p] = String(port).split(/[/#]/);
    const ver = await (await fetch(`http://127.0.0.1:${p}/json/version`)).json();
    const b = new WebSocket(ver.webSocketDebuggerUrl);
    await new Promise((res, rej) => {
      b.onopen = res;
      b.onerror = () => rej(new Error("không nối được kênh trình duyệt"));
    });
    let n = 0;
    const bsend = (method, params = {}, sessionId) => b.send(JSON.stringify({ id: ++n, method, params, ...(sessionId ? { sessionId } : {}) }));
    b.onmessage = (ev) => {
      const m = JSON.parse(ev.data);
      if (m.method === "Target.attachedToTarget") {
        const { sessionId, targetInfo, waitingForDebugger } = m.params;
        const isOther = targetInfo.type === "page" && !/taskpane\.html/.test(targetInfo.url);
        if (isOther) {
          bsend("Network.enable", {}, sessionId);
          bsend("Network.setCacheDisabled", { cacheDisabled: true }, sessionId);
          bsend("Fetch.enable", { patterns: [{ urlPattern: prefix + "*" }] }, sessionId);
          console.log(`gắn vào trang mới: ${safeUrl(targetInfo.url) || "(trống)"}`);
        }
        if (waitingForDebugger) bsend("Runtime.runIfWaitingForDebugger", {}, sessionId);
      } else if (m.method === "Fetch.requestPaused" && m.sessionId) {
        const pr = m.params;
        try {
          const u = new URL(pr.request.url);
          const rel = decodeURIComponent((u.origin + u.pathname).slice(prefix.length));
          const file = path.join(base, rel);
          if (rel && file.startsWith(base) && fs.existsSync(file) && fs.statSync(file).isFile()) {
            const ext = path.extname(file).slice(1).toLowerCase();
            bsend(
              "Fetch.fulfillRequest",
              {
                requestId: pr.requestId,
                responseCode: 200,
                responseHeaders: [
                  { name: "Content-Type", value: MIME[ext] || "application/octet-stream" },
                  { name: "Cache-Control", value: "no-store" },
                ],
                body: fs.readFileSync(file).toString("base64"),
              },
              m.sessionId
            );
            served++;
            console.log(`cục bộ  ${rel}  (trang mới)`);
          } else {
            bsend("Fetch.continueRequest", { requestId: pr.requestId }, m.sessionId);
            console.log(`mạng    ${rel}  (trang mới)`);
          }
        } catch (e) {
          console.error("LỖI khi trả tệp cho trang mới:", e.message);
        }
      }
    };
    bsend("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: true, flatten: true });
    c.closed.then(() => b.close());
  } catch (e) {
    console.log(`Không tự gắn được vào trang mở sau (hộp thoại): ${e.message}`);
  }
  console.log(`Đang trả bản cục bộ từ ${base} cho ${prefix}* — giữ tiến trình này tới khi thử xong.`);
  await c.closed;
  console.log(`Khung đã đóng. Đã trả ${served} tệp cục bộ.`);
}

try {
  if (cmd === "make") await make(a1, a2);
  else if (cmd === "list") await listTargets(a1);
  else if (cmd === "nav") await navigate(a1, a2);
  else if (cmd === "reload") await reload(a1);
  else if (cmd === "thu") await testMode(a1, a2);
  else if (cmd === "serve") await serve(a1, a2);
  else if (cmd === "file") await setFile(a1, a2, process.argv.slice(5));
  else if (cmd === "eval") await evalFile(a1, files);
  else if (cmd === "shot") await shot(a1, a2);
  else console.log("Cách dùng: xem chú thích đầu tệp kiem-thu/cdp.mjs");
} catch (e) {
  console.error("LỖI:", e.message);
  process.exitCode = 1;
}
