/**
 * B17 (rà soát kiến trúc 20/11/2026, làm 27/11): chèn Content-Security-Policy vào mọi trang HTML lúc build.
 * Mã băm của <script> nội tuyến tính trên HTML ĐÃ rút gọn (html-loader đổi chữ của mã nội tuyến), nên không ghi tay.
 * connect-src để mở: nguồn AI tự khai có thể là địa chỉ bất kỳ.
 */
const crypto = require("crypto");
const HtmlWebpackPlugin = require("html-webpack-plugin");

/** Máy chủ của Office.js (office.js tự nạp thêm tệp từ các máy chủ này) */
const OFFICE_SCRIPT_HOSTS = [
  "https://officeapis.public.onecdn.static.microsoft",
  "https://appsforoffice.microsoft.com",
  "https://*.officeapps.live.com",
  "https://*.microsoft.com",
];

function buildCsp(html, { dev = false } = {}) {
  const hashes = [];
  const re = /<script(\s[^>]*)?>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html))) {
    const attrs = m[1] || "";
    if (/\ssrc\s*=/i.test(attrs)) continue;
    if (!m[2]) continue;
    hashes.push(`'sha256-${crypto.createHash("sha256").update(m[2], "utf8").digest("base64")}'`);
  }
  const script = ["'self'", ...OFFICE_SCRIPT_HOSTS, ...hashes];
  return [
    `default-src 'self' https: data: blob:`,
    `script-src ${script.join(" ")}`,
    `worker-src 'self' blob:`,
    `style-src 'self' 'unsafe-inline' https:`,
    `img-src 'self' data: blob: https:`,
    `font-src 'self' data: https:`,
    `connect-src * data: blob:` + (dev ? " ws: wss:" : ""),
    `frame-src https:`,
    `object-src 'none'`,
    `base-uri 'self'`,
  ].join("; ");
}

function injectCsp(html, opts) {
  if (/http-equiv=["']?Content-Security-Policy/i.test(html)) return html;
  const meta = `<meta http-equiv="Content-Security-Policy" content="${buildCsp(html, opts)}">`;
  // Đặt ngay sau <meta charset> để có hiệu lực trước mọi <script>; thiếu thì dừng build (không lặng lẽ bỏ CSP)
  if (!/<meta charset=[^>]*>/i.test(html)) throw new Error("CspPlugin: trang HTML thiếu <meta charset>, không chèn được CSP");
  return html.replace(/(<meta charset=[^>]*>)/i, `$1${meta}`);
}

class CspPlugin {
  constructor(opts = {}) {
    this.opts = opts;
  }
  apply(compiler) {
    compiler.hooks.compilation.tap("CspPlugin", (compilation) => {
      HtmlWebpackPlugin.getHooks(compilation).beforeEmit.tap("CspPlugin", (data) => {
        data.html = injectCsp(data.html, this.opts);
        return data;
      });
    });
  }
}

module.exports = { CspPlugin, buildCsp, injectCsp };
