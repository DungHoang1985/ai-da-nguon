/** Phiên bản add-in, hiện ở đầu khung để biết đang chạy bản nào.
 *  Nguồn duy nhất: trường "version" trong package.json, webpack gắn vào lúc đóng gói (DefinePlugin). */
declare const __APP_VERSION__: string | undefined;

export const APP_VERSION_FULL: string = typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "0.0.0-dev";

/** "3.6.0" → "3.6"; "3.6.1" → "3.6.1"; giữ nguyên chuỗi lạ. */
export function displayVersion(v: string): string {
  const m = /^(\d+)\.(\d+)\.(\d+)(.*)$/.exec(v.trim());
  if (!m) return v;
  return m[3] === "0" && !m[4] ? `${m[1]}.${m[2]}` : v;
}

export const APP_VERSION = displayVersion(APP_VERSION_FULL);
