/** Bước 0 nhóm 2 (26/11/2026): khung đang chuẩn bị gửi (đã bấm Gửi, chưa chuyển "đang chạy") thì không nhận lệnh ghi chéo. */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import { receiveBlocked } from "../../src/core/crossapp";

describe("Nhóm 2 · bước 0 — lệnh ghi chéo không chen vào lúc khung đang chuẩn bị gửi", () => {
  const free = { working: false, hasTools: true, busy: false, sending: false };
  it("rảnh thì nhận; đang chạy, đang chuẩn bị gửi, đang làm lệnh khác, không có công cụ thì không", () => {
    assert.strictEqual(receiveBlocked(free), false);
    assert.strictEqual(receiveBlocked({ ...free, sending: true }), true);
    assert.strictEqual(receiveBlocked({ ...free, busy: true }), true);
    assert.strictEqual(receiveBlocked({ ...free, working: true }), true);
    assert.strictEqual(receiveBlocked({ ...free, hasTools: false }), true);
  });
  it("Chat.tsx dùng receiveBlocked với cờ đang chuẩn bị gửi (sendingRef)", () => {
    const src = fs.readFileSync(path.join(__dirname, "../../src/taskpane/components/Chat.tsx"), "utf8");
    assert.match(src, /receiveBlocked\(\{[^}]*sending:\s*sendingRef\.current[^}]*\}\)/);
    assert.doesNotMatch(src, /if \(working \|\| !hostTools \|\| busyRef\.current\) return;/);
    // Bận ngay khi lấy được lệnh, TRƯỚC lần chờ giành lệnh (không để người dùng gửi chen vào 80–140 ms đó)
    const body = src.slice(src.indexOf("const receive = async"), src.indexOf("const onStorage"));
    assert.ok(body.indexOf("busyRef.current = true") > -1 && body.indexOf("busyRef.current = true") < body.indexOf("await claimCommand"), "busyRef phải bật trước await claimCommand");
  });
});
