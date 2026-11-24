# Kiểm trong Office bằng mã

Không cần điều khiển màn hình, dùng được cả khi anh không ngồi trước máy. Giới hạn và bảng thăm dò desktop ẩn ở `docs/bai-hoc.md`.

## Công cụ

- `kiem-thu/cdp.mjs`: điều khiển khung add-in qua cổng gỡ lỗi (cách dùng ở đầu tệp). Bài thử mẫu ở `kiem-thu/cdp-bai/`.
- Nạp lại khung: `node kiem-thu/cdp.mjs reload 9229` (tay: con trỏ trong khung rồi Ctrl+F5). Số phiên bản hiện ở đầu khung.
- **Thử mã mới trước khi phát hành**: `npm run build` rồi `node kiem-thu/cdp.mjs serve 9229 dist` (chạy nền) nạp bản cục bộ vào khung add-in, không đẩy gì lên `main`.
- Mở và đóng Office thử: `kiem-thu/mo-office.ps1` — **mặc định mở trên desktop ẩn** (`kiem-thu/desktop-rieng.ps1`), không hiện lên màn hình của anh; tự dừng nếu anh đang mở Office ở desktop chính; `-HienLen` chỉ dùng khi anh đã đồng ý.
- Lệnh COM tới Office thử phải chạy bên trong desktop ẩn (`desktop-rieng.ps1 -Lenh powershell.exe -ThamSo …`).
- Ba ứng dụng cùng lúc: mở Word bằng `mo-office.ps1`, rồi `Start-Process excel.exe` / `powerpnt.exe` kèm biến môi trường cổng gỡ lỗi; mỗi khung một tiến trình `cdp.mjs serve 9229/<Ứng dụng> dist`.
- Hộp thoại đăng nhập là một trang riêng trên cùng cổng: `node kiem-thu/cdp.mjs list 9229` liệt kê các trang (giá trị tham số nhạy cảm chỉ in độ dài); `9229#openrouter.ai` hoặc `9229#dang-nhap` chọn trang đó cho `eval`, `shot`, `nav`; `serve` tự gắn vào hộp thoại để nó cũng nhận bản cục bộ.
- Xem bản in của văn bản Word: Word COM `ExportAsFixedFormat` ra PDF rồi `pypdfium2` ra PNG (cách làm ở `docs/bai-hoc.md`).
- Đo token mỗi lượt: `kiem-thu/do-token.ts`.

## Tệp thử

- Word: `kiem-thu/docx-thu.mjs` (thăm dò đánh số), `kiem-thu/hop-dong-thu.mjs` (hợp đồng 50 trang có sửa đổi và nhận xét của "Đối tác"), `kiem-thu/soat-thu.mjs` (hợp đồng sạch / cài sẵn 9 lỗi / phiên bản khác lệch 10 chỗ / công văn lề sai, có chú thích, đầu và chân trang, bookmark), `kiem-thu/pdf-quet-thu.mjs` (PDF quét 3 trang).
- PowerPoint: `kiem-thu/ppt-thu.mjs` (10 slide theo mẫu riêng, cài 5 lỗi lệch mẫu và 1 slide chữ tràn).

## Bài thử theo đợt (`kiem-thu/cdp-bai/`)

- Đợt 2 (đăng nhập): `dot-2-*.js`.
- Đợt 3: `dot-3-*.js`, `tham-do-3-*.js`.
- Đợt 4 (Excel): `dot-4-excel-du-lieu.js` dựng ba sheet thử (bảng tiến độ, mô hình cài 4 lỗi, bảng hồ sơ có dòng trùng), `dot-4-excel-bai.js` giao việc rồi đọc lại sổ tính và số lần gọi AI, `dot-4-cong-tac.js` bật / tắt hai công tắc nhật ký, `dot-4-word-gui.js` (Word thu nhỏ, Groq 413), `tham-do-4-excel.js`.
