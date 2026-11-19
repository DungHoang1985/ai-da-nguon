// Mở bảng chọn model (để chụp); ESC = đóng
await toChat();
if (typeof DONG !== "undefined") { if (document.querySelector('select[aria-label="Nhà cung cấp"]')) btn(/^Model đang dùng/).click(); await sleep(300); return { dong: true }; }
btn(/^Model đang dùng/).click();
await sleep(1500);
return { mo: true };
