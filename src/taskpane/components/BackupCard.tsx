/** Thẻ "Sao lưu và khôi phục" (trang Cài đặt). Logic ở src/core/backup.ts (có kiểm thử). */
import * as React from "react";
import { Button, Text, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { getKV } from "../../core/db";
import { BackupFile, backupFileName, createBackup, parseBackup, restoreBackup, summarizeBackup } from "../../core/backup";
import { APP_VERSION_FULL } from "../../core/version";
import { copyText, downloadText } from "../fileio";

const useStyles = makeStyles({
  card: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: "8px",
    padding: "10px",
    marginBottom: "10px",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  row: { display: "flex", gap: "6px", alignItems: "center", flexWrap: "wrap", minWidth: 0 },
  warn: { color: tokens.colorPaletteDarkOrangeForeground1, fontWeight: 600 },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
  box: { border: `1px dashed ${tokens.colorNeutralStroke1}`, borderRadius: "6px", padding: "8px", display: "flex", flexDirection: "column", gap: "6px" },
});

const BackupCard: React.FC<{ host?: string; onRestored: () => void }> = ({ host, onRestored }) => {
  const s = useStyles();
  const [msg, setMsg] = React.useState<{ text: string; err?: boolean }>();
  const [pending, setPending] = React.useState<BackupFile | null>(null);
  const [paste, setPaste] = React.useState<string | null>(null);
  const [confirmReplace, setConfirmReplace] = React.useState(false);
  const fileRef = React.useRef<HTMLInputElement>(null);

  const make = () => createBackup(localStorage, getKV(), { appVersion: APP_VERSION_FULL, host });

  const download = async () => {
    try {
      const b = await make();
      const name = backupFileName(b);
      const ok = downloadText(name, JSON.stringify(b, null, 1));
      const sum = summarizeBackup(b);
      setMsg(
        ok
          ? {
              text: `Đã tạo tệp “${name}” (thư mục Tải xuống): ${sum.keys} key, ${sum.conversations} hội thoại, ${sum.skills} kỹ năng, ${sum.commands} lệnh. Nếu không thấy tệp, dùng “Chép vào bộ nhớ tạm”.`,
            }
          : { text: "Khung add-in không cho tải tệp. Dùng “Chép vào bộ nhớ tạm” rồi dán vào một tệp .json.", err: true }
      );
    } catch (e: any) {
      setMsg({ text: "Không tạo được bản sao lưu: " + (e?.message || e), err: true });
    }
  };

  const copy = async () => {
    try {
      const b = await make();
      const ok = await copyText(JSON.stringify(b));
      setMsg(
        ok
          ? { text: "Đã chép bản sao lưu vào bộ nhớ tạm. Dán vào Notepad, lưu thành tệp .json và cất kỹ (chứa API key)." }
          : { text: "Không chép được vào bộ nhớ tạm trong môi trường này.", err: true }
      );
    } catch (e: any) {
      setMsg({ text: "Không tạo được bản sao lưu: " + (e?.message || e), err: true });
    }
  };

  const load = (text: string) => {
    try {
      setPending(parseBackup(text));
      setPaste(null);
      setConfirmReplace(false);
      setMsg(undefined);
    } catch (e: any) {
      setMsg({ text: e?.message || String(e), err: true });
    }
  };

  const onFile = async (files: FileList | null) => {
    const f = files?.[0];
    if (fileRef.current) fileRef.current.value = "";
    if (!f) return;
    try {
      load(await f.text());
    } catch (e: any) {
      setMsg({ text: "Không đọc được tệp: " + (e?.message || e), err: true });
    }
  };

  const apply = async (mode: "merge" | "replace") => {
    if (!pending) return;
    if (mode === "replace" && !confirmReplace) {
      setConfirmReplace(true);
      return;
    }
    try {
      const r = await restoreBackup(pending, localStorage, getKV(), mode);
      setPending(null);
      setConfirmReplace(false);
      setMsg({ text: `Đã khôi phục (${mode === "merge" ? "gộp" : "thay toàn bộ"}): ${r.local} mục cài đặt, ${r.kv} mục dữ liệu.` });
      onRestored();
    } catch (e: any) {
      setMsg({ text: "Khôi phục lỗi: " + (e?.message || e), err: true });
    }
  };

  const sum = pending ? summarizeBackup(pending) : null;

  return (
    <div className={s.card}>
      <Text weight="semibold">Sao lưu và khôi phục</Text>
      <Text size={200} className={s.muted}>
        Gom toàn bộ dữ liệu add-in (key, cài đặt, chỉ dẫn riêng, model đã nhớ, kỹ năng, lệnh, lịch sử hội thoại) vào một tệp. Office lưu dữ liệu add-in
        theo từng tài khoản Office: đổi tài khoản hoặc sau khi Office cập nhật, dữ liệu có thể như “biến mất” — khi đó khôi phục từ tệp này.
      </Text>
      <Text size={200} className={s.warn}>
        ⚠️ Tệp sao lưu chứa API key: là tệp bí mật. Không gửi cho người khác, không đưa lên mạng.
      </Text>
      <div className={s.row}>
        <Button size="small" appearance="primary" onClick={download}>
          Tải tệp sao lưu
        </Button>
        <Button size="small" onClick={copy}>
          Chép vào bộ nhớ tạm
        </Button>
      </div>
      <div className={s.row}>
        <Button size="small" onClick={() => fileRef.current?.click()}>
          Khôi phục từ tệp…
        </Button>
        <Button size="small" onClick={() => setPaste((v) => (v === null ? "" : null))}>
          {paste === null ? "Dán bản sao lưu" : "Đóng ô dán"}
        </Button>
      </div>
      <input ref={fileRef} type="file" accept=".json,application/json" style={{ display: "none" }} onChange={(e) => onFile(e.target.files)} />
      {paste !== null && (
        <div className={s.box}>
          <Textarea rows={4} value={paste} onChange={(_, d) => setPaste(d.value)} placeholder="Dán nội dung tệp sao lưu (JSON) vào đây" resize="vertical" />
          <Button size="small" onClick={() => load(paste)} disabled={!paste.trim()} style={{ alignSelf: "flex-start" }}>
            Đọc bản sao lưu
          </Button>
        </div>
      )}
      {pending && sum && (
        <div className={s.box}>
          <Text size={200}>
            Tệp tạo lúc {new Date(pending.createdAt).toLocaleString("vi-VN")} (bản {pending.appVersion}): {sum.keys} key, {sum.instructions} chỉ dẫn riêng,{" "}
            {sum.conversations} hội thoại, {sum.skills} kỹ năng, {sum.commands} lệnh.
          </Text>
          <div className={s.row}>
            <Button size="small" appearance="primary" onClick={() => apply("merge")}>
              Gộp vào dữ liệu hiện có
            </Button>
            <Button size="small" onClick={() => apply("replace")}>
              {confirmReplace ? "Bấm lần nữa để thay toàn bộ" : "Thay toàn bộ"}
            </Button>
            <Button size="small" appearance="subtle" onClick={() => setPending(null)}>
              Hủy
            </Button>
          </div>
          <Text size={100} className={s.muted}>
            Gộp: thêm những gì máy này chưa có, giữ nguyên dữ liệu đang có. Thay toàn bộ: xóa dữ liệu add-in hiện có trên máy này rồi ghi lại theo tệp.
          </Text>
        </div>
      )}
      {msg && (
        <Text size={200} className={msg.err ? s.err : s.ok}>
          {msg.text}
        </Text>
      )}
    </div>
  );
};

export default BackupCard;
