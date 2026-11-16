/** Thẻ của nhà cung cấp có đăng nhập (OpenRouter): đăng nhập một lần bấm trong hộp thoại của Office, lối dự phòng dán mã,
 *  hạn mức của key, Đăng xuất. Phần thuần ở src/core/oauth.ts và src/core/credits.ts (có kiểm thử). */
import * as React from "react";
import { Button, Input, Link, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { Provider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { fetchModels } from "../../core/models";
import { cleanKey, describeError, keyProblem } from "../../core/errors";
import { LoginSession, buildAuthUrl, createLoginSession, dialogStartUrl, exchangeCode, extractCode, readDialogMessage } from "../../core/oauth";
import { describeKeyInfo, fetchKeyInfo, hasKeyInfo } from "../../core/credits";
import { DialogHandle, openDialog, openInBrowser } from "../../office/dialog";
import { invalidateKeyInfo, useCredits } from "./useCredits";

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
  row: { display: "flex", gap: "6px", alignItems: "center", minWidth: 0, flexWrap: "wrap" },
  box: { display: "flex", flexDirection: "column", gap: "6px", minWidth: 0 },
  limits: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    padding: "6px 8px",
    borderRadius: "6px",
    background: tokens.colorNeutralBackground2,
  },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  warn: { color: tokens.colorPaletteDarkOrangeForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
});

type Status = { kind: "ok" | "err" | "busy" | "info"; text: string } | undefined;

/** Địa chỉ trang hộp thoại: cùng thư mục với taskpane.html */
function dialogPageUrl(): string {
  return new URL("dang-nhap.html", window.location.href.split(/[?#]/)[0]).toString();
}

/** Luồng đăng nhập của một nguồn: hộp thoại Office, lối dán mã. Mã kiểm bí mật chỉ nằm trong bộ nhớ của hook này. */
export function useSignIn(storage: Storage, p: Provider, onChanged: () => void) {
  const [status, setStatus] = React.useState<Status>();
  const sessionRef = React.useRef<LoginSession | null>(null);
  const [ready, setReady] = React.useState(false);
  const dialogRef = React.useRef<DialogHandle | null>(null);
  const aliveRef = React.useRef(true);

  const renew = React.useCallback(async () => {
    sessionRef.current = null;
    setReady(false);
    try {
      sessionRef.current = await createLoginSession();
      if (aliveRef.current) setReady(true);
    } catch (e: any) {
      if (aliveRef.current) setStatus({ kind: "err", text: String(e?.message || e) });
    }
  }, []);

  React.useEffect(() => {
    aliveRef.current = true;
    renew();
    return () => {
      aliveRef.current = false;
      dialogRef.current?.close();
      dialogRef.current = null;
      sessionRef.current = null;
    };
  }, [renew]);

  /** Lưu key vừa nhận rồi tải danh sách model */
  const adopt = async (rawKey: string, how: string) => {
    const k = cleanKey(rawKey);
    const problem = keyProblem(k);
    if (!k || problem) {
      setStatus({ kind: "err", text: problem || `${p.name} trả về khóa rỗng. Bấm đăng nhập lại.` });
      return false;
    }
    // OpenRouter trả danh sách model cả khi key sai (đo 16/11/2026): hỏi thẳng máy chủ về chính key này trước khi lưu,
    // để một key sai không thay mất key đang dùng được. Lỗi khác (mất mạng…) thì vẫn lưu và báo ở bước sau.
    if (hasKeyInfo(p)) {
      setStatus({ kind: "busy", text: "Đang kiểm tra khóa…" });
      try {
        await fetchKeyInfo(p, k);
      } catch (e) {
        const info = describeError(e, { provider: p });
        if (info.reason === "key_invalid" || info.reason === "key_expired") {
          setStatus({ kind: "err", text: `${info.text} Chưa lưu khóa này.` });
          return false;
        }
      }
    }
    storage.setKey(p.id, k);
    invalidateKeyInfo(p.id);
    setStatus({ kind: "busy", text: "Đã nhận khóa, đang tải danh sách model…" });
    try {
      const models = await fetchModels(p, k);
      storage.setModels(p.id, models);
      setStatus({ kind: "ok", text: `${how} · ${models.length} model` });
    } catch (e) {
      setStatus({ kind: "err", text: `Đã lưu khóa nhưng chưa tải được danh sách model: ${describeError(e, { provider: p }).text}` });
    }
    onChanged();
    return true;
  };

  /** Đổi mã lấy khóa. Dù được hay không, cặp mã kiểm / thách thức của lần này bỏ đi, tạo cặp mới. */
  const finish = async (code: string) => {
    const session = sessionRef.current;
    if (!session) {
      setStatus({ kind: "err", text: "Lần đăng nhập này đã kết thúc. Bấm đăng nhập lại." });
      return;
    }
    setStatus({ kind: "busy", text: `Đang đổi mã lấy khóa ${p.name}…` });
    try {
      const key = await exchangeCode(p, code, session.verifier);
      await adopt(key, "Đã đăng nhập");
    } catch (e: any) {
      setStatus({ kind: "err", text: String(e?.message || "Không đổi được mã đăng nhập.") });
    } finally {
      await renew();
    }
  };

  /** Lối chính: hộp thoại của Office */
  const signInWithDialog = async () => {
    const session = sessionRef.current;
    if (!session || !p.oauth) return;
    dialogRef.current?.close();
    dialogRef.current = null;
    setStatus({ kind: "info", text: `Đang mở cửa sổ đăng nhập ${p.name}… Đăng nhập (nếu được hỏi) rồi bấm Authorize trong cửa sổ đó.` });
    try {
      const handle = await openDialog(dialogStartUrl(dialogPageUrl(), p.id, session), {
        onMessage: (message) => {
          const r = readDialogMessage(message, sessionRef.current);
          dialogRef.current?.close();
          dialogRef.current = null;
          if (r.ok === true) finish(r.code);
          else setStatus({ kind: "err", text: r.reason });
        },
        onClosed: (_why, text) => {
          dialogRef.current = null;
          setStatus((cur) => (cur?.kind === "info" ? { kind: "err", text } : cur));
        },
      });
      dialogRef.current = handle;
    } catch (e: any) {
      setStatus({ kind: "err", text: String(e?.message || "Không mở được cửa sổ đăng nhập. Dùng lối dán mã bên dưới.") });
    }
  };

  /** Lối dự phòng, bước 1: mở trang của nhà cung cấp trong trình duyệt ngoài; trang đó hiện mã để chép */
  const openCodePage = () => {
    const session = sessionRef.current;
    if (!session || !p.oauth) return;
    const ok = openInBrowser(buildAuthUrl(p.oauth, { challenge: session.challenge }));
    setStatus(
      ok
        ? { kind: "info", text: `Đã mở trang ${p.name} trong trình duyệt. Bấm Authorize, chép mã trên trang rồi dán vào ô bên dưới (mã dùng trong 10 phút).` }
        : { kind: "err", text: "Không mở được trình duyệt. Dùng nút “Lấy key” rồi dán key theo cách cũ." }
    );
  };

  /** Lối dự phòng, bước 2: người dùng dán mã */
  const submitCode = async (pasted: string): Promise<boolean> => {
    const r = extractCode(pasted, p);
    if (r.kind === "invalid") {
      setStatus({ kind: "err", text: r.reason });
      return false;
    }
    if (r.kind === "key") {
      // Dán nhầm API key vào ô mã: dùng luôn như key dán tay
      return adopt(r.key, "Đã kết nối bằng key vừa dán");
    }
    await finish(r.code);
    return true;
  };

  /** Lưu key dán tay (cách cũ) */
  const saveKey = async (raw: string): Promise<boolean> => {
    const k = cleanKey(raw);
    if (!k) return false;
    const problem = keyProblem(k);
    if (problem) {
      setStatus({ kind: "err", text: problem });
      return false;
    }
    return adopt(k, "Đã kết nối");
  };

  /** Đăng xuất: xóa khóa và danh mục model của nguồn này khỏi kho của add-in */
  const signOut = () => {
    dialogRef.current?.close();
    dialogRef.current = null;
    storage.removeKey(p.id);
    invalidateKeyInfo(p.id);
    setStatus({
      kind: "ok",
      text: `Đã xóa khóa ${p.name} khỏi máy này. Khóa vẫn còn hiệu lực ở phía ${p.name} cho tới khi anh xóa nó ở trang quản lý key.`,
    });
    onChanged();
  };

  return { status, setStatus, ready, signInWithDialog, openCodePage, submitCode, saveKey, signOut };
}

type SignIn = ReturnType<typeof useSignIn>;

/** Nút đăng nhập + lối dán mã (dùng ở thẻ Cài đặt và ở màn chào) */
export const LoginControls: React.FC<{ p: Provider; signIn: SignIn; primaryLabel?: string }> = ({ p, signIn, primaryLabel }) => {
  const s = useStyles();
  const [code, setCode] = React.useState("");
  const busy = signIn.status?.kind === "busy";
  const submit = async () => {
    if (await signIn.submitCode(code)) setCode("");
  };
  return (
    <div className={s.box}>
      <div className={s.row}>
        <Button appearance="primary" onClick={signIn.signInWithDialog} disabled={!signIn.ready}>
          {primaryLabel || `Đăng nhập ${p.name}`}
        </Button>
        {!signIn.ready && <Spinner size="tiny" />}
      </div>
      <Text size={200} className={s.muted}>
        Mở cửa sổ của {p.name}: đăng nhập (nếu được hỏi), bấm Authorize là xong — không phải chép key. Add-in không thấy mật khẩu của anh.
      </Text>
      <Text size={200} weight="semibold">
        Cửa sổ không mở hoặc không đăng nhập được? Dùng lối dán mã:
      </Text>
      <div className={s.row}>
        <Button size="small" onClick={signIn.openCodePage} disabled={!signIn.ready}>
          Mở trang lấy mã trong trình duyệt
        </Button>
      </div>
      <div className={s.row}>
        <Input
          size="small"
          autoComplete="off"
          name="aidn-oauth-code"
          placeholder="Dán mã vào đây"
          aria-label="Dán mã vào đây"
          value={code}
          onChange={(_, d) => setCode(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && code.trim() && !busy) submit();
          }}
          style={{ flex: 1, minWidth: 0 }}
        />
        <Button size="small" onClick={submit} disabled={!code.trim() || busy}>
          Xác nhận mã
        </Button>
      </div>
    </div>
  );
};

export const StatusLine: React.FC<{ status: Status }> = ({ status }) => {
  const s = useStyles();
  if (!status) return null;
  return (
    <Text size={200} className={status.kind === "ok" ? s.ok : status.kind === "err" ? s.err : s.muted} role="status">
      {status.text}
    </Text>
  );
};

interface Props {
  p: Provider;
  storage: Storage;
  onChanged: () => void;
}

const SignInCard: React.FC<Props> = ({ p, storage, onChanged }) => {
  const s = useStyles();
  const signIn = useSignIn(storage, p, onChanged);
  const hasKey = !!storage.getKey(p.id);
  const cached = storage.getModelCache()[p.id];
  const credits = useCredits(storage, p);
  const [relogin, setRelogin] = React.useState(false);
  const [confirmOut, setConfirmOut] = React.useState(false);
  const [key, setKey] = React.useState("");
  const view = credits.info ? describeKeyInfo(credits.info, { balanceAllowed: credits.balanceAllowed }) : undefined;

  // Đăng nhập xong thì thu gọn phần đăng nhập lại
  React.useEffect(() => {
    if (signIn.status?.kind === "ok") setRelogin(false);
  }, [signIn.status]);

  const saveKey = async () => {
    if (await signIn.saveKey(key)) setKey("");
  };

  return (
    <div className={s.card} data-aidn-card={`signin-${p.id}`}>
      <div className={s.row} style={{ justifyContent: "space-between" }}>
        <Text weight="semibold">{p.name}</Text>
        <Link href={p.keyUrl} target="_blank">
          Lấy key
        </Link>
      </div>
      {p.note && (
        <Text size={200} className={s.muted}>
          {p.note}
        </Text>
      )}
      {hasKey ? (
        <>
          {!signIn.status && (
            <Text size={200} className={s.ok}>
              {cached ? `Đã đăng nhập · ${cached.models.length} model` : "Đã có khóa trên máy này"}
            </Text>
          )}
          <StatusLine status={signIn.status} />
          {credits.enabled && (
            <div className={s.limits} data-aidn="han-muc">
              <div className={s.row} style={{ justifyContent: "space-between" }}>
                <Text size={200} weight="semibold">
                  Hạn mức
                </Text>
                <div className={s.row}>
                  {credits.loading && <Spinner size="extra-tiny" />}
                  <Button size="small" appearance="subtle" onClick={() => credits.reload(true)} disabled={credits.loading}>
                    Làm mới
                  </Button>
                  {p.billingUrl && (
                    <Link href={p.billingUrl} target="_blank">
                      Xem số dư
                    </Link>
                  )}
                </div>
              </div>
              {view?.lines.map((l, i) => (
                <Text key={i} size={200}>
                  {l}
                </Text>
              ))}
              {view?.warnings.map((w, i) => (
                <Text key={`w${i}`} size={200} className={view.level === "out" ? s.err : s.warn}>
                  ⚠️ {w}
                </Text>
              ))}
              {!view && credits.error && (
                <Text size={200} className={s.err}>
                  Không xem được hạn mức: {credits.error}
                </Text>
              )}
              {!view && !credits.error && !credits.loading && (
                <Text size={200} className={s.muted}>
                  Bấm “Làm mới” để xem hạn mức.
                </Text>
              )}
            </div>
          )}
          {confirmOut ? (
            <div className={s.box}>
              <Text size={200}>
                Xóa khóa {p.name} khỏi máy này? Muốn dùng lại thì đăng nhập lại (hoặc dán key). Hội thoại đã lưu không bị xóa.
              </Text>
              <div className={s.row}>
                <Button
                  size="small"
                  appearance="primary"
                  onClick={() => {
                    setConfirmOut(false);
                    setRelogin(false);
                    signIn.signOut();
                  }}
                >
                  Xóa khóa trên máy
                </Button>
                <Button size="small" onClick={() => setConfirmOut(false)}>
                  Thôi
                </Button>
              </div>
            </div>
          ) : (
            <div className={s.row}>
              <Button size="small" onClick={() => setConfirmOut(true)}>
                Đăng xuất
              </Button>
              <Button size="small" appearance="subtle" onClick={() => setRelogin((v) => !v)}>
                {relogin ? "Ẩn đăng nhập lại" : "Đăng nhập lại (đổi khóa)"}
              </Button>
            </div>
          )}
          {relogin && <LoginControls p={p} signIn={signIn} primaryLabel="Đăng nhập lại" />}
        </>
      ) : (
        <>
          {signIn.status ? (
            <StatusLine status={signIn.status} />
          ) : (
            <Text size={200} className={s.muted}>
              Chưa đăng nhập
            </Text>
          )}
          {signIn.status?.kind === "ok" && p.manageKeysUrl && (
            <Link href={p.manageKeysUrl} target="_blank">
              Mở trang quản lý key của {p.name}
            </Link>
          )}
          <LoginControls p={p} signIn={signIn} />
          <Text size={200} weight="semibold">
            Hoặc dán API key như trước:
          </Text>
          <div className={s.row}>
            <Input
              type="password"
              size="small"
              placeholder="Dán API key"
              value={key}
              onChange={(_, d) => setKey(d.value)}
              style={{ flex: 1, minWidth: 0 }}
            />
            <Button size="small" onClick={saveKey} disabled={!key.trim() || signIn.status?.kind === "busy"}>
              Lưu
            </Button>
          </div>
        </>
      )}
    </div>
  );
};

/** Nút đăng nhập ở màn chào (chưa có key nào) */
export const WelcomeSignIn: React.FC<{ p: Provider; storage: Storage; onChanged: () => void }> = ({ p, storage, onChanged }) => {
  const signIn = useSignIn(storage, p, onChanged);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6, minWidth: 0 }} data-aidn-card={`welcome-${p.id}`}>
      <LoginControls p={p} signIn={signIn} />
      <StatusLine status={signIn.status} />
    </div>
  );
};

export default SignInCard;
