import * as React from "react";
import { Button, Input, Link, Text, makeStyles, tokens } from "@fluentui/react-components";
import { PROVIDERS, Provider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { fetchModels } from "../../core/models";
import { networkErrorMessage } from "../../core/errors";

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
  row: { display: "flex", gap: "6px", alignItems: "center" },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
});

interface Props {
  storage: Storage;
  onChanged: () => void;
}

type Status = { kind: "ok" | "err" | "busy"; text: string } | undefined;

const ProviderCard: React.FC<{ p: Provider } & Props> = ({ p, storage, onChanged }) => {
  const s = useStyles();
  const [key, setKey] = React.useState(storage.getKey(p.id));
  const cached = storage.getModelCache()[p.id];
  const [status, setStatus] = React.useState<Status>(
    cached ? { kind: "ok", text: `Đã kết nối · ${cached.models.length} model` } : undefined
  );

  const save = async () => {
    if (!key.trim()) return;
    storage.setKey(p.id, key);
    setStatus({ kind: "busy", text: "Đang kiểm tra key…" });
    try {
      const models = await fetchModels(p, key.trim());
      storage.setModels(p.id, models);
      setStatus({ kind: "ok", text: `Đã kết nối · ${models.length} model` });
    } catch (e) {
      setStatus({ kind: "err", text: networkErrorMessage(e) });
    }
    onChanged();
  };

  const remove = () => {
    storage.removeKey(p.id);
    setKey("");
    setStatus(undefined);
    onChanged();
  };

  return (
    <div className={s.card}>
      <div className={s.row} style={{ justifyContent: "space-between" }}>
        <Text weight="semibold">{p.name}</Text>
        <Link href={p.keyUrl} target="_blank">
          Lấy key
        </Link>
      </div>
      {p.note && <Text size={200} className={s.muted}>{p.note}</Text>}
      <div className={s.row}>
        <Input
          type="password"
          placeholder="Dán API key"
          value={key}
          onChange={(_, d) => setKey(d.value)}
          style={{ flex: 1 }}
        />
        <Button appearance="primary" onClick={save} disabled={!key.trim() || status?.kind === "busy"}>
          Lưu
        </Button>
        {storage.getKey(p.id) && <Button onClick={remove}>Xóa</Button>}
      </div>
      {status && (
        <Text size={200} className={status.kind === "ok" ? s.ok : status.kind === "err" ? s.err : s.muted}>
          {status.text}
        </Text>
      )}
    </div>
  );
};

const Settings: React.FC<Props> = (props) => {
  const s = useStyles();
  return (
    <div>
      <Text size={200} className={s.muted} block style={{ marginBottom: 10 }}>
        Key chỉ lưu trên máy này, trong add-in; chỉ gửi tới đúng nhà cung cấp khi gọi. Bấm “Lưu” để kiểm tra key và tải
        danh sách model.
      </Text>
      {PROVIDERS.map((p) => (
        <ProviderCard key={p.id} p={p} {...props} />
      ))}
    </div>
  );
};

export default Settings;
