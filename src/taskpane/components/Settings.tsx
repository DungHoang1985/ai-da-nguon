import * as React from "react";
import { Button, Input, Link, Spinner, Switch, Text, makeStyles, tokens } from "@fluentui/react-components";
import { HostName } from "../../office/host";
import { getHostTools } from "../../office/tools";
import { TestStep, runSelfTest } from "../../office/selftest";
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
  row: { display: "flex", gap: "6px", alignItems: "center", minWidth: 0 },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
});

interface Props {
  storage: Storage;
  onChanged: () => void;
  host?: HostName;
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
          style={{ flex: 1, minWidth: 0 }}
        />
        <Button appearance="primary" style={{ minWidth: 0 }} onClick={save} disabled={!key.trim() || status?.kind === "busy"}>
          Lưu
        </Button>
        {storage.getKey(p.id) && <Button style={{ minWidth: 0 }} onClick={remove}>Xóa</Button>}
      </div>
      {status && (
        <Text size={200} className={status.kind === "ok" ? s.ok : status.kind === "err" ? s.err : s.muted}>
          {status.text}
        </Text>
      )}
    </div>
  );
};

const AgentSettings: React.FC<Props> = ({ storage, onChanged, host }) => {
  const s = useStyles();
  const [st, setSt] = React.useState(storage.getSettings());
  const [steps, setSteps] = React.useState<TestStep[] | null>(null);
  const [running, setRunning] = React.useState(false);
  const set = (patch: any) => {
    storage.setSettings(patch);
    setSt(storage.getSettings());
    onChanged();
  };
  const tools = host ? getHostTools(host, () => storage.getSettings()) : null;
  const test = async () => {
    if (!tools) return;
    setRunning(true);
    setSteps([]);
    try {
      await runSelfTest(tools, setSteps);
    } finally {
      setRunning(false);
    }
  };
  const passed = steps?.filter((x) => x.ok === true).length ?? 0;
  const failed = steps?.filter((x) => x.ok === false).length ?? 0;
  return (
    <div className={s.card}>
      <Text weight="semibold">AI tự thao tác trên tài liệu</Text>
      <Switch label="Cho AI tự đọc, sửa, tạo nội dung trong Office" checked={st.agentMode} onChange={(_, d) => set({ agentMode: d.checked })} />
      <Switch label="Hỏi tôi trước mỗi thao tác sửa tài liệu" checked={st.askBeforeWrite} onChange={(_, d) => set({ askBeforeWrite: d.checked })} />
      <Switch label="Word: ghi sửa đổi bằng Theo dõi thay đổi (Track Changes)" checked={st.trackChanges} onChange={(_, d) => set({ trackChanges: d.checked })} />
      <Text weight="semibold" style={{ marginTop: 6 }}>
        Tự kiểm tra chức năng {host && host !== "Khác" ? `(${host})` : ""}
      </Text>
      <Text size={200} className={s.muted}>
        Chạy thật từng thao tác (đọc, ghi, định dạng, tạo bảng/biểu đồ/slide, nhận xét…) trên tài liệu đang mở rồi tự dọn sạch. Không cần API key. Nên chạy trên tài liệu trống.
      </Text>
      <div className={s.row}>
        <Button onClick={test} disabled={!tools || running}>
          Chạy tự kiểm tra
        </Button>
        {running && <Spinner size="tiny" />}
        {steps && !running && (
          <Text weight="semibold" className={failed ? s.err : s.ok}>
            {failed ? `${failed} lỗi` : "Tất cả đạt"} · {passed}/{steps.length} bước đạt
          </Text>
        )}
      </div>
      {steps?.map((x, i) => (
        <Text key={i} size={200} className={x.ok === false ? s.err : x.ok ? s.ok : s.muted}>
          {x.ok === true ? "✅" : x.ok === false ? "❌" : "➖"} {x.name}
          {x.detail ? ` — ${x.detail}` : ""}
        </Text>
      ))}
    </div>
  );
};

const Settings: React.FC<Props> = (props) => {
  const s = useStyles();
  return (
    <div>
      <AgentSettings {...props} />
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
