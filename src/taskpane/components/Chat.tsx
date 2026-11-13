import * as React from "react";
import {
  Button,
  Checkbox,
  Input,
  Select,
  Spinner,
  Switch,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { PROVIDERS, getProvider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { ModelInfo, fetchModels, isStale } from "../../core/models";
import { ChatMessage, streamChat } from "../../core/chat";
import { networkErrorMessage } from "../../core/errors";
import { HostName, INSERT_LABEL, insertAnswer, readSelection } from "../../office/host";

const useStyles = makeStyles({
  wrap: { display: "flex", flexDirection: "column", gap: "8px", height: "100%" },
  row: { display: "flex", gap: "6px", alignItems: "center", flexWrap: "wrap" },
  list: { flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: "8px", minHeight: "120px" },
  user: {
    alignSelf: "flex-end",
    background: tokens.colorBrandBackground2,
    borderRadius: "8px",
    padding: "8px",
    maxWidth: "90%",
    whiteSpace: "pre-wrap",
  },
  bot: {
    background: tokens.colorNeutralBackground3,
    borderRadius: "8px",
    padding: "8px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  muted: { color: tokens.colorNeutralForeground3 },
  err: { color: tokens.colorPaletteRedForeground1 },
});

interface Props {
  storage: Storage;
  host: HostName;
  version: number;
  goSettings: () => void;
}

const SYSTEM = "Bạn là trợ lý trong Microsoft Office. Trả lời bằng tiếng Việt, ngắn gọn, đúng trọng tâm. Khi trình bày dữ liệu dạng bảng, dùng bảng Markdown.";

const Chat: React.FC<Props> = ({ storage, host, version, goSettings }) => {
  const s = useStyles();
  const keys = storage.getKeys();
  const available = PROVIDERS.filter((p) => keys[p.id]);
  const sel = storage.getSelection();

  const [providerId, setProviderId] = React.useState<string>(
    sel.providerId && keys[sel.providerId] ? sel.providerId : available[0]?.id || ""
  );
  const [modelId, setModelId] = React.useState<string>(sel.modelId || "");
  const [freeOnly, setFreeOnly] = React.useState<boolean>(!!sel.freeOnly);
  const [search, setSearch] = React.useState("");
  const [models, setModels] = React.useState<ModelInfo[]>([]);
  const [catalogNote, setCatalogNote] = React.useState("");
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [input, setInput] = React.useState("");
  const [withSelection, setWithSelection] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [notice, setNotice] = React.useState<{ text: string; err?: boolean }>();
  const abortRef = React.useRef<AbortController | null>(null);
  const listRef = React.useRef<HTMLDivElement>(null);

  const loadModels = React.useCallback(
    async (force = false) => {
      const p = getProvider(providerId);
      if (!p) return;
      const cached = storage.getModelCache()[p.id];
      if (cached) {
        setModels(cached.models);
        setCatalogNote(`Cập nhật lúc ${new Date(cached.at).toLocaleString("vi-VN")}`);
      }
      if (force || isStale(cached?.at)) {
        setCatalogNote("Đang tải danh sách model…");
        try {
          const list = await fetchModels(p, storage.getKey(p.id));
          storage.setModels(p.id, list);
          setModels(list);
          setCatalogNote(`Cập nhật lúc ${new Date().toLocaleString("vi-VN")} · ${list.length} model`);
        } catch (e) {
          setCatalogNote(
            (cached ? `Dùng danh sách cũ (${new Date(cached.at).toLocaleString("vi-VN")}). ` : "") +
              "Không tải được: " +
              networkErrorMessage(e)
          );
        }
      }
    },
    [providerId, storage]
  );

  React.useEffect(() => {
    loadModels();
  }, [loadModels, version]);

  React.useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight });
  }, [messages]);

  const shown = models
    .filter((m) => !freeOnly || m.free)
    .filter((m) => !search || m.name.toLowerCase().includes(search.toLowerCase()));

  React.useEffect(() => {
    if (shown.length && !shown.some((m) => m.id === modelId)) {
      setModelId(shown[0].id);
    }
  }, [shown.map((m) => m.id).join("|")]); // eslint-disable-line react-hooks/exhaustive-deps

  React.useEffect(() => {
    storage.setSelection({ providerId, modelId, freeOnly });
  }, [providerId, modelId, freeOnly, storage]);

  if (!available.length) {
    return (
      <div className={s.wrap}>
        <Text>Chưa có API key nào.</Text>
        <Button appearance="primary" onClick={goSettings}>
          Mở Cài đặt để nhập key
        </Button>
      </div>
    );
  }

  const send = async () => {
    const p = getProvider(providerId);
    if (!p || !modelId || !input.trim()) return;
    setNotice(undefined);
    let content = input.trim();
    if (withSelection) {
      try {
        const sel = await readSelection();
        if (sel.text.trim()) content = `${content}\n\n--- Nội dung đang chọn ---\n${sel.text}`;
        setNotice({ text: sel.note });
      } catch (e) {
        setNotice({ text: "Không đọc được nội dung đang chọn: " + networkErrorMessage(e), err: true });
      }
    }
    const history: ChatMessage[] = [...messages, { role: "user", content }];
    setMessages([...history, { role: "assistant", content: "" }]);
    setInput("");
    setBusy(true);
    const ac = new AbortController();
    abortRef.current = ac;
    try {
      await streamChat({
        provider: p,
        key: storage.getKey(p.id),
        model: modelId,
        messages: [{ role: "system", content: SYSTEM }, ...history],
        signal: ac.signal,
        onDelta: (d) =>
          setMessages((prev) => {
            const next = prev.slice();
            const last = next[next.length - 1];
            next[next.length - 1] = { ...last, content: last.content + d };
            return next;
          }),
      });
    } catch (e) {
      setNotice({ text: networkErrorMessage(e), err: !(e instanceof DOMException && e.name === "AbortError") });
    } finally {
      setBusy(false);
      abortRef.current = null;
    }
  };

  const stop = () => abortRef.current?.abort();

  const insert = async (text: string) => {
    try {
      setNotice({ text: await insertAnswer(text) });
    } catch (e) {
      setNotice({ text: "Không chèn được: " + networkErrorMessage(e), err: true });
    }
  };

  const copy = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setNotice({ text: "Đã sao chép." });
    } catch {
      setNotice({ text: "Không sao chép được trong môi trường này.", err: true });
    }
  };

  return (
    <div className={s.wrap}>
      <div className={s.row}>
        <Select value={providerId} onChange={(_, d) => setProviderId(d.value)} style={{ flex: 1, minWidth: 0 }}>
          {available.map((p) => (
            <option key={p.id} value={p.id}>
              {p.name}
            </option>
          ))}
        </Select>
        <Button size="small" onClick={() => loadModels(true)} title="Tải lại danh sách model">
          Làm mới
        </Button>
      </div>
      <div className={s.row}>
        <Input
          placeholder="Tìm model…"
          value={search}
          onChange={(_, d) => setSearch(d.value)}
          size="small"
          style={{ flex: 1, minWidth: 0 }}
        />
        {providerId === "openrouter" && (
          <Switch label="Chỉ miễn phí" checked={freeOnly} onChange={(_, d) => setFreeOnly(d.checked)} />
        )}
      </div>
      <Select value={modelId} onChange={(_, d) => setModelId(d.value)}>
        {shown.map((m) => (
          <option key={m.id} value={m.id}>
            {m.free ? "🆓 " : ""}
            {m.name}
          </option>
        ))}
      </Select>
      <Text size={200} className={s.muted}>
        {catalogNote} {shown.length !== models.length ? `· đang hiện ${shown.length}/${models.length}` : ""}
      </Text>

      <div className={s.list} ref={listRef}>
        {messages.map((m, i) =>
          m.role === "user" ? (
            <div key={i} className={s.user}>
              {m.content.split("\n--- Nội dung đang chọn ---")[0]}
            </div>
          ) : (
            <div key={i} className={s.bot}>
              {m.content || (busy && i === messages.length - 1 ? <Spinner size="tiny" /> : "")}
              {m.content && !(busy && i === messages.length - 1) && (
                <div className={s.row} style={{ marginTop: 6 }}>
                  <Button size="small" appearance="primary" onClick={() => insert(m.content)}>
                    {INSERT_LABEL[host]}
                  </Button>
                  <Button size="small" onClick={() => copy(m.content)}>
                    Sao chép
                  </Button>
                </div>
              )}
            </div>
          )
        )}
      </div>

      {notice && (
        <Text size={200} className={notice.err ? s.err : s.muted}>
          {notice.text}
        </Text>
      )}
      <Checkbox
        label={host === "Excel" ? "Gửi kèm vùng ô đang chọn" : "Gửi kèm nội dung đang chọn"}
        checked={withSelection}
        onChange={(_, d) => setWithSelection(!!d.checked)}
      />
      <Textarea
        placeholder="Nhập câu hỏi… (Ctrl+Enter để gửi)"
        value={input}
        onChange={(_, d) => setInput(d.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) send();
        }}
        resize="vertical"
        rows={3}
      />
      <div className={s.row}>
        {busy ? (
          <Button onClick={stop}>Dừng</Button>
        ) : (
          <Button appearance="primary" onClick={send} disabled={!input.trim() || !modelId}>
            Gửi
          </Button>
        )}
        <Button onClick={() => setMessages([])} disabled={busy || !messages.length}>
          Cuộc trò chuyện mới
        </Button>
      </div>
    </div>
  );
};

export default Chat;
