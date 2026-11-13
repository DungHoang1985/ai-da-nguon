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
import { streamChat } from "../../core/chat";
import { networkErrorMessage } from "../../core/errors";
import { AgentMessage, ToolCall, UndoEntry } from "../../core/agent/types";
import { compactHistory, runAgent } from "../../core/agent/loop";
import { isToolUnsupportedError } from "../../core/agent/llm";
import { HostName, INSERT_LABEL, insertAnswer, readSelection } from "../../office/host";
import { QUICK_PROMPTS, describeCall, getHostTools } from "../../office/tools";

const useStyles = makeStyles({
  wrap: { display: "flex", flexDirection: "column", gap: "8px", height: "100%", minWidth: 0 },
  row: { display: "flex", gap: "6px", alignItems: "center", flexWrap: "wrap", minWidth: 0 },
  list: {
    flex: 1,
    overflowY: "auto",
    overflowX: "hidden",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    minHeight: "140px",
  },
  user: {
    alignSelf: "flex-end",
    background: tokens.colorBrandBackground2,
    borderRadius: "8px",
    padding: "8px",
    maxWidth: "90%",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  bot: {
    background: tokens.colorNeutralBackground3,
    borderRadius: "8px",
    padding: "8px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  step: { fontSize: "12px", color: tokens.colorNeutralForeground2, whiteSpace: "normal" },
  stepErr: { fontSize: "12px", color: tokens.colorPaletteRedForeground1, whiteSpace: "normal" },
  approve: {
    border: `1px solid ${tokens.colorPaletteYellowBorder2}`,
    background: tokens.colorPaletteYellowBackground1,
    borderRadius: "6px",
    padding: "6px",
    whiteSpace: "normal",
  },
  chip: { whiteSpace: "normal", textAlign: "left", height: "auto", minHeight: "28px" },
  muted: { color: tokens.colorNeutralForeground3 },
  err: { color: tokens.colorPaletteRedForeground1 },
});

type Segment =
  | { type: "text"; text: string }
  | { type: "step"; id: string; label: string; status: "running" | "ok" | "error"; detail?: string };

interface Item {
  role: "user" | "assistant";
  text: string;
  segments?: Segment[];
  undo?: UndoEntry[];
  undone?: boolean;
  wrote?: boolean;
  error?: string;
}

interface Props {
  storage: Storage;
  host: HostName;
  version: number;
  goSettings: () => void;
}

const PLAIN_SYSTEM =
  "Bạn là trợ lý trong Microsoft Office. Trả lời bằng tiếng Việt, ngắn gọn, đúng trọng tâm. Khi trình bày dữ liệu dạng bảng, dùng bảng Markdown.";

/** Hiển thị Markdown đơn giản: **đậm**, tiêu đề #, gạch đầu dòng */
const inline = (s: string) =>
  s.split(/(\*\*[^*]+\*\*)/g).map((part, i) =>
    part.startsWith("**") && part.endsWith("**") && part.length > 4 ? <strong key={i}>{part.slice(2, -2)}</strong> : part
  );

const Rich: React.FC<{ text: string }> = ({ text }) => (
  <span style={{ whiteSpace: "normal" }}>
    {text.split("\n").map((line, i) => {
      const h = line.match(/^\s*#{1,4}\s+(.*)$/);
      if (h) return <div key={i} style={{ fontWeight: 600, marginTop: 4 }}>{inline(h[1])}</div>;
      const b = line.match(/^\s*[-*•]\s+(.*)$/);
      if (b)
        return (
          <div key={i} style={{ paddingLeft: 14, textIndent: -10 }}>
            • {inline(b[1])}
          </div>
        );
      return <div key={i}>{line.trim() ? inline(line) : "\u00a0"}</div>;
    })}
  </span>
);

const finalText = (it: Item) =>
  it.segments
    ? it.segments
        .filter((sg) => sg.type === "text")
        .map((sg: any) => sg.text)
        .join("\n")
        .trim()
    : it.text;

const Chat: React.FC<Props> = ({ storage, host, version, goSettings }) => {
  const s = useStyles();
  const keys = storage.getKeys();
  const available = PROVIDERS.filter((p) => keys[p.id]);
  const sel = storage.getSelection();
  const [settings, setSettingsState] = React.useState(storage.getSettings());
  const updateSettings = (patch: any) => {
    storage.setSettings(patch);
    setSettingsState(storage.getSettings());
  };
  React.useEffect(() => setSettingsState(storage.getSettings()), [version, storage]);

  const [providerId, setProviderId] = React.useState<string>(
    sel.providerId && keys[sel.providerId] ? sel.providerId : available[0]?.id || ""
  );
  const [modelId, setModelId] = React.useState<string>(sel.modelId || "");
  const [freeOnly, setFreeOnly] = React.useState<boolean>(!!sel.freeOnly);
  const [search, setSearch] = React.useState("");
  const [models, setModels] = React.useState<ModelInfo[]>([]);
  const [catalogNote, setCatalogNote] = React.useState("");
  const [items, setItems] = React.useState<Item[]>([]);
  const [input, setInput] = React.useState("");
  const [withSelection, setWithSelection] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [notice, setNotice] = React.useState<{ text: string; err?: boolean }>();
  const [pending, setPending] = React.useState<{ call: ToolCall; resolve: (ok: boolean) => void } | null>(null);
  const abortRef = React.useRef<AbortController | null>(null);
  const historyRef = React.useRef<AgentMessage[]>([]);
  const allowAllRef = React.useRef(false);
  const listRef = React.useRef<HTMLDivElement>(null);

  const hostTools = React.useMemo(() => getHostTools(host, () => storage.getSettings()), [host, storage]);
  const agentOn = settings.agentMode && !!hostTools;

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
  }, [items, pending]);

  const toolsFilter = agentOn && providerId === "openrouter" && settings.toolsOnly;
  const shown = models
    .filter((m) => !freeOnly || m.free)
    .filter((m) => !toolsFilter || m.tools !== false)
    .filter((m) => !search || m.name.toLowerCase().includes(search.toLowerCase()));

  React.useEffect(() => {
    if (shown.length && !shown.some((m) => m.id === modelId)) setModelId(shown[0].id);
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

  /** Cập nhật mục trả lời đang chạy (mục cuối) */
  const patchLast = (fn: (it: Item) => Item) =>
    setItems((prev) => {
      const next = prev.slice();
      next[next.length - 1] = fn(next[next.length - 1]);
      return next;
    });

  const appendText = (d: string) =>
    patchLast((it) => {
      const segs = (it.segments || []).slice();
      const last = segs[segs.length - 1];
      if (last && last.type === "text") segs[segs.length - 1] = { ...last, text: last.text + d };
      else segs.push({ type: "text", text: d.replace(/^\n+/, "") });
      return { ...it, segments: segs };
    });

  const askApproval = (call: ToolCall) =>
    allowAllRef.current ? Promise.resolve(true) : new Promise<boolean>((resolve) => setPending({ call, resolve }));

  const answerApproval = (ok: boolean, all = false) => {
    if (all) allowAllRef.current = true;
    pending?.resolve(ok);
    setPending(null);
  };

  const send = async (text?: string) => {
    const p = getProvider(providerId);
    const question = (text ?? input).trim();
    if (!p || !modelId || !question) return;
    setNotice(undefined);
    let content = question;
    if (withSelection) {
      try {
        const selection = await readSelection();
        if (selection.text.trim()) content = `${content}\n\n--- Nội dung đang chọn ---\n${selection.text}`;
        setNotice({ text: selection.note });
      } catch (e) {
        setNotice({ text: "Không đọc được nội dung đang chọn: " + networkErrorMessage(e), err: true });
      }
    }
    setItems((prev) => [...prev, { role: "user", text: question }, { role: "assistant", text: "", segments: [] }]);
    setInput("");
    setBusy(true);
    allowAllRef.current = false;
    const ac = new AbortController();
    abortRef.current = ac;
    const model = models.find((m) => m.id === modelId);
    const useAgent = agentOn && !!hostTools && model?.tools !== false;

    try {
      if (useAgent && hostTools) {
        let ov: any;
        try {
          ov = await hostTools.overview();
        } catch (e) {
          ov = { error: "Không đọc được trạng thái tài liệu: " + networkErrorMessage(e) };
        }
        const log: UndoEntry[] = [];
        let needSep = false;
        const msgs = await runAgent({
          provider: p,
          key: storage.getKey(p.id),
          model: modelId,
          tools: hostTools,
          system: hostTools.systemPrompt(ov),
          messages: [...compactHistory(historyRef.current), { role: "user", content }],
          signal: ac.signal,
          log,
          approve: settings.askBeforeWrite ? askApproval : undefined,
          onEvent: (ev) => {
            if (ev.type === "step") {
              if (ev.step > 1) needSep = true;
            } else if (ev.type === "text") {
              appendText(needSep ? "\n" + ev.delta : ev.delta);
              needSep = false;
            } else if (ev.type === "tool_start") {
              patchLast((it) => ({
                ...it,
                segments: [
                  ...(it.segments || []),
                  { type: "step", id: ev.call.id, label: describeCall(ev.call), status: "running" },
                ],
              }));
            } else if (ev.type === "tool_end") {
              const isWrite = hostTools.specs.some((t) => t.name === ev.call.name && t.write);
              patchLast((it) => ({
                ...it,
                wrote: it.wrote || (isWrite && !ev.error),
                segments: (it.segments || []).map((sg) =>
                  sg.type === "step" && sg.id === ev.call.id
                    ? { ...sg, status: ev.error ? "error" : "ok", detail: ev.error || ev.result?.message }
                    : sg
                ),
              }));
            } else if (ev.type === "limit") {
              appendText(`\n(Đã đạt giới hạn ${ev.maxSteps} bước. Gửi "tiếp tục" nếu muốn AI làm tiếp.)`);
            }
          },
        });
        historyRef.current = msgs;
        if (log.length) patchLast((it) => ({ ...it, undo: log }));
      } else {
        const history = items
          .map((it) => ({ role: it.role, content: it.role === "user" ? it.text : finalText(it) }))
          .filter((m) => m.content);
        let answer = "";
        await streamChat({
          provider: p,
          key: storage.getKey(p.id),
          model: modelId,
          messages: [{ role: "system", content: PLAIN_SYSTEM }, ...history, { role: "user", content }],
          signal: ac.signal,
          onDelta: (d) => {
            answer += d;
            appendText(d);
          },
        });
        historyRef.current = [...historyRef.current, { role: "user", content }, { role: "assistant", content: answer }];
      }
    } catch (e: any) {
      const aborted = e instanceof DOMException && e.name === "AbortError";
      const msg = aborted
        ? "Đã dừng."
        : isToolUnsupportedError(e)
          ? "Model này không gọi được công cụ nên không thao tác được trên tài liệu. Hãy chọn model khác (OpenRouter: bật “Chỉ model thao tác được”) hoặc tắt “AI tự thao tác” để chỉ trò chuyện."
          : networkErrorMessage(e);
      patchLast((it) => ({ ...it, error: msg }));
    } finally {
      setBusy(false);
      setPending(null);
      abortRef.current = null;
    }
  };

  const stop = () => {
    abortRef.current?.abort();
    pending?.resolve(false);
    setPending(null);
  };

  const undo = async (index: number) => {
    const it = items[index];
    if (!it.undo?.length) return;
    let ok = 0;
    const errors: string[] = [];
    for (const u of it.undo.slice().reverse()) {
      try {
        await u.undo();
        ok++;
      } catch (e) {
        errors.push(`${u.label}: ${networkErrorMessage(e)}`);
      }
    }
    setItems((prev) => prev.map((x, i) => (i === index ? { ...x, undone: true } : x)));
    setNotice(
      errors.length
        ? { text: `Hoàn tác ${ok}/${it.undo.length} thao tác. Lỗi: ${errors.join("; ")}`, err: true }
        : { text: `Đã hoàn tác ${ok} thao tác.` }
    );
  };

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

  const newChat = () => {
    setItems([]);
    historyRef.current = [];
    setNotice(undefined);
  };

  const model = models.find((m) => m.id === modelId);
  const isLastBusy = (i: number) => busy && i === items.length - 1;

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
          <Switch label="Miễn phí" checked={freeOnly} onChange={(_, d) => setFreeOnly(d.checked)} />
        )}
        {providerId === "openrouter" && agentOn && (
          <Switch
            label="Chỉ model thao tác được"
            checked={settings.toolsOnly}
            onChange={(_, d) => updateSettings({ toolsOnly: d.checked })}
          />
        )}
      </div>
      <Select value={modelId} onChange={(_, d) => setModelId(d.value)}>
        {shown.map((m) => (
          <option key={m.id} value={m.id}>
            {m.free ? "🆓 " : ""}
            {m.tools ? "🛠 " : ""}
            {m.name}
          </option>
        ))}
      </Select>
      <Text size={200} className={s.muted}>
        {catalogNote} {shown.length !== models.length ? `· đang hiện ${shown.length}/${models.length}` : ""}
      </Text>
      <Switch
        label={hostTools ? `AI tự thao tác trong ${host}` : "AI tự thao tác (chỉ trong Excel/Word/PowerPoint)"}
        checked={agentOn}
        disabled={!hostTools}
        onChange={(_, d) => updateSettings({ agentMode: d.checked })}
      />
      {agentOn && model?.tools === false && (
        <Text size={200} className={s.err}>
          Model này không gọi được công cụ → chỉ trò chuyện. Chọn model có biểu tượng 🛠 để AI tự thao tác.
        </Text>
      )}

      <div className={s.list} ref={listRef}>
        {!items.length && agentOn && (
          <div className={s.row} style={{ flexDirection: "column", alignItems: "stretch" }}>
            <Text size={200} className={s.muted}>
              Gợi ý (bấm để gửi):
            </Text>
            {(QUICK_PROMPTS[host] || []).map((q) => (
              <Button key={q} size="small" className={s.chip} onClick={() => send(q)} disabled={busy || !modelId}>
                {q}
              </Button>
            ))}
          </div>
        )}
        {items.map((it, i) =>
          it.role === "user" ? (
            <div key={i} className={s.user}>
              {it.text}
            </div>
          ) : (
            <div key={i} className={s.bot}>
              {(it.segments || []).map((sg, k) =>
                sg.type === "text" ? (
                  <Rich key={k} text={sg.text} />
                ) : (
                  <span key={k} className={sg.status === "error" ? s.stepErr : s.step}>
                    {sg.status === "running" ? "⏳" : sg.status === "ok" ? "✅" : "⚠️"} {sg.label}
                    {sg.detail && sg.status === "error" ? ` — ${sg.detail}` : ""}
                  </span>
                )
              )}
              {isLastBusy(i) && !pending && <Spinner size="tiny" />}
              {pending && i === items.length - 1 && (
                <div className={s.approve}>
                  <Text size={200}>AI muốn: {describeCall(pending.call)}</Text>
                  <div className={s.row} style={{ marginTop: 4 }}>
                    <Button size="small" appearance="primary" onClick={() => answerApproval(true)}>
                      Cho phép
                    </Button>
                    <Button size="small" onClick={() => answerApproval(true, true)}>
                      Cho phép hết lần này
                    </Button>
                    <Button size="small" onClick={() => answerApproval(false)}>
                      Bỏ qua
                    </Button>
                  </div>
                </div>
              )}
              {it.error && <Text className={s.err}>{it.error}</Text>}
              {!isLastBusy(i) && (
                <div className={s.row} style={{ marginTop: 4 }}>
                  {it.undo?.length && !it.undone ? (
                    <Button size="small" onClick={() => undo(i)}>
                      Hoàn tác thay đổi ({it.undo.length})
                    </Button>
                  ) : null}
                  {it.undone && (
                    <Text size={200} className={s.muted}>
                      Đã hoàn tác.
                    </Text>
                  )}
                  {finalText(it) && (
                    <>
                      <Button size="small" onClick={() => insert(finalText(it))}>
                        {INSERT_LABEL[host]}
                      </Button>
                      <Button size="small" onClick={() => copy(finalText(it))}>
                        Sao chép
                      </Button>
                    </>
                  )}
                </div>
              )}
              {it.wrote && host === "Word" && settings.trackChanges && !isLastBusy(i) && (
                <Text size={200} className={s.muted}>
                  Sửa đổi được ghi bằng Theo dõi thay đổi — vào thẻ Review để Chấp nhận / Từ chối.
                </Text>
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
        placeholder={agentOn ? `Yêu cầu AI làm gì trong ${host}… (Ctrl+Enter để gửi)` : "Nhập câu hỏi… (Ctrl+Enter để gửi)"}
        value={input}
        onChange={(_, d) => setInput(d.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey) && !busy) send();
        }}
        resize="vertical"
        rows={3}
      />
      <div className={s.row}>
        {busy ? (
          <Button onClick={stop}>Dừng</Button>
        ) : (
          <Button appearance="primary" onClick={() => send()} disabled={!input.trim() || !modelId}>
            Gửi
          </Button>
        )}
        <Button onClick={newChat} disabled={busy || !items.length}>
          Cuộc trò chuyện mới
        </Button>
      </div>
    </div>
  );
};

export default Chat;
