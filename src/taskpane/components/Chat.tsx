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
import { networkErrorMessage } from "../../core/errors";
import { AgentMessage, ToolCall, UndoEntry } from "../../core/agent/types";
import { compactHistory, runAgent } from "../../core/agent/loop";
import { callModelStep, isToolUnsupportedError } from "../../core/agent/llm";
import { compactIfNeeded } from "../../core/agent/compaction";
import { getKV } from "../../core/db";
import { ConversationMeta, ConversationStore, newConversationId, slimForStorage, titleFrom } from "../../core/conversations";
import { Command, Skill, SkillStore, resolveSlash, skillCatalog } from "../../core/skills";
import { CORE_GUIDES, autoSelectSkills } from "../../core/skillpack";
import { Attachment, readAttachment, withAttachments } from "../../core/attachments";
import { publishSnapshot } from "../../core/share";
import { RefTarget, findRefs } from "../../core/refs";
import { parseMarkdownTable } from "../../core/table";
import { HostName, INSERT_LABEL, insertAnswer, readSelection } from "../../office/host";
import { QUICK_PROMPTS, describeCall, getHostTools } from "../../office/tools";
import { createCommonTools } from "../../office/common";
import { navigateTo } from "../../office/navigate";

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
    minWidth: 0,
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
  risky: { color: tokens.colorPaletteRedForeground1, fontWeight: 600 },
  chip: { whiteSpace: "normal", textAlign: "left", height: "auto", minHeight: "28px" },
  muted: { color: tokens.colorNeutralForeground3 },
  err: { color: tokens.colorPaletteRedForeground1 },
  link: {
    color: tokens.colorBrandForegroundLink,
    textDecorationLine: "underline",
    cursor: "pointer",
  },
  tableWrap: { overflowX: "auto", maxWidth: "100%" },
  table: {
    borderCollapse: "collapse",
    fontSize: "12px",
    "& td, & th": { border: `1px solid ${tokens.colorNeutralStroke2}`, padding: "2px 4px", verticalAlign: "top" },
    "& th": { background: tokens.colorNeutralBackground4, fontWeight: 600 },
  },
  menu: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    maxHeight: "180px",
    overflowY: "auto",
    background: tokens.colorNeutralBackground1,
  },
  menuItem: {
    padding: "4px 6px",
    cursor: "pointer",
    whiteSpace: "normal",
    ":hover": { background: tokens.colorNeutralBackground1Hover },
  },
  fileChip: {
    fontSize: "12px",
    padding: "2px 6px",
    borderRadius: "10px",
    background: tokens.colorNeutralBackground4,
    display: "inline-flex",
    gap: "4px",
    alignItems: "center",
    maxWidth: "100%",
    overflow: "hidden",
  },
  histItem: {
    display: "flex",
    gap: "6px",
    alignItems: "center",
    padding: "6px",
    borderRadius: "6px",
    background: tokens.colorNeutralBackground3,
    minWidth: 0,
  },
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
  /** Số thao tác có Hoàn tác (giữ khi mở lại hội thoại cũ) */
  undoCount?: number;
  wrote?: boolean;
  error?: string;
  files?: string[];
  label?: string;
  skills?: string[];
}

interface Props {
  storage: Storage;
  host: HostName;
  version: number;
  goSettings: () => void;
}

const PLAIN_SYSTEM =
  "Bạn là trợ lý trong Microsoft Office. Trả lời bằng tiếng Việt, ngắn gọn, đúng trọng tâm. Khi trình bày dữ liệu dạng bảng, dùng bảng Markdown.";

const WEB_PROVIDERS = (p: { id: string; kind: string }) => p.id === "openrouter" || p.kind === "anthropic" || p.kind === "gemini";

/** Chữ thường + liên kết bấm được tới ô/đoạn/slide */
const Linked: React.FC<{ text: string; host: string; onNav: (r: RefTarget) => void; cls: string }> = ({ text, host, onNav, cls }) => (
  <>
    {findRefs(text, host).map((p, i) =>
      p.ref ? (
        <span
          key={i}
          className={cls}
          role="link"
          tabIndex={0}
          title="Bấm để chuyển tới"
          onClick={() => onNav(p.ref!)}
          onKeyDown={(e) => e.key === "Enter" && onNav(p.ref!)}
        >
          {p.text}
        </span>
      ) : (
        <React.Fragment key={i}>{p.text}</React.Fragment>
      )
    )}
  </>
);

/** Hiển thị Markdown đơn giản: **đậm**, tiêu đề #, gạch đầu dòng, bảng; địa chỉ ô/đoạn/slide bấm được */
const Rich: React.FC<{ text: string; host: string; onNav: (r: RefTarget) => void }> = ({ text, host, onNav }) => {
  const s = useStyles();
  const inline = (str: string) =>
    str.split(/(\*\*[^*]+\*\*)/g).map((part, i) =>
      part.startsWith("**") && part.endsWith("**") && part.length > 4 ? (
        <strong key={i}>
          <Linked text={part.slice(2, -2)} host={host} onNav={onNav} cls={s.link} />
        </strong>
      ) : (
        <Linked key={i} text={part} host={host} onNav={onNav} cls={s.link} />
      )
    );
  const lines = text.split("\n");
  const out: React.ReactNode[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.trim().startsWith("|")) {
      const block: string[] = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) block.push(lines[i++]);
      i--;
      const rows = parseMarkdownTable(block.join("\n"));
      if (rows && rows.length) {
        out.push(
          <div key={`t${i}`} className={s.tableWrap}>
            <table className={s.table}>
              <tbody>
                {rows.map((r, ri) => (
                  <tr key={ri}>
                    {r.map((c, ci) =>
                      ri === 0 ? <th key={ci}>{inline(c)}</th> : <td key={ci}>{inline(c)}</td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
        continue;
      }
      block.forEach((b, k) => out.push(<div key={`r${i}-${k}`}>{inline(b)}</div>));
      continue;
    }
    const h = line.match(/^\s*#{1,4}\s+(.*)$/);
    if (h) {
      out.push(
        <div key={i} style={{ fontWeight: 600, marginTop: 4 }}>
          {inline(h[1])}
        </div>
      );
      continue;
    }
    const b = line.match(/^\s*[-*•]\s+(.*)$/);
    if (b) {
      out.push(
        <div key={i} style={{ paddingLeft: 14, textIndent: -10 }}>
          • {inline(b[1])}
        </div>
      );
      continue;
    }
    out.push(<div key={i}>{line.trim() ? inline(line) : " "}</div>);
  }
  return <span style={{ whiteSpace: "normal" }}>{out}</span>;
};

const finalText = (it: Item) =>
  it.segments
    ? it.segments
        .filter((sg) => sg.type === "text")
        .map((sg: any) => sg.text)
        .join("\n")
        .trim()
    : it.text;

/** Hội thoại gửi model ở chế độ trò chuyện thường: bỏ lời gọi công cụ */
const plainHistory = (h: AgentMessage[]): AgentMessage[] =>
  h
    .filter((m) => m.role === "user" || (m.role === "assistant" && m.content))
    .map((m) => (m.role === "assistant" ? { role: "assistant", content: m.content } : { role: "user", content: m.content })) as AgentMessage[];

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
  const [pending, setPending] = React.useState<{ call: ToolCall; reason?: string; resolve: (ok: boolean) => void } | null>(null);
  const [attachments, setAttachments] = React.useState<Attachment[]>([]);
  const [attachBusy, setAttachBusy] = React.useState(false);
  const [skills, setSkills] = React.useState<Skill[]>([]);
  const [commands, setCommands] = React.useState<Command[]>([]);
  const [convId, setConvId] = React.useState("");
  const [convList, setConvList] = React.useState<ConversationMeta[]>([]);
  const [showHistory, setShowHistory] = React.useState(false);
  const abortRef = React.useRef<AbortController | null>(null);
  const historyRef = React.useRef<AgentMessage[]>([]);
  const allowAllRef = React.useRef(false);
  const listRef = React.useRef<HTMLDivElement>(null);
  const fileRef = React.useRef<HTMLInputElement>(null);
  const inputRef = React.useRef<HTMLTextAreaElement>(null);
  const busyRef = React.useRef(false);
  const prevBusy = React.useRef(false);

  const kv = React.useMemo(() => getKV(), []);
  const convStore = React.useMemo(() => new ConversationStore(kv), [kv]);
  const skillStore = React.useMemo(() => new SkillStore(kv), [kv]);

  const hostTools = React.useMemo(() => getHostTools(host, () => storage.getSettings()), [host, storage]);
  const agentOn = settings.agentMode && !!hostTools;

  const refreshConvs = React.useCallback(async () => {
    try {
      setConvList(await convStore.list(host));
    } catch {
      setConvList([]);
    }
  }, [convStore, host]);

  React.useEffect(() => {
    (async () => {
      try {
        setSkills(await skillStore.forHost(host));
        setCommands(await skillStore.commands());
      } catch {
        /* bộ nhớ không dùng được → không có skill người dùng */
      }
    })();
    refreshConvs();
  }, [skillStore, host, version, refreshConvs]);

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

  /** Chia sẻ nội dung đang mở cho ứng dụng Office khác (đọc chéo) */
  const share = React.useCallback(async () => {
    if (!settings.shareContext || !hostTools?.snapshot || busyRef.current) return;
    try {
      const snap = await hostTools.snapshot();
      publishSnapshot(host, snap.data, snap.title);
    } catch {
      /* tài liệu đang bận → lần sau */
    }
  }, [settings.shareContext, hostTools, host]);

  React.useEffect(() => {
    if (!settings.shareContext) return undefined;
    share();
    const t = setInterval(share, 30000);
    return () => clearInterval(t);
  }, [share, settings.shareContext]);

  /** Lưu hội thoại sau mỗi lượt (khi AI chạy xong) */
  React.useEffect(() => {
    busyRef.current = busy;
    if (prevBusy.current && !busy && items.length) {
      const id = convId || newConversationId();
      if (!convId) setConvId(id);
      const firstUser = items.find((it) => it.role === "user");
      convStore
        .save({
          id,
          host,
          title: titleFrom(firstUser?.text || ""),
          updatedAt: Date.now(),
          items: items.map((it) => ({ ...it, undo: undefined, undoCount: it.undo?.length || it.undoCount })),
          history: slimForStorage(historyRef.current),
        })
        .then(refreshConvs)
        .catch(() => undefined);
      share();
    }
    prevBusy.current = busy;
  }, [busy]); // eslint-disable-line react-hooks/exhaustive-deps

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

  const provider = getProvider(providerId);
  const webCapable = !!provider && WEB_PROVIDERS(provider);

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

  /** Thao tác rủi ro (reason) luôn hỏi từng lần; thao tác ghi thường có thể "cho phép hết lần này" */
  const askApproval = (call: ToolCall, reason?: string) =>
    !reason && allowAllRef.current
      ? Promise.resolve(true)
      : new Promise<boolean>((resolve) => setPending({ call, reason, resolve }));

  const answerApproval = (ok: boolean, all = false) => {
    if (all) allowAllRef.current = true;
    pending?.resolve(ok);
    setPending(null);
  };

  const onNav = (ref: RefTarget) =>
    navigateTo(host, ref).catch((e) => setNotice({ text: "Không chuyển tới được: " + networkErrorMessage(e), err: true }));

  const onFiles = async (files: FileList | null) => {
    if (!files || !files.length) return;
    setAttachBusy(true);
    const out: Attachment[] = [];
    const errs: string[] = [];
    for (const f of Array.from(files)) {
      try {
        out.push(await readAttachment(f));
      } catch (e: any) {
        errs.push(e?.message || `Không đọc được ${f.name}`);
      }
    }
    setAttachments((prev) => [...prev, ...out]);
    setAttachBusy(false);
    if (errs.length) setNotice({ text: errs.join(" · "), err: true });
    if (fileRef.current) fileRef.current.value = "";
  };

  const send = async (text?: string) => {
    const p = getProvider(providerId);
    const typed = (text ?? input).trim();
    if (!p || !modelId || !typed) return;
    setNotice(undefined);
    const slash = resolveSlash(typed, skills, commands);
    let content = slash.content;
    if (withSelection) {
      try {
        const selection = await readSelection();
        if (selection.text.trim()) content = `${content}\n\n--- Nội dung đang chọn ---\n${selection.text}`;
        setNotice({ text: selection.note });
      } catch (e) {
        setNotice({ text: "Không đọc được nội dung đang chọn: " + networkErrorMessage(e), err: true });
      }
    }
    const atts = attachments;
    content = withAttachments(content, atts);
    const images = atts.filter((a) => a.kind === "image" && a.data).map((a) => ({ mime: a.mime || "image/png", data: a.data! }));
    const userMsg: AgentMessage = images.length ? { role: "user", content, images } : { role: "user", content };

    const model = models.find((m) => m.id === modelId);
    const useAgent = agentOn && !!hostTools && model?.tools !== false;
    const auto = useAgent && !slash.label ? autoSelectSkills(typed, skills as any) : [];

    setItems((prev) => [
      ...prev,
      { role: "user", text: typed, files: atts.map((a) => a.name), label: slash.label },
      { role: "assistant", text: "", segments: [], skills: auto.map((k) => k.name) },
    ]);
    setInput("");
    setAttachments([]);
    setBusy(true);
    allowAllRef.current = false;
    const ac = new AbortController();
    abortRef.current = ac;
    const key = storage.getKey(p.id);
    const instr = (settings.instructions || {})[host]?.trim();

    try {
      const comp = await compactIfNeeded({ history: historyRef.current, provider: p, key, model: modelId, signal: ac.signal });
      if (comp.compacted) {
        historyRef.current = comp.history;
        setNotice({ text: "Cuộc trò chuyện dài: đã tự tóm tắt phần đầu để AI nhớ gọn hơn." });
      }

      if (useAgent && hostTools) {
        let ov: any;
        try {
          ov = await hostTools.overview();
        } catch (e) {
          ov = { error: "Không đọc được trạng thái tài liệu: " + networkErrorMessage(e) };
        }
        const system = [
          hostTools.systemPrompt(ov),
          CORE_GUIDES[host] || "",
          instr ? `## Chỉ dẫn riêng của người dùng cho ${host} (luôn tuân theo)\n${instr}` : "",
          skillCatalog(skills),
          auto.length
            ? `## Kỹ năng đã nạp sẵn cho yêu cầu này (làm theo)\n${auto.map((k) => `<skill name="${k.name}">\n${k.body}\n</skill>`).join("\n")}`
            : "",
        ]
          .filter(Boolean)
          .join("\n\n");
        const log: UndoEntry[] = [];
        let needSep = false;
        const msgs = await runAgent({
          provider: p,
          key,
          model: modelId,
          tools: hostTools,
          extra: createCommonTools({ host, skills, shareContext: settings.shareContext }),
          system,
          messages: [...compactHistory(historyRef.current), userMsg],
          signal: ac.signal,
          log,
          approve: askApproval,
          askBeforeWrite: settings.askBeforeWrite,
          confirmRisky: settings.confirmRisky,
          webSearch: settings.webSearch && webCapable,
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
        const system = [PLAIN_SYSTEM, instr ? `## Chỉ dẫn riêng của người dùng (luôn tuân theo)\n${instr}` : ""].filter(Boolean).join("\n\n");
        const res = await callModelStep({
          provider: p,
          key,
          model: modelId,
          tools: [],
          messages: [{ role: "system", content: system }, ...plainHistory(historyRef.current), userMsg],
          signal: ac.signal,
          onText: (d) => appendText(d),
          webSearch: settings.webSearch && webCapable,
        });
        historyRef.current = [...historyRef.current, userMsg, { role: "assistant", content: res.text }];
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
    setConvId("");
    setNotice(undefined);
    setShowHistory(false);
  };

  const openConv = async (id: string) => {
    try {
      const c = await convStore.get(id);
      if (!c) return;
      setItems(c.items as Item[]);
      historyRef.current = c.history || [];
      setConvId(c.id);
      setShowHistory(false);
      setNotice({ text: "Đã mở lại hội thoại. Nút Hoàn tác chỉ dùng được trong phiên làm việc đã tạo ra thay đổi." });
    } catch (e) {
      setNotice({ text: "Không mở được hội thoại: " + networkErrorMessage(e), err: true });
    }
  };

  const deleteConv = async (id: string) => {
    await convStore.delete(id).catch(() => undefined);
    if (id === convId) newChat();
    refreshConvs();
  };

  const model = models.find((m) => m.id === modelId);
  const isLastBusy = (i: number) => busy && i === items.length - 1;

  const slashQuery = /^\/([\w-]*)$/.exec(input)?.[1];
  const slashItems =
    slashQuery !== undefined
      ? [
          ...commands.map((c) => ({ name: c.name, desc: c.description || "Lệnh của tôi", kind: "Lệnh" })),
          ...skills.map((k) => ({ name: k.name, desc: k.description, kind: "Kỹ năng" })),
        ]
          .filter((x) => x.name.toLowerCase().includes(slashQuery.toLowerCase()))
          .slice(0, 15)
      : [];

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
          autoComplete="off"
          name="aidn-model-search"
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
      <div className={s.row}>
        <Switch
          label={hostTools ? `AI tự thao tác trong ${host}` : "AI tự thao tác (chỉ trong Excel/Word/PowerPoint)"}
          checked={agentOn}
          disabled={!hostTools}
          onChange={(_, d) => updateSettings({ agentMode: d.checked })}
        />
        {webCapable && (
          <Switch label="🌐 Tìm trên web" checked={settings.webSearch} onChange={(_, d) => updateSettings({ webSearch: d.checked })} />
        )}
      </div>
      {agentOn && model?.tools === false && (
        <Text size={200} className={s.err}>
          Model này không gọi được công cụ → chỉ trò chuyện. Chọn model có biểu tượng 🛠 để AI tự thao tác.
        </Text>
      )}
      {settings.webSearch && webCapable && provider?.kind === "gemini" && agentOn && (
        <Text size={200} className={s.muted}>
          Gemini chỉ tìm trên web khi tắt “AI tự thao tác”.
        </Text>
      )}

      {showHistory ? (
        <div className={s.list}>
          <div className={s.row}>
            <Text weight="semibold" style={{ flex: 1 }}>
              Lịch sử trong {host} ({convList.length})
            </Text>
            <Button size="small" onClick={() => setShowHistory(false)}>
              Đóng
            </Button>
          </div>
          {!convList.length && (
            <Text size={200} className={s.muted}>
              Chưa có hội thoại nào được lưu.
            </Text>
          )}
          {convList.map((c) => (
            <div key={c.id} className={s.histItem}>
              <div style={{ flex: 1, minWidth: 0, cursor: "pointer" }} onClick={() => openConv(c.id)} title="Mở lại">
                <Text size={300} style={{ display: "block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {c.id === convId ? "▶ " : ""}
                  {c.title}
                </Text>
                <Text size={100} className={s.muted}>
                  {new Date(c.updatedAt).toLocaleString("vi-VN")}
                </Text>
              </div>
              <Button size="small" appearance="subtle" onClick={() => deleteConv(c.id)} title="Xóa hội thoại">
                Xóa
              </Button>
            </div>
          ))}
        </div>
      ) : (
        <div className={s.list} ref={listRef}>
          {!items.length && agentOn && (
            <div className={s.row} style={{ flexDirection: "column", alignItems: "stretch" }}>
              <Text size={200} className={s.muted}>
                Gợi ý (bấm để gửi) · gõ “/” để gọi kỹ năng:
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
                {it.files?.length ? (
                  <div>
                    <Text size={100} className={s.muted}>
                      📎 {it.files.join(", ")}
                    </Text>
                  </div>
                ) : null}
              </div>
            ) : (
              <div key={i} className={s.bot}>
                {it.skills?.length ? (
                  <Text size={100} className={s.muted}>
                    Kỹ năng áp dụng: {it.skills.join(", ")}
                  </Text>
                ) : null}
                {(it.segments || []).map((sg, k) =>
                  sg.type === "text" ? (
                    <Rich key={k} text={sg.text} host={host} onNav={onNav} />
                  ) : (
                    <span key={k} className={sg.status === "error" ? s.stepErr : s.step}>
                      {sg.status === "running" ? "⏳" : sg.status === "ok" ? "✅" : "⚠️"}{" "}
                      <Linked text={sg.label} host={host} onNav={onNav} cls={s.link} />
                      {sg.detail && sg.status === "error" ? ` — ${sg.detail}` : ""}
                    </span>
                  )
                )}
                {isLastBusy(i) && !pending && <Spinner size="tiny" />}
                {pending && i === items.length - 1 && (
                  <div className={s.approve}>
                    {pending.reason && <Text className={s.risky}>⚠️ {pending.reason}</Text>}
                    <div>
                      <Text size={200}>AI muốn: {describeCall(pending.call)}</Text>
                    </div>
                    <div className={s.row} style={{ marginTop: 4 }}>
                      <Button size="small" appearance="primary" onClick={() => answerApproval(true)}>
                        Cho phép
                      </Button>
                      {!pending.reason && (
                        <Button size="small" onClick={() => answerApproval(true, true)}>
                          Cho phép hết lần này
                        </Button>
                      )}
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
      )}

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
      {attachments.length > 0 && (
        <div className={s.row}>
          {attachments.map((a, i) => (
            <span key={i} className={s.fileChip} title={a.note || a.name}>
              {a.kind === "image" ? "🖼" : "📄"} {a.name}
              <span
                role="button"
                tabIndex={0}
                style={{ cursor: "pointer" }}
                title="Bỏ tệp"
                onClick={() => setAttachments((prev) => prev.filter((_, k) => k !== i))}
              >
                ✕
              </span>
            </span>
          ))}
        </div>
      )}
      {slashItems.length > 0 && (
        <div className={s.menu}>
          {slashItems.map((x) => (
            <div
              key={x.kind + x.name}
              className={s.menuItem}
              onClick={() => {
                setInput(`/${x.name} `);
                inputRef.current?.focus();
              }}
            >
              <Text size={200} weight="semibold">
                /{x.name}
              </Text>{" "}
              <Text size={100} className={s.muted}>
                {x.kind} · {x.desc.slice(0, 110)}
              </Text>
            </div>
          ))}
        </div>
      )}
      <Textarea
        ref={inputRef}
        placeholder={
          agentOn ? `Yêu cầu AI làm gì trong ${host}… (gõ / để gọi kỹ năng · Ctrl+Enter để gửi)` : "Nhập câu hỏi… (Ctrl+Enter để gửi)"
        }
        value={input}
        onChange={(_, d) => setInput(d.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey) && !busy) send();
        }}
        resize="vertical"
        rows={3}
      />
      <input
        ref={fileRef}
        type="file"
        multiple
        accept=".pdf,.docx,.txt,.md,.csv,.tsv,.json,.xml,.html,.htm,.yaml,.yml,.log,image/png,image/jpeg,image/gif,image/webp"
        style={{ display: "none" }}
        onChange={(e) => onFiles(e.target.files)}
      />
      <div className={s.row}>
        {busy ? (
          <Button onClick={stop}>Dừng</Button>
        ) : (
          <Button appearance="primary" onClick={() => send()} disabled={!input.trim() || !modelId || attachBusy}>
            Gửi
          </Button>
        )}
        <Button onClick={() => fileRef.current?.click()} disabled={busy || attachBusy} title="Đính kèm PDF, DOCX, TXT, CSV, ảnh…">
          {attachBusy ? "Đang đọc…" : "📎 Tệp"}
        </Button>
        <Button onClick={newChat} disabled={busy || (!items.length && !convId)}>
          Mới
        </Button>
        <Button
          onClick={() => {
            refreshConvs();
            setShowHistory((v) => !v);
          }}
          disabled={busy}
        >
          Lịch sử{convList.length ? ` (${convList.length})` : ""}
        </Button>
      </div>
    </div>
  );
};

export default Chat;
