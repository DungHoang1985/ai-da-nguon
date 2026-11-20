import * as React from "react";
import {
  Button,
  Menu,
  MenuDivider,
  MenuItem,
  MenuItemCheckbox,
  MenuItemRadio,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Spinner,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  AddRegular,
  AttachRegular,
  ChatAddRegular,
  HistoryRegular,
  SendRegular,
  SettingsRegular,
  ShieldCheckmarkRegular,
  StopRegular,
} from "@fluentui/react-icons";
import ModelPicker from "./ModelPicker";
import { allProviders, getProvider } from "../../core/providers";
import { UsageMode, normalizeFreeOrder, providerInMode } from "../../core/modes";
import { FREE_STALL_MS, FREE_STALL_THINK_MS, FreeRotation, buildFreeCandidates } from "../../core/agent/rotation";
import { fetchModels, isStale } from "../../core/models";
import { Storage } from "../../core/storage";
import { ErrorReason, describeError, networkErrorMessage } from "../../core/errors";
import { chatCreditWarning } from "../../core/credits";
import { useCredits } from "./useCredits";
import { WelcomeSignIn } from "./SignInCard";
import { useCatalog } from "./useCatalog";
import { AgentMessage, ToolCall, UndoEntry, Usage } from "../../core/agent/types";
import { isToolUnsupportedError } from "../../core/agent/llm";
import { stepCapFor } from "../../core/agent/loop";
import { addUsage, formatUsage, runTurn } from "../../core/agent/turn";
import { getKV } from "../../core/db";
import { ConversationMeta, ConversationStore, MAX_CONVERSATIONS, newConversationId, slimForStorage, testModeOn, titleFrom } from "../../core/conversations";
import { Command, Skill, SkillStore, resolveSlash, skillCatalog } from "../../core/skills";
import { CORE_GUIDES, autoSelectSkills } from "../../core/skillpack";
import { WORD_TOOL_GROUPS } from "../../core/toolgroups";
import { Attachment, DOCX_MIME, MAX_FILES_TOTAL_B64, PDF_MIME, filesBase64Chars, readAttachment, withAttachments } from "../../core/attachments";
import { attachmentPrivacyNote, canSeePages, routeScanPdf } from "../../core/pdfPages";
import { clearAttachedFiles, registerAttachedFile } from "../../core/attachedFiles";
import { pdfCapable, thinkingToggleWorks } from "../../core/agent/thinking";
import { publishSnapshot } from "../../core/share";
import { REDO_WITH_TOOLS, unverifiedWriteClaim, writeSucceeded } from "../../core/claims";
import { APPS, CMD_PREFIX, Connection, HEARTBEAT_MS, beat, claimCommand, clearBeat, connectionLine, connections, describePayload, normalizeApp, reply, sweep, takeCommand } from "../../core/crossapp";
import { executeIncoming } from "../../office/crossExec";
import { RefTarget, findRefs } from "../../core/refs";
import { parseMarkdownTable } from "../../core/table";
import { HostName, INSERT_LABEL, insertAnswer, readSelection } from "../../office/host";
import { QUICK_PROMPTS, describeCall, getHostTools } from "../../office/tools";
import { createCommonTools } from "../../office/common";
import { navigateTo } from "../../office/navigate";
import { APP_VERSION, APP_VERSION_FULL } from "../../core/version";
import { inputKeyAction } from "../../core/keys";
import { copyText } from "../fileio";

const useStyles = makeStyles({
  wrap: { display: "flex", flexDirection: "column", gap: "6px", height: "100%", minWidth: 0 },
  row: { display: "flex", gap: "6px", alignItems: "center", flexWrap: "wrap", minWidth: 0 },
  /** Đầu khung một dòng: tên + phiên bản · Lịch sử · Hội thoại mới · Cài đặt */
  header: { display: "flex", alignItems: "center", gap: "4px", flexShrink: 0, minWidth: 0, minHeight: "28px" },
  title: { whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  /** Vùng hội thoại chiếm toàn bộ phần giữa */
  list: {
    flex: 1,
    overflowY: "auto",
    overflowX: "hidden",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    minHeight: "80px",
  },
  fixed: { flexShrink: 0 },
  composer: { display: "flex", gap: "6px", alignItems: "flex-end", flexShrink: 0, minWidth: 0 },
  input: { flex: 1, minWidth: 0, "& textarea": { maxHeight: "132px", overflowY: "auto" } },
  toolbar: { display: "flex", gap: "2px", alignItems: "center", flexShrink: 0, minWidth: 0, flexWrap: "wrap" },
  modelSlot: { flex: 1, minWidth: 0, display: "flex", justifyContent: "flex-end" },
  tag: {
    fontSize: "12px",
    padding: "1px 6px",
    borderRadius: "10px",
    background: tokens.colorNeutralBackground4,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "nowrap",
  },
  welcome: { flex: 1, display: "flex", flexDirection: "column", gap: "10px", justifyContent: "center", padding: "8px" },
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
  /** Cảnh báo "AI nói đã sửa nhưng không ghi": phải nổi hơn chữ thường để không tin nhầm */
  warnBox: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    alignItems: "flex-start",
    padding: "6px 8px",
    borderRadius: "4px",
    borderLeft: `3px solid ${tokens.colorPaletteDarkOrangeBorder2}`,
    backgroundColor: tokens.colorPaletteYellowBackground1,
    color: tokens.colorNeutralForeground1,
    fontWeight: 600,
  },
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
    // Không cho khung cột co về 0 khi khung add-in thấp
    flexShrink: 0,
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
  | { type: "step"; id: string; label: string; status: "running" | "ok" | "error"; detail?: string; /** Một dòng kết quả của bước chỉ đọc (soát, so sánh…) */ brief?: string };

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
  /** Chi tiết kỹ thuật của lỗi (nút "Chép chi tiết") */
  errorDetail?: string;
  /** Lỗi về key hoặc tiền: hiện nút Đổi nguồn / Mở Cài đặt */
  errorReason?: ErrorReason;
  files?: string[];
  label?: string;
  skills?: string[];
  /** Số token của lượt (cộng mọi bước) */
  usage?: Usage;
  /** Số lần gọi model của nhiệm vụ này (mỗi bước của vòng "model → công cụ → kết quả" là một lần gọi) */
  calls?: number;
  /** Mục do khung ở ứng dụng khác nhờ ghi (không phải câu trả lời của model; không vào lịch sử gửi model) */
  incoming?: string;
  /** AI nói đã thay đổi tài liệu nhưng lượt này không có thao tác ghi nào thành công (chỉ để hiện; không gửi cho model) */
  claimWarning?: string;
  /** Lượt dừng vì chạm trần số bước hoặc vì một công cụ lặp cùng một lỗi: hiện thẻ "Làm tiếp?" (chỉ để hiện) */
  paused?: { steps: number; writes: number; stuck?: string };
  /** Số lần gọi AI của các lượt trước thuộc CÙNG nhiệm vụ (khi người dùng bấm Làm tiếp) */
  callsBase?: number;
}

/** Câu gửi model khi người dùng bấm "Làm tiếp" */
export const CONTINUE_TASK =
  "[Add-in: lượt trước do ADD-IN dừng lại (chạm trần số bước của một lượt, hoặc một công cụ lặp cùng một lỗi) — không phải do trợ lý tự dừng. Người dùng vừa bấm “Làm tiếp”.] " +
  "Làm tiếp nhiệm vụ đang dở từ chỗ vừa dừng và làm cho tới khi XONG HẲN nhiệm vụ. Không làm lại các bước đã có kết quả thành công; xong thì báo ngắn gọn đã làm những gì.";

interface Props {
  storage: Storage;
  host: HostName;
  version: number;
  onOpenSettings: () => void;
  /** Vừa đăng nhập ở màn chào: báo App dựng lại để danh sách nguồn có key được đọc lại */
  onKeysChanged?: () => void;
}

const WEB_PROVIDERS = (p: { id: string; kind: string; dialect?: { webSearch?: boolean } }) =>
  p.id === "openrouter" || (p.kind === "anthropic" && (!p.dialect || !!p.dialect.webSearch)) || p.kind === "gemini";

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

const Chat: React.FC<Props> = ({ storage, host, version, onOpenSettings, onKeysChanged }) => {
  const s = useStyles();
  const [settings, setSettingsState] = React.useState(storage.getSettings());
  const updateSettings = (patch: any) => {
    storage.setSettings(patch);
    setSettingsState(storage.getSettings());
  };
  React.useEffect(() => setSettingsState(storage.getSettings()), [version, storage]);

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
  /** Yêu cầu gần nhất người dùng gõ (không tính "Làm tiếp"): cột "Yêu cầu của lượt" của Nhật ký AI */
  const lastRequestRef = React.useRef<string>("");
  const allowAllRef = React.useRef(false);
  const listRef = React.useRef<HTMLDivElement>(null);
  const fileRef = React.useRef<HTMLInputElement>(null);
  const inputRef = React.useRef<HTMLTextAreaElement>(null);
  const [slashIdx, setSlashIdx] = React.useState(0);
  /** Esc đóng danh sách "/" cho đến khi nội dung ô nhập đổi */
  const [slashClosedFor, setSlashClosedFor] = React.useState<string | null>(null);
  const busyRef = React.useRef(false);
  const prevBusy = React.useRef(false);
  /** Bản mới nhất của các mục hiển thị và mã hội thoại, để lưu theo từng bước mà không chờ React vẽ lại */
  const itemsRef = React.useRef<Item[]>([]);
  itemsRef.current = items;
  const convIdRef = React.useRef("");
  const saveChain = React.useRef<Promise<void>>(Promise.resolve());
  const lastSaveAt = React.useRef(0);
  /** Token đầu vào thực đo của bước gần nhất: đo độ đầy ngữ cảnh để quyết định tóm tắt */
  const lastInputRef = React.useRef<number | undefined>(undefined);
  const setConv = (id: string) => {
    convIdRef.current = id;
    setConvId(id);
  };

  // Ô nhập tự giãn 1–6 dòng
  React.useLayoutEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = `${Math.min(el.scrollHeight + 2, 132)}px`;
  }, [input]);

  const kv = React.useMemo(() => getKV(), []);
  /** Chế độ thử (bài thử tự động bật qua sessionStorage): không lưu lịch sử */
  const testMode = React.useMemo(() => testModeOn(), []);
  const convStore = React.useMemo(() => new ConversationStore(kv, testModeOn), [kv]);
  /** Lịch sử đầy: hội thoại cũ nhất sắp bị xóa khi hội thoại mới này được lưu */
  const [histWarn, setHistWarn] = React.useState("");
  const skillStore = React.useMemo(() => new SkillStore(kv), [kv]);

  const hostTools = React.useMemo(() => getHostTools(host, () => storage.getSettings()), [host, storage]);
  const agentOn = settings.agentMode && !!hostTools;
  /** Đợt 9: chế độ Miễn phí / Trả phí; chưa chọn → hiện thẻ chọn chế độ */
  const [mode, setModeState] = React.useState<UsageMode | undefined>(() => storage.getSelection().mode);
  /** Nguồn tải danh mục model lỗi lần gần nhất (để không tải lại mỗi lần gửi) */
  const catalogFailRef = React.useRef<Record<string, number>>({});
  const setMode = (m: UsageMode) => {
    storage.setSelection({ mode: m });
    setModeState(m);
  };
  // Đổi chế độ hoặc thứ tự xoay vòng trong Cài đặt: đọc lại
  React.useEffect(() => {
    const m = storage.getSelection().mode;
    if (m !== mode) setModeState(m);
  }, [version]); // eslint-disable-line react-hooks/exhaustive-deps
  const cat = useCatalog(storage, { agentOn, toolsOnly: settings.toolsOnly, version, mode });
  const { available, providerId, modelId, models, shown } = cat;
  /** Mở bảng chọn nguồn và model từ nút "Đổi nguồn" cạnh thông báo lỗi */
  const openPickerRef = React.useRef<(() => void) | null>(null);
  /** Hạn mức OpenRouter: đọc lại sau mỗi lượt (thưa nhất 60 giây một lần) để cảnh báo khi thấp */
  const credits = useCredits(storage, getProvider(providerId));
  const creditWarn = chatCreditWarning(credits.info, { freeModel: !!cat.model?.free, balanceAllowed: credits.balanceAllowed });

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

  // ----- Ghi chéo ứng dụng: nhịp sống, chỉ báo kết nối, nhận lệnh ghi hộ (core/crossapp.ts) -----
  const selfApp = normalizeApp(host);
  /** Các ứng dụng khác đang bật kết nối trong Cài đặt của khung này */
  const linkedApps = React.useMemo(
    () => (settings.shareContext ? APPS.filter((a) => a !== selfApp && (settings.crossApps || {})[a] !== false) : []),
    [settings.shareContext, settings.crossApps, selfApp]
  );
  const [links, setLinks] = React.useState<Connection[]>([]);
  const seenCmds = React.useRef(new Set<string>());
  const acceptRef = React.useRef(false);
  acceptRef.current = !!settings.shareContext && settings.crossWrite !== false && !!hostTools;
  const linkedRef = React.useRef<string[]>([]);
  linkedRef.current = linkedApps;

  React.useEffect(() => {
    if (!selfApp || !settings.shareContext) {
      setLinks([]);
      return undefined;
    }
    let ls: typeof window.localStorage;
    try {
      ls = window.localStorage;
    } catch {
      return undefined;
    }
    const title = (() => {
      try {
        const url = String((globalThis as any).Office?.context?.document?.url || "");
        return decodeURIComponent(url.split(/[\\/]/).pop() || "").slice(0, 80) || undefined;
      } catch {
        return undefined;
      }
    })();
    const tick = () => {
      beat(ls, { host: selfApp, title, accept: acceptRef.current });
      const cs = connections(ls, selfApp).filter((c) => linkedRef.current.includes(c.app));
      setLinks((prev) => (JSON.stringify(prev) === JSON.stringify(cs) ? prev : cs));
    };
    /** Nhận một lệnh ghi hộ (nếu có): chạy bằng công cụ ghi của ứng dụng này, trả kết quả, hiện một mục có nút Hoàn tác */
    let working = false;
    const token = Math.random().toString(36).slice(2);
    const receive = async () => {
      // Khung đang chạy một lượt của chính nó (hoặc đang hỏi người dùng): để lệnh nằm trong hộp thư, lát nữa lấy; khung gửi tự rút nếu chờ quá lâu
      if (working || !hostTools || busyRef.current) return;
      const cmd = takeCommand(ls, selfApp, { accept: acceptRef.current, seen: seenCmds.current, title });
      if (!cmd) return;
      working = true;
      if (!(await claimCommand(ls, cmd.id, token))) {
        working = false;
        return;
      }
      if (!linkedRef.current.includes(cmd.from)) {
        reply(ls, cmd.id, { ok: false, error: `Khung AI Đa Nguồn ở ${selfApp} đang tắt kết nối với ${cmd.from}.` });
        working = false;
        return;
      }
      // Trong lúc ghi hộ, khung này coi như đang bận: người dùng chưa gửi được yêu cầu mới (tránh hai luồng ghi chen nhau)
      busyRef.current = true;
      setBusy(true);
      const log: UndoEntry[] = [];
      const what = describePayload(cmd.payload);
      try {
        const message = await executeIncoming(hostTools, cmd.payload, log);
        reply(ls, cmd.id, { ok: true, message });
        setItems((prev) => [...prev, { role: "assistant", text: "", incoming: `Vừa ghi theo yêu cầu từ khung AI Đa Nguồn ở ${cmd.from}: ${what}.`, wrote: true, undo: log.length ? log : undefined }]);
      } catch (e: any) {
        // Ghi dở thì gỡ phần đã làm rồi mới báo lỗi, để khung gửi không tưởng là đã ghi
        for (const u of log.slice().reverse()) await u.undo().catch(() => undefined);
        reply(ls, cmd.id, { ok: false, error: `${selfApp} không ghi được: ${networkErrorMessage(e)}` });
      } finally {
        working = false;
        busyRef.current = false;
        setBusy(false);
      }
    };
    const onStorage = (e: StorageEvent) => {
      if (e.key === CMD_PREFIX + selfApp && e.newValue) receive();
    };
    tick();
    sweep(ls as any);
    window.addEventListener("storage", onStorage);
    const t1 = setInterval(tick, HEARTBEAT_MS);
    // Lối dự phòng khi lỡ sự kiện storage
    const t2 = setInterval(receive, 1000);
    const bye = () => clearBeat(ls, selfApp);
    window.addEventListener("pagehide", bye);
    return () => {
      clearInterval(t1);
      clearInterval(t2);
      window.removeEventListener("storage", onStorage);
      window.removeEventListener("pagehide", bye);
      bye();
    };
  }, [selfApp, settings.shareContext, hostTools]);

  /** Lưu hội thoại xuống máy. Gọi sau mỗi bước của AI (để Office có đổ giữa chừng vẫn còn phần đã làm) và khi lượt kết thúc. */
  const saveConversation = React.useCallback((): Promise<void> => {
    const its = itemsRef.current;
    if (!its.length) return saveChain.current;
    if (!convIdRef.current) {
      convIdRef.current = newConversationId();
      setConvId(convIdRef.current);
    }
    const firstUser = its.find((it) => it.role === "user");
    const conv = {
      id: convIdRef.current,
      host,
      title: titleFrom(firstUser?.text || ""),
      updatedAt: Date.now(),
      items: its.map((it) => ({ ...it, undo: undefined, undoCount: it.undo?.length || it.undoCount })),
      history: slimForStorage(historyRef.current),
    };
    lastSaveAt.current = Date.now();
    // Xếp hàng để hai lần lưu không ghi đè chỉ mục của nhau
    saveChain.current = saveChain.current.then(() => convStore.save(conv)).catch(() => undefined);
    return saveChain.current;
  }, [convStore, host]);

  /** Lưu hội thoại khi lượt kết thúc (xong, bị dừng hoặc lỗi) */
  React.useEffect(() => {
    busyRef.current = busy;
    if (prevBusy.current && !busy && items.length) {
      saveConversation().then(refreshConvs);
      share();
      credits.reload(false);
    }
    prevBusy.current = busy;
  }, [busy]); // eslint-disable-line react-hooks/exhaustive-deps

  // Đợt 9: chọn chế độ khi bắt đầu (cả người dùng bản cũ); chọn xong không hiện lại, đổi được ở bảng chọn model và Cài đặt
  if (!mode) {
    return (
      <div className={s.wrap}>
        <div className={s.header}>
          <Text weight="semibold" size={300}>
            AI Đa Nguồn
          </Text>
          <Text size={100} className={s.muted}>
            v{APP_VERSION}
          </Text>
          <div style={{ flex: 1 }} />
          <Button appearance="subtle" size="small" icon={<SettingsRegular />} title="Cài đặt" aria-label="Cài đặt" onClick={onOpenSettings} />
        </div>
        <div className={s.welcome} data-aidn="chon-che-do">
          <Text size={500} weight="semibold">
            Chọn chế độ dùng AI
          </Text>
          <Button appearance="primary" size="large" onClick={() => setMode("free")} style={{ justifyContent: "flex-start" }} data-mode="free">
            🆓 Miễn phí
          </Button>
          <Text size={200} className={s.muted}>
            Dùng key gói miễn phí của Gemini, Groq, OpenRouter (model :free), Cohere. Nguồn hết lượt hoặc không phản hồi thì tự đổi sang nguồn miễn phí kế
            tiếp. Gói miễn phí có giới hạn lượt và nhà cung cấp có thể dùng nội dung gửi lên để cải thiện dịch vụ: không gửi tài liệu mật.
          </Text>
          <Button appearance="primary" size="large" onClick={() => setMode("paid")} style={{ justifyContent: "flex-start" }} data-mode="paid">
            💳 Trả phí
          </Button>
          <Text size={200} className={s.muted}>
            Dùng key trả tiền (Anthropic Claude, OpenAI, DeepSeek, OpenRouter…). Tính tiền theo key của anh, không tự đổi nguồn.
          </Text>
          <Text size={100} className={s.muted}>
            Đổi chế độ lúc nào cũng được: bấm nhãn model cạnh ô nhập.
          </Text>
        </div>
      </div>
    );
  }

  // Màn chào: chỉ hiện khi chưa có key nào (của chế độ đang bật)
  if (!available.length) {
    const otherKeys = allProviders().some((p) => storage.getKeys()[p.id] && !providerInMode(p, mode));
    return (
      <div className={s.wrap}>
        <div className={s.header}>
          <Text weight="semibold" size={300}>
            AI Đa Nguồn
          </Text>
          <Text size={100} className={s.muted}>
            v{APP_VERSION}
          </Text>
          <div style={{ flex: 1 }} />
          <Button appearance="subtle" size="small" icon={<SettingsRegular />} title="Cài đặt" aria-label="Cài đặt" onClick={onOpenSettings} />
        </div>
        <div className={s.welcome}>
          <Text size={500} weight="semibold">
            Chào anh 👋
          </Text>
          {otherKeys && (
            <Text data-aidn="che-do-chua-co-key">
              Chế độ {mode === "free" ? "Miễn phí" : "Trả phí"} chưa có nguồn nào có key.{" "}
              <Button size="small" onClick={() => setMode(mode === "free" ? "paid" : "free")}>
                Chuyển sang {mode === "free" ? "Trả phí" : "Miễn phí"}
              </Button>
            </Text>
          )}
          <Text>
            AI Đa Nguồn làm việc ngay trong {host === "Khác" ? "Word, Excel, PowerPoint" : host}: đọc, sửa, tạo nội dung theo yêu cầu, bằng API
            key của anh (OpenRouter, Anthropic, Gemini, Groq…).
          </Text>
          <Text size={200} className={s.muted}>
            Key chỉ lưu trên máy này và chỉ gửi tới đúng nhà cung cấp. OpenRouter có nhiều model miễn phí.
          </Text>
          {getProvider("openrouter")?.oauth && (
            <WelcomeSignIn p={getProvider("openrouter")!} storage={storage} onChanged={() => onKeysChanged?.()} />
          )}
          <Button onClick={onOpenSettings} style={{ alignSelf: "flex-start" }}>
            Nhập API key của nhà cung cấp khác
          </Button>
          <Text size={200} className={s.muted}>
            Từng dùng AI Đa Nguồn trên máy này mà giờ không thấy key, lịch sử? Office lưu dữ liệu add-in theo từng tài khoản Office: nếu vừa đổi tài khoản
            hoặc Office vừa cập nhật, dữ liệu cũ nằm ở kho khác. Vào Cài đặt → Sao lưu → “Khôi phục từ tệp…”.
          </Text>
          <Button onClick={onOpenSettings} style={{ alignSelf: "flex-start" }}>
            Khôi phục từ tệp sao lưu
          </Button>
        </div>
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
        const att = await readAttachment(f);
        // Bản gốc của tệp .docx: để công cụ Word so sánh phiên bản đọc được (giữ trong bộ nhớ của phiên này)
        if (att.mime === DOCX_MIME && att.data) registerAttachedFile(att.name, att.data);
        // PDF có trang quét (hoặc bị cắt): giữ bản gốc để AI đọc theo trang bằng read_attachment_pages
        if (att.mime === PDF_MIME && att.data) registerAttachedFile(att.name, att.data);
        out.push(att);
      } catch (e: any) {
        errs.push(e?.message || `Không đọc được ${f.name}`);
      }
    }
    setAttachments((prev) => [...prev, ...out]);
    setAttachBusy(false);
    if (errs.length) setNotice({ text: errs.join(" · "), err: true });
    if (fileRef.current) fileRef.current.value = "";
  };

  const send = async (text?: string, opts?: { continueTask?: boolean }) => {
    // Nguồn và model lấy tại lúc gửi: đổi nguồn khi đang chạy chỉ có hiệu lực từ lượt sau
    const p = getProvider(providerId);
    const typed = (text ?? input).trim();
    if (!p || !typed) return;
    if (cat.blockReason) {
      setNotice({ text: cat.blockReason, err: true });
      return;
    }
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
    // Bấm "Làm tiếp": không gửi tệp đang chờ, không xóa chữ đang gõ dở (soát mã Đợt 7, #14)
    const atts = opts?.continueTask ? [] : attachments;
    // Yêu cầu thật của người dùng cho cột "Yêu cầu của lượt" trong Nhật ký AI: lượt Làm tiếp dùng lại yêu cầu gốc (soát mã Đợt 7, #5)
    const request = opts?.continueTask ? lastRequestRef.current || typed : typed;
    if (!opts?.continueTask) lastRequestRef.current = typed;
    const model = models.find((m) => m.id === modelId);
    const vision = model?.vision ?? p.vision;
    const useAgent = agentOn && !!hostTools && model?.tools !== false;
    const auto = useAgent && !slash.label ? autoSelectSkills(typed, skills as any) : [];
    // PDF quét đi đường nào: nguyên tệp (tệp ngắn, nguồn đọc được) · AI đọc theo trang bằng công cụ · không đọc được
    const scans = atts.filter((a) => a.kind === "file" && a.data);
    const routeOf = (a: Attachment) => routeScanPdf({ pages: a.pages ?? 1, pdfCapable: pdfCapable(p), vision, tools: useAgent });
    const byPages = scans.filter((a) => routeOf(a) === "pages");
    const whole = scans.filter((a) => routeOf(a) !== "pages");
    // Tệp gửi nguyên vẹn nằm lại trong lịch sử và được gửi lại ở mọi lượt: quá trần thì nhà cung cấp từ chối mọi lượt sau
    const newFiles = whole.reduce((n, a) => n + a.data!.length, 0);
    if (newFiles && filesBase64Chars(historyRef.current as any) + newFiles > MAX_FILES_TOTAL_B64) {
      setNotice({
        text: "Các tệp PDF quét trong hội thoại này đã quá lớn để gửi tiếp (tệp được gửi lại ở mỗi lượt). Bấm “Hội thoại mới” rồi đính kèm lại tệp cần hỏi, hoặc tách tệp nhỏ hơn.",
        err: true,
      });
      return;
    }
    // Bản gốc của tệp .docx và .pdf: đăng ký (lại) lúc gửi để công cụ so sánh phiên bản và công cụ đọc theo trang đọc được
    atts.forEach((a) => (a.mime === DOCX_MIME || a.mime === PDF_MIME) && a.data && registerAttachedFile(a.name, a.data));
    content = withAttachments(content, atts, { compare: host === "Word" && agentOn, pageTool: useAgent, byPages: byPages.map((a) => a.name) });
    const images = atts.filter((a) => a.kind === "image" && a.data).map((a) => ({ mime: a.mime || "image/png", data: a.data! }));
    // PowerPoint: giữ ảnh đính kèm trong bộ nhớ của phiên và nêu tên để AI đặt được ảnh vào slide (ppt_insert_image)
    if (host === "PowerPoint" && useAgent) {
      const named = atts.filter((a) => a.kind === "image" && a.data);
      named.forEach((a) => registerAttachedFile(a.name, a.data!));
      if (named.length) content = `${content}\n\n[Ảnh đính kèm: ${named.map((a) => `“${a.name}”`).join(", ")}. Muốn đặt ảnh vào slide thì gọi ppt_insert_image với image = tên ảnh.]`;
    }
    // PDF quét gửi nguyên tệp cho model tự đọc; nguồn không đọc được thì báo rõ (và model cũng được báo trong tin nhắn)
    const pdfs = whole.map((a) => ({ mime: a.mime || "application/pdf", data: a.data!, name: a.name }));
    if (pdfs.length && !pdfCapable(p)) {
      setNotice({
        text: `${p.name} không đọc được PDF quét (${pdfs.map((f) => f.name).join(", ")}): tệp không được gửi đi. ${vision === true && !useAgent ? "Bật “AI tự thao tác” để AI đọc từng trang bằng ảnh, hoặc đổi" : "Đổi"} sang nguồn Anthropic hoặc Gemini rồi đính kèm lại.`,
        err: true,
      });
    } else if (atts.length) {
      // Nhắc riêng tư (không chặn): tệp đính kèm sắp tới nguồn miễn phí hoặc qua nơi trung chuyển
      const priv = attachmentPrivacyNote({ providerId: p.id, providerName: p.name, modelId, free: model?.free });
      if (priv) setNotice({ text: priv });
    }


    setItems((prev) => {
      // Bấm "Làm tiếp": vẫn là nhiệm vụ cũ → số lần gọi AI cộng dồn; thẻ hỏi của lượt trước được gỡ
      const last = prev[prev.length - 1];
      const base = opts?.continueTask && last?.role === "assistant" ? (last.callsBase || 0) + (last.calls || 0) : 0;
      return [
        ...prev.map((x) => (x.paused ? { ...x, paused: undefined } : x)),
        { role: "user", text: opts?.continueTask ? "Làm tiếp" : typed, files: atts.map((a) => a.name), label: slash.label },
        { role: "assistant", text: "", segments: [], skills: auto.map((k) => k.name), callsBase: base || undefined },
      ];
    });
    // Hội thoại mới mà lịch sử đã đầy: báo trước hội thoại cũ nhất sẽ bị xóa (trước đây xóa lặng lẽ)
    if (!convIdRef.current) {
      convStore
        .nextToDrop()
        .then((d) => setHistWarn(d ? `Lịch sử đã đầy (${MAX_CONVERSATIONS} hội thoại): hội thoại cũ nhất “${d.title}” bị xóa để lấy chỗ. Muốn giữ hội thoại nào thì xóa bớt hội thoại khác trong Lịch sử.` : ""))
        .catch(() => undefined);
    }
    if (!opts?.continueTask) {
      setInput("");
      setAttachments([]);
    }
    setShowHistory(false);
    setBusy(true);
    allowAllRef.current = false;
    const ac = new AbortController();
    abortRef.current = ac;
    const key = storage.getKey(p.id);
    const instr = (settings.instructions || {})[host]?.trim();
    const log: UndoEntry[] = [];
    const extraTools = createCommonTools({ host, skills, shareContext: settings.shareContext, apps: linkedApps, signal: ac.signal, seesPages: canSeePages({ pdfCapable: pdfCapable(p), vision }) });
    /** Công cụ đã gọi trong lượt và số lệnh ghi không thành công — để nhận ra ca "nói đã sửa mà không ghi" */
    const called: string[] = [];
    let failedWrites = 0;
    let needSep = false;
    let retrying = false;
    const clearRetry = () => {
      if (!retrying) return;
      retrying = false;
      setNotice(undefined);
    };

    // Đợt 9: chế độ Miễn phí → mọi lần gọi model của lượt đi qua vòng xoay nguồn miễn phí; nguồn đang chọn đứng đầu
    let rotation: FreeRotation | undefined;
    if (mode === "free") {
      const keys = storage.getKeys();
      const order = normalizeFreeOrder(storage.getSelection().freeOrder, allProviders());
      // Nguồn trong vòng chưa có danh mục model (hoặc đã cũ): tải trước để chọn được model. Tối đa 8 giây, bấm Dừng là thôi chờ;
      // nguồn tải lỗi thì 10 phút sau mới thử tải lại (soát mã Đợt 9, M2)
      const fetches = order
        .filter((pid) => keys[pid] && pid !== p.id)
        .map(async (pid) => {
          const c = storage.getModelCache()[pid];
          const pp = getProvider(pid);
          if (!pp || (c && !isStale(c.at)) || (catalogFailRef.current[pid] || 0) > Date.now() - 600000) return;
          try {
            storage.setModels(pid, await fetchModels(pp, keys[pid]));
          } catch {
            catalogFailRef.current[pid] = Date.now();
          }
        });
      await Promise.race([
        Promise.all(fetches),
        new Promise((r) => setTimeout(r, 8000)),
        new Promise((r) => ac.signal.addEventListener("abort", r, { once: true })),
      ]);
      if (ac.signal.aborted) {
        setBusy(false);
        abortRef.current = null;
        return;
      }
      const candidates = buildFreeCandidates({
        order,
        selected: { providerId: p.id, modelId },
        keys,
        getProvider,
        modelCache: storage.getModelCache(),
        remembered: (pid) => storage.getModelFor(pid),
        agentOn: useAgent,
      });
      try {
        rotation = new FreeRotation(candidates, {
          store: { get: () => storage.getCooldowns(), set: (v) => storage.setCooldowns(v) },
          stallMs: settings.deepThink ? FREE_STALL_THINK_MS : FREE_STALL_MS,
        });
      } catch (e) {
        const text = describeError(e, { provider: p }).text;
        patchLast((it) => ({ ...it, error: text }));
        setNotice({ text, err: true });
        // Trả lại chữ đã gõ để gửi lại không phải gõ lại (soát mã Đợt 9, L4)
        if (!opts?.continueTask) {
          setInput(typed);
          setAttachments(atts);
        }
        setBusy(false);
        abortRef.current = null;
        return;
      }
      // Chỉ có một nguồn miễn phí dùng được: nhắc nhập thêm key để có chỗ đổi khi hết lượt (dot-9 R9)
      if (candidates.length < 2)
        setNotice({ text: `Chế độ Miễn phí mới có một nguồn dùng được (${candidates[0].provider.name}): hết lượt là phải dừng. Nhập thêm key miễn phí (Gemini, Groq, OpenRouter, Cohere) trong Cài đặt để add-in tự đổi nguồn.` });
    }

    try {
      const res = await runTurn({
        rotation,
        provider: p,
        key,
        model: modelId,
        context: model?.context,
        maxOutput: model?.maxOutput,
        vision,
        host,
        history: historyRef.current,
        text: content,
        request,
        images: images.length ? images : undefined,
        files: pdfs.length ? pdfs : undefined,
        think: !!settings.deepThink,
        instructions: instr,
        agent:
          useAgent && hostTools
            ? {
                tools: hostTools,
                extra: extraTools,
                guide: CORE_GUIDES[host],
                skillCatalog: skillCatalog(skills),
                autoSkills: auto.map((k) => ({ name: k.name, body: k.body })),
                askBeforeWrite: settings.askBeforeWrite,
                confirmRisky: settings.confirmRisky,
                reviewEdits: host === "Word" && !!settings.reviewEdits,
                approve: askApproval,
                log,
                maxSteps: stepCapFor(host),
                // Đợt 10: chế độ Miễn phí chỉ gửi công cụ lõi, nhóm chuyên bật theo việc (Trả phí gửi đủ để giữ bộ nhớ đệm)
                toolGroups: mode === "free" && host === "Word" ? WORD_TOOL_GROUPS : undefined,
                groupHints: [...(atts.some((a) => a.mime === PDF_MIME) ? ["ngoai"] : []), ...(atts.some((a) => a.mime === DOCX_MIME) ? ["soat"] : [])],
              }
            : undefined,
        webSearch: settings.webSearch && webCapable,
        lastInputTokens: lastInputRef.current,
        signal: ac.signal,
        onHistory: (h) => {
          historyRef.current = h;
          // Lưu theo từng bước, thưa nhất 1,5 giây một lần; lần lưu cuối lượt mới là bản đầy đủ
          if (Date.now() - lastSaveAt.current > 1500) setTimeout(saveConversation, 0);
        },
        onEvent: (ev) => {
          if (ev.type === "compacted") {
            setNotice({ text: "Cuộc trò chuyện dài: đã tự tóm tắt phần đầu để AI nhớ gọn hơn." });
          } else if (ev.type === "step") {
            if (ev.step > 1) needSep = true;
            patchLast((it) => ({ ...it, calls: Math.max(it.calls || 0, ev.step) }));
          } else if (ev.type === "text") {
            clearRetry();
            appendText(needSep ? "\n" + ev.delta : ev.delta);
            needSep = false;
          } else if (ev.type === "tool_start") {
            clearRetry();
            patchLast((it) => ({
              ...it,
              segments: [...(it.segments || []), { type: "step", id: ev.call.id, label: describeCall(ev.call), status: "running" }],
            }));
          } else if (ev.type === "tool_end") {
            const isWrite = [...(hostTools?.specs || []), ...extraTools.specs].some((t) => t.name === ev.call.name && t.write);
            called.push(ev.call.name);
            // Công cụ ghi báo lỗi, trả ok=false hoặc "thay 0 chỗ" thì tài liệu chưa đổi
            const done = isWrite && writeSucceeded(ev.result, ev.error);
            if (isWrite && !done) failedWrites++;
            patchLast((it) => ({
              ...it,
              wrote: it.wrote || done,
              segments: (it.segments || []).map((sg) =>
                sg.type === "step" && sg.id === ev.call.id
                  ? { ...sg, status: ev.error ? "error" : "ok", detail: ev.error || ev.result?.message, brief: !ev.error && typeof ev.result?.brief === "string" ? ev.result.brief : undefined }
                  : sg
              ),
            }));
          } else if (ev.type === "review_start") {
            clearRetry();
            patchLast((it) => ({
              ...it,
              segments: [...(it.segments || []), { type: "step", id: "review:" + ev.call.id, label: `Soát lập luận trước khi ghi: ${describeCall(ev.call)}`, status: "running" }],
            }));
          } else if (ev.type === "review_end") {
            patchLast((it) => ({
              ...it,
              segments: (it.segments || []).map((sg) =>
                sg.type === "step" && sg.id === "review:" + ev.call.id
                  ? { ...sg, status: ev.ok ? "ok" : "error", detail: ev.ok ? ev.note || "Đạt" : `Chưa ghi — ${ev.note || "có vấn đề"}`, brief: ev.ok ? ev.note || "đạt" : undefined }
                  : sg
              ),
            }));
          } else if (ev.type === "switch") {
            clearRetry();
            appendText(`
↪ ${ev.from}: ${ev.reason} — đổi sang ${ev.to}${ev.partial ? " (trả lời lại bước này)" : ""}.
`);
            needSep = false;
          } else if (ev.type === "thinking_off") {
            setNotice({ text: "Nguồn đang dùng không nhận chế độ suy nghĩ với model này: đã gửi lại yêu cầu ở chế độ thường." });
          } else if (ev.type === "limit") {
            patchLast((it) => ({ ...it, paused: { steps: ev.maxSteps, writes: ev.writes } }));
          } else if (ev.type === "stuck") {
            patchLast((it) => ({ ...it, paused: { steps: ev.step, writes: ev.writes, stuck: `${ev.tool} báo cùng một lỗi ${ev.times} bước liên tiếp` } }));
          } else if (ev.type === "retry") {
            retrying = true;
            const why = ev.reason.charAt(0).toUpperCase() + ev.reason.slice(1);
            setNotice({ text: `${why}: đang thử lại lần ${ev.attempt}/${ev.max} sau ${Math.round(ev.waitMs / 1000)} giây… Bấm Dừng để hủy.` });
          } else if (ev.type === "usage") {
            clearRetry();
            patchLast((it) => ({ ...it, usage: addUsage(it.usage || {}, ev.usage) }));
          } else if (ev.type === "truncated") {
            if (!ev.skipped) appendText("\n(Câu trả lời bị cắt vì quá dài. Gửi “tiếp tục” để AI viết tiếp.)");
            else if (ev.gaveUp)
              appendText("\n(Dừng lượt: câu trả lời bị cắt vì quá dài nhiều lần liên tiếp, các lệnh ghi dở chưa được thực hiện. Hãy nhờ AI làm từng phần nhỏ hơn.)");
          }
        },
      });
      historyRef.current = res.history;
      lastInputRef.current = res.lastInputTokens;
      // Kể cả khi bị dừng hay lỗi giữa chừng: các bước đã làm vẫn hoàn tác được
      if (log.length) patchLast((it) => ({ ...it, undo: log }));
      if (res.error) {
        const e: any = res.error;
        const info = describeError(e, { provider: p });
        const msg =
          info.kind === "aborted"
            ? "Đã dừng."
            : isToolUnsupportedError(e)
              ? "Model này không gọi được công cụ nên không thao tác được trên tài liệu. Hãy chọn model khác (OpenRouter: bật “Chỉ model thao tác được”) hoặc tắt “AI tự thao tác” để chỉ trò chuyện."
              : info.text;
        const detail = `${info.detail}\nModel: ${modelId}\nPhiên bản: ${APP_VERSION_FULL}\nLúc: ${new Date().toLocaleString("vi-VN")}`;
        patchLast((it) => ({ ...it, error: msg, errorDetail: info.kind === "aborted" ? undefined : detail, errorReason: info.reason }));
      } else if (useAgent && hostTools) {
        // Lượt kết thúc bình thường: AI khẳng định đã thay đổi tài liệu mà không có thao tác ghi nào thành công → cảnh báo trên khung
        patchLast((it) => {
          const w = unverifiedWriteClaim({ userText: typed, answer: finalText(it), wrote: !!it.wrote, toolsCalled: called, failedWrites });
          return w ? { ...it, claimWarning: w.text } : it;
        });
      }
    } catch (e: any) {
      // runTurn không ném lỗi; tới đây là lỗi trong chính add-in
      const info = describeError(e, { provider: p });
      patchLast((it) => ({ ...it, error: info.text, errorDetail: info.detail }));
    } finally {
      retrying = false;
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
    setNotice((await copyText(text)) ? { text: "Đã sao chép." } : { text: "Không sao chép được trong môi trường này.", err: true });
  };

  const newChat = () => {
    // Tệp .docx đính kèm để so sánh thuộc về hội thoại cũ (tệp đang chờ gửi được đăng ký lại lúc gửi)
    clearAttachedFiles();
    setItems([]);
    historyRef.current = [];
    lastRequestRef.current = "";
    lastInputRef.current = undefined;
    setConv("");
    setHistWarn("");
    setNotice(undefined);
    setShowHistory(false);
  };

  const openConv = async (id: string) => {
    try {
      const c = await convStore.get(id);
      if (!c) return;
      setItems(c.items as Item[]);
      historyRef.current = c.history || [];
      lastInputRef.current = undefined;
      setConv(c.id);
      setHistWarn("");
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
  const menuOpen = slashItems.length > 0 && slashClosedFor !== input;
  const activeSlash = Math.min(slashIdx, Math.max(0, slashItems.length - 1));
  const pickSlash = (name: string) => {
    setInput(`/${name} `);
    inputRef.current?.focus();
  };

  const onInputKey = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const act = inputKeyAction(
      {
        key: e.key,
        shiftKey: e.shiftKey,
        ctrlKey: e.ctrlKey,
        metaKey: e.metaKey,
        altKey: e.altKey,
        isComposing: (e.nativeEvent as KeyboardEvent).isComposing,
        keyCode: e.keyCode,
      },
      { menuOpen, busy }
    );
    if (act === "none" || act === "newline") return;
    e.preventDefault();
    if (act === "send") send();
    else if (act === "stop") stop();
    else if (act === "menu-next") setSlashIdx((activeSlash + 1) % slashItems.length);
    else if (act === "menu-prev") setSlashIdx((activeSlash - 1 + slashItems.length) % slashItems.length);
    else if (act === "menu-pick") pickSlash(slashItems[activeSlash].name);
    else if (act === "menu-close") setSlashClosedFor(input);
  };

  const selectionLabel = host === "Excel" ? "Gửi kèm vùng ô đang chọn" : "Gửi kèm nội dung đang chọn";
  const thinkWorks = !!provider && thinkingToggleWorks(provider, modelId);
  const plusChecked = [...(withSelection ? ["sel"] : []), ...(settings.webSearch && webCapable ? ["web"] : []), ...(settings.deepThink ? ["think"] : [])];

  return (
    <div className={s.wrap}>
      <div className={s.header}>
        <Text weight="semibold" size={300} className={s.title}>
          AI Đa Nguồn
        </Text>
        <Text size={100} className={s.muted} title={`Phiên bản ${APP_VERSION_FULL}`}>
          v{APP_VERSION}
        </Text>
        {testMode && (
          <Text size={100} className={s.err} data-test-mode="1" title="Bài thử tự động đang chạy: hội thoại trong phiên này không được lưu vào Lịch sử. Tự hết khi đóng Office.">
            chế độ thử — không lưu lịch sử
          </Text>
        )}
        {links.length > 0 && (
          <Text
            size={100}
            className={s.muted}
            data-links="1"
            style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", marginLeft: 4 }}
            title={
              "Kết nối với khung AI Đa Nguồn ở ứng dụng khác (✓ đang mở, — chưa mở):\n" +
              links.map((c) => `${c.app}: ${c.open ? `đang mở${c.title ? ` “${c.title}”` : ""}${c.accept ? "" : " (không nhận ghi)"}` : "chưa mở"}`).join("\n")
            }
          >
            {connectionLine(links)}
          </Text>
        )}
        <div style={{ flex: 1 }} />
        <Button
          appearance={showHistory ? "secondary" : "subtle"}
          size="small"
          icon={<HistoryRegular />}
          title="Lịch sử hội thoại"
          aria-label="Lịch sử hội thoại"
          aria-pressed={showHistory}
          disabled={busy}
          onClick={() => {
            refreshConvs();
            setShowHistory((v) => !v);
          }}
        />
        <Button
          appearance="subtle"
          size="small"
          icon={<ChatAddRegular />}
          title="Hội thoại mới"
          aria-label="Hội thoại mới"
          disabled={busy || (!items.length && !convId)}
          onClick={newChat}
        />
        <Button appearance="subtle" size="small" icon={<SettingsRegular />} title="Cài đặt" aria-label="Cài đặt" onClick={onOpenSettings} />
      </div>

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
          {!items.length && (
            <div className={s.row} style={{ flexDirection: "column", alignItems: "stretch" }}>
              <Text size={200} className={s.muted}>
                {agentOn
                  ? `AI làm việc trực tiếp trong ${host}. Gợi ý (bấm để gửi) · gõ “/” để gọi kỹ năng:`
                  : "Chế độ chỉ trò chuyện. Bật “AI tự thao tác” trong Cài đặt để AI sửa trực tiếp trên tài liệu."}
              </Text>
              {agentOn &&
                (QUICK_PROMPTS[host] || []).map((q) => (
                  <Button key={q} size="small" className={s.chip} onClick={() => send(q)} disabled={busy || !!cat.blockReason}>
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
                      {sg.detail && sg.status === "error" ? ` — ${sg.detail}` : sg.brief && sg.status === "ok" ? ` — ${sg.brief}` : ""}
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
                {it.usage && formatUsage(it.usage) ? (
                  <Text
                    size={100}
                    className={s.muted}
                    data-calls={(it.callsBase || 0) + (it.calls || 0) || undefined}
                    title="Một nhiệm vụ anh giao được AI làm qua nhiều bước nối tiếp: mỗi bước là một lần gọi AI (AI ra lệnh công cụ → add-in làm trên tài liệu → trả kết quả cho AI). Lần gọi nào cũng kèm toàn bộ yêu cầu và các bước trước, nên AI không mất mạch. Nhà cung cấp tính mỗi bước là một lượt dùng. Số token là tổng mọi bước; phần đọc từ bộ nhớ đệm được tính rẻ hơn nhiều."
                  >
                    {(it.callsBase || 0) + (it.calls || 0) > 1 ? `1 nhiệm vụ · ${(it.callsBase || 0) + (it.calls || 0)} lần gọi AI · ` : ""}
                    {formatUsage(it.usage)}
                  </Text>
                ) : null}
                {it.paused && !isLastBusy(i) && i === items.length - 1 && (
                  <div className={s.warnBox} role="status" data-paused="1">
                    <Text size={200}>
                      {it.paused.stuck
                        ? `AI dừng ở bước ${it.paused.steps}: ${it.paused.stuck}. Đã ghi ${it.paused.writes} thao tác vào tài liệu.`
                        : `AI đã làm ${it.paused.steps} bước (hết một lượt), đã ghi ${it.paused.writes} thao tác vào tài liệu. Làm tiếp?`}
                    </Text>
                    <div className={s.row}>
                      <Button size="small" appearance="primary" data-continue="1" onClick={() => send(CONTINUE_TASK, { continueTask: true })} disabled={busy || !!cat.blockReason}>
                        {it.paused.stuck ? "Thử tiếp" : "Làm tiếp"}
                      </Button>
                      <Button size="small" data-stop-here="1" onClick={() => setItems((prev) => prev.map((x, k) => (k === i ? { ...x, paused: undefined } : x)))}>
                        Dừng ở đây
                      </Button>
                    </div>
                  </div>
                )}
                {it.incoming && (
                  <Text size={200} data-incoming="1">
                    📥 {it.incoming}
                  </Text>
                )}
                {it.error && <Text className={s.err}>{it.error}</Text>}
                {it.claimWarning && (
                  <div className={s.warnBox} role="alert" data-claim-warning="1">
                    <Text size={200}>⚠ {it.claimWarning}</Text>
                    {!isLastBusy(i) && i === items.length - 1 && (
                      <div className={s.row}>
                        <Button size="small" appearance="primary" onClick={() => send(REDO_WITH_TOOLS)} disabled={busy || !!cat.blockReason}>
                          Yêu cầu làm lại bằng công cụ
                        </Button>
                        <Button size="small" onClick={() => openPickerRef.current?.()}>
                          Đổi model
                        </Button>
                      </div>
                    )}
                  </div>
                )}
                {it.error && it.errorReason && !isLastBusy(i) && i === items.length - 1 && (
                  <div className={s.row}>
                    {available.length > 1 && (
                      <Button size="small" appearance="primary" onClick={() => openPickerRef.current?.()}>
                        Đổi nguồn
                      </Button>
                    )}
                    <Button size="small" onClick={onOpenSettings}>
                      Mở Cài đặt
                    </Button>
                  </div>
                )}
                {it.error && it.errorDetail && !isLastBusy(i) && (
                  <Button size="small" appearance="subtle" style={{ alignSelf: "flex-start" }} onClick={() => copy(`${it.error}\n\n${it.errorDetail}`)}>
                    Chép chi tiết lỗi
                  </Button>
                )}
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
                {it.wrote && host === "Word" && !isLastBusy(i) && (
                  <Text size={200} className={s.muted}>
                    Sửa đổi được ghi bằng Theo dõi thay đổi — vào thẻ Review để Chấp nhận / Từ chối từng chỗ, hoặc bấm Hoàn tác thay đổi để gỡ cả lượt.
                  </Text>
                )}
              </div>
            )
          )}
        </div>
      )}

      {notice && (
        <Text size={200} className={`${s.fixed} ${notice.err ? s.err : s.muted}`}>
          {notice.text}
        </Text>
      )}
      {agentOn && model?.tools === false && (
        <Text size={200} className={`${s.fixed} ${s.err}`}>
          Model này không gọi được công cụ → chỉ trò chuyện. Chọn model có biểu tượng 🛠 để AI tự thao tác.
        </Text>
      )}
      {settings.webSearch && webCapable && provider?.kind === "gemini" && agentOn && (
        <Text size={200} className={`${s.fixed} ${s.muted}`}>
          Gemini chỉ tìm trên web khi tắt “AI tự thao tác”.
        </Text>
      )}
      {attachments.length > 0 && (
        <div className={`${s.row} ${s.fixed}`}>
          {attachments.map((a, i) => (
            <span key={i} className={s.fileChip} title={a.note || a.name}>
              {a.kind === "image" ? "🖼" : a.kind === "file" ? "📑" : "📄"} {a.name}
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
      {menuOpen && (
        <div className={s.menu} role="listbox" aria-label="Kỹ năng và lệnh">
          {slashItems.map((x, k) => (
            <div
              key={x.kind + x.name}
              role="option"
              aria-selected={k === activeSlash}
              className={s.menuItem}
              style={k === activeSlash ? { background: tokens.colorNeutralBackground1Selected } : undefined}
              onMouseEnter={() => setSlashIdx(k)}
              onClick={() => pickSlash(x.name)}
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

      <div className={s.composer}>
        <Textarea
          ref={inputRef}
          className={s.input}
          resize="none"
          rows={1}
          placeholder={agentOn ? `Yêu cầu AI làm gì trong ${host}…` : "Nhập câu hỏi…"}
          title="Enter gửi · Shift+Enter xuống dòng · gõ / để gọi kỹ năng"
          value={input}
          onChange={(_, d) => {
            setInput(d.value);
            setSlashIdx(0);
          }}
          onKeyDown={onInputKey}
        />
        {busy ? (
          <Button appearance="primary" icon={<StopRegular />} title="Dừng (Esc)" aria-label="Dừng" onClick={stop} />
        ) : (
          <Button
            appearance="primary"
            icon={<SendRegular />}
            title="Gửi (Enter)"
            aria-label="Gửi"
            onClick={() => send()}
            disabled={!input.trim() || !!cat.blockReason || attachBusy}
          />
        )}
      </div>

      <div className={s.toolbar}>
        <Menu
          checkedValues={{ opts: plusChecked }}
          onCheckedValueChange={(_, d) => {
            const v = d.checkedItems;
            setWithSelection(v.includes("sel"));
            updateSettings({ ...(webCapable ? { webSearch: v.includes("web") } : {}), deepThink: v.includes("think") });
          }}
        >
          <MenuTrigger disableButtonEnhancement>
            <Button
              size="small"
              appearance="subtle"
              icon={attachBusy ? <Spinner size="extra-tiny" /> : <AddRegular />}
              title="Đính kèm tệp, kỹ năng, tìm trên web, gửi kèm nội dung đang chọn"
              aria-label="Thêm: tệp, kỹ năng, tìm web, nội dung đang chọn"
            />
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<AttachRegular />} disabled={busy || attachBusy} onClick={() => fileRef.current?.click()}>
                Đính kèm tệp (PDF, DOCX, TXT, CSV, ảnh…)
              </MenuItem>
              <MenuItem
                onClick={() => {
                  setInput("/");
                  setSlashClosedFor(null);
                  setTimeout(() => inputRef.current?.focus(), 0);
                }}
              >
                / Kỹ năng và lệnh
              </MenuItem>
              <MenuDivider />
              <MenuItemCheckbox name="opts" value="sel">
                {selectionLabel}
              </MenuItemCheckbox>
              {webCapable && (
                <MenuItemCheckbox name="opts" value="web">
                  Tìm trên web
                </MenuItemCheckbox>
              )}
              <MenuItemCheckbox name="opts" value="think">
                Suy nghĩ kỹ (chậm hơn, tốn hơn)
              </MenuItemCheckbox>
            </MenuList>
          </MenuPopover>
        </Menu>
        {agentOn ? (
          <Menu
            checkedValues={{ perm: [settings.askBeforeWrite ? "ask" : "auto"] }}
            onCheckedValueChange={(_, d) => updateSettings({ askBeforeWrite: d.checkedItems[0] === "ask" })}
          >
            <MenuTrigger disableButtonEnhancement>
              <Button
                size="small"
                appearance="subtle"
                icon={<ShieldCheckmarkRegular />}
                title="Mức quyền của AI khi sửa tài liệu"
                aria-label={`Mức quyền: ${settings.askBeforeWrite ? "hỏi trước khi sửa" : "tự sửa, hỏi khi rủi ro"}`}
              >
                {settings.askBeforeWrite ? "Hỏi trước" : "Tự sửa"}
              </Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItemRadio name="perm" value="ask">
                  Hỏi trước mỗi thao tác sửa
                </MenuItemRadio>
                <MenuItemRadio name="perm" value="auto">
                  Tự sửa, hỏi khi rủi ro
                </MenuItemRadio>
                <MenuDivider />
                <MenuItem disabled>Ghi đè, xóa, duyệt tất cả: luôn hỏi kèm lý do</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        ) : (
          <Text size={200} className={s.tag} title="Bật “AI tự thao tác” trong Cài đặt để AI sửa trực tiếp trên tài liệu">
            Chỉ trò chuyện
          </Text>
        )}
        {withSelection && (
          <Text size={100} className={s.tag} title={selectionLabel} onClick={() => setWithSelection(false)} style={{ cursor: "pointer" }}>
            📄 Kèm vùng chọn ✕
          </Text>
        )}
        {settings.webSearch && webCapable && (
          <Text size={100} className={s.tag} title="Đang bật tìm trên web" onClick={() => updateSettings({ webSearch: false })} style={{ cursor: "pointer" }}>
            🌐 Web ✕
          </Text>
        )}
        {settings.deepThink && (
          <Text
            size={100}
            className={s.tag}
            title={thinkWorks ? "Đang bật Suy nghĩ kỹ: AI suy nghĩ nhiều hơn trước khi trả lời" : "Suy nghĩ kỹ không có tác dụng với nguồn hoặc model đang chọn"}
            onClick={() => updateSettings({ deepThink: false })}
            style={{ cursor: "pointer", ...(thinkWorks ? {} : { textDecoration: "line-through" }) }}
          >
            🧠 Nghĩ kỹ ✕
          </Text>
        )}
        <div className={s.modelSlot}>
          <ModelPicker cat={cat} agentOn={agentOn} toolsOnly={settings.toolsOnly} onToolsOnly={(v) => updateSettings({ toolsOnly: v })} openRef={openPickerRef} mode={mode} onMode={setMode} />
        </div>
      </div>
      {histWarn && (
        <Text size={100} className={`${s.fixed} ${s.muted}`} data-aidn="lich-su-day">
          ⚠️ {histWarn}
        </Text>
      )}
      {!busy && creditWarn && (
        <Text size={100} className={`${s.fixed} ${creditWarn.level === "out" ? s.err : s.muted}`} data-aidn="canh-bao-han-muc">
          ⚠️ {creditWarn.text}
        </Text>
      )}
      {!busy && cat.blockReason && (
        <Text size={100} className={`${s.fixed} ${cat.catalog.status === "error" ? s.err : s.muted}`}>
          {cat.blockReason}
        </Text>
      )}
      <input
        ref={fileRef}
        type="file"
        multiple
        accept=".pdf,.docx,.txt,.md,.csv,.tsv,.json,.xml,.html,.htm,.yaml,.yml,.log,image/png,image/jpeg,image/gif,image/webp"
        style={{ display: "none" }}
        onChange={(e) => onFiles(e.target.files)}
      />
    </div>
  );
};

export default Chat;
