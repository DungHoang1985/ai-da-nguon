import * as React from "react";
import { Button, Input, Link, Select, Spinner, Switch, Text, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowLeftRegular } from "@fluentui/react-icons";
import { APP_VERSION, APP_VERSION_FULL } from "../../core/version";
import { HostName } from "../../office/host";
import { getHostTools } from "../../office/tools";
import { TestStep, runSelfTest } from "../../office/selftest";
import { PROVIDERS, Provider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { fetchModels } from "../../core/models";
import { cleanKey, describeError, keyProblem, networkErrorMessage } from "../../core/errors";
import ProviderCheckCard from "./ProviderCheck";
import SignInCard from "./SignInCard";
import BackupCard from "./BackupCard";
import { getKV } from "../../core/db";
import { Command, Skill, SkillStore, importZip, parseCommandMd, parseSkillMd, toSkillName } from "../../core/skills";

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
  page: { display: "flex", flexDirection: "column", height: "100%", minHeight: 0 },
  header: { display: "flex", alignItems: "center", gap: "4px", flexShrink: 0, minHeight: "28px", paddingBottom: "4px" },
  scroll: { flex: 1, minHeight: 0, overflowY: "auto", overflowX: "hidden", paddingRight: "2px" },
  section: { fontSize: "13px", fontWeight: 600, margin: "10px 0 6px", color: tokens.colorNeutralForeground2 },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
});

interface Props {
  storage: Storage;
  onChanged: () => void;
  host?: HostName;
  /** Về khung chat (trang Cài đặt có nút quay lại) */
  onBack?: () => void;
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
    // Bỏ ký tự ẩn, khoảng trắng, dấu nháy lẫn vào khi dán; key còn ký tự lạ thì không lưu
    const k = cleanKey(key);
    if (!k) return;
    const problem = keyProblem(k);
    if (problem) {
      setStatus({ kind: "err", text: problem });
      return;
    }
    setKey(k);
    storage.setKey(p.id, k);
    setStatus({ kind: "busy", text: "Đang kiểm tra key…" });
    try {
      const models = await fetchModels(p, k);
      storage.setModels(p.id, models);
      setStatus({ kind: "ok", text: `Đã kết nối · ${models.length} model` });
    } catch (e) {
      setStatus({ kind: "err", text: describeError(e, { provider: p }).text });
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
      {p.blockedReason && (
        <Text size={200} className={s.err}>
          ⚠️ {p.blockedReason}
        </Text>
      )}
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

function useSettingsState(storage: Storage, onChanged: () => void) {
  const [st, setSt] = React.useState(storage.getSettings());
  const set = (patch: any) => {
    storage.setSettings(patch);
    setSt(storage.getSettings());
    onChanged();
  };
  return { st, set };
}

/** Quyền và an toàn. Word luôn ghi sửa đổi bằng Theo dõi thay đổi (không còn công tắc tắt — anh duyệt 15/11/2026). */
const PermissionsCard: React.FC<Props> = ({ storage, onChanged }) => {
  const s = useStyles();
  const { st, set } = useSettingsState(storage, onChanged);
  return (
    <div className={s.card}>
      <Switch label="AI tự thao tác: cho AI tự đọc, sửa, tạo nội dung trong Office" checked={st.agentMode} onChange={(_, d) => set({ agentMode: d.checked })} />
      {!st.agentMode && (
        <Text size={200} className={s.muted}>
          Đang tắt: khung chat chỉ trò chuyện (có nhãn “Chỉ trò chuyện” cạnh ô nhập).
        </Text>
      )}
      <Switch label="Hỏi tôi trước mỗi thao tác sửa tài liệu" checked={st.askBeforeWrite} onChange={(_, d) => set({ askBeforeWrite: d.checked })} />
      <Switch
        label="Luôn hỏi trước khi ghi đè dữ liệu có sẵn hoặc xóa (ô, đoạn, slide, sheet)"
        checked={st.confirmRisky}
        onChange={(_, d) => set({ confirmRisky: d.checked })}
      />
      <Text size={200} className={s.muted}>
        Word: mọi sửa đổi của AI luôn ghi bằng Theo dõi thay đổi (Track Changes) để anh duyệt trong thẻ Review, kể cả khi đang tắt Theo dõi thay
        đổi; xong việc add-in trả lại chế độ cũ.
      </Text>
      <Switch
        label="Suy nghĩ kỹ: cho AI suy nghĩ nhiều hơn trước khi trả lời (chậm hơn, tốn token hơn)"
        checked={!!st.deepThink}
        onChange={(_, d) => set({ deepThink: d.checked })}
      />
      <Text size={200} className={s.muted}>
        Có tác dụng với Anthropic, Gemini, OpenRouter và OpenAI, tùy model. Bật tắt nhanh ở dấu + cạnh ô nhập.
      </Text>
      <Switch
        label="Word: soát lập luận trước khi ghi sửa đổi vào điều khoản"
        checked={!!st.reviewEdits}
        onChange={(_, d) => set({ reviewEdits: d.checked })}
      />
      <Text size={200} className={s.muted}>
        Trước mỗi lần AI viết lại, thay chữ hoặc chèn điều khoản, một lượt AI riêng đối chiếu đề xuất với yêu cầu của anh và chữ hiện tại; thấy
        vấn đề thì chưa ghi mà báo lại để AI sửa đề xuất. Mỗi lần sửa tốn thêm một lượt gọi model.
      </Text>
    </div>
  );
};

const CrossAppCard: React.FC<Props> = ({ storage, onChanged }) => {
  const s = useStyles();
  const { st, set } = useSettingsState(storage, onChanged);
  return (
    <div className={s.card}>
      <Switch
        label="Cho Excel, Word, PowerPoint đọc chéo nội dung đang mở của nhau"
        checked={st.shareContext}
        onChange={(_, d) => set({ shareContext: d.checked })}
      />
    </div>
  );
};

const SelfTestCard: React.FC<Props> = ({ storage, host }) => {
  const s = useStyles();
  const [steps, setSteps] = React.useState<TestStep[] | null>(null);
  const [running, setRunning] = React.useState(false);
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
      <Text weight="semibold">Tự kiểm tra chức năng {host && host !== "Khác" ? `(${host})` : ""}</Text>
      <Text size={200} className={s.muted}>
        Chạy thật từng thao tác (đọc, ghi, định dạng, tạo bảng/biểu đồ/slide, nhận xét…) trên tài liệu đang mở rồi tự dọn sạch. Không cần API key. Trong
        Word, phần thử nằm trong một vùng riêng ở cuối văn bản và phần còn lại được kiểm là không đổi; Excel và PowerPoint nên chạy trên tài liệu trống.
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


const HOSTS = ["Excel", "Word", "PowerPoint"];

const InstructionsCard: React.FC<Props> = ({ storage, onChanged, host }) => {
  const s = useStyles();
  const [target, setTarget] = React.useState<string>(host && HOSTS.includes(host) ? host : "Word");
  const [text, setText] = React.useState((storage.getSettings().instructions || {})[target] || "");
  const [saved, setSaved] = React.useState("");
  React.useEffect(() => {
    setText((storage.getSettings().instructions || {})[target] || "");
    setSaved("");
  }, [target, storage]);
  const save = () => {
    const all = { ...(storage.getSettings().instructions || {}) };
    if (text.trim()) all[target] = text.trim();
    else delete all[target];
    storage.setSettings({ instructions: all });
    setSaved("Đã lưu. Áp dụng cho mọi yêu cầu tiếp theo.");
    onChanged();
  };
  return (
    <div className={s.card}>
      <Text weight="semibold">Chỉ dẫn riêng (luôn áp dụng)</Text>
      <Text size={200} className={s.muted}>
        Ví dụ: “Tôi làm pháp lý dự án bất động sản, bên tôi là chủ đầu tư; số tiền ghi đơn vị triệu đồng; văn phong hành chính.”
      </Text>
      <div className={s.row}>
        <Text size={200}>Cho</Text>
        <Select value={target} onChange={(_, d) => setTarget(d.value)} size="small">
          {HOSTS.map((h) => (
            <option key={h} value={h}>
              {h}
            </option>
          ))}
        </Select>
      </div>
      <Textarea value={text} onChange={(_, d) => setText(d.value)} rows={4} resize="vertical" placeholder={`Chỉ dẫn cho AI khi làm việc trong ${target}…`} />
      <div className={s.row}>
        <Button size="small" appearance="primary" onClick={save}>
          Lưu chỉ dẫn
        </Button>
        {saved && (
          <Text size={200} className={s.ok}>
            {saved}
          </Text>
        )}
      </div>
    </div>
  );
};

const SkillsCard: React.FC<Props> = ({ onChanged }) => {
  const s = useStyles();
  const store = React.useMemo(() => new SkillStore(getKV()), []);
  const [skills, setSkills] = React.useState<Skill[]>([]);
  const [commands, setCommands] = React.useState<Command[]>([]);
  const [open, setOpen] = React.useState(false);
  const [paste, setPaste] = React.useState("");
  const [cmdName, setCmdName] = React.useState("");
  const [cmdBody, setCmdBody] = React.useState("");
  const [report, setReport] = React.useState<{ text: string; err?: boolean } | null>(null);
  const fileRef = React.useRef<HTMLInputElement>(null);

  const reload = React.useCallback(async () => {
    try {
      setSkills(await store.all());
      setCommands(await store.commands());
    } catch (e) {
      setReport({ text: "Không đọc được bộ nhớ kỹ năng: " + networkErrorMessage(e), err: true });
    }
  }, [store]);
  React.useEffect(() => {
    reload();
  }, [reload]);

  const done = async (text: string, err = false) => {
    setReport({ text, err });
    await reload();
    onChanged();
  };

  const importFiles = async (files: FileList | null) => {
    if (!files || !files.length) return;
    const msgs: string[] = [];
    let nSkills = 0;
    let nCmds = 0;
    for (const f of Array.from(files)) {
      try {
        if (/\.(zip|skill|plugin)$/i.test(f.name)) {
          const r = await importZip(await f.arrayBuffer(), f.name);
          await store.addSkills(r.skills);
          await store.addCommands(r.commands);
          nSkills += r.skills.length;
          nCmds += r.commands.length;
          if (r.ignored.length) msgs.push(`${f.name}: bỏ qua ${r.ignored.length} mục (${r.ignored.slice(0, 3).join("; ")}${r.ignored.length > 3 ? "…" : ""})`);
        } else if (/\.md$/i.test(f.name)) {
          const text = await f.text();
          const base = f.name.replace(/\.md$/i, "");
          const fm = /^---\r?\n([\s\S]*?)\r?\n---/.exec(text)?.[1] || "";
          const isSkill = /^skill$/i.test(base) || (/^name\s*:/m.test(fm) && /^description\s*:/m.test(fm));
          if (isSkill) {
            await store.addSkills([parseSkillMd(text, /^skill$/i.test(base) ? "" : base, f.name)]);
            nSkills++;
          } else {
            await store.addCommands([parseCommandMd(text, base, f.name)]);
            nCmds++;
          }
        } else msgs.push(`${f.name}: chỉ nhận .md, .zip, .skill`);
      } catch (e: any) {
        msgs.push(`${f.name}: ${e?.message || e}`);
      }
    }
    if (fileRef.current) fileRef.current.value = "";
    await done(`Đã nhập ${nSkills} kỹ năng, ${nCmds} lệnh.${msgs.length ? " " + msgs.join(" · ") : ""}`, !nSkills && !nCmds);
  };

  const importPaste = async () => {
    try {
      const sk = parseSkillMd(paste, "", "dán tay");
      await store.addSkills([sk]);
      setPaste("");
      await done(`Đã thêm kỹ năng "${sk.name}".`);
    } catch (e: any) {
      setReport({ text: e?.message || String(e), err: true });
    }
  };

  const addCommand = async () => {
    const name = toSkillName(cmdName);
    if (!cmdBody.trim()) return;
    await store.addCommands([{ name, body: cmdBody.trim(), description: cmdBody.trim().slice(0, 80) }]);
    setCmdName("");
    setCmdBody("");
    await done(`Đã thêm lệnh /${name}.`);
  };

  const builtin = skills.filter((k) => k.source === "builtin");
  const user = skills.filter((k) => k.source !== "builtin");
  const on = skills.filter((k) => k.enabled).length;

  return (
    <div className={s.card}>
      <Text weight="semibold">Kỹ năng (Skills) và Lệnh của tôi</Text>
      <Text size={200} className={s.muted}>
        {builtin.length} kỹ năng tích hợp sẵn (chuyển thể từ skill chính thức xlsx, docx, pptx, dataviz và plugin Legal của Claude) + {user.length} kỹ năng
        nhập thêm · {on} đang bật · {commands.length} lệnh. AI tự chọn kỹ năng phù hợp; gõ “/” trong khung chat để gọi trực tiếp.
      </Text>
      <div className={s.row}>
        <Button size="small" onClick={() => setOpen((v) => !v)}>
          {open ? "Ẩn danh sách" : "Xem danh sách"}
        </Button>
        <Button size="small" onClick={() => fileRef.current?.click()}>
          Nhập tệp .md / .zip / .skill
        </Button>
      </div>
      <input
        ref={fileRef}
        type="file"
        multiple
        accept=".md,.zip,.skill,.plugin"
        style={{ display: "none" }}
        onChange={(e) => importFiles(e.target.files)}
      />
      {report && (
        <Text size={200} className={report.err ? s.err : s.ok}>
          {report.text}
        </Text>
      )}
      {open && (
        <>
          {skills.map((k) => (
            <div key={k.name} className={s.row} style={{ alignItems: "flex-start" }}>
              <Switch checked={!!k.enabled} onChange={async (_, d) => {
                await store.setEnabled(k.name, d.checked);
                await reload();
                onChanged();
              }} />
              <div style={{ flex: 1, minWidth: 0 }}>
                <Text size={200} weight="semibold">
                  /{k.name}
                </Text>{" "}
                <Text size={100} className={s.muted}>
                  {k.source === "builtin" ? "tích hợp" : k.origin || "nhập thêm"}
                  {k.hosts?.length ? ` · ${k.hosts.join(", ")}` : " · mọi ứng dụng"}
                </Text>
                <Text size={100} block className={s.muted}>
                  {k.description.slice(0, 160)}
                </Text>
              </div>
              {k.source !== "builtin" && (
                <Button size="small" appearance="subtle" onClick={async () => {
                  await store.remove(k.name);
                  await done(`Đã xóa kỹ năng "${k.name}".`);
                }}>
                  Xóa
                </Button>
              )}
            </div>
          ))}
          <Text weight="semibold" size={200} style={{ marginTop: 6 }}>
            Dán nội dung SKILL.md
          </Text>
          <Textarea value={paste} onChange={(_, d) => setPaste(d.value)} rows={4} resize="vertical" placeholder={"---\nname: ten-ky-nang\ndescription: Dùng khi…\n---\nHướng dẫn…"} />
          <Button size="small" onClick={importPaste} disabled={!paste.trim()}>
            Thêm kỹ năng
          </Button>
          <Text weight="semibold" size={200} style={{ marginTop: 6 }}>
            Lệnh của tôi ({commands.length})
          </Text>
          {commands.map((c) => (
            <div key={c.name} className={s.row}>
              <Text size={200} style={{ flex: 1, minWidth: 0 }}>
                /{c.name} <span className={s.muted}>{(c.description || c.body).slice(0, 80)}</span>
              </Text>
              <Button size="small" appearance="subtle" onClick={async () => {
                await store.removeCommand(c.name);
                await done(`Đã xóa lệnh /${c.name}.`);
              }}>
                Xóa
              </Button>
            </div>
          ))}
          <Input size="small" autoComplete="off" placeholder="Tên lệnh, vd: tom-tat" value={cmdName} onChange={(_, d) => setCmdName(d.value)} />
          <Textarea
            value={cmdBody}
            onChange={(_, d) => setCmdBody(d.value)}
            rows={3}
            resize="vertical"
            placeholder="Nội dung lệnh. $ARGUMENTS = phần chữ gõ sau tên lệnh."
          />
          <Button size="small" onClick={addCommand} disabled={!cmdName.trim() || !cmdBody.trim()}>
            Thêm lệnh
          </Button>
        </>
      )}
    </div>
  );
};

/** Trang Cài đặt (mở bằng nút bánh răng ở đầu khung chat). Thứ tự theo kế hoạch v4 mục 3.2. */
const Settings: React.FC<Props> = (props) => {
  const s = useStyles();
  const keys = props.storage.getKeys();
  // Nguồn có đăng nhập (OpenRouter) luôn đứng đầu, có thẻ riêng: đăng nhập, hạn mức, đăng xuất
  const signIns = PROVIDERS.filter((p) => p.oauth);
  const keyed = PROVIDERS.filter((p) => keys[p.id] && !p.oauth);
  const others = PROVIDERS.filter((p) => !keys[p.id] && !p.oauth);
  const [showOthers, setShowOthers] = React.useState(keyed.length === 0 && !signIns.some((p) => keys[p.id]));
  /** Khôi phục xong thì dựng lại các thẻ để đọc dữ liệu mới */
  const [gen, setGen] = React.useState(0);
  return (
    <div className={s.page} key={gen}>
      <div className={s.header}>
        {props.onBack && (
          <Button appearance="subtle" size="small" icon={<ArrowLeftRegular />} onClick={props.onBack} title="Quay lại khung chat">
            Quay lại
          </Button>
        )}
        <Text weight="semibold" size={300}>
          Cài đặt
        </Text>
        <div style={{ flex: 1 }} />
        <Text size={100} className={s.muted} title={`Phiên bản ${APP_VERSION_FULL}`}>
          v{APP_VERSION}
          {props.host && props.host !== "Khác" ? ` · ${props.host}` : ""}
        </Text>
      </div>
      <div className={s.scroll}>
        <div className={s.section}>Tài khoản và key</div>
        <Text size={200} className={s.muted} block style={{ marginBottom: 8 }}>
          Key chỉ lưu trên máy này, trong add-in; chỉ gửi tới đúng nhà cung cấp khi gọi. Bấm “Lưu” để kiểm tra key và tải danh sách model. Model đã chọn
          được nhớ riêng cho từng nhà cung cấp (đổi ở nhãn model cạnh ô nhập).
        </Text>
        {signIns.map((p) => (
          <SignInCard key={p.id} p={p} storage={props.storage} onChanged={props.onChanged} />
        ))}
        {keyed.map((p) => (
          <ProviderCard key={p.id} p={p} {...props} />
        ))}
        {keyed.length > 0 && (
          <Button size="small" appearance="subtle" onClick={() => setShowOthers((v) => !v)} style={{ marginBottom: 8 }}>
            {showOthers ? "Ẩn nhà cung cấp chưa có key" : `Thêm key nhà cung cấp khác (${others.length})`}
          </Button>
        )}
        {showOthers && others.map((p) => <ProviderCard key={p.id} p={p} {...props} />)}

        <div className={s.section}>Quyền và an toàn</div>
        <PermissionsCard {...props} />
        <div className={s.section}>Chỉ dẫn và kỹ năng</div>
        <InstructionsCard {...props} />
        <SkillsCard {...props} />
        <div className={s.section}>Đọc chéo ứng dụng</div>
        <CrossAppCard {...props} />
        <div className={s.section}>Sao lưu</div>
        <BackupCard
          host={props.host}
          onRestored={() => {
            props.onChanged();
            setGen((g) => g + 1);
          }}
        />
        <div className={s.section}>Kiểm tra</div>
        <SelfTestCard {...props} />
        <ProviderCheckCard storage={props.storage} />
      </div>
    </div>
  );
};


export default Settings;
