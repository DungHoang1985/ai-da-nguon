/** Thẻ "Kiểm tra nhà cung cấp" (trang Cài đặt): chạy thật với từng key đã lưu; bảng đổi nguồn mọi cặp, cả hai chiều. */
import * as React from "react";
import { Button, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { allProviders, quirksOf } from "../../core/providers";
import { Storage } from "../../core/storage";
import { fetchModels, isStale } from "../../core/models";
import { describeError } from "../../core/errors";
import { pickModel } from "../../core/catalog";
import { CheckResult, CheckTarget, MatrixCell, matrixReport, runSwitchMatrix } from "../../core/providerCheck";
import { copyText } from "../fileio";

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
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
  grid: {
    borderCollapse: "collapse",
    fontSize: "11px",
    "& td, & th": { border: `1px solid ${tokens.colorNeutralStroke2}`, padding: "1px 3px", textAlign: "center" },
    "& th": { background: tokens.colorNeutralBackground3, fontWeight: 600 },
  },
});

const SHORT: Record<string, string> = {
  openrouter: "OR",
  gemini: "Gem",
  groq: "Grq",
  anthropic: "Ant",
  openai: "OAI",
  mistral: "Mis",
  cerebras: "Cer",
  deepseek: "DS",
  sambanova: "Sam",
  nvidia: "NV",
  cohere: "Coh",
};

/** Chọn model cho mỗi nguồn: model anh dùng lần trước nếu gọi được công cụ, không thì model gọi được công cụ phù hợp nhất. */
async function prepareTargets(storage: Storage): Promise<{ targets: CheckTarget[]; pre: Record<string, CheckResult> }> {
  const keys = storage.getKeys();
  const targets: CheckTarget[] = [];
  const pre: Record<string, CheckResult> = {};
  for (const p of allProviders().filter((x) => keys[x.id])) {
    let cached = storage.getModelCache()[p.id];
    let loadErr = "";
    if (!cached || isStale(cached.at)) {
      try {
        const list = await fetchModels(p, storage.getKey(p.id));
        storage.setModels(p.id, list);
        cached = { at: Date.now(), models: list };
      } catch (e) {
        loadErr = describeError(e, { provider: p }).text;
      }
    }
    const models = cached?.models || [];
    if (!models.length) {
      pre[p.id] = { ok: false, ms: 0, message: `Không tải được danh sách model: ${loadErr || "không có model"}` };
      continue;
    }
    const remembered = storage.getModelFor(p.id);
    const rememberedOk = models.find((m) => m.id === remembered && m.tools !== false);
    const model = pickModel(models, {
      providerId: p.id,
      remembered: rememberedOk ? remembered : undefined,
      agentOn: true,
      freeOnly: !!quirksOf(p).mixedFreeCatalog && !!storage.getSelection().freeOnly,
    });
    const info = models.find((m) => m.id === model);
    targets.push({ provider: p, key: storage.getKey(p.id), model, vision: info?.vision ?? p.vision });
  }
  return { targets, pre };
}

const ProviderCheckCard: React.FC<{ storage: Storage }> = ({ storage }) => {
  const s = useStyles();
  const [running, setRunning] = React.useState<"" | "single" | "matrix">("");
  const [targets, setTargets] = React.useState<CheckTarget[]>([]);
  const [singles, setSingles] = React.useState<Record<string, CheckResult>>({});
  const [cells, setCells] = React.useState<MatrixCell[]>([]);
  const [note, setNote] = React.useState<{ text: string; err?: boolean }>();
  const abortRef = React.useRef<AbortController | null>(null);

  const run = async (mode: "single" | "matrix") => {
    setRunning(mode);
    setNote({ text: "Đang chuẩn bị (tải danh sách model)…" });
    setSingles({});
    setCells([]);
    const ac = new AbortController();
    abortRef.current = ac;
    try {
      const { targets: ts, pre } = await prepareTargets(storage);
      setTargets(ts);
      setSingles({ ...pre });
      const n = ts.length;
      setNote({ text: mode === "matrix" ? `Đang chạy: ${n} nguồn, ${n * (n - 1)} cặp đổi nguồn…` : `Đang chạy ${n} nguồn…` });
      const res = await runSwitchMatrix({
        targets: ts,
        singlesOnly: mode === "single",
        signal: ac.signal,
        onSingle: (pid, r) => setSingles((prev) => ({ ...prev, [pid]: r })),
        onCell: (c) => setCells((prev) => [...prev, c]),
      });
      const okS = Object.values(res.singles).filter((r) => r.ok).length;
      const okC = res.cells.filter((c) => c.ok).length;
      setNote({
        text: `Xong. Từng nguồn: đạt ${okS}/${n}${mode === "matrix" ? ` · Đổi nguồn: đạt ${okC}/${res.cells.length}` : ""}.`,
        err: okS < n || okC < res.cells.length,
      });
    } catch (e) {
      setNote({ text: describeError(e).text, err: true });
    } finally {
      setRunning("");
      abortRef.current = null;
    }
  };

  const copyReport = async () => {
    const ok = await copyText(matrixReport(targets, singles, cells));
    setNote(ok ? { text: "Đã chép kết quả." } : { text: "Không chép được trong môi trường này.", err: true });
  };

  const ids = targets.map((t) => t.provider.id);
  const cellOf = (a: string, b: string) => cells.find((c) => c.from === a && c.to === b);
  const failed = cells.filter((c) => !c.ok);
  const keyed = allProviders().filter((p) => storage.getKeys()[p.id]);

  return (
    <div className={s.card}>
      <Text weight="semibold">Kiểm tra nhà cung cấp</Text>
      <Text size={200} className={s.muted}>
        Chạy thật với key đã lưu: model gọi một công cụ thử → nhận kết quả → trả lời có chứa kết quả. “Đổi nguồn” gửi tiếp hội thoại của
        nguồn này sang từng nguồn khác (cả hai chiều) và kiểm AI còn nhớ lượt trước. Tốn rất ít token; {keyed.length} nguồn có key.
      </Text>
      <div className={s.row}>
        <Button size="small" onClick={() => run("single")} disabled={!!running || !keyed.length}>
          Kiểm tra từng nguồn
        </Button>
        <Button size="small" onClick={() => run("matrix")} disabled={!!running || keyed.length < 2}>
          Kiểm tra đổi nguồn (mọi cặp)
        </Button>
        {running && (
          <>
            <Spinner size="tiny" />
            <Button size="small" onClick={() => abortRef.current?.abort()}>
              Dừng
            </Button>
          </>
        )}
        {!running && (Object.keys(singles).length > 0 || cells.length > 0) && (
          <Button size="small" appearance="subtle" onClick={copyReport}>
            Chép kết quả
          </Button>
        )}
      </div>
      {note && (
        <Text size={200} className={note.err ? s.err : s.muted}>
          {note.text}
        </Text>
      )}
      {keyed
        .filter((p) => singles[p.id] || running)
        .map((p) => {
          const r = singles[p.id];
          const t = targets.find((x) => x.provider.id === p.id);
          return (
            <Text key={p.id} size={200} className={r ? (r.ok ? s.ok : s.err) : s.muted}>
              {r ? (r.ok ? "✅" : "❌") : "⏳"} {p.name}
              {t?.model ? ` · ${t.model}` : ""}
              {r ? ` — ${r.message}` : ""}
            </Text>
          );
        })}
      {cells.length > 0 && ids.length > 1 && (
        <>
          <Text size={200} weight="semibold">
            Đổi nguồn (hàng: nguồn trước → cột: nguồn sau): đạt {cells.filter((c) => c.ok).length}/{ids.length * (ids.length - 1)}
          </Text>
          <div style={{ overflowX: "auto" }}>
            <table className={s.grid}>
              <tbody>
                <tr>
                  <th>từ \ sang</th>
                  {ids.map((b) => (
                    <th key={b}>{SHORT[b] || b}</th>
                  ))}
                </tr>
                {ids.map((a) => (
                  <tr key={a}>
                    <th>{SHORT[a] || a}</th>
                    {ids.map((b) => {
                      const c = a === b ? undefined : cellOf(a, b);
                      return (
                        <td key={b} title={c ? c.message : ""} className={c ? (c.ok ? s.ok : s.err) : s.muted}>
                          {a === b ? "·" : c ? (c.ok ? "✓" : "✗") : "…"}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {failed.map((c) => (
            <Text key={c.from + c.to} size={100} className={s.err}>
              {c.from} → {c.to}: {c.message}
            </Text>
          ))}
        </>
      )}
    </div>
  );
};

export default ProviderCheckCard;
