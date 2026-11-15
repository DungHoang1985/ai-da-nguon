/** Nhãn model cạnh ô nhập; bấm vào mới bung bảng chọn: nhà cung cấp, ô tìm, lọc Miễn phí, Chỉ model thao tác được, tải lại. */
import * as React from "react";
import {
  Button,
  Input,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Select,
  Spinner,
  Switch,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ArrowClockwiseRegular, ChevronDownRegular } from "@fluentui/react-icons";
import type { ModelInfo } from "../../core/models";
import { useCatalog } from "./useCatalog";

type Catalog = ReturnType<typeof useCatalog>;

const useStyles = makeStyles({
  trigger: { minWidth: 0, maxWidth: "100%", justifyContent: "flex-start", fontWeight: 400 },
  label: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 },
  prov: { color: tokens.colorNeutralForeground3 },
  surface: { width: "min(320px, calc(100vw - 24px))", padding: "10px", display: "flex", flexDirection: "column", gap: "6px" },
  row: { display: "flex", gap: "6px", alignItems: "center", flexWrap: "wrap", minWidth: 0 },
  list: {
    maxHeight: "min(260px, 45vh)",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: "6px",
    display: "flex",
    flexDirection: "column",
  },
  item: {
    padding: "5px 8px",
    cursor: "pointer",
    fontSize: "12px",
    lineHeight: "16px",
    ":hover": { background: tokens.colorNeutralBackground1Hover },
  },
  sel: { background: tokens.colorBrandBackground2, fontWeight: 600 },
  muted: { color: tokens.colorNeutralForeground3 },
  err: { color: tokens.colorPaletteRedForeground1 },
});

/** Tên gọn để hiện trên nhãn: bỏ phần "(mã model)" ở cuối tên. */
export function shortModelName(m: ModelInfo | undefined, fallback = ""): string {
  if (!m) return fallback;
  return m.name.replace(/\s*\([^()]*\)\s*$/, "").trim() || m.id;
}

const ModelPicker: React.FC<{
  cat: Catalog;
  agentOn: boolean;
  toolsOnly: boolean;
  onToolsOnly: (v: boolean) => void;
}> = ({ cat, agentOn, toolsOnly, onToolsOnly }) => {
  const s = useStyles();
  const [open, setOpen] = React.useState(false);
  const status = cat.catalog.status;
  const label =
    status === "loading" ? "Đang tải model…" : status === "error" ? "Chưa dùng được" : shortModelName(cat.model, "Chọn model");
  const provShort = (cat.provider?.name || "").replace(/\s*\(.*\)$/, "");

  return (
    <Popover open={open} onOpenChange={(_, d) => setOpen(d.open)} positioning="above-start" trapFocus>
      <PopoverTrigger disableButtonEnhancement>
        <Button
          size="small"
          appearance="subtle"
          icon={<ChevronDownRegular />}
          iconPosition="after"
          className={s.trigger}
          title={`${cat.provider?.name || ""} · ${cat.modelId || "chưa chọn model"} — bấm để đổi`}
          aria-label={`Model đang dùng: ${cat.provider?.name || ""} ${label}. Bấm để đổi`}
        >
          <span className={s.label}>
            <span className={s.prov}>{provShort} · </span>
            <span className={status === "error" ? s.err : undefined}>{label}</span>
          </span>
        </Button>
      </PopoverTrigger>
      <PopoverSurface className={s.surface} aria-label="Chọn nhà cung cấp và model">
        <div className={s.row}>
          <Select
            size="small"
            value={cat.providerId}
            onChange={(_, d) => cat.selectProvider(d.value)}
            style={{ flex: 1, minWidth: 0 }}
            aria-label="Nhà cung cấp"
          >
            {cat.available.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </Select>
          <Button size="small" icon={<ArrowClockwiseRegular />} onClick={() => cat.reload()} title="Tải lại danh sách model" aria-label="Tải lại danh sách model" />
        </div>
        {cat.provider?.blockedReason && (
          <Text size={200} className={s.err}>
            ⚠️ {cat.provider.name}: {cat.provider.blockedReason}
          </Text>
        )}
        <Input
          size="small"
          placeholder="Tìm model…"
          autoComplete="off"
          name="aidn-model-search"
          value={cat.search}
          onChange={(_, d) => cat.setSearch(d.value)}
        />
        {cat.providerId === "openrouter" && (
          <div className={s.row}>
            <Switch label="Miễn phí" checked={cat.freeOnly} onChange={(_, d) => cat.setFreeOnly(d.checked)} />
            {agentOn && <Switch label="Chỉ model thao tác được" checked={toolsOnly} onChange={(_, d) => onToolsOnly(d.checked)} />}
          </div>
        )}
        <div className={s.list} role="listbox" aria-label="Model">
          {status === "loading" && (
            <div className={s.item}>
              <Spinner size="extra-tiny" label="Đang tải danh sách model…" labelPosition="after" />
            </div>
          )}
          {status === "error" && (
            <Text size={200} className={`${s.item} ${s.err}`}>
              {cat.blockReason}
            </Text>
          )}
          {cat.shown.map((m) => (
            <div
              key={m.id}
              role="option"
              aria-selected={m.id === cat.modelId}
              tabIndex={0}
              title={m.id}
              className={`${s.item} ${m.id === cat.modelId ? s.sel : ""}`}
              onClick={() => {
                cat.selectModel(m.id);
                setOpen(false);
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") {
                  e.preventDefault();
                  cat.selectModel(m.id);
                  setOpen(false);
                }
              }}
            >
              {m.free ? "🆓 " : ""}
              {m.tools ? "🛠 " : ""}
              {m.name}
            </div>
          ))}
          {status === "ready" && !cat.shown.length && (
            <Text size={200} className={`${s.item} ${s.muted}`}>
              Không có model khớp bộ lọc.
            </Text>
          )}
        </div>
        <Text size={100} className={s.muted}>
          {cat.note}
          {status === "ready" && cat.shown.length !== cat.models.length ? ` · đang hiện ${cat.shown.length}/${cat.models.length}` : ""}
          {agentOn && status === "ready" ? " · 🛠 = gọi được công cụ (thao tác trên tài liệu)" : ""}
        </Text>
      </PopoverSurface>
    </Popover>
  );
};

export default ModelPicker;
