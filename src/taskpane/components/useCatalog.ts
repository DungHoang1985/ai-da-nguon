/** Nguồn + danh mục model + model đang chọn của khung chat. Logic nằm ở src/core/catalog.ts (có kiểm thử). */
import * as React from "react";
import { PROVIDERS, getProvider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { fetchModels, isStale } from "../../core/models";
import { describeError, keyProblem } from "../../core/errors";
import { CatalogState, applyFailed, applyLoaded, filterModels, pickModel, sendBlockReason, startCatalog } from "../../core/catalog";

export function useCatalog(storage: Storage, o: { agentOn: boolean; toolsOnly: boolean; version: number }) {
  const keys = storage.getKeys();
  const available = PROVIDERS.filter((p) => keys[p.id]);
  const firstPid = (() => {
    const sel = storage.getSelection();
    return sel.providerId && keys[sel.providerId] ? sel.providerId : available[0]?.id || "";
  })();

  const tokenRef = React.useRef(0);
  const [catalog, setCatalog] = React.useState<CatalogState>(() => startCatalog(firstPid, storage.getModelCache()[firstPid], 0));
  const [modelId, setModelIdState] = React.useState<string>(() => storage.getModelFor(firstPid));
  const [freeOnly, setFreeOnlyState] = React.useState<boolean>(!!storage.getSelection().freeOnly);
  const [search, setSearch] = React.useState("");

  const providerId = catalog.providerId;
  const provider = getProvider(providerId);

  /** Đổi nguồn là thay ngay danh sách (bộ nhớ đệm của nguồn mới hoặc "đang tải"); kết quả tải muộn của lần cũ bị bỏ. */
  const load = React.useCallback(
    async (pid: string, force: boolean) => {
      const cached = storage.getModelCache()[pid];
      const token = ++tokenRef.current;
      setCatalog(startCatalog(pid, cached, token));
      const p = getProvider(pid);
      if (!p || (!force && cached && !isStale(cached.at))) return;
      const key = storage.getKey(pid);
      const problem = keyProblem(key);
      if (problem) {
        setCatalog((s) => applyFailed(s, pid, token, problem));
        return;
      }
      try {
        const list = await fetchModels(p, key);
        storage.setModels(pid, list);
        setCatalog((s) => applyLoaded(s, pid, token, list, Date.now()));
      } catch (e) {
        setCatalog((s) => applyFailed(s, pid, token, describeError(e, { provider: p }).text));
      }
    },
    [storage]
  );

  // Mở khung / đổi cài đặt (thêm, xóa key): nạp lại nguồn đang chọn; nguồn mất key thì sang nguồn khác
  React.useEffect(() => {
    const ks = storage.getKeys();
    const pid = providerId && ks[providerId] ? providerId : PROVIDERS.find((p) => ks[p.id])?.id || "";
    if (pid !== providerId) setModelIdState(storage.getModelFor(pid));
    if (pid) load(pid, false);
  }, [o.version]); // eslint-disable-line react-hooks/exhaustive-deps

  const selectProvider = (pid: string) => {
    if (pid === providerId) return;
    storage.setSelection({ providerId: pid });
    setModelIdState(storage.getModelFor(pid));
    setSearch("");
    load(pid, false);
  };

  const selectModel = (id: string) => {
    setModelIdState(id);
    storage.setModelFor(providerId, id);
    storage.setSelection({ providerId });
  };

  const setFreeOnly = (v: boolean) => {
    setFreeOnlyState(v);
    storage.setSelection({ freeOnly: v });
  };

  // Danh sách về (hoặc bộ lọc đổi) mà model đang chọn không thuộc nguồn này / bị lọc: chọn model đã nhớ hoặc model gọi được công cụ
  const eligible = filterModels(catalog.models, { providerId, freeOnly, toolsOnly: o.toolsOnly, agentOn: o.agentOn, search: "" });
  React.useEffect(() => {
    if (catalog.status !== "ready" || !eligible.length) return;
    if (eligible.some((m) => m.id === modelId)) return;
    const id = pickModel(eligible, { providerId, remembered: storage.getModelFor(providerId), agentOn: o.agentOn, freeOnly });
    if (id) selectModel(id);
  }, [catalog.status, catalog.providerId, eligible.map((m) => m.id).join("|")]); // eslint-disable-line react-hooks/exhaustive-deps

  const shown = filterModels(catalog.models, { providerId, freeOnly, toolsOnly: o.toolsOnly, agentOn: o.agentOn, search });
  const model = catalog.models.find((m) => m.id === modelId);
  const blockReason = sendBlockReason({ provider, catalog, modelId });
  const note =
    catalog.status === "loading"
      ? "Đang tải danh sách model…"
      : catalog.status === "error"
        ? ""
        : `Cập nhật lúc ${catalog.at ? new Date(catalog.at).toLocaleString("vi-VN") : "?"} · ${catalog.models.length} model` +
          (catalog.error ? ` · Không tải lại được: ${catalog.error}` : "");

  return {
    available,
    provider,
    providerId,
    catalog,
    models: catalog.models,
    shown,
    model,
    modelId,
    selectProvider,
    selectModel,
    freeOnly,
    setFreeOnly,
    search,
    setSearch,
    reload: () => load(providerId, true),
    blockReason,
    note,
  };
}
