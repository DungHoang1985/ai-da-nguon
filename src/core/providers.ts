/** Danh sách nhà cung cấp. Thêm nguồn OpenAI-compatible mới: chỉ cần thêm 1 dòng. */

export type ProviderKind = "openai" | "gemini" | "anthropic";

export interface Provider {
  id: string;
  name: string;
  kind: ProviderKind;
  /** URL gốc API (không có dấu / cuối) */
  baseUrl: string;
  /** Trang lấy API key */
  keyUrl: string;
  /** URL liệt kê model nếu khác `${baseUrl}/models` */
  modelsUrl?: string;
  /** Ghi chú hiển thị trong Cài đặt */
  note?: string;
  /** Nhà cung cấp không cho gọi trực tiếp từ add-in (đo được): hiện lý do ngay khi chọn */
  blockedReason?: string;
  /** Mọi model của nguồn này nhận ảnh (true) / không nhận (false); không ghi = tùy model hoặc chưa rõ */
  vision?: boolean;
  /** Nguồn cho đăng nhập để nhận key (OAuth PKCE) thay cho dán key. Thêm nguồn cùng kiểu: chỉ thêm cấu hình này. */
  oauth?: OAuthConfig;
  /** Trang quản lý key của người dùng (để thu hồi key sau khi đăng xuất) */
  manageKeysUrl?: string;
  /** Trang xem số dư, nạp tiền */
  billingUrl?: string;
  /** Nguồn kind "anthropic" nhưng không phải api.anthropic.com: các chỗ khác với bản gốc (Đợt 6) */
  dialect?: AnthropicDialect;
  /** Dùng chung key với nguồn này (một tài khoản, hai kênh): không có ô nhập key riêng */
  keyFrom?: string;
  /** Nguồn do người dùng tự khai (địa chỉ, key, tên model) */
  custom?: boolean;
  /** Chế độ dùng (Đợt 9): "free" = chỉ hiện ở chế độ Miễn phí, "both" = có bậc miễn phí và bậc trả phí (cùng key),
   *  không ghi = chỉ Trả phí. Lọc model theo chế độ ở src/core/modes.ts. */
  tier?: "free" | "paid" | "both";
}

/** "Phương ngữ" của một nguồn nhận kiểu yêu cầu Anthropic (POST …/messages). Thiếu trường nào = giống api.anthropic.com.
 *  Căn cứ: tài liệu chính thức và phép đo CORS ngày 17/11/2026, bảng ở tasks/dot-6.md. */
export interface AnthropicDialect {
  /** Cách gửi key: "x-api-key" (mặc định) hoặc "bearer" (Authorization: Bearer) */
  auth?: "x-api-key" | "bearer";
  /** Chỉ gửi tiêu đề key và content-type: không anthropic-version, anthropic-beta, anthropic-dangerous-direct-browser-access
   *  (nhiều nguồn không cho các tiêu đề này ở bước dò trước của trình duyệt) */
  bare?: boolean;
  /** false = nguồn bỏ qua cache_control: không gắn mốc bộ nhớ đệm */
  cache?: boolean;
  /** Tham số suy nghĩ: "anthropic" (mặc định, theo đời model Claude) · "enabled" (chỉ bật / tắt) · "none" (không gửi gì) */
  thinking?: "anthropic" | "enabled" | "none";
  /** Nguồn nhận nguyên tệp PDF (khối document base64) */
  pdf?: boolean;
  /** Nguồn có công cụ tìm web kiểu Anthropic */
  webSearch?: boolean;
  /** Địa chỉ liệt kê model kiểu OpenAI (GET, Authorization: Bearer); lỗi hoặc không có thì dùng `models` */
  modelsUrl?: string;
  /** Chỉ giữ model có mã khớp mẫu này (khi lấy từ modelsUrl) */
  modelFilter?: string;
  /** Danh sách model cố định theo tài liệu của nguồn */
  models?: { id: string; name?: string; context?: number; vision?: boolean }[];
}

/** Cấu hình đăng nhập kiểu OAuth PKCE trả về API key (đối chiếu openrouter.ai/docs/guides/overview/auth/oauth, 15/11/2026). */
export interface OAuthConfig {
  /** Trang đăng nhập và cho phép của nhà cung cấp */
  authUrl: string;
  /** Địa chỉ đổi mã lấy key (POST JSON: code, code_verifier, code_challenge_method) */
  exchangeUrl: string;
  /** Tên gợi ý cho key mới, hiện ở trang quản lý key của nhà cung cấp */
  keyLabel?: string;
  /** Tên miền của authUrl: trang hộp thoại chỉ chuyển tới đúng tên miền này */
  host: string;
}

/** Đo ngày 14 và 15/11/2026: máy chủ không trả tiêu đề Access-Control-Allow-Origin nên trình duyệt trong Office chặn. */
const NO_CORS = "nhà cung cấp này chặn gọi trực tiếp từ add-in (đo 15/11/2026: máy chủ không trả tiêu đề CORS). Dùng model này qua OpenRouter nếu cần.";

export const PROVIDERS: Provider[] = [
  {
    id: "openrouter",
    name: "OpenRouter",
    tier: "both",
    kind: "openai",
    baseUrl: "https://openrouter.ai/api/v1",
    keyUrl: "https://openrouter.ai/keys",
    note: "Nhiều model miễn phí (đuôi :free)",
    oauth: {
      authUrl: "https://openrouter.ai/auth",
      exchangeUrl: "https://openrouter.ai/api/v1/auth/keys",
      keyLabel: "AI Da Nguon (Office)",
      host: "openrouter.ai",
    },
    manageKeysUrl: "https://openrouter.ai/settings/keys",
    billingUrl: "https://openrouter.ai/settings/credits",
  },
  {
    id: "gemini",
    name: "Google Gemini",
    tier: "both",
    kind: "gemini",
    baseUrl: "https://generativelanguage.googleapis.com/v1beta",
    keyUrl: "https://aistudio.google.com/apikey",
    vision: true,
  },
  {
    id: "groq",
    name: "Groq",
    tier: "both",
    kind: "openai",
    baseUrl: "https://api.groq.com/openai/v1",
    keyUrl: "https://console.groq.com/keys",
  },
  {
    id: "anthropic",
    name: "Anthropic (Claude)",
    kind: "anthropic",
    baseUrl: "https://api.anthropic.com/v1",
    keyUrl: "https://console.anthropic.com/settings/keys",
    vision: true,
  },
  {
    // Kênh Messages của OpenRouter (đối chiếu openrouter.ai/docs/api/api-reference/anthropic-messages, 17/11/2026):
    // model Claude đi đúng kiểu Anthropic — có bộ nhớ đệm chủ động, PDF nguyên tệp, khối suy nghĩ. Dùng chung key với OpenRouter.
    id: "openrouter-claude",
    name: "OpenRouter · Claude (kiểu Anthropic)",
    kind: "anthropic",
    baseUrl: "https://openrouter.ai/api/v1",
    keyUrl: "https://openrouter.ai/keys",
    keyFrom: "openrouter",
    billingUrl: "https://openrouter.ai/settings/credits",
    note: "Model Claude qua OpenRouter, gọi đúng kiểu Anthropic: có bộ nhớ đệm, PDF nguyên tệp. Dùng chung key với OpenRouter.",
    vision: true,
    dialect: {
      auth: "bearer",
      bare: true,
      cache: true,
      thinking: "anthropic",
      pdf: true,
      webSearch: false,
      modelsUrl: "https://openrouter.ai/api/v1/models",
      modelFilter: "^anthropic/",
    },
  },
  {
    // api-docs.deepseek.com/guides/anthropic_api (17/11/2026): bỏ qua cache_control, PDF base64 bị từ chối, ảnh chỉ deepseek-flash
    id: "deepseek-anthropic",
    name: "DeepSeek (kiểu Anthropic)",
    kind: "anthropic",
    baseUrl: "https://api.deepseek.com/anthropic/v1",
    keyUrl: "https://platform.deepseek.com/api_keys",
    keyFrom: "deepseek",
    note: "DeepSeek gọi theo kiểu Anthropic. Không nhận PDF quét; ảnh chỉ deepseek-flash. Dùng chung key với DeepSeek.",
    dialect: {
      auth: "x-api-key",
      cache: false,
      thinking: "enabled",
      pdf: false,
      webSearch: false,
      models: [
        { id: "deepseek-flash", name: "DeepSeek Flash (deepseek-flash)", vision: true },
        { id: "deepseek-v4-pro", name: "DeepSeek V4 Pro (deepseek-v4-pro)", vision: false },
      ],
    },
  },
  {
    // platform.minimax.io/docs/api-reference/text-anthropic-api (17/11/2026). Đo: bước dò trước chỉ cho Authorization.
    id: "minimax",
    name: "MiniMax",
    kind: "anthropic",
    baseUrl: "https://api.minimax.io/anthropic/v1",
    keyUrl: "https://platform.minimax.io/user-center/basic-information/interface-key",
    note: "Gọi theo kiểu Anthropic. Nhận ảnh; chưa rõ có nhận PDF quét.",
    vision: true,
    dialect: {
      auth: "bearer",
      bare: true,
      cache: false,
      thinking: "none",
      pdf: false,
      webSearch: false,
      models: [
        { id: "MiniMax-M3", name: "MiniMax M3 (MiniMax-M3)", context: 1000000, vision: true },
        { id: "MiniMax-M3.1-Flash", name: "MiniMax M3.1 Flash (MiniMax-M3.1-Flash)", context: 1000000, vision: true },
      ],
    },
  },
  {
    // platform.kimi.ai/docs/guide/agent-support (17/11/2026). Đo: bước dò trước chỉ cho Authorization và Content-Type.
    id: "kimi",
    name: "Kimi (Moonshot)",
    kind: "anthropic",
    baseUrl: "https://api.moonshot.ai/anthropic/v1",
    keyUrl: "https://platform.kimi.ai/console/api-keys",
    note: "Gọi theo kiểu Anthropic.",
    dialect: {
      auth: "bearer",
      bare: true,
      cache: false,
      thinking: "none",
      pdf: false,
      webSearch: false,
      modelsUrl: "https://api.moonshot.ai/v1/models",
      models: [
        { id: "kimi-k3.1-pro", name: "Kimi K3.1 Pro (kimi-k3.1-pro)" },
        { id: "kimi-k3.1-pro-highspeed", name: "Kimi K3.1 Pro Highspeed (kimi-k3.1-pro-highspeed)" },
      ],
    },
  },
  {
    id: "openai",
    name: "OpenAI",
    kind: "openai",
    baseUrl: "https://api.openai.com/v1",
    keyUrl: "https://platform.openai.com/api-keys",
  },
  {
    id: "mistral",
    name: "Mistral",
    kind: "openai",
    baseUrl: "https://api.mistral.ai/v1",
    keyUrl: "https://console.mistral.ai/api-keys",
  },
  {
    id: "cerebras",
    name: "Cerebras",
    kind: "openai",
    baseUrl: "https://api.cerebras.ai/v1",
    keyUrl: "https://cloud.cerebras.ai/",
  },
  {
    id: "deepseek",
    name: "DeepSeek",
    kind: "openai",
    baseUrl: "https://api.deepseek.com/v1",
    keyUrl: "https://platform.deepseek.com/api_keys",
    vision: false,
  },
  {
    id: "sambanova",
    name: "SambaNova",
    kind: "openai",
    baseUrl: "https://api.sambanova.ai/v1",
    keyUrl: "https://cloud.sambanova.ai/apis",
    blockedReason: NO_CORS,
  },
  {
    id: "nvidia",
    name: "NVIDIA NIM",
    kind: "openai",
    baseUrl: "https://integrate.api.nvidia.com/v1",
    keyUrl: "https://build.nvidia.com/",
    blockedReason: NO_CORS,
  },
  {
    id: "cohere",
    name: "Cohere",
    tier: "both",
    kind: "openai",
    baseUrl: "https://api.cohere.ai/compatibility/v1",
    modelsUrl: "https://api.cohere.com/v1/models?endpoint=chat&page_size=1000",
    keyUrl: "https://dashboard.cohere.com/api-keys",
  },
];

/** Nguồn do người dùng tự khai trong Cài đặt (lưu cùng kho của add-in) */
export interface CustomSource {
  id: string;
  name: string;
  /** Địa chỉ gốc kiểu ANTHROPIC_BASE_URL, ví dụ https://api.minimax.io/anthropic (add-in tự thêm /v1/messages) */
  baseUrl: string;
  /** Tên model, mỗi dòng một model */
  models: string[];
  /** Gửi key kiểu Anthropic gốc (x-api-key kèm anthropic-version) thay cho Authorization: Bearer */
  anthropicHeaders?: boolean;
  /** Người dùng khai đây là nguồn miễn phí (Đợt 9): chỉ hiện ở chế độ Miễn phí và vào vòng xoay. Không ghi = Trả phí. */
  free?: boolean;
}

export const CUSTOM_PREFIX = "tuy-chinh-";

/** Chuẩn hóa địa chỉ gốc người dùng nhập: chỉ nhận https (hoặc máy này), bỏ / cuối và phần /messages thừa, bảo đảm kết thúc bằng /v1.
 *  Sai thì trả error (câu tiếng Việt). */
export function normalizeBaseUrl(input: string): { url?: string; host?: string; error?: string } {
  let u: URL;
  try {
    u = new URL(String(input || "").trim());
  } catch {
    return { error: "Địa chỉ không hợp lệ. Ví dụ: https://api.minimax.io/anthropic" };
  }
  const local = u.hostname === "localhost" || u.hostname === "127.0.0.1" || u.hostname === "[::1]";
  if (u.protocol !== "https:" && !(u.protocol === "http:" && local)) return { error: "Chỉ nhận địa chỉ https (key được gửi tới địa chỉ này)." };
  if (u.username || u.password) return { error: "Địa chỉ không được chứa tên đăng nhập hay mật khẩu." };
  if (u.search || u.hash) return { error: "Địa chỉ không được chứa tham số (?…) hay dấu #." };
  let path = u.pathname.replace(/\/+$/, "").replace(/\/messages$/, "");
  if (!/\/v\d+$/.test(path)) path += "/v1";
  return { url: `${u.protocol}//${u.host}${path}`, host: u.host };
}

export function customProvider(c: CustomSource): Provider | undefined {
  const n = normalizeBaseUrl(c.baseUrl);
  if (!n.url || !String(c.id || "").startsWith(CUSTOM_PREFIX)) return undefined;
  return {
    id: c.id,
    name: c.name || n.host || c.id,
    kind: "anthropic",
    baseUrl: n.url,
    keyUrl: "",
    custom: true,
    ...(c.free ? { tier: "free" as const } : {}),
    note: `Nguồn tự khai, kiểu Anthropic · key được gửi tới ${n.host}`,
    dialect: {
      auth: c.anthropicHeaders ? "x-api-key" : "bearer",
      bare: !c.anthropicHeaders,
      cache: true,
      thinking: "none",
      pdf: false,
      webSearch: false,
      models: (c.models || []).map((id) => ({ id })),
    },
  };
}

let customProviders: Provider[] = [];
/** Nạp các nguồn tự khai (gọi khi mở add-in và mỗi lần người dùng sửa danh sách) */
export function setCustomSources(list: CustomSource[]): void {
  customProviders = (Array.isArray(list) ? list : []).map(customProvider).filter((p): p is Provider => !!p);
}

/** Mọi nguồn: có sẵn + tự khai */
export function allProviders(): Provider[] {
  return [...PROVIDERS, ...customProviders];
}

export function getProvider(id: string): Provider | undefined {
  return allProviders().find((p) => p.id === id);
}
