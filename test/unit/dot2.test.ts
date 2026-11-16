/** Đợt 2 — Đăng nhập và chi phí: khối đăng nhập (PKCE), hạn mức OpenRouter, phân loại lỗi key và tiền. */
import * as assert from "assert";
import {
  CHALLENGE_METHOD,
  OAuthError,
  base64Url,
  buildAuthUrl,
  createLoginSession,
  dialogMessage,
  dialogStartUrl,
  exchangeCode,
  exchangeErrorMessage,
  exchangeRequest,
  extractCode,
  newState,
  newVerifier,
  parseDialogQuery,
  readDialogMessage,
  redact,
  s256Challenge,
  validChallenge,
} from "../../src/core/oauth";
import { PROVIDERS, getProvider } from "../../src/core/providers";
import { chatCreditWarning, describeKeyInfo, fetchKeyInfo, keyInfoRequest, parseBalance, parseKeyInfo, usd } from "../../src/core/credits";
import { ApiError, classifyError, describeError, ensureOk, explainApiError, isRetryable, messageForReason } from "../../src/core/errors";
import { callModelStep, createStepAccumulator } from "../../src/core/agent/llm";
import { fetchModels } from "../../src/core/models";
import { Storage } from "../../src/core/storage";
import { createBackup } from "../../src/core/backup";

const OR = getProvider("openrouter")!;
const ANTHROPIC = getProvider("anthropic")!;
const GEMINI = getProvider("gemini")!;
const OPENAI = getProvider("openai")!;
const DEEPSEEK = getProvider("deepseek")!;

/** Chuỗi giả, không phải key hay mã thật */
const FAKE_KEY = "sk-or-v1-" + "0".repeat(24) + "gia";
const FAKE_CODE = "ma-gia-0123456789abcdef";

const json = (status: number, body: any, headers: Record<string, string> = {}) =>
  new Response(typeof body === "string" ? body : JSON.stringify(body), { status, headers: { "content-type": "application/json", ...headers } });

function memStore() {
  const m: Record<string, string> = {};
  return {
    data: m,
    getItem: (k: string) => (k in m ? m[k] : null),
    setItem: (k: string, v: string) => {
      m[k] = v;
    },
    removeItem: (k: string) => {
      delete m[k];
    },
  };
}

describe("Đợt 2 — mã kiểm và mã thách thức (PKCE)", () => {
  it("base64url không đệm, không có + / =", () => {
    assert.strictEqual(base64Url(new Uint8Array([251, 255, 254])), "-__-");
    assert.strictEqual(base64Url(new Uint8Array([1])), "AQ");
  });

  it("mã thách thức S256 đúng véc-tơ chuẩn của RFC 7636 (phụ lục B)", async () => {
    assert.strictEqual(await s256Challenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk"), "E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM");
  });

  it("mã kiểm dài 43 ký tự base64url, mỗi lần một khác; state cũng vậy", () => {
    const a = newVerifier();
    const b = newVerifier();
    assert.match(a, /^[A-Za-z0-9_-]{43}$/);
    assert.notStrictEqual(a, b);
    assert.match(newState(), /^[A-Za-z0-9_-]{22}$/);
    assert.notStrictEqual(newState(), newState());
  });

  it("lần đăng nhập: mã thách thức khớp mã kiểm", async () => {
    const s = await createLoginSession(1000);
    assert.strictEqual(s.challenge, await s256Challenge(s.verifier));
    assert.ok(validChallenge(s.challenge));
    assert.strictEqual(s.createdAt, 1000);
  });
});

describe("Đợt 2 — địa chỉ đăng nhập", () => {
  const challenge = "E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM";
  const state = "abcDEF123_-xyz0123456x";

  it("OpenRouter được khai bằng cấu hình; nguồn khác chưa có đăng nhập", () => {
    assert.ok(OR.oauth);
    assert.strictEqual(OR.oauth!.authUrl, "https://openrouter.ai/auth");
    assert.strictEqual(OR.oauth!.exchangeUrl, "https://openrouter.ai/api/v1/auth/keys");
    assert.deepStrictEqual(
      PROVIDERS.filter((p) => p.oauth).map((p) => p.id),
      ["openrouter"]
    );
  });

  it("lối hộp thoại: có callback_url, code_challenge, S256, state, nhãn key", () => {
    const u = new URL(buildAuthUrl(OR.oauth!, { challenge, state, callbackUrl: "https://dunghoang1985.github.io/ai-da-nguon/dang-nhap.html" }));
    assert.strictEqual(u.origin + u.pathname, "https://openrouter.ai/auth");
    assert.strictEqual(u.searchParams.get("callback_url"), "https://dunghoang1985.github.io/ai-da-nguon/dang-nhap.html");
    assert.strictEqual(u.searchParams.get("code_challenge"), challenge);
    assert.strictEqual(u.searchParams.get("code_challenge_method"), "S256");
    assert.strictEqual(u.searchParams.get("state"), state);
    assert.ok(u.searchParams.get("key_label"));
  });

  it("lối dán mã: không có callback_url và state (nhà cung cấp hiện mã để chép)", () => {
    const u = new URL(buildAuthUrl(OR.oauth!, { challenge, state }));
    assert.strictEqual(u.searchParams.get("callback_url"), null);
    assert.strictEqual(u.searchParams.get("state"), null);
    assert.strictEqual(u.searchParams.get("code_challenge"), challenge);
  });

  it("địa chỉ không bao giờ chứa mã kiểm", async () => {
    const s = await createLoginSession();
    const all = buildAuthUrl(OR.oauth!, { challenge: s.challenge, state: s.state, callbackUrl: "https://a.example/dang-nhap.html" }) + dialogStartUrl("https://a.example/dang-nhap.html", "openrouter", s);
    assert.ok(!all.includes(s.verifier));
  });

  it("từ chối mã thách thức sai dạng và cấu hình trỏ sai tên miền", () => {
    assert.throws(() => buildAuthUrl(OR.oauth!, { challenge: "ngan" }), OAuthError);
    assert.throws(() => buildAuthUrl({ ...OR.oauth!, authUrl: "https://xau.example/auth" }, { challenge }), OAuthError);
    assert.throws(() => buildAuthUrl({ ...OR.oauth!, authUrl: "http://openrouter.ai/auth" }, { challenge }), OAuthError);
  });

  it("địa chỉ mở hộp thoại: trang của add-in, bỏ tham số cũ, chỉ mang nguồn, mã thách thức, state", () => {
    const u = new URL(dialogStartUrl("https://dunghoang1985.github.io/ai-da-nguon/dang-nhap.html?_host_Info=x#y", "openrouter", { challenge, state }));
    assert.strictEqual(u.origin + u.pathname, "https://dunghoang1985.github.io/ai-da-nguon/dang-nhap.html");
    assert.deepStrictEqual([...u.searchParams.keys()], ["p", "c", "s"]);
    assert.strictEqual(u.hash, "");
  });

  it("trang hộp thoại đọc địa chỉ: lần mở đầu, lần chuyển về, bị từ chối, địa chỉ lạ", () => {
    assert.deepStrictEqual(parseDialogQuery(`?p=openrouter&c=${challenge}&s=${state}`), { kind: "start", providerId: "openrouter", challenge, state });
    assert.deepStrictEqual(parseDialogQuery(`?code=${FAKE_CODE}&state=${state}`), { kind: "callback", code: FAKE_CODE, state });
    assert.deepStrictEqual(parseDialogQuery(`?code=${FAKE_CODE}`), { kind: "callback", code: FAKE_CODE });
    assert.strictEqual(parseDialogQuery("?error=access_denied").kind, "denied");
    assert.strictEqual(parseDialogQuery("").kind, "none");
    // Tham số sai dạng: không chuyển đi đâu cả
    assert.strictEqual(parseDialogQuery(`?p=openrouter&c=ngan&s=${state}`).kind, "none");
    assert.strictEqual(parseDialogQuery(`?p=https://xau.example&c=${challenge}&s=${state}`).kind, "none");
    assert.strictEqual(parseDialogQuery("?code=co%20dau%20cach").kind, "none");
  });
});

describe("Đợt 2 — lối dán mã: nhận và từ chối", () => {
  it("nhận mã trần, bỏ khoảng trắng, dấu nháy, ký tự ẩn", () => {
    assert.deepStrictEqual(extractCode(`  ${FAKE_CODE}\n`, OR), { kind: "code", code: FAKE_CODE });
    assert.deepStrictEqual(extractCode(`"${FAKE_CODE}"`, OR), { kind: "code", code: FAKE_CODE });
    assert.deepStrictEqual(extractCode(`​${FAKE_CODE}﻿`, OR), { kind: "code", code: FAKE_CODE });
  });

  it("nhận cả khi dán nguyên địa chỉ có ?code=…", () => {
    assert.deepStrictEqual(extractCode(`https://a.example/dang-nhap.html?code=${FAKE_CODE}&state=abc`, OR), { kind: "code", code: FAKE_CODE });
    assert.deepStrictEqual(extractCode(`code=${FAKE_CODE}`, OR), { kind: "code", code: FAKE_CODE });
  });

  it("từ chối: rỗng, có dấu cách ở giữa, chữ có dấu, địa chỉ không có mã, quá ngắn", () => {
    for (const bad of ["", "   ", "ma co dau cach", "mã-có-dấu-12345", "https://openrouter.ai/auth", "abc"]) {
      const r = extractCode(bad, OR);
      assert.strictEqual(r.kind, "invalid", bad);
      assert.ok((r as any).reason.length > 10);
    }
  });

  it("dán nhầm API key vào ô mã thì nhận ra là key", () => {
    assert.deepStrictEqual(extractCode(FAKE_KEY, OR), { kind: "key", key: FAKE_KEY });
  });
});

describe("Đợt 2 — đổi mã lấy key", () => {
  it("yêu cầu đúng tài liệu: POST JSON gồm code, code_verifier, code_challenge_method", () => {
    const r = exchangeRequest(OR.oauth!, FAKE_CODE, "ma-kiem");
    assert.strictEqual(r.url, "https://openrouter.ai/api/v1/auth/keys");
    assert.strictEqual(r.init.method, "POST");
    assert.deepStrictEqual(JSON.parse(String(r.init.body)), { code: FAKE_CODE, code_verifier: "ma-kiem", code_challenge_method: CHALLENGE_METHOD });
    // Không gửi kèm khóa hay tiêu đề xác thực nào
    assert.deepStrictEqual(Object.keys(r.init.headers as any), ["Content-Type"]);
  });

  it("200 có key: trả key", async () => {
    let seen: any;
    const key = await exchangeCode(OR, FAKE_CODE, "ma-kiem", (async (url: any, init: any) => {
      seen = { url, init };
      return json(200, { key: FAKE_KEY, user_id: "u" });
    }) as any);
    assert.strictEqual(key, FAKE_KEY);
    assert.strictEqual(seen.url, OR.oauth!.exchangeUrl);
  });

  it("200 không có key: báo lỗi rõ", async () => {
    await assert.rejects(exchangeCode(OR, FAKE_CODE, "v", (async () => json(200, {})) as any), /không kèm key/);
  });

  const cases: [number, any, RegExp][] = [
    [403, { error: { code: 403, message: "Invalid code or code_verifier" } }, /không nhận mã này/],
    [403, { error: { code: 403, message: "Authorization code expired" } }, /hết hạn.*10 phút/],
    // Câu thật của máy chủ với mã giả (đo 16/11/2026): 400, không phải 403 như tài liệu
    [400, { error: { code: 400, message: "Invalid code" } }, /không nhận mã này/],
    [400, { error: { code: 400, message: "Invalid code_challenge_method" } }, /từ chối yêu cầu đổi mã \(400\)/],
    [405, "Method Not Allowed", /405/],
    [429, { error: { message: "Too many requests" } }, /quá nhiều/],
    [502, "<html>bad gateway</html>", /đang lỗi \(502\)/],
  ];
  for (const [status, body, re] of cases) {
    it(`lỗi ${status} (${typeof body === "string" ? body.slice(0, 20) : body.error.message}): câu tiếng Việt đúng lý do`, async () => {
      await assert.rejects(
        exchangeCode(OR, FAKE_CODE, "ma-kiem-bi-mat-0123456789", (async () => json(status, body)) as any),
        (e: any) => e instanceof OAuthError && e.status === status && re.test(e.message)
      );
    });
  }

  it("mất mạng: báo không kết nối được, không ném lỗi thô", async () => {
    await assert.rejects(
      exchangeCode(OR, FAKE_CODE, "v", (async () => {
        throw new TypeError("Failed to fetch");
      }) as any),
      (e: any) => e instanceof OAuthError && /Không kết nối được/.test(e.message)
    );
  });

  it("thông báo lỗi không bao giờ chứa mã đăng nhập, mã kiểm hay key — kể cả khi máy chủ nhắc lại", async () => {
    const verifier = "ma-kiem-bi-mat-0123456789-ABCDEFGHIJKLMNOPQ";
    const body = { error: { message: `Invalid code ${FAKE_CODE} for verifier ${verifier} and key ${FAKE_KEY}` } };
    await assert.rejects(exchangeCode(OR, FAKE_CODE, verifier, (async () => json(403, body)) as any), (e: any) => {
      assert.ok(!e.message.includes(FAKE_CODE), "lộ mã");
      assert.ok(!e.message.includes(verifier), "lộ mã kiểm");
      assert.ok(!e.message.includes(FAKE_KEY), "lộ key");
      return true;
    });
    assert.strictEqual(redact("a sk-or-v1-abcdefgh123 b", []), "a sk-… b");
    assert.ok(!exchangeErrorMessage("OpenRouter", 403, "").includes("undefined"));
  });

  it("nguồn không có cấu hình đăng nhập: từ chối", async () => {
    await assert.rejects(exchangeCode(ANTHROPIC, FAKE_CODE, "v", (async () => json(200, { key: "x" })) as any), /không hỗ trợ đăng nhập/);
  });
});

describe("Đợt 2 — tin nhắn từ hộp thoại về khung", () => {
  const session = { state: "state-dung-0123456789" };
  it("đúng loại, đúng state: nhận mã", () => {
    assert.deepStrictEqual(readDialogMessage(dialogMessage({ code: FAKE_CODE, state: session.state }), session), { ok: true, code: FAKE_CODE });
  });
  it("state không khớp, thiếu state, không còn lần đăng nhập, dữ liệu lạ: từ chối", () => {
    assert.strictEqual(readDialogMessage(dialogMessage({ code: FAKE_CODE, state: "state-khac-0123456789" }), session).ok, false);
    assert.strictEqual(readDialogMessage(dialogMessage({ code: FAKE_CODE }), session).ok, false);
    assert.strictEqual(readDialogMessage(dialogMessage({ code: FAKE_CODE, state: session.state }), null).ok, false);
    assert.strictEqual(readDialogMessage("không phải json", session).ok, false);
    assert.strictEqual(readDialogMessage(JSON.stringify({ t: "khac", code: FAKE_CODE, state: session.state }), session).ok, false);
    assert.strictEqual(readDialogMessage(JSON.stringify({ t: "aidn-oauth", code: "co dau cach", state: session.state }), session).ok, false);
    assert.strictEqual(readDialogMessage(JSON.stringify({ t: "aidn-oauth", denied: true }), session).ok, false);
  });
});

describe("Đợt 2 — đăng xuất xóa sạch khóa trên máy", () => {
  it("removeKey xóa khóa và danh mục model của nguồn đó, giữ nguyên nguồn khác", () => {
    const st = memStore();
    const s = new Storage(st);
    s.setKey("openrouter", FAKE_KEY);
    s.setKey("groq", "gsk_gia_0123456789");
    s.setModels("openrouter", [{ id: "a/b:free", name: "A" }]);
    s.setModels("groq", [{ id: "g", name: "G" }]);
    s.removeKey("openrouter");
    assert.strictEqual(s.getKey("openrouter"), "");
    assert.ok(!("openrouter" in s.getKeys()));
    assert.ok(!("openrouter" in s.getModelCache()));
    assert.strictEqual(s.getKey("groq"), "gsk_gia_0123456789");
    assert.ok("groq" in s.getModelCache());
    // Không còn sót ở bất kỳ mục nào của kho
    assert.ok(!JSON.stringify(st.data).includes(FAKE_KEY));
  });

  it("sau khi đăng xuất, tệp sao lưu mới không còn khóa đó", async () => {
    const st = memStore();
    const s = new Storage(st);
    s.setKey("openrouter", FAKE_KEY);
    s.removeKey("openrouter");
    const kv = { keys: async () => [] as string[], get: async () => undefined, set: async () => {}, del: async () => {} };
    const ls = {
      get length() {
        return Object.keys(st.data).length;
      },
      key: (i: number) => Object.keys(st.data)[i] ?? null,
      getItem: st.getItem,
      setItem: st.setItem,
      removeItem: st.removeItem,
    };
    const b = await createBackup(ls, kv as any, { appVersion: "4.2.0", host: "Word" });
    assert.ok("aidn.keys" in b.local, "sao lưu vẫn có mục key (của nguồn khác)");
    assert.ok(!JSON.stringify(b).includes(FAKE_KEY));
  });
});

describe("Đợt 2 — hạn mức OpenRouter", () => {
  const sample = (over: any = {}) => ({
    data: {
      label: "sk-or-v1-abc…xyz",
      limit: null,
      limit_remaining: null,
      usage: 1.25,
      usage_daily: 0.02,
      usage_monthly: 0.9,
      is_free_tier: true,
      free_model_daily_requests: { used: 12, limit: 50, remaining: 38 },
      ...over,
    },
  });

  it("yêu cầu GET /key kèm Bearer", () => {
    const r = keyInfoRequest(OR, FAKE_KEY);
    assert.strictEqual(r.url, "https://openrouter.ai/api/v1/key");
    assert.strictEqual((r.init.headers as any).Authorization, `Bearer ${FAKE_KEY}`);
  });

  it("đọc đủ trường; thiếu remaining thì tự tính", () => {
    const i = parseKeyInfo(sample(), 5);
    assert.deepStrictEqual(i.free, { used: 12, limit: 50, remaining: 38 });
    assert.strictEqual(i.limit, null);
    assert.strictEqual(i.usageDaily, 0.02);
    assert.strictEqual(i.at, 5);
    assert.strictEqual(parseKeyInfo(sample({ free_model_daily_requests: { used: 48, limit: 50 } })).free!.remaining, 2);
    assert.strictEqual(parseBalance({ data: { total_credits: 10, total_usage: 2.5 } }), 7.5);
    assert.strictEqual(parseBalance({ data: {} }), undefined);
  });

  it("còn nhiều: không cảnh báo; thẻ có dòng lượt miễn phí và gợi ý 1.000 lượt", () => {
    const v = describeKeyInfo(parseKeyInfo(sample()), { balanceAllowed: false });
    assert.strictEqual(v.level, "ok");
    assert.deepStrictEqual(v.warnings, []);
    assert.ok(v.lines.some((l) => /còn 38\/50 lượt/.test(l)));
    assert.ok(v.lines.some((l) => /1\.000 lượt/.test(l)));
    assert.ok(v.lines.some((l) => /không được xem/.test(l)));
    assert.ok(v.lines.some((l) => /hôm nay 0,02 USD · tháng này 0,90 USD/.test(l)));
  });

  it("còn ít lượt miễn phí: cảnh báo; hết: báo hết", () => {
    const low = describeKeyInfo(parseKeyInfo(sample({ free_model_daily_requests: { used: 47, limit: 50, remaining: 3 } })));
    assert.strictEqual(low.level, "warn");
    assert.match(low.warnings[0], /Chỉ còn 3 lượt/);
    const out = describeKeyInfo(parseKeyInfo(sample({ free_model_daily_requests: { used: 50, limit: 50, remaining: 0 } })));
    assert.strictEqual(out.level, "out");
    assert.match(out.warnings[0], /hết lượt/);
  });

  it("key có trần: còn dưới 1 USD thì cảnh báo, chạm trần thì báo hết", () => {
    const low = describeKeyInfo(parseKeyInfo(sample({ limit: 5, limit_remaining: 0.4 })));
    assert.strictEqual(low.level, "warn");
    assert.ok(low.lines.some((l) => /còn 0,40 USD trên 5,00 USD/.test(l)));
    const out = describeKeyInfo(parseKeyInfo(sample({ limit: 5, limit_remaining: 0 })));
    assert.strictEqual(out.level, "out");
  });

  it("số dư tài khoản: âm thì báo hết, thấp (tài khoản đã nạp) thì cảnh báo", () => {
    const neg = describeKeyInfo({ ...parseKeyInfo(sample()), balance: -0.5 });
    assert.strictEqual(neg.level, "out");
    assert.ok(neg.lines.some((l) => /Số dư tài khoản: -0,50 USD/.test(l)));
    const low = describeKeyInfo({ ...parseKeyInfo(sample({ is_free_tier: false, free_model_daily_requests: { used: 0, limit: 1000, remaining: 1000 } })), balance: 0.3 });
    assert.strictEqual(low.level, "warn");
    // Tài khoản chưa từng nạp, số dư 0: bình thường, không cảnh báo
    assert.strictEqual(describeKeyInfo({ ...parseKeyInfo(sample()), balance: 0 }).level, "ok");
  });

  it("tiền kiểu Việt", () => {
    assert.strictEqual(usd(1.5), "1,50 USD");
    assert.strictEqual(usd(0.0031), "0,0031 USD");
    assert.strictEqual(usd(0), "0,00 USD");
  });

  it("dòng cảnh báo ở khung chat: chỉ khi thấp, và đúng loại model đang chọn", () => {
    const many = parseKeyInfo(sample());
    assert.strictEqual(chatCreditWarning(many, { freeModel: true }), undefined);
    assert.strictEqual(chatCreditWarning(undefined), undefined);
    const few = parseKeyInfo(sample({ free_model_daily_requests: { used: 46, limit: 50, remaining: 4 } }));
    assert.match(chatCreditWarning(few, { freeModel: true })!.text, /còn 4\/50 lượt/);
    // Đang chọn model trả tiền thì số lượt miễn phí không liên quan
    assert.strictEqual(chatCreditWarning(few, { freeModel: false }), undefined);
    const none = parseKeyInfo(sample({ free_model_daily_requests: { used: 50, limit: 50, remaining: 0 } }));
    assert.strictEqual(chatCreditWarning(none, { freeModel: true })!.level, "out");
    assert.strictEqual(chatCreditWarning(parseKeyInfo(sample({ limit: 5, limit_remaining: 0 })), { freeModel: true })!.level, "out");
  });

  it("fetchKeyInfo: /credits bị 403 (key thường) thì vẫn trả phần của /key, không coi là lỗi", async () => {
    const calls: string[] = [];
    const r = await fetchKeyInfo(OR, FAKE_KEY, {
      tryBalance: true,
      fetchFn: (async (url: any) => {
        calls.push(String(url));
        return String(url).endsWith("/key") ? json(200, sample()) : json(403, { error: { code: 403, message: "Only management keys can perform this operation" } });
      }) as any,
    });
    assert.deepStrictEqual(calls, ["https://openrouter.ai/api/v1/key", "https://openrouter.ai/api/v1/credits"]);
    assert.strictEqual(r.balanceAllowed, false);
    assert.strictEqual(r.info.balance, undefined);
    assert.strictEqual(r.info.free!.remaining, 38);
  });

  it("fetchKeyInfo: /credits được phép thì có số dư; không bật tryBalance thì không gọi", async () => {
    const r = await fetchKeyInfo(OR, FAKE_KEY, {
      tryBalance: true,
      fetchFn: (async (url: any) => (String(url).endsWith("/key") ? json(200, sample()) : json(200, { data: { total_credits: 10, total_usage: 4 } }))) as any,
    });
    assert.strictEqual(r.info.balance, 6);
    assert.strictEqual(r.balanceAllowed, true);
    let n = 0;
    await fetchKeyInfo(OR, FAKE_KEY, {
      fetchFn: (async () => {
        n++;
        return json(200, sample());
      }) as any,
    });
    assert.strictEqual(n, 1);
  });

  it("fetchKeyInfo: key sai (401) thì báo lỗi tiếng Việt về key", async () => {
    await assert.rejects(
      fetchKeyInfo(OR, FAKE_KEY, { fetchFn: (async () => json(401, { error: { code: 401, message: "User not found." } })) as any }),
      (e: any) => e instanceof ApiError && e.reason === "key_invalid" && /Key không đúng: OpenRouter không nhận/.test(e.message)
    );
  });
});

describe("Đợt 2 — phân loại lỗi key và tiền", () => {
  it("Anthropic hết tín dụng (400, câu trong tài liệu): hết tiền, có gợi ý đổi nguồn, không thử lại", () => {
    const detail = "Your credit balance is too low to access the Anthropic API. Please go to Plans & Billing to upgrade or purchase credits.";
    const body = { type: "error", error: { type: "invalid_request_error", message: detail } };
    assert.deepStrictEqual(classifyError(400, detail, body), { reason: "no_credit" });
    const x = explainApiError(400, detail, { provider: ANTHROPIC, body });
    assert.match(x.text, /Anthropic \(Claude\): tài khoản đã hết tiền/);
    assert.match(x.text, /đổi sang nguồn khác ở nhãn model/);
    assert.match(x.text, /tín dụng kỳ này đã hết/);
    assert.match(x.text, /Hội thoại vẫn còn nguyên/);
    assert.ok(!isRetryable(new ApiError(x.text, 400, undefined, x.reason)));
  });

  it("Anthropic: 402 billing_error, 401 key sai, 401 key hết hạn, 401 tổ chức chặn trình duyệt, 400 key chưa gắn workspace", () => {
    assert.strictEqual(classifyError(402, "Payment issue", { error: { type: "billing_error", message: "Payment issue" } }).reason, "no_credit");
    assert.strictEqual(classifyError(401, "invalid x-api-key", { error: { type: "authentication_error", message: "invalid x-api-key" } }).reason, "key_invalid");
    assert.strictEqual(classifyError(401, "API key has expired", { error: { type: "authentication_error" } }).reason, "key_expired");
    assert.strictEqual(classifyError(401, "CORS requests are not allowed for this Organization because of its settings").reason, "browser_blocked");
    const ws = "This API key is not scoped to a workspace, so this request must include the anthropic-workspace-id header with the ID of the workspace to use.";
    assert.strictEqual(classifyError(400, ws).reason, "workspace");
    assert.match(messageForReason("workspace", ANTHROPIC), /chọn một workspace/);
    assert.match(messageForReason("browser_blocked", ANTHROPIC), /key không sai/);
    assert.strictEqual(classifyError(400, "You have reached your specified API usage limits. You will regain access on 2026-12-01.").reason, "key_limit");
  });

  it("Gemini: hết tiền trả trước (429) là hết tiền, không thử lại; hết hạn mức thường vẫn thử lại", () => {
    const depleted = "Your prepayment credits are depleted. Please go to AI Studio at https://ai.studio/projects to manage your project and billing.";
    const body = { error: { code: 429, message: depleted, status: "RESOURCE_EXHAUSTED" } };
    assert.strictEqual(classifyError(429, depleted, body).reason, "no_credit");
    assert.match(explainApiError(429, depleted, { provider: GEMINI, body }).text, /Google Gemini: tài khoản đã hết tiền/);
    const quota = "You exceeded your current quota, please check your plan and billing details.";
    const qBody = { error: { code: 429, message: quota, status: "RESOURCE_EXHAUSTED" } };
    assert.deepStrictEqual(classifyError(429, quota, qBody), {});
    assert.match(explainApiError(429, quota, { provider: GEMINI, body: qBody }).text, /Hết hạn mức/);
    assert.ok(isRetryable(new ApiError("x", 429)));
  });

  it("Gemini: key sai (400 API_KEY_INVALID), key hết hạn, key bị báo lộ", () => {
    const bad = { error: { code: 400, message: "API key not valid. Please pass a valid API key.", status: "INVALID_ARGUMENT", details: [{ reason: "API_KEY_INVALID" }] } };
    assert.strictEqual(classifyError(400, bad.error.message, bad).reason, "key_invalid");
    assert.strictEqual(classifyError(400, "API key expired. Please renew the API key.", { error: { details: [{ reason: "API_KEY_INVALID" }] } }).reason, "key_expired");
    assert.strictEqual(classifyError(403, "Your API key was reported as leaked. Please use another API key.").reason, "key_expired");
  });

  it("OpenAI: hết tiền (429 insufficient_quota), key sai (401 invalid_api_key)", () => {
    const body = { error: { message: "You exceeded your current quota, please check your plan and billing details.", type: "insufficient_quota", code: "insufficient_quota" } };
    assert.strictEqual(classifyError(429, body.error.message, body).reason, "no_credit");
    assert.match(explainApiError(429, body.error.message, { provider: OPENAI, body }).text, /OpenAI: tài khoản đã hết tiền/);
    assert.strictEqual(classifyError(401, "Incorrect API key provided: sk-…", { error: { code: "invalid_api_key" } }).reason, "key_invalid");
  });

  it("DeepSeek: 402 Insufficient Balance là hết tiền", () => {
    assert.match(explainApiError(402, "Insufficient Balance", { provider: DEEPSEEK }).text, /DeepSeek: tài khoản đã hết tiền/);
  });

  it("OpenRouter 402 theo nguồn giới hạn: hết credit, chạm trần key, ngân sách đang chạy (tạm thời, thử lại được)", () => {
    const mk = (src: string) => ({ error: { code: 402, message: "Insufficient credits", metadata: { limit_source: src } } });
    assert.strictEqual(classifyError(402, "Insufficient credits", mk("openrouter_credits")).reason, "no_credit");
    assert.strictEqual(classifyError(402, "Key limit exceeded", mk("openrouter_key_limit")).reason, "key_limit");
    assert.deepStrictEqual(classifyError(402, "Insufficient credits", mk("openrouter_in_flight_budget")), { transient: true });
    const x = explainApiError(402, "Insufficient credits", { provider: OR, body: mk("openrouter_in_flight_budget") });
    const err = new ApiError(x.text, 402);
    err.transient = x.transient;
    assert.ok(isRetryable(err));
    assert.strictEqual(classifyError(403, "Budget exceeded", mk("openrouter_workspace_budget")).reason, "key_limit");
    assert.strictEqual(classifyError(402, "This request requires more credits, or fewer max_tokens. You requested up to 4096 tokens, but can only afford 120.").reason, "no_credit");
  });

  it("OpenRouter: hết lượt miễn phí trong ngày (429) — không phải hết tiền, không thử lại; hết lượt theo phút vẫn thử lại", () => {
    const day = "Rate limit exceeded: free-models-per-day. Add 10 credits to unlock 1000 free model requests per day";
    assert.strictEqual(classifyError(429, day).reason, "free_limit");
    assert.match(explainApiError(429, day, { provider: OR }).text, /hết lượt dùng model miễn phí của hôm nay/);
    assert.deepStrictEqual(classifyError(429, "Rate limit exceeded: free-models-per-min."), {});
  });

  it("OpenRouter: 401 key bị xóa hoặc sai; 401 model bị giới hạn nơi dùng thì không đổ cho key", () => {
    assert.match(explainApiError(401, "User not found.", { provider: OR }).text, /Key không đúng: OpenRouter không nhận.*Đăng nhập lại/);
    assert.match(explainApiError(401, "This key has been disabled", { provider: OR }).text, /hết hạn, bị tắt hoặc bị thu hồi/);
    assert.match(explainApiError(401, "x is only available on agentic harnesses", { provider: OR }).text, /từ chối yêu cầu \(401\)/);
  });

  it("lỗi thường không bị gán lý do: 400 sai tham số, 404, 429 thường, 500, 403 không rõ", () => {
    for (const [st, d] of [
      [400, "messages.0.content: Field required"],
      [404, "model not found"],
      [429, "Rate limit reached for requests"],
      [500, "Internal error"],
      [403, "This model is only available in agentic harnesses"],
      [400, "max_tokens: 99999 > 32000, which is the maximum allowed"],
      [400, "tool_choice: type \"tool\" and \"any\" are not supported for this model."],
    ] as [number, string][]) {
      assert.deepStrictEqual(classifyError(st, d), {}, `${st} ${d}`);
    }
  });

  it("ensureOk: gắn lý do, tên nguồn, và giữ mã trạng thái", async () => {
    const depleted = { error: { code: 429, message: "Your prepayment credits are depleted.", status: "RESOURCE_EXHAUSTED" } };
    await assert.rejects(ensureOk(json(429, depleted), GEMINI), (e: any) => e instanceof ApiError && e.status === 429 && e.reason === "no_credit" && /Google Gemini/.test(e.message));
    // Gemini có lúc trả mảng
    await assert.rejects(ensureOk(json(429, [depleted]), GEMINI), (e: any) => e.reason === "no_credit");
    const info = describeError(new ApiError("x", 402, undefined, "no_credit"), { provider: OR });
    assert.strictEqual(info.reason, "no_credit");
  });
});

describe("Đợt 2 — lỗi hết tiền trong lượt gọi model", () => {
  const base = { model: "m", messages: [{ role: "user" as const, content: "chào" }], tools: [], onText: () => {} };

  it("Gemini báo hết tiền bằng 429: gọi đúng một lần, không tự thử lại", async () => {
    let n = 0;
    await assert.rejects(
      callModelStep({
        ...base,
        provider: GEMINI,
        key: "k",
        retryDelaysMs: [1, 1, 1],
        fetchFn: (async () => {
          n++;
          return json(429, { error: { code: 429, message: "Your prepayment credits are depleted.", status: "RESOURCE_EXHAUSTED" } });
        }) as any,
      }),
      (e: any) => e.reason === "no_credit"
    );
    assert.strictEqual(n, 1);
  });

  it("OpenRouter 402 do ngân sách đang chạy: tự thử lại rồi chạy được", async () => {
    let n = 0;
    const r = await callModelStep({
      ...base,
      provider: OR,
      key: "k",
      retryDelaysMs: [1, 1],
      sleep: async () => {},
      fetchFn: (async () => {
        n++;
        if (n === 1) return json(402, { error: { code: 402, message: "Insufficient credits", metadata: { limit_source: "openrouter_in_flight_budget" } } });
        return new Response('data: {"choices":[{"delta":{"content":"xin chào"},"finish_reason":"stop"}]}\n\ndata: [DONE]\n\n', { status: 200 });
      }) as any,
    });
    assert.strictEqual(n, 2);
    assert.strictEqual(r.text, "xin chào");
  });

  it("lỗi hết credit báo giữa luồng (đã trả 200): vẫn ra câu hết tiền kèm lý do", () => {
    const acc = createStepAccumulator("openai", () => {}, undefined, OR);
    assert.throws(
      () => acc.push(JSON.stringify({ error: { code: 402, message: "Insufficient credits", metadata: { limit_source: "openrouter_credits" } }, choices: [{ finish_reason: "error" }] })),
      (e: any) => e instanceof ApiError && e.reason === "no_credit" && /OpenRouter: tài khoản đã hết tiền/.test(e.message)
    );
    // Lỗi khác giữa luồng giữ cách báo cũ
    const acc2 = createStepAccumulator("anthropic", () => {});
    assert.throws(() => acc2.push(JSON.stringify({ type: "error", error: { type: "overloaded_error", message: "Overloaded" } })), (e: any) => e.status === 529 && !e.reason && /Nhà cung cấp báo lỗi/.test(e.message));
  });

  it("tải danh sách model với key sai: thông báo nêu tên nguồn và lý do", async () => {
    await assert.rejects(
      fetchModels(ANTHROPIC, "sai", (async () => json(401, { type: "error", error: { type: "authentication_error", message: "invalid x-api-key" } })) as any),
      (e: any) => e.reason === "key_invalid" && /Key không đúng: Anthropic \(Claude\) không nhận/.test(e.message)
    );
  });
});
