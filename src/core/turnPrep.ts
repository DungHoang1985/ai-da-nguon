/** Chuẩn bị một lượt — phần thuần tách khỏi Chat.tsx (B7 phần 1, nhóm 2 rà soát kiến trúc, 26/11/2026) để kiểm thử máy được:
 *  xử lý tệp đính kèm (đường đi của PDF quét, trần dung lượng, ảnh, thông báo), gợi ý nhóm công cụ, xử lý một lệnh ghi chéo.
 *  Chỉ dời nguyên văn logic; trạng thái React, hẹn giờ, kho cục bộ ở lại Chat.tsx. */
import { Attachment, DOCX_MIME, MAX_FILES_TOTAL_B64, PDF_MIME, withAttachments } from "./attachments";
import { attachmentPrivacyNote, routeScanPdf } from "./pdfPages";
import { pdfCapable } from "./agent/thinking";
import { Provider } from "./providers";
import { UndoEntry } from "./agent/types";

export interface Notice {
  text: string;
  err?: boolean;
}

export interface AttachmentsInput {
  /** Nội dung tin nhắn đã ghép lệnh /kỹ năng và vùng chọn */
  content: string;
  /** Tệp đính kèm của lượt (lượt "Làm tiếp": rỗng) */
  atts: Attachment[];
  host: string;
  provider: Provider;
  modelId: string;
  modelFree?: boolean;
  vision?: boolean;
  /** AI tự thao tác (bật và model gọi được công cụ) */
  useAgent: boolean;
  /** Công tắc "AI tự thao tác" của khung (chưa xét model) */
  agentOn: boolean;
  /** Tổng ký tự base64 của tệp đã nằm trong lịch sử (filesBase64Chars) */
  historyFilesB64: number;
}

export type AttachmentsResult =
  | { stop: true; notice: Notice }
  | {
      stop: false;
      content: string;
      images: { mime: string; data: string }[];
      pdfs: { mime: string; data: string; name: string }[];
      /** Tệp cần đăng ký vào bộ nhớ phiên (so sánh phiên bản, đọc theo trang, chèn ảnh) — theo đúng thứ tự cũ */
      register: { name: string; data: string }[];
      /** Thông báo đặt SAU thông báo đọc vùng chọn (ghi đè nó, như trước) */
      notice?: Notice;
    };

export function prepareAttachments(o: AttachmentsInput): AttachmentsResult {
  const { atts, provider: p, vision, useAgent, host } = o;
  let content = o.content;
  // PDF quét đi đường nào: nguyên tệp (tệp ngắn, nguồn đọc được) · AI đọc theo trang bằng công cụ · không đọc được
  const scans = atts.filter((a) => a.kind === "file" && a.data);
  const routeOf = (a: Attachment) => routeScanPdf({ pages: a.pages ?? 1, pdfCapable: pdfCapable(p), vision, tools: useAgent });
  const byPages = scans.filter((a) => routeOf(a) === "pages");
  const whole = scans.filter((a) => routeOf(a) !== "pages");
  // Tệp gửi nguyên vẹn nằm lại trong lịch sử và được gửi lại ở mọi lượt: quá trần thì nhà cung cấp từ chối mọi lượt sau
  const newFiles = whole.reduce((n, a) => n + a.data!.length, 0);
  if (newFiles && o.historyFilesB64 + newFiles > MAX_FILES_TOTAL_B64) {
    return {
      stop: true,
      notice: {
        text: "Các tệp PDF quét trong hội thoại này đã quá lớn để gửi tiếp (tệp được gửi lại ở mỗi lượt). Bấm “Hội thoại mới” rồi đính kèm lại tệp cần hỏi, hoặc tách tệp nhỏ hơn.",
        err: true,
      },
    };
  }
  const register: { name: string; data: string }[] = [];
  // Bản gốc của tệp .docx và .pdf: đăng ký (lại) lúc gửi để công cụ so sánh phiên bản và công cụ đọc theo trang đọc được
  atts.forEach((a) => (a.mime === DOCX_MIME || a.mime === PDF_MIME) && a.data && register.push({ name: a.name, data: a.data }));
  content = withAttachments(content, atts, { compare: host === "Word" && o.agentOn, pageTool: useAgent, byPages: byPages.map((a) => a.name), docTool: host === "Word" && useAgent });
  const images = atts.filter((a) => a.kind === "image" && a.data).map((a) => ({ mime: a.mime || "image/png", data: a.data! }));
  // PowerPoint: giữ ảnh đính kèm trong bộ nhớ của phiên và nêu tên để AI đặt được ảnh vào slide (ppt_insert_image)
  if (host === "PowerPoint" && useAgent) {
    const named = atts.filter((a) => a.kind === "image" && a.data);
    named.forEach((a) => register.push({ name: a.name, data: a.data! }));
    if (named.length) content = `${content}\n\n[Ảnh đính kèm: ${named.map((a) => `“${a.name}”`).join(", ")}. Muốn đặt ảnh vào slide thì gọi công cụ chèn ảnh đính kèm với image = tên ảnh.]`;
  }
  // PDF quét gửi nguyên tệp cho model tự đọc; nguồn không đọc được thì báo rõ (và model cũng được báo trong tin nhắn)
  const pdfs = whole.map((a) => ({ mime: a.mime || "application/pdf", data: a.data!, name: a.name }));
  let notice: Notice | undefined;
  if (pdfs.length && !pdfCapable(p)) {
    notice = {
      text: `${p.name} không đọc được PDF quét (${pdfs.map((f) => f.name).join(", ")}): tệp không được gửi đi. ${vision === true && !useAgent ? "Bật “AI tự thao tác” để AI đọc từng trang bằng ảnh, hoặc đổi" : "Đổi"} sang nguồn Anthropic hoặc Gemini rồi đính kèm lại.`,
      err: true,
    };
  } else if (atts.length) {
    // Nhắc riêng tư (không chặn): tệp đính kèm sắp tới nguồn miễn phí hoặc qua nơi trung chuyển
    const priv = attachmentPrivacyNote({ providerId: p.id, providerName: p.name, modelId: o.modelId, free: o.modelFree });
    if (priv) notice = { text: priv };
  }
  return { stop: false, content, images, pdfs, register, notice };
}

/** Nhóm công cụ bật thêm theo tệp đính kèm (chế độ Miễn phí): .docx trong Word → "soat"; ảnh (đính kèm lượt này hoặc còn trong
 *  phiên) trong PowerPoint → "hinh_anh" (lời "đặt hình đó vào slide 4" không có từ khóa — soát mã Đợt 11, #7) */
export function groupHintsFor(o: { host: string; atts: Attachment[]; images: number; sessionImages: number }): string[] {
  return [...(o.host === "Word" && o.atts.some((a) => a.mime === DOCX_MIME) ? ["soat"] : []), ...(o.host === "PowerPoint" && (o.images || o.sessionImages) ? ["hinh_anh"] : [])];
}

/** Kết quả xử lý một lệnh ghi chéo đã giành được */
export interface IncomingResult {
  ok: boolean;
  /** Trả cho khung gửi */
  reply: { ok: true; message: string } | { ok: false; error: string };
  /** Mục hiện trên khung nhận (chỉ khi ghi được) */
  item?: { role: "assistant"; text: ""; incoming: string; wrote: true; undo?: UndoEntry[] };
}

/** Xử lý một lệnh ghi chéo đã giành được: khung tắt kết nối với nguồn gửi → từ chối; ghi bằng `exec`; ghi dở thì gỡ phần đã làm rồi
 *  mới báo lỗi (để khung gửi không tưởng là đã ghi). Không đụng React hay kho cục bộ: bên gọi trả lời và thêm mục. */
export async function handleIncoming(
  cmd: { id: string; from: string; payload: any },
  o: {
    selfApp: string;
    linked: string[];
    what: string;
    exec: (payload: any, log: UndoEntry[]) => Promise<string>;
    errText: (e: any) => string;
    /** Trả lời khung gửi — gọi TRONG khối thử như mã cũ: trả lời thành công mà lỗi (kho cục bộ đầy…) thì gỡ phần đã ghi (phản biện 4a) */
    reply?: (res: { ok: true; message: string } | { ok: false; error: string }) => void;
  }
): Promise<IncomingResult> {
  const send = o.reply || (() => undefined);
  if (!o.linked.includes(cmd.from)) {
    const res = { ok: false as const, error: `Khung AI Đa Nguồn ở ${o.selfApp} đang tắt kết nối với ${cmd.from}.` };
    send(res);
    return { ok: false, reply: res };
  }
  const log: UndoEntry[] = [];
  try {
    const message = await o.exec(cmd.payload, log);
    const res = { ok: true as const, message };
    send(res);
    return {
      ok: true,
      reply: res,
      item: { role: "assistant", text: "", incoming: `Vừa ghi theo yêu cầu từ khung AI Đa Nguồn ở ${cmd.from}: ${o.what}.`, wrote: true, undo: log.length ? log : undefined },
    };
  } catch (e: any) {
    for (const u of log.slice().reverse()) await u.undo().catch(() => undefined);
    const res = { ok: false as const, error: `${o.selfApp} không ghi được: ${o.errText(e)}` };
    send(res);
    return { ok: false, reply: res };
  }
}
