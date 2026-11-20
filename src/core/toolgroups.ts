/** Nạp công cụ theo việc (Đợt 10): ở chế độ Miễn phí mỗi lần gọi model chỉ gửi nhóm công cụ lõi; nhóm chuyên được gửi khi
 *  yêu cầu của người dùng khớp từ khóa, kỹ năng tự chọn kéo theo, tệp đính kèm cần tới, khi model tự xin bằng load_tools,
 *  hoặc khi nhóm đã được dùng trước đó trong hội thoại. Chế độ Trả phí không dùng tới (gửi đủ công cụ, giữ bộ nhớ đệm).
 *  Nhóm đang bật được suy ra lại từ lịch sử ở mỗi bước: không có trạng thái riêng phải lưu. */
import { AgentMessage, ExtraTools, ToolSpec } from "./agent/types";

export interface ToolGroup {
  id: string;
  /** Tên nhóm tiếng Việt cho nhãn bước trên khung */
  label: string;
  tools: string[];
  /** Từ khóa: chuỗi chỉ có chữ ASCII so trên câu đã bỏ dấu; chuỗi có dấu so trên câu gốc (phân biệt "bảng" với "bằng") */
  keywords: string[];
  /** Kỹ năng tự chọn kéo theo nhóm này */
  skills?: string[];
}

export const LOAD_TOOLS = "load_tools";

export const WORD_TOOL_GROUPS: ToolGroup[] = [
  {
    id: "hanh_chinh",
    label: "văn bản hành chính, thể thức trang",
    tools: ["word_create_admin_document", "word_set_page_format", "word_check_page_format"],
    keywords: ["cong van", "to trinh", "bao cao", "quyet dinh", "thong bao", "giay moi", "bien ban", "ke hoach", "the thuc", "hanh chinh", "le trang", "can le", "kho giay", "kho a4", "phong chu", "co chu"],
    skills: ["soan-thao-van-ban-hanh-chinh", "soan-phan-hoi-phap-ly"],
  },
  {
    id: "dien_mau",
    label: "điền mẫu",
    tools: ["word_list_form_fields", "word_fill_form"],
    keywords: ["dien mau", "dien vao", "dien thong tin", "dien hop dong", "dien to khai", "dien bieu mau", "dien giup", "dien cac", "dien het", "dien day du", "to khai", "bieu mau", "cho trong", "hoan thien mau", "mau co san", "o dien"],
    skills: ["dien-mau"],
  },
  {
    id: "soat",
    label: "soát nhất quán, so phiên bản, phần ngoài thân văn bản",
    tools: ["word_check_consistency", "word_compare_versions", "word_read_parts"],
    keywords: ["soat", "kiem tra", "ra soat", "nhat quan", "truoc khi ky", "trinh ky", "ban cuoi", "so sanh", "phien ban", "doi chieu", "chu thich", "dau trang", "chan trang", "bookmark", "footnote", "header", "footer", "hop van ban", "loi chinh ta", "danh so"],
    skills: ["kiem-tra-truoc-ky", "ra-soat-hop-dong"],
  },
  {
    id: "nhan_xet_sua_doi",
    label: "nhận xét và sửa đổi theo dõi",
    tools: ["word_list_comments", "word_reply_comment", "word_resolve_comment", "word_list_revisions", "word_review_revision"],
    keywords: ["nhan xet", "comment", "gop y", "sua doi", "track", "redline", "doi tac", "chap nhan", "tu choi", "theo doi", "ai sua", "da sua gi"],
    skills: ["xu-ly-nhan-xet", "tom-tat-sua-doi-doi-tac", "kiem-tra-truoc-ky"],
  },
  {
    id: "bang",
    label: "bảng",
    tools: ["word_read_tables", "word_insert_table", "word_fill_table"],
    keywords: ["bảng", "table", "cột", "kẻ bảng"],
  },
  {
    id: "ngoai",
    label: "ứng dụng khác, tệp PDF đính kèm",
    tools: ["read_other_app", "write_other_app", "read_attachment_pages"],
    keywords: ["excel", "powerpoint", "slide", "trinh chieu", "bang tinh", "dinh kem", "pdf", "trang so", "ban quet", "ban scan"],
  },
];

/** Chữ thường, bỏ dấu tiếng Việt, gộp khoảng trắng */
export function fold(s: string): string {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");
}

const isAscii = (s: string) => /^[\x00-\x7f]*$/.test(s);

/** Nhóm bật theo lời người dùng của lượt này, kỹ năng tự chọn và gợi ý từ tệp đính kèm (id nhóm) */
export function requestGroups(groups: ToolGroup[], o: { text: string; skills?: string[]; hints?: string[] }): Set<string> {
  const raw = ` ${String(o.text || "").toLowerCase().replace(/\s+/g, " ")} `;
  const folded = ` ${fold(o.text)} `;
  const on = new Set<string>();
  for (const g of groups) {
    const hit = g.keywords.some((k) => (isAscii(k) ? folded.includes(k) : raw.includes(k.toLowerCase())));
    if (hit || g.skills?.some((s) => o.skills?.includes(s)) || o.hints?.includes(g.id)) on.add(g.id);
  }
  return on;
}

/** Nhóm đã dùng trong lịch sử: model đã gọi load_tools cho nhóm đó, hoặc đã gọi một công cụ của nhóm */
export function usedGroups(groups: ToolGroup[], msgs: AgentMessage[]): Set<string> {
  const byTool = new Map<string, string>();
  for (const g of groups) for (const t of g.tools) byTool.set(t, g.id);
  const ids = new Set(groups.map((g) => g.id));
  const on = new Set<string>();
  for (const m of msgs as any[]) {
    for (const c of m.toolCalls || []) {
      if (c?.name === LOAD_TOOLS) {
        for (const id of asList(c.args?.group)) if (ids.has(id)) on.add(id);
      } else if (c?.name === "use_skill") {
        // Model tự đọc một kỹ năng: bật luôn nhóm mà kỹ năng đó cần
        for (const g of groups) if (g.skills?.includes(String(c.args?.name || ""))) on.add(g.id);
      } else if (byTool.has(c?.name)) on.add(byTool.get(c.name)!);
    }
  }
  return on;
}

const asList = (v: any): string[] => (Array.isArray(v) ? v : v === undefined || v === null ? [] : String(v).split(/[,\s]+/)).map((x) => String(x).trim()).filter(Boolean);

/** Bộ công cụ gửi model: công cụ không thuộc nhóm nào (lõi) + công cụ của nhóm đang bật; thứ tự giữ như bộ đầy đủ */
export function visibleSpecs(all: ToolSpec[], groups: ToolGroup[], active: Set<string>): ToolSpec[] {
  const hidden = new Set<string>();
  for (const g of groups) if (!active.has(g.id)) for (const t of g.tools) hidden.add(t);
  return all.filter((t) => !hidden.has(t.name));
}

export function loadToolsSpec(groups: ToolGroup[]): ToolSpec {
  return {
    name: LOAD_TOOLS,
    description: "Make a group of specialised tools available from your next step (see the tool groups list in the system prompt).",
    parameters: {
      type: "object",
      properties: { group: { type: "string", enum: groups.map((g) => g.id) } },
      required: ["group"],
    },
  };
}

/** Dòng lời nhắc hệ thống (cố định trong cả hội thoại): danh sách nhóm và cách xin nhóm */
export function groupsNote(groups: ToolGroup[]): string {
  return [
    "## Nhóm công cụ chuyên",
    `Để tiết kiệm, chỉ công cụ cơ bản được gửi sẵn. Cần công cụ chuyên mà danh sách công cụ chưa có thì GỌI ${LOAD_TOOLS}(group) rồi dùng ở bước sau (có thể gọi cùng lúc với lệnh đọc). Các nhóm:`,
    ...groups.map((g) => `- ${g.id}: ${g.tools.join(", ")}`),
  ].join("\n");
}

/** Bọc công cụ dùng chung: thêm load_tools (chỉ trả lời, nhóm được bật nhờ usedGroups đọc lại lịch sử) */
export function withLoadTools(extra: ExtraTools | undefined, groups: ToolGroup[]): ExtraTools {
  const spec = loadToolsSpec(groups);
  return {
    specs: [...(extra?.specs || []), spec],
    execute: async (name, args) => {
      if (name !== LOAD_TOOLS) {
        if (!extra) throw new Error(`Không có công cụ tên "${name}".`);
        return extra.execute(name, args);
      }
      const want = asList(args?.group);
      const found = groups.filter((g) => want.includes(g.id));
      if (!found.length) throw new Error(`Không có nhóm "${want.join(", ")}". Các nhóm: ${groups.map((g) => g.id).join(", ")}.`);
      return { ok: true, loaded: found.map((g) => g.id), tools: found.flatMap((g) => g.tools), note: "Các công cụ này có từ bước tiếp theo." };
    },
  };
}
