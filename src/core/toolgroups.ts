/** Nạp công cụ theo việc (Đợt 10 Word, Đợt 11 Excel và PowerPoint): ở chế độ Miễn phí mỗi lần gọi model chỉ gửi nhóm công cụ lõi; nhóm chuyên được gửi khi
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

/** Đợt 12 phần E: nhóm bo_nho (công cụ propose_memory) ở chế độ Miễn phí chỉ gửi khi người dùng nói tới một quy ước lâu dài —
 *  nguồn miễn phí sức chứa nhỏ, không gửi mô tả công cụ này ở mọi lượt */
export const MEMORY_KEYWORDS = ["tu gio", "tu nay", "tu lan sau", "lan sau", "ve sau", "ghi nho", "hay nho", "nho giup", "nho rang", "nho la", "luon luon", "lan nao cung", "quy uoc", "nhớ"];


export const WORD_TOOL_GROUPS: ToolGroup[] = [
  {
    id: "hanh_chinh",
    label: "văn bản hành chính, thể thức trang",
    tools: ["word_create_admin_document", "word_set_page_format", "word_check_page_format"],
    keywords: ["cong van", "to trinh", "bao cao", "quyet dinh", "thong bao", "giay moi", "bien ban", "ke hoach", "the thuc", "hanh chinh", "le trang", "can le", "kho giay", "kho a4", "so trang", "danh so trang", "phong chu", "lề", "cỡ chữ"],
    skills: ["soan-thao-van-ban-hanh-chinh", "soan-phan-hoi-phap-ly"],
  },
  {
    id: "dien_mau",
    label: "điền mẫu",
    tools: ["word_list_form_fields", "word_fill_form"],
    keywords: ["dien mau", "dien vao", "dien thong tin", "dien hop dong", "dien to khai", "dien bieu mau", "dien giup", "dien cac", "dien het", "dien day du", "to khai", "bieu mau", "cho trong", "hoan thien mau", "mau co san", "ô điền"],
    skills: ["dien-mau"],
  },
  {
    id: "soat",
    label: "soát nhất quán, so phiên bản, phần ngoài thân văn bản",
    tools: ["word_check_consistency", "word_compare_versions", "word_read_parts"],
    keywords: ["soat", "kiem tra", "ra soat", "nhat quan", "truoc khi ky", "trinh ky", "ban cuoi", "so sanh", "phien ban", "doi chieu", "chu thich", "dau trang", "chan trang", "bookmark", "footnote", "header", "footer", "hop van ban", "loi chinh ta", "danh so"],
    skills: ["kiem-tra-truoc-ky", "ra-soat-hop-dong", "soan-thao-van-ban-hanh-chinh"],
  },
  {
    id: "nhan_xet_sua_doi",
    label: "nhận xét và sửa đổi theo dõi",
    tools: ["word_list_comments", "word_reply_comment", "word_resolve_comment", "word_list_revisions", "word_review_revision"],
    keywords: ["nhan xet", "comment", "gop y", "sua doi", "track changes", "redline", "doi tac", "chap nhan", "tu choi", "theo doi", "ai sua", "da sua gi"],
    skills: ["xu-ly-nhan-xet", "tom-tat-sua-doi-doi-tac", "kiem-tra-truoc-ky", "ra-soat-hop-dong"],
  },
  {
    id: "bang",
    label: "bảng",
    tools: ["word_read_tables", "word_insert_table", "word_fill_table"],
    keywords: ["bảng", "table", "cột", "kẻ bảng"],
    skills: ["kiem-tra-tuan-thu"],
  },
  {
    id: "ngoai",
    label: "ứng dụng Office khác",
    tools: ["read_other_app", "write_other_app"],
    keywords: ["excel", "powerpoint", "slide", "trinh chieu", "bang tinh", "ung dung khac"],
    skills: ["kiem-tra-tuan-thu", "chuan-bi-hop", "dien-mau"],
  },
  {
    id: "bo_nho",
    label: "ghi nhớ cho hội thoại sau",
    tools: ["propose_memory"],
    keywords: MEMORY_KEYWORDS,
  },
];

/** Đợt 11: nhóm chuyên của Excel. Lõi (luôn gửi): tổng quan, đọc, ghi, định dạng, thêm sheet, tạo bảng, sắp xếp, xóa vùng, chọn vùng,
 *  tìm, quét lỗi, lần công thức, nhận xét ô, chèn/xóa hàng cột, thao tác sheet, xem đối tượng có sẵn, kỹ năng, đọc trang tệp đính kèm */
export const EXCEL_TOOL_GROUPS: ToolGroup[] = [
  {
    id: "bieu_do",
    label: "biểu đồ",
    tools: ["excel_create_chart", "excel_edit_chart"],
    keywords: ["bieu do", "do thi", "chart", "ve bieu", "graph", "truc quan"],
    skills: ["thiet-ke-bieu-do", "phan-tich-du-lieu"],
  },
  {
    id: "pivot",
    label: "PivotTable",
    tools: ["excel_create_pivot", "excel_edit_pivot"],
    keywords: ["pivot", "pivottable", "pivot table", "tong hop theo", "bang tong hop", "thong ke theo", "bao cao tong hop", "nhom theo"],
    skills: ["phan-tich-du-lieu"],
  },
  {
    id: "loc_dinh_dang",
    label: "lọc, định dạng có điều kiện, danh sách chọn",
    tools: ["excel_filter", "excel_conditional_format", "excel_data_validation"],
    keywords: ["lọc", "filter", "dinh dang co dieu kien", "tô màu", "to mau theo", "highlight", "lam noi bat", "bat thuong", "danh sach chon", "tha xuong", "dropdown", "drop down", "validation", "conditional formatting", "chon tu danh sach", "gioi han nhap", "rang buoc nhap", "kiem soat nhap"],
    skills: ["phan-tich-du-lieu", "kiem-tra-tuan-thu"],
  },
  {
    id: "in_an",
    label: "thiết lập in",
    tools: ["excel_page_setup"],
    keywords: ["in ra", "ban in", "khi in", "de in", "vung in", "trang in", "print", "kho giay", "huong giay", "in ngang", "in doc", "vua trang", "vua mot trang", "vua 1 trang", "chuan bi in", "thiet lap in", "lap lai tieu de", "le trang"],
  },
  {
    id: "mo_hinh",
    label: "soát mô hình tài chính, tên ô giả định, giải ngược, bảng độ nhạy",
    tools: ["excel_audit_model", "excel_named_range", "excel_what_if"],
    keywords: ["mo hinh", "o gia dinh", "bang gia dinh", "cac gia dinh", "gia dinh dau vao", "do nhay", "goal seek", "tim nguoc", "giai nguoc", "what if", "dat ten vung", "dat ten o", "dat ten cho o", "ten vung", "named range", "npv", "irr", "dong tien", "hieu qua dau tu", "phuong an tai chinh", "so cung", "hard code", "kiem toan", "audit"],
    skills: ["mo-hinh-tai-chinh-du-an", "kiem-tra-loi-bang-tinh"],
  },
  {
    id: "lam_sach_khoa",
    label: "xóa trùng, thay hàng loạt, điền chuỗi, khóa ô công thức",
    tools: ["excel_clean_data", "excel_protect"],
    keywords: ["trùng", "lam sach", "don dep", "chuan hoa", "thay the", "thay hang loat", "dien chuoi", "dien tiep", "dien ngay", "keo xuong", "khóa", "khoá", "bao ve", "protect", "lock"],
    skills: ["chuan-hoa-bang-tinh"],
  },
  {
    id: "ngoai",
    label: "ứng dụng Office khác",
    tools: ["read_other_app", "write_other_app"],
    keywords: ["word", "powerpoint", "slide", "trinh chieu", "van ban dang mo", "ung dung khac"],
    skills: ["mo-hinh-tai-chinh-du-an", "kiem-tra-tuan-thu"],
  },
  {
    id: "bo_nho",
    label: "ghi nhớ cho hội thoại sau",
    tools: ["propose_memory"],
    keywords: MEMORY_KEYWORDS,
  },
];

/** Đợt 11: nhóm chuyên của PowerPoint. Lõi (luôn gửi): tổng quan, thêm slide (một và nhiều), sửa chữ, ô chữ, hình, bảng mới, chuyển slide,
 *  sửa / xóa hình, định dạng chữ, đọc vùng chọn, chụp slide, đọc mẫu, kỹ năng, đọc trang tệp đính kèm */
export const PPT_TOOL_GROUPS: ToolGroup[] = [
  {
    id: "bieu_do_so_do",
    label: "biểu đồ, sơ đồ quy trình và mốc thời gian",
    tools: ["ppt_add_chart", "ppt_add_process_flow"],
    keywords: ["bieu do", "do thi", "chart", "so lieu", "quy trinh", "cac buoc", "moc thoi gian", "timeline", "so do", "tien do", "lo trinh", "giai doan", "truc quan"],
    skills: ["thiet-ke-bieu-do", "bao-cao-tien-do-du-an", "trinh-chieu-tu-tai-lieu"],
  },
  {
    id: "sap_xep",
    label: "sắp xếp, nhân bản, xóa slide, đổi bố cục",
    tools: ["ppt_move_slide", "ppt_duplicate_slide", "ppt_delete_slide", "ppt_set_layout"],
    keywords: ["di chuyen", "doi cho", "doi thu tu", "sap xep", "len truoc", "len dau", "ra sau", "xuong cuoi", "hoan doi", "doi vi tri", "nhan ban", "sao chep slide", "xoa slide", "bo slide", "bo cuc", "layout"],
    skills: ["kiem-tra-bai-trinh-chieu"],
  },
  {
    id: "kiem_bai",
    label: "kiểm cả bài theo mẫu",
    tools: ["ppt_check_deck"],
    keywords: ["kiem tra", "soat", "ra soat", "hoan thien", "truoc khi gui", "dung mau", "phong chu", "co chu", "tran chu", "lỗi"],
    skills: ["kiem-tra-bai-trinh-chieu"],
  },
  {
    id: "hinh_anh",
    label: "chèn ảnh đính kèm",
    tools: ["ppt_insert_image"],
    keywords: ["ảnh", "hinh anh", "logo", "photo", "picture"],
  },
  {
    id: "bang",
    label: "bảng",
    tools: ["ppt_edit_table"],
    keywords: ["bảng", "table", "cột"],
  },
  {
    id: "ngoai",
    label: "ứng dụng Office khác",
    tools: ["read_other_app", "write_other_app"],
    keywords: ["word", "excel", "van ban dang mo", "so tinh", "bang tinh", "ung dung khac"],
    skills: ["bao-cao-tien-do-du-an", "trinh-chieu-tu-tai-lieu", "chuan-bi-hop"],
  },
  {
    id: "bo_nho",
    label: "ghi nhớ cho hội thoại sau",
    tools: ["propose_memory"],
    keywords: MEMORY_KEYWORDS,
  },
];

/** Dòng chỉ đường tới công cụ chuyên trong lời nhắc Word (chế độ Trả phí: đủ công cụ, gọi thẳng tên) */
export const WORD_SPECIAL_LINE =
  "- CÔNG CỤ CHUYÊN (làm theo mô tả của công cụ): soát trước khi ký, rà lỗi → word_check_consistency; đầu/chân trang, bookmark, trường, ô điền mẫu, hộp văn bản → word_read_parts; soạn văn bản hành chính → MỘT lần word_create_admin_document; lề, khổ giấy → word_check_page_format, word_set_page_format; so với tệp .docx đính kèm → word_compare_versions; điền mẫu → word_list_form_fields rồi MỘT lần word_fill_form.";

/** Chế độ Miễn phí: chỉ đường bằng TÊN NHÓM, không nêu tên công cụ có thể đang ẩn (soát mã Đợt 10, H2: model gọi tên công cụ
 *  không có trong danh sách gửi đi thì nhà cung cấp có thể từ chối cả yêu cầu) */
export const WORD_SPECIAL_LINE_FREE =
  "- CÔNG CỤ CHUYÊN (xem mục Nhóm công cụ chuyên): soát trước khi ký, rà lỗi, so với tệp .docx đính kèm, đầu/chân trang, bookmark, trường, hộp văn bản → nhóm soat; soạn văn bản hành chính (một lần gọi), lề, khổ giấy → nhóm hanh_chinh; điền mẫu → nhóm dien_mau; nhận xét, sửa đổi được theo dõi → nhóm nhan_xet_sua_doi; bảng → nhóm bang; ứng dụng Office khác → nhóm ngoai.";

/** Dòng chỉ đường tới công cụ chuyên trong lời nhắc Excel (Đợt 11, chế độ Trả phí) */
export const EXCEL_SPECIAL_LINE =
  "- CÔNG CỤ CHUYÊN (làm theo mô tả của công cụ): lọc → excel_filter; tô màu theo điều kiện → excel_conditional_format; danh sách chọn → excel_data_validation; PivotTable → excel_create_pivot; sửa biểu đồ, PivotTable CÓ SẴN → excel_list_objects rồi excel_edit_chart, excel_edit_pivot (không tạo mới); in → MỘT lần excel_page_setup; soát mô hình tài chính → excel_audit_model (đọc ô được chỉ rồi mới kết luận); tên ô giả định → excel_named_range; giải ngược, bảng độ nhạy → excel_what_if (sổ tính không đổi; muốn ghi kết quả thì hỏi người dùng rồi mới ghi); xóa dòng trùng, thay hàng loạt, điền chuỗi → excel_clean_data; khóa ô công thức → excel_protect (không đặt mật khẩu).";

export const EXCEL_SPECIAL_LINE_FREE =
  "- CÔNG CỤ CHUYÊN (xem mục Nhóm công cụ chuyên): biểu đồ → nhóm bieu_do; PivotTable → nhóm pivot; lọc, tô màu theo điều kiện, danh sách chọn → nhóm loc_dinh_dang; in → nhóm in_an; soát mô hình tài chính, tên ô giả định, giải ngược, bảng độ nhạy → nhóm mo_hinh; xóa dòng trùng, thay hàng loạt, điền chuỗi, khóa ô công thức → nhóm lam_sach_khoa; ứng dụng Office khác → nhóm ngoai. Sửa biểu đồ, PivotTable CÓ SẴN: excel_list_objects rồi công cụ của nhóm, không tạo mới.";

/** Dòng chỉ đường tới công cụ chuyên trong lời nhắc PowerPoint (Đợt 11, chế độ Trả phí) */
export const PPT_SPECIAL_LINE =
  "- CÔNG CỤ CHUYÊN: số liệu → ppt_add_chart (biểu đồ gốc theo màu chủ đề); quy trình, các bước, mốc thời gian → ppt_add_process_flow; sắp xếp slide → ppt_move_slide, ppt_duplicate_slide, ppt_delete_slide; đổi bố cục → ppt_set_layout; sửa bảng có sẵn → ppt_edit_table; ảnh người dùng đính kèm → ppt_insert_image; kiểm cả bài → ppt_check_deck.";

export const PPT_SPECIAL_LINE_FREE =
  "- CÔNG CỤ CHUYÊN (xem mục Nhóm công cụ chuyên): biểu đồ số liệu, quy trình, các bước, mốc thời gian → nhóm bieu_do_so_do; sắp xếp, nhân bản, xóa slide, đổi bố cục → nhóm sap_xep; sửa bảng có sẵn → nhóm bang; ảnh đính kèm → nhóm hinh_anh; kiểm cả bài → nhóm kiem_bai; ứng dụng Office khác → nhóm ngoai.";

/** Cấu hình nạp theo việc của từng ứng dụng: nhóm chuyên và cặp dòng chỉ đường (Trả phí theo tên công cụ, Miễn phí theo tên nhóm) */
const HOST_GROUPS: Record<string, { groups: ToolGroup[]; line: string; freeLine: string }> = {
  Word: { groups: WORD_TOOL_GROUPS, line: WORD_SPECIAL_LINE, freeLine: WORD_SPECIAL_LINE_FREE },
  Excel: { groups: EXCEL_TOOL_GROUPS, line: EXCEL_SPECIAL_LINE, freeLine: EXCEL_SPECIAL_LINE_FREE },
  PowerPoint: { groups: PPT_TOOL_GROUPS, line: PPT_SPECIAL_LINE, freeLine: PPT_SPECIAL_LINE_FREE },
};

/** Nhóm công cụ theo chế độ và ứng dụng: chỉ chế độ Miễn phí nạp theo việc; Trả phí luôn gửi đủ (giữ bộ nhớ đệm lời nhắc) */
export function toolGroupsFor(mode: string, host: string): ToolGroup[] | undefined {
  return mode === "free" ? HOST_GROUPS[host]?.groups : undefined;
}

/** Mọi nhóm của mọi ứng dụng (nhãn bước của load_tools: cùng id thì cùng nhãn) */
export const ALL_TOOL_GROUPS: ToolGroup[] = Object.values(HOST_GROUPS).flatMap((h) => h.groups);

/** Đổi tên công cụ chuyên trong một đoạn lời nhắc thành tên nhóm (chế độ Miễn phí: không nêu tên công cụ có thể đang ẩn) */
export function hideGroupTools(text: string, groups: ToolGroup[]): string {
  let out = text;
  for (const g of groups) for (const t of g.tools) out = out.split(t).join(`công cụ nhóm ${g.id}`);
  return out;
}

/** Lời nhắc riêng của ứng dụng cho chế độ Miễn phí: thay dòng chỉ đường theo tên công cụ bằng dòng theo tên nhóm, rồi đổi
 *  tên công cụ chuyên còn sót (vd word_list_revisions ở mục Track Changes) thành tên nhóm */
export function freeBasePrompt(base: string, groups: ToolGroup[] = WORD_TOOL_GROUPS): string {
  let out = base;
  for (const h of Object.values(HOST_GROUPS)) if (h.groups === groups) out = out.replace(h.line, h.freeLine);
  return hideGroupTools(out, groups);
}

/** Chữ thường, bỏ dấu tiếng Việt, gộp khoảng trắng */
export function fold(s: string): string {
  return String(s || "")
    .normalize("NFC")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");
}

const isAscii = (s: string) => /^[\x00-\x7f]*$/.test(s);
const escRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
/** Khớp trọn từ (không khớp chuỗi con: "số điện thoại" không chứa "ô điền", "lại sửa" không chứa "ai sửa") */
const hasWord = (text: string, kw: string) => new RegExp(`(^|[^\\p{L}\\p{N}])${escRe(kw)}($|[^\\p{L}\\p{N}])`, "u").test(text);

/** Đợt 12 E: câu người dùng gõ có nói tới việc ghi nhớ / quy ước lâu dài không (cùng từ khóa của nhóm bo_nho).
 *  Công cụ propose_memory chỉ nhận đề xuất khi có — chỉ dẫn ẩn trong tài liệu không tự sinh được đề xuất nhớ (soát mã phần E, #4). */
export function mentionsMemory(text: string): boolean {
  const raw = String(text || "").normalize("NFC").toLowerCase();
  const folded = fold(raw);
  return MEMORY_KEYWORDS.some((k) => (isAscii(k) ? hasWord(folded, k) : hasWord(raw, k.normalize("NFC").toLowerCase())));
}

/** Nhóm bật theo lời người dùng của lượt này, kỹ năng (tự chọn hoặc gọi bằng /tên) và gợi ý từ tệp đính kèm (id nhóm) */
export function requestGroups(groups: ToolGroup[], o: { text: string; skills?: string[]; hints?: string[] }): Set<string> {
  // NFC: chữ gõ kiểu tổ hợp (dấu tách rời) vẫn khớp từ khóa có dấu
  const raw = String(o.text || "").normalize("NFC").toLowerCase().replace(/\s+/g, " ");
  const folded = fold(o.text);
  const on = new Set<string>();
  for (const g of groups) {
    const hit = g.keywords.some((k) => (isAscii(k) ? hasWord(folded, k) : hasWord(raw, k.normalize("NFC").toLowerCase())));
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
    `Để tiết kiệm, chỉ công cụ cơ bản được gửi sẵn. CHỈ gọi công cụ có trong danh sách công cụ hiện có. Cần công cụ chuyên thì GỌI ${LOAD_TOOLS}(group) trước (có thể gọi cùng lúc với lệnh đọc); công cụ của nhóm có từ bước sau. Các nhóm:`,
    ...groups.map((g) => `- ${g.id}: ${g.label}`),
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
