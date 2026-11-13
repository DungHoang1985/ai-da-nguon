/** Chọn bộ công cụ theo ứng dụng đang chạy + mô tả từng bước bằng tiếng Việt cho khung chat. */
import { HostToolSet, ToolCall } from "../core/agent/types";
import { excelTools } from "./excelTools";
import { createWordTools } from "./wordTools";
import { pptTools } from "./pptTools";
import { HostName } from "./host";

export function getHostTools(host: HostName, settings: () => { trackChanges: boolean }): HostToolSet | null {
  if (host === "Excel") return excelTools;
  if (host === "Word") return createWordTools(settings);
  if (host === "PowerPoint") return pptTools;
  return null;
}

const where = (a: any) => {
  const addr = String(a?.address || a?.source_address || "");
  return (a?.sheet && !addr.includes("!") ? `${a.sheet}!` : "") + addr;
};

const CHART_VI: Record<string, string> = {
  column: "cột",
  bar: "thanh ngang",
  line: "đường",
  pie: "tròn",
  area: "vùng",
  scatter: "phân tán",
  doughnut: "vành khuyên",
  stacked_column: "cột chồng",
  stacked_bar: "thanh chồng",
};

export function describeCall(call: ToolCall): string {
  const a = call.args || {};
  switch (call.name) {
    case "excel_get_overview":
      return "Xem tổng quan sổ tính";
    case "excel_read_range":
      return `Đọc ${where(a)}`;
    case "excel_write_range":
      return `Ghi dữ liệu vào ${where(a)}`;
    case "excel_format_range":
      return `Định dạng ${where(a)}`;
    case "excel_add_sheet":
      return `Tạo sheet "${a.name}"`;
    case "excel_create_table":
      return `Tạo bảng ${where(a)}`;
    case "excel_create_chart":
      return `Vẽ biểu đồ ${CHART_VI[a.chart_type] || a.chart_type || ""} từ ${where(a)}`;
    case "excel_sort_range":
      return `Sắp xếp ${where(a)} theo cột ${a.column}`;
    case "excel_clear_range":
      return `Xóa ${where(a)}`;
    case "excel_select_range":
      return `Chọn ${where(a)}`;
    case "word_get_overview":
      return "Xem tổng quan văn bản";
    case "word_read_paragraphs":
      return a.selection_only ? "Đọc đoạn đang chọn" : `Đọc đoạn ${a.start ?? 0}–${a.end ?? ""}`;
    case "word_replace_text":
      return `Thay "${String(a.find || "").slice(0, 40)}" → "${String(a.replace || "").slice(0, 40)}"`;
    case "word_set_paragraph_text":
      return `Sửa đoạn ${a.index}`;
    case "word_insert_paragraphs":
      return `Chèn nội dung (${a.position || "cuối văn bản"})`;
    case "word_replace_selection":
      return "Thay nội dung đang chọn";
    case "word_format_text":
      return "Định dạng chữ";
    case "word_add_comment":
      return `Thêm nhận xét: ${String(a.text || "").slice(0, 50)}`;
    case "word_insert_table":
      return "Chèn bảng";
    case "word_delete_paragraphs":
      return `Xóa đoạn ${a.start}${a.end !== undefined ? `–${a.end}` : ""}`;
    case "ppt_get_overview":
      return "Xem tổng quan bài trình chiếu";
    case "ppt_add_slide":
      return `Thêm slide "${String(a.title || a.layout || "").slice(0, 50)}"`;
    case "ppt_set_shape_text":
      return `Sửa chữ slide ${a.slide} (${a.shape})`;
    case "ppt_add_textbox":
      return `Thêm ô chữ ở slide ${a.slide}`;
    case "ppt_add_shape":
      return `Thêm hình ở slide ${a.slide}`;
    case "ppt_add_table":
      return `Thêm bảng ở slide ${a.slide}`;
    case "ppt_go_to_slide":
      return `Chuyển tới slide ${a.slide}`;
    // Excel nâng cao
    case "excel_find":
      return `Tìm "${String(a.text || "").slice(0, 40)}"${a.sheet ? ` ở sheet ${a.sheet}` : ""}`;
    case "excel_find_errors":
      return `Quét lỗi công thức${a.sheet ? ` ở sheet ${a.sheet}` : " cả sổ tính"}`;
    case "excel_trace":
      return a.direction === "dependents" ? `Lần các ô phụ thuộc vào ${a.cell}` : `Lần các ô mà ${a.cell} tham chiếu tới`;
    case "excel_filter":
      return a.clear ? "Bỏ lọc" : `Lọc ${where(a)}${a.column ? ` theo cột ${a.column}` : ""}`;
    case "excel_create_pivot":
      return `Tạo PivotTable từ ${where(a)}`;
    case "excel_conditional_format":
      return `Định dạng có điều kiện ${where(a)}`;
    case "excel_data_validation":
      return `Tạo danh sách chọn ${where(a)}`;
    case "excel_add_comment":
      return `Thêm ghi chú ô ${a.cell}`;
    case "excel_insert_delete": {
      const del = String(a.action).startsWith("delete");
      const rows = String(a.action).includes("rows");
      return `${del ? "Xóa" : "Chèn"} ${a.count || 1} ${rows ? "hàng" : "cột"} tại ${a.at}${a.sheet ? ` (${a.sheet})` : ""}`;
    }
    case "excel_sheet":
      return a.action === "rename"
        ? `Đổi tên sheet "${a.sheet || "đang mở"}" → "${a.new_name}"`
        : a.action === "freeze_rows"
          ? `Cố định ${a.count || 1} dòng đầu`
          : a.action === "freeze_columns"
            ? `Cố định ${a.count || 1} cột đầu`
            : a.action === "unfreeze"
              ? "Bỏ cố định dòng/cột"
              : a.action === "delete"
                ? `Xóa sheet "${a.sheet || "đang mở"}"`
                : `Mở sheet "${a.sheet}"`;
    // Word nâng cao
    case "word_list_comments":
      return "Đọc các nhận xét";
    case "word_reply_comment":
      return `Trả lời nhận xét: ${String(a.text || "").slice(0, 50)}`;
    case "word_resolve_comment":
      return a.resolved === false ? "Mở lại nhận xét" : "Đánh dấu nhận xét đã xử lý";
    case "word_list_revisions":
      return "Đọc các sửa đổi được theo dõi";
    case "word_review_revision":
      return `${a.action === "accept" ? "Chấp nhận" : "Từ chối"} ${a.all ? "tất cả sửa đổi" : `sửa đổi số ${(a.indexes || []).join(", ")}`}`;
    case "word_read_tables":
      return "Đọc các bảng";
    case "word_fill_table":
      return `Điền bảng số ${a.table_index}`;
    case "word_go_to":
      return `Chuyển tới [¶${a.paragraph_index}]`;
    // PowerPoint nâng cao
    case "ppt_add_chart":
      return `Chèn biểu đồ ${CHART_VI[a.chart_type] || a.chart_type || ""} (slide mới)`;
    case "ppt_add_process_flow":
      return `Vẽ sơ đồ ${Array.isArray(a.steps) ? a.steps.length + " bước " : ""}ở slide ${a.slide}`;
    case "ppt_move_slide":
      return `Chuyển slide ${a.slide} tới vị trí ${a.to}`;
    case "ppt_duplicate_slide":
      return `Nhân bản slide ${a.slide}`;
    case "ppt_delete_slide":
      return `Xóa slide ${a.slide}`;
    case "ppt_update_shape":
      return `Sửa hình "${a.shape}" ở slide ${a.slide}`;
    case "ppt_delete_shape":
      return `Xóa hình "${a.shape}" ở slide ${a.slide}`;
    // Dùng chung
    case "use_skill":
      return `Đọc kỹ năng ${a.name}${a.file ? ` (${a.file})` : ""}`;
    case "read_other_app":
      return `Đọc nội dung đang mở ở ${a.app}`;
  }
  return call.name;
}

export const QUICK_PROMPTS: Record<string, string[]> = {
  Excel: [
    "Quét lỗi công thức trong sổ tính và sửa",
    "Tạo PivotTable tổng hợp theo nhóm rồi vẽ biểu đồ",
    "Tô màu các giá trị vượt ngưỡng bằng định dạng có điều kiện",
    "Chuẩn hóa bảng: tiêu đề, định dạng số, cố định dòng tiêu đề",
  ],
  Word: [
    "Rà soát hợp đồng này cho bên mình và ghi nhận xét vào các điều khoản rủi ro",
    "Tóm tắt các sửa đổi của đối tác (Track Changes) và đề xuất chấp nhận hay từ chối",
    "Xử lý các nhận xét trong văn bản",
    "Kiểm tra văn bản trước khi trình ký",
  ],
  PowerPoint: [
    "Tạo 5 slide giới thiệu về chủ đề tôi đưa",
    "Thêm slide sơ đồ quy trình các bước từ nội dung tôi đưa",
    "Thêm biểu đồ từ số liệu đang mở ở Excel",
    "Kiểm tra bài trình chiếu và sửa lỗi",
  ],
};
