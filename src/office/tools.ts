/** Chọn bộ công cụ theo ứng dụng đang chạy + mô tả từng bước bằng tiếng Việt cho khung chat. */
import { HostToolSet, ToolCall } from "../core/agent/types";
import { excelTools } from "./excelTools";
import { createWordTools, reasonOf } from "./wordTools";
import { pptTools } from "./pptTools";
import { HostName } from "./host";
import { withAuditTrail } from "./excelTrail";

export function getHostTools(host: HostName, settings: () => unknown): HostToolSet | null {
  if (host === "Excel") return withAuditTrail(excelTools, { settings: () => settings() as any, describe: describeCall });
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

/** Nhãn của một bước; công cụ sửa Word kèm lý do AI nêu */
export function describeCall(call: ToolCall): string {
  const why = reasonOf(call.name, call.args);
  return describeBase(call) + (why ? ` — lý do: ${why}` : "");
}

function describeBase(call: ToolCall): string {
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
    case "word_get_outline":
      return "Xem mục lục điều khoản";
    case "word_find":
      return `Tìm "${String(a.query || "").slice(0, 40)}" trong văn bản`;
    case "word_read_clause":
      return `Đọc ${String(a.number || "").slice(0, 40)}`;
    case "word_read_paragraphs":
      return a.selection_only ? "Đọc đoạn đang chọn" : `Đọc đoạn ${a.start ?? 0}–${a.end ?? ""}${a.view === "original" ? " (bản gốc)" : ""}`;
    case "word_replace_text":
      return (
        `Thay "${String(a.find || "").slice(0, 40)}" → "${String(a.replace || "").slice(0, 40)}"` +
        (a.paragraph_index !== undefined && a.paragraph_index !== null ? ` ở [¶${a.paragraph_index}]` : a.all ? " (mọi chỗ)" : "")
      );
    case "word_set_paragraph_text":
      return `Sửa [¶${a.index}]`;
    case "word_insert_clause":
      return `Chèn điều khoản mới sau [¶${a.after_index}]`;
    case "word_change_level":
      return `Đổi cấp [¶${a.index}] (${a.direction === "deeper" ? "xuống" : "lên"} một cấp)`;
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
    case "ppt_add_slides": {
      const n = Array.isArray(a.slides) ? a.slides.length : 0;
      const first = Array.isArray(a.slides) ? String(a.slides[0]?.title || "").slice(0, 40) : "";
      return `Thêm ${n || "nhiều"} slide${first ? ` (từ "${first}")` : ""}`;
    }
    case "ppt_insert_image":
      return `Đặt ảnh đính kèm${a.image ? ` "${String(a.image).slice(0, 40)}"` : ""} vào slide ${a.slide}`;
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
      if (/group_/.test(String(a.action))) return `${String(a.action).startsWith("un") ? "Bỏ nhóm" : "Nhóm"} ${a.count || 1} ${rows ? "hàng" : "cột"} từ ${a.at}${a.sheet ? ` (${a.sheet})` : ""}`;
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
    // Excel Đợt 4
    case "excel_list_objects":
      return `Xem biểu đồ, PivotTable, tên vùng, thiết lập in đang có${a.sheet ? ` ở sheet ${a.sheet}` : ""}`;
    case "excel_edit_chart":
      return a.delete ? `Xóa biểu đồ ${a.chart || ""}` : `Sửa biểu đồ ${a.chart || "có sẵn"}${a.chart_type ? `: kiểu ${CHART_VI[a.chart_type] || a.chart_type}` : ""}${a.title ? `, tiêu đề “${String(a.title).slice(0, 40)}”` : ""}`;
    case "excel_edit_pivot":
      return `Sửa PivotTable ${a.pivot || "có sẵn"}${a.refresh ? " (làm mới)" : ""}`;
    case "excel_page_setup":
      return `Thiết lập in${a.sheet ? ` sheet ${a.sheet}` : ""}${a.print_area ? `: vùng in ${a.print_area}` : ""}${a.orientation ? `, giấy ${/^land/i.test(a.orientation) ? "ngang" : "dọc"}` : ""}`;
    case "excel_named_range":
      return a.action === "delete" ? `Xóa tên vùng ${a.name}` : `Đặt tên vùng ${a.name} cho ${a.address || ""}`;
    case "excel_audit_model":
      return `Kiểm toán mô hình${a.address ? ` ${where(a)}` : a.sheet ? ` sheet ${a.sheet}` : ""}: số gõ cứng, công thức lệch, liên kết ngoài`;
    case "excel_what_if":
      return a.mode === "sensitivity" ? `Bảng nhạy cảm của ${a.output_cell || ""} theo ${a.input_cell}${a.input_cell_2 ? ` và ${a.input_cell_2}` : ""}` : `Giải ngược: tìm ${a.input_cell} để ${a.target_cell || ""} = ${a.target_value ?? ""}`;
    case "excel_clean_data":
      return a.action === "remove_duplicates" ? `Xóa dòng trùng ở ${where(a)}` : a.action === "replace_all" ? `Thay “${String(a.find || "").slice(0, 30)}” → “${String(a.replace || "").slice(0, 30)}” ở ${where(a)}` : `Tự điền ${a.fill_to || where(a)}`;
    case "excel_protect":
      return a.action === "unprotect" ? `Bỏ bảo vệ sheet ${a.sheet || "đang mở"}` : `Khóa ô công thức, bảo vệ sheet ${a.sheet || "đang mở"}`;
    // Word nâng cao
    case "word_list_comments":
      return "Đọc các nhận xét";
    case "word_reply_comment":
      return `Trả lời nhận xét: ${String(a.text || "").slice(0, 50)}`;
    case "word_resolve_comment":
      return a.resolved === false ? "Mở lại nhận xét" : "Đánh dấu nhận xét đã xử lý";
    case "word_list_revisions":
      return `Đọc các sửa đổi được theo dõi${a.author ? ` của ${a.author}` : ""}${a.start ? ` (từ số ${a.start})` : ""}`;
    case "word_review_revision":
      return `${a.action === "accept" ? "Chấp nhận" : "Từ chối"} ${a.all ? `tất cả sửa đổi${a.author ? ` của ${a.author}` : ""}` : `sửa đổi số ${(a.indexes || []).join(", ")}`}`;
    case "word_read_tables":
      return "Đọc các bảng";
    case "word_fill_table":
      return `Điền bảng số ${a.table_index}`;
    case "word_go_to":
      return `Chuyển tới [¶${a.paragraph_index}]`;
    case "word_check_consistency":
      return "Soát nhất quán: đánh số, dẫn chiếu, thuật ngữ, tên các bên, chỗ trống, số và chữ";
    case "word_read_parts":
      return "Đọc chú thích, đầu và chân trang, bookmark, trường, ô điền mẫu, hộp văn bản";
    case "word_check_page_format":
      return "Kiểm thể thức trang: lề, khổ giấy, phông, cỡ chữ";
    case "ppt_get_slide_image":
      return `Chụp ảnh slide ${a.slide ?? ""} để tự kiểm`;
    case "ppt_check_deck":
      return "Kiểm bài theo mẫu: phông, cỡ chữ, tiêu đề, chữ tràn, hình ra ngoài";
    case "ppt_get_theme":
      return "Đọc mẫu: màu chủ đề, phông, bố cục, ô giữ chỗ";
    case "ppt_format_text":
      return `Định dạng chữ ở slide ${a.slide ?? ""}${a.paragraph ? `, đoạn ${a.paragraph}` : a.find ? `: “${String(a.find).slice(0, 30)}”` : ""}`;
    case "ppt_edit_table":
      return `Sửa bảng ở slide ${a.slide ?? ""}`;
    case "ppt_get_selection":
      return "Đọc hình và chữ đang chọn";
    case "ppt_set_layout":
      return `Đổi bố cục slide ${a.slide ?? ""} sang “${String(a.layout || "").slice(0, 30)}”`;
    case "write_other_app":
      return `Ghi sang ${a.app || "ứng dụng khác"}${a.title ? `: slide “${String(a.title).slice(0, 40)}”` : a.sheet_name ? `: sheet “${String(a.sheet_name).slice(0, 30)}”` : ""}`;
    case "word_create_admin_document":
      return `Dựng văn bản hành chính đúng thể thức${a.subject ? `: ${String(a.subject).slice(0, 50)}` : ""}`;
    case "word_set_page_format":
      return "Đặt khổ giấy và lề theo chuẩn";
    case "word_compare_versions":
      return `So với phiên bản khác${a.file ? ` (${String(a.file).slice(0, 40)})` : ""}${a.open_redline ? ", mở bản so sánh của Word" : ""}`;
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
    case "read_attachment_pages":
      return `Đọc trang ${a.pages ?? "?"} của tệp đính kèm${a.file ? ` ${a.file}` : ""}`;
  }
  return call.name;
}

export const QUICK_PROMPTS: Record<string, string[]> = {
  Excel: [
    "Quét lỗi công thức trong sổ tính và sửa",
    "Tạo PivotTable tổng hợp theo nhóm rồi vẽ biểu đồ",
    "Tô màu các giá trị vượt ngưỡng bằng định dạng có điều kiện",
    "Chuẩn hóa bảng: tiêu đề, định dạng số, cố định dòng tiêu đề",
    "Kiểm toán mô hình: tìm số gõ cứng và công thức lệch hàng",
    "Chuẩn bị in sheet này: A4 ngang, vừa một trang ngang, lặp dòng tiêu đề",
  ],
  Word: [
    "Rà soát hợp đồng này cho bên mình và ghi nhận xét vào các điều khoản rủi ro",
    "Tóm tắt các sửa đổi của đối tác (Track Changes) và đề xuất chấp nhận hay từ chối",
    "Xử lý các nhận xét trong văn bản",
    "Soát nhất quán văn bản trước khi trình ký: đánh số, dẫn chiếu, thuật ngữ, chỗ trống, số và chữ",
    "Soạn công văn đúng thể thức theo nội dung tôi đưa (hỏi tôi cơ quan gửi, nơi nhận và nội dung chính)",
  ],
  PowerPoint: [
    "Tạo 5 slide giới thiệu về chủ đề tôi đưa",
    "Thêm slide sơ đồ quy trình các bước từ nội dung tôi đưa",
    "Thêm biểu đồ từ số liệu đang mở ở Excel",
    "Kiểm tra bài trình chiếu và sửa lỗi",
  ],
};
