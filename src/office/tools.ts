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
  }
  return call.name;
}

export const QUICK_PROMPTS: Record<string, string[]> = {
  Excel: [
    "Tóm tắt dữ liệu ở sheet đang mở",
    "Tạo bảng tổng hợp theo nhóm và vẽ biểu đồ",
    "Kiểm tra công thức lỗi và sửa",
    "Định dạng bảng cho đẹp (tiêu đề, viền, số có dấu phân cách)",
  ],
  Word: [
    "Rà soát chính tả, ngữ pháp và sửa bằng Track Changes",
    "Thêm nhận xét cho các điều khoản có rủi ro",
    "Tóm tắt văn bản này và chèn phần tóm tắt lên đầu",
    "Chuẩn hóa định dạng tiêu đề các mục",
  ],
  PowerPoint: [
    "Tạo 5 slide giới thiệu về chủ đề tôi đưa",
    "Tóm tắt nội dung bài trình chiếu",
    "Viết lại tiêu đề các slide cho gọn và rõ",
    "Thêm slide kết luận ở cuối",
  ],
};
