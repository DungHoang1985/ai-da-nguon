// B15 (rà soát kiến trúc 20/11/2026, làm 26/11): lint riêng của dự án thay cấu hình mặc định của office-addin-lint (cấu hình đó
// báo 4.536 lỗi, phần lớn là định dạng prettier / CRLF và biến trình duyệt "chưa khai báo" — tiếng ồn che lỗi thật).
// Giữ: quy tắc lỗi của eslint-plugin-office-addins (load / sync Office.js) và react-hooks/rules-of-hooks — 26/11/2026 suýt đặt
// useCallback sau return sớm trong Chat.tsx (khung vỡ khi chuyển từ màn chọn chế độ sang chat). Kiểu và biến chưa khai báo: tsc lo.
import officeAddins from "eslint-plugin-office-addins";
import reactHooks from "eslint-plugin-react-hooks";
import tsParser from "@typescript-eslint/parser";
import tsPlugin from "@typescript-eslint/eslint-plugin";

const recommended = officeAddins.configs.recommended;

const ALL = ["src/**/*.{ts,tsx}", "test/**/*.ts"];
const OFFICE_RULES = Object.fromEntries(
  recommended.flatMap((c) => Object.entries(c.rules || {})).filter(([k]) => k.startsWith("office-addins/"))
);
const CORE_RULES = Object.fromEntries(
  recommended.flatMap((c) => Object.entries(c.rules || {})).filter(([k]) => !k.includes("/"))
);

export default [
  { ignores: ["dist/**", "node_modules/**", "kiem-thu/**", "tasks/**", "_luu-cuc-bo/**", ".claude/**", "*.config.js", "*.mjs"] },
  {
    files: ALL,
    // Plugin @typescript-eslint chỉ đăng ký (không bật quy tắc) để các chú thích eslint-disable cũ trong mã nhận ra tên quy tắc
    plugins: { "@typescript-eslint": tsPlugin },
    languageOptions: { parser: tsParser, parserOptions: { ecmaFeatures: { jsx: true } } },
    linterOptions: { reportUnusedDisableDirectives: "off" },
    rules: {
      ...CORE_RULES,
      // tsc đã kiểm (và hiểu kiểu, overload, biến toàn cục của Office / DOM); để eslint báo nữa chỉ ra lỗi giả
      "no-undef": "off",
      "no-unused-vars": "off",
      "no-redeclare": "off",
      "no-dupe-class-members": "off",
      // Chữ tiếng Việt, khoảng trắng không ngắt và ký tự điều khiển trong biểu thức chính quy là CỐ Ý (làm sạch văn bản Office)
      "no-irregular-whitespace": "off",
      "no-control-regex": "off",
      "no-misleading-character-class": "off",
      "no-useless-escape": "off",
    },
  },
  // Quy tắc Office.js (load / sync) chỉ có nghĩa với mã gọi Office: ở lõi và giao diện, biến tên "sel", "range"… là đối tượng thường.
  // Mức CẢNH BÁO: quy tắc đoán theo cú pháp, không lần qua hàm phụ (loadRows, snapshot…) nên báo nhầm nhiều (285 chỗ ngày 26/11/2026,
  // trong mã đã qua Tự kiểm tra); xem khi sửa đúng tệp đó, không chặn.
  {
    files: ["src/office/**/*.ts"],
    plugins: { "office-addins": officeAddins },
    rules: Object.fromEntries(Object.keys(OFFICE_RULES).map((k) => [k, "warn"])),
  },
  // Hook React chỉ ở giao diện: lõi có hàm thường tên useText, useSkill… (không phải hook)
  {
    files: ["src/taskpane/**/*.{ts,tsx}"],
    plugins: { "react-hooks": reactHooks },
    rules: { "react-hooks/rules-of-hooks": "error" },
  },
];
