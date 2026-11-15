import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import Chat from "./Chat";
import Settings from "./Settings";
import { Storage } from "../../core/storage";
import { hostName } from "../../office/host";

/** Bố cục "một dòng đầu" (anh chọn 15/11/2026): không còn thanh thẻ Chat / Cài đặt.
 *  Khung chat có đầu khung một dòng; Cài đặt là trang riêng mở bằng nút bánh răng, có nút quay lại.
 *  Khung chat vẫn được giữ nguyên trạng thái khi mở Cài đặt (ẩn bằng display: none). */
const useStyles = makeStyles({
  root: {
    padding: "6px 8px 8px",
    height: "100vh",
    width: "100%",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    overflowX: "hidden",
  },
  page: { flex: 1, minHeight: 0, display: "flex", flexDirection: "column" },
});

const storage = new Storage();

const App: React.FC = () => {
  const s = useStyles();
  const host = hostName();
  const [view, setView] = React.useState<"chat" | "settings">("chat");
  const [version, setVersion] = React.useState(0);

  return (
    <div className={s.root}>
      <div className={s.page} style={{ display: view === "chat" ? "flex" : "none" }}>
        <Chat storage={storage} host={host} version={version} onOpenSettings={() => setView("settings")} />
      </div>
      {view === "settings" && (
        <div className={s.page}>
          <Settings storage={storage} host={host} onChanged={() => setVersion((v) => v + 1)} onBack={() => setView("chat")} />
        </div>
      )}
    </div>
  );
};

export default App;
