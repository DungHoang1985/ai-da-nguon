import * as React from "react";
import { Tab, TabList, Text, makeStyles, tokens } from "@fluentui/react-components";
import Chat from "./Chat";
import Settings from "./Settings";
import { Storage } from "../../core/storage";
import { hostName } from "../../office/host";

const useStyles = makeStyles({
  root: { padding: "10px", height: "100vh", boxSizing: "border-box", display: "flex", flexDirection: "column" },
  head: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  body: { flex: 1, overflowY: "auto", paddingTop: "8px" },
  muted: { color: tokens.colorNeutralForeground3 },
});

const storage = new Storage();

const App: React.FC = () => {
  const s = useStyles();
  const host = hostName();
  const hasKey = Object.keys(storage.getKeys()).length > 0;
  const [tab, setTab] = React.useState<string>(hasKey ? "chat" : "settings");
  const [version, setVersion] = React.useState(0);

  return (
    <div className={s.root}>
      <div className={s.head}>
        <Text weight="semibold" size={400}>
          AI Đa Nguồn
        </Text>
        <Text size={200} className={s.muted}>
          {host === "Khác" ? "" : `Đang chạy trong ${host}`}
        </Text>
      </div>
      <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(String(d.value))} size="small">
        <Tab value="chat">Chat</Tab>
        <Tab value="settings">Cài đặt</Tab>
      </TabList>
      <div className={s.body}>
        {tab === "chat" ? (
          <Chat storage={storage} host={host} version={version} goSettings={() => setTab("settings")} />
        ) : (
          <Settings storage={storage} onChanged={() => setVersion((v) => v + 1)} />
        )}
      </div>
    </div>
  );
};

export default App;
