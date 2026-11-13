/* global Office */
import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./components/App";

const render = () => {
  const root = createRoot(document.getElementById("container") as HTMLElement);
  root.render(
    <FluentProvider theme={webLightTheme} style={{ height: "100%" }}>
      <App />
    </FluentProvider>
  );
};

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => render());
} else {
  render();
}
