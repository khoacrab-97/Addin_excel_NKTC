import { initializeIcons } from "@fluentui/react";
import ReactDOM from "react-dom";
import { App } from "./App";

Office.onReady(() => {
  initializeIcons();
  // React 17: dùng ReactDOM.render để tương thích Internet Explorer 11 (Excel 2016/2019).
  ReactDOM.render(<App />, document.getElementById("root"));
});
