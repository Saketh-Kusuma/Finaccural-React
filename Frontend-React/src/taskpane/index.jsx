import { createRoot } from "react-dom/client";
import { App } from "./App";
const mount = () => createRoot(document.getElementById("root")).render(<App />);

if (window.Office?.onReady) {
  window.Office.onReady(mount);
} else {
  mount();
}
