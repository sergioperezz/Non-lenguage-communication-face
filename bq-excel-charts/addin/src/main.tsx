import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";

// Office.onReady garantiza que el host (Excel) está listo antes de montar la UI.
// En un navegador normal (fuera de Office) seguimos montando para poder
// desarrollar la UI, aunque las acciones de Excel no estarán disponibles.
function mount() {
  const root = createRoot(document.getElementById("root")!);
  root.render(<React.StrictMode><App /></React.StrictMode>);
}

if (typeof Office !== "undefined") {
  Office.onReady().then(mount);
} else {
  mount();
}
