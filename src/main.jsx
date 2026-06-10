import React from "react";
import { createRoot } from "react-dom/client";
import FleetDataAnalyzer from "./fleet-data-analyzer.jsx";

const style = document.createElement("style");
style.textContent = "html,body,#root{margin:0;height:100%;background:#eef1f4;}";
document.head.appendChild(style);

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <FleetDataAnalyzer />
  </React.StrictMode>
);
