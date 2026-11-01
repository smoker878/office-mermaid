/* global document, window, Office, HTMLTextAreaElement, HTMLImageElement, setTimeout, clearTimeout */
import { DiagramError, renderDiagram } from "../renderer/renderer";

const defaultMermaid = `flowchart LR
  A[輸入 Mermaid] --> B[調整主題]
  B --> C[預覽 SVG]
  C --> D[準備插入 Office]`;
const defaultCss = `.pele {
  --pele-bg: #ffffff;
  --pele-fg: #243247;
  --pele-line: #526581;
  --pele-surface: #edf3fc;
  --pele-border: #8da9d2;
  --pele-accent: #2563eb;
  --pele-font: sans-serif;
  --pele-radius: 4px;
}`;

function initialize() {
  const mermaid = document.getElementById("mermaid-source") as HTMLTextAreaElement;
  const css = document.getElementById("css-source") as HTMLTextAreaElement;
  const status = document.getElementById("render-status")!;
  const errorBox = document.getElementById("render-error")!;
  const image = document.getElementById("svg-preview") as HTMLImageElement;
  const placeholder = document.getElementById("preview-placeholder")!;
  const dimensions = document.getElementById("preview-dimensions")!;
  const preview = document.getElementById("preview-region")!;
  let timer: ReturnType<typeof setTimeout>;
  let hasPreview = false;
  let dirty = false;
  mermaid.value = defaultMermaid;
  css.value = defaultCss;

  function refresh() {
    clearTimeout(timer);
    mermaid.removeAttribute("aria-invalid");
    css.removeAttribute("aria-invalid");
    errorBox.hidden = true;
    try {
      const result = renderDiagram(mermaid.value, css.value);
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(result.svg)}`;
      image.hidden = false;
      placeholder.hidden = true;
      hasPreview = true;
      dimensions.textContent = `${result.type} · ${Math.round(result.width)} × ${Math.round(result.height)} px`;
      status.textContent = "預覽已更新";
      status.dataset.state = "success";
    } catch (error) {
      const field = error instanceof DiagramError && error.field === "css" ? css : mermaid;
      field.setAttribute("aria-invalid", "true");
      errorBox.textContent =
        error instanceof Error ? error.message : "渲染失敗，請修改輸入後重試。";
      errorBox.hidden = false;
      status.textContent = hasPreview ? "輸入有誤 · 保留上次成功預覽" : "尚無有效預覽";
      status.dataset.state = "error";
    } finally {
      preview.setAttribute("aria-busy", "false");
    }
  }

  function scheduleRefresh() {
    dirty = true;
    clearTimeout(timer);
    status.textContent = "等待更新預覽…";
    status.dataset.state = "pending";
    preview.setAttribute("aria-busy", "true");
    timer = setTimeout(refresh, 400);
  }
  mermaid.addEventListener("input", scheduleRefresh);
  css.addEventListener("input", scheduleRefresh);
  document.getElementById("refresh-preview")!.addEventListener("click", refresh);
  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
      event.preventDefault();
      refresh();
    }
  });
  window.addEventListener("beforeunload", (event) => {
    if (dirty) {
      event.preventDefault();
      event.returnValue = "";
    }
  });
  refresh();
  initializeOffice();
}

function initializeOffice() {
  // preview.html never loads Office.js: it also works with the network disabled.
  if (window.location.pathname.endsWith("preview.html")) return;
  const label = document.getElementById("host-status")!;
  const script = document.createElement("script");
  script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
  script.async = true;
  script.onerror = () => {
    label.textContent = "Office.js 無法連線；編輯器與預覽仍可使用。";
  };
  script.onload = () => {
    if (typeof Office !== "undefined") {
      Office.onReady((info) => {
        label.textContent = info.host
          ? `${info.host} 已連線；文件插入與重新編輯尚未整合。`
          : "瀏覽器預覽模式；文件操作需要 Office 宿主。";
      });
    }
  };
  document.head.appendChild(script);
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initialize);
} else {
  initialize();
}
