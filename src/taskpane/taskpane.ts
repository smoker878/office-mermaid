/* global document, window, Office, HTMLTextAreaElement, HTMLImageElement, HTMLInputElement, HTMLButtonElement, setTimeout, clearTimeout */
import { DiagramError, renderDiagram } from "../renderer/renderer";
import { defaultLayout, validateDraft, type DiagramDraft } from "../diagram/record";
import { WordAdapter, wordCapability } from "../office/word-adapter";
import type { DiagramTarget, OfficeAdapter } from "../office/adapter";
import { DiagramTargetError } from "../office/adapter";

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
  const layoutWidth = document.getElementById("layout-width") as HTMLInputElement;
  const slider = document.getElementById("layout-width-slider") as HTMLInputElement;
  const autoDirection = document.getElementById("auto-direction") as HTMLInputElement;
  const displayWidth = document.getElementById("display-width") as HTMLInputElement;
  const preserveWidth = document.getElementById("preserve-width") as HTMLInputElement;
  const insert = document.getElementById("insert-diagram") as HTMLButtonElement;
  const load = document.getElementById("load-diagram") as HTMLButtonElement;
  const update = document.getElementById("update-diagram") as HTMLButtonElement;
  const newDiagram = document.getElementById("new-diagram") as HTMLButtonElement;
  const officeStatus = document.getElementById("office-status")!;
  const officeError = document.getElementById("office-error")!;
  const editingTarget = document.getElementById("editing-target")!;
  let timer: ReturnType<typeof setTimeout>;
  let hasPreview = false;
  let adapter: OfficeAdapter | null = null;
  let target: DiagramTarget | null = null;
  let busy = false;
  let needsReload = false;
  let validPreview = false;
  let baseline: string;
  mermaid.value = defaultMermaid;
  css.value = defaultCss;
  layoutWidth.value = String(defaultLayout.maxWidth);

  function draft(): DiagramDraft {
    return {
      mermaidSource: mermaid.value,
      cssSource: css.value,
      layout: { maxWidth: layoutWidth.valueAsNumber, autoDirection: autoDirection.checked },
      displayWidthCm: displayWidth.valueAsNumber,
    };
  }

  function fingerprint() {
    return JSON.stringify([
      mermaid.value,
      css.value,
      layoutWidth.value,
      autoDirection.checked,
      displayWidth.value,
    ]);
  }
  baseline = fingerprint();

  function buttons() {
    insert.disabled = !adapter || busy || !validPreview || !!target;
    load.disabled = !adapter || busy;
    update.disabled = !adapter || busy || !validPreview || !target || needsReload;
    newDiagram.hidden = !target;
    newDiagram.disabled = busy;
    insert.classList.toggle("primary", !target);
    update.classList.toggle("primary", !!target);
    editingTarget.textContent = target
      ? `正在編輯圖表 · ${target.record.id.slice(0, 8)}`
      : "新增圖表";
    if (needsReload) editingTarget.textContent += " · 請重新載入確認（草稿保留中）";
    for (const element of [
      mermaid,
      css,
      layoutWidth,
      slider,
      autoDirection,
      displayWidth,
      preserveWidth,
    ])
      element.disabled = busy;
    displayWidth.disabled = busy || (!!target && preserveWidth.checked);
    document.getElementById("refresh-preview")!.toggleAttribute("disabled", busy);
    document.querySelector(".office-panel")!.setAttribute("aria-busy", String(busy));
  }

  function refresh() {
    clearTimeout(timer);
    validPreview = false;
    mermaid.removeAttribute("aria-invalid");
    css.removeAttribute("aria-invalid");
    layoutWidth.removeAttribute("aria-invalid");
    errorBox.hidden = true;
    try {
      const result = renderDiagram(mermaid.value, css.value, draft().layout);
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(result.svg)}`;
      image.hidden = false;
      placeholder.hidden = true;
      hasPreview = true;
      validPreview = true;
      dimensions.textContent = `${result.type} · ${Math.round(result.width)} × ${Math.round(result.height)} px`;
      status.textContent = "預覽已更新";
      status.dataset.state = "success";
    } catch (error) {
      const field =
        error instanceof DiagramError && error.field === "css"
          ? css
          : error instanceof DiagramError && error.field === "layout"
            ? layoutWidth
            : mermaid;
      field.setAttribute("aria-invalid", "true");
      errorBox.textContent =
        error instanceof Error ? error.message : "渲染失敗，請修改輸入後重試。";
      errorBox.hidden = false;
      status.textContent = hasPreview ? "輸入有誤 · 保留上次成功預覽" : "尚無有效預覽";
      status.dataset.state = "error";
    } finally {
      preview.setAttribute("aria-busy", "false");
      buttons();
    }
  }

  function scheduleRefresh() {
    validPreview = false;
    markUnsaved();
    buttons();
    clearTimeout(timer);
    status.textContent = "等待更新預覽…";
    status.dataset.state = "pending";
    preview.setAttribute("aria-busy", "true");
    timer = setTimeout(refresh, 400);
  }
  mermaid.addEventListener("input", scheduleRefresh);
  css.addEventListener("input", scheduleRefresh);
  layoutWidth.addEventListener("input", () => {
    slider.value = layoutWidth.value;
    scheduleRefresh();
  });
  slider.addEventListener("input", () => {
    layoutWidth.value = slider.value;
    scheduleRefresh();
  });
  autoDirection.addEventListener("change", scheduleRefresh);
  preserveWidth.addEventListener("change", buttons);
  displayWidth.addEventListener("input", () => {
    displayWidth.removeAttribute("aria-invalid");
    markUnsaved();
  });

  function markUnsaved() {
    officeStatus.textContent = needsReload
      ? "草稿已保留；請重新載入圖表確認目標後再更新。"
      : "有尚未寫入文件的變更；預覽更新不會自動修改文件。";
    officeStatus.dataset.state = "pending";
  }
  document.getElementById("refresh-preview")!.addEventListener("click", refresh);
  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
      event.preventDefault();
      if (!busy) refresh();
    }
  });
  window.addEventListener("beforeunload", (event) => {
    if (busy || fingerprint() !== baseline) {
      event.preventDefault();
      event.returnValue = "";
    }
  });

  async function operate(action: "insert" | "load" | "update") {
    if (!adapter || busy) return;
    if (
      action === "load" &&
      fingerprint() !== baseline &&
      !window.confirm("載入選取圖表會取代目前尚未寫入文件的編輯內容。要繼續載入嗎？")
    )
      return;
    busy = true;
    clearTimeout(timer);
    officeError.hidden = true;
    officeStatus.textContent = {
      insert: "正在準備並插入圖表…",
      load: "正在載入選取圖表…",
      update: "正在更新原圖…",
    }[action];
    officeStatus.dataset.state = "pending";
    buttons();
    try {
      let result: DiagramTarget;
      if (action === "load") {
        result = await adapter.loadSelection();
        // Validate the loaded source before replacing any in-progress editor content.
        renderDiagram(result.record.mermaidSource, result.record.cssSource, result.record.layout);
        mermaid.value = result.record.mermaidSource;
        css.value = result.record.cssSource;
        layoutWidth.value = String(result.record.layout.maxWidth);
        slider.value = layoutWidth.value;
        autoDirection.checked = result.record.layout.autoDirection;
        displayWidth.value = String(Math.round(result.displayWidthCm * 100) / 100);
        preserveWidth.checked = true;
      } else {
        const current = draft();
        validateDraft(current);
        // Always render the CURRENT input, including during the debounce interval.
        const rendered = renderDiagram(current.mermaidSource, current.cssSource, current.layout);
        result =
          action === "insert"
            ? await adapter.insert(current, rendered)
            : await adapter.update(target!, current, rendered, preserveWidth.checked);
        displayWidth.value = String(Math.round(result.displayWidthCm * 100) / 100);
      }
      target = result;
      needsReload = false;
      baseline = fingerprint();
      refresh();
      officeStatus.textContent =
        action === "load"
          ? "圖表已載入，可修改後按「更新圖表」。"
          : "圖表與原始碼已寫入目前文件；請在 Word 儲存文件。";
      officeStatus.dataset.state = "success";
    } catch (error) {
      officeError.textContent =
        error instanceof Error
          ? error.message
          : "文件操作失敗。請確認文件可編輯、Word 版本支援，並重新載入圖表。";
      officeError.hidden = false;
      officeStatus.textContent = "操作未完成；編輯內容已保留。";
      officeStatus.dataset.state = "error";
      if (action === "update" && error instanceof DiagramTargetError) needsReload = true;
      if (
        !Number.isFinite(displayWidth.valueAsNumber) ||
        displayWidth.valueAsNumber < 1 ||
        displayWidth.valueAsNumber > 50
      )
        displayWidth.setAttribute("aria-invalid", "true");
    } finally {
      busy = false;
      if (!validPreview) refresh();
      buttons();
      if (action === "load" && officeError.hidden) mermaid.focus();
    }
  }
  insert.addEventListener("click", () => void operate("insert"));
  load.addEventListener("click", () => void operate("load"));
  update.addEventListener("click", () => void operate("update"));
  newDiagram.addEventListener("click", () => {
    target = null;
    needsReload = false;
    // Keep the current source as a draft for the next insertion.
    officeStatus.textContent = "已切換為新增，保留目前內容；請將 Word 游標移到新的插入位置。";
    officeStatus.dataset.state = "pending";
    officeError.hidden = true;
    buttons();
  });
  refresh();
  initializeOffice((connected) => {
    adapter = connected;
    buttons();
  });
}

function initializeOffice(connect: (adapter: OfficeAdapter) => void) {
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
        if (info.host === Office.HostType.Word) {
          const reason = wordCapability();
          label.textContent = reason || "Word 已連線 · 可插入、載入與更新圖表（尚待實機驗收）。";
          if (!reason) {
            document.getElementById("office-help")!.textContent =
              "新增：將游標放在插入位置。編輯：選取既有圖表，再按「載入選取圖表」。更新會以已載入的圖表為目標，不跟隨游標。";
            connect(new WordAdapter());
          }
        } else {
          label.textContent = info.host
            ? `${info.host} 尚未支援文件操作；仍可編輯與預覽。`
            : "瀏覽器預覽模式；文件操作需要 Office 宿主。";
        }
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
