/* global document, window, Office, HTMLElement, HTMLTextAreaElement, HTMLImageElement, HTMLInputElement, HTMLButtonElement, ResizeObserver, setTimeout, clearTimeout */
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
  const selectionNotice = document.getElementById("selection-notice")!;
  const editSelected = document.getElementById("edit-selected") as HTMLButtonElement;
  let selected: DiagramTarget | null = null;
  let selectionVersion = 0;
  let selectionTimer: ReturnType<typeof setTimeout>;
  let scanning = false;
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
  const actionBar = document.querySelector<HTMLElement>(".office-panel")!;
  new ResizeObserver(() => {
    document.documentElement.style.setProperty(
      "--action-bar-height",
      `${actionBar.getBoundingClientRect().height + 24}px`
    );
  }).observe(actionBar);

  function buttons() {
    insert.disabled = !adapter || busy || !validPreview || !!target;
    load.disabled = !adapter || busy;
    update.disabled = !adapter || busy || !validPreview || !target || needsReload;
    insert.hidden = !!target;
    update.hidden = !target;
    newDiagram.disabled = busy;
    editSelected.disabled = busy;
    insert.classList.toggle("primary", !target);
    update.classList.toggle("primary", !!target);
    editingTarget.textContent = target
      ? `正在編輯：${target.record.mermaidSource.replace(/\s+/g, " ").slice(0, 64)}`
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
    document.getElementById("size-summary")!.textContent =
      `${layoutWidth.value} px · ${autoDirection.checked ? "自動方向" : "保留方向"} · ${target && preserveWidth.checked ? "保留文件寬度" : `${displayWidth.value} 公分`}`;
    selectionNotice.hidden =
      !selected ||
      busy ||
      (!!target && selected.controlId === target.controlId && selected.tag === target.tag);
    document.getElementById("selection-message")!.textContent = target
      ? "已選取另一張 Mermaid 圖表；目前編輯目標不變。"
      : "已選取 Mermaid 圖表，可載入編輯。";
    if (!adapter) {
      officeStatus.textContent = "文件操作尚未連線；仍可編輯與預覽。";
    } else if (!busy && !validPreview && !needsReload) {
      officeStatus.textContent = "請先修正輸入錯誤，再插入或更新圖表。";
    }
  }

  // Selection only offers an action: it never replaces the draft or editing target.
  function selectionChanged() {
    const version = ++selectionVersion;
    selected = null;
    selectionNotice.hidden = true;
    clearTimeout(selectionTimer);
    selectionTimer = setTimeout(async () => {
      if (!adapter || busy) return;
      if (scanning) {
        selectionChanged();
        return;
      }
      scanning = true;
      try {
        const candidate = await adapter.loadSelection();
        if (version === selectionVersion && !busy) selected = candidate;
      } catch {
        // Ordinary pictures, empty selection, and multi-selection are not errors here.
        // The manual edit action still exposes actionable loading errors.
      } finally {
        scanning = false;
        if (version === selectionVersion) buttons();
      }
    }, 200);
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
      if (field === css) document.getElementById("style-options")!.setAttribute("open", "");
      if (field === layoutWidth) document.getElementById("size-options")!.setAttribute("open", "");
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
    buttons();
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

  async function operate(action: "insert" | "load" | "update", expected?: DiagramTarget) {
    if (!adapter || busy) return;
    if (
      action === "load" &&
      fingerprint() !== baseline &&
      !window.confirm("目前有尚未套用的修改。要放棄修改並編輯選取圖表嗎？取消可繼續目前的編輯。")
    )
      return;
    busy = true;
    ++selectionVersion;
    clearTimeout(selectionTimer);
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
        if (expected && (result.controlId !== expected.controlId || result.tag !== expected.tag))
          throw new Error("選取的圖表已改變，請重新選取後按「編輯選取圖表」。目前草稿已保留。");
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
          ? "圖表已載入；修改後按「更新這張圖表」。移動游標不會改變編輯目標。"
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
      ) {
        displayWidth.setAttribute("aria-invalid", "true");
        document.getElementById("size-options")!.setAttribute("open", "");
      }
    } finally {
      busy = false;
      if (!validPreview) refresh();
      buttons();
      if (action === "load" && officeError.hidden) mermaid.focus();
      selectionChanged();
    }
  }
  insert.addEventListener("click", () => void operate("insert"));
  load.addEventListener("click", () => void operate("load"));
  editSelected.addEventListener("click", () => {
    if (selected) void operate("load", selected);
  });
  update.addEventListener("click", () => void operate("update"));
  newDiagram.addEventListener("click", () => {
    target = null;
    needsReload = false;
    // Keep the current source as a draft for the next insertion.
    officeStatus.textContent =
      "新增圖表：已保留目前內容，請將 Word 游標放到空白插入位置，再按「插入文件」。";
    officeStatus.dataset.state = "pending";
    officeError.hidden = true;
    buttons();
    mermaid.focus();
  });
  refresh();
  initializeOffice((connected) => {
    adapter = connected;
    officeStatus.textContent =
      "將游標放到空白位置後插入；要修改既有圖表，請選取後按上方「編輯選取圖表」。";
    buttons();
    selectionChanged();
  }, selectionChanged);
}

function initializeOffice(connect: (adapter: OfficeAdapter) => void, selectionChanged: () => void) {
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
              "新增：將游標放在空白位置。編輯：選取圖表，再按「編輯選取圖表」。更新目標不隨游標改變。";
            connect(new WordAdapter());
            const fallback = () => {
              const help = document.getElementById("selection-help")!;
              help.textContent = "選圖提示無法啟用；請選取圖表後，按上方「編輯選取圖表」。";
              help.hidden = false;
            };
            try {
              Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                selectionChanged,
                (result) => {
                  if (result.status === Office.AsyncResultStatus.Failed) fallback();
                }
              );
              window.addEventListener("unload", () => {
                Office.context.document.removeHandlerAsync(
                  Office.EventType.DocumentSelectionChanged,
                  { handler: selectionChanged }
                );
              });
            } catch {
              fallback();
            }
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
