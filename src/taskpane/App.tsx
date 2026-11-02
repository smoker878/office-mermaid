/* global ResizeObserver, HTMLDivElement, HTMLElement */
import { useEffect, useRef } from "react";
import { EditorPanel } from "./EditorPanel";
import { PreviewPanel } from "./PreviewPanel";
import { SwitchDiagramDialog } from "./SwitchDiagramDialog";
import { useEditor } from "./use-editor";
import { button, errorMessage, hint, panel, primaryButton, statusMessage } from "./styles";

export function App() {
  const editor = useEditor();
  const { target, selected, busy, host } = editor;
  const appRef = useRef<HTMLDivElement>(null);
  const actionBarRef = useRef<HTMLElement>(null);
  useEffect(() => {
    const bar = actionBarRef.current!;
    const observer = new ResizeObserver(() => {
      appRef.current?.style.setProperty(
        "--action-bar-height",
        `${bar.getBoundingClientRect().height + 24}px`
      );
    });
    observer.observe(bar);
    return () => observer.disconnect();
  }, []);

  const showSelection =
    !!selected &&
    !busy &&
    (!target || selected.controlId !== target.controlId || selected.tag !== target.tag);
  const officeHelp = host.adapter
    ? "新增：將游標放在空白位置。編輯：選取圖表後載入；編輯期間選到另一張圖會詢問是否切換，更新目標不會自動改變。"
    : "文件操作需要支援 WordApi 1.4 的 Word；瀏覽器仍可調整排版與預覽。";
  return (
    <div
      ref={appRef}
      data-mode={target ? "edit" : "insert"}
      className={`min-h-screen pb-[var(--action-bar-height)] [--action-bar-height:200px] ${target ? "bg-edit-page" : "bg-page"}`}
    >
      <header className="mx-auto max-w-300 p-4 pb-0 min-[421px]:p-6 min-[421px]:pb-0">
        <h1 className="mt-2 text-title font-bold leading-tight">Office Mermaid</h1>
        <p id="host-status" className={`${hint} mt-2`} role="status">
          {host.hostStatus}
        </p>
        <div className="mt-4 mb-3 flex flex-wrap gap-2">
          <button
            id="new-diagram"
            className={button}
            type="button"
            disabled={busy}
            onClick={editor.newDiagram}
          >
            {target ? "＋ 改為插入新圖" : "＋ 新增圖表"}
          </button>
          <button
            id="load-diagram"
            className={button}
            type="button"
            disabled={!host.adapter || busy}
            onClick={() => editor.requestLoad(selected || undefined)}
          >
            編輯選取圖表
          </button>
        </div>
        <p
          id="editing-target"
          className={`mb-2 rounded-panel border p-3 font-semibold wrap-anywhere ${target ? "border-edit-border bg-edit-tint text-edit-ink" : "border-border bg-tint text-primary"}`}
          aria-live="polite"
        >
          <span className="block">{target ? "✎ 編輯既有圖表" : "＋ 插入新圖表"}</span>
          {target ? (
            <span className="mt-1 block text-small font-normal">
              {target.record.mermaidSource.replace(/\s+/g, " ").slice(0, 64)}
            </span>
          ) : null}
          {editor.needsReload ? " · 請重新載入確認（草稿保留中）" : ""}
        </p>
        <p id="office-help" className={`${hint} mt-2`}>
          {officeHelp}
        </p>
        <div
          id="selection-notice"
          className="mt-3 flex flex-wrap items-center gap-3 rounded-panel bg-tint p-3"
          hidden={!showSelection}
        >
          <p id="selection-message" className="min-w-35 flex-1" role="status">
            {target
              ? "已選取另一張 Mermaid 圖表；目前編輯目標不變。"
              : "已選取 Mermaid 圖表，可載入編輯。"}
          </p>
          <button
            id="edit-selected"
            className={button}
            type="button"
            disabled={busy}
            onClick={() => {
              if (selected) editor.requestLoad(selected);
            }}
          >
            {target ? "切換編輯這張圖" : "編輯這張"}
          </button>
        </div>
        <p id="selection-help" className={`${hint} mt-2`} hidden={!host.selectionHelp}>
          {host.selectionHelp}
        </p>
      </header>
      <main className="mx-auto grid max-w-300 grid-cols-1 gap-6 p-4 min-[421px]:p-6 min-[800px]:grid-cols-2 min-[800px]:items-start">
        <EditorPanel editor={editor} />
        <PreviewPanel editor={editor} />
        <section
          ref={actionBarRef}
          className={`${panel} office-panel fixed inset-x-0 bottom-0 z-10 max-h-[45vh] overflow-y-auto shadow-action ${target ? "border-edit-border bg-edit-tint!" : ""}`}
          aria-labelledby="office-title"
          aria-busy={busy}
        >
          <h2 id="office-title" className="sr-only">
            文件操作
          </h2>
          <p className={`mb-3 font-semibold ${target ? "text-edit-ink" : "text-primary"}`}>
            {target ? "更新原圖，不會新增圖片" : "在 Word 游標位置插入新圖表"}
          </p>
          <div className="flex flex-wrap items-center gap-3" aria-describedby="office-help">
            <button
              id="insert-diagram"
              className={`${primaryButton} w-full`}
              type="button"
              hidden={!!target}
              disabled={!editor.canWrite}
              onClick={() => void editor.operate("insert")}
            >
              {editor.operation === "insert" ? "插入中…" : "插入新圖表"}
            </button>
            <button
              id="update-diagram"
              className={`${primaryButton} w-full`}
              type="button"
              hidden={!target}
              disabled={!editor.canWrite || !target}
              onClick={() => void editor.operate("update")}
            >
              {editor.operation === "update" ? "更新中…" : "更新原圖"}
            </button>
          </div>
          <p id="draft-status" className={`${hint} mt-3`} aria-live="polite">
            {busy
              ? editor.operation === "load"
                ? "正在載入圖表…"
                : "正在寫入文件…"
              : editor.needsReload
                ? "草稿已保留，請重新載入確認目標"
                : editor.dirty
                  ? "尚未套用變更 · 預覽不會自動修改文件"
                  : target
                    ? "已與文件同步 · 請在 Word 儲存文件"
                    : "保留目前內容，插入成另一張圖"}
          </p>
          <p
            id="office-status"
            className={`${statusMessage} mb-0`}
            role="status"
            aria-live="polite"
            data-state={editor.officeStatus.state}
          >
            {editor.officeStatus.text}
          </p>
          <p
            className={`${hint} mt-2`}
            hidden={!host.adapter || busy || editor.preview.valid || editor.needsReload}
          >
            {editor.preview.pending
              ? "請等待有效預覽，再插入或更新圖表。"
              : "請先修正輸入錯誤，再插入或更新圖表。"}
          </p>
          <p id="office-error" className={errorMessage} role="alert" hidden={!editor.officeError}>
            {editor.officeError}
          </p>
        </section>
      </main>
      <SwitchDiagramDialog editor={editor} />
    </div>
  );
}
