import type { EditorController } from "./use-editor";
import { button, errorMessage, hint, panel, textInput } from "./styles";

export function EditorPanel({ editor }: { editor: EditorController }) {
  const { fields, changeField, busy, preview } = editor;
  const codeInput = `${textInput} min-h-40 resize-y bg-surface! p-3 font-mono leading-relaxed [tab-size:2]`;
  const settings = "mt-6 min-w-0 rounded-panel border border-border p-4";
  const checkbox = "flex min-h-11 items-center gap-2 font-normal";
  return (
    <section className={panel} aria-labelledby="editor-title">
      <h2 id="editor-title" className="text-heading font-bold leading-tight">
        圖表內容
      </h2>
      <div className="mt-6">
        <label htmlFor="mermaid-source" className="block font-semibold">
          Mermaid 原始碼
        </label>
        <p id="mermaid-help" className={`${hint} mt-1 mb-2`}>
          輸入後自動預覽。支援 Pele 的 Mermaid 語法，不載入外部圖片或連結。
        </p>
        <textarea
          id="mermaid-source"
          ref={editor.sourceRef}
          className={codeInput}
          rows={10}
          spellCheck={false}
          autoCapitalize="off"
          aria-describedby="mermaid-help render-error"
          aria-invalid={preview.error?.field === "mermaid" || undefined}
          disabled={busy}
          value={fields.mermaidSource}
          onChange={(event) => changeField("mermaidSource", event.target.value)}
        />
      </div>
      <details
        id="style-options"
        className="mt-6 border-t border-border pt-2"
        open={editor.styleOpen}
        onToggle={(event) => editor.setStyleOpen(event.currentTarget.open)}
      >
        <summary className="min-h-11 cursor-pointer pt-1 font-semibold">樣式與 CSS</summary>
        <div className="mt-6">
          <label htmlFor="css-source" className="block font-semibold">
            CSS／Pele 主題
          </label>
          <p id="css-help" className={`${hint} mt-1 mb-2`}>
            使用單一 .pele 區塊；顏色限 HEX，字型限本機字型。
          </p>
          <textarea
            id="css-source"
            className={codeInput}
            rows={11}
            spellCheck={false}
            autoCapitalize="off"
            aria-describedby="css-help render-error"
            aria-invalid={preview.error?.field === "css" || undefined}
            disabled={busy}
            value={fields.cssSource}
            onChange={(event) => changeField("cssSource", event.target.value)}
          />
          <details className={`mt-3 ${hint}`}>
            <summary className="min-h-8 cursor-pointer pt-1">支援哪些主題設定？</summary>
            <p className="mt-2">
              顏色：bg、fg、fg-on-light、fg-on-dark、muted、line、surface、surface-alt、border、accent、critical、series-1
              至 series-8。
            </p>
            <p className="mt-2">
              字型：font、font-mono。字重：title-weight、heading-weight、bold-weight（normal、bold
              或 100–900）。圓角：radius（0–64px）。皆須加上 --pele- 前綴。
            </p>
            <p className="mt-2">
              不支援其他 selector、@import、@media、外部字型、url() 或變數巢狀引用。Mermaid
              內的固定樣式可能覆蓋主題。
            </p>
          </details>
        </div>
      </details>
      <details
        id="size-options"
        className="mt-6 border-t border-border pt-2"
        open={editor.sizeOpen}
        onToggle={(event) => editor.setSizeOpen(event.currentTarget.open)}
      >
        <summary className="min-h-11 cursor-pointer pt-1 font-semibold">
          排版與圖片大小
          <span id="size-summary" className={`block pl-4 font-normal ${hint}`}>
            {fields.layoutWidth} px · {fields.autoDirection ? "自動方向" : "保留方向"} ·{" "}
            {editor.target && editor.preserveWidth ? "保留文件寬度" : `${fields.displayWidth} 公分`}
          </span>
        </summary>
        <fieldset id="layout-settings" className={settings} disabled={busy}>
          <legend className="px-2 font-semibold">圖表排版大小</legend>
          <label htmlFor="layout-width" className="block font-semibold">
            可用寬度（px）
          </label>
          <input
            id="layout-width"
            className={`${textInput} my-2 min-h-11`}
            type="number"
            min={200}
            max={4000}
            step={1}
            aria-describedby="layout-help render-error"
            aria-invalid={preview.error?.field === "layout" || undefined}
            value={fields.layoutWidth}
            onChange={(event) => changeField("layoutWidth", event.target.value)}
          />
          <input
            id="layout-width-slider"
            className="block min-h-11 w-full accent-primary"
            type="range"
            min={200}
            max={4000}
            step={50}
            aria-label="調整排版寬度"
            value={fields.layoutWidth || "800"}
            onChange={(event) => changeField("layoutWidth", event.target.value)}
          />
          <label className={checkbox}>
            <input
              id="auto-direction"
              className="shrink-0 accent-primary"
              type="checkbox"
              checked={fields.autoDirection}
              onChange={(event) => changeField("autoDirection", event.target.checked)}
            />{" "}
            依寬度自動調整方向
          </label>
          <p id="layout-help" className={`${hint} mt-2`}>
            改變排版，不只是縮放。部分圖表會換行、調整間距；開啟自動方向後，窄版可能轉為直向。實際寬度可能超過設定。
          </p>
        </fieldset>
        <fieldset id="display-settings" className={settings} disabled={busy}>
          <legend className="px-2 font-semibold">文件顯示大小</legend>
          <label htmlFor="display-width" className="block font-semibold">
            圖片寬度（公分）
          </label>
          <input
            id="display-width"
            className={`${textInput} my-2 min-h-11`}
            type="number"
            min={1}
            max={50}
            step={0.1}
            aria-describedby="display-help"
            aria-invalid={editor.displayInvalid || undefined}
            disabled={busy || (!!editor.target && editor.preserveWidth)}
            value={fields.displayWidth}
            onChange={(event) => changeField("displayWidth", event.target.value)}
          />
          <label className={checkbox}>
            <input
              id="preserve-width"
              className="shrink-0 accent-primary"
              type="checkbox"
              checked={editor.preserveWidth}
              onChange={(event) => editor.setPreserveWidth(event.target.checked)}
            />{" "}
            更新時保留文件中的圖片寬度
          </label>
          <p id="display-help" className={`${hint} mt-2`}>
            只縮放圖片，不改排版；高度依比例計算。更新時取消保留，才會使用新的寬度。
          </p>
        </fieldset>
      </details>
      <div className="mt-6 flex flex-wrap items-center gap-3">
        <button
          id="refresh-preview"
          className={button}
          type="button"
          disabled={busy}
          onClick={editor.refresh}
        >
          重新預覽
        </button>
        <span className={hint}>⌘ / Ctrl + Enter</span>
      </div>
      <p id="render-error" className={errorMessage} role="alert" hidden={!preview.error}>
        {preview.error?.message}
      </p>
    </section>
  );
}
