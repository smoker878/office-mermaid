import type { EditorController } from "./use-editor";
import { hint, panel, statusMessage } from "./styles";

export function PreviewPanel({ editor }: { editor: EditorController }) {
  const { result, pending, error } = editor.preview;
  const state = pending ? "pending" : error ? "error" : "success";
  const text = pending
    ? "等待更新預覽…"
    : error
      ? result
        ? "輸入有誤 · 保留上次成功預覽"
        : "尚無有效預覽"
      : "預覽已更新";
  return (
    <section
      className={`${panel} min-[800px]:sticky min-[800px]:top-6`}
      aria-labelledby="preview-title"
    >
      <div className="flex flex-wrap items-center justify-between gap-3">
        <h2 id="preview-title" className="text-heading font-bold leading-tight">
          SVG 預覽
        </h2>
        <span className="rounded-panel bg-tint px-2 py-1 text-small text-primary">Pele 1.13</span>
      </div>
      <p
        id="render-status"
        className={statusMessage}
        role="status"
        aria-live="polite"
        data-state={state}
      >
        {text}
      </p>
      <div
        id="preview-region"
        className="flex min-h-60 items-center justify-center overflow-auto rounded-panel border border-border bg-page p-4"
        aria-busy={pending}
      >
        <p id="preview-placeholder" className={hint} hidden={!!result}>
          輸入 Mermaid，即可在這裡檢查圖表。
        </p>
        <img
          id="svg-preview"
          className="block h-auto max-w-full"
          alt="Mermaid 圖表的 SVG 預覽"
          hidden={!result}
          src={
            result
              ? `data:image/svg+xml;charset=utf-8,${encodeURIComponent(result.svg)}`
              : undefined
          }
        />
      </div>
      <p id="preview-dimensions" className={`${hint} mt-3`}>
        {result
          ? `${result.type} · ${Math.round(result.width)} × ${Math.round(result.height)} px`
          : ""}
      </p>
      <p className={`${hint} mt-3`}>
        預覽使用已解析主題的自包含 SVG，不依賴頁面 CSS。字型仍須由檢視裝置提供。
      </p>
    </section>
  );
}
