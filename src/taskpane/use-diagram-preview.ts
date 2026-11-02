/* global setTimeout, clearTimeout */
import { useCallback, useEffect, useState } from "react";
import { DiagramError, renderDiagram, type DiagramPreview } from "../renderer/renderer";
import type { DiagramDraft } from "../diagram/record";

interface PreviewState {
  result: DiagramPreview | null;
  error: DiagramError | null;
  key: string | null;
}

type PreviewInput = Pick<DiagramDraft, "mermaidSource" | "cssSource" | "layout">;

function renderKey(draft: PreviewInput): string {
  return JSON.stringify([draft.mermaidSource, draft.cssSource, draft.layout]);
}

export function useDiagramPreview(draft: PreviewInput, busy: boolean) {
  const [preview, setPreview] = useState<PreviewState>({ result: null, error: null, key: null });
  const key = renderKey(draft);
  const refresh = useCallback((current: PreviewInput) => {
    try {
      const result = renderDiagram(current.mermaidSource, current.cssSource, current.layout);
      setPreview({ result, error: null, key: renderKey(current) });
    } catch (error) {
      const failure =
        error instanceof DiagramError
          ? error
          : new DiagramError(
              "mermaid",
              error instanceof Error ? error.message : "渲染失敗，請重試。"
            );
      setPreview((previous) => ({ ...previous, error: failure, key: renderKey(current) }));
    }
  }, []);

  const { mermaidSource, cssSource } = draft;
  const { maxWidth, autoDirection } = draft.layout;
  useEffect(() => {
    if (busy || preview.key === key) return;
    const timer = setTimeout(
      () =>
        refresh({
          mermaidSource,
          cssSource,
          layout: { maxWidth, autoDirection },
        }),
      preview.key === null ? 0 : 400
    );
    return () => clearTimeout(timer);
  }, [mermaidSource, cssSource, maxWidth, autoDirection, key, preview.key, busy, refresh]);

  const pending = preview.key !== key;
  return {
    ...preview,
    pending,
    valid: !pending && !preview.error && !!preview.result,
    refresh,
  };
}
