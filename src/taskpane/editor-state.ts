import { defaultLayout, type DiagramDraft } from "../diagram/record";

export interface EditorFields {
  mermaidSource: string;
  cssSource: string;
  layoutWidth: string;
  autoDirection: boolean;
  displayWidth: string;
}

export const initialFields: EditorFields = {
  mermaidSource: `flowchart LR
  A[輸入 Mermaid] --> B[調整主題]
  B --> C[預覽 SVG]
  C --> D[準備插入 Office]`,
  cssSource: `.pele {
  --pele-bg: #ffffff;
  --pele-fg: #243247;
  --pele-line: #526581;
  --pele-surface: #edf3fc;
  --pele-border: #8da9d2;
  --pele-accent: #2563eb;
  --pele-font: sans-serif;
  --pele-radius: 4px;
}`,
  layoutWidth: String(defaultLayout.maxWidth),
  autoDirection: defaultLayout.autoDirection,
  displayWidth: "14",
};

export function toDraft(fields: EditorFields): DiagramDraft {
  return {
    mermaidSource: fields.mermaidSource,
    cssSource: fields.cssSource,
    layout: {
      maxWidth: fields.layoutWidth === "" ? NaN : Number(fields.layoutWidth),
      autoDirection: fields.autoDirection,
    },
    displayWidthCm: fields.displayWidth === "" ? NaN : Number(fields.displayWidth),
  };
}

export function fingerprint(fields: EditorFields): string {
  return JSON.stringify(fields);
}

export interface StatusMessage {
  text: string;
  state?: "pending" | "success" | "error";
}
