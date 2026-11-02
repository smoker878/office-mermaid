import { PeleError, render } from "pele";
import { parseTheme } from "./theme";
import { exportSvg } from "./svg";
import { validateLayout, type LayoutSettings } from "../diagram/record";

export interface DiagramPreview {
  svg: string;
  width: number;
  height: number;
  type: string;
}

export class DiagramError extends Error {
  constructor(
    public readonly field: "mermaid" | "css" | "layout",
    message: string
  ) {
    super(message);
    this.name = "DiagramError";
  }
}

export function renderDiagram(
  mermaidSource: string,
  cssSource: string,
  layout?: LayoutSettings
): DiagramPreview {
  if (layout) {
    try {
      validateLayout(layout);
    } catch (error) {
      throw new DiagramError("layout", error instanceof Error ? error.message : "排版設定無效。");
    }
  }
  let theme;
  try {
    theme = parseTheme(cssSource);
  } catch (error) {
    throw new DiagramError("css", error instanceof Error ? error.message : "CSS 解析失敗。");
  }
  if (!mermaidSource.trim())
    throw new DiagramError("mermaid", "請輸入 Mermaid 語法，例如 flowchart LR，再加入節點與連線。");
  try {
    const result = render(mermaidSource, {
      fontFamily: theme["--pele-font"] || "sans-serif",
      fontFamilyMono: theme["--pele-font-mono"] || "monospace",
      responsive: false,
      autoDirection: layout?.autoDirection ?? false,
      maxWidth: layout?.maxWidth,
      links: false,
      images: false,
      limit: 50000,
      outputLimit: 4000000,
      maxEdges: 1000,
      idPrefix: "office-mermaid",
    });
    return { ...result, svg: exportSvg(result.svg, theme, result.width, result.height) };
  } catch (error) {
    const message =
      error instanceof PeleError || error instanceof Error ? error.message : "未知的渲染錯誤。";
    throw new DiagramError("mermaid", `無法產生預覽，請檢查 Mermaid 語法或圖表大小。\n${message}`);
  }
}
