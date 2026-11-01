/* global document, DOMParser */
import { DiagramError, renderDiagram } from "../src/renderer/renderer";
import { parseTheme } from "../src/renderer/theme";
import { exportSvg } from "../src/renderer/svg";

const results: string[] = [];
let failures = 0;
function assert(condition: unknown, message: string) {
  if (!condition) throw new Error(message);
}
function test(name: string, action: () => void) {
  try {
    action();
    results.push(`PASS ${name}`);
  } catch (error) {
    failures++;
    results.push(`FAIL ${name}: ${error instanceof Error ? error.message : error}`);
  }
}
function rejects(action: () => void) {
  let rejected = false;
  try {
    action();
  } catch {
    rejected = true;
  }
  assert(rejected, "應拒絕不安全或無效輸入");
}

function run() {
  test("空白 CSS 使用預設值", () => assert(Object.keys(parseTheme(" ")).length === 0, "預設主題"));
  test("CSS 註解、色彩、字型、圓角與字重", () => {
    const theme = parseTheme(
      '/* 主題 */ .pele { --pele-fg: #123; --pele-font: "Arial", sans-serif; --pele-radius: 8px; --pele-title-weight: 600; }'
    );
    assert(theme["--pele-fg"] === "#123" && theme["--pele-radius"] === "8px", "解析結果");
  });
  for (const input of [
    "body { color: red; }",
    ".pele { --pele-bg: url(https://example.com/x); }",
    '@import "x"; .pele {}',
    ".pele { --pele-fg: var(--x); }",
    ".pele { color: #fff; }",
    ".pele { --pele-unknown: #fff; }",
    ".pele { --pele-radius: 999px; }",
    ".pele { --pele-font: sans-serif !important; }",
    ".pele {} .pele {}",
  ])
    test(`拒絕 CSS：${input}`, () => rejects(() => parseTheme(input)));
  for (const [type, source] of [
    ["flowchart", "flowchart LR\nA[開始] --> B[完成]"],
    ["sequence", "sequenceDiagram\nAlice->>Bob: Hello"],
    ["class", "classDiagram\nAnimal <|-- Duck"],
    ["state", "stateDiagram-v2\n[*] --> Ready\nReady --> [*]"],
    ["pie", 'pie\n"A" : 40\n"B" : 60'],
    ["subgraph", "flowchart LR\nsubgraph Group\nA --> B\nend"],
  ])
    test(`${type} → 自包含 SVG`, () => {
      const result = renderDiagram(
        source,
        ".pele { --pele-fg: #123456; --pele-surface: #abcdef; }"
      );
      assert(result.width > 0 && result.height > 0, "有效尺寸");
      assert(
        !/var\(|color-mix\(|<script|<style|<foreignObject|<image|\sstyle=/i.test(result.svg),
        "不得殘留 CSS 變數或動態內容"
      );
      assert(
        !new DOMParser().parseFromString(result.svg, "image/svg+xml").querySelector("parsererror"),
        "有效 XML"
      );
    });
  test("主題顏色確實套用，字型在渲染前傳入", () => {
    const result = renderDiagram(
      "flowchart LR\nA --> B",
      ".pele { --pele-fg: #123456; --pele-font: serif; }"
    );
    assert(result.svg.includes("rgb(18, 52, 86)"), "前景色");
    assert(result.svg.includes('font-family="serif"'), "字型");
  });
  test("空 Mermaid 回傳欄位錯誤", () => {
    try {
      renderDiagram("", "");
    } catch (error) {
      assert(error instanceof DiagramError && error.field === "mermaid", "Mermaid 欄位");
      return;
    }
    throw new Error("未拒絕空內容");
  });
  test("錯誤 Mermaid", () => rejects(() => renderDiagram("not a diagram", "")));
  test("過長 Mermaid", () =>
    rejects(() => renderDiagram("flowchart LR\n" + "A".repeat(50001), "")));
  const wrap = (content: string) =>
    `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20">${content}</svg>`;
  for (const content of [
    "<script>alert(1)</script>",
    '<rect onload="alert(1)"/>',
    '<image href="https://example.com/a.svg"/>',
    "<foreignObject/>",
    '<rect fill="url(https://example.com/a)"/>',
    "<style>rect{fill:red}</style>",
    '<rect style="fill:u\\72l(https://example.com/a)"/>',
    '<animate attributeName="fill"/>',
    '<a href="javascript:alert(1)"><rect/></a>',
  ])
    test(`SVG 安全檢查：${content}`, () => rejects(() => exportSvg(wrap(content), {}, 20, 20)));
  test("尺寸限制", () => rejects(() => exportSvg(wrap(""), {}, Infinity, 20)));
  test("XML 外部實體拒絕", () =>
    rejects(() =>
      exportSvg('<!DOCTYPE svg SYSTEM "https://example.com/a">' + wrap(""), {}, 20, 20)
    ));
  test("頁面 CSS 不得污染 SVG", () => {
    const style = document.createElement("style");
    style.textContent = "svg, text { fill: red !important; font-family: fantasy !important; }";
    document.head.appendChild(style);
    try {
      const result = renderDiagram("flowchart LR\nA --> B", ".pele { --pele-fg: #123456; }");
      assert(result.svg.includes("rgb(18, 52, 86)"), "Shadow DOM 隔離");
    } finally {
      style.remove();
    }
  });
  test("拒絕未知變數，且不繼承祖先的 Pele 變數", () => {
    rejects(() => exportSvg(wrap('<rect style="fill:var(--attack)"/>'), {}, 20, 20));
    const previous = document.body.style.getPropertyValue("--pele-fg");
    document.body.style.setProperty("--pele-fg", "url(https://example.com/paint.svg#x)");
    try {
      const result = renderDiagram("flowchart LR\nA --> B", "");
      assert(!result.svg.includes("example.com"), "不得繼承外部資源");
      assert(result.svg.includes("rgb(31, 31, 31)"), "使用 Pele 預設前景色");
    } finally {
      if (previous) document.body.style.setProperty("--pele-fg", previous);
      else document.body.style.removeProperty("--pele-fg");
    }
  });
  test("保留連續空白的 xml:space", () => {
    const result = exportSvg(wrap('<text xml:space="preserve">A    B</text>'), {}, 20, 20);
    assert(result.includes('xml:space="preserve"'), "保留空白處理屬性");
    assert(result.includes("A    B"), "保留文字");
  });
  document.getElementById("test-results")!.textContent =
    `${results.join("\n")}\n\n${results.length - failures}/${results.length} passed`;
  document.body.dataset.testStatus = failures ? "failed" : "passed";
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", run);
else run();
