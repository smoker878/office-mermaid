export type DiagramTheme = Record<string, string>;

const colorNames = new Set([
  ..."bg fg fg-on-light fg-on-dark muted line surface surface-alt border accent critical".split(
    " "
  ),
  ...Array.from({ length: 8 }, (_, index) => `series-${index + 1}`),
]);

/** Deliberately small CSS subset: one .pele rule and literal theme values. */
export function parseTheme(source: string): DiagramTheme {
  if (source.length > 10000) throw new Error("CSS 不可超過 10,000 字元。");
  const text = source.replace(/\/\*[\s\S]*?\*\//g, "").trim();
  if (!text) return {};
  const match = /^\.pele\s*\{([^{}]*)\}\s*$/.exec(text);
  if (!match) throw new Error("CSS 只支援一個 .pele { … } 區塊，不支援其他 selector 或 @ 規則。");
  const theme: DiagramTheme = {};
  for (const declaration of match[1].split(";")) {
    if (!declaration.trim()) continue;
    const property = /^\s*--pele-([a-z0-9-]+)\s*:\s*(.*?)\s*$/.exec(declaration);
    if (!property) throw new Error("請使用 --pele-* 主題變數，每個設定以分號隔開。");
    const [, name, value] = property;
    let valid = false;
    if (colorNames.has(name)) {
      valid = /^#(?:[\da-f]{3}|[\da-f]{4}|[\da-f]{6}|[\da-f]{8})$/i.test(value);
    } else if (name === "font" || name === "font-mono") {
      valid =
        /^(?:[\p{L}\p{N}_ -]+|"[\p{L}\p{N}_ -]+"|'[\p{L}\p{N}_ -]+')(?:\s*,\s*(?:[\p{L}\p{N}_ -]+|"[\p{L}\p{N}_ -]+"|'[\p{L}\p{N}_ -]+'))*$/u.test(
          value
        );
    } else if (name === "radius") {
      valid = /^(?:0|\d+(?:\.\d+)?px)$/.test(value) && parseFloat(value) <= 64;
    } else if (["title-weight", "heading-weight", "bold-weight"].includes(name)) {
      valid = /^(?:normal|bold|[1-9]00)$/.test(value);
    } else {
      throw new Error(`不支援的主題變數：--pele-${name}。`);
    }
    if (!valid)
      throw new Error(
        `--pele-${name} 的值無效。顏色請用 HEX，圓角用 0–64px，字型請用本機字型名稱。`
      );
    theme[`--pele-${name}`] = value;
  }
  return theme;
}
