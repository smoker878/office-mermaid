/* global document, DOMParser, SVGSVGElement, SVGElement, getComputedStyle, XMLSerializer */
import type { DiagramTheme } from "./theme";

const namespace = "http://www.w3.org/2000/svg";
const tags = new Set(
  `svg g path rect circle ellipse line polyline polygon text tspan defs marker clipPath
  mask pattern linearGradient radialGradient stop title desc`.split(/\s+/)
);
const properties = `fill fill-opacity fill-rule stroke stroke-width stroke-opacity
  stroke-dasharray stroke-dashoffset stroke-linecap stroke-linejoin stroke-miterlimit
  opacity color font-family font-size font-weight font-style text-anchor dominant-baseline
  alignment-baseline letter-spacing word-spacing text-decoration paint-order visibility
  display stop-color stop-opacity rx ry direction text-transform`.split(/\s+/);
const attributes = new Set([
  ...properties,
  ...`id class style viewBox width height x y x1 x2 y1 y2 cx cy r dx dy d points transform
  preserveAspectRatio marker-start marker-mid marker-end markerWidth markerHeight markerUnits
  refX refY orient clip-path clip-rule clipPathUnits mask maskUnits maskContentUnits
  patternUnits patternContentUnits patternTransform gradientUnits gradientTransform offset
  spreadMethod fx fy fr textLength lengthAdjust vector-effect role aria-label
  aria-labelledby aria-describedby aria-roledescription xml:space`.split(/\s+/),
]);

function referencedVariables(value: string): string[] {
  return (value.match(/var\s*\(\s*--[\w-]+/gi) || []).map((reference) =>
    reference.replace(/^var\s*\(\s*/i, "")
  );
}

function safeValue(value: string): boolean {
  if (/[\\<>]|\/\*|@|expression\s*\(/i.test(value)) return false;
  if (referencedVariables(value).some((name) => !/^--(?:pele-[a-z0-9-]+|_[a-z]+)$/.test(name)))
    return false;
  const withoutLocalUrls = value.replace(/url\(["']?#[\w.:-]+["']?\)/gi, "");
  return !/url\s*\(|(?:https?|data|javascript):/i.test(withoutLocalUrls);
}

/** Validate before attaching to the DOM; no scripts, links, images or external CSS. */
function parseSafeSvg(source: string): SVGSVGElement {
  if (/<!DOCTYPE|<!ENTITY|<\?/i.test(source)) throw new Error("SVG 包含不支援的 XML 宣告。");
  const xml = new DOMParser().parseFromString(source, "image/svg+xml");
  if (xml.querySelector("parsererror")) throw new Error("渲染結果不是有效的 SVG。");
  const root = xml.documentElement;
  if (root.localName !== "svg" || root.namespaceURI !== namespace)
    throw new Error("缺少 SVG 根元素。");
  for (const element of [root, ...Array.from(root.querySelectorAll("*"))]) {
    if (element.namespaceURI !== namespace || !tags.has(element.localName)) {
      throw new Error(`SVG 包含不支援的元素：${element.localName}。`);
    }
    for (const attribute of Array.from(element.attributes)) {
      if (attribute.name === "xmlns" && element === root) continue;
      if (attribute.name === "xml:space" && !/^(?:preserve|default)$/.test(attribute.value)) {
        throw new Error("SVG 包含無效的 xml:space 屬性。");
      }
      if (
        /^on/i.test(attribute.name) ||
        /href$/i.test(attribute.name) ||
        !safeValue(attribute.value)
      ) {
        throw new Error("SVG 包含不安全的屬性或外部資源。");
      }
      if (!attributes.has(attribute.name)) element.removeAttributeNode(attribute);
    }
    const styled = element as unknown as SVGElement;
    for (const name of Array.from(styled.style)) {
      // Browsers expand Pele's font shorthand into these harmless reset values.
      if (
        (name.startsWith("font-") || name === "line-height") &&
        /^(?:normal|auto|none|100%)$/.test(styled.style.getPropertyValue(name))
      )
        continue;
      if (!properties.includes(name) && !/^--(?:pele-[a-z0-9-]+|_[a-z]+)$/.test(name)) {
        throw new Error(`SVG 包含不支援的樣式：${name}。`);
      }
    }
  }
  return document.importNode(root, true) as unknown as SVGSVGElement;
}

/** Freeze browser-resolved theme values into presentation attributes, in isolation. */
export function exportSvg(
  source: string,
  theme: DiagramTheme,
  width: number,
  height: number
): string {
  if (![width, height].every((size) => Number.isFinite(size) && size > 0 && size <= 30000)) {
    throw new Error("圖表尺寸超出支援範圍，請縮小圖表。");
  }
  const svg = parseSafeSvg(source);
  for (const [name, value] of Object.entries(theme)) svg.style.setProperty(name, value);
  const host = document.createElement("div");
  host.style.cssText = "all:initial;position:fixed;left:-100000px;top:0;pointer-events:none";
  // all:initial does not reset custom properties. Block inheritance BEFORE attachment.
  for (const name of referencedVariables(source)) host.style.setProperty(name, "initial");
  host.setAttribute("aria-hidden", "true");
  const shadow = host.attachShadow({ mode: "closed" });
  shadow.appendChild(svg);
  document.body.appendChild(host);
  try {
    const canvas = document.createElement("canvas");
    canvas.width = canvas.height = 1;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) throw new Error("瀏覽器無法解析 SVG 色彩。");
    const colors = new Map<string, string>();
    function color(value: string): string {
      if (!/^(?:color\(|color-mix\(|oklch\(|oklab\(|lab\(|lch\()/i.test(value)) return value;
      if (colors.has(value)) return colors.get(value)!;
      context!.clearRect(0, 0, 1, 1);
      context!.fillStyle = value;
      context!.fillRect(0, 0, 1, 1);
      const [r, g, b, alpha] = context!.getImageData(0, 0, 1, 1).data;
      const result =
        alpha === 255 ? `rgb(${r}, ${g}, ${b})` : `rgba(${r}, ${g}, ${b}, ${alpha / 255})`;
      colors.set(value, result);
      return result;
    }
    const elements = [svg, ...Array.from(svg.querySelectorAll<SVGElement>("*"))];
    // Snapshot ALL computed styles before removing any inherited variables.
    const snapshots = elements.map((element) => {
      const computed = getComputedStyle(element);
      return properties.map((name) => [name, computed.getPropertyValue(name)]);
    });
    elements.forEach((element, index) => {
      element.removeAttribute("style");
      element.removeAttribute("class");
      for (const [name, raw] of snapshots[index]) {
        if (!raw || /var\(/i.test(raw)) throw new Error(`無法解析 SVG 樣式：${name}。`);
        const value = color(raw);
        if (!safeValue(value)) throw new Error("SVG 樣式解析後包含不安全的資源。");
        // auto is not a valid SVG rx/ry presentation attribute.
        if ((name === "rx" || name === "ry") && value === "auto") continue;
        element.setAttribute(name, value);
      }
    });
    svg.setAttribute("width", String(width));
    svg.setAttribute("height", String(height));
    const background = document.createElementNS(namespace, "rect");
    background.setAttribute("width", "100%");
    background.setAttribute("height", "100%");
    background.setAttribute("fill", theme["--pele-bg"] || "#ffffff");
    svg.insertBefore(background, svg.firstChild);
    const result = new XMLSerializer().serializeToString(svg);
    if (result.length > 4000000) throw new Error("SVG 超過 4 MB，請縮小圖表。");
    return result;
  } finally {
    host.remove();
  }
}
