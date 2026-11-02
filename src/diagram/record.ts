/* global crypto, DOMParser */
export interface LayoutSettings {
  maxWidth: number;
  autoDirection: boolean;
}

export interface DiagramDraft {
  mermaidSource: string;
  cssSource: string;
  layout: LayoutSettings;
  displayWidthCm: number;
}

export interface DiagramRecord extends DiagramDraft {
  id: string;
  schemaVersion: 1;
  rendererVersion: string;
}

export const defaultLayout: LayoutSettings = { maxWidth: 800, autoDirection: false };
export const recordNamespace = "urn:office-mermaid:diagram:1";
export const tagPrefix = "office-mermaid:";
// XML 1.0 cannot represent these characters, even when escaped.
// eslint-disable-next-line no-control-regex
const invalidXmlCharacters = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/;

export function validateLayout(layout: LayoutSettings): void {
  if (
    !layout ||
    !Number.isInteger(layout.maxWidth) ||
    layout.maxWidth < 200 ||
    layout.maxWidth > 4000 ||
    typeof layout.autoDirection !== "boolean"
  )
    throw new Error("排版寬度須為 200–4000 px 的整數。自動方向須為開啟或關閉。");
}

export function validateDraft(draft: DiagramDraft): void {
  validateLayout(draft.layout);
  if (
    typeof draft.mermaidSource !== "string" ||
    !draft.mermaidSource.trim() ||
    draft.mermaidSource.length > 50000 ||
    typeof draft.cssSource !== "string" ||
    draft.cssSource.length > 10000
  )
    throw new Error("圖表原始碼無效或超過大小限制。");
  if (
    !Number.isFinite(draft.displayWidthCm) ||
    draft.displayWidthCm < 1 ||
    draft.displayWidthCm > 50
  )
    throw new Error("文件顯示寬度須介於 1–50 公分。");
  if (invalidXmlCharacters.test(draft.mermaidSource + draft.cssSource))
    throw new Error("原始碼含有無法存入 XML 的控制字元。");
}

export function createRecord(draft: DiagramDraft, id: string = crypto.randomUUID()): DiagramRecord {
  validateDraft(draft);
  return {
    mermaidSource: draft.mermaidSource,
    cssSource: draft.cssSource,
    layout: { ...draft.layout },
    displayWidthCm: draft.displayWidthCm,
    id,
    schemaVersion: 1,
    rendererVersion: "pele@1.13.0",
  };
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function serializeRecord(record: DiagramRecord): string {
  validateDraft(record);
  return `<diagram xmlns="${recordNamespace}">${escapeXml(JSON.stringify(record))}</diagram>`;
}

export function parseRecord(xml: string): DiagramRecord {
  if (xml.length > 500000 || /<!DOCTYPE|<!ENTITY/i.test(xml))
    throw new Error("圖表資料過大或含有不支援的 XML。");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (
    doc.querySelector("parsererror") ||
    root.localName !== "diagram" ||
    root.namespaceURI !== recordNamespace ||
    root.children.length
  )
    throw new Error("文件內的圖表資料已損壞或格式不支援。");
  let value: DiagramRecord;
  try {
    value = JSON.parse(root.textContent || "");
  } catch {
    throw new Error("無法讀取文件內的圖表原始碼。");
  }
  if (!value || value.schemaVersion !== 1)
    throw new Error("此圖表資料版本不支援，請使用相容版本的外掛。");
  if (
    typeof value.id !== "string" ||
    !/^[\da-f-]{36}$/i.test(value.id) ||
    typeof value.rendererVersion !== "string" ||
    value.rendererVersion.length > 100
  )
    throw new Error("圖表識別資料無效。");
  validateDraft(value);
  return { ...createRecord(value, value.id), rendererVersion: value.rendererVersion };
}

export function partIdFromTag(tag: string): string {
  if (!tag.startsWith(tagPrefix) || !/^\{?[\da-f-]{36}\}?$/i.test(tag.slice(tagPrefix.length)))
    throw new Error("請選取 Office Mermaid 插入的圖表，再按「載入選取圖表」。");
  return tag.slice(tagPrefix.length);
}
