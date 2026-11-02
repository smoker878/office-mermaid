/* global document, Image, btoa, TextEncoder, setTimeout, clearTimeout, crypto */
import type { DiagramPreview } from "../renderer/renderer";
import { escapeXml } from "../diagram/record";

export function utf8Base64(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return btoa(binary);
}

/** Real raster fallback for Office versions which cannot display SVG. No external requests. */
export function createPngFallback(preview: DiagramPreview): Promise<string> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const timer = setTimeout(
      () => finish(new Error("圖片準備逾時，尚未寫入文件，請重試。")),
      10000
    );
    function finish(error?: Error, result?: string) {
      clearTimeout(timer);
      image.onload = null;
      image.onerror = null;
      if (error) reject(error);
      else resolve(result!);
    }
    image.onerror = () => finish(new Error("無法解碼 SVG，尚未寫入文件。"));
    image.onload = () => {
      try {
        const scale = Math.min(
          2,
          4096 / Math.max(preview.width, preview.height),
          Math.sqrt(4000000 / (preview.width * preview.height))
        );
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(preview.width * scale));
        canvas.height = Math.max(1, Math.round(preview.height * scale));
        const context = canvas.getContext("2d");
        if (!context) throw new Error("此環境無法準備 Office 圖片。");
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        const data = canvas.toDataURL("image/png");
        if (!data.startsWith("data:image/png;base64,")) throw new Error("PNG 圖片產生失敗。");
        finish(undefined, data.split(",")[1]);
      } catch (error) {
        finish(error instanceof Error ? error : new Error("準備圖片失敗。"));
      }
    };
    image.src = `data:image/svg+xml;base64,${utf8Base64(preview.svg)}`;
  });
}

/** Flat OPC with embedded SVG + PNG, no links to the task pane or external resources. */
export function buildPictureOoxml(
  preview: DiagramPreview,
  png: string,
  widthCm: number,
  tag?: string
): string {
  const cx = Math.round(widthCm * 360000);
  const cy = Math.round((cx * preview.height) / preview.width);
  if (![cx, cy].every((v) => Number.isFinite(v) && v > 0 && v <= 20000000))
    throw new Error("文件圖片尺寸過大，請降低顯示寬度或調整圖表排版。");
  const name = `Mermaid-${crypto.randomUUID()}`;
  const drawingId = 1 + (crypto.getRandomValues(new Uint32Array(1))[0] % 2147483646);
  const drawing = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${drawingId}" name="${name}" descr="Mermaid 圖表"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>
    <pic:blipFill><a:blip r:embed="png"><a:extLst>
    <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="svg"/></a:ext>
    </a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
    </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const content = tag
    ? `<w:sdt><w:sdtPr><w:alias w:val="Office Mermaid"/><w:tag w:val="${escapeXml(tag)}"/>
      <w:id w:val="${crypto.getRandomValues(new Uint32Array(1))[0] % 2147483647}"/>
      </w:sdtPr><w:sdtContent>${drawing}</w:sdtContent></w:sdt>`
    : drawing;
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships></pkg:xmlData></pkg:part>
    <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
    <w:body><w:p>${content}</w:p></w:body></w:document></pkg:xmlData></pkg:part>
    <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="png" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/diagram.png"/>
    <Relationship Id="svg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/diagram.svg"/>
    </Relationships></pkg:xmlData></pkg:part>
    <pkg:part pkg:name="/word/media/diagram.png" pkg:contentType="image/png"><pkg:binaryData>${png}</pkg:binaryData></pkg:part>
    <pkg:part pkg:name="/word/media/diagram.svg" pkg:contentType="image/svg+xml"><pkg:binaryData>${utf8Base64(preview.svg)}</pkg:binaryData></pkg:part>
    </pkg:package>`;
}
