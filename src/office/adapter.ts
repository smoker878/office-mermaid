import type { DiagramDraft, DiagramRecord } from "../diagram/record";
import type { DiagramPreview } from "../renderer/renderer";

/** The editor must reload the document target before attempting another update. */
export class DiagramTargetError extends Error {}

export interface DiagramTarget {
  controlId: number;
  tag: string;
  record: DiagramRecord;
  displayWidthCm: number;
}

export interface OfficeAdapter {
  insert(draft: DiagramDraft, preview: DiagramPreview): Promise<DiagramTarget>;
  loadSelection(): Promise<DiagramTarget>;
  update(
    target: DiagramTarget,
    draft: DiagramDraft,
    preview: DiagramPreview,
    preserveWidth: boolean
  ): Promise<DiagramTarget>;
}
