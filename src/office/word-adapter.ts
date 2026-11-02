/* global Word, Office */
/* These rules don't follow our helper syncs, ClientResult.value, navigation proxies or recovery context.
   Scalar properties are explicitly loaded and synced; navigation proxies need no load themselves. */
/* eslint-disable office-addins/load-object-before-read, office-addins/call-sync-before-read, office-addins/call-sync-after-load, office-addins/no-navigational-load */
import {
  createRecord,
  parseRecord,
  partIdFromTag,
  serializeRecord,
  tagPrefix,
  validateDraft,
  type DiagramDraft,
  type DiagramRecord,
} from "../diagram/record";
import type { DiagramPreview } from "../renderer/renderer";
import type { DiagramTarget, OfficeAdapter } from "./adapter";
import { DiagramTargetError } from "./adapter";
import { buildPictureOoxml, createPngFallback } from "./ooxml";

const pointsPerCm = 72 / 2.54;

export function wordCapability(): string | null {
  if (typeof Office === "undefined" || typeof Word === "undefined")
    return "文件操作需要 Word 宿主。";
  if (!Office.context.requirements.isSetSupported("WordApi", "1.4"))
    return "此 Word 不支援 WordApi 1.4；請更新 Office。編輯與預覽仍可使用。";
  return null;
}

async function readRecord(context: Word.RequestContext, tag: string): Promise<DiagramRecord> {
  const part = context.document.customXmlParts.getItemOrNullObject(partIdFromTag(tag));
  part.load("id");
  await context.sync();
  if (part.isNullObject) throw new Error("圖表原始資料已遺失，無法重新編輯；文件圖片不受影響。");
  const xml = part.getXml();
  await context.sync();
  return parseRecord(xml.value);
}

async function stageRecord(context: Word.RequestContext, record: DiagramRecord): Promise<string> {
  // Immutable snapshots: never overwrite source used by another pasted copy or by Undo.
  const part = context.document.customXmlParts.add(serializeRecord(record));
  part.load("id");
  await context.sync();
  return tagPrefix + part.id;
}

async function inspectPicture(context: Word.RequestContext, control: Word.ContentControl) {
  control.load("id,tag,text,cannotEdit,cannotDelete");
  const pictures = control.inlinePictures;
  pictures.load("items/width,items/height");
  const children = control.contentControls;
  children.load("items/id");
  await context.sync();
  if (control.isNullObject) throw new Error("原圖已被刪除，請重新插入圖表。");
  if (pictures.items.length !== 1 || control.text.trim() || children.items.length)
    throw new Error("圖表容器已被修改或包含其他內容，為避免覆蓋，請先還原為單張圖表。");
  return pictures.items[0];
}

async function requireWritable(context: Word.RequestContext) {
  context.document.load("changeTrackingMode");
  await context.sync();
  if (context.document.changeTrackingMode !== "Off")
    throw new Error("請先關閉 Word 的追蹤修訂，再插入或更新圖表，避免留下多份修訂圖片。");
}

export class WordAdapter implements OfficeAdapter {
  async insert(draft: DiagramDraft, preview: DiagramPreview): Promise<DiagramTarget> {
    validateDraft(draft);
    const png = await createPngFallback(preview);
    const record = createRecord(draft);
    // Validate dimensions before making any document writes.
    buildPictureOoxml(preview, png, draft.displayWidthCm);
    return Word.run(async (context) => {
      await requireWritable(context);
      const selection = context.document.getSelection();
      selection.load("isEmpty");
      const parent = selection.parentContentControlOrNullObject;
      parent.load("id");
      await context.sync();
      if (!selection.isEmpty || !parent.isNullObject)
        throw new Error("請將游標放在圖表／內容控制項以外的空白插入位置，不要選取文字或圖片。");
      const tag = await stageRecord(context, record);
      try {
        selection.insertOoxml(
          buildPictureOoxml(preview, png, draft.displayWidthCm, tag),
          "Replace"
        );
        await context.sync();
        const controls = context.document.contentControls.getByTag(tag);
        controls.load("items/id");
        await context.sync();
        if (controls.items.length !== 1) throw new Error("插入結果無法識別。");
        const picture = await inspectPicture(context, controls.items[0]);
        return {
          controlId: controls.items[0].id,
          tag,
          record,
          displayWidthCm: picture.width / pointsPerCm,
        };
      } catch {
        // Do not retry automatically: the host may have applied the insertion before an error.
        throw new Error(
          "無法確認插入完成。請先檢查文件是否已有圖表；若有，選取後載入，避免重複插入。原始碼快照已保留於文件。"
        );
      }
    });
  }

  async loadSelection(): Promise<DiagramTarget> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      const contained = selection.contentControls;
      contained.load("items/id,items/tag");
      const parent = selection.parentContentControlOrNullObject;
      parent.load("id,tag");
      await context.sync();
      const candidates = contained.items.filter((control) => control.tag.startsWith(tagPrefix));
      if (
        !parent.isNullObject &&
        parent.tag.startsWith(tagPrefix) &&
        !candidates.some((control) => control.id === parent.id)
      )
        candidates.push(parent);
      if (candidates.length !== 1)
        throw new Error(
          candidates.length > 1
            ? "請只選取一張圖表。"
            : "請選取 Office Mermaid 圖表（或將游標放入圖表容器），再按「載入選取圖表」。"
        );
      const control = candidates[0];
      const picture = await inspectPicture(context, control);
      const record = await readRecord(context, control.tag);
      return {
        controlId: control.id,
        tag: control.tag,
        record,
        displayWidthCm: picture.width / pointsPerCm,
      };
    });
  }

  async update(
    target: DiagramTarget,
    draft: DiagramDraft,
    preview: DiagramPreview,
    preserveWidth: boolean
  ): Promise<DiagramTarget> {
    validateDraft(draft);
    const png = await createPngFallback(preview);
    return Word.run(async (context) => {
      await requireWritable(context);
      const control = context.document.contentControls.getByIdOrNullObject(target.controlId);
      control.load("id,tag");
      await context.sync();
      if (control.isNullObject) throw new DiagramTargetError("原圖已被刪除，請重新插入圖表。");
      if (control.tag !== target.tag)
        throw new DiagramTargetError("圖表已被其他操作更新，請重新載入後再編輯。");
      const picture = await inspectPicture(context, control);
      if (control.cannotEdit || control.cannotDelete)
        throw new Error("此圖表已鎖定，請先在 Word 解除鎖定再更新。");
      const saved = await readRecord(context, control.tag);
      if (serializeRecord(saved) !== serializeRecord(target.record))
        throw new DiagramTargetError("原始資料已變更，請重新載入後再編輯。");
      const copies = context.document.contentControls.getByTag(target.tag);
      copies.load("items/id");
      const backup = control.getRange("Content").getOoxml();
      await context.sync();
      const width = preserveWidth ? picture.width / pointsPerCm : draft.displayWidthCm;
      const record = createRecord(
        { ...draft, displayWidthCm: width },
        copies.items.length > 1 ? undefined : saved.id
      );
      const ooxml = buildPictureOoxml(preview, png, width);
      const tag = await stageRecord(context, record);
      try {
        control.insertOoxml(ooxml, "Replace");
        control.tag = tag;
        await context.sync();
        const updated = await inspectPicture(context, control);
        if (control.tag !== tag) throw new Error("圖表資料關聯更新失敗。");
        return { controlId: control.id, tag, record, displayWidthCm: updated.width / pointsPerCm };
      } catch {
        // Office batches are NOT transactions. Recover the previous content and tag explicitly.
        try {
          await Word.run(async (recovery) => {
            const original = recovery.document.contentControls.getByIdOrNullObject(
              target.controlId
            );
            original.load("id,tag");
            await recovery.sync();
            if (original.isNullObject || (original.tag !== tag && original.tag !== target.tag))
              throw new Error("無法安全還原目標。");
            original.insertOoxml(backup.value, "Replace");
            original.tag = target.tag;
            await recovery.sync();
          });
        } catch {
          throw new DiagramTargetError(
            "更新失敗，且無法確認原圖已還原。請先檢查文件並使用 Word 復原；不要直接重試。新舊原始碼快照均保留。"
          );
        }
        throw new DiagramTargetError(
          "更新失敗，已嘗試還原原圖與原始資料。請重新載入圖表確認後再重試。"
        );
      }
    });
  }
}
