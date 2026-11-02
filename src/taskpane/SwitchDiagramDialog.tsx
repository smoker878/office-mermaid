/* global HTMLDialogElement, HTMLButtonElement */
import { useEffect, useRef } from "react";
import type { EditorController } from "./use-editor";
import { button, primaryButton } from "./styles";

export function SwitchDiagramDialog({ editor }: { editor: EditorController }) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const cancelRef = useRef<HTMLButtonElement>(null);
  const open = !!editor.switchRequest;

  useEffect(() => {
    if (!open) return;
    const dialog = dialogRef.current!;
    dialog.showModal();
    cancelRef.current?.focus();
    return () => dialog.close();
  }, [open]);

  function cancel() {
    dialogRef.current?.close();
    editor.cancelSwitch();
    editor.sourceRef.current?.focus();
  }

  return (
    <dialog
      id="switch-diagram-dialog"
      ref={dialogRef}
      className="m-auto max-h-[85vh] w-[calc(100%-2rem)] max-w-100 overflow-y-auto rounded-panel border border-border bg-surface p-4 text-ink shadow-action backdrop:bg-dialog-backdrop min-[421px]:p-6"
      aria-labelledby="switch-diagram-title"
      aria-describedby="switch-diagram-description"
      onCancel={(event) => {
        event.preventDefault();
        cancel();
      }}
    >
      <h2 id="switch-diagram-title" className="text-heading font-bold">
        要切換編輯另一張圖嗎？
      </h2>
      <p id="switch-diagram-description" className="mt-3">
        {editor.dirty
          ? "目前有尚未套用的修改，切換後將捨棄這些修改。文件中的原圖不會改變。"
          : "切換後會載入你剛選取的圖表，文件中的原圖不會改變。"}
      </p>
      <div className="mt-6 flex flex-col gap-3">
        <button
          id="cancel-switch"
          ref={cancelRef}
          type="button"
          className={button}
          onClick={cancel}
        >
          繼續編輯目前圖表
        </button>
        <button
          id="confirm-switch"
          type="button"
          className={primaryButton}
          onClick={editor.confirmSwitch}
        >
          {editor.dirty ? "捨棄修改並切換" : "切換編輯這張圖"}
        </button>
      </div>
    </dialog>
  );
}
