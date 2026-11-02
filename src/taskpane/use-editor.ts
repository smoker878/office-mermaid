/* global window, document, setTimeout, clearTimeout, HTMLTextAreaElement, BeforeUnloadEvent, KeyboardEvent */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import { validateDraft } from "../diagram/record";
import { DiagramTargetError, type DiagramTarget, type OfficeAdapter } from "../office/adapter";
import { renderDiagram } from "../renderer/renderer";
import {
  fingerprint,
  initialFields,
  toDraft,
  type EditorFields,
  type StatusMessage,
} from "./editor-state";
import { useDiagramPreview } from "./use-diagram-preview";
import { useOfficeHost } from "./use-office-host";

export function useEditor() {
  const [fields, setFields] = useState(initialFields);
  const [baseline, setBaseline] = useState(() => fingerprint(initialFields));
  const [target, setTarget] = useState<DiagramTarget | null>(null);
  const [selected, setSelected] = useState<DiagramTarget | null>(null);
  const [busy, setBusy] = useState(false);
  const [needsReload, setNeedsReload] = useState(false);
  const [preserveWidth, setPreserveWidth] = useState(true);
  const [styleOpen, setStyleOpen] = useState(false);
  const [sizeOpen, setSizeOpen] = useState(false);
  const [displayInvalid, setDisplayInvalid] = useState(false);
  const [officeError, setOfficeError] = useState("");
  const [officeStatus, setOfficeStatus] = useState<StatusMessage>({
    text: "文件操作尚未連線；仍可編輯與預覽。",
  });
  const sourceRef = useRef<HTMLTextAreaElement>(null);
  const focusAfterLoad = useRef(false);
  const operationLock = useRef(false);
  const active = useRef(true);
  const selectionVersion = useRef(0);
  const selectionTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const scanning = useRef(false);
  const hostRef = useRef<OfficeAdapter | null>(null);

  const selectionChanged = useCallback(function scanSelection() {
    const version = ++selectionVersion.current;
    setSelected(null);
    clearTimeout(selectionTimer.current);
    selectionTimer.current = setTimeout(async () => {
      const adapter = hostRef.current;
      if (!adapter || operationLock.current || !active.current) return;
      if (scanning.current) {
        scanSelection();
        return;
      }
      scanning.current = true;
      try {
        const candidate = await adapter.loadSelection();
        if (active.current && version === selectionVersion.current && !operationLock.current)
          setSelected(candidate);
      } catch {
        // Ordinary pictures and empty/multiple selections don't replace the current target.
      } finally {
        scanning.current = false;
      }
    }, 200);
  }, []);

  const host = useOfficeHost(selectionChanged);
  useLayoutEffect(() => {
    hostRef.current = host.adapter;
  }, [host.adapter]);
  useEffect(() => {
    if (!host.adapter) return;
    setOfficeStatus({
      text: "將游標放到空白位置後插入；要修改既有圖表，請選取後按上方「編輯選取圖表」。",
    });
    selectionChanged();
  }, [host.adapter, selectionChanged]);
  useEffect(() => {
    active.current = true;
    return () => {
      active.current = false;
      ++selectionVersion.current;
      clearTimeout(selectionTimer.current);
    };
  }, []);

  const draft = toDraft(fields);
  const preview = useDiagramPreview(draft, busy);
  useEffect(() => {
    if (!busy && focusAfterLoad.current) {
      focusAfterLoad.current = false;
      sourceRef.current?.focus();
    }
  }, [busy]);
  useEffect(() => {
    if (preview.error?.field === "css") setStyleOpen(true);
    if (preview.error?.field === "layout") setSizeOpen(true);
  }, [preview.error]);

  const dirty = fingerprint(fields) !== baseline;
  useEffect(() => {
    const beforeUnload = (event: BeforeUnloadEvent) => {
      if (operationLock.current || dirty) {
        event.preventDefault();
        event.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", beforeUnload);
    return () => window.removeEventListener("beforeunload", beforeUnload);
  }, [dirty]);
  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
        event.preventDefault();
        if (!operationLock.current) preview.refresh(toDraft(fields));
      }
    };
    document.addEventListener("keydown", keydown);
    return () => document.removeEventListener("keydown", keydown);
  }, [fields, preview.refresh]);

  function changeField<K extends keyof EditorFields>(key: K, value: EditorFields[K]) {
    if (operationLock.current) return;
    setFields((current) => ({ ...current, [key]: value }));
    if (key === "displayWidth") setDisplayInvalid(false);
    setOfficeStatus({
      text: needsReload
        ? "草稿已保留；請重新載入圖表確認目標後再更新。"
        : "有尚未寫入文件的變更；預覽更新不會自動修改文件。",
      state: "pending",
    });
  }

  async function operate(action: "insert" | "load" | "update", expected?: DiagramTarget) {
    const adapter = host.adapter;
    if (!adapter || operationLock.current || (action === "update" && (!target || needsReload)))
      return;
    if (
      action === "load" &&
      dirty &&
      !window.confirm("目前有尚未套用的修改。要放棄修改並編輯選取圖表嗎？取消可繼續目前的編輯。")
    )
      return;
    // A ref locks immediately, before React commits the disabled button state.
    operationLock.current = true;
    setBusy(true);
    ++selectionVersion.current;
    clearTimeout(selectionTimer.current);
    setOfficeError("");
    setOfficeStatus({
      text: { insert: "正在準備並插入圖表…", load: "正在載入選取圖表…", update: "正在更新原圖…" }[
        action
      ],
      state: "pending",
    });
    let loaded = false;
    try {
      let result: DiagramTarget;
      let nextFields = fields;
      if (action === "load") {
        result = await adapter.loadSelection();
        if (!active.current) return;
        if (expected && (result.controlId !== expected.controlId || result.tag !== expected.tag))
          throw new Error("選取的圖表已改變，請重新選取後按「編輯選取圖表」。目前草稿已保留。");
        // Validate before replacing a single character of the user's in-progress draft.
        renderDiagram(result.record.mermaidSource, result.record.cssSource, result.record.layout);
        nextFields = {
          mermaidSource: result.record.mermaidSource,
          cssSource: result.record.cssSource,
          layoutWidth: String(result.record.layout.maxWidth),
          autoDirection: result.record.layout.autoDirection,
          displayWidth: String(Math.round(result.displayWidthCm * 100) / 100),
        };
        setPreserveWidth(true);
        loaded = true;
      } else {
        const current = toDraft(fields);
        validateDraft(current);
        // Never write a stale preview, even when invoked in the debounce interval.
        const rendered = renderDiagram(current.mermaidSource, current.cssSource, current.layout);
        result =
          action === "insert"
            ? await adapter.insert(current, rendered)
            : await adapter.update(target!, current, rendered, preserveWidth);
        if (!active.current) return;
        nextFields = {
          ...fields,
          displayWidth: String(Math.round(result.displayWidthCm * 100) / 100),
        };
      }
      setFields(nextFields);
      setTarget(result);
      setNeedsReload(false);
      setDisplayInvalid(false);
      setBaseline(fingerprint(nextFields));
      preview.refresh(toDraft(nextFields));
      setOfficeStatus({
        text:
          action === "load"
            ? "圖表已載入；修改後按「更新這張圖表」。移動游標不會改變編輯目標。"
            : "圖表與原始碼已寫入目前文件；請在 Word 儲存文件。",
        state: "success",
      });
    } catch (error) {
      if (!active.current) return;
      setOfficeError(
        error instanceof Error
          ? error.message
          : "文件操作失敗。請確認文件可編輯、Word 版本支援，並重新載入圖表。"
      );
      setOfficeStatus({ text: "操作未完成；編輯內容已保留。", state: "error" });
      if (action === "update" && error instanceof DiagramTargetError) setNeedsReload(true);
      const width = toDraft(fields).displayWidthCm;
      if (!Number.isFinite(width) || width < 1 || width > 50) {
        setDisplayInvalid(true);
        setSizeOpen(true);
      }
      if (!preview.valid) preview.refresh(toDraft(fields));
    } finally {
      operationLock.current = false;
      if (active.current) {
        setBusy(false);
        focusAfterLoad.current = loaded;
        selectionChanged();
      }
    }
  }

  function newDiagram() {
    if (operationLock.current) return;
    setTarget(null);
    setNeedsReload(false);
    setOfficeError("");
    setOfficeStatus({
      text: "新增圖表：已保留目前內容，請將 Word 游標放到空白插入位置，再按「插入文件」。",
      state: "pending",
    });
    sourceRef.current?.focus();
  }

  return {
    fields,
    changeField,
    sourceRef,
    preview,
    host,
    target,
    selected,
    busy,
    needsReload,
    preserveWidth,
    setPreserveWidth,
    styleOpen,
    setStyleOpen,
    sizeOpen,
    setSizeOpen,
    displayInvalid,
    officeError,
    officeStatus,
    operate,
    newDiagram,
    refresh: () => preview.refresh(toDraft(fields)),
    canWrite: !!host.adapter && !busy && preview.valid && !needsReload,
  };
}

export type EditorController = ReturnType<typeof useEditor>;
