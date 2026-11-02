/* global window, document, Office */
import { useEffect, useState } from "react";
import type { OfficeAdapter } from "../office/adapter";
import { WordAdapter, wordCapability } from "../office/word-adapter";

const browserStatus = "瀏覽器預覽模式；文件操作需要 Office 宿主。";
let officeScript: Promise<void> | undefined;

function loadOffice(): Promise<void> {
  if (!officeScript) {
    officeScript = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
      script.async = true;
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("Office.js 無法連線；編輯器與預覽仍可使用。"));
      document.head.appendChild(script);
    });
  }
  return officeScript;
}

export function useOfficeHost(selectionChanged: () => void) {
  const [adapter, setAdapter] = useState<OfficeAdapter | null>(null);
  const [hostStatus, setHostStatus] = useState(browserStatus);
  const [selectionHelp, setSelectionHelp] = useState("");

  useEffect(() => {
    // Never request Office.js from browser preview, including after a StrictMode remount.
    if (window.location.pathname.endsWith("preview.html")) return;
    let active = true;
    let registered = false;
    const removeHandler = () => {
      if (registered) {
        registered = false;
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
          handler: selectionChanged,
        });
      }
    };
    const fallback = () => {
      if (active) setSelectionHelp("選圖提示無法啟用；請選取圖表後，按上方「編輯選取圖表」。");
    };
    const dispose = () => {
      active = false;
      removeHandler();
    };
    window.addEventListener("unload", dispose);

    void loadOffice()
      .then(() => {
        if (!active || typeof Office === "undefined") return;
        Office.onReady((info) => {
          if (!active) return;
          if (info.host !== Office.HostType.Word) {
            setHostStatus(
              info.host ? `${info.host} 尚未支援文件操作；仍可編輯與預覽。` : browserStatus
            );
            return;
          }
          const reason = wordCapability();
          setHostStatus(reason || "Word 已連線 · 可插入、載入與更新圖表（尚待實機驗收）。");
          if (reason) return;
          setAdapter(new WordAdapter());
          try {
            Office.context.document.addHandlerAsync(
              Office.EventType.DocumentSelectionChanged,
              selectionChanged,
              (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) fallback();
                else {
                  registered = true;
                  // Registration may finish after the component has unmounted.
                  if (!active) removeHandler();
                }
              }
            );
          } catch {
            fallback();
          }
        });
      })
      .catch((error: Error) => {
        if (active) setHostStatus(error.message);
      });

    return () => {
      window.removeEventListener("unload", dispose);
      dispose();
    };
  }, [selectionChanged]);

  return { adapter, hostStatus, selectionHelp };
}
