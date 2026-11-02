/* global document, window, DOMParser, atob, TextDecoder */
import {
  createRecord,
  parseRecord,
  serializeRecord,
  defaultLayout,
  partIdFromTag,
  type DiagramDraft,
} from "../src/diagram/record";
import { renderDiagram } from "../src/renderer/renderer";
import { buildPictureOoxml, createPngFallback } from "../src/office/ooxml";
import { WordAdapter, wordCapability } from "../src/office/word-adapter";

function assert(condition: unknown, message: string) {
  if (!condition) throw new Error(message);
}
async function rejects(action: () => unknown | Promise<unknown>) {
  let rejected = false;
  try {
    await action();
  } catch {
    rejected = true;
  }
  assert(rejected, "應拒絕此操作");
}

const draft: DiagramDraft = {
  mermaidSource: "flowchart LR\nA[中文 & 測試] --> B[完成]",
  cssSource: ".pele { --pele-bg: #fff; }",
  layout: { ...defaultLayout },
  displayWidthCm: 14,
};

/** Small Office API contract fake, not a substitute for an actual Word host. */
class FakeWord {
  parts = new Map<string, string>();
  controls = new Map<number, FakeControl>();
  selected: number[] = [];
  emptySelection = true;
  changeTrackingMode = "Off";
  failPicture = false;
  failRecovery = false;
  nextId = 1;
  queue: (() => void)[] = [];
  context = {
    document: {
      load: () => {},
      get changeTrackingMode() {
        return "Off";
      },
      getSelection: () => ({
        load: () => {},
        isEmpty: this.emptySelection,
        parentContentControlOrNullObject:
          this.selected.length === 1
            ? this.controls.get(this.selected[0])
            : { load: () => {}, isNullObject: true },
        contentControls: this.collection(this.selected.map((id) => this.controls.get(id)!)),
        insertOoxml: (xml: string) =>
          this.queue.push(() => {
            const doc = new DOMParser().parseFromString(xml, "application/xml");
            const tag = doc.getElementsByTagNameNS("*", "tag")[0].getAttribute("w:val")!;
            const control = new FakeControl(this, this.nextId++, tag, xml);
            this.controls.set(control.id, control);
          }),
      }),
      customXmlParts: {
        add: (xml: string) => {
          const id = window.crypto.randomUUID();
          this.queue.push(() => this.parts.set(id, xml));
          return { id, load: () => {} };
        },
        getItemOrNullObject: (id: string) => ({
          load: () => {},
          isNullObject: !this.parts.has(id),
          getXml: () => ({ value: this.parts.get(id) }),
        }),
      },
      contentControls: {
        getByTag: (tag: string) =>
          this.collection([...this.controls.values()].filter((c) => c.tag === tag)),
        getByIdOrNullObject: (id: number) =>
          this.controls.get(id) || { load: () => {}, isNullObject: true },
      },
    },
    sync: async () => {
      const pending = this.queue.splice(0);
      for (const action of pending) action();
    },
  };

  constructor() {
    Object.defineProperty(this.context.document, "changeTrackingMode", {
      get: () => this.changeTrackingMode,
    });
  }
  collection(items: unknown[]) {
    return { items, load: () => {} };
  }
  install() {
    (window as unknown as { Word: unknown }).Word = {
      run: (callback: (context: unknown) => unknown) => callback(this.context),
    };
  }
}

class FakeControl {
  isNullObject = false;
  cannotEdit = false;
  cannotDelete = false;
  text = "";
  width = (14 * 72) / 2.54;
  height = 80;
  children: FakeControl[] = [];
  constructor(
    public host: FakeWord,
    public id: number,
    private currentTag: string,
    public xml: string
  ) {}
  get tag() {
    return this.currentTag;
  }
  set tag(value: string) {
    this.host.queue.push(() => {
      this.currentTag = value;
    });
  }
  load() {}
  get inlinePictures() {
    return this.host.collection([{ width: this.width, height: this.height }]);
  }
  get contentControls() {
    return this.host.collection(this.children);
  }
  getRange() {
    return { getOoxml: () => ({ value: this.xml }) };
  }
  insertOoxml(xml: string) {
    this.host.queue.push(() => {
      if (this.host.failRecovery) throw new Error("simulated recovery failure");
      this.xml = xml;
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const extent = doc.getElementsByTagNameNS(
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "extent"
      )[0];
      this.width = Number(extent.getAttribute("cx")) / 12700;
      if (this.host.failPicture) {
        this.host.failPicture = false;
        throw new Error("simulated partial host write");
      }
    });
  }
}

async function run() {
  const results: string[] = [];
  let failures = 0;
  async function test(name: string, action: () => unknown | Promise<unknown>) {
    try {
      await action();
      results.push(`PASS ${name}`);
    } catch (error) {
      failures++;
      results.push(`FAIL ${name}: ${error instanceof Error ? error.message : error}`);
    }
  }
  await test("原始碼、CSS、排版、版本可完整 XML 往返", () => {
    const record = createRecord(draft);
    assert(
      JSON.stringify(parseRecord(serializeRecord(record))) === JSON.stringify(record),
      "資料往返"
    );
  });
  await test("拒絕未知版本、不合法 XML、外部實體與過大資料", async () => {
    const xml = serializeRecord(createRecord(draft));
    for (const value of [
      xml.replace("&quot;schemaVersion&quot;:1", "&quot;schemaVersion&quot;:2"),
      "<broken",
      '<!DOCTYPE diagram SYSTEM "https://example.com/x">' + xml,
      "x".repeat(500001),
    ])
      await rejects(() => parseRecord(value));
  });
  await test("拒絕無效尺寸、空原始碼與非外掛 tag", async () => {
    for (const width of [NaN, 0, 199, 4001, 800.5])
      await rejects(() =>
        createRecord({ ...draft, layout: { maxWidth: width, autoDirection: false } })
      );
    await rejects(() => createRecord({ ...draft, displayWidthCm: 0 }));
    await rejects(() => createRecord({ ...draft, mermaidSource: "" }));
    await rejects(() => partIdFromTag("other"));
  });
  await test("排版設定無效回傳 layout 欄位錯誤", async () => {
    await rejects(() =>
      renderDiagram(draft.mermaidSource, "", { maxWidth: 0, autoDirection: false })
    );
  });
  await test("窄版自動方向會重排，關閉則保留原方向", () => {
    const source =
      "flowchart LR\nA[開始流程] --> B[讀取資料] --> C[驗證資料] --> D[執行處理] --> E[完成流程]";
    const vertical = renderDiagram(source, "", { maxWidth: 300, autoDirection: true });
    const horizontal = renderDiagram(source, "", { maxWidth: 300, autoDirection: false });
    assert(
      vertical.height > horizontal.height && vertical.svg !== horizontal.svg,
      `應重新排版而非僅縮放：${vertical.width}×${vertical.height} / ${horizontal.width}×${horizontal.height}`
    );
  });
  const preview = renderDiagram(draft.mermaidSource, draft.cssSource, draft.layout);
  await test("PNG fallback 與 SVG 同時嵌入有效 OOXML，含中文", async () => {
    const png = await createPngFallback(preview);
    assert(atob(png).slice(1, 4) === "PNG", "真實 PNG");
    const xml = buildPictureOoxml(preview, png, 14, "office-mermaid:test");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    assert(!doc.querySelector("parsererror"), "有效 XML");
    const binaries = doc.getElementsByTagNameNS("*", "binaryData");
    const bytes = Uint8Array.from(atob(binaries[1].textContent!), (c) => c.charCodeAt(0));
    assert(new TextDecoder().decode(bytes) === preview.svg, "SVG UTF-8 正確");
    assert(xml.includes("SVG/main") && xml.includes('r:embed="png"'), "SVG 與 fallback 關聯");
    await rejects(() => buildPictureOoxml(preview, png, Infinity));
  });
  await test("無 Word 環境顯示能力限制", () => assert(!!wordCapability(), "不可啟用操作"));

  const host = new FakeWord();
  host.install();
  const adapter = new WordAdapter();
  let target = await adapter.insert(draft, preview);
  await test("插入建立圖表容器與文件內快照", () => {
    assert(host.controls.size === 1 && host.parts.size === 1, "圖表與原始碼");
    assert(target.record.layout.maxWidth === 800, "排版已保存");
  });
  await test("重新建立 adapter 可由文件載入（非側邊欄記憶體）", async () => {
    host.selected = [target.controlId];
    host.emptySelection = false;
    const loaded = await new WordAdapter().loadSelection();
    assert(
      loaded.record.mermaidSource === draft.mermaidSource && loaded.tag === target.tag,
      "讀回文件資料"
    );
  });
  await test("沒有選取與多選均拒絕且不改文件", async () => {
    host.selected = [];
    await rejects(() => adapter.loadSelection());
    const copy = new FakeControl(
      host,
      host.nextId++,
      target.tag,
      host.controls.get(target.controlId)!.xml
    );
    host.controls.set(copy.id, copy);
    host.selected = [target.controlId, copy.id];
    await rejects(() => adapter.loadSelection());
  });
  await test("選取文字／容器時拒絕新增，不誤刪內容", async () => {
    const before = host.parts.size;
    await rejects(() => adapter.insert(draft, preview));
    assert(host.parts.size === before, "不寫入資料");
  });
  await test("更新鎖定已載入目標；複製副本不受影響且取得獨立 ID", async () => {
    const copy = [...host.controls.values()].find((c) => c.id !== target.controlId)!;
    host.selected = [copy.id];
    const previousId = target.record.id;
    host.controls.get(target.controlId)!.width = (10 * 72) / 2.54;
    const changed = { ...draft, mermaidSource: "flowchart LR\nA --> C", displayWidthCm: 20 };
    target = await adapter.update(
      target,
      changed,
      renderDiagram(changed.mermaidSource, changed.cssSource, changed.layout),
      true
    );
    assert(target.controlId !== copy.id && copy.tag !== target.tag, "只更新原目標");
    assert(
      target.record.id !== previousId && Math.abs(target.displayWidthCm - 10) < 0.01,
      "獨立 ID 與目前寬度"
    );
    const old = parseRecord(host.parts.get(partIdFromTag(copy.tag))!);
    assert(old.mermaidSource === draft.mermaidSource, "副本保留原始資料");
  });
  await test("指定新寬度可縮放文件圖片，不改變排版設定", async () => {
    target = await adapter.update(target, { ...draft, displayWidthCm: 12 }, preview, false);
    assert(
      Math.abs(target.displayWidthCm - 12) < 0.01 && target.record.layout.maxWidth === 800,
      "縮放與排版分離"
    );
  });
  await test("過時目標拒絕更新", async () => {
    await rejects(() =>
      adapter.update({ ...target, tag: "office-mermaid:stale" }, draft, preview, true)
    );
  });
  await test("鎖定圖表拒絕更新且不建立新快照", async () => {
    const control = host.controls.get(target.controlId)!;
    const before = host.parts.size;
    control.cannotEdit = true;
    await rejects(() => adapter.update(target, draft, preview, true));
    control.cannotEdit = false;
    assert(host.parts.size === before, "鎖定時不寫入");
  });
  await test("容器新增其他文字時拒絕覆蓋", async () => {
    const control = host.controls.get(target.controlId)!;
    control.text = "使用者新增的內容";
    await rejects(() => adapter.update(target, draft, preview, true));
    control.text = "";
  });
  await test("追蹤修訂時不寫入", async () => {
    host.changeTrackingMode = "TrackAll";
    await rejects(() => adapter.update(target, draft, preview, true));
    host.changeTrackingMode = "Off";
  });
  await test("部分更新失敗會執行原內容與 tag 還原，舊快照仍保留", async () => {
    const before = host.controls.get(target.controlId)!.xml;
    host.failPicture = true;
    await rejects(() => adapter.update(target, draft, preview, false));
    assert(
      host.controls.get(target.controlId)!.xml === before &&
        host.controls.get(target.controlId)!.tag === target.tag,
      "還原完成"
    );
    assert(host.parts.has(partIdFromTag(target.tag)), "舊資料不刪除");
  });
  await test("無法還原時清楚警告，不假稱更新成功", async () => {
    host.failRecovery = true;
    try {
      await adapter.update(target, draft, preview, false);
      throw new Error("未拒絕失敗操作");
    } catch (error) {
      assert(error instanceof Error && error.message.includes("無法確認原圖已還原"), "明確警告");
    } finally {
      host.failRecovery = false;
    }
  });
  await test("資料遺失仍保留圖片並拒絕載入", async () => {
    host.selected = [target.controlId];
    host.parts.delete(partIdFromTag(target.tag));
    await rejects(() => adapter.loadSelection());
    assert(host.controls.has(target.controlId), "圖片仍存在");
  });
  await test("目標刪除後拒絕更新，不自動新增", async () => {
    host.controls.delete(target.controlId);
    const before = host.controls.size;
    await rejects(() => adapter.update(target, draft, preview, true));
    assert(host.controls.size === before, "不新增圖片");
  });
  document.getElementById("test-results")!.textContent =
    `${results.join("\n")}\n\n${results.length - failures}/${results.length} passed\n（Office API 模擬測試，不代表 Word 實機驗收）`;
  document.body.dataset.testStatus = failures ? "failed" : "passed";
}

void run().catch((error) => {
  document.getElementById("test-results")!.textContent = String(error);
  document.body.dataset.testStatus = "failed";
});
