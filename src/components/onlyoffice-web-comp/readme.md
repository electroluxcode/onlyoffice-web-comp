# OnlyOffice Comp Documentation

> 📖 English | [中文](readme.zh.md)

OnlyOffice Comp is a browser-side wrapper around the OnlyOffice Document Server static SDK. It supports online editing, read-only preview, export, and x2t conversion for Word, Excel, and PowerPoint. **No self-hosted Document Server is required**—only static SDK assets on your site.

## Table of Contents

- [Architecture & Entry Point](#architecture--entry-point)
- [Prerequisites](#prerequisites)
- [Quick Start](#quick-start)
- [OnlyOfficeManager (Recommended)](#onlyofficemanager-recommended)
- [EditorManager (Low-Level)](#editormanager-low-level)
- [Event System](#event-system)
- [Export & Save](#export--save)
- [Comments & Revisions](#comments--revisions)
- [Word Custom Fonts](#word-custom-fonts)
- [Constants & Store](#constants--store)
- [Utilities](#utilities)
- [Multi-Instance](#multi-instance)
- [Supported Formats](#supported-formats)
- [Notes](#notes)

---

## Architecture & Entry Point

```
onlyoffice-comp/
├── const/          Constants, static resource paths, file types
├── store/          Document / language cross-page state
├── util/           SDK init, x2t conversion, download helpers
├── core/           EditorManager, OnlyOfficeManager, EventBus
├── feature/        Comments, revisions, Word fonts
└── internal/       Mock server / x2t worker (not exported)
```

**Import everything from the package entry:**

```typescript
import {
  OnlyOfficeManager,
  EditorManager,
  editorManager,
  editorManagerFactory,
  createEditorView,
  onlyofficeEventbus,
  ONLYOFFICE_EVENT_KEYS,
  FILE_TYPE,
  ONLYOFFICE_ID,
} from "@/onlyoffice-comp";
```

Types are exported from `type/word-api.ts` and `type/sdk-internal.ts`.

---

## Prerequisites

1. **Static assets**: Host the OnlyOffice SDK (including `web-apps/`, `sdkjs/`, `fonts/`, `x2t/`) under a public path, default `public/9.3.0/`.
2. **Env var** (optional): `NEXT_PUBLIC_APP_ROOT=/9.3.0` — must match `STATIC_RESOURCE.onlyoffice.root`.
3. **x2t Brotli**: `x2t/x2t.js` and `x2t.wasm` are Brotli-precompressed on disk. **No** `Content-Encoding: br` headers are required; the x2t worker decompresses them via built-in `fetch-brotli`.
4. **DOM container**: Reserve a mount point in your page (see examples below).

---

## Quick Start

### 1. Page container

```tsx
import { ONLYOFFICE_ID, ONLYOFFICE_CONTAINER_CONFIG } from "@/onlyoffice-comp";

export function EditorHost() {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} relative flex-1`}
    >
      <div id={ONLYOFFICE_ID} className="absolute inset-0" />
    </div>
  );
}
```

### 2. Create editor (recommended: `OnlyOfficeManager`)

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
} from "@/onlyoffice-comp";

const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New Document.docx",
  readOnly: false,
});

await manager.openFile(file);
await manager.downloadExport();
```

### 3. Create editor (low-level: `createEditorView`)

For full control without the document store:

```typescript
import { createEditorView, initializeOnlyOffice } from "@/onlyoffice-comp";

await initializeOnlyOffice(); // also called inside createEditorView

await createEditorView({
  isNew: false,
  fileName: "report.docx",
  file,
  readOnly: false,
  lang: "zh",
  containerId: ONLYOFFICE_ID,
});
```

---

## OnlyOfficeManager (Recommended)

High-level facade for pages: init, open, export, read-only, language.

### Construction & factories

| API | Description |
|-----|-------------|
| `OnlyOfficeManager.create(options)` | Load DocsAPI and open `defaultFileName` as new doc |
| `OnlyOfficeManager.fromEditor(editor, options)` | Wrap existing `EditorManager` without auto-open |
| `onlyOfficeManagerFactory.open(options, document)` | Multi-container: cache manager by `containerId` |
| `getDefaultOnlyOfficeManager()` / `setDefaultOnlyOfficeManager()` | Optional single-page default |

**`OnlyOfficeManagerOptions`**

| Field | Type | Description |
|-------|------|-------------|
| `containerId` | `string?` | DOM id, default `ONLYOFFICE_ID` |
| `fileType` | `FileType` | Export format: `FILE_TYPE.DOCX` / `XLSX` / `PPTX` |
| `defaultFileName` | `string` | Initial document name |
| `readOnly` | `boolean?` | Initial read-only mode |
| `lang` | `'zh' \| 'en'?` | UI language |

**`OpenDocumentInput`**

| Field | Description |
|-------|-------------|
| `fileName` | File name with extension |
| `file` | Local `File` |
| `isNew` | Create blank document |
| `readOnly` | Override facade read-only flag |

### Instance methods

| Method | Description |
|--------|-------------|
| `openDocument(input)` | Open/switch document (updates document store) |
| `openNew(fileName, readOnly?)` | New document |
| `openFile(file, readOnly?)` | Open local file |
| `isReady()` | Whether open succeeded |
| `getReadOnly()` / `setReadOnly()` / `toggleReadOnly()` | Toggle edit/view via `asc_setRestriction` (no full reload) |
| `getLanguage()` / `setLanguage()` / `toggleLanguage()` | zh/en (remounts iframe) |
| `getEditor()` | Underlying `EditorManager` |
| `exportDocument()` | Export `Editor.bin` snapshot |
| `exportAsBlob()` | x2t → Office file as `Blob` |
| `downloadExport()` | `exportAsBlob` + browser download |
| `onLoadingChange(handler)` | Subscribe to `LOADING_CHANGE`, returns unsubscribe |
| `destroy()` | Tear down editor |

---

## EditorManager (Low-Level)

One `EditorManager` per `containerId`: mock server, iframe proxies, DocsAPI lifecycle.

### Factory

```typescript
import { editorManager, editorManagerFactory } from "@/onlyoffice-comp";

editorManagerFactory.getDefault();
editorManagerFactory.get("editor-1");
editorManagerFactory.create("editor-1");
editorManagerFactory.getAll();
editorManagerFactory.destroy("editor-1");
editorManagerFactory.destroyAll();
```

### `createEditorView(options)`

Same as `manager.create(options)`, returns `Promise<EditorManager>`.

**`CreateEditorViewOptions`**

| Field | Description |
|-------|-------------|
| `isNew` | Create new document |
| `fileName` | File name |
| `file` | Local file |
| `url` | Remote URL |
| `loader` | Custom `(url) => ArrayBuffer` |
| `fileType` | Extension; inferred from `fileName` if omitted |
| `readOnly` | Read-only mode |
| `lang` | `'zh' \| 'en'` |
| `containerId` | Container id |
| `editorManager` | Existing instance |
| `theme` | `OfficeTheme`, e.g. `'theme-white'` |
| `plugins` | `'featured' \| 'all' \| 'none'` |

### Instance methods

| Category | Methods |
|----------|---------|
| Lifecycle | `create()`, `exists()`, `destroy()` |
| Export | `export()` |
| Read-only / lang | `setReadOnly()`, `getReadOnly()`, `setLanguage()` |
| State | `getInstanceId()`, `getContainerId()`, `getFileName()`, `isDirty()` |
| Container | `getContainerParentSelector()`, `getContainerStyle()` |
| Media | `updateMedia(key, data)`, `getMedia()` |
| SDK subscribe | `subscribe({ type, fn })` |
| Comments | `getAllComments()`, `addComment()`, `updateComment()`, `removeComment()`, `goToComment()`, `registerCommentCallbacks()` |
| Revisions | `setTrackRevisions()`, `isTrackRevisions()`, `getAllRevisions()`, `acceptRevision()`, `rejectRevision()`, `acceptAllRevisions()`, `rejectAllRevisions()`, `registerRevisionCallbacks()` |

**Default review policy**: `review: true`, `trackChanges: false` — accept/reject existing tracked changes without auto-recording new ones.

---

## Event System

```typescript
import { onlyofficeEventbus, ONLYOFFICE_EVENT_KEYS } from "@/onlyoffice-comp";
```

| Key | Value | When | Payload |
|-----|-------|------|---------|
| `DOCUMENT_READY` | `documentReady` | Document loaded | `{ fileName, fileType, instanceId? }` |
| `SAVE_DOCUMENT` | `saveDocument` | Save/export with bin | `{ fileName, fileType, binData, instanceId, media? }` |
| `ONSAVE` | `onSave` | User save completed (lightweight) | `{ fileName, instanceId }` |
| `LOADING_CHANGE` | `loadingChange` | Language remount, etc. | `{ loading: boolean }` |

```typescript
onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, (data) => {
  upload(data.binData);
});

const data = await onlyofficeEventbus.waitFor(
  ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY,
  30_000,
);

onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
```

**Multi-instance**: filter with `data.instanceId === manager.getInstanceId()`.

**When to use which event**:

- Need binary data → `SAVE_DOCUMENT` or `export()` / `exportDocument()`
- UI “saved” toast only → `ONSAVE`
- User save shortcut / toolbar save → both `SAVE_DOCUMENT` and `ONSAVE` (browser download is intercepted)

---

## Export & Save

### Programmatic export

```typescript
const snapshot = await manager.exportDocument();
const { blob, fileName } = await manager.exportAsBlob();
await manager.downloadExport();
```

Manual x2t:

```typescript
import { convertBinToDocument, downloadBlob, getOfficeMimeType } from "@/onlyoffice-comp";

const snapshot = await editor.export();
const { data, fileName } = await convertBinToDocument(
  snapshot.binData,
  snapshot.fileName,
  FILE_TYPE.DOCX,
  snapshot.media,
);
downloadBlob(new Blob([data], { type: getOfficeMimeType(FILE_TYPE.DOCX) }), fileName);
```

- In read-only mode, `export()` returns cached snapshot without triggering iframe save.
- `convertBinToDocument` injects registered Word fonts for x2t export aliases.

### User save (inside editor)

When the user saves from the editor UI:

1. Mock server updates in-memory `Editor.bin` (**no** browser docx download)
2. Emits `SAVE_DOCUMENT` (with `binData`) and `ONSAVE`

```typescript
onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, async (data) => {
  await uploadToServer(data.binData, data.fileName);
});
```

---

## Comments & Revisions

Call via `EditorManager` or `manager.getEditor()`.

```typescript
const editor = manager.getEditor();

const id = editor.addComment("Please review this paragraph");
editor.goToComment(id, { showBalloon: true });
await editor.registerCommentCallbacks({ onAdd: (id, data) => {} });

editor.acceptAllRevisions();
editor.rejectRevision(revisionId);
```

Types `CommentData`, `CommentItem`, `RevisionItem`, etc. are exported from `@/onlyoffice-comp`.

---

## Word Custom Fonts

自定义字体由 **`AllFonts.js` 内联维护**（同步加载，无额外请求）：

- `window.__fonts_files` / `window.__fonts_infos`：catalog 二进制与族名映射
- `window.__custom_font_registry__`：id、displayName、aliases
- 末尾 `custom-fonts.js`：运行时注入二进制、L1b 别名、字体解析补丁

构建时编辑 `public/9.3.0/fonts/custom-fonts.json`（脚本源文件），运行：

```bash
node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs --id 1000 --verify
```

脚本会：写入线格式 → 更新 `__fonts_files` / `__fonts_infos`（仅 displayName 一行）→ 把 JSON 同步进 `__custom_font_registry__` → 追加最新 `custom-fonts.js`。

只改了 registry 或补丁逻辑时：

```bash
node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs --sync-bootstrap
```

---

## Constants & Store

```typescript
import {
  ONLYOFFICE_ID,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  ONLYOFFICE_LANG_KEY,
  FILE_TYPE,
  STATIC_RESOURCE,
  getDocumentType,
  getNewUrl,
} from "@/onlyoffice-comp";
```

| Name | Description |
|------|-------------|
| `STATIC_RESOURCE` | SDK / x2t paths; bump SDK via `NEXT_PUBLIC_APP_ROOT` |
| `FILE_TYPE` | `DOCX` / `XLSX` / `PPTX` for x2t export |
| `ONLYOFFICE_LANG_KEY` | `{ ZH: 'zh', EN: 'en' }` |

### Document store

```typescript
import {
  setDocmentObj,
  getDocmentObj,
  setNewDocument,
  setDocumentFile,
  setDocumentUrl,
} from "@/onlyoffice-comp";
```

### Language store

```typescript
import { getCurrentLang, setCurrentLang, getOnlyOfficeLang } from "@/onlyoffice-comp";
```

---

## Utilities

| Function | Description |
|----------|-------------|
| `initializeOnlyOffice()` | Load DocsAPI + preload iframe (singleton Promise) |
| `createEditorView(options)` | Create editor view |
| `convertBinToDocument(bin, fileName, fileType, media?)` | Editor.bin → Office file |
| `downloadBlob(blob, fileName)` | Trigger browser download |
| `getOfficeMimeType(fileType)` | Office MIME type |

---

## Multi-Instance

```tsx
<div className="onlyoffice-container relative" data-onlyoffice-container-id="editor-1">
  <div id="editor-1" className="absolute inset-0" />
</div>
```

```typescript
await createEditorView({
  containerId: "editor-1",
  isNew: true,
  fileName: "a.docx",
});
```

Use matching `data-onlyoffice-container-id` and `containerId` to avoid cross-instance routing.

---

## Supported Formats

| Type | Common extensions |
|------|-------------------|
| Word | docx, doc, odt, rtf, txt, … |
| Excel | xlsx, xls, ods, csv, … |
| PowerPoint | pptx, ppt, odp, … |
| Other | pdf (read-only), vsdx (Draw) |

x2t export targets: `FILE_TYPE.DOCX | XLSX | PPTX` only.

---

## Notes

1. **Language change** remounts the iframe (`lang` is in the init URL) and emits `LOADING_CHANGE`.
2. **Read-only toggle** uses `asc_setRestriction` without destroying the iframe.
3. **Fonts**: system fonts not registered in `AllFonts.js` may fall back to Arial after x2t round-trip.
4. Do not import from `internal/` — only use APIs exported from `index.ts`.
5. See `onlyoffice-comp/docs/` for topic-specific guides.
