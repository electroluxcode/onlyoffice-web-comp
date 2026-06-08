# OnlyOffice Comp 使用文档

> 📖 [English](readme.md) | 中文

OnlyOffice Comp 是基于 OnlyOffice Document Server 静态 SDK 的浏览器端文档编辑器封装，支持 Word / Excel / PowerPoint 的在线编辑、只读预览、导出与 x2t 格式转换。**无需自建 Document Server**，只需在站点托管 SDK 静态资源。

## 目录

- [架构与入口](#架构与入口)
- [前置条件](#前置条件)
- [快速开始](#快速开始)
- [OnlyOfficeManager（推荐）](#onlyofficemanager推荐)
- [EditorManager（底层）](#editormanager底层)
- [事件系统](#事件系统)
- [导出与保存](#导出与保存)
- [批注与修订](#批注与修订)
- [Word 自定义字体](#word-自定义字体)
- [常量与 Store](#常量与-store)
- [工具函数](#工具函数)
- [多实例](#多实例)
- [支持格式](#支持格式)
- [注意事项](#注意事项)

---

## 架构与入口

```
onlyoffice-comp/
├── const/          常量、静态资源路径、文件类型
├── store/          文档 / 语言等跨页面状态
├── util/           SDK 初始化、x2t 转换、下载
├── core/           EditorManager、OnlyOfficeManager、EventBus
├── feature/        批注、修订、Word 字体
└── internal/       mock server / x2t worker（不对外导出）
```

**统一从包入口导入：**

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

`index.ts` 会 re-export 上述模块；类型见 `type/word-api.ts`、`type/sdk-internal.ts`。

---

## 前置条件

1. **静态资源**：将 OnlyOffice SDK（含 `web-apps/`、`sdkjs/`、`fonts/`、`x2t/`）放到站点可访问目录，默认 `public/9.3.0/`。
2. **环境变量**（可选）：`NEXT_PUBLIC_APP_ROOT=/9.3.0`，与 `STATIC_RESOURCE.onlyoffice.root` 一致。
3. **x2t Brotli**：`x2t/` 下 `x2t.js`、`x2t.wasm` 为 Brotli 预压缩文件；**无需**配置 `Content-Encoding: br`，Worker 内 `fetch-brotli` + 内置解码器会自动解压。
4. **DOM 容器**：页面需预留编辑器挂载点（见下方示例）。

---

## 快速开始

### 1. 页面容器

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

### 2. 创建编辑器（推荐：`OnlyOfficeManager`）

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

// 打开本地文件
await manager.openFile(file);

// 导出为 docx 并触发浏览器下载
await manager.downloadExport();
```

### 3. 创建编辑器（底层：`createEditorView`）

适合需要完全控制 `EditorManager`、不经过 document store 的场景：

```typescript
import { createEditorView, initializeOnlyOffice } from "@/onlyoffice-comp";

await initializeOnlyOffice(); // createEditorView 内部也会调用，可省略

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

## OnlyOfficeManager（推荐）

面向业务页面的门面，封装 initialize、开文档、导出、只读、语言切换。

### 构造与工厂

| API | 说明 |
|-----|------|
| `OnlyOfficeManager.create(options)` | 初始化 DocsAPI 并以 `defaultFileName` 新建文档 |
| `OnlyOfficeManager.fromEditor(editor, options)` | 绑定已有 `EditorManager`，不自动 open |
| `onlyOfficeManagerFactory.open(options, document)` | 多容器场景：按 `containerId` 缓存门面 |
| `getDefaultOnlyOfficeManager()` / `setDefaultOnlyOfficeManager()` | 单页默认门面（可选） |

**`OnlyOfficeManagerOptions`**

| 字段 | 类型 | 说明 |
|------|------|------|
| `containerId` | `string?` | DOM 容器 id，默认 `ONLYOFFICE_ID` |
| `fileType` | `FileType` | 导出用主格式：`FILE_TYPE.DOCX` / `XLSX` / `PPTX` |
| `defaultFileName` | `string` | 首次 bootstrap 的文件名 |
| `readOnly` | `boolean?` | 初始只读 |
| `lang` | `'zh' \| 'en'?` | 初始界面语言 |

**`OpenDocumentInput`**

| 字段 | 说明 |
|------|------|
| `fileName` | 文件名（含扩展名） |
| `file` | 本地 `File` |
| `isNew` | 是否新建空白文档 |
| `readOnly` | 覆盖门面只读状态 |

### 实例方法

| 方法 | 说明 |
|------|------|
| `openDocument(input)` | 打开/切换文档（同步更新 document store） |
| `openNew(fileName, readOnly?)` | 新建 |
| `openFile(file, readOnly?)` | 打开本地文件 |
| `isReady()` | 是否已成功 open |
| `getReadOnly()` / `setReadOnly()` / `toggleReadOnly()` | 只读切换（底层 `asc_setRestriction`，不整页刷新） |
| `getLanguage()` / `setLanguage()` / `toggleLanguage()` | 中/英切换（会 remount iframe） |
| `getEditor()` | 获取底层 `EditorManager` |
| `exportDocument()` | 导出 `Editor.bin` 快照（见 [导出与保存](#导出与保存)） |
| `exportAsBlob()` | x2t 转为 Office 文件并返回 `Blob` |
| `downloadExport()` | `exportAsBlob` + 浏览器下载 |
| `onLoadingChange(handler)` | 订阅 `LOADING_CHANGE`，返回取消函数 |
| `destroy()` | 销毁编辑器 |

---

## EditorManager（底层）

每个 `containerId` 对应一个 `EditorManager`，管理 mock server、iframe 代理、DocsAPI 生命周期。

### 工厂

```typescript
import { editorManager, editorManagerFactory } from "@/onlyoffice-comp";

editorManagerFactory.getDefault();           // 单实例默认
editorManagerFactory.get("editor-1");        // 按容器获取/创建
editorManagerFactory.create("editor-1");
editorManagerFactory.getAll();
editorManagerFactory.destroy("editor-1");
editorManagerFactory.destroyAll();
```

### `createEditorView(options)`

等价于 `manager.create(options)`，返回 `Promise<EditorManager>`。

**`CreateEditorViewOptions`**

| 字段 | 说明 |
|------|------|
| `isNew` | 是否新建 |
| `fileName` | 文件名 |
| `file` | 本地文件 |
| `url` | 远程 URL（需配合 `loader` 或浏览器 fetch） |
| `loader` | 自定义 `(url) => ArrayBuffer` |
| `fileType` | 扩展名，默认从 `fileName` 解析 |
| `readOnly` | 只读 |
| `lang` | `'zh' \| 'en'` |
| `containerId` | 容器 id |
| `editorManager` | 指定已有实例 |
| `theme` | `OfficeTheme`，如 `'theme-white'` |
| `plugins` | `'featured' \| 'all' \| 'none'` |

### 实例方法一览

| 分类 | 方法 |
|------|------|
| 生命周期 | `create()`、`exists()`、`destroy()` |
| 导出 | `export()` |
| 只读/语言 | `setReadOnly()`、`getReadOnly()`、`setLanguage()` |
| 状态 | `getInstanceId()`、`getContainerId()`、`getFileName()`、`isDirty()` |
| 容器样式 | `getContainerParentSelector()`、`getContainerStyle()` |
| 媒体 | `updateMedia(key, data)`、`getMedia()` |
| SDK 订阅 | `subscribe({ type, fn })` |
| 批注 | `getAllComments()`、`addComment()`、`updateComment()`、`removeComment()`、`goToComment()`、`registerCommentCallbacks()` |
| 修订 | `setTrackRevisions()`、`isTrackRevisions()`、`getAllRevisions()`、`acceptRevision()`、`rejectRevision()`、`acceptAllRevisions()`、`rejectAllRevisions()`、`registerRevisionCallbacks()` |

**默认修订策略**：打开带修订的文档时 `review: true`、`trackChanges: false`——可接受/拒绝已有修订，但不自动进入「录制修订」模式。

---

## 事件系统

```typescript
import { onlyofficeEventbus, ONLYOFFICE_EVENT_KEYS } from "@/onlyoffice-comp";
```

| 事件键 | 常量值 | 触发时机 | 载荷 |
|--------|--------|----------|------|
| `DOCUMENT_READY` | `documentReady` | 文档加载完成 | `{ fileName, fileType, instanceId? }` |
| `SAVE_DOCUMENT` | `saveDocument` | 保存/导出完成，含 bin | `{ fileName, fileType, binData, instanceId, media? }` |
| `ONSAVE` | `onSave` | 用户保存流程结束（轻量） | `{ fileName, instanceId }` |
| `LOADING_CHANGE` | `loadingChange` | 语言切换等 remount | `{ loading: boolean }` |

```typescript
// 监听
onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, (data) => {
  // 上传 data.binData 到服务端
});

// 等待（带超时）
const data = await onlyofficeEventbus.waitFor(
  ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY,
  30_000,
);

// 取消
onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
```

**多实例**：用 `data.instanceId === manager.getInstanceId()` 过滤。

**选择建议**：

- 需要二进制 → `SAVE_DOCUMENT` 或 `editor.export()` / `manager.exportDocument()`
- 仅需「已保存」提示 → `ONSAVE`
- 用户 Ctrl+Shift+S / 工具栏保存 → 触发 `SAVE_DOCUMENT` + `ONSAVE`（已拦截浏览器直接下载）

---

## 导出与保存

### 主动导出

```typescript
const snapshot = await manager.exportDocument();
// snapshot: { fileName, fileType, binData, instanceId, media? }

const { blob, fileName } = await manager.exportAsBlob();
await manager.downloadExport();
```

或手动 x2t：

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

- 只读模式下 `export()` 直接返回已缓存快照，不触发 iframe 保存。
- `convertBinToDocument` 会注入 Word 注册字体包（x2t 导出别名）。

### 用户保存（编辑器内）

用户按保存快捷键或点工具栏保存时：

1. mock server 更新 `Editor.bin`（**不会**触发浏览器下载 docx）
2. 广播 `SAVE_DOCUMENT`（含 `binData`）与 `ONSAVE`

```typescript
onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, async (data) => {
  await uploadToServer(data.binData, data.fileName);
});
```

---

## 批注与修订

通过 `EditorManager` 或 `manager.getEditor()` 调用。

```typescript
const editor = manager.getEditor();

// 批注
const id = editor.addComment("请确认此处表述");
editor.goToComment(id, { showBalloon: true });
await editor.registerCommentCallbacks({
  onAdd: (id, data) => console.log("added", id, data),
});

// 修订（文档内已有修订时可接受/拒绝）
editor.acceptAllRevisions();
editor.rejectRevision(revisionId);
await editor.registerRevisionCallbacks({
  onShowChanges: (items) => console.log(items),
});
```

类型：`CommentData`、`CommentItem`、`RevisionItem` 等从 `@/onlyoffice-comp` 导出。

---

## Word 自定义字体

自定义字体直接在 OnlyOffice SDK 的 `AllFonts.js` 注册，不再经过 `onlyoffice-comp` 运行时补丁。

### 步骤

1. 将 TTF 放到 `onlyoffice-comp/scripts/fonts/{id}.ttf`
2. 生成 catalog 线格式并 patch `AllFonts.js`：

```bash
node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs --id 1000 --verify
```

3. 在 `public/v9.x/sdkjs/common/AllFonts.js` 的 `__fonts_infos` 中为同一 `fileIndex` 追加文档里可能出现的族名别名，例如：

```javascript
["STHupo", 218, 0, -1, -1, -1, -1, -1, -1],
["方正小标宋简体", 218, 1, -1, -1, -1, -1, -1, -1],
["FZXiaoBiaoSong-Z05S", 218, 2, -1, -1, -1, -1, -1, -1],
```

`218` 是 `"1000"` 在 `__fonts_files` 中的下标；第三个数字为同文件内的 face 序号（0、1、2…）。

字体二进制位于 `public/{version}/fonts/{id}`（catalog 线格式，无扩展名）。编辑器 iframe 会通过 AllFonts catalog 自动拉取。

---

## 常量与 Store

### 常用常量

```typescript
import {
  ONLYOFFICE_ID,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  ONLYOFFICE_LANG_KEY,
  FILE_TYPE,
  STATIC_RESOURCE,
  ASC_RESTRICTION_VIEW,
  ASC_RESTRICTION_NONE,
  getDocumentType,
  getNewUrl,
  resolveSiteUrl,
} from "@/onlyoffice-comp";
```

| 名称 | 说明 |
|------|------|
| `STATIC_RESOURCE` | SDK / x2t 路径总入口；升级 SDK 改 `NEXT_PUBLIC_APP_ROOT` 或 `version.onlyofficeSdk` |
| `FILE_TYPE` | `DOCX` / `XLSX` / `PPTX`（大写，用于 x2t 导出） |
| `ONLYOFFICE_LANG_KEY` | `{ ZH: 'zh', EN: 'en' }` |

### Document Store

```typescript
import {
  setDocmentObj,
  getDocmentObj,
  clearDocmentObj,
  setNewDocument,
  setDocumentFile,
  setDocumentUrl,
} from "@/onlyoffice-comp";
```

`OnlyOfficeManager.openDocument` 内部会同步 store。

### Language Store

```typescript
import {
  getCurrentLang,
  setCurrentLang,
  getOnlyOfficeLang,
} from "@/onlyoffice-comp";
```

---

## 工具函数

| 函数 | 说明 |
|------|------|
| `initializeOnlyOffice()` | 加载 DocsAPI 脚本 + preload iframe（单例 Promise） |
| `createEditorView(options)` | 创建编辑器视图 |
| `convertBinToDocument(bin, fileName, fileType, media?)` | Editor.bin → docx/xlsx/pptx |
| `downloadBlob(blob, fileName)` | 触发浏览器下载 |
| `getOfficeMimeType(fileType)` | Office MIME 类型 |

---

## 多实例

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

// 或使用 factory
await onlyOfficeManagerFactory.open(
  { containerId: "editor-1", fileType: FILE_TYPE.DOCX, defaultFileName: "a.docx" },
  { fileName: "a.docx", isNew: true },
);
```

**必须**使用 `data-onlyoffice-container-id` 与 `containerId` 一致，避免 iframe / 上传路由串实例。

---

## 支持格式

| 类型 | 扩展名（常见） |
|------|----------------|
| Word | docx, doc, odt, rtf, txt, … |
| Excel | xlsx, xls, ods, csv, … |
| PowerPoint | pptx, ppt, odp, … |
| 其他 | pdf（只读）, vsdx（Draw） |

x2t 导出主格式仅 `FILE_TYPE.DOCX | XLSX | PPTX`。

---

## 注意事项

1. **语言切换**需 remount iframe（`lang` 在初始化 URL 中），会触发 `LOADING_CHANGE`。
2. **只读切换**优先走 `asc_setRestriction`，不 destroy iframe；与语言切换不同。
3. **字体**：docx/xlsx 内系统字体（如宋体、等线）若未在 `AllFonts.js` 注册，x2t 可能回退为 Arial；见上文「Word 自定义字体」。
4. **internal/** 模块（`EditorServer`、`x2t.worker` 等）不保证对外稳定，请只使用 `index.ts` 导出 API。
5. 更细的专题文档见 `onlyoffice-comp/docs/` 目录。
