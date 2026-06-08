# 02 - 核心 API

[← 快速开始](./01-快速开始.md) | [事件系统 →](./03-事件系统.md)

## `OnlyOfficeManager`（推荐）

面向业务页面的门面，收敛初始化、开文档、导出、只读、语言等调用。

### `OnlyOfficeManager.create(options)`

加载 DocsAPI 并以 `defaultFileName` 新建空白文档。

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,   // 可选，默认 ONLYOFFICE_ID
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
  readOnly: false,
  lang: "zh",                   // 可选，默认 zh
});
```

### `OnlyOfficeManager.createWithFile(options, file)`

先拿到 `File`，再一次性挂载编辑器（不会先打开空白文档）。

```typescript
const manager = await OnlyOfficeManager.createWithFile(
  {
    containerId: ONLYOFFICE_ID,
    fileType: FILE_TYPE.XLSX,
    defaultFileName: "test.xlsx",
  },
  file,
);
```

### 实例方法

| 方法 | 说明 |
|------|------|
| `openDocument(input)` | 打开/切换文档（上传、新建、重开） |
| `openFile(file, readOnly?)` | 打开本地文件 |
| `openNew(fileName, readOnly?)` | 新建文档 |
| `isReady()` | 是否已打开文档 |
| `getReadOnly()` / `setReadOnly()` / `toggleReadOnly()` | 只读切换（同步，底层 `asc_setRestriction`） |
| `getLanguage()` / `setLanguage()` / `toggleLanguage()` | 语言切换 |
| `exportDocument()` | 导出 bin 数据 |
| `exportAsBlob()` | 导出为 Blob |
| `downloadExport()` | 导出并触发浏览器下载 |
| `onLoadingChange(handler)` | 监听 loading，返回取消函数 |
| `getEditor()` | 获取底层 `EditorManager` |
| `destroy()` | 销毁实例 |

### `OpenDocumentInput`

```typescript
type OpenDocumentInput = {
  fileName: string;
  file?: File;
  isNew?: boolean;
  readOnly?: boolean;
};
```

## `onlyOfficeManagerFactory`

多容器场景按 `containerId` 缓存 `OnlyOfficeManager` 门面。

```typescript
import {
  onlyOfficeManagerFactory,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

const manager = await onlyOfficeManagerFactory.open(
  {
    containerId: "editor-1",
    fileType: FILE_TYPE.DOCX,
    defaultFileName: "New_Document.docx",
    readOnly: false,
  },
  {
    fileName: "New_Document.docx",
    isNew: true,
  },
);

onlyOfficeManagerFactory.get("editor-1");
onlyOfficeManagerFactory.destroy("editor-1");
onlyOfficeManagerFactory.destroyAll();
```

## 底层 API

### `initializeOnlyOffice()`

```typescript
import { initializeOnlyOffice } from "@/components/onlyoffice-web-comp";

await initializeOnlyOffice();
```

- 单例模式，多次调用只初始化一次
- `OnlyOfficeManager.create` / `EditorManager.create` 内部会自动调用
- 仅在手动 `fromEditor` 绑定等高级场景需要显式调用

### `createEditorView(options)`

```typescript
import { createEditorView } from "@/components/onlyoffice-web-comp";

await createEditorView({
  isNew: boolean;
  fileName: string;
  file?: File;
  url?: string;
  loader?: (url: string) => Promise<ArrayBuffer>;
  fileType?: string;
  readOnly?: boolean;
  lang?: string;              // 默认跟随 store，初始为 zh
  containerId?: string;
  editorManager?: EditorManager;
  theme?: OfficeTheme;
  plugins?: PluginMode;
});
```

**返回值：** `Promise<EditorManager>`

**支持的文件类型：**

- Word: `.docx`, `.doc`, `.odt`, `.rtf`, `.txt`
- Excel: `.xlsx`, `.xls`, `.ods`, `.csv`
- PowerPoint: `.pptx`, `.ppt`, `.odp`

### `editorManagerFactory` 和 `EditorManager`

#### 单实例

```typescript
import { editorManagerFactory } from "@/components/onlyoffice-web-comp";

const editorManager = editorManagerFactory.getDefault();

if (editorManager.exists()) {
  // 编辑器已创建
}

const binData = await editorManager.export();

editorManager.setReadOnly(true);   // 同步方法
editorManager.setReadOnly(false);

const isReadOnly = editorManager.getReadOnly();

editorManager.destroy();
```

#### 多实例

```typescript
const manager1 = editorManagerFactory.create("editor-1");
const manager2 = editorManagerFactory.get("editor-2"); // 不存在时自动 create

const allManagers = editorManagerFactory.getAll();

editorManagerFactory.destroy("editor-1");
editorManagerFactory.destroyAll();
```

#### `EditorManager` 实例方法

| 方法 | 说明 |
|------|------|
| `exists()` | 检查编辑器是否存在 |
| `export()` | 导出文档二进制数据 |
| `setReadOnly(readOnly)` | 切换只读/可编辑（同步） |
| `getReadOnly()` | 获取当前只读状态 |
| `getInstanceId()` | 获取实例唯一 ID |
| `getContainerId()` | 获取容器 ID |
| `getFileName()` | 获取当前文件名 |
| `updateMedia(key, url)` | 更新媒体文件映射 |
| `getMedia()` | 获取媒体文件映射 |
| `destroy()` | 销毁编辑器实例 |
| `subscribe({ type, fn })` | 订阅 Word SDK 回调，见 [07](./07-批注修订与-Word-API.md) |

**`export()` 返回值：**

```typescript
{
  fileName: string;
  fileType: string;
  binData: Uint8Array;
  instanceId?: string;
  media?: Record<string, Uint8Array>;
}
```

多实例下 `export()` 通过 `instanceId` 过滤 `SAVE_DOCUMENT` 事件。只读模式下直接返回已缓存的 `binData`。

### `convertBinToDocument()`

```typescript
import { convertBinToDocument, FILE_TYPE } from "@/components/onlyoffice-web-comp";

const result = await convertBinToDocument(
  binData.binData,
  binData.fileName,
  FILE_TYPE.DOCX,
  binData.media,
);

// result: { fileName: string, data: ArrayBuffer }
```

业务页面导出优先使用 `OnlyOfficeManager.downloadExport()`，无需手动转换。
