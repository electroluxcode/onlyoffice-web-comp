# 05 - API 参考

[← 完整示例](./04-完整示例.md) | [注意事项 →](./06-注意事项与支持格式.md)

## 常量

### `ONLYOFFICE_ID`

编辑器容器的 DOM ID，默认为 `'iframe-office-id'`。

### `ONLYOFFICE_CONTAINER_CONFIG`

| 字段 | 说明 |
|------|------|
| `PARENT_SELECTOR` | 父元素选择器 `.onlyoffice-container` |
| `PARENT_CLASS_NAME` | 父元素类名 `onlyoffice-container` |
| `STYLE` | 容器绝对定位样式 `{ position, inset }` |

### `ONLYOFFICE_EVENT_KEYS`

| 常量 | 值 | 说明 |
|------|-----|------|
| `SAVE_DOCUMENT` | `saveDocument` | 保存完成，含 `binData` |
| `DOCUMENT_READY` | `documentReady` | 文档就绪 |
| `LOADING_CHANGE` | `loadingChange` | Loading 状态 |
| `ONSAVE` | `onSave` | 保存流程结束（轻量） |

### `FILE_TYPE`

- `FILE_TYPE.DOCX` — Word
- `FILE_TYPE.XLSX` — Excel
- `FILE_TYPE.PPTX` — PowerPoint

### `ONLYOFFICE_LANG_KEY`

- `ONLYOFFICE_LANG_KEY.ZH` — `zh`（**默认语言**）
- `ONLYOFFICE_LANG_KEY.EN` — `en`

### `READONLY_SWITCH_MIN_DELAY_MS`

只读 ↔ 编辑切换时 loading 最短展示时长，值为 `200`（ms）。

### `STATIC_RESOURCE`

OnlyOffice SDK 与 x2t 静态资源路径总入口。

```typescript
import { STATIC_RESOURCE } from "@/components/onlyoffice-web-comp";

STATIC_RESOURCE.onlyoffice.root   // 默认 /packages/onlyoffice/9.3.0
STATIC_RESOURCE.onlyoffice.apiUrl   // api.js 绝对 URL
STATIC_RESOURCE.x2t.script          // x2t.js 路径
STATIC_RESOURCE.x2t.wasm            // x2t.wasm 路径
```

可通过环境变量 `NEXT_PUBLIC_APP_ROOT` 覆盖 SDK 根路径。

> `ONLYOFFICE_RESOURCE` / `X2T_RESOURCE` 已标记 `@deprecated`，请使用 `STATIC_RESOURCE`。

## 类型定义

### `OnlyOfficeManagerOptions`

```typescript
type OnlyOfficeManagerOptions = {
  containerId?: string;
  fileType: FileType;
  defaultFileName: string;
  readOnly?: boolean;
  lang?: OnlyOfficeLang;
};
```

### `OpenDocumentInput`

```typescript
type OpenDocumentInput = {
  fileName: string;
  file?: File;
  isNew?: boolean;
  readOnly?: boolean;
};
```

### `DocumentReadyData`

```typescript
type DocumentReadyData = {
  fileName: string;
  fileType: string;
  instanceId?: string;
};
```

### `SaveDocumentData`

```typescript
type SaveDocumentData = {
  fileName: string;
  fileType: string;
  binData: Uint8Array;
  instanceId: string;
  media?: Record<string, Uint8Array>;
};
```

### `OnSaveData`

```typescript
type OnSaveData = {
  fileName: string;
  instanceId: string;
};
```

### `LoadingChangeData`

```typescript
type LoadingChangeData = {
  loading: boolean;
};
```

### `AscWordApiMethod`

Word 编辑器 iframe 内 SDK 方法名联合类型，定义于 `type/word-api.ts`。用于 `EditorManager.subscribe({ type, fn })` 的 `type` 参数。

常用条目（节选）：

```typescript
// 批注
| 'asc_onAddComment'
| 'asc_onChangeCommentData'
| 'asc_onRemoveComment'
// 修订
| 'asc_onShowRevisionsChange'
// 文档状态
| 'asc_onDocumentModifiedChanged'
| 'asc_onSaveCallback'
```

完整列表见源码 `type/word-api.ts`。

## 导出清单

统一从包入口导入：

```typescript
import {
  // 门面
  OnlyOfficeManager,
  onlyOfficeManagerFactory,
  // 底层
  EditorManager,
  editorManager,
  editorManagerFactory,
  createEditorView,
  initializeOnlyOffice,
  convertBinToDocument,
  // 事件
  onlyofficeEventbus,
  ONLYOFFICE_EVENT_KEYS,
  // 常量
  FILE_TYPE,
  ONLYOFFICE_ID,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_LANG_KEY,
  STATIC_RESOURCE,
  // store
  setDocmentObj,
  getDocmentObj,
} from "@/components/onlyoffice-web-comp";
```

类型通过 `export type *` 从 `type/word-api.ts`、`type/sdk-internal.ts` 导出。
