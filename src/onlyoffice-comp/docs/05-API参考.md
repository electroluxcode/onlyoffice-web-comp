# 05 - API 参考

[← 完整示例](./04-完整示例.md) | [注意事项 →](./06-注意事项与支持格式.md)

## 常量

### `ONLYOFFICE_ID`

编辑器容器的 DOM ID，默认为 `'iframe-office-id'`。

### `ONLYOFFICE_CONTAINER_CONFIG`

容器挂载配置：

| 字段 | 说明 |
|------|------|
| `ID` | 同 `ONLYOFFICE_ID` |
| `PARENT_SELECTOR` | 父元素选择器 `.onlyoffice-container` |
| `PARENT_CLASS_NAME` | 父元素类名 `onlyoffice-container` |
| `STYLE` | 容器绝对定位样式 |

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

- `ONLYOFFICE_LANG_KEY.ZH` — `zh`
- `ONLYOFFICE_LANG_KEY.EN` — `en`

### `READONLY_TIMEOUT_CONFIG`

见 [03-事件系统](./03-事件系统.md#等待事件)。

### `ONLYOFFICE_RESOURCE`

OnlyOffice 7.0.7 静态资源路径（`api.js`、`x2t.js`、`xlsx` 等）。

## 类型定义

### `DocumentReadyData`

```typescript
type DocumentReadyData = {
  fileName: string;
  fileType: string;
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
