# 02 - 核心 API

[← 快速开始](./01-快速开始.md) | [事件系统 →](./03-事件系统.md)

## `initializeOnlyOffice()`

初始化 OnlyOffice 编辑器环境，包括加载脚本、API 和 X2T 转换器。

```typescript
import { initializeOnlyOffice } from '@/onlyoffice-comp/lib/utils';

await initializeOnlyOffice();
```

**特点：**

- 使用单例模式，多次调用只会初始化一次
- 自动加载所有必需的资源
- 返回 Promise，支持异步等待

## `createEditorView(options)`

创建编辑器视图，支持新建或打开文档。支持单实例和多实例两种模式。

```typescript
import { createEditorView } from '@/onlyoffice-comp/lib/x2t';

await createEditorView({
  isNew: boolean;           // 是否新建文档
  fileName: string;         // 文件名（包含扩展名）
  file?: File;              // 文件对象（打开现有文档时必需）
  readOnly?: boolean;       // 是否只读模式，默认为 false
  lang?: string;            // 界面语言，默认为 'en'
  containerId?: string;     // 容器ID（多实例模式必需，单实例模式可选）
  editorManager?: EditorManager; // 编辑器管理器实例（可选）
});
```

**返回值：** `Promise<EditorManager>` - 返回编辑器管理器实例

**单实例模式：**

```typescript
await createEditorView({
  isNew: true,
  fileName: 'document.docx',
});
```

**多实例模式：**

```typescript
const manager = await createEditorView({
  isNew: true,
  fileName: 'document.docx',
  containerId: 'editor-1',
});
```

**支持的文件类型：**

- Word: `.docx`, `.doc`, `.odt`, `.rtf`, `.txt`
- Excel: `.xlsx`, `.xls`, `.ods`, `.csv`
- PowerPoint: `.pptx`, `.ppt`, `.odp`

## `editorManagerFactory` 和 `EditorManager`

编辑器管理器工厂和编辑器管理器，提供编辑器的操作和控制功能。

### 单实例模式（向后兼容）

```typescript
import { editorManagerFactory } from '@/onlyoffice-comp/lib/editor-manager';

const editorManager = editorManagerFactory.getDefault();

if (editorManager.exists()) {
  // 编辑器已创建
}

const binData = await editorManager.export();

await editorManager.setReadOnly(true);
await editorManager.setReadOnly(false);

const isReadOnly = editorManager.getReadOnly();

editorManager.destroy();
```

### 多实例模式

```typescript
import { editorManagerFactory } from '@/onlyoffice-comp/lib/editor-manager';

const manager1 = editorManagerFactory.create('editor-1');
const manager2 = editorManagerFactory.create('editor-2');

const manager = editorManagerFactory.get('editor-1');
const allManagers = editorManagerFactory.getAll();

editorManagerFactory.destroy('editor-1');
editorManagerFactory.destroyAll();
```

### `EditorManager` 实例方法

| 方法 | 说明 |
|------|------|
| `exists()` | 检查编辑器是否存在 |
| `export()` | 导出文档二进制数据 |
| `setReadOnly(readOnly)` | 切换只读/可编辑（会触发 `LOADING_CHANGE`） |
| `getReadOnly()` | 获取当前只读状态 |
| `getInstanceId()` | 获取实例唯一 ID |
| `getContainerId()` | 获取容器 ID |
| `getFileName()` | 获取当前文件名 |
| `getContainerParentSelector()` | 获取容器父元素选择器 |
| `getContainerStyle()` | 获取容器样式配置 |
| `updateMedia(key, url)` | 更新媒体文件映射 |
| `getMedia()` | 获取媒体文件映射 |
| `destroy()` | 销毁编辑器实例 |
| `subscribe({ type, fn })` | 订阅 Word SDK 回调，见 [07-批注修订与 Word API](./07-批注修订与-Word-API.md) |

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

在多实例模式下，`export()` 通过 `instanceId` 过滤 `SAVE_DOCUMENT` 事件，只接收当前实例的保存结果。只读模式下直接返回已缓存的 `binData`，不触发编辑器保存。

批注、修订相关方法见 [07-批注修订与 Word API](./07-批注修订与-Word-API.md)。

## `convertBinToDocument()`

将二进制数据转换为指定格式的文档。

```typescript
import { convertBinToDocument } from '@/onlyoffice-comp/lib/x2t';
import { FILE_TYPE } from '@/onlyoffice-comp/lib/const';

const result = await convertBinToDocument(
  binData.binData,
  binData.fileName,
  FILE_TYPE.DOCX,
);

// result: { fileName: string, data: Uint8Array }
```

**支持的文件类型：**

- `FILE_TYPE.DOCX` - Word 文档
- `FILE_TYPE.XLSX` - Excel 表格
- `FILE_TYPE.PPTX` - PowerPoint 演示文稿
