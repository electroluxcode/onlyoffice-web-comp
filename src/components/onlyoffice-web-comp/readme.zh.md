# OnlyOffice Web Comp

> 📖 [English](readme.md) | 中文

基于 OnlyOffice 静态 SDK 的 **Web 端文档编辑组件库**，支持 Word / Excel / PowerPoint 的在线编辑、只读预览、导出与 x2t 格式转换。**无需自建 Document Server**，只需托管 SDK 静态资源。

> 本文档为**入口页**。完整说明见 [`docs/`](./docs/00-概述.md)。

## 文档导航

| 编号 | 文档 | 说明 |
|------|------|------|
| 00 | [概述](./docs/00-概述.md) | 文档索引与阅读路径 |
| 01 | [快速开始](./docs/01-快速开始.md) | 初始化、容器挂载、创建编辑器 |
| 02 | [核心 API](./docs/02-核心API.md) | `OnlyOfficeManager`、`EditorManager`、多实例 |
| 03 | [事件系统](./docs/03-事件系统.md) | EventBus、事件类型与监听 |
| 04 | [完整示例](./docs/04-完整示例.md) | React 集成示例 |
| 05 | [API 参考](./docs/05-API参考.md) | 常量、类型定义 |
| 06 | [注意事项与支持格式](./docs/06-注意事项与支持格式.md) | 前置条件、文件格式、常见坑 |
| 07 | [批注修订与 Word API](./docs/07-批注修订与-Word-API.md) | 批注、修订、SDK 回调 |

**推荐阅读路径**：首次接入 → [01](./docs/01-快速开始.md) → [02](./docs/02-核心API.md) → [04](./docs/04-完整示例.md)

## 目录结构

```
onlyoffice-web-comp/
├── const/       常量、静态资源路径、文件类型
├── store/       文档 / 语言等跨页面状态
├── util/        SDK 初始化、x2t 转换、下载
├── core/        EditorManager、OnlyOfficeManager、EventBus
├── feature/     批注、修订
├── docs/        完整使用文档（从这里开始）
└── internal/    mock server / x2t worker（不对外导出）
```

## 最小示例

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

// 新建空白文档
const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
});

// 打开已有 File（先取文件，再挂载）
const file = await fetch("/test.xlsx").then((r) => r.blob())
  .then((blob) => new File([blob], "test.xlsx", { type: blob.type }));
await OnlyOfficeManager.createWithFile({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.XLSX,
  defaultFileName: "test.xlsx",
}, file);
```

更多用法（事件、导出、多实例、只读切换）见 [docs/02-核心API.md](./docs/02-核心API.md)。

## 本仓库中的演示

组件库旁的演示页面位于 `src/components/onlyoffice-web-demo/`，可在以下路由体验：

- `/docs/base` — Word 单实例
- `/excel/base` — Excel 单实例（默认加载 `public/test.xlsx`）
- `/ppt/base` — PPT 单实例
- `/multi/base` — 多实例并排
- `/multi/tabs` — 多实例 Tab 切换

## 相关链接

- [项目根 README（仓库总览）](../../../README.zh.md)
- [OnlyOffice 官方 API](https://api.onlyoffice.com/zh-CN/docs/docs-api/usage-api/config/document/)
