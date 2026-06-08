# OnlyOffice Web Comp

> 📖 [English](README.md) | 中文

🌐 **在线演示**: [https://onlyoffice-web-comp.vercel.app/](https://onlyoffice-web-comp.vercel.app/)

基于 OnlyOffice 静态 SDK 的浏览器端文档处理方案：在客户端完成 Word / Excel / PowerPoint 的查看、编辑与转换，**无需 Document Server**。

本仓库包含两部分：


| 部分       | 路径                                                                                                    | 说明              |
| -------- | ----------------------------------------------------------------------------------------------------- | --------------- |
| **组件库**  | `[src/components/onlyoffice-web-comp/](src/components/onlyoffice-web-comp/)`                          | 可复用的 Web 端编辑器封装 |
| **演示应用** | `[src/app/](src/app/)` + `[src/components/onlyoffice-web-demo/](src/components/onlyoffice-web-demo/)` | Next.js 演示与集成参考 |


## 核心优势

- **数据留在本地**：文档处理在浏览器内完成
- **格式兼容**：Word、Excel、PowerPoint 主流格式
- **零后端**：托管静态 SDK 即可使用
- **多实例**：支持多容器并行、Tab 切换缓存

## 快速体验（演示应用）

1. 访问 [在线编辑器](https://onlyoffice-web-comp.vercel.app/) 或本地 `http://localhost:3001`
2. 选择路由：


| 路由            | 说明                          |
| ------------- | --------------------------- |
| `/excel/base` | Excel 单实例（默认打开 `test.xlsx`） |
| `/docs/base`  | Word 单实例                    |
| `/ppt/base`   | PowerPoint 单实例              |
| `/multi/base` | 多实例并排演示                     |
| `/multi/tabs` | 多实例 Tab 演示                  |


1. 上传本地文件 → 编辑 → 导出

### URL 参数


| 参数       | 说明   | 可选值        |
| -------- | ---- | ---------- |
| `locale` | 界面语言 | `en`, `zh` |


## 组件库文档

**API 与接入说明不在本 README 重复**，请阅读组件库文档：

- **入口**：[组件库 README（中文）](src/components/onlyoffice-web-comp/readme.zh.md)
- **完整文档**：[docs/00-概述.md](src/components/onlyoffice-web-comp/docs/00-概述.md)


| 文档                                                                      | 内容                      |
| ----------------------------------------------------------------------- | ----------------------- |
| [01-快速开始](src/components/onlyoffice-web-comp/docs/01-快速开始.md)           | 初始化与容器挂载                |
| [02-核心API](src/components/onlyoffice-web-comp/docs/02-核心API.md)         | `OnlyOfficeManager`、多实例 |
| [03-事件系统](src/components/onlyoffice-web-comp/docs/03-事件系统.md)           | EventBus                |
| [04-完整示例](src/components/onlyoffice-web-comp/docs/04-完整示例.md)           | React 示例                |
| [05-API参考](src/components/onlyoffice-web-comp/docs/05-API参考.md)         | 常量与类型                   |
| [06-注意事项](src/components/onlyoffice-web-comp/docs/06-注意事项与支持格式.md)      | 前置条件与格式                 |
| [07-批注修订](src/components/onlyoffice-web-comp/docs/07-批注修订与-Word-API.md) | 批注、修订                   |


```typescript
import { OnlyOfficeManager, FILE_TYPE, ONLYOFFICE_ID } from "@/components/onlyoffice-web-comp";
```

## 本地开发

```bash
git clone <repository-url>
cd onlyoffice-web-comp
pnpm install
pnpm dev
# http://localhost:3001
```

## 项目结构

```
onlyoffice-web-comp/
├── src/
│   ├── app/                          # Next.js 路由（演示入口）
│   │   ├── excel/base/               # Excel 演示
│   │   ├── docs/base/                # Word 演示
│   │   ├── ppt/base/                 # PPT 演示
│   │   └── multi/                    # 多实例演示
│   ├── components/
│   │   ├── onlyoffice-web-comp/      # 组件库 + docs/
│   │   ├── onlyoffice-web-demo/      # 演示页面封装
│   │   └── studio-layout/            # 演示站布局
│   └── ...
├── public/                           # OnlyOffice SDK 静态资源
└── scripts/                          # 构建工具（如 minify）
```

## 技术栈

- **OnlyOffice SDK**：文档编辑核心
- **x2t + WebAssembly**：格式转换
- **Next.js 15 + React 19**：演示应用

## 部署

```bash
pnpm install
pnpm build
```

可部署至 Vercel 或任意静态托管。演示地址：[https://onlyoffice-web-comp.vercel.app/](https://onlyoffice-web-comp.vercel.app/)

## 字体配置

自定义字体通过 `**__custom_font_registry__**` 注册，配合 `**ttf-to-catalog-font.mjs**` 生成 OnlyOffice catalog 线格式。

### 1. TTF/OTF 转为 catalog 线格式

脚本位置：

`public/packages/onlyoffice/9.3.0/fonts/ttf-to-catalog-font.mjs`

将源字体放到脚本同目录（如 `1001.ttf`），或显式指定输入文件：

```bash
# 从同目录读取 1001.ttf，写入 public/packages/onlyoffice/9.3.0/fonts/1001
node public/packages/onlyoffice/9.3.0/fonts/ttf-to-catalog-font.mjs --id 1001 --verify

# 指定源文件
node public/packages/onlyoffice/9.3.0/fonts/ttf-to-catalog-font.mjs ./MyFont.ttf --id 1001 --verify
```

产物为 **无扩展名** 的 catalog 文件：`public/packages/onlyoffice/9.3.0/fonts/{id}`。

组件库内也有一份脚本副本：`src/components/onlyoffice-web-comp/scripts/fonts/ttf-to-catalog-font.mjs`。

### 2. 在 `__custom_font_registry__` 中注册别名

编辑 `public/packages/onlyoffice/9.3.0/sdkjs/common/AllFonts.js`，在 `window["__custom_font_registry__"]` 中声明 **字体文件 id → 文档内可能出现的名称列表**：

```javascript
window["__custom_font_registry__"] = {
  "1001": [
    "仿宋_GB2312",
    "Slidefu",
    "Slidefu Regular",
    "演示佛系体",
  ],
};
```

- **键**（如 `"1001"`）必须与上一步 `--id` 及 `fonts/` 下的文件名一致
- **值**为别名数组，需覆盖 Word / Excel / PPT 文档中实际使用的字体名

=

### 3. 内置字体（可选）

SDK 自带的内置字体仍通过 `AllFonts.js` 中 `__fonts_files` 的**数字索引**引用。若需替换某内置字形，将 catalog 线格式文件放到 `public/packages/onlyoffice/9.3.0/fonts/{索引号}`（无扩展名），索引号可在 `__fonts_files` 数组中查找。

请确保所用字体文件符合相关许可协议。

## 相关资源

- [OnlyOffice API 文档](https://api.onlyoffice.com/zh-CN/docs/docs-api/usage-api/config/document/)
- [OnlyOffice Web Apps](https://github.com/ONLYOFFICE/web-apps)
- [OnlyOffice SDK](https://github.com/ONLYOFFICE/sdkjs)
- [x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)

## 参与贡献

欢迎提交 Issue 和 Pull Request。

## 开源许可

详见 [LICENSE](LICENSE)。