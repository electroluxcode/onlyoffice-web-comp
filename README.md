# OnlyOffice Web Comp

> 📖 English | [中文](README.zh.md)

🌐 **Live Demo**: https://onlyoffice-web-comp.vercel.app/

A browser-based document solution built on the OnlyOffice static SDK. View, edit, and convert Word, Excel, and PowerPoint entirely on the client—**no Document Server required**.

This repository has two parts:

| Part | Path | Description |
|------|------|-------------|
| **Component library** | [`src/components/onlyoffice-web-comp/`](src/components/onlyoffice-web-comp/) | Reusable Web editor wrapper |
| **Demo app** | [`src/app/`](src/app/) + [`src/components/onlyoffice-web-demo/`](src/components/onlyoffice-web-demo/) | Next.js demos and integration reference |

## Core Advantages

- **Local processing**: Documents stay in the browser
- **Format support**: Word, Excel, PowerPoint
- **No backend**: Host static SDK assets only
- **Multi-instance**: Side-by-side editors and tabbed caching

## Quick Try (Demo App)

1. Visit the [live demo](https://onlyoffice-web-comp.vercel.app/) or run locally at `http://localhost:3001`
2. Open a route:

| Route | Description |
|-------|-------------|
| `/excel/base` | Excel single instance (loads `test.xlsx` by default) |
| `/docs/base` | Word single instance |
| `/ppt/base` | PowerPoint single instance |
| `/multi/base` | Multi-instance side-by-side |
| `/multi/tabs` | Multi-instance with tabs |

3. Upload a file → edit → export

### URL Parameters

| Param | Description | Values |
|-------|-------------|--------|
| `locale` | UI language | `en`, `zh` |

## Component Library Docs

**API and integration details are not duplicated here.** Use the component library documentation:

- **Entry**: [Component README (English)](src/components/onlyoffice-web-comp/readme.md)
- **Full docs**: [docs/00-概述.md](src/components/onlyoffice-web-comp/docs/00-概述.md)

| Doc | Topic |
|-----|-------|
| [01 Quick Start](src/components/onlyoffice-web-comp/docs/01-快速开始.md) | Init and container mount |
| [02 Core API](src/components/onlyoffice-web-comp/docs/02-核心API.md) | `OnlyOfficeManager`, multi-instance |
| [03 Events](src/components/onlyoffice-web-comp/docs/03-事件系统.md) | EventBus |
| [04 Examples](src/components/onlyoffice-web-comp/docs/04-完整示例.md) | React examples |
| [05 Reference](src/components/onlyoffice-web-comp/docs/05-API参考.md) | Constants and types |
| [06 Notes](src/components/onlyoffice-web-comp/docs/06-注意事项与支持格式.md) | Prerequisites and formats |
| [07 Comments](src/components/onlyoffice-web-comp/docs/07-批注修订与-Word-API.md) | Comments and revisions |

```typescript
import { OnlyOfficeManager, FILE_TYPE, ONLYOFFICE_ID } from "@/components/onlyoffice-web-comp";
```

## Local Development

```bash
git clone <repository-url>
cd onlyoffice-web-comp
pnpm install
pnpm dev
# http://localhost:3001
```

## Project Structure

```
onlyoffice-web-comp/
├── src/
│   ├── app/                          # Next.js routes (demo entry)
│   │   ├── excel/base/               # Excel demo
│   │   ├── docs/base/                # Word demo
│   │   ├── ppt/base/                 # PPT demo
│   │   └── multi/                    # Multi-instance demos
│   ├── components/
│   │   ├── onlyoffice-web-comp/      # Component library + docs/
│   │   ├── onlyoffice-web-demo/      # Demo page wrappers
│   │   └── studio-layout/            # Demo site layout
│   └── ...
├── public/                           # OnlyOffice SDK static assets
└── scripts/                          # Build tools (e.g. minify)
```

## Tech Stack

- **OnlyOffice SDK**: Core editing
- **x2t + WebAssembly**: Format conversion
- **Next.js 15 + React 19**: Demo application

## Deployment

```bash
pnpm install
pnpm build
```

Deploy to Vercel or any static host. Live demo: https://onlyoffice-web-comp.vercel.app/

## Fonts

Custom fonts are registered via **`__custom_font_registry__`**, with **`ttf-to-catalog-font.mjs`** producing OnlyOffice catalog wire-format files.

### 1. Convert TTF/OTF to catalog wire format

Script:

`public/packages/onlyoffice/9.3.0/fonts/ttf-to-catalog-font.mjs`

Place the source font next to the script (e.g. `1001.ttf`), or pass an explicit input path:

```bash
# Read 1001.ttf from the same directory → public/packages/onlyoffice/9.3.0/fonts/1001
node public/packages/onlyoffice/9.3.0/fonts/ttf-to-catalog-font.mjs --id 1001 --verify

# Explicit source file
node public/packages/onlyoffice/9.3.0/fonts/ttf-to-catalog-font.mjs ./MyFont.ttf --id 1001 --verify
```

Output is an extensionless catalog file: `public/packages/onlyoffice/9.3.0/fonts/{id}`.

A copy also lives at `src/components/onlyoffice-web-comp/scripts/fonts/ttf-to-catalog-font.mjs`.

### 2. Register aliases in `__custom_font_registry__`

Edit `public/packages/onlyoffice/9.3.0/sdkjs/common/AllFonts.js`:

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

- **Keys** (e.g. `"1001"`) must match the `--id` and the filename under `fonts/`
- **Values** are alias arrays covering every font name used in your documents


### 3. Built-in fonts (optional)

Built-in glyphs still use **numeric indices** in `__fonts_files`. To replace one, place a catalog wire file at `public/packages/onlyoffice/9.3.0/fonts/{index}` (no extension) using the index from that array.

Ensure all font files comply with applicable licenses.

## Related Resources

- [OnlyOffice API docs](https://api.onlyoffice.com/docs/docs-api/usage-api/config/document/)
- [OnlyOffice Web Apps](https://github.com/ONLYOFFICE/web-apps)
- [OnlyOffice SDK](https://github.com/ONLYOFFICE/sdkjs)
- [x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)

## Contributing

Issues and Pull Requests are welcome.

## License

See [LICENSE](LICENSE).
