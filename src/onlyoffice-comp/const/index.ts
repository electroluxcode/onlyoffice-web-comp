import { DocumentType } from "../internal/editor/types";

// ── 编辑器容器 / 事件 ──────────────────────────────────────────

export const ONLYOFFICE_ID = "iframe-office-id";

export const ONLYOFFICE_CONTAINER_CONFIG = {
  PARENT_SELECTOR: ".onlyoffice-container",
  PARENT_CLASS_NAME: "onlyoffice-container",
  STYLE: {
    position: "absolute",
    inset: 0,
  },
} as const;

export const ONLYOFFICE_EVENT_KEYS = {
  SAVE_DOCUMENT: "saveDocument",
  DOCUMENT_READY: "documentReady",
  LOADING_CHANGE: "loadingChange",
  ONSAVE: "onSave",
} as const;

export type OnlyOfficeEventKey =
  (typeof ONLYOFFICE_EVENT_KEYS)[keyof typeof ONLYOFFICE_EVENT_KEYS];

export const ONLYOFFICE_LANG_KEY = {
  ZH: "zh",
  EN: "en",
} as const;

/** 只读 ↔ 编辑切换时，loading 最少展示时长（ms） */
export const READONLY_SWITCH_MIN_DELAY_MS = 200;

/** Asc.c_oAscRestrictionType（sdk-all-min.js: k.Mf / k.Hca） */
export const ASC_RESTRICTION_NONE = 0;
export const ASC_RESTRICTION_VIEW = 128;

/** OnlyOffice 编辑器左上角 logo（jsDelivr 固定版本，避免依赖站点本地资源） */
export const OFFICE_EDITOR_LOGO = {
  /** 浅色主题：Office 品牌色图标 */
  image:
    "https://cdn.jsdelivr.net/npm/simple-icons@9.21.0/icons/microsoftoffice.svg",
  /** 深色主题：同图标（品牌色在深色背景上同样清晰） */
  imageDark:
    "https://cdn.jsdelivr.net/npm/simple-icons@9.21.0/icons/microsoftoffice.svg",
} as const;

// ── 静态资源（SDK / x2t）────────────────────────────────────────

type StaticResource = {
  /** 版本目录：升级资源时只改这里 */
  version: {
    onlyofficeSdk: string;
    x2t: string;
  };
  onlyoffice: {
    root: string;
    apiJs: string;
    preloadHtml: string;
    apiUrl: string;
    preloadUrl: string;
    fontUrl: (fontId: string) => string;
  };
  x2t: {
    root: string;
    script: string;
    wasm: string;
  };
};

function createStaticResource(): StaticResource {
  const onlyofficeSdkRoot =
    process.env.NEXT_PUBLIC_APP_ROOT || "/packages/onlyoffice/9.3.0";
  /** x2t 与 SDK 同版本目录；磁盘上为 Brotli 预压缩，由 x2t.worker 内 fetch-brotli 自动解压 */
  const x2tRoot = `${onlyofficeSdkRoot}/x2t-1`;
  const apiJs = "/web-apps/apps/api/documents/api.js";
  const preloadHtml = "/web-apps/apps/api/documents/preload.html";

  return {
    version: {
      onlyofficeSdk: onlyofficeSdkRoot,
      x2t: x2tRoot,
    },
    onlyoffice: {
      root: onlyofficeSdkRoot,
      apiJs,
      preloadHtml,
      apiUrl: onlyofficeSdkRoot + apiJs,
      preloadUrl: onlyofficeSdkRoot + preloadHtml,
      fontUrl: (fontId: string) => `${onlyofficeSdkRoot}/fonts/${fontId}`,
    },
    x2t: {
      root: x2tRoot,
      script: `${x2tRoot}/x2t.js`,
      wasm: `${x2tRoot}/x2t.wasm`,
    },
  };
}

/**
 * 静态资源路径总入口。
 * 升级 OnlyOffice SDK → 改 `version.onlyofficeSdk`（或 env `NEXT_PUBLIC_APP_ROOT`）
 * 升级 x2t WASM     → 改 `version.x2t`（Brotli 解压见 internal/editor/fetch-brotli.ts）
 */
export const STATIC_RESOURCE = createStaticResource();

/** @deprecated 使用 STATIC_RESOURCE.onlyoffice */
export const ONLYOFFICE_RESOURCE = {
  APP_ROOT: STATIC_RESOURCE.onlyoffice.root,
  API_JS: STATIC_RESOURCE.onlyoffice.apiJs,
  PRELOAD_HTML: STATIC_RESOURCE.onlyoffice.preloadHtml,
  API_URL: STATIC_RESOURCE.onlyoffice.apiUrl,
  PRELOAD_URL: STATIC_RESOURCE.onlyoffice.preloadUrl,
  fontUrl: STATIC_RESOURCE.onlyoffice.fontUrl,
} as const;

/** @deprecated 使用 STATIC_RESOURCE.x2t */
export const X2T_RESOURCE = {
  ROOT: STATIC_RESOURCE.x2t.root,
  SCRIPT: STATIC_RESOURCE.x2t.script,
  WASM: STATIC_RESOURCE.x2t.wasm,
} as const;

/** 站点相对路径 → 绝对 URL（Worker 内 origin 用 self.location.origin） */
export function resolveSiteUrl(origin: string, path: string): string {
  if (/^https?:\/\//i.test(path)) {
    return path;
  }
  return `${origin}${path.startsWith("/") ? path : `/${path}`}`;
}

export function getX2tBaseUrl(origin: string): string {
  return resolveSiteUrl(origin, `${STATIC_RESOURCE.x2t.root}/`);
}

// ── 文档类型 ────────────────────────────────────────────────────

/** x2t / 编辑器入参用的三类主格式（大写） */
export const FILE_TYPE = {
  DOCX: "DOCX",
  XLSX: "XLSX",
  PPTX: "PPTX",
} as const;

export type FileType = (typeof FILE_TYPE)[keyof typeof FILE_TYPE];

/** 与 SDK AppType 数值一致，仅内部映射用 */
const AppType = {
  word: 1,
  slide: 3,
  cell: 2,
  draw: 4,
  pdf: 5,
} as const;

const docTypeMap: Record<string, (typeof AppType)[keyof typeof AppType]> = {
  docx: AppType.word,
  doc: AppType.word,
  odt: AppType.word,
  rtf: AppType.word,
  txt: AppType.word,
  html: AppType.word,
  mht: AppType.word,
  epub: AppType.word,
  fb2: AppType.word,
  mobi: AppType.word,
  docm: AppType.word,
  dotx: AppType.word,
  dotm: AppType.word,
  oform: AppType.word,
  docxf: AppType.word,
  pptx: AppType.slide,
  ppt: AppType.slide,
  odp: AppType.slide,
  ppsx: AppType.slide,
  pptm: AppType.slide,
  ppsm: AppType.slide,
  potx: AppType.slide,
  potm: AppType.slide,
  otp: AppType.slide,
  odg: AppType.slide,
  xlsx: AppType.cell,
  xls: AppType.cell,
  ods: AppType.cell,
  csv: AppType.cell,
  xlsm: AppType.cell,
  xltx: AppType.cell,
  xltm: AppType.cell,
  xlsb: AppType.cell,
  ots: AppType.cell,
  vsdx: AppType.draw,
  vssx: AppType.draw,
  vstx: AppType.draw,
  vsdm: AppType.draw,
  vssm: AppType.draw,
  vstm: AppType.draw,
  pdf: AppType.pdf,
};

const appTypeName: Record<number, DocumentType> = {
  [AppType.word]: DocumentType.Word,
  [AppType.slide]: DocumentType.Slide,
  [AppType.cell]: DocumentType.Cell,
  [AppType.draw]: DocumentType.Draw,
  [AppType.pdf]: DocumentType.Pdf,
};

export function getDocumentType(ext: string) {
  const code = docTypeMap[ext.toLowerCase()];
  if (code === undefined) {
    return DocumentType.Word;
  }
  return appTypeName[code] ?? DocumentType.Word;
}

/** 新建文档页 URL */
export function getNewUrl(type: string) {
  return `/editor?new=${type}`;
}
