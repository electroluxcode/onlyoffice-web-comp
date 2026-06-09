/**
 * OnlyOffice Word 编辑器 iframe 运行时类型（对照 sdk-all-min.js 混淆名）。
 * 升级 SDK 时优先核对此文件。
 */


export type AscLogicDocument = {
  Ai?: () => void;
};


export type OnlyOfficeIframeWindow = typeof window & {
  Common?: {
    NotificationCenter?: {
      trigger: (
        event: string,
        disabled: boolean,
        options: Record<string, unknown>,
      ) => void;
    };
  };
  Asc?: {
  };
  __ONLYOFFICE_PROXIES_INSTALLED__?: boolean;
  __ONLYOFFICE_SAVE_BLOCKED__?: boolean;
};
