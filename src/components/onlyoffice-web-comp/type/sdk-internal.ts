/**
 * OnlyOffice 编辑器 iframe / Reporter popup 的运行时扩展类型。
 *
 * 对照 SDK 源码：public/packages/onlyoffice/9.3.0/sdkjs/.../sdk-all-min.js
 * 接入层在 iframe 内挂载的 __ONLYOFFICE_* 标记也在此集中声明，升级 SDK 时优先核对此文件。
 *
 * 关键静态路径（默认根目录 /packages/onlyoffice/9.3.0）：
 * - 主编辑器 iframe：web-apps/apps/{documenteditor|spreadsheeteditor|presentationeditor}/main/index.html
 * - PPT 演示者视图 popup：web-apps/apps/presentationeditor/main/index.reporter.html
 * - 协作 Mock 注入：core/editor-manager.ts → installIframeProxies() /
 *   internal/editor/install-proxies.ts
 */

/** Word 逻辑文档运行时（SDK 混淆名，如 ra.Ea）；仅声明接入层用到的字段。 */
export type AscLogicDocument = {
  /** 触发内容同步 / 重算（混淆名因 SDK 版本而异，升级时需对照 sdk-all-min.js）。 */
  Ai?: () => void;
};

/**
 * DocsAPI iframe 的 contentWindow 扩展形状。
 *
 * 生命周期：EditorManager.mountDocEditor() → onAppReady → installIframeProxies() 写入下列标记。
 */
export type OnlyOfficeIframeWindow = typeof window & {
  /** OnlyOffice UI 事件总线；只读切换等场景会调用 Common.NotificationCenter.trigger。 */
  Common?: {
    NotificationCenter?: {
      trigger: (
        event: string,
        disabled: boolean,
        options: Record<string, unknown>,
      ) => void;
    };
  };
  /** SDK 全局命名空间；Asc.editor 为当前实例 API（见 EditorManager.getSdkApi()）。 */
  Asc?: {
  };

  /**
   * XHR / fetch / socket.io 已代理到内存 EditorServer。
   * 写入：install-proxies.ts → installOnlyOfficeProxies()
   * 读取：EditorManager.installIframeProxies() 防重复安装。
   */
  __ONLYOFFICE_PROXIES_INSTALLED__?: boolean;

  /**
   * iframe 内 Ctrl/Cmd+S 保存快捷键已禁用。
   * 写入：EditorManager.installSaveShortcutBlocker()
   * 原因：保存与 export() 共用 /downloadas/ 管道，需避免用户触发原生保存。
   */
  __ONLYOFFICE_SAVE_BLOCKED__?: boolean;

  /**
   * 演示者视图（Reporter）跨窗口注入入口。
   * 写入：install-proxies.ts → installReporterWindowHook()（挂在 iframe window 上）
   * 读取：index.reporter.html 启动脚本通过 window.opener 链调用 install(target)，
   *       在 RequireJS 加载 SDK 前完成 Mock 注入。
   */
  __ONLYOFFICE_REPORTER_BRIDGE__?: {
    install: (target: Window) => void;
  };

  /**
   * iframe 的 window.open 已劫持，用于在打开 index.reporter.html 时尽早注入 Mock。
   * 写入：install-proxies.ts → installReporterWindowHook()
   */
  __ONLYOFFICE_REPORTER_HOOK__?: boolean;
};
